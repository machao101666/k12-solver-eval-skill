# -*- coding: utf-8 -*-
"""
stage4_consensus_judge.py - K12作业解题共识裁判脚本
读取 evaluation.xlsx，从5个模型的 reasoning/answer 列获取作答，
调用API进行一致性校验、置信度分级、择优选定标准伪标签，
将结果写入最后5列（confidence, selected_ref_id, consensus_answer, pseudo_label, filter_reason）。
"""

import argparse
import os
import re
import json
import time
import threading
import requests
import openpyxl
import base64
from io import BytesIO
from PIL import Image

from logger import get_logger
from config import cfg, column, prompt, soft_secret, require_secrets

# 并发控制
API_SEMAPHORE = threading.Semaphore(int(cfg('runtime.consensus_concurrency', 8)))
RATE_LIMIT_LOCK = threading.Lock()
_rate_limit_until = 0
EXCEL_LOCK = threading.Lock()


def wait_if_rate_limited():
    """检查是否处于限流状态，如果是则等待"""
    global _rate_limit_until
    with RATE_LIMIT_LOCK:
        remaining = _rate_limit_until - time.time()
    if remaining > 0:
        print(f"    [限流等待] {remaining:.0f}s...")
        time.sleep(remaining + 1)


# ============================================
# 配置：全部来自 config.yaml / .env，见 config.py
# ============================================
_JUDGE_PROVIDER = cfg('models.consensus_judge.provider', 'openai_compat')
API_URL = cfg(f'providers.{_JUDGE_PROVIDER}.base_url')
API_KEY = soft_secret(f'providers.{_JUDGE_PROVIDER}.api_key')

EXCEL_FILE = cfg('paths.excel_file', 'evaluation.xlsx')
BASE_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))

HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {API_KEY}',
}

# 共识裁判模型配置
MODEL_CONFIG = {
    'label': 'consensus_judge',
    'model': cfg('models.consensus_judge.model'),
    'provider': 'claude',
    'temperature': cfg('models.consensus_judge.temperature', None),
    'enable_thinking': cfg('models.consensus_judge.enable_thinking', True),
}

# 参考模型 ID 到模型名的映射，顺序与 config.yaml 的 models.reference_models 一致
REF_ID_TO_MODEL_NAME = {
    f'M{index}': item['model']
    for index, item in enumerate(cfg('models.reference_models'), start=1)
}

# 列配置（1-based），全部来自 config.yaml 的 excel.columns
COL_IMAGE_PATH = column('image_path')
COL_XUEDUAN = column('grade_band')
COL_XUEKE = column('subject')
COL_TIXING = column('question_type')

# 5 个参考模型的输入列，来自 config.yaml 的 models.reference_models[*].excel_columns
_REF_COLUMNS = [item['excel_columns'] for item in cfg('models.reference_models')]
COL_M1_REASONING, COL_M1_ANSWER = int(_REF_COLUMNS[0]['reasoning']), int(_REF_COLUMNS[0]['answer'])
COL_M2_REASONING, COL_M2_ANSWER = int(_REF_COLUMNS[1]['reasoning']), int(_REF_COLUMNS[1]['answer'])
COL_M3_REASONING, COL_M3_ANSWER = int(_REF_COLUMNS[2]['reasoning']), int(_REF_COLUMNS[2]['answer'])
COL_M4_REASONING, COL_M4_ANSWER = int(_REF_COLUMNS[3]['reasoning']), int(_REF_COLUMNS[3]['answer'])
COL_M5_REASONING, COL_M5_ANSWER = int(_REF_COLUMNS[4]['reasoning']), int(_REF_COLUMNS[4]['answer'])

# 输出列
COL_CONFIDENCE = column('confidence')
COL_SELECTED_REF_ID = column('selected_ref_id')
COL_CONSENSUS_ANSWER = column('consensus_answer')
COL_PSEUDO_LABEL = column('pseudo_label')
COL_FILTER_REASON = column('filter_reason')

# 提示词模板（外置到 prompts/，见 prompts/README.md）
PROMPT_TEMPLATE = prompt('consensus')


def build_prompt(
    xueduan,
    xueke,
    tixing,
    m1_reasoning,
    m1_answer,
    m2_reasoning,
    m2_answer,
    m3_reasoning,
    m3_answer,
    m4_reasoning,
    m4_answer,
    m5_reasoning,
    m5_answer,
):
    """构建共识裁判提示词"""
    if not xueduan or not xueke or not tixing:
        raise ValueError("学段、学科、题型不能为空")
    if not m1_reasoning or not m1_answer:
        raise ValueError("模型1的推理过程和答案不能为空")
    if not m2_reasoning or not m2_answer:
        raise ValueError("模型2的推理过程和答案不能为空")
    if not m3_reasoning or not m3_answer:
        raise ValueError("模型3的推理过程和答案不能为空")
    if not m4_reasoning or not m4_answer:
        raise ValueError("模型4的推理过程和答案不能为空")
    if not m5_reasoning or not m5_answer:
        raise ValueError("模型5的推理过程和答案不能为空")

    return PROMPT_TEMPLATE.format(
        学段=xueduan,
        学科=xueke,
        题型=tixing,
        M1推理过程=m1_reasoning,
        M1最终答案=m1_answer,
        M2推理过程=m2_reasoning,
        M2最终答案=m2_answer,
        M3推理过程=m3_reasoning,
        M3最终答案=m3_answer,
        M4推理过程=m4_reasoning,
        M4最终答案=m4_answer,
        M5推理过程=m5_reasoning,
        M5最终答案=m5_answer,
    )


def image_to_base64_data_url(image_path):
    """将本地图片文件转为 base64 data URL，超5MB时自动压缩到限制内"""
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"图片文件不存在: {image_path}")

    ext = os.path.splitext(image_path)[1].lower()
    mime_map = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
        '.webp': 'image/webp',
    }
    mime_type = mime_map.get(ext, 'image/png')
    max_bytes = int(cfg('runtime.max_image_bytes', 5 * 1024 * 1024))

    with open(image_path, 'rb') as f:
        raw = f.read()

    # 原图在限制内，直接使用
    if len(raw) <= max_bytes:
        img_data = base64.b64encode(raw).decode('utf-8')
        return f"data:{mime_type};base64,{img_data}"

    # 超过5MB时，统一转JPEG并逐步压缩
    img = Image.open(image_path)
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')

    quality = 85
    scale = 1.0
    compressed = None

    for _ in range(10):
        target_w = max(1, int(img.width * scale))
        target_h = max(1, int(img.height * scale))
        resized = img.resize((target_w, target_h), Image.LANCZOS) if scale < 1.0 else img

        buf = BytesIO()
        resized.save(buf, format='JPEG', quality=quality, optimize=True)
        data = buf.getvalue()

        if len(data) <= max_bytes:
            compressed = data
            break

        # 先降质量，质量过低后再缩尺寸
        if quality > 45:
            quality -= 10
        else:
            scale *= 0.85

    if compressed is None:
        compressed = data

    img_data = base64.b64encode(compressed).decode('utf-8')
    return f"data:image/jpeg;base64,{img_data}"


def call_api(image_data_url, prompt_text, max_retries=1):
    """
    调用百度云API进行共识裁判
    返回 parsed_result 或 None
    """
    import base64

    messages = [
        {
            'role': 'user',
            'content': [
                {
                    'type': 'image_url',
                    'image_url': {'url': image_data_url},
                },
                {
                    'type': 'text',
                    'text': prompt_text,
                },
            ],
        }
    ]

    payload = {
        'model': MODEL_CONFIG['model'],
        'messages': messages,
        'stream': False,
        'enable_thinking': MODEL_CONFIG['enable_thinking'],
    }
    if MODEL_CONFIG.get('temperature') is not None:
        payload['temperature'] = MODEL_CONFIG['temperature']

    last_error = ''

    for attempt in range(1, max_retries + 2):
        try:
            response = requests.post(
                API_URL,
                headers=HEADERS,
                data=json.dumps(payload),
                timeout=int(cfg('runtime.request_timeout', 200)),
            )
            result = response.json()

            if 'error' in result:
                error_code = result['error'].get('code', '')
                error_msg = result['error'].get('message', str(result['error']))

                # 限流错误：指数退避后重试
                if 'qps' in str(error_msg).lower() or 'rate' in str(error_msg).lower() or error_code in ['18', '336005', '336006', 18, 336005, 336006]:
                    wait = min(5 * (attempt ** 2), 60)
                    print(f"    限流(第{attempt}次)，等待{wait}s...")
                    with RATE_LIMIT_LOCK:
                        global _rate_limit_until
                        _rate_limit_until = time.time() + wait
                    time.sleep(wait)
                    if attempt <= max_retries:
                        continue
                    return {'error': f"限流(第{attempt}次): {error_msg}"}

                last_error = f"API错误(第{attempt}次): {error_msg}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return {'error': last_error}

            choices = result.get('choices', [])
            if not choices:
                last_error = f"API返回无choices(第{attempt}次): {json.dumps(result, ensure_ascii=False)[:200]}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return {'error': last_error}

            message = choices[0].get('message', {})
            content = message.get('content', '')

            # 解析 JSON
            if content:
                result = parse_json_response(content)
                return result

            last_error = f"返回内容为空(第{attempt}次)"
            if attempt <= max_retries:
                print(f"    重试 {attempt}/{max_retries}...")
                time.sleep(2)
                continue

        except requests.exceptions.Timeout:
            last_error = f'请求超时(第{attempt}次)'
            if attempt <= max_retries:
                print(f"    重试 {attempt}/{max_retries}...")
                time.sleep(2)
                continue
            break
        except requests.exceptions.RequestException as e:
            last_error = f'网络错误(第{attempt}次): {str(e)}'
            if attempt <= max_retries:
                print(f"    重试 {attempt}/{max_retries}...")
                time.sleep(2)
                continue
            break
        except Exception as e:
            last_error = f'未知错误(第{attempt}次): {str(e)}'
            if attempt <= max_retries:
                print(f"    重试 {attempt}/{max_retries}...")
                time.sleep(2)
                continue
            break

    return {'error': f'{last_error}（已重试{max_retries}次）'}


def parse_json_response(content):
    """解析API返回的JSON内容"""
    # 尝试提取JSON（可能在```json块中或纯文本）
    json_text = content.strip()

    # 移除可能的 markdown 代码块
    json_text = re.sub(r'^```json\s*', '', json_text, flags=re.IGNORECASE)
    json_text = re.sub(r'^```\s*', '', json_text)
    json_text = re.sub(r'\s*```$', '', json_text)

    # 尝试找到 JSON 对象
    try:
        # 尝试直接解析
        return json.loads(json_text)
    except json.JSONDecodeError:
        # 尝试查找 JSON 对象
        match = re.search(r'\{[\s\S]*\}', json_text)
        if match:
            try:
                return json.loads(match.group(0))
            except json.JSONDecodeError:
                pass

    return {'raw_content': content}


def write_result_to_excel(ws, row_idx, result):
    """将结果写入Excel对应列"""
    confidence = result.get('confidence', '')
    filter_reason = result.get('filter_reason', '')

    # 写入 confidence 和 filter_reason（必定有）
    ws.cell(row=row_idx, column=COL_CONFIDENCE, value=confidence)
    ws.cell(row=row_idx, column=COL_FILTER_REASON, value=filter_reason)

    # 只有 high 时才写入其他字段
    if confidence == 'high':
        selected_ref_id = result.get('selected_ref_id', '')
        selected_model_name = REF_ID_TO_MODEL_NAME.get(selected_ref_id, selected_ref_id)
        ws.cell(row=row_idx, column=COL_SELECTED_REF_ID, value=selected_model_name)
        ws.cell(row=row_idx, column=COL_CONSENSUS_ANSWER, value=result.get('consensus_answer', ''))
        ws.cell(row=row_idx, column=COL_PSEUDO_LABEL, value=result.get('pseudo_label', ''))
    else:
        # medium/low 时清空其他字段
        ws.cell(row=row_idx, column=COL_SELECTED_REF_ID, value='')
        ws.cell(row=row_idx, column=COL_CONSENSUS_ANSWER, value='')
        ws.cell(row=row_idx, column=COL_PSEUDO_LABEL, value='')


def parse_args():
    parser = argparse.ArgumentParser(description="共识投票与伪标签生成")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    return parser.parse_args()


def process_excel(work_dir=None):
    """主流程：读取 evaluation.xlsx，并发处理共识裁判"""
    require_secrets(f'providers.{_JUDGE_PROVIDER}.base_url', f'providers.{_JUDGE_PROVIDER}.api_key')
    import base64

    base_dir = os.path.abspath(work_dir or BASE_DIR)
    os.environ["AUTO_EVAL_WORK_DIR"] = base_dir
    xlsx_path = os.path.join(base_dir, EXCEL_FILE)

    if not os.path.exists(xlsx_path):
        print(f"[ERROR] 找不到文件: {xlsx_path}")
        return

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    total_rows = ws.max_row

    logger = get_logger(base_dir=base_dir)

    stage_ctx = logger.stage("consensus", desc="共识投票与伪标签生成（B模块）", stage_order="04")
    stage = stage_ctx.__enter__()
    stage.log_file_input(xlsx_path)
    stage.info(f"共 {total_rows - 1} 行数据，开始处理...")

    # 预加载所有行数据
    rows_data = []
    for row_idx in range(2, total_rows + 1):
        # 读取基本信息
        image_rel_path = ws.cell(row=row_idx, column=COL_IMAGE_PATH).value
        xueduan = ws.cell(row=row_idx, column=COL_XUEDUAN).value or ''
        xueke = ws.cell(row=row_idx, column=COL_XUEKE).value or ''
        tixing = ws.cell(row=row_idx, column=COL_TIXING).value or ''

        # 读取五个模型的作答
        m1_reasoning = ws.cell(row=row_idx, column=COL_M1_REASONING).value or ''
        m1_answer = ws.cell(row=row_idx, column=COL_M1_ANSWER).value or ''
        m2_reasoning = ws.cell(row=row_idx, column=COL_M2_REASONING).value or ''
        m2_answer = ws.cell(row=row_idx, column=COL_M2_ANSWER).value or ''
        m3_reasoning = ws.cell(row=row_idx, column=COL_M3_REASONING).value or ''
        m3_answer = ws.cell(row=row_idx, column=COL_M3_ANSWER).value or ''
        m4_reasoning = ws.cell(row=row_idx, column=COL_M4_REASONING).value or ''
        m4_answer = ws.cell(row=row_idx, column=COL_M4_ANSWER).value or ''
        m5_reasoning = ws.cell(row=row_idx, column=COL_M5_REASONING).value or ''
        m5_answer = ws.cell(row=row_idx, column=COL_M5_ANSWER).value or ''

        # 检查是否有有效图片路径
        if not image_rel_path:
            rows_data.append({
                'row_idx': row_idx,
                'skip': True,
                'reason': '图片路径为空'
            })
            continue

        image_full_path = os.path.join(base_dir, image_rel_path)
        if not os.path.exists(image_full_path):
            rows_data.append({
                'row_idx': row_idx,
                'skip': True,
                'reason': f'图片文件不存在: {image_rel_path}'
            })
            continue

        try:
            image_data_url = image_to_base64_data_url(image_full_path)
            rows_data.append({
                'row_idx': row_idx,
                'skip': False,
                'image_data_url': image_data_url,
                'xueduan': xueduan,
                'xueke': xueke,
                'tixing': tixing,
                'm1_reasoning': m1_reasoning,
                'm1_answer': m1_answer,
                'm2_reasoning': m2_reasoning,
                'm2_answer': m2_answer,
                'm3_reasoning': m3_reasoning,
                'm3_answer': m3_answer,
                'm4_reasoning': m4_reasoning,
                'm4_answer': m4_answer,
                'm5_reasoning': m5_reasoning,
                'm5_answer': m5_answer,
            })
        except Exception as e:
            rows_data.append({
                'row_idx': row_idx,
                'skip': True,
                'reason': f'图片加载失败: {str(e)}'
            })

    total_requests = sum(1 for r in rows_data if not r.get('skip'))
    completed_count = [0]
    save_counter = [0]

    def process_row(row_data):
        """处理单行数据"""
        row_idx = row_data['row_idx']

        if row_data.get('skip'):
            reason = row_data.get('reason', '未知原因')
            print(f"[行{row_idx}] 跳过: {reason}")
            with EXCEL_LOCK:
                ws.cell(row=row_idx, column=COL_CONFIDENCE, value='skip')
                ws.cell(row=row_idx, column=COL_FILTER_REASON, value=reason)
            return

        # 构建提示词
        prompt_text = build_prompt(
            row_data['xueduan'],
            row_data['xueke'],
            row_data['tixing'],
            row_data['m1_reasoning'],
            row_data['m1_answer'],
            row_data['m2_reasoning'],
            row_data['m2_answer'],
            row_data['m3_reasoning'],
            row_data['m3_answer'],
            row_data['m4_reasoning'],
            row_data['m4_answer'],
            row_data['m5_reasoning'],
            row_data['m5_answer'],
        )

        print(f"[行{row_idx}] 调用共识裁判 API...")

        wait_if_rate_limited()

        # 获取信号量后调用 API
        API_SEMAPHORE.acquire()
        try:
            result = call_api(row_data['image_data_url'], prompt_text)
        finally:
            API_SEMAPHORE.release()

        # 处理结果
        if 'error' in result:
            print(f"[行{row_idx}] API错误: {result['error']}")
            with EXCEL_LOCK:
                ws.cell(row=row_idx, column=COL_CONFIDENCE, value='error')
                ws.cell(row=row_idx, column=COL_FILTER_REASON, value=result['error'])
        else:
            write_result_to_excel(ws, row_idx, result)
            confidence = result.get('confidence', 'unknown')
            print(f"[行{row_idx}] 完成: confidence={confidence}")

        completed_count[0] += 1
        save_counter[0] += 1
        if save_counter[0] % 5 == 0:
            with EXCEL_LOCK:
                wb.save(xlsx_path)
            print(f"    [自动保存] {save_counter[0]} 条记录已保存")

    # 启动并发处理
    threads = []
    for row_data in rows_data:
        t = threading.Thread(target=process_row, args=(row_data,))
        t.start()
        threads.append(t)
        # 控制并发数
        time.sleep(0.5)

    for t in threads:
        t.join()

    # 最终保存
    wb.save(xlsx_path)
    stage.log_file_output(xlsx_path)
    stage.info(f"全部处理完成！共处理 {completed_count[0]} 条记录")
    stage_ctx.__exit__(None, None, None)
    logger.finish()


if __name__ == '__main__':
    cli_args = parse_args()
    process_excel(work_dir=cli_args.work_dir)

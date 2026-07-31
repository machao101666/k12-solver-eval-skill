# -*- coding: utf-8 -*-
"""
stage7_score_judge.py - K12作业解题模型专业评分裁判脚本
读取 evaluation.xlsx，获取伪标签与待测输出，调用裁判模型评分，
将 score_0_3 写入 AI 列，将 judge_reason 写入 AJ 列。
"""

import argparse
import base64
import os
import re
import json
import time
import threading
from io import BytesIO

import requests
import openpyxl

from openpyxl.utils.exceptions import IllegalCharacterError
from PIL import Image

from logger import get_logger
from config import cfg, column, prompt, soft_secret, require_secrets


# 并发与限流控制
API_SEMAPHORE = threading.Semaphore(int(cfg('runtime.score_concurrency', 8)))
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
# 配置
# ============================================
_JUDGE_PROVIDER = cfg('models.score_judge.provider', 'openai_compat')
API_URL = cfg(f'providers.{_JUDGE_PROVIDER}.base_url')
API_KEY = soft_secret(f'providers.{_JUDGE_PROVIDER}.api_key')

EXCEL_FILE = cfg('paths.excel_file', 'evaluation.xlsx')
_ONLY_CONFIDENCE = str(cfg('evaluation.only_confidence', 'high')).strip().lower()
BASE_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))

HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {API_KEY}',
}

# 评分裁判模型配置
MODEL_CONFIG = {
    'label': 'score_judge',
    'model': cfg('models.score_judge.model'),
    'provider': 'claude',
    'temperature': cfg('models.score_judge.temperature', None),
    'enable_thinking': cfg('models.score_judge.enable_thinking', True),
}

# 列配置（1-based），全部来自 config.yaml 的 excel.columns
COL_IMAGE_PATH = column('image_path')           # B
COL_GRADE_BAND = column('grade_band')           # C
COL_SUBJECT = column('subject')                 # D
COL_QUESTION_TYPE = column('question_type')     # E
COL_CONFIDENCE = column('confidence')           # Z
COL_PSEUDO_LABEL = column('pseudo_label')       # AC
COL_MODEL_FULL_OUTPUT = column('model_output')  # AF

# 输出列：待测模型评分
COL_SCORE_0_3 = column('target_score')          # AI
COL_JUDGE_REASON = column('target_reason')      # AJ

# 参考模型横向评分目标：解题输出列来自 models.reference_models[*].excel_columns，
# 评分写回列来自 excel.reference_score_columns，两者顺序必须一一对应。
def _build_reference_targets():
    ref_models = cfg('models.reference_models')
    score_columns = cfg('excel.reference_score_columns')
    if len(ref_models) != len(score_columns):
        raise RuntimeError(
            f"配置不一致：models.reference_models 有 {len(ref_models)} 个模型，"
            f"excel.reference_score_columns 有 {len(score_columns)} 组评分列，必须一一对应。"
        )
    targets = []
    for model_item, score_item in zip(ref_models, score_columns):
        label = score_item.get('label') or model_item['label']
        targets.append({
            'label': label,
            'reasoning_col': int(model_item['excel_columns']['reasoning']),
            'answer_col': int(model_item['excel_columns']['answer']),
            'score_col': int(score_item['score']),
            'reason_col': int(score_item['reason']),
            'score_header': f'{label}_score_0_3',
            'reason_header': 'judge_reason',
        })
    return targets


REFERENCE_MODEL_SCORE_TARGETS = _build_reference_targets()

# 提示词模板（外置到 prompts/，见 prompts/README.md）
PROMPT_TEMPLATE = prompt('score_target')


REFERENCE_PROMPT_TEMPLATE = prompt('score_reference')


def build_prompt(grade_band, subject, question_type, pseudo_label_text, model_full_output):
    """构建待测模型评分裁判提示词。"""
    return PROMPT_TEMPLATE.format(
        image='见输入图片',
        grade_band=grade_band or '',
        subject=subject or '',
        question_type=question_type or '',
        pseudo_label_text=pseudo_label_text or '',
        model_full_output=model_full_output or '',
    )


def build_reference_prompt(grade_band, subject, question_type, pseudo_label_text, model_full_text):
    """构建五个参考模型横向评分提示词。"""
    return REFERENCE_PROMPT_TEMPLATE.format(
        image='见输入图片',
        grade_band=grade_band or '',
        subject=subject or '',
        question_type=question_type or '',
        pseudo_label_text=pseudo_label_text or '',
        model_full_text=model_full_text or '',
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

    if len(raw) <= max_bytes:
        img_data = base64.b64encode(raw).decode('utf-8')
        return f"data:{mime_type};base64,{img_data}"

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

        if quality > 45:
            quality -= 10
        else:
            scale *= 0.85

    if compressed is None:
        compressed = data

    img_data = base64.b64encode(compressed).decode('utf-8')
    return f"data:image/jpeg;base64,{img_data}"


def parse_json_response(content):
    """解析裁判响应；出现多个JSON时采用最后一个完整评分结果。"""
    json_text = content.strip()
    json_text = re.sub(r'^```json\s*', '', json_text, flags=re.IGNORECASE)
    json_text = re.sub(r'^```\s*', '', json_text)
    json_text = re.sub(r'\s*```$', '', json_text)

    def valid_score_object(value):
        return isinstance(value, dict) and 'score_0_3' in value and 'judge_reason' in value

    try:
        parsed = json.loads(json_text)
        if valid_score_object(parsed):
            return parsed
        if isinstance(parsed, dict) and isinstance(parsed.get('raw_content'), str):
            nested = parse_json_response(parsed['raw_content'])
            if valid_score_object(nested):
                return nested
    except json.JSONDecodeError:
        pass

    decoder = json.JSONDecoder()
    objects = []
    position = 0
    while position < len(json_text):
        start = json_text.find('{', position)
        if start < 0:
            break
        try:
            parsed, end = decoder.raw_decode(json_text, start)
            if valid_score_object(parsed):
                objects.append(parsed)
            position = end
        except json.JSONDecodeError:
            position = start + 1

    if objects:
        return objects[-1]
    return {'raw_content': content}


def normalize_score(value):
    """将模型输出严格标准化为 0~3 整数，失败返回空字符串。"""
    if isinstance(value, bool):
        return ''
    if isinstance(value, int) and value in (0, 1, 2, 3):
        return value
    if isinstance(value, float) and value.is_integer() and int(value) in (0, 1, 2, 3):
        return int(value)
    if isinstance(value, str):
        match = re.fullmatch(r'\s*([0-3])(?:\.0+)?\s*(?:分)?\s*', value)
        if match:
            return int(match.group(1))
    return ''


def sanitize_excel_text(value):
    """移除Excel不支持的控制字符，避免IllegalCharacterError"""
    if value is None:
        return ''
    text = str(value)
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)


def call_api(image_data_url, prompt_text, max_retries=1):
    """调用评分裁判API，返回 dict；首次调用 + 最多 1 次重试，每次限时 200s。"""
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
            if content:
                return parse_json_response(content)

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


def write_result_to_excel(ws, row_idx, result, score_col=COL_SCORE_0_3, reason_col=COL_JUDGE_REASON):
    """写入评分结果"""
    score = normalize_score(result.get('score_0_3', ''))
    reason = sanitize_excel_text(result.get('judge_reason', '')).strip()

    ws.cell(row=row_idx, column=score_col, value=score)
    ws.cell(row=row_idx, column=reason_col, value=reason)


def build_reference_model_output(reasoning, answer):
    """将参考模型的 reasoning/answer 拼为待评分完整输出。"""
    reasoning = sanitize_excel_text(reasoning or '')
    answer = sanitize_excel_text(answer or '')
    if not reasoning and not answer:
        return ''
    return f"Reasoning：{reasoning}\n\nAnswer：{answer}"


def ensure_output_headers(ws, targets="all"):
    """补齐 Stage 7 输出列表头。"""
    score_header = 'score_0_3' if targets == 'target' else 'target_score_0_3'
    ws.cell(row=1, column=COL_SCORE_0_3, value=score_header)
    ws.cell(row=1, column=COL_JUDGE_REASON, value='judge_reason')
    for target in REFERENCE_MODEL_SCORE_TARGETS:
        ws.cell(row=1, column=target['score_col'], value=target['score_header'])
        ws.cell(row=1, column=target['reason_col'], value=target['reason_header'])


def parse_args():
    parser = argparse.ArgumentParser(description="待测模型与五参考模型评分裁判")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    parser.add_argument("--concurrency", type=int, default=None, help="评分请求级并发数，默认取 config.yaml 的 runtime.score_concurrency")
    parser.add_argument("--targets", choices=("all", "target", "references"), default="all", help="评分目标：全部、仅待测模型或仅五参考模型")
    parser.add_argument("--invalid-retries", type=int, default=None, help="裁判响应无法解析为合法评分时的额外重试次数，默认取 config.yaml 的 evaluation.invalid_retries")
    parser.add_argument("--only-invalid", action="store_true", help="只补跑分数非法或理由为空的评分目标")
    return parser.parse_args()


def process_excel(work_dir=None, concurrency=None, targets="all", invalid_retries=None, only_invalid=False):
    """主流程：读取 evaluation.xlsx，以单个评分请求为粒度并发评分。"""
    require_secrets(f'providers.{_JUDGE_PROVIDER}.base_url', f'providers.{_JUDGE_PROVIDER}.api_key')
    from concurrent.futures import ThreadPoolExecutor, as_completed

    global API_SEMAPHORE
    if invalid_retries is None:
        invalid_retries = int(cfg('evaluation.invalid_retries', 2))
    concurrency = max(1, int(concurrency or cfg('runtime.score_concurrency', 8)))
    API_SEMAPHORE = threading.Semaphore(concurrency)

    base_dir = os.path.abspath(work_dir or BASE_DIR)
    os.environ["AUTO_EVAL_WORK_DIR"] = base_dir
    logger = get_logger(base_dir=base_dir)
    xlsx_path = os.path.join(base_dir, EXCEL_FILE)
    if not os.path.exists(xlsx_path):
        print(f"[ERROR] 找不到文件: {xlsx_path}")
        return

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    ensure_output_headers(ws, targets=targets)
    total_rows = ws.max_row

    stage_ctx = logger.stage("score", desc="专业评分裁判（C模块）", stage_order="07")
    stage = stage_ctx.__enter__()
    stage.log_file_input(xlsx_path)
    stage.info(f"共 {total_rows - 1} 行数据，开始处理（仅confidence={_ONLY_CONFIDENCE}，评分请求级并发={concurrency}）...")

    row_indices = list(range(2, total_rows + 1))
    stats = {"processed": 0, "skipped": 0, "failed": 0, "target_scored": 0, "score_tasks": 0, "score_task_failed": 0}
    row_states = {}
    score_tasks = []

    def mark_row_failed(row_idx, score_col, reason_col, reason):
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=score_col, value='')
            ws.cell(row=row_idx, column=reason_col, value=reason)
        row_states[row_idx] = {'failed': True, 'target_scored': False, 'summaries': [reason]}

    for row_idx in row_indices:
        confidence = (ws.cell(row=row_idx, column=COL_CONFIDENCE).value or '').strip().lower()
        if confidence != _ONLY_CONFIDENCE:
            stats['skipped'] += 1
            message = f"[行{row_idx}] 跳过: confidence={confidence or '空'}"
            print(message)
            stage.info(message)
            continue

        row_states[row_idx] = {'failed': False, 'target_scored': False, 'summaries': []}
        image_rel_path = ws.cell(row=row_idx, column=COL_IMAGE_PATH).value
        grade_band = ws.cell(row=row_idx, column=COL_GRADE_BAND).value or ''
        subject = ws.cell(row=row_idx, column=COL_SUBJECT).value or ''
        question_type = ws.cell(row=row_idx, column=COL_QUESTION_TYPE).value or ''
        pseudo_label_text = ws.cell(row=row_idx, column=COL_PSEUDO_LABEL).value or ''
        model_full_output = ws.cell(row=row_idx, column=COL_MODEL_FULL_OUTPUT).value or ''

        if only_invalid and targets == "target":
            existing_score = normalize_score(ws.cell(row=row_idx, column=COL_SCORE_0_3).value)
            existing_reason = sanitize_excel_text(ws.cell(row=row_idx, column=COL_JUDGE_REASON).value).strip()
            if existing_score in (0, 1, 2, 3) and existing_reason:
                row_states[row_idx]['target_scored'] = True
                row_states[row_idx]['summaries'].append(f"target={existing_score}(已有合法评分，跳过)")
                continue

        if not image_rel_path:
            mark_row_failed(row_idx, COL_SCORE_0_3, COL_JUDGE_REASON, '图片路径为空，未评分')
            continue

        image_full_path = os.path.join(base_dir, str(image_rel_path))
        if not os.path.exists(image_full_path):
            mark_row_failed(row_idx, COL_SCORE_0_3, COL_JUDGE_REASON, f'图片不存在: {image_rel_path}')
            continue

        try:
            image_data_url = image_to_base64_data_url(image_full_path)
        except Exception as e:
            mark_row_failed(row_idx, COL_SCORE_0_3, COL_JUDGE_REASON, f'图片编码失败: {e}')
            continue

        if targets in ("all", "target"):
            existing_score = normalize_score(ws.cell(row=row_idx, column=COL_SCORE_0_3).value)
            existing_reason = sanitize_excel_text(ws.cell(row=row_idx, column=COL_JUDGE_REASON).value).strip()
            if only_invalid and existing_score in (0, 1, 2, 3) and existing_reason:
                row_states[row_idx]['target_scored'] = True
                row_states[row_idx]['summaries'].append(f"target={existing_score}(已有合法评分，跳过)")
            else:
                score_tasks.append({
                    'row_idx': row_idx,
                    'label': 'target',
                    'image_data_url': image_data_url,
                    'prompt': build_prompt(
                        grade_band=grade_band,
                        subject=subject,
                        question_type=question_type,
                        pseudo_label_text=pseudo_label_text,
                        model_full_output=model_full_output,
                    ),
                    'score_col': COL_SCORE_0_3,
                    'reason_col': COL_JUDGE_REASON,
                    'is_target': True,
                })

        if targets in ("all", "references"):
            for target in REFERENCE_MODEL_SCORE_TARGETS:
                reference_output = build_reference_model_output(
                    ws.cell(row=row_idx, column=target['reasoning_col']).value,
                    ws.cell(row=row_idx, column=target['answer_col']).value,
                )
                existing_score = normalize_score(ws.cell(row=row_idx, column=target['score_col']).value)
                existing_reason = sanitize_excel_text(ws.cell(row=row_idx, column=target['reason_col']).value).strip()
                if only_invalid and existing_score in (0, 1, 2, 3) and existing_reason:
                    row_states[row_idx]['summaries'].append(f"{target['label']}={existing_score}(已有合法评分，跳过)")
                    continue
                score_tasks.append({
                    'row_idx': row_idx,
                    'label': target['label'],
                    'image_data_url': image_data_url,
                    'prompt': build_reference_prompt(
                        grade_band=grade_band,
                        subject=subject,
                        question_type=question_type,
                        pseudo_label_text=pseudo_label_text,
                        model_full_text=reference_output,
                    ),
                    'score_col': target['score_col'],
                    'reason_col': target['reason_col'],
                    'is_target': False,
                })

    stats['score_tasks'] = len(score_tasks)
    stage.info(f"已生成评分请求任务 {len(score_tasks)} 个，开始请求级并发执行...")

    def process_score_task(task):
        t0 = time.time()
        attempts = max(0, int(invalid_retries or 0)) + 1
        result = None
        for attempt in range(1, attempts + 1):
            wait_if_rate_limited()
            API_SEMAPHORE.acquire()
            try:
                result = call_api(task['image_data_url'], task['prompt'])
            finally:
                API_SEMAPHORE.release()
            score = normalize_score(result.get('score_0_3', '')) if isinstance(result, dict) else ''
            reason = sanitize_excel_text(result.get('judge_reason', '')).strip() if isinstance(result, dict) else ''
            if 'error' not in result and score in (0, 1, 2, 3) and reason:
                break
            if attempt < attempts:
                print(f"    [行{task['row_idx']}] {task['label']} 裁判结果无效，额外重试 {attempt}/{attempts - 1}...")
        score = normalize_score(result.get('score_0_3', '')) if isinstance(result, dict) else ''
        reason = sanitize_excel_text(result.get('judge_reason', '')).strip() if isinstance(result, dict) else ''
        if 'error' not in result and (score not in (0, 1, 2, 3) or not reason):
            raw = sanitize_excel_text(result.get('raw_content', '')) if isinstance(result, dict) else ''
            result = {'error': f"裁判响应无法解析为合法评分（已尝试{attempts}次）: {raw[:500]}"}
        elapsed = round(time.time() - t0, 2)
        return {**task, 'result': result, 'elapsed': elapsed}

    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        futures = [executor.submit(process_score_task, task) for task in score_tasks]
        for future in as_completed(futures):
            item = future.result()
            row_idx = item['row_idx']
            label = item['label']
            result = item['result']
            elapsed = item['elapsed']
            state = row_states.setdefault(row_idx, {'failed': False, 'target_scored': False, 'summaries': []})

            with EXCEL_LOCK:
                if 'error' in result:
                    ws.cell(row=row_idx, column=item['score_col'], value='')
                    ws.cell(row=row_idx, column=item['reason_col'], value=result['error'])
                    state['summaries'].append(f"{label}=API错误({elapsed}s)")
                    state['failed'] = True
                    stats['score_task_failed'] += 1
                    message = f"[行{row_idx}] {label}评分失败，耗时 {elapsed}s"
                else:
                    write_result_to_excel(ws, row_idx, result, item['score_col'], item['reason_col'])
                    score = ws.cell(row=row_idx, column=item['score_col']).value
                    state['summaries'].append(f"{label}={score}({elapsed}s)")
                    if item['is_target'] and score != '':
                        state['target_scored'] = True
                    message = f"[行{row_idx}] {label}评分完成: {score}，耗时 {elapsed}s"

            print(message)
            stage.info(message)

    for row_idx, state in sorted(row_states.items()):
        summary = ', '.join(state['summaries'])
        if state['target_scored']:
            stats['processed'] += 1
            stats['target_scored'] += 1
            status = 'success'
        elif state['failed']:
            stats['failed'] += 1
            status = 'partial'
        else:
            status = 'skipped'
        message = f"[行{row_idx}] 汇总: {status}; {summary}"
        print(message)
        stage.info(message)

    temp_xlsx_path = f"{xlsx_path}.tmp.xlsx"
    wb.save(temp_xlsx_path)
    os.replace(temp_xlsx_path, xlsx_path)
    stage.log_file_output(xlsx_path)
    stage.info(
        f"全部处理完成: processed={stats['processed']}, skipped={stats['skipped']}, "
        f"failed={stats['failed']}, target_scored={stats['target_scored']}, "
        f"score_tasks={stats['score_tasks']}, score_task_failed={stats['score_task_failed']}；"
        f"单个评分请求失败已写入对应 reason 列，不影响流水线继续"
    )
    stage_ctx.__exit__(None, None, None)

    logger.finish()


if __name__ == '__main__':
    cli_args = parse_args()
    process_excel(
        work_dir=cli_args.work_dir,
        concurrency=cli_args.concurrency,
        targets=cli_args.targets,
        invalid_retries=cli_args.invalid_retries,
        only_invalid=cli_args.only_invalid,
    )

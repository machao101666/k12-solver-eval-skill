# -*- coding: utf-8 -*-
"""
stage3_multi_model_solve.py - K12题目视觉理解大模型批量评测脚本
读取 evaluation.xlsx，对每道题目分别调用 5 个解题大模型，
将返回的 model、reasoning、answer 写回 Excel 对应列。
"""

import argparse
import os
import re
import base64
import json
import time
import threading
import requests
import openpyxl

from logger import get_logger
from config import cfg, prompt, soft_secret, require_secrets

# 并发控制
API_SEMAPHORE = threading.Semaphore(int(cfg('runtime.solve_concurrency', 20)))
RATE_LIMIT_LOCK = threading.Lock()      # 限流等待锁
_rate_limit_until = 0                   # 限流恢复时间戳


def wait_if_rate_limited():
    """检查是否处于限流状态，如果是则等待"""
    global _rate_limit_until
    with RATE_LIMIT_LOCK:
        remaining = _rate_limit_until - time.time()
    if remaining > 0:
        print(f"    [限流等待] {remaining:.0f}s...")
        time.sleep(remaining + 1)
EXCEL_LOCK = threading.Lock()            # Excel 写入锁

# ============================================
# 配置：全部来自 config.yaml / .env，见 config.py
# ============================================
API_URL = cfg('providers.qianfan.base_url')
_QIANFAN_KEY = soft_secret('providers.qianfan.api_key')
API_KEY = _QIANFAN_KEY if _QIANFAN_KEY.startswith('Bearer ') else f'Bearer {_QIANFAN_KEY}'

# OpenAI 兼容聚合网关（doubao / claude / gemini 等走这里）
COMPAT_URL = cfg('providers.openai_compat.base_url', '') or ''
# Anthropic 原生 /v1/messages 端点；未配置时由 base_url 自动推导
COMPAT_MESSAGES_URL = cfg('providers.openai_compat.messages_url', None) \
    or COMPAT_URL.replace('/chat/completions', '/messages')
COMPAT_KEY = soft_secret('providers.openai_compat.api_key')
COMPAT_HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {COMPAT_KEY}',
}

EXCEL_FILE = cfg('paths.excel_file', 'evaluation.xlsx')
BASE_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))

HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': API_KEY,
}

# 提示词模板（外置到 prompts/，见 prompts/README.md）
PROMPT_TEMPLATE = prompt('solve')

# 参考模型配置：全部来自 config.yaml 的 models.reference_models。
# 想换模型集合或调整写回列，改 config.yaml，不用动代码。
def _build_models():
    models = []
    for item in cfg('models.reference_models'):
        columns = item.get('excel_columns', {})
        provider_name = item.get('provider', 'qianfan')
        models.append({
            'label': item['label'],
            'model': item['model'],
            'provider': provider_name,
            'api_style': item.get('api_style') or ('qianfan' if provider_name == 'qianfan' else 'openai_compat'),
            'temperature': item.get('temperature'),
            'enable_thinking': item.get('enable_thinking', True),
            'col_time': int(columns['time']),
            'col_model': int(columns['model']),
            'col_reasoning': int(columns['reasoning']),
            'col_answer': int(columns['answer']),
        })
    return models


MODELS = _build_models()


def image_to_base64_data_url(image_path):
    """将本地图片文件转为 base64 data URL"""
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

    with open(image_path, 'rb') as f:
        img_data = base64.b64encode(f.read()).decode('utf-8')

    return f"data:{mime_type};base64,{img_data}"


def image_to_base64_raw(image_path):
    """将本地图片文件转为纯 base64 字符串（不含 data URL 前缀）"""
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"图片文件不存在: {image_path}")
    with open(image_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')


def build_prompt(xueduan, xueke, tixing):
    """根据学段、学科、题型构建最终提示词"""
    return PROMPT_TEMPLATE.format(学段=xueduan, 学科=xueke, 题型=tixing)


def parse_content(content):
    """
    从模型返回的 content 中提取 Reasoning 和 Answer 部分
    返回 (reasoning, answer)
    """
    reasoning = ''
    answer = ''

    if not content:
        return reasoning, answer

    # 匹配 "Reasoning：" 之后到 "Answer：" 之前的内容
    reasoning_match = re.search(
        r'Reasoning[：:]\s*\n?(.*?)(?=Answer[：:]|\Z)',
        content, re.DOTALL
    )
    if reasoning_match:
        reasoning = reasoning_match.group(1).strip()

    # 匹配 "Answer：" 之后的内容
    answer_match = re.search(
        r'Answer[：:]\s*\n?(.*?)$',
        content, re.DOTALL
    )
    if answer_match:
        answer = answer_match.group(1).strip()

    return reasoning, answer


def call_model(model_config, image_data_url, prompt_text, max_retries=1):
    """
    调用千帆大模型 API，单次限时 200s，最多重试 max_retries 次
    返回 (model_name, reasoning, answer)
    """
    model_name = model_config['model']

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
        'model': model_name,
        'messages': messages,
        'stream': False,
        'temperature': model_config['temperature'],
    }

    if model_config.get('enable_thinking'):
        if 'deepseek' in model_name.lower():
            payload['thinking'] = {'type': 'enabled'}
        else:
            payload['enable_thinking'] = True

    last_error = ''

    for attempt in range(1, max_retries + 2):  # 1 次初始 + max_retries 次重试
        try:
            response = requests.post(
                API_URL,
                headers=HEADERS,
                data=json.dumps(payload),
                timeout=200,
            )
            result = response.json()

            if 'error' in result:
                error_code = result['error'].get('code', '')
                error_msg = result['error'].get('message', str(result['error']))

                # QPS/限流错误：指数退避后重试
                if 'qps' in str(error_msg).lower() or 'rate' in str(error_msg).lower() or error_code in ['18', '336005', '336006', 18, 336005, 336006]:
                    wait = min(5 * (attempt ** 2), 60)
                    print(f"    限流(第{attempt}次)，等待{wait}s...")
                    with RATE_LIMIT_LOCK:
                        global _rate_limit_until
                        _rate_limit_until = time.time() + wait
                    time.sleep(wait)
                    if attempt <= max_retries:
                        continue
                    return model_name, f"限流(第{attempt}次): {error_msg}", ''

                last_error = f"API错误(第{attempt}次): {error_msg}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return model_name, last_error, ''

            choices = result.get('choices', [])
            if not choices:
                last_error = f"API返回无choices(第{attempt}次): {json.dumps(result, ensure_ascii=False)[:200]}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return model_name, last_error, ''

            message = choices[0].get('message', {})
            returned_model = result.get('model', model_name)
            content = message.get('content', '')
            reasoning_content = message.get('reasoning_content', '')

            reasoning, answer = parse_content(content)

            if not reasoning and reasoning_content:
                reasoning = reasoning_content.strip()

            return returned_model, reasoning, answer

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

    # 所有重试耗尽
    return model_name, f'{last_error}（已重试{max_retries}次，跳过）', ''


def call_openai_compat_model(model_config, image_data_url, prompt_text, max_retries=1):
    """
    调用华研 API (OpenAI 兼容格式)，单次限时 200s，最多重试 max_retries 次
    返回 (model_name, reasoning, answer)
    """
    model_name = model_config['model']

    messages = [
        {
            'role': 'user',
            'content': [
                {'type': 'image_url', 'image_url': {'url': image_data_url}},
                {'type': 'text', 'text': prompt_text},
            ],
        }
    ]

    payload = {
        'model': model_name,
        'messages': messages,
        'stream': False,
    }
    # gpt-5系列不接受temperature参数
    if model_config.get('temperature') is not None and 'gpt-5' not in model_name.lower():
        payload['temperature'] = model_config['temperature']

    # gpt-5系列用reasoning_effort，其他用thinking
    if model_config.get('enable_thinking'):
        if 'gpt-5' in model_name.lower():
            payload['reasoning_effort'] = 'high'
        else:
            payload['thinking'] = {'type': 'enabled'}

    last_error = ''
    for attempt in range(1, max_retries + 2):
        try:
            response = requests.post(
                COMPAT_URL,
                headers=COMPAT_HEADERS,
                data=json.dumps(payload),
                timeout=200,
            )
            result = response.json()

            if 'error' in result:
                error_msg = str(result['error'])
                if 'rate' in error_msg.lower() or '429' in error_msg:
                    wait = min(5 * (attempt ** 2), 60)
                    print(f"    限流(第{attempt}次)，等待{wait}s...")
                    time.sleep(wait)
                    if attempt <= max_retries:
                        continue
                    return model_name, f"限流(第{attempt}次): {error_msg}", ''
                last_error = f"API错误(第{attempt}次): {error_msg}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return model_name, last_error, ''

            choices = result.get('choices', [])
            if not choices:
                last_error = f"API返回无choices(第{attempt}次): {json.dumps(result, ensure_ascii=False)[:200]}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return model_name, last_error, ''

            message = choices[0].get('message', {})
            returned_model = result.get('model', model_name)
            content = message.get('content', '')
            reasoning, answer = parse_content(content)
            return returned_model, reasoning, answer

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

    return model_name, f'{last_error}（已重试{max_retries}次，跳过）', ''


def call_claude_model(model_config, image_path, prompt_text, max_retries=1):
    """
    调用 Anthropic 原生 /v1/messages 格式接口（端点见 providers.openai_compat.messages_url）
    返回 (model_name, reasoning, answer)
    """
    model_name = model_config['model']

    img_b64 = image_to_base64_raw(image_path)
    ext = os.path.splitext(image_path)[1].lower()
    media_type = 'image/png' if ext == '.png' else 'image/jpeg'

    content = [
        {'type': 'image', 'source': {'type': 'base64', 'media_type': media_type, 'data': img_b64}},
        {'type': 'text', 'text': prompt_text},
    ]

    payload = {
        'model': model_name,
        'max_tokens': 8192,
        'system': prompt_text,
        'messages': [{'role': 'user', 'content': content}],
        'stream': False,
        'thinking': {'type': 'adaptive'},
        'output_config': {'effort': 'high'},
    }

    claude_headers = {
        'Content-Type': 'application/json',
        'x-api-key': COMPAT_KEY,
        'anthropic-version': '2023-06-01',
    }

    last_error = ''
    for attempt in range(1, max_retries + 2):
        try:
            response = requests.post(
                COMPAT_MESSAGES_URL,
                headers=claude_headers,
                data=json.dumps(payload),
                timeout=200,
            )
            result = response.json()

            if 'error' in result:
                error_msg = str(result['error'])
                if 'rate' in error_msg.lower() or '429' in error_msg:
                    wait = min(5 * (attempt ** 2), 60)
                    print(f"    限流(第{attempt}次)，等待{wait}s...")
                    time.sleep(wait)
                    if attempt <= max_retries:
                        continue
                    return model_name, f"限流(第{attempt}次): {error_msg}", ''
                last_error = f"API错误(第{attempt}次): {error_msg}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return model_name, last_error, ''

            content_blocks = result.get('content', [])
            text_content = ''
            thinking_content = ''
            for block in content_blocks:
                if block.get('type') == 'text':
                    text_content += block.get('text', '')
                elif block.get('type') == 'thinking':
                    thinking_content += block.get('thinking', '')

            returned_model = result.get('model', model_name)
            reasoning, answer = parse_content(text_content)

            if not reasoning and thinking_content:
                reasoning = thinking_content.strip()

            return returned_model, reasoning, answer

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

    return model_name, f'{last_error}（已重试{max_retries}次，跳过）', ''


def dispatch_call(model_config, image_data_url, image_full_path, prompt_text):
    """根据 api_style 分发到不同的 API 调用函数"""
    api_style = model_config.get('api_style') or (
        'qianfan' if model_config.get('provider', 'qianfan') == 'qianfan' else 'openai_compat'
    )
    if api_style == 'claude':
        return call_claude_model(model_config, image_full_path, prompt_text)
    elif api_style == 'openai_compat':
        return call_openai_compat_model(model_config, image_data_url, prompt_text)
    else:
        return call_model(model_config, image_data_url, prompt_text)


def parse_args():
    parser = argparse.ArgumentParser(description="五模型并行解题")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    return parser.parse_args()


def process_excel(work_dir=None):
    """主流程：读取 evaluation.xlsx，并发处理（5模型各自独立流水线，全局并发受限）"""
    require_secrets(*sorted({
        f"providers.{model.get('provider', 'qianfan')}.{field}"
        for model in MODELS
        for field in ('base_url', 'api_key')
    }))
    base_dir = os.path.abspath(work_dir or BASE_DIR)
    os.environ["AUTO_EVAL_WORK_DIR"] = base_dir
    logger = get_logger(base_dir=base_dir)
    xlsx_path = os.path.join(base_dir, EXCEL_FILE)

    if not os.path.exists(xlsx_path):
        print(f"[ERROR] 找不到文件: {xlsx_path}")
        return

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    total_rows = ws.max_row

    stage_ctx = logger.stage("parta", desc="5模型并行解题（A模块）", stage_order="03")
    stage = stage_ctx.__enter__()
    stage.log_file_input(xlsx_path)
    stage.info(f"共 {total_rows - 1} 行数据，并发数=2，开始处理...")

    # 预加载所有行数据
    rows_data = []
    for row_idx in range(2, total_rows + 1):
        image_rel_path = ws.cell(row=row_idx, column=2).value
        if not image_rel_path:
            rows_data.append({'row_idx': row_idx, 'skip': True})
            continue

        xueduan = ws.cell(row=row_idx, column=3).value or ''
        xueke = ws.cell(row=row_idx, column=4).value or ''
        tixing = ws.cell(row=row_idx, column=5).value or ''

        image_full_path = os.path.join(base_dir, image_rel_path)
        try:
            image_data_url = image_to_base64_data_url(image_full_path)
            prompt_text = build_prompt(xueduan, xueke, tixing)
            rows_data.append({
                'row_idx': row_idx,
                'skip': False,
                'image_data_url': image_data_url,
                'image_full_path': image_full_path,
                'prompt_text': prompt_text,
                'xueduan': xueduan,
                'xueke': xueke,
                'tixing': tixing,
            })
        except FileNotFoundError as e:
            rows_data.append({'row_idx': row_idx, 'skip': True, 'error': str(e)})

    total_requests = sum(1 for r in rows_data if not r.get('skip')) * len(MODELS)
    completed_count = [0]  # 用列表共享
    save_counter = [0]     # 每 3 次写入保存一次

    def process_model(model_idx):
        """每个模型独立流水线，顺序处理所有行"""
        model_cfg = MODELS[model_idx]
        label = model_cfg['label']

        for row_data in rows_data:
            row_idx = row_data['row_idx']

            # 跳过无效行
            if row_data.get('skip'):
                err = row_data.get('error', '图片路径为空')
                with EXCEL_LOCK:
                    ws.cell(row=row_idx, column=model_cfg['col_model'], value=model_cfg['model'])
                    ws.cell(row=row_idx, column=model_cfg['col_reasoning'], value=err)
                    ws.cell(row=row_idx, column=model_cfg['col_answer'], value='')
                    ws.cell(row=row_idx, column=model_cfg['col_time'], value=0)
                    save_counter[0] += 1
                    if save_counter[0] % 3 == 0:
                        wb.save(xlsx_path)
                continue

            image_data_url = row_data['image_data_url']
            image_full_path = row_data['image_full_path']
            prompt_text = row_data['prompt_text']
            xd, xk, tx = row_data['xueduan'], row_data['xueke'], row_data['tixing']

            print(f"[行{row_idx}][{label}] 调用 {model_cfg['model']} (学段={xd}, 学科={xk}, 题型={tx})")

            wait_if_rate_limited()

            # 获取信号量后调用 API
            API_SEMAPHORE.acquire()
            try:
                t0 = time.time()
                model_name, reasoning, answer = dispatch_call(model_cfg, image_data_url, image_full_path, prompt_text)
                elapsed = round(time.time() - t0, 1)
            finally:
                API_SEMAPHORE.release()

            # 线程安全写入 Excel
            with EXCEL_LOCK:
                ws.cell(row=row_idx, column=model_cfg['col_model'], value=model_name)
                ws.cell(row=row_idx, column=model_cfg['col_reasoning'], value=reasoning)
                ws.cell(row=row_idx, column=model_cfg['col_answer'], value=answer)
                ws.cell(row=row_idx, column=model_cfg['col_time'], value=elapsed)
                save_counter[0] += 1
                if save_counter[0] % 3 == 0:
                    wb.save(xlsx_path)

            completed_count[0] += 1
            print(f"[行{row_idx}][{label}] 完成 ({elapsed}s, {completed_count[0]}/{total_requests})")
            print(f"  Answer: {answer[:150]}{'...' if len(answer) > 150 else ''}")

    # 启动所有模型线程
    threads = []
    for i in range(len(MODELS)):
        t = threading.Thread(target=process_model, args=(i,), daemon=False)
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    # 最终保存
    wb.save(xlsx_path)
    stage.log_file_output(xlsx_path)
    stage.set_row_count(len(rows_data))
    stage.info(f"全部完成: 共 {completed_count[0]}/{total_requests} 次请求")
    stage_ctx.__exit__(None, None, None)

    logger.finish()


if __name__ == '__main__':
    cli_args = parse_args()
    process_excel(work_dir=cli_args.work_dir)

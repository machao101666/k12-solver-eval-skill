#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
待测模型解题批量调用脚本（AE/B -> AF）
=====================================
读取 evaluation.xlsx 中的题图（AE 列的公网 URL，或 B 列的本地路径），
调用「待测模型」解题接口，结果写入 AF 列、耗时写入 AG 列。

待测模型接口完全可自定义，由 config.yaml 的 providers.target_model.api_style 决定：

  - baidu_solve  （默认 / 推荐）百度智能云 OCR「拍照解题」
                 https://ai.baidu.com/tech/ocr/solve
                 鉴权：API Key + Secret Key 换 access_token
                 端点：POST https://aip.baidubce.com/stream/2.0/ocr/v1/solve
  - openai_compat 任意 OpenAI 兼容的多模态 chat 接口（自建 vLLM / OpenRouter /
                 one-api / 各家官方 OpenAI 兼容端点均可），提示词见
                 prompts/stage6_target_solve.md

两种风格都支持 input_mode = url（用 AE 列的公网 URL）或 image（直接读 B 列本地图片，
可跳过 Stage 5 图床上传）。

用法：
    python3 stage6_target_solve.py [--start-row 2] [--end-row 176] [--concurrency 1]
"""
import base64
import json
import os
import socket
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import argparse
from io import BytesIO

from openpyxl import load_workbook
from PIL import Image

from logger import get_logger
from config import cfg, column, prompt, secret, require_secrets

# ============ 接口配置：全部来自 config.yaml / .env，见 config.py ============
API_STYLE = str(cfg('providers.target_model.api_style', 'baidu_solve')).strip().lower()
INPUT_MODE = str(cfg('providers.target_model.input_mode', 'url')).strip().lower()
TIMEOUT = int(cfg('providers.target_model.timeout', 200))
MAX_RETRIES = int(cfg('providers.target_model.max_retries', 1))

_SUPPORTED_STYLES = ('baidu_solve', 'openai_compat')
if API_STYLE not in _SUPPORTED_STYLES:
    raise SystemExit(
        f"providers.target_model.api_style = '{API_STYLE}' 不受支持，"
        f"可选值：{', '.join(_SUPPORTED_STYLES)}"
    )
if INPUT_MODE not in ('url', 'image'):
    raise SystemExit(
        f"providers.target_model.input_mode = '{INPUT_MODE}' 不受支持，可选值：url, image"
    )

# ============ 表格配置 ============
WORK_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(WORK_DIR, cfg('paths.excel_file', 'evaluation.xlsx'))
COL_B = column('image_path')     # B列 - 本地图片路径（input_mode=image 时使用）
COL_Z = column('confidence')     # Z列 - confidence
COL_AE = column('image_url')     # AE列 - 公网图片 URL（input_mode=url 时使用）
COL_AF = column('model_output')  # AF列 - 输出文本
COL_AG = column('model_time')    # AG列 - 待测模型请求耗时（秒）

# ============ 并发控制 ============
DEFAULT_CONCURRENCY = int(cfg('providers.target_model.concurrency', 1))
SEMAPHORE = None
EXCEL_LOCK = threading.Lock()

_MIME_MAP = {
    '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png',
    '.bmp': 'image/bmp', '.webp': 'image/webp', '.gif': 'image/gif',
}


def required_secret_paths():
    """当前 api_style 必须配齐的凭证路径。"""
    if API_STYLE == 'baidu_solve':
        return (
            'providers.target_model.baidu_solve.api_key',
            'providers.target_model.baidu_solve.secret_key',
        )
    return (
        'providers.target_model.openai_compat.base_url',
        'providers.target_model.openai_compat.api_key',
        'providers.target_model.openai_compat.model',
    )


def read_image_base64(image_path):
    """读取本地图片，返回 (纯 base64, mime_type)。

    超过 runtime.max_image_bytes 时统一转 JPEG 并逐步压缩，压缩策略与
    Stage 4 / Stage 7 保持一致，避免大图被接口按体积拒绝。
    """
    if not os.path.isabs(image_path):
        image_path = os.path.join(WORK_DIR, image_path)
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"图片文件不存在: {image_path}")

    ext = os.path.splitext(image_path)[1].lower()
    mime_type = _MIME_MAP.get(ext, 'image/png')
    max_bytes = int(cfg('runtime.max_image_bytes', 5 * 1024 * 1024))

    with open(image_path, 'rb') as f:
        raw = f.read()
    if len(raw) <= max_bytes:
        return base64.b64encode(raw).decode('utf-8'), mime_type

    img = Image.open(image_path)
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')

    quality, scale, compressed, data = 85, 1.0, None, raw
    for _ in range(10):
        resized = img.resize(
            (max(1, int(img.width * scale)), max(1, int(img.height * scale))), Image.LANCZOS
        ) if scale < 1.0 else img
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

    return base64.b64encode(compressed or data).decode('utf-8'), 'image/jpeg'


def http_post_json(url, payload, headers, timeout=None):
    """发一个 JSON POST，返回解析后的 dict。"""
    body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    req = urllib.request.Request(url=url, data=body, method='POST')
    for key, value in headers.items():
        req.add_header(key, value)
    with urllib.request.urlopen(req, timeout=timeout or TIMEOUT) as resp:
        if resp.status != 200:
            raise RuntimeError(f"HTTP 状态码异常: {resp.status}")
        return json.loads(resp.read().decode('utf-8', 'replace'))


# ============================================================
# api_style = baidu_solve：百度智能云 OCR「拍照解题」
# ============================================================
_TOKEN_LOCK = threading.Lock()
_TOKEN_CACHE = {"value": None, "expires_at": 0.0}


def get_access_token():
    """用 API Key / Secret Key 换 access_token，进程内缓存（官方有效期 30 天）。"""
    with _TOKEN_LOCK:
        if _TOKEN_CACHE["value"] and time.time() < _TOKEN_CACHE["expires_at"]:
            return _TOKEN_CACHE["value"]

        params = urllib.parse.urlencode({
            "grant_type": "client_credentials",
            "client_id": secret('providers.target_model.baidu_solve.api_key'),
            "client_secret": secret('providers.target_model.baidu_solve.secret_key'),
        })
        token_url = cfg('providers.target_model.baidu_solve.token_url')
        req = urllib.request.Request(url=f"{token_url}?{params}", data=b"", method='POST')
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode('utf-8', 'replace'))

        token = data.get("access_token")
        if not token:
            raise RuntimeError(f"获取 access_token 失败: {data}")
        # 留 1 小时余量，避免临界过期
        _TOKEN_CACHE["value"] = token
        _TOKEN_CACHE["expires_at"] = time.time() + max(int(data.get("expires_in", 2592000)) - 3600, 60)
        return token


def solve_baidu(image_url=None, image_path=None):
    """调用「拍照解题」接口，返回完整解题文本。"""
    payload = {
        "disable_preprocess": bool(cfg('providers.target_model.baidu_solve.disable_preprocess', False)),
        "stream_switch": bool(cfg('providers.target_model.baidu_solve.stream_switch', False)),
    }
    if INPUT_MODE == 'image':
        payload["image"] = read_image_base64(image_path)[0]
    else:
        payload["url"] = image_url

    solve_url = cfg('providers.target_model.baidu_solve.solve_url')
    url = f"{solve_url}?access_token={urllib.parse.quote(get_access_token())}"
    data = http_post_json(url, payload, {"Content-Type": "application/json"})

    if data.get("error_code"):
        raise RuntimeError(
            f"接口返回错误: error_code={data.get('error_code')}, "
            f"error_msg={data.get('error_msg')}, log_id={data.get('log_id')}"
        )

    result = data.get("result") or {}
    content = result.get("content")
    if not content:
        # stream_switch=true 时只有 data_block，需要自己拼接
        content = "".join(
            block.get("answer_content", "") for block in (result.get("data_block") or [])
        )
    content = (content or "").strip()
    if not content:
        raise RuntimeError(f"接口返回为空: log_id={data.get('log_id')}")
    return content


# ============================================================
# api_style = openai_compat：任意 OpenAI 兼容多模态 chat 接口
# ============================================================
def solve_openai_compat(image_url=None, image_path=None):
    """调用 OpenAI 兼容多模态接口，返回完整解题文本。"""
    if INPUT_MODE == 'image':
        raw, mime = read_image_base64(image_path)
        image_ref = f"data:{mime};base64,{raw}"
    else:
        image_ref = image_url

    payload = {
        "model": cfg('providers.target_model.openai_compat.model'),
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt('target_solve')},
                {"type": "image_url", "image_url": {"url": image_ref}},
            ],
        }],
    }
    temperature = cfg('providers.target_model.openai_compat.temperature', None)
    if temperature is not None:
        payload["temperature"] = temperature
    if cfg('providers.target_model.openai_compat.enable_thinking', None) is not None:
        payload["enable_thinking"] = bool(
            cfg('providers.target_model.openai_compat.enable_thinking')
        )

    api_key = secret('providers.target_model.openai_compat.api_key')
    headers = {
        "Content-Type": "application/json",
        "Authorization": api_key if api_key.startswith('Bearer ') else f"Bearer {api_key}",
    }
    data = http_post_json(
        cfg('providers.target_model.openai_compat.base_url'), payload, headers
    )

    if data.get("error"):
        raise RuntimeError(f"接口返回错误: {data['error']}")
    choices = data.get("choices") or []
    if not choices:
        raise RuntimeError(f"接口返回缺少 choices: {str(data)[:300]}")
    content = ((choices[0].get("message") or {}).get("content") or "").strip()
    if not content:
        raise RuntimeError("接口返回为空")
    return content


_SOLVERS = {
    'baidu_solve': solve_baidu,
    'openai_compat': solve_openai_compat,
}


def solve_with_retry(image_url=None, image_path=None, max_retries=MAX_RETRIES):
    """按模型调用粒度执行：首次调用 + 最多 max_retries 次重试，每次限时 TIMEOUT 秒。"""
    solver = _SOLVERS[API_STYLE]
    last_error = None
    for attempt in range(1, max_retries + 2):
        try:
            return solver(image_url=image_url, image_path=image_path)
        except Exception as exc:
            last_error = exc
            if attempt <= max_retries:
                print(f"    待测模型解题重试 {attempt}/{max_retries}: {str(exc)[:120]}")
                time.sleep(2)
                continue
            raise last_error


def process_row(wb, ws, row_idx, stats):
    """处理单行"""
    confidence = ws.cell(row=row_idx, column=COL_Z).value
    if str(confidence).strip().lower() != "high":
        stats["skipped"] += 1
        print(f"[row {row_idx}] 跳过: confidence不是high")
        return

    if INPUT_MODE == 'image':
        source_col, source_name = COL_B, "B列(本地图片路径)"
    else:
        source_col, source_name = COL_AE, "AE列(图片URL)"
    source = ws.cell(row=row_idx, column=source_col).value

    if not source:
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value=f"ERROR: {source_name}为空")
            ws.cell(row=row_idx, column=COL_AG, value=None)
            wb.save(XLSX_PATH)
        stats["failed"] += 1
        print(f"[row {row_idx}] 失败: {source_name}为空")
        return

    source = str(source).strip()
    kwargs = {"image_path": source} if INPUT_MODE == 'image' else {"image_url": source}
    print(f"[row {row_idx}] start")

    SEMAPHORE.acquire()
    t0 = time.time()
    try:
        result = solve_with_retry(**kwargs)
        elapsed = round(time.time() - t0, 3)
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value=result)
            ws.cell(row=row_idx, column=COL_AG, value=elapsed)
            wb.save(XLSX_PATH)
        stats["success"] += 1
        print(f"[row {row_idx}] 成功, {elapsed}s, len={len(result)}")

    except urllib.error.HTTPError as e:
        elapsed = round(time.time() - t0, 3)
        detail = e.read().decode("utf-8", "replace")
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value=f"ERROR: HTTP错误 {e.code}: {detail}")
            ws.cell(row=row_idx, column=COL_AG, value=elapsed)
            wb.save(XLSX_PATH)
        stats["failed"] += 1
        print(f"[row {row_idx}] 失败: HTTP错误 {e.code}")

    except urllib.error.URLError as e:
        elapsed = round(time.time() - t0, 3)
        reason = str(e.reason)
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value=f"ERROR: 网络错误: {reason}")
            ws.cell(row=row_idx, column=COL_AG, value=elapsed)
            wb.save(XLSX_PATH)
        stats["failed"] += 1
        print(f"[row {row_idx}] 失败: 网络错误")

    except (socket.timeout, TimeoutError):
        elapsed = round(time.time() - t0, 3)
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value=f"ERROR: 请求超时({TIMEOUT}秒)")
            ws.cell(row=row_idx, column=COL_AG, value=elapsed)
            wb.save(XLSX_PATH)
        stats["failed"] += 1
        print(f"[row {row_idx}] 失败: 请求超时({TIMEOUT}秒)")

    except Exception as e:
        elapsed = round(time.time() - t0, 3)
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value=f"ERROR: {str(e)}")
            ws.cell(row=row_idx, column=COL_AG, value=elapsed)
            wb.save(XLSX_PATH)
        stats["failed"] += 1
        print(f"[row {row_idx}] 失败: {str(e)[:120]}")

    finally:
        SEMAPHORE.release()


def main():
    global SEMAPHORE, WORK_DIR, XLSX_PATH

    parser = argparse.ArgumentParser(description="待测模型解题批量调用（AE/B->AF）")
    parser.add_argument("--start-row", type=int, default=2, help="起始行号，默认2")
    parser.add_argument("--end-row", type=int, default=None, help="结束行号，默认到最后一行")
    parser.add_argument("--concurrency", type=int, default=DEFAULT_CONCURRENCY, help=f"并发数，默认{DEFAULT_CONCURRENCY}")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    args = parser.parse_args()

    require_secrets(*required_secret_paths())

    WORK_DIR = os.path.abspath(args.work_dir or WORK_DIR)
    os.environ["AUTO_EVAL_WORK_DIR"] = WORK_DIR
    XLSX_PATH = os.path.join(WORK_DIR, cfg('paths.excel_file', 'evaluation.xlsx'))
    SEMAPHORE = threading.Semaphore(args.concurrency)

    wb = load_workbook(XLSX_PATH)
    ws = wb.active
    ws.cell(row=1, column=COL_AG, value="time")

    end_row = args.end_row if args.end_row else ws.max_row
    rows_to_process = list(range(args.start_row, end_row + 1))

    total = len(rows_to_process)

    logger = get_logger(base_dir=WORK_DIR)

    with logger.stage("target_solve", desc="待测模型解题", stage_order="06") as stage:
        stage.log_file_input(XLSX_PATH)
        stage.info(
            f"接口风格={API_STYLE}, 输入方式={INPUT_MODE}, "
            f"处理范围: 第{args.start_row}行 到 第{end_row}行，共 {total} 行，并发数={args.concurrency}"
        )

        stats = {"success": 0, "failed": 0, "skipped": 0}

        threads = []
        for row_idx in rows_to_process:
            t = threading.Thread(target=process_row, args=(wb, ws, row_idx, stats), daemon=False)
            t.start()
            threads.append(t)

        for t in threads:
            t.join()

        wb.save(XLSX_PATH)

        stage.log_file_output(XLSX_PATH)
        stage.set_row_count(total)
        stage.info(f"完成: total={total}, success={stats['success']}, failed={stats['failed']}, skipped={stats['skipped']}；单行失败已写入 AF/AG，不影响流水线继续")

    logger.finish()


if __name__ == "__main__":
    main()

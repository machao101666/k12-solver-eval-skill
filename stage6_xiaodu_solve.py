#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小度解题接口批量调用脚本（AE -> AF）
=====================================
读取 evaluation.xlsx 中 AE 列的图片 URL，
调用小度解题接口，结果写入 AF 列。

用法：
    python3 stage6_xiaodu_solve.py [--start-row 2] [--end-row 176] [--concurrency 1]
"""
import json
import time
import uuid
import hmac
import hashlib
import os
import urllib.request
import urllib.error
import socket
import threading
import argparse
import sys
from openpyxl import load_workbook

from logger import get_logger
import config

# ============ 接口凭证 ============
# 小度作业接口 App Key / Secret：从环境变量 / .env 读取，源码不含真实密钥
APP_KEY = config.require("XIAODU_APP_KEY", "小度待测模型解题")
APP_SECRET = config.require("XIAODU_APP_SECRET", "小度待测模型解题")

# ============ 接口配置 ============
HOST = config.get("XIAODU_HOST", "xiaodu.baidu.com")
URI = config.get("XIAODU_URI", "/edu-gateway/gateway/openapi/homework-tob/solve")
CONTENT_TYPE = "application/json"
METHOD = "POST"
TIMEOUT = 200
MAX_RETRIES = 1

# ============ 表格配置 ============
WORK_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(WORK_DIR, "evaluation.xlsx")
COL_Z = 26    # Z列 - confidence
COL_AE = 31   # AE列 - 输入URL
COL_AF = 32   # AF列 - 输出文本
COL_AG = 33   # AG列 - 小度请求耗时（秒）

# ============ 并发控制 ============
DEFAULT_CONCURRENCY = 1
SEMAPHORE = None
EXCEL_LOCK = threading.Lock()


def _sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _hmac_sha256_hex(key: str, msg: str) -> str:
    return hmac.new(key.encode("utf-8"), msg.encode("utf-8"), hashlib.sha256).hexdigest()


def build_authorization(body: bytes, timestamp: str) -> str:
    hashed_payload = _sha256_hex(body)
    canonical_headers = f"content-type:{CONTENT_TYPE}\nhost:{HOST}\n"
    canonical_request = f"{METHOD}\n{URI}\n\n{canonical_headers}\n{hashed_payload}"
    string_to_sign = f"{timestamp}\n{canonical_request}"
    return _hmac_sha256_hex(APP_SECRET, string_to_sign)


def solve_once(image_url: str) -> str:
    """调用小度解题接口，返回完整解题文本"""
    payload = {"imageUrl": image_url, "logid": str(int(time.time()))}
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")

    timestamp = str(int(time.time()))
    request_id = str(uuid.uuid4())
    authorization = build_authorization(body, timestamp)

    url = f"https://{HOST}{URI}"
    headers = {
        "Appkey": APP_KEY,
        "Content-Type": CONTENT_TYPE,
        "Authorization": authorization,
        "Timestamp": timestamp,
        "requestId": request_id,
        "Accept": "text/event-stream",
    }

    req = urllib.request.Request(url=url, data=body, method=METHOD)
    for k, v in headers.items():
        req.add_header(k, v)

    full_content = []
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        if resp.status != 200:
            raise RuntimeError(f"HTTP 状态码异常: {resp.status}")
        for raw in resp:
            line = raw.decode("utf-8", "replace").strip()
            if not line or not line.startswith("data:"):
                continue
            data_str = line[len("data:"):].strip()
            if not data_str:
                continue
            try:
                obj = json.loads(data_str)
            except json.JSONDecodeError:
                continue

            if "code" in obj and obj.get("code") not in (0, None):
                raise RuntimeError(
                    f"接口返回错误: code={obj.get('code')}, message={obj.get('message')}, msg={obj.get('msg')}"
                )

            piece = obj.get("content", "")
            if piece:
                full_content.append(piece)
            if obj.get("done") is True:
                break

    result = "".join(full_content).strip()
    if not result:
        raise RuntimeError("接口返回为空")
    return result


def solve_with_retry(image_url: str, max_retries: int = MAX_RETRIES) -> str:
    """按模型调用粒度执行：首次调用 + 最多 1 次重试，每次限时 TIMEOUT 秒。"""
    last_error = None
    for attempt in range(1, max_retries + 2):
        try:
            return solve_once(image_url)
        except Exception as exc:
            last_error = exc
            if attempt <= max_retries:
                print(f"    小度解题重试 {attempt}/{max_retries}: {str(exc)[:120]}")
                time.sleep(2)
                continue
            raise last_error


def process_row(wb, ws, row_idx: int, stats: dict):
    """处理单行"""
    confidence = ws.cell(row=row_idx, column=COL_Z).value
    if str(confidence).strip().lower() != "high":
        stats["skipped"] += 1
        print(f"[row {row_idx}] 跳过: confidence不是high")
        return

    image_url = ws.cell(row=row_idx, column=COL_AE).value

    if not image_url:
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value="ERROR: AE列为空")
            ws.cell(row=row_idx, column=COL_AG, value=None)
            wb.save(XLSX_PATH)
        stats["failed"] += 1
        print(f"[row {row_idx}] 失败: AE列为空")
        return

    image_url = str(image_url).strip()
    print(f"[row {row_idx}] start")

    SEMAPHORE.acquire()
    try:
        t0 = time.time()
        result = solve_with_retry(image_url)
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

    except socket.timeout:
        elapsed = round(time.time() - t0, 3)
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=COL_AF, value=f"ERROR: 请求超时({TIMEOUT}秒)")
            ws.cell(row=row_idx, column=COL_AG, value=elapsed)
            wb.save(XLSX_PATH)
        stats["failed"] += 1
        print(f"[row {row_idx}] 失败: 请求超时({TIMEOUT}秒)")

    except TimeoutError:
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

    parser = argparse.ArgumentParser(description="小度解题接口批量调用（AE->AF）")
    parser.add_argument("--start-row", type=int, default=2, help="起始行号，默认2")
    parser.add_argument("--end-row", type=int, default=None, help="结束行号，默认到最后一行")
    parser.add_argument("--concurrency", type=int, default=DEFAULT_CONCURRENCY, help=f"并发数，默认{DEFAULT_CONCURRENCY}")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    args = parser.parse_args()

    WORK_DIR = os.path.abspath(args.work_dir or WORK_DIR)
    os.environ["AUTO_EVAL_WORK_DIR"] = WORK_DIR
    XLSX_PATH = os.path.join(WORK_DIR, "evaluation.xlsx")
    SEMAPHORE = threading.Semaphore(args.concurrency)

    wb = load_workbook(XLSX_PATH)
    ws = wb.active
    ws.cell(row=1, column=COL_AG, value="time")

    end_row = args.end_row if args.end_row else ws.max_row
    rows_to_process = list(range(args.start_row, end_row + 1))

    total = len(rows_to_process)

    logger = get_logger(base_dir=WORK_DIR)

    with logger.stage("xiaodu", desc="小度待测模型解题", stage_order="06") as stage:
        stage.log_file_input(XLSX_PATH)
        stage.info(f"处理范围: 第{args.start_row}行 到 第{end_row}行，共 {total} 行，并发数={args.concurrency}")

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

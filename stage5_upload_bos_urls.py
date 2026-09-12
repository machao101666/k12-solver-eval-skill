# -*- coding: utf-8 -*-
"""
将本地题图上传到百度智能云 BOS，返回公开访问 URL，可选写入 evaluation.xlsx 的 AE 列。

策略：只使用 BOS SDK。SDK 不可用、上传失败、上传后对象校验失败或 AE 写入失败时，Stage 5 直接失败退出，
不使用 dodo_cli 或任何手动降级方案。
"""

import argparse
import os
import sys
from datetime import datetime
from urllib.parse import quote

from openpyxl import load_workbook

from logger import get_logger
import config

# ============ 配置区 ============
# BOS 访问密钥：从环境变量 / .env 读取，源码不含真实密钥
AK = config.require("BOS_ACCESS_KEY", "百度 BOS 上传")
SK = config.require("BOS_SECRET_KEY", "百度 BOS 上传")
ENDPOINT = config.get("BOS_ENDPOINT", "bj.bcebos.com")   # 北京区域，按需改为 gz / su
BUCKET = config.get("BOS_BUCKET", "machao")
BOS_ROOT_NAME = "image_ocr_clean"  # 本地根目录名，上传时替换成批次时间
WORK_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(WORK_DIR, "evaluation.xlsx")
COL_IMAGE_PATH = 2   # B列
COL_CONFIDENCE = 26  # Z列
COL_AE_URL = 31      # AE列
# ================================


class UploadError(RuntimeError):
    pass


def get_bos_client():
    try:
        from baidubce.bce_client_configuration import BceClientConfiguration
        from baidubce.auth.bce_credentials import BceCredentials
        from baidubce.services.bos.bos_client import BosClient
    except Exception as exc:
        raise UploadError(
            "BOS SDK 不可用，请先执行 pip3 install -r requirements.txt，"
            f"确认 bce-python-sdk 安装成功。原始错误: {exc}"
        ) from exc

    config = BceClientConfiguration(
        credentials=BceCredentials(AK, SK),
        endpoint=ENDPOINT,
    )
    return BosClient(config)


def build_object_key(local_path, upload_batch_dir):
    normalized = os.path.normpath(local_path)
    parts = normalized.split(os.sep)

    if BOS_ROOT_NAME in parts:
        root_index = parts.index(BOS_ROOT_NAME)
        relative_parts = parts[root_index + 1:]
    else:
        relative_parts = [os.path.basename(local_path)]

    object_parts = [upload_batch_dir] + relative_parts
    return "/".join(object_parts)


def build_public_url(object_key):
    encoded_key = "/".join(quote(part) for part in object_key.split("/"))
    return f"https://{BUCKET}.{ENDPOINT}/{encoded_key}"


def iter_list_objects(response):
    contents = getattr(response, "contents", None)
    if contents is None and isinstance(response, dict):
        contents = response.get("contents") or response.get("Contents")
    return contents or []


def object_key_from_list_item(item):
    if isinstance(item, dict):
        return item.get("key") or item.get("Key")
    return getattr(item, "key", None) or getattr(item, "Key", None)


def assert_object_exists(client, object_key):
    try:
        response = client.list_objects(BUCKET, prefix=object_key, max_keys=1000)
    except TypeError:
        response = client.list_objects(BUCKET, prefix=object_key)
    except Exception as exc:
        raise UploadError(f"BOS 上传后 list_objects 校验失败: {object_key}; 原始错误: {exc}") from exc

    matched = any(object_key_from_list_item(item) == object_key for item in iter_list_objects(response))
    if not matched:
        raise UploadError(f"BOS 上传后校验失败，list_objects 未找到 object_key: {object_key}")


def upload_with_bos(client, local_path, upload_batch_dir):
    object_key = build_object_key(local_path, upload_batch_dir)
    client.put_object_from_file(BUCKET, object_key, local_path)
    assert_object_exists(client, object_key)
    return object_key, build_public_url(object_key)


def collect_files(target):
    if os.path.isdir(target):
        files = []
        for root, _, filenames in os.walk(target):
            for filename in filenames:
                if filename.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                    files.append(os.path.join(root, filename))
        return sorted(files)
    return [target]


def normalize_path(path):
    return os.path.normpath(str(path)).replace(os.sep, "/")


def write_urls_to_excel(upload_results, xlsx_path=XLSX_PATH, only_high=True):
    wb = load_workbook(xlsx_path)
    ws = wb.active

    by_full_path = {normalize_path(path): url for path, _, url in upload_results}
    by_basename = {os.path.basename(path): url for path, _, url in upload_results}

    written = 0
    skipped = 0
    missing = []
    target_high_rows = 0
    for row in range(2, ws.max_row + 1):
        if only_high:
            confidence = ws.cell(row=row, column=COL_CONFIDENCE).value
            if str(confidence).strip().lower() != "high":
                skipped += 1
                continue
            target_high_rows += 1

        image_path = ws.cell(row=row, column=COL_IMAGE_PATH).value
        if not image_path:
            continue

        normalized = normalize_path(image_path)
        url = by_full_path.get(normalized) or by_basename.get(os.path.basename(normalized))
        if not url:
            missing.append((row, image_path))
            continue

        ws.cell(row=row, column=COL_AE_URL, value=url)
        written += 1

    wb.save(xlsx_path)
    return written, skipped, missing, target_high_rows


def upload_files(files, upload_batch_dir, stage):
    if not files:
        raise UploadError("没有找到可上传的图片文件")

    client = get_bos_client()
    upload_results = []
    for file_path in files:
        object_key, url = upload_with_bos(client, file_path, upload_batch_dir)
        upload_results.append((file_path, object_key, url))
        stage.log_file_output(url)
        stage.log_item_result(file_path, "uploaded", object_key)
    return upload_results


def main():
    parser = argparse.ArgumentParser(description="上传本地图片到 BOS，并可选写入 evaluation.xlsx 的 AE 列")
    parser.add_argument("target", help="图片路径或目录")
    parser.add_argument("--write-excel", action="store_true", help="将上传 URL 自动写入 evaluation.xlsx 的 AE 列")
    parser.add_argument("--all-rows", action="store_true", help="写入所有匹配行；默认只写 confidence=high 的行")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    parser.add_argument("--xlsx", default=None, help="Excel 路径，默认 <work-dir>/evaluation.xlsx")
    args = parser.parse_args()

    work_dir = os.path.abspath(args.work_dir or WORK_DIR)
    os.environ["AUTO_EVAL_WORK_DIR"] = work_dir
    xlsx_path = args.xlsx or os.path.join(work_dir, "evaluation.xlsx")
    if not os.path.isabs(xlsx_path):
        xlsx_path = os.path.abspath(xlsx_path)
    target = args.target
    if not os.path.isabs(target):
        target = os.path.join(work_dir, target)

    logger = get_logger(base_dir=work_dir)

    try:
        with logger.stage("upload_bos", desc="本地题图上传 BOS 并生成 URL", stage_order="05") as stage:
            stage.log_file_input(target)
            if args.write_excel:
                stage.log_file_input(xlsx_path)

            # 只生成一次上传批次时间；object_key 和 URL 均来自同一个 object_key，避免上传和读取时间不一致。
            upload_batch_dir = datetime.now().strftime("%Y%m%d_%H%M%S")
            files = collect_files(target)

            print(f"upload_batch_dir\t{upload_batch_dir}")
            print("upload_mode\tbos")
            stage.info(f"BOS 上传批次目录: {upload_batch_dir}")
            stage.info(f"待上传文件数: {len(files)}")
            stage.info("上传策略: 仅使用 BOS，失败即停止，不使用 dodo 降级")

            upload_results = upload_files(files, upload_batch_dir, stage)

            for file_path, object_key, url in upload_results:
                print(f"{file_path}\t{object_key}\t{url}")

            if args.write_excel:
                written, skipped, missing, target_high_rows = write_urls_to_excel(
                    upload_results,
                    xlsx_path,
                    only_high=not args.all_rows,
                )
                print(f"excel_written\t{written}")
                print(f"excel_skipped_non_high\t{skipped}")
                stage.info(
                    f"Excel 写入 AE 列: written={written}, skipped_non_high={skipped}, "
                    f"missing={len(missing)}, target_high_rows={target_high_rows}"
                )
                stage.log_file_output(xlsx_path)
                for row, image_path in missing:
                    print(f"excel_missing\trow={row}\t{image_path}")
                    stage.warn(f"Excel 未匹配: row={row}, image={image_path}")

                if target_high_rows > 0 and written == 0:
                    raise UploadError("存在 confidence=high 行，但 AE 列没有成功写入任何 URL")

            stage.set_row_count(len(files))

        logger.finish()
    except Exception as exc:
        logger.finish()
        print(f"[ERROR] Stage 5 BOS 上传失败，流程停止: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

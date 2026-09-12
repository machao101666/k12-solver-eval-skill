# -*- coding: utf-8 -*-
"""
百度OCR多模态试卷切题 - 批量脚本
输入：图片URL
流程：
  1. 下载图片到 image/ 目录，自动编号命名（1.jpg, 2.jpg...）
  2. 调用百度OCR识别并切割题目
  3. 切割后的题目保存到 image_ocr/ 目录，命名 {原图编号}-{序号}.png
"""

import os
import re
import sys
import csv
import time
import base64
import json
import argparse

import requests
from PIL import Image
from io import BytesIO

from logger import get_logger
import config

# ============================================
# 凭证配置：从环境变量 / .env 读取，源码不含真实密钥
# ============================================
# 百度 OCR 应用的 API Key / Secret Key（client_credentials 换取 access_token）
API_KEY = config.get('BAIDU_OCR_API_KEY', '')
SECRET_KEY = config.get('BAIDU_OCR_SECRET_KEY', '')

# API 地址保留默认值，可用同名环境变量覆盖
PAPER_CUT_VLM_CREATE_URL = config.get('BAIDU_OCR_CREATE_URL', 'https://aip.baidubce.com/rest/2.0/ocr/v1/paper_cut_edu_vlm/create_task')
PAPER_CUT_VLM_RESULT_URL = config.get('BAIDU_OCR_RESULT_URL', 'https://aip.baidubce.com/rest/2.0/ocr/v1/paper_cut_edu_vlm/get_result')

# Access Token 缓存
access_token = None
token_expires_at = 0


def is_api_error(result):
    """判断百度 API 返回是否为错误；无 error_code 表示非错误响应。"""
    error_code = result.get('error_code')
    return error_code not in (None, 0, '0')


def get_next_filename(folder="image"):
    """获取 image 文件夹中下一个可用的编号文件名"""
    if not os.path.exists(folder):
        os.makedirs(folder)
        return os.path.join(folder, "1.jpg")

    existing = os.listdir(folder)
    numbers = []
    for name in existing:
        match = re.match(r"^(\d+)\.", name)
        if match:
            numbers.append(int(match.group(1)))

    next_num = max(numbers, default=0) + 1
    return os.path.join(folder, f"{next_num}.jpg")


def download_image(url, folder="image"):
    """下载图片到指定文件夹，自动编号命名"""
    filepath = get_next_filename(folder)
    response = requests.get(url, timeout=30)
    response.raise_for_status()

    with open(filepath, "wb") as f:
        f.write(response.content)

    print(f"[INFO] 已下载: {filepath}")
    return filepath


def get_access_token():
    """获取百度Access Token（从 app.py 复用）"""
    global access_token, token_expires_at

    if access_token and time.time() < token_expires_at:
        return access_token

    auth_url = 'https://aip.baidubce.com/oauth/2.0/token'
    params = {
        'grant_type': 'client_credentials',
        'client_id': config.require('BAIDU_OCR_API_KEY', '百度 OCR 切题'),
        'client_secret': config.require('BAIDU_OCR_SECRET_KEY', '百度 OCR 切题')
    }

    try:
        response = requests.post(auth_url, params=params)
        result = response.json()

        if 'access_token' in result:
            access_token = result['access_token']
            expires_in = result.get('expires_in', 2592000) - 300
            token_expires_at = time.time() + expires_in
            return access_token
        else:
            error_msg = result.get('error_description', 'Unknown error')
            raise Exception(f"获取Token失败: {error_msg}")

    except Exception as e:
        raise Exception(f"获取Access Token失败: {str(e)}")


def recognize_paper_cut_vlm_by_image(image_path):
    """
    调用百度OCR多模态试卷切题识别（通过本地图片）
    流程：读取本地图片 -> base64编码 -> 创建任务 -> 轮询查询结果
    """
    try:
        # 读取本地图片并转为base64
        with open(image_path, 'rb') as f:
            image_base64 = base64.b64encode(f.read()).decode('utf-8')

        token = get_access_token()

        # ========== 第一步：创建任务 ==========
        create_url = f"{PAPER_CUT_VLM_CREATE_URL}?access_token={token}"
        headers = {'Content-Type': 'application/json'}

        # only_split=true 走同步“仅切题”流程；该接口实际要求 JSON 布尔值
        data = json.dumps({
            'image': image_base64,
            'only_split': True,
            'scene_type': 'paper'
        }, ensure_ascii=False)

        # 创建任务（带QPS超限重试）
        max_create_retries = 3
        for create_attempt in range(max_create_retries):
            response = requests.post(create_url, data=data, headers=headers, timeout=30)
            response.raise_for_status()
            create_result = response.json()

            if create_result.get('error_code') == 18:
                if create_attempt < max_create_retries - 1:
                    time.sleep(1)
                    continue
                else:
                    return create_result

            if is_api_error(create_result):
                return create_result

            sync_result = create_result.get('result', {})
            if isinstance(sync_result, dict) and (
                sync_result.get('qus_result') is not None
                or sync_result.get('qus_result_num') is not None
                or sync_result.get('result')
            ):
                print("[INFO] 同步切题任务执行成功")
                return create_result

            break

        task_id = create_result.get('result', {}).get('task_id')
        if not task_id:
            return {'error': '创建任务失败，未返回task_id', 'error_code': -1}

        print(f"[INFO] 任务已创建，task_id: {task_id}")

        # ========== 第二步：轮询查询结果 ==========
        result_url = f"{PAPER_CUT_VLM_RESULT_URL}?access_token={token}"
        result_headers = {'Content-Type': 'application/x-www-form-urlencoded'}

        query_attempts = 3
        max_retries = 30
        retry_interval = 3

        for query_attempt in range(1, query_attempts + 1):
            # 文档建议：提交请求后5～10秒开始轮询
            print(f"[INFO] 等待任务处理中（约10秒），查询尝试 {query_attempt}/{query_attempts}...")
            time.sleep(10)
            consecutive_qps_errors = 0

            for i in range(max_retries):
                result_data = {'task_id': str(task_id)}
                response = requests.post(result_url, data=result_data, headers=result_headers, timeout=30)
                response.raise_for_status()
                query_result = response.json()

                # QPS超限处理
                if query_result.get('error_code') == 18:
                    consecutive_qps_errors += 1
                    wait_time = min(consecutive_qps_errors * 1, 5)
                    time.sleep(wait_time)
                    continue
                else:
                    consecutive_qps_errors = 0

                result_info = query_result.get('result', {})
                status = str(result_info.get('status', '')).lower()

                if status in ('running', 'processing', 'pending') or query_result.get('error_code') in (60, '60'):
                    print(f"[INFO] 任务处理中... ({i + 1}/{max_retries})")
                    time.sleep(retry_interval)
                    continue

                # 官方说明：216630 recognize error 可再次请求
                if query_result.get('error_code') in (216630, '216630'):
                    if query_attempt < query_attempts:
                        print(f"[WARN] OCR识别错误(216630)，等待后重试查询: {query_attempt}/{query_attempts}")
                        time.sleep(5 * query_attempt)
                        break
                    return query_result

                # 其他错误
                if is_api_error(query_result):
                    return query_result

                if status == 'success':
                    print("[INFO] 任务执行成功")
                    return query_result

                if status == 'failed':
                    return {'error': '任务执行失败', 'error_code': -1, 'task_status': status}

                # 还在处理中
                print(f"[INFO] 任务处理中... ({i + 1}/{max_retries})")
                time.sleep(retry_interval)

        return {'error': '轮询超时', 'error_code': -1, 'task_status': 'timeout'}

    except Exception as e:
        return {'error': str(e), 'error_code': -1}


def cut_and_save_questions(image_path, result_json, base_name):
    """
    根据API返回的坐标信息，从本地图片中切割题目区域并保存
    coordinate 格式: [x1, y1, x2, y2]
    保存到 image_ocr/ 目录，命名规则: {base_name}-{序号}.png
    """
    output_dir = 'image_ocr'
    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"[INFO] 创建输出目录: {output_dir}/")

    # 获取题目结果，兼容百度 OCR 不同返回结构
    result_data = result_json.get('result', {})
    if isinstance(result_data, dict) and 'result' in result_data:
        # 异步任务嵌套结构
        result_data = result_data.get('result', {})

    qus_result = result_data.get('qus_result') or result_data.get('question_list') or result_data.get('questions') or []
    if not qus_result:
        print("[WARN] 未识别到题目，无图片可切割")
        return 0

    # 加载本地原图
    print(f"[INFO] 正在加载原图: {image_path}")
    try:
        original_img = Image.open(image_path)
        print(f"[INFO] 原图尺寸: {original_img.size}")
    except Exception as e:
        print(f"[ERROR] 加载原图失败: {e}")
        return 0

    saved_count = 0

    for question in qus_result:
        qus_id = question.get('qus_id') or question.get('que_id') or '?'
        location = question.get('location', question)
        qus_locations = location.get('qus_location', [])

        if not qus_locations:
            print(f"[WARN] 第 {qus_id} 题无位置信息，跳过")
            continue

        # 获取第一个位置信息；coordinate 实测返回 [x1, y1, x2, y2]，PDF 示例字段为 {x,y,w,h}
        loc = qus_locations[0] if isinstance(qus_locations, list) else qus_locations
        coordinate = loc.get('coordinate') or loc.get('points') or []

        if coordinate and len(coordinate) >= 4:
            x1, y1, x2, y2 = coordinate[:4]
            left = int(x1)
            top = int(y1)
            right = int(x2)
            bottom = int(y2)
        elif all(k in loc for k in ('x', 'y', 'w', 'h')):
            left = int(loc['x'])
            top = int(loc['y'])
            right = int(loc['x'] + loc['w'])
            bottom = int(loc['y'] + loc['h'])
        elif all(k in loc for k in ('left_x', 'left_y', 'right_x', 'right_y')):
            left = int(loc['left_x'])
            top = int(loc['left_y'])
            right = int(loc['right_x'])
            bottom = int(loc['right_y'])
        else:
            print(f"[WARN] 第 {qus_id} 题坐标格式不正确: {coordinate or loc}")
            continue

        # 确保坐标不越界
        img_width, img_height = original_img.size
        left = max(0, left)
        top = max(0, top)
        right = min(img_width, right)
        bottom = min(img_height, bottom)

        if right <= left or bottom <= top:
            print(f"[WARN] 第 {qus_id} 题坐标无效，跳过")
            continue

        # 切割图片
        try:
            cropped = original_img.crop((left, top, right, bottom))

            # 文件名按 {base_name}-1, {base_name}-2... 累加
            saved_count += 1
            filename = f"{base_name}-{saved_count}.png"
            filepath = os.path.join(output_dir, filename)
            cropped.save(filepath, 'PNG')

            print(f"[INFO] 第 {qus_id} 题已保存: {filename} "
                  f"(坐标: [{left}, {top}, {right}, {bottom}], 尺寸: {cropped.size})")

        except Exception as e:
            print(f"[ERROR] 切割第 {qus_id} 题失败: {e}")

    return saved_count


def process_single_url(url):
    """
    处理单个 URL：下载 -> OCR识别 -> 切割保存
    返回切割保存的数量（>0 表示成功）
    """
    print("=" * 60)
    print("百度OCR 多模态试卷切题 - URL批量脚本")
    print("=" * 60)
    print(f"[INFO] 输入URL: {url}")

    # 步骤0：下载图片到 image/ 目录
    print("\n[步骤0/3] 下载图片到 image/ 目录...")
    try:
        image_path = download_image(url, folder="image")
    except Exception as e:
        print(f"[ERROR] 下载图片失败: {e}")
        return 0

    # 提取 base_name（如 "1" 从 "image/1.jpg"）
    base_name = os.path.splitext(os.path.basename(image_path))[0]
    print(f"[INFO] 图片编号: {base_name}")

    # 步骤1：调用API获取题目切分结果
    print("\n[步骤1/3] 调用API获取题目切分结果...")
    result = recognize_paper_cut_vlm_by_image(image_path)

    # 检查结果
    if 'error' in result or is_api_error(result):
        print(f"[ERROR] API调用失败: {result}")
        return 0

    result_data = result.get('result', {})
    if isinstance(result_data, dict) and 'result' in result_data:
        result_data = result_data.get('result', {})

    qus_result = result_data.get('qus_result') or result_data.get('question_list') or result_data.get('questions') or []
    print(f"[INFO] 共识别到 {len(qus_result)} 道题目")

    # 步骤2：加载本地图片并切割保存
    print("\n[步骤2/3] 加载原图并切割题目区域...")
    saved = cut_and_save_questions(image_path, result, base_name=base_name)

    print("\n" + "=" * 60)
    print(f"处理完成！共保存 {saved} 张切割图片到 image_ocr/ 目录")
    print("=" * 60)

    return saved


def main():
    """
    主函数
    用法: python stage1_ocr_cut_questions.py <CSV文件路径> [--work-dir <目录>]
    示例: python stage1_ocr_cut_questions.py urls.csv --work-dir runs/run_xxx
    """
    parser = argparse.ArgumentParser(description="OCR 试卷切题")
    parser.add_argument("csv_path", help="CSV 文件路径，第一列为图片 URL")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或当前目录")
    args = parser.parse_args()

    work_dir = os.path.abspath(args.work_dir or os.getcwd())
    os.environ["AUTO_EVAL_WORK_DIR"] = work_dir
    os.makedirs(work_dir, exist_ok=True)
    logger = get_logger(base_dir=work_dir)

    csv_path = args.csv_path.strip()
    if not os.path.isabs(csv_path):
        csv_path = os.path.abspath(csv_path)

    if not os.path.exists(csv_path):
        print(f"[ERROR] CSV文件不存在: {csv_path}")
        sys.exit(1)

    old_cwd = os.getcwd()
    os.chdir(work_dir)
    try:
        # 读取CSV全部内容
        rows = []
        with open(csv_path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                rows.append(row)

        total = len(rows)

        with logger.stage("batch", desc="OCR 试卷切题", stage_order="01") as stage:
            stage.log_file_input(csv_path)
            stage.info(f"共读取到 {total} 行数据，开始按顺序处理...")

            success_count = 0
            fail_count = 0

            for idx, row in enumerate(rows):
                if not row:
                    continue

                url = row[0].strip()
                if not url:
                    stage.warn(f"第 {idx + 1} 行 URL 为空，跳过")
                    continue

                stage.info(f"处理第 {idx + 1}/{total} 行")
                t0 = time.time()
                saved_count = process_single_url(url)
                elapsed = round(time.time() - t0, 1)

                if saved_count > 0:
                    success_count += 1
                    stage.log_item_result(f"row_{idx+1}", "success",
                                          f"切割 {saved_count} 张图片，耗时 {elapsed}s")
                else:
                    fail_count += 1
                    stage.log_item_result(f"row_{idx+1}", "failed",
                                          f"未识别到题目，耗时 {elapsed}s")

                stage.set_progress(idx + 1, total)

                if idx < total - 1:
                    time.sleep(3)

            # 不回写 CSV，CSV 只作为第一列 URL 输入源
            stage.log_file_output("image/")
            stage.log_file_output("image_ocr/")
            stage.set_row_count(total)
            stage.info(f"完成: 成功 {success_count}, 失败 {fail_count}")

        logger.finish()
    finally:
        os.chdir(old_cwd)


if __name__ == '__main__':
    main()

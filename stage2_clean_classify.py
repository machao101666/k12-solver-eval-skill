# -*- coding: utf-8 -*-
"""
批量处理image_ocr文件夹中的图片，调用千帆视觉理解API进行有效性判定与分类
"""
import argparse
import base64
import os
import shutil
import time
import requests
import sys
import signal
import openpyxl

from logger import get_logger
from config import cfg, secret, prompt


class TimeoutException(Exception):
    pass


def timeout_handler(signum, frame):
    raise TimeoutException()


def image_to_base64(image_path: str) -> str:
    """将本地图片转换为base64编码"""
    with open(image_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')


def print_progress(msg):
    """打印进度并实时刷新"""
    print(msg)
    sys.stdout.flush()


def check_question(image_path: str, timeout=None, max_retries=None) -> str:
    """调用视觉理解模型判定图片有效性与分类。

    接入点、模型、提示词、超时与重试次数均来自 config.yaml 的
    ``providers.qianfan`` / ``models.classify`` / ``prompts.files.clean_classify``。

    Args:
        image_path: 图片路径
        timeout: 单次请求超时（秒），None 表示用配置值
        max_retries: 重试次数，None 表示用配置值

    Returns:
        API返回的原始content内容
    """
    provider_name = cfg('models.classify.provider', 'qianfan')
    url = cfg(f'providers.{provider_name}.base_url')
    api_key = secret(f'providers.{provider_name}.api_key')
    if timeout is None:
        timeout = int(cfg('models.classify.timeout', 200))
    if max_retries is None:
        max_retries = int(cfg('models.classify.max_retries', 1))

    base64_image = image_to_base64(image_path)

    payload = {
        "model": cfg('models.classify.model'),
        "stream": False,
        "temperature": cfg('models.classify.temperature', 0.2),
        "enable_thinking": cfg('models.classify.enable_thinking', False),
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}"
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt("clean_classify")           }
                ]
            }
        ]
    }

    headers = {
        'Content-Type': 'application/json',
        'Authorization': api_key if api_key.startswith('Bearer ') else f'Bearer {api_key}',
    }

    last_error = ''
    for attempt in range(1, max_retries + 2):
        try:
            response = requests.request("POST", url, headers=headers, json=payload, timeout=timeout)
            result = response.json()

            if 'choices' in result and len(result['choices']) > 0:
                return result['choices'][0]['message']['content']

            if 'error' in result:
                last_error = f"API错误(第{attempt}次): {result['error']}"
                if attempt <= max_retries:
                    print(f"    重试 {attempt}/{max_retries}...")
                    time.sleep(2)
                    continue
                return last_error

            return str(result)
        except requests.exceptions.Timeout:
            last_error = f"请求超时(第{attempt}次)"
            if attempt <= max_retries:
                print(f"    重试 {attempt}/{max_retries}...")
                time.sleep(2)
                continue
            raise
        except requests.exceptions.RequestException as exc:
            last_error = f"网络错误(第{attempt}次): {exc}"
            if attempt <= max_retries:
                print(f"    重试 {attempt}/{max_retries}...")
                time.sleep(2)
                continue
            return last_error

    return last_error


def parse_result(content: str) -> dict:
    """解析模型返回的内容"""
    lines = content.strip().split('\n')
    result = {
        '有效性': None,
        '学段': None,
        '学科': None,
        '题型': None,
        '无效原因': None
    }

    for line in lines:
        if line.startswith('有效性判定：'):
            result['有效性'] = line.split('：')[1].strip()
        elif line.startswith('学段：'):
            result['学段'] = line.split('：')[1].strip()
        elif line.startswith('学科：'):
            result['学科'] = line.split('：')[1].strip()
        elif line.startswith('题型：'):
            result['题型'] = line.split('：')[1].strip()
        elif line.startswith('无效原因：'):
            result['无效原因'] = line.split('：')[1].strip()

    return result


def get_image_files(ocr_dir: str, start_filename: str = None) -> list:
    """获取image_ocr文件夹中所有png文件，按文件名排序，可指定起始文件"""
    files = []
    for f in os.listdir(ocr_dir):
        if f.endswith('.png'):
            files.append(f)
    files = sorted(files, key=lambda x: x.replace('-', ' ').replace('.png', ''))

    # 如果指定了起始文件，从该文件开始
    if start_filename:
        try:
            start_idx = files.index(start_filename)
            return files[start_idx:]
        except ValueError:
            print(f"未找到起始文件: {start_filename}")
            return files
    return files


def main():
    parser = argparse.ArgumentParser(description="有效性判定、分类并写入 evaluation.xlsx")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    args = parser.parse_args()

    work_dir = os.path.abspath(args.work_dir or os.path.dirname(os.path.abspath(__file__)))
    os.environ["AUTO_EVAL_WORK_DIR"] = work_dir
    logger = get_logger(base_dir=work_dir)

    old_cwd = os.getcwd()
    os.chdir(work_dir)
    try:
        ocr_dir = "image_ocr"
        output_dir = "image_ocr_clean"
        unused_dir = os.path.join(output_dir, "unused")

        # 创建输出目录
        os.makedirs(unused_dir, exist_ok=True)

        # 打开 evaluation.xlsx
        XLSX_PATH = os.path.join(work_dir, "evaluation.xlsx")
        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb.active
        next_row = ws.max_row + 1
        save_counter = 0

        # 获取所有图片文件
        files = get_image_files(ocr_dir, None)
        total_files = len(files)

        with logger.stage("clean", desc="有效性判定与分类", stage_order="02") as stage:
            stage.log_file_input(ocr_dir)
            stage.log_file_input(XLSX_PATH)
            stage.info(f"共找到 {total_files} 个图片文件")

            valid_count = 0
            invalid_count = 0
            timeout_count = 0
            parse_error_count = 0
            valid_stats = {}

            for i, filename in enumerate(files, 1):
                filepath = os.path.join(ocr_dir, filename)
                stage.info(f"[{i}/{total_files}] {filename}")

                try:
                    t0 = time.time()
                    content = check_question(filepath)
                    result = parse_result(content)
                    elapsed = round(time.time() - t0, 1)

                    stage.info(f"  判定={result['有效性']} 学段={result['学段']} 学科={result['学科']} 题型={result['题型']} 耗时={elapsed}s")
                    stage.log_api_call("ernie-5.0-thinking-latest", f"image={filename}", content[:800], elapsed)

                    if result['有效性'] == '2':
                        invalid_count += 1
                        reason = result['无效原因'] if result['无效原因'] else "未知原因"
                        reason = reason.replace('/', '-').replace('\\', '-').replace(':', '-')
                        new_name = f"{invalid_count}-{reason}.png"
                        new_path = os.path.join(unused_dir, new_name)
                        shutil.copy2(filepath, new_path)
                        stage.log_file_moved(filepath, new_path)
                        stage.log_item_result(filename, "invalid", reason)

                    elif result['有效性'] == '1':
                        valid_count += 1
                        grade = result['学段'] if result['学段'] else "无"
                        subject = result['学科'] if result['学科'] else "无"
                        qtype = result['题型'] if result['题型'] else "无"

                        type_dir = os.path.join(output_dir, grade, subject)
                        os.makedirs(type_dir, exist_ok=True)

                        new_name = f"{valid_count}-{qtype}.png"
                        new_path = os.path.join(type_dir, new_name)
                        shutil.copy2(filepath, new_path)
                        stage.log_file_moved(filepath, new_path)

                        # 写入 Excel
                        rel_path = os.path.relpath(new_path, os.path.dirname(XLSX_PATH))
                        row = next_row
                        next_row += 1
                        ws.cell(row=row, column=2, value=rel_path)
                        ws.cell(row=row, column=3, value=grade)
                        ws.cell(row=row, column=4, value=subject)
                        ws.cell(row=row, column=5, value=qtype)

                        stage.log_item_result(filename, "valid", f"{grade}/{subject}/{qtype}")

                        save_counter += 1
                        if save_counter % 10 == 0:
                            wb.save(XLSX_PATH)
                            stage.info(f"  Excel 已保存 {save_counter} 条")

                        # 统计
                        if grade not in valid_stats:
                            valid_stats[grade] = {}
                        if subject not in valid_stats[grade]:
                            valid_stats[grade][subject] = {}
                        if qtype not in valid_stats[grade][subject]:
                            valid_stats[grade][subject][qtype] = 0
                        valid_stats[grade][subject][qtype] += 1

                    else:
                        parse_error_count += 1
                        stage.warn(f"  无法解析判定结果: {filename}")
                        stage.log_item_result(filename, "parse_error", "")

                    stage.set_progress(i, total_files)

                except requests.exceptions.Timeout:
                    timeout_count += 1
                    stage.warn(f"  超时: {filename}")
                    stage.log_item_result(filename, "timeout", "")
                except Exception as e:
                    stage.error(f"  错误: {filename} - {e}")
                    stage.log_item_result(filename, "error", str(e))

            # 最终保存 Excel
            wb.save(XLSX_PATH)
            stage.log_file_output(XLSX_PATH)
            stage.log_file_output(output_dir)
            stage.set_row_count(save_counter)

            # 统计摘要
            for grade in sorted(valid_stats.keys()):
                for subject in sorted(valid_stats[grade].keys()):
                    for qtype, count in sorted(valid_stats[grade][subject].items()):
                        stage.info(f"  分类统计: {grade}/{subject}/{qtype}: {count} 个")

            stage.info(f"完成: 有效={valid_count}, 无效={invalid_count}, 超时={timeout_count}, 解析失败={parse_error_count}")

        logger.finish()
    finally:
        os.chdir(old_cwd)


if __name__ == '__main__':
    main()
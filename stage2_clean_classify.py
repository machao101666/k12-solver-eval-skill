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
import config

# 千帆接口地址（保留默认，可用环境变量覆盖）；Token 从环境变量 / .env 读取
QIANFAN_API_URL = config.get('QIANFAN_API_URL', 'https://qianfan.baidubce.com/v2/chat/completions')


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


def check_question(image_path: str, timeout: int = 200, max_retries: int = 1) -> str:
    """
    调用千帆视觉理解API检查图片，首次调用 + 最多 1 次重试，每次限时 200 秒。

    Args:
        image_path: 图片路径
        timeout: 超时时间（秒），默认200秒

    Returns:
        API返回的原始content内容
    """
    url = QIANFAN_API_URL

    base64_image = image_to_base64(image_path)

    payload = {
        "model": "ernie-5.0-thinking-latest",
        "stream": False,
        "temperature": 0.2,
        "enable_thinking": False,
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
                        "text": """# 任务说明
你是专业试题视觉审核与分类专家，依据图片内容执行判定、分类与原因说明，严格遵守所有规则、输出格式及示例范式，禁止自由发挥、自创内容。

# 一、有效性判定总规则
同时满足全部条件，有效性判定输出【1】（有效完整单题）；任意触发以下任一情况，有效性判定输出【2】（无效内容）：
1. 题目数量：图片内以**一道独立大题**为主体，大题附带多个子小题（如数学大题含多小问、阅读理解大题含多道小题），统一视为1道完整试题，判定为有效；
   图片包含**两道及以上独立大题**、多组互不归属的试题、题目片段拼接 → 判定无效。
2. 内容完整性：题目无裁切、无关键文字缺失、内容完整 → 残缺则无效
3. 图像质量：画面无严重模糊、反光、遮挡、扭曲，文字图形可正常识别 → 画质异常则无效
4. 配图校验：题干出现"如图、见下图、附图"等字样，但图片内无对应配图 → 无效
5. 学科限制：试题仅限数学、语文、英语；物理、化学、生物、道法等其他学科 → 无效

补充规则：题目附带手写答案、印刷答案、解题步骤、草稿笔迹，不影响有效性判定。

# 二、输出规则
1. 有效性判定：有效输出数字 1，无效输出数字 2
2. 判定为 1（有效）：依次填写 学段、学科、题型，严格使用指定名称
3. 判定为 2（无效）：学段、学科、题型 统一填写「无」，并在最后用**一句话简短说明无效原因**

# 三、分类可选值（严格照搬，不可修改、简写、新增）
## 学段（三选一）
小学、初中、高中
## 学科（三选一）
数学、语文、英语
## 题型清单
### 数学
选择题、填空题、计算题（口算）、计算题（竖式计算）、计算题（脱式计算）、计算题（解方程）、普通应用大题（无图）、普通应用大题（带图）、证明题、解析几何题
### 语文
选择题、拼音打勾题、连线题、填空题、阅读理解类
### 英语
单项选择、完形填空、阅读理解（选择型）、匹配题/连线题、阅读理解（简答）、书面表达（作文）、填空题

# 四、题型判定口径
1. 数学
纯数字算式 → 计算题（口算）；竖式格式运算 → 计算题（竖式计算）；递等式运算 → 计算题（脱式计算）；含未知数求解 → 计算题（解方程）；文字应用题无图 → 普通应用大题（无图）；文字应用题带图表 → 普通应用大题（带图）；几何求证（必须为解答题题型） → 证明题；坐标/函数综合大题（必须为解答题题型） → 解析几何题；一道大题包含多个小问，仍按该大题整体归类题型。
2. 语文
拼音选读音/字形 → 拼音打勾题；两两内容配对作答 → 连线题；长文加问答（含多道小题） → 阅读理解类。
3. 英语
四选一客观题 → 单项选择；短文挖空选择题 → 完形填空；阅读+选择题作答（含多小题） → 阅读理解（选择型）；两两配对作答 → 匹配题/连线题；阅读+手写作答（含多小题） → 阅读理解（简答）；短文写作类 → 书面表达（作文）。

# 五、Few-shot 示例（参考格式与判断逻辑，严格模仿）
## 示例1（有效：单道小学数学竖式计算题）
有效性判定：1
学段：小学
学科：数学
题型：计算题（竖式计算）
无效原因：无

## 示例2（有效：数学大题包含多个小问，视为一道完整题目）
有效性判定：1
学段：初中
学科：数学
题型：普通应用大题（无图）
无效原因：无

## 示例3（有效：英语阅读理解大题附带多道小题）
有效性判定：1
学段：高中
学科：英语
题型：阅读理解（选择型）
无效原因：无

## 示例4（无效：图片包含两道相互独立的大题）
有效性判定：2
学段：无
学科：无
题型：无
无效原因：图片内存在两道及以上独立大题

## 示例5（无效：题干写有如图，但无对应配图）
有效性判定：2
学段：无
学科：无
题型：无
无效原因：题干提及如图，图片缺少对应配图

## 示例6（无效：属于物理学科，非语数英）
有效性判定：2
学段：无
学科：无
题型：无
无效原因：试题为物理学科，超出指定学科范围

## 示例7（无效：画面严重模糊，无法识别内容）
有效性判定：2
学段：无
学科：无
题型：无
无效原因：图片画质模糊，无法识别题目内容

## 示例8（无效：题目被裁切，内容不完整）
有效性判定：2
学段：无
学科：无
题型：无
无效原因：题目画面裁切，内容残缺不全

## 示例9（有效：语文阅读理解含多道小题）
有效性判定：1
学段：小学
学科：语文
题型：阅读理解类
无效原因：无

# 六、固定输出格式（严格按顺序分行，不增删空行、符号）
有效性判定：
学段：
学科：
题型：
无效原因："""
                    }
                ]
            }
        ]
    }

    headers = {
        'Content-Type': 'application/json',
        'Authorization': f"Bearer {config.require('QIANFAN_API_TOKEN', '千帆有效性判定/分类')}"
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
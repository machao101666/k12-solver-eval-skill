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
import config

# 并发控制
API_SEMAPHORE = threading.Semaphore(20)   # 同时最多 20 个 API 请求（最大并发）
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
# 配置
# ============================================
API_URL = config.get('QIANFAN_API_URL', 'https://qianfan.baidubce.com/v2/chat/completions')
# 千帆 Bearer Token（从环境变量 / .env 读取，源码不含真实密钥）
API_KEY = f"Bearer {config.require('QIANFAN_API_TOKEN', '千帆五模型解题')}"

# Claude API (华研) - 已弃用，改用 OpenAI 兼容格式；如需启用请配置 HUAYAN_API_KEY
HUAYAN_URL = config.get('HUAYAN_API_URL', 'https://www.huayanapi.com/v1/chat/completions')
HUAYAN_KEY = config.get('HUAYAN_API_KEY', '')
HUAYAN_HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {HUAYAN_KEY}',
}

EXCEL_FILE = 'evaluation.xlsx'
BASE_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))

HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': API_KEY,
}

# 提示词模板
PROMPT_TEMPLATE = """你是严格遵循K12课标、输出格式高度标准化的校内作业解题助手，作答严格遵守全部规则，严格模仿示例的写实写法、分步逻辑、句式风格，禁止使用模板化空话、套话、泛泛概括式表述，必须写实、具体、有实际内容。

一、基础作答约束（核心干扰项过滤规则，优先级最高）
1.  执行顺序：先过滤所有干扰信息，再识别纯题干，最后独立解题。
2.  干扰项定义：图片中出现的**任何手写笔迹、学生作答痕迹、学生写的答案（无论对错）、印刷参考答案、已写的解题步骤、草稿、涂改痕迹、标注、打勾/打叉符号**，全部视为无效干扰信息，**必须100%忽略，不得读取、参考、使用**。
3.  解题依据：仅以图片内**题干原题**为唯一依据，完全独立解题，不判断、不评价、不修改图片中已有的任何作答内容。
4.  输出要求：不闲聊、不复述任何题干、不做教学点评、不拓展课外知识点，只输出标准解题过程与最终结论，禁止在输出中提及“图片中的答案”“手写痕迹”等干扰信息。

二、学段防超纲强约束
本题所属学段：{学段}
必须严格遵循{学段}教学大纲，只使用本学段课本常规解法、标准术语；
禁止使用高年级知识、超纲公式、竞赛解法、偏门技巧，只用校内课堂标准基础解法。

三、学科与题型约束
本题所属学科：{学科}
本题所属题型：{题型}
严格按本学科本题型标准答题规范、步骤格式、作答详略程度书写，不凑字数、不写无意义套话，不用“应该”“可能”等话术，要用确定性的语言回答。

四、固定强制输出格式（固定标签，便于代码精准提取）
必须严格使用固定标识开头，标识文字不许改、不许加符号、不许换行乱排：
Reasoning：
书写完整分步推理、演算、证明、分析过程，每一步都要有实际内容，禁止只写"梳理要点、分维度概括"这类空泛话术。

Answer：
1. 计算、选择、填空、简答类有明确结果的题目：给出**具体数值、选项、完整明确结论**，不模糊、不笼统、不敷衍。
2. 数学证明题、几何推导、语文纯分析说理题等无独立单句答案的题型：
Answer 统一固定只写：该命题成立，不空白、不自行编造多余内容。
3.如果一个图片有多个小题，直接输出每个小题的答案，中间用空格间隔，无需写各个小题的题号序号，禁止输出其他任何题干、过程等内容

五、数学格式规范
所有算式、方程、几何表达式统一用标准LaTeX格式包裹。

六、写实标准示例（严格模仿这种写实写法，禁止模板化空话）
【示例1 数学计算题-解方程】
Reasoning：
1. 把常数项移到等式右侧：$3x = 15 - 6$
2. 计算右侧结果：$3x = 9$
3. 两边同除以3，求得未知数：$x = 3$
Answer：$x=3$

【示例2 语文阅读理解简答-人物形象】
Reasoning：
1. 从文中主动帮老人干活、遇事懂得谦让的行为事件切入；
2. 结合人物语言和心理描写，提炼性格品质；
3. 整合具体情节对应的性格特点，逐条凝练表述。
Answer：文中少年孝顺懂事、善良热心、谦逊有礼，做事踏实稳重，懂得体谅长辈。

【示例3 英语阅读简答-原因类】
Reasoning：
1. 定位原文第二段对应语句，抓取人物行为动机；
2. 筛选原文关键信息，用初中标准简单句式整合；
3. 保证语法正确、句意完整，不额外主观臆测。
Answer：He likes summer because he can swim and spend time with his parents on holiday.

【示例4 数学几何证明题】
Reasoning：
1. 由题意可知AB平行于CD，可得内错角相等；
2. 结合已知边相等条件，套用三角形全等判定SAS；
3. 逐层推导对应边、对应角相等，完成完整逻辑证明。
Answer：该命题成立

七、客观题作答规则
选择、填空、计算、竖式、脱式、解方程、解析几何等：
步骤写实分步书写，每一步都有实际运算或推导内容；
Answer 给出确切数值、选项、结果，不用模糊概括；如果有多个题目答案，用空格分割即可，禁止输出题号序号等无关内容。

八、开放题作答强约束（防泛泛而谈）
语文阅读简答、赏析、感悟、英语简答类：
禁止写"从多维度、从主旨角度、概括大意"这类模板空话；
必须结合题目场景写实分析、具体表述，要点明确、语句落地；
同题只允许少量措辞差异，核心要点、结论保持一致，便于后续语义比对。

九、异常兜底规则
若图片模糊、题干残缺、无法识别完整题目，仅固定输出一句话：无法正常解题
"""

# 五个模型配置
MODELS = [
    {
        'label': '模型1(ernie)',
        'model': 'ernie-5.0-thinking-latest',
        'temperature': 0.1,
        'enable_thinking': True,
        'col_model': 7,  # G
        'col_reasoning': 8,  # H
        'col_answer': 9,  # I
        'col_time': 6,  # F
    },
    {
        'label': '模型2(doubao)',
        'model': 'doubao-seed-2-0-pro-260215',
        'provider': 'huayan',
        'temperature': 0.1,
        'enable_thinking': True,
        'col_model': 11,  # K
        'col_reasoning': 12,  # L
        'col_answer': 13,  # M
        'col_time': 10,  # J
    },
    {
        'label': '模型3(qwen)',
        'model': 'qwen3-vl-235b-a22b-thinking',
        'temperature': 0.1,
        'enable_thinking': True,
        'col_model': 15,  # O
        'col_reasoning': 16,  # P
        'col_answer': 17,  # Q
        'col_time': 14,  # N
    },
    {
        'label': '模型4(claude)',
        'model': 'anthropic/claude-opus-4-8',
        'provider': 'claude',
        'temperature': None,
        'enable_thinking': True,
        'col_model': 19,  # S
        'col_reasoning': 20,  # T
        'col_answer': 21,  # U
        'col_time': 18,  # R
    },
    {
        'label': '模型5(gemini)',
        'model': 'gemini-3.1-pro-preview',
        'provider': 'huayan',
        'temperature': 0.1,
        'enable_thinking': True,
        'col_model': 23,  # W
        'col_reasoning': 24,  # X
        'col_answer': 25,  # Y
        'col_time': 22,  # V
    },
]


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


def call_huayan_model(model_config, image_data_url, prompt_text, max_retries=1):
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
                HUAYAN_URL,
                headers=HUAYAN_HEADERS,
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
    调用 Claude API (Anthropic 原生格式 /v1/messages)，单次限时 200s
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
        'x-api-key': HUAYAN_KEY,
        'anthropic-version': '2023-06-01',
    }

    last_error = ''
    for attempt in range(1, max_retries + 2):
        try:
            response = requests.post(
                'https://www.huayanapi.com/v1/messages',
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
    """根据 provider 分发到不同的 API 调用函数"""
    if model_config.get('provider') == 'claude':
        return call_claude_model(model_config, image_full_path, prompt_text)
    elif model_config.get('provider') == 'huayan':
        return call_huayan_model(model_config, image_data_url, prompt_text)
    else:
        return call_model(model_config, image_data_url, prompt_text)


def parse_args():
    parser = argparse.ArgumentParser(description="五模型并行解题")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    return parser.parse_args()


def process_excel(work_dir=None):
    """主流程：读取 evaluation.xlsx，并发处理（5模型各自独立流水线，全局并发受限）"""
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

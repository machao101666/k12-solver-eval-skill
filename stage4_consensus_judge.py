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
import config

# 并发控制
API_SEMAPHORE = threading.Semaphore(8)
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
API_URL = config.get('HUAYAN_API_URL', 'https://www.huayanapi.com/v1/chat/completions')
API_KEY = config.require('HUAYAN_API_KEY', '华研共识裁判')

EXCEL_FILE = 'evaluation.xlsx'
BASE_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))

HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {API_KEY}',
}

# 模型配置
MODEL_CONFIG = {
    'label': '模型4(claude)',
    'model': 'anthropic/claude-opus-4-8',
    'provider': 'claude',
    'temperature': None,
    'enable_thinking': True,
}

REF_ID_TO_MODEL_NAME = {
    'M1': 'ernie-5.0-thinking-latest',
    'M2': 'doubao-seed-2-0-pro-260215',
    'M3': 'qwen3-vl-235b-a22b-thinking',
    'M4': 'claude-opus-4-8',
    'M5': 'gemini-3.1-pro-preview',
}

# 列配置（1-based）
COL_IMAGE_PATH = 2
COL_XUEDUAN = 3
COL_XUEKE = 4
COL_TIXING = 5

# 5个模型输入列
COL_M1_REASONING = 8
COL_M1_ANSWER = 9
COL_M2_REASONING = 12
COL_M2_ANSWER = 13
COL_M3_REASONING = 16
COL_M3_ANSWER = 17
COL_M4_REASONING = 20
COL_M4_ANSWER = 21
COL_M5_REASONING = 24
COL_M5_ANSWER = 25

# 输出列（最后5列）
COL_CONFIDENCE = 26
COL_SELECTED_REF_ID = 27
COL_CONSENSUS_ANSWER = 28
COL_PSEUDO_LABEL = 29
COL_FILTER_REASON = 30

# 提示词模板
PROMPT_TEMPLATE = """你是K12作业解题共识裁判，参照原题图片仅做**纯读题理解**，不读取、不参考图片内任何手写演算、勾选答案、批注草稿、解题痕迹与暗示思路；
图片只用来识别题干、已知条件、设问、小题数量结构、学段学科题型，**严禁以图片内容作为答案、解题过程的评判标准**。
兼容两种出题形式：
① 一张图片仅包含单个独立小题
② 一张图片为一道大题、内部包含多个子小题
输入为5个模型的Reasoning和Answer，若模型无法正常作答，统一内容已标记为「无法作答」。
你的任务：先识别无效作答、再按剩余有效模型做共识投票、置信度分级、择优生成标准伪标签。
严格遵守所有硬性规则，严格模仿下方示例的判定逻辑、格式要求；
严禁任何自行改写、润色、优化、二次创作，只能从5个模型原始有效作答里原样挑选、原样照搬，一字不改动。

【题目基础信息与素材】
原题图片：以图片中题干、设问、小题结构、题型属性作为唯一评判依据
学段：{学段}
学科：{学科}
题型：{题型}

【5个模型原始作答内容】
模型1原始作答：
Reasoning：{M1推理过程}
Answer：{M1最终答案}

模型2原始作答：
Reasoning：{M2推理过程}
Answer：{M2最终答案}

模型3原始作答：
Reasoning：{M3推理过程}
Answer：{M3最终答案}

模型4原始作答：
Reasoning：{M4推理过程}
Answer：{M4最终答案}

模型5原始作答：
Reasoning：{M5推理过程}
Answer：{M5最终答案}

【核心定义：无效作答判定】
凡是 Reasoning 或 Answer 等于「无法作答」，统一判定为无效模型，不参与共识投票、不参与择优；
其余有完整解题过程+有效答案的，判定为有效模型。

【全局统一：答案一致性详细判定规则】
1. 以下差异全部忽略，不作为答案分歧依据：
有无小题题号、序号括号编号、多余文字、空格换行排版、表述语序、句式措辞、答案详略长短、等价写法格式差异。
2. 判定唯一标准：
剥离题号、排版、修饰文字后，以**核心答案内容是否等价一致**为唯一依据；
核心答案相同、只是外在格式表述不同，直接判定答案一致；
核心答案本质不同，才判定答案分歧。

【开放型/半开放型的主观试题 专属规则】
1. 语文简答、阅读理解赏析、主观论述、表达类等开放题无唯一标准答案，允许多维度、多角度、不同行文表述；
2. 开放题不要求文字完全雷同、表述角度完全一致、结论句式一模一样；
3. 开放题判定答案与推理过程一致，仅考核三项：
   ① 核心立意、答题主旨贴合题意一致；
   ② 关键采分点、核心答题要点覆盖完整；
   ③ 逻辑通顺、观点合理、价值观合规、论证自洽；
4. 只要满足以上三项，即便答题话术、行文结构、切入角度、篇幅长短不同，一律视为答案一致、推理过程一致；
5. 禁止用客观题唯一标准答案标准，生硬卡死开放主观题的文字和表述形式。

【全局统一：推理过程Reasoning一致性详细判定规则】
一、总体原则
判断解题推理过程是否一致，不看字数长短、不看表述措辞、不看步骤拆分粗细、不看行文语序、不看过渡话术、不看排版分段、不看题号标注；
只以**核心知识点、解题思路主线、关键推理逻辑、解题方法、核心推导步骤**是否本质同源同向为唯一判定标准。

二、一律忽略、视为过程一致的差异
1. 步骤拆分粒度不同：有的分多步、有的合并简写，不判分歧；
2. 语言表述差异：同义替换、书面/口语化、句式不同，不判分歧；
3. 有无连接过渡词：有无首先/其次/因此/综上等话术，不判分歧；
4. 省略基础平凡步骤：省略简单口算、常识铺垫步骤，只保留关键推导，不判分歧；
5. 排版分段、换行缩进、书写格式不同，不判分歧；
6. 结论顺序不同：先给答案后推导、先推导后给答案，不判分歧；
7. 附带小题题号、编号格式不同，全部忽略，不影响过程一致性。

三、判定为【推理过程一致】硬性标准（需同时满足）
1. 所用核心知识点、公式、定理、语法考点完全相同；
2. 破题切入点、整体解题思考逻辑主线完全一致；
3. 关键核心推导步骤、因果关系、等量关系、排除/分析逻辑一致；
4. 采用的解题方法完全相同（方程法/算术法/代入排除/原文定位等）；
5. 推理逻辑闭环，指向同一核心答案结论。

四、判定为【推理过程不一致】硬性标准（满足任意一条即判不一致）
1. 使用的核心知识点、公式、定理、考点完全不同；
2. 破题切入点、整体思考逻辑主线完全相反、分析角度完全不同；
3. 关键推理步骤逻辑矛盾、因果推导相悖、判定依据对立；
4. 解题方法完全不同，解题路径本质不一样；
5. 推理无实质内容、空话凑字、逻辑混乱、无有效解题推导；
6. 一方有完整严谨推理，另一方只有零散话术、无关键支撑步骤。

五、分题型细化口径
1. 数学：核心公式定理、列式逻辑、推导主线一致，步骤简写/语序变形均算一致；
2. 语文简答/阅读：采分点、原文定位、主旨分析角度一致，表述长短不同算一致；
3. 英语：语法考点、时态从句、固定搭配判断逻辑、原文定位思路一致即算一致；
4. 证明/理科大题：论证主线、引用定理、推导链条一致，省略平凡步骤一律算一致。

【第一步：统计有效模型数量】
先逐个筛查5个模型，剔除所有标记「无法作答」的无效模型，统计当前有效模型总数N。

【第二步：单小题 动态自适应共识投票规则（所有模型等价权重）】
按实际有效数量N，严格执行以下分级阈值：
1. N=5（5个全有效）：≥4个模型核心答案等价 + 推理过程一致 → confidence=high
2. N=4（4个有效）：≥3个模型核心答案等价 + 推理过程一致 → confidence=high
3. N=3（3个有效）：必须3个模型核心答案等价 + 推理过程全部一致 → confidence=high
4. N=2（2个有效）：即便核心答案与推理过程一致，也只定为 medium，不允许判high
5. N≤1（仅1个或0个有效）：直接定为 low，无共识参考价值

未达到high标准的，按实际核心答案与推理分歧程度判为 medium 或 low。

【第三步：大题含多个子小题 判定规则】
1. 先识别图片内大题包含的所有独立小题；
2. 逐小题单独执行上面「单小题动态投票规则」，该小题达到high标准记为合格小题；
3. 计算本大题合格小题占比 = 合格小题数 ÷ 本大题总小题数；
4. 合格小题占比 ≥80% → 整题 confidence=high；
5. 占比不足80%或多数小题共识弱 → 按实际分歧判 medium / low。

【特殊题型兜底规则】
1. 英语听力类题型：本身无文字作答条件，若多数模型为「无法作答」，直接判low；
2. 初中/高中超高难度数学题，有效模型少、难以达成高共识，不强行拉高置信度，按实际投票结果正常分级；
3. 5个模型全部为「无法作答」，直接判low，标注原因即可。

【最优模型选取规则】
1. 所有模型完全等价权重，不优先任何序号模型；
2. 仅在达成high共识的有效模型群体中择优；
3. 择优只看内容：解题推理步骤最完整、逻辑最严谨、贴合课内答题规范；
4. 必须原样照搬选中模型原始Reasoning、Answer，禁止改写、润色、调整语句、删减原有题号与格式。

【输出字段规范 & 显示规则】
1. pseudo_label 固定格式：
Reasoning：选中模型完整原始解题推理过程
Answer：选中模型完整原始最终答案
原样复刻原文，不删减、不修改、不额外加内容，保留原有题号、排版与输出格式。

2. JSON固定字段顺序：
confidence → selected_ref_id → consensus_answer → pseudo_label → filter_reason

3. 分支输出强制规则：
① 当 confidence=high：完整输出全部5个字段；
② 当 confidence=medium / low：不输出 selected_ref_id、consensus_answer、pseudo_label，仅保留 confidence、filter_reason 两个字段。

4. 字段说明：
consensus_answer：复用选中模型原始Answer全文，保留原有格式题号；
filter_reason：high填空；medium/low必须写明具体原因：包含但不限于
有效模型数量少、未达高共识投票阈值、多小题合格占比不足80%、题型为英语听力无作答条件、5个模型均无法作答、初高中难题模型普遍输出失效、模型仅题号/排版/措辞有差异但核心答案或推理逻辑存在分歧等。

【Few-Shot 示例】
示例1 5个模型全有效，4个核心答案一致、推理逻辑一致，仅题号表述格式不同 → high
{{
"confidence": "high",
"selected_ref_id": "M2",
"consensus_answer": "C",
"pseudo_label": "Reasoning：梳理题目条件，分步推导知识点，逐项辨析选项对错。\\nAnswer：C",
"filter_reason": ""
}}

示例2 仅2个有效模型，虽核心答案与推理一致但只判medium
{{
"confidence": "medium",
"filter_reason": "仅2个模型有效作答，虽核心答案与推理思路一致，未达到高置信投票阈值"
}}

示例3 5个模型全部无法作答 → low
{{
"confidence": "low",
"filter_reason": "本题难度过高，5个模型均输出无法作答，无有效共识模型"
}}

示例4 英语听力题型多数无法作答 → low
{{
"confidence": "low",
"filter_reason": "本题为英语听力题型，无文字作答条件，多数模型输出无法作答"
}}

【输出强制要求】
1. 必须先自动过滤「无法作答」的无效模型，再统计有效数量N进行动态投票；
2. 严格遵循答案一致性、推理过程一致性全套详细规则，忽略题号、编号、排版、句式、步骤详略等非本质差异；
3. 仅以核心答案等价度、知识点、解题主线、关键推理逻辑作为共识判定唯一依据；
4. 严格按N对应的阈值判定high/medium/low，不放宽、不自定义标准；
5. 严格遵守字段顺序与中低置信字段隐藏规则；
6. pseudo_label 保留固定Reasoning、Answer双标签，原样复用模型原文，不修改原有题号格式；
7. 只返回纯净JSON，无多余解释、无多余文字、无Markdown。
"""


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
    max_bytes = 5 * 1024 * 1024

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
                timeout=200,
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

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
import config


# 并发与限流控制
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
HUAYAN_API_URL = config.get('HUAYAN_API_URL', 'https://www.huayanapi.com/v1/chat/completions')
# 华研与内部 One API 网关的 Token 均从环境变量 / .env 读取，源码不含真实密钥
ONEAPI_API_URL = config.get('ONEAPI_API_URL', 'https://oneapi-comate.baidu-int.com/v1/chat/completions')

EXCEL_FILE = 'evaluation.xlsx'
BASE_DIR = os.environ.get("AUTO_EVAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))

# 评分裁判配置：多套并存，默认走内部 One API 网关的 Opus 5；
# 切换方式：AUTO_EVAL_JUDGE_MODEL=oneapi-opus-5 | huayan-opus-5 | huayan-opus-4-8
JUDGE_MODEL_PRESETS = {
    'oneapi-opus-5': {
        'label': '模型4(claude)',
        'model': 'Opus 5',
        'provider': 'claude',
        'temperature': None,
        'enable_thinking': True,
        'api_url': ONEAPI_API_URL,
        'api_key_env': 'ONE_API_TOKEN',
    },
    'huayan-opus-5': {
        'label': '模型4(claude)',
        'model': 'anthropic/claude-opus-5',
        'provider': 'claude',
        'temperature': None,
        'enable_thinking': True,
        'api_url': HUAYAN_API_URL,
        'api_key_env': 'HUAYAN_API_KEY',
    },
    # 历史基线：0730vis 等既往报告均由该模型判分，保留以便复现对比
    'huayan-opus-4-8': {
        'label': '模型4(claude)',
        'model': 'anthropic/claude-opus-4-8',
        'provider': 'claude',
        'temperature': None,
        'enable_thinking': True,
        'api_url': HUAYAN_API_URL,
        'api_key_env': 'HUAYAN_API_KEY',
    },
}

JUDGE_MODEL_KEY = os.environ.get('AUTO_EVAL_JUDGE_MODEL', 'oneapi-opus-5')
if JUDGE_MODEL_KEY not in JUDGE_MODEL_PRESETS:
    raise SystemExit(
        f"未知的裁判模型 AUTO_EVAL_JUDGE_MODEL={JUDGE_MODEL_KEY}，可选: {sorted(JUDGE_MODEL_PRESETS)}"
    )
MODEL_CONFIG = JUDGE_MODEL_PRESETS[JUDGE_MODEL_KEY]

API_URL = MODEL_CONFIG['api_url']
if MODEL_CONFIG.get('api_key_env'):
    API_KEY = os.environ.get(MODEL_CONFIG['api_key_env'], '')
    if not API_KEY:
        raise SystemExit(
            f"裁判 {JUDGE_MODEL_KEY} 需要环境变量 {MODEL_CONFIG['api_key_env']}，当前为空"
        )
else:
    API_KEY = MODEL_CONFIG['api_key']

HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {API_KEY}',
}

# 列配置（1-based）
COL_IMAGE_PATH = 2          # B
COL_GRADE_BAND = 3          # C
COL_SUBJECT = 4             # D
COL_QUESTION_TYPE = 5       # E
COL_CONFIDENCE = 26         # Z
COL_PSEUDO_LABEL = 29       # AC
COL_MODEL_FULL_OUTPUT = 32  # AF

# 输出列
COL_SCORE_0_3 = 35          # AI - 小度 score_0_3
COL_JUDGE_REASON = 36       # AJ - 小度 judge_reason

# 五个参考模型评分输出列：AK/AL、AM/AN、AO/AP、AQ/AR、AS/AT
REFERENCE_MODEL_SCORE_TARGETS = [
    {
        'label': 'ernie',
        'reasoning_col': 8,    # H
        'answer_col': 9,       # I
        'score_col': 37,       # AK
        'reason_col': 38,      # AL
        'score_header': 'ernie_score_0_3',
        'reason_header': 'judge_reason',
    },
    {
        'label': 'doubao',
        'reasoning_col': 12,   # L
        'answer_col': 13,      # M
        'score_col': 39,       # AM
        'reason_col': 40,      # AN
        'score_header': 'doubao_score_0_3',
        'reason_header': 'judge_reason',
    },
    {
        'label': 'qwen',
        'reasoning_col': 16,   # P
        'answer_col': 17,      # Q
        'score_col': 41,       # AO
        'reason_col': 42,      # AP
        'score_header': 'qwen_score_0_3',
        'reason_header': 'judge_reason',
    },
    {
        'label': 'claude',
        'reasoning_col': 20,   # T
        'answer_col': 21,      # U
        'score_col': 43,       # AQ
        'reason_col': 44,      # AR
        'score_header': 'claude_score_0_3',
        'reason_header': 'judge_reason',
    },
    {
        'label': 'gemini',
        'reasoning_col': 24,   # X
        'answer_col': 25,      # Y
        'score_col': 45,       # AS
        'reason_col': 46,      # AT
        'score_header': 'gemini_score_0_3',
        'reason_header': 'judge_reason',
    },
]

PROMPT_TEMPLATE = """你是K12作业解题模型专业评分裁判，必须同时严格遵守两类规则：
第一类：历史全局约束规则（必须逐条严格执行，不得简化、不得省略）
第二类：严格逐字遵循《作业解题模型评测方案》3.1、3.2 完整0～3分评分量表所有条款，逐条匹配判分，不自行增减、不放宽、不主观变通。

==================== 第一部分 全局固定约束规则（逐条强制执行） ====================
0. 最高优先级字段非空强制拦截规则
所有输入字段：原题图片、题目学段、题目学科、题目题型、伪标签原文、待测模型完整输出原文，
任意一个字段为空、无有效内容、缺失，**无需后续任何判分、不自行从图片读取答案补标答、不做内容比对**，直接强制判0分，不再执行后续所有评分逻辑。
本规则优先级高于所有读图、内容比对、量表打分规则，无条件优先执行。

1. 原图使用约束
仅把原题图片用来读取题干、已知条件、设问、自动识别大题包含的小题数量、学段、学科、题型；
严格忽略图片内所有手写笔迹、演算过程、草稿批注、勾选答案、人为标注、解题痕迹；
禁止参考图片里任何手写答案、解题步骤作为评分依据；
仅以模块B产出的伪标签作为唯一标准答案依据。

2. 答案与推理一致性比对约束
评判模型答案、推理是否一致时，以下所有差异全部忽略，不作为分歧、扣分依据：
有无小题题号、序号括号编号、多余文字、空格换行排版、表述语序、句式措辞、答案详略长短、等价写法格式差异；
步骤拆分粗细不同、有的分多步有的合并简写、有无过渡连接词、先给答案后推导/先推导后答案、排版分段缩进差异，全部忽略；
只以核心答案内容等价性、核心知识点、解题思路主线、关键推理逻辑、核心推导步骤作为唯一评判标准。

3. 开放型主观试题专属完整约束
适用：语文阅读理解、赏析简答、主观论述、表达类、作文类无唯一标准答案题目；
1）开放题无唯一标准答案，允许多维度、多角度、不同行文结构、不同话术表述、不同篇幅长短；
2）不要求文字完全雷同、表述角度完全一致、句式结论一模一样；
3）评分仅考核三项：核心立意答题主旨贴合题意、关键采分点核心答题要点覆盖完整、逻辑通顺观点合理价值观合规论证自洽；
4）只要满足以上三项，仅话术、行文结构、切入角度、篇幅不同，一律不扣分、不压低分数；
5）禁止用客观题唯一标准答案标准生硬卡死开放题文字表述；
6）开放题同样必须有明确收尾作答观点与最终结论，仅有分析推理无最终明确观点结论，同样按规则判0分。

4. 待测模型结构化输出 官方固定格式定义（严格执行）
必选标题（两个标题缺一不可，必须按以下原样、顺序出现）：
==**解题过程**==
==**最终答案**==
格式合规判定规则：
1. 只有完整包含上述两个精确标题，且“==**解题过程**==”在前、“==**最终答案**==”在后，才满足3分的基础格式要求；
2. 标题文字、星号数量、等号数量或标题顺序任一不一致，均视为未按约定结构化格式输出；
3. 使用旧格式【详细解析】【最终答案】或其他近似、自创标题，严格不能评3分，即使内容完全正确也最高只能评2分；
4. 除上述两个必选标题外，不设置可选结构化模块；出现其他额外结构化大标题，视为格式不完全合规，严格不能评3分；
5. 无结构化纯自由行文不满足3分格式标准，仅按内容打0-2分。

5. 待测模型内容解析完整规则
1）严格符合官方结构化标题时：
按固定标题精准拆分：
“==**解题过程**==”之后、“==**最终答案**==”之前的全部内容，作为正式解题推理过程；
“==**最终答案**==”之后的全部内容，作为正式最终作答结论。
2）未严格符合结构化标题时的兼容规则：
若待测输出缺少任一精确标题、使用旧标题或近似标题、标题顺序错误，禁止仅因格式问题直接判0分；
必须通读全文，自动语义区分解题推理过程与最终明确答案，正常参与内容对标打分，但最高只能给2分，严格不参与3分评定。
3）硬性零分强约束（无条件执行）：
无论是否有结构化标题，只要仅有深度思考、推理、解析过程，没有给出任何明确最终答案、没有收尾作答结论，直接强制判0分，不按过程酌情给分。

6. 通用输出约束
1）仅模块C对入集高置信题目待测输出打分，模块A、模块B不参与打分；
2）严禁改写、润色、优化、修正、二次创作伪标签内容与待测模型输出内容，只做解析、识别、比对、按规则客观打分；
3）模型标注「无法作答」、明确说明题型限制无法作答、无任何有效解题内容，统一纳入0分判定；
4）打分只看内容合规性、逻辑正确性、完整度、量表条款，不因输出格式不标准无故压分；
5）一图多题场景：由裁判自行阅读原题图片，自动识别本大题包含的小题总数量，自行判断漏答占比，不依赖外部传入小题总数字段。


==================== 第二部分 官方评测方案 3.1 统一评分原则 ====================
1. 采用同一套 0～3 分可用性量表，0～3 分仅用于待测模型输出；
2. 模块A、模块B不打分，仅模块C评分裁判对入集题的待测返回进行打分；
3. 主指标：2分率 = 待测输出得分≥2的题数 / 有效评测题数；
4. 风险指标：0分率 = 待测输出得分=0的题数 / 有效评测题数；
5. 3分率以本次约定固定输出格式为唯一判定依据，符合格式+内容满分即可评3分；
6. 总分官方口径：
0分：数理/逻辑硬伤，完全不可用
1分：核心任务有瑕疵，基本不可用或勉强可参考
2分：核心任务无问题（数理/逻辑正确），用户可看懂、可使用
3分：在2分基础上，输出完全符合本次约定的固定格式与字段要求

==================== 第三部分 官方评测方案 3.2 完整0～3分评分量表 ====================
一、0分 完全不可用（满足任意一条即判0分）
1. 推理不可用：解题/分析类推理步骤可用占比不足50%，不论最终答案是否正确；
2. 答非所问：内容与题目无关，或严重理解错题意；
3. 胡言乱语：输出无意义、乱码、重复堆砌、无实质解题内容；
4. 漏答严重：一图多题时，裁判自行阅读原题图片识别小题总数，漏答小题数量占比超过50%；
5. 输入异常仍硬答：题图缺失/不可读仍强行输出作答；
6. 仅有思考推理解析过程，无任何明确最终答案、无收尾作答结论；
7. 模型输出无法作答、无有效解题推理与答案内容。

二、1分 基本不可用（不满足0分条件，满足任意一条即判1分）
1. 过程尚可但答案错：推理步骤可用占比超过50%，但最终答案错误；
2. 答案对但缺关键过程：最终答案正确，但缺少必要推理关键步骤，无法支撑学习使用；
3. 选对但过程错：选择/填空类题目，解题过程明显错误，答案碰巧正确；
4. 严重啰嗦：超过30%内容为废话、车轱辘话、无意义重复表述；
5. 漏答较多：一图多题时，裁判自行阅读原题图片识别小题总数，仅少数小题答对，漏答占比未超50%但缺陷明显；
6. 幻觉多题：无中生有编造额外小题，原题部分小题答对；
7. 逻辑对结果错：推理逻辑链条完整严谨，但最终结果错误或幻觉错误；
8. 公式硬伤：数学公式存在明显原理性错误，单纯文字笔误不判1分。

三、2分 核心可用（本期主达标线，不满足0分、1分所有条件）
1. 核心数理、逻辑、主旨立意无问题，与伪标答/consensus无实质性冲突；
2. 关键推理步骤、最终答案正确完整，用户可看懂、可正常使用；
3. 格式宽松兼容：未严格按本次固定结构化标题输出，但解题过程与答案齐全可辨识；
4. 允许非致命轻微瑕疵：LaTeX渲染小问题、缺少$符号、少量错别字、表述简略；
5. 一图多题场景，裁判自行阅读原题图片识别小题总数，漏答小题占比不足50%，无严重漏答；
6. 未同时、按顺序、原样包含“==**解题过程**==”和“==**最终答案**==”时，最高顶格给到2分，严格不评定3分；
7. 凡解题过程略有简略、少量步骤省略、表述口语化、存在轻微可感知不严谨瑕疵，均属于正常可用范围，统一判定为2分，不晋级3分。

四、3分 完全可用（严格高标准，缺一不可，门槛抬高）
必须同时满足以下所有严苛条件，才可评定3分，任意一条不满足一律最高只给2分：
1. 先完全达到2分所有标准，无0分、1分任何硬伤；
2. 结构化输出严格符合唯一约定格式：必须按顺序、原样完整包含“==**解题过程**==”和“==**最终答案**==”两个标题；
3. 两个标题的文字、星号数量、等号数量和先后顺序必须完全一致；使用【详细解析】【最终答案】等旧格式、近似格式或其他自创标题，严格不能评3分；
4. “==**解题过程**==”模块中的推理步骤完整饱满，无跳步、无关键步骤省略、无过度简略，逻辑推导严谨闭环；
5. “==**最终答案**==”模块中必须包含明确、正确、完整的最终作答结论；
6. 表述专业规范、逻辑严谨，无口语化随意表述、无逻辑断层、无解释敷衍；
7. 无任何可感知瑕疵：无轻微逻辑疏漏、无简略省略、无表述不严谨、无冗余废话、无公式小瑕疵、无错别字；
8. 无严重啰嗦、无幻觉多题、无较多漏答、无公式原理硬伤；
9. 只要格式不完全匹配、解析简略、跳步骤、表述随意或存在轻微不严谨任一情况，直接取消3分资格，顶格判定2分。

==================== 第四部分 伪标签固定解析规则 ====================
伪标签固定标准格式：
Reasoning：标准解题推理全过程
Answer：标准最终答案
严格按前缀拆分解析：
Reasoning：后面全部内容 ＝ 标准推理过程
Answer：后面全部内容 ＝ 标准答案

==================== 第五部分 输入输出规范 ====================
输入信息
原题图片：{image}
题目学段：{grade_band}
题目学科：{subject}
题目题型：{question_type}
伪标签原文：{pseudo_label_text}
待测模型完整输出原文：{model_full_output}

输出强制要求
1. 仅返回纯净JSON，无多余解释、无多余文字、无Markdown；
2. 固定两个字段：
score_0_3：只能为 0、1、2、3 整数
judge_reason：必须完整写明：
① 伪标签是否正常解析出Reasoning与标准答案；
② 待测是否按顺序、原样包含“==**解题过程**==”和“==**最终答案**==”，并说明两个模块的拆分情况；
③ 明确说明是否识别到有效最终作答结论，无结论需明确说明并引用0分规则；
④ 严格对标官方量表，写明触发0分/1分/2分/3分的具体条款依据；
⑤ 说明通过原题图片识别到的小题数量、推理可用占比、漏答比例、啰嗦程度、公式是否有硬伤；
⑥ 明确判定是否完全符合本次唯一约定结构化格式；任一标题缺失、改写、顺序错误或使用旧格式时，必须明确说明不具备3分格式条件。

输出示例
{{"score_0_3":0,"judge_reason":"已正常解析伪标签Reasoning与Answer，待测缺少==**最终答案**==标题且仅输出思考推理过程，无任何明确最终作答结论，按全局约束及官方0分规则判定为0分"}}
{{"score_0_3":1,"judge_reason":"已解析标准伪标签内容，成功识别待测推理与答案；推理可用占比超50%但最终答案错误，符合官方1分「过程尚可但答案错」条款，判定为1分"}}
{{"score_0_3":2,"judge_reason":"已解析伪标签并识别待测作答内容，无0分、1分任一违规情形，核心逻辑与答案正确；待测未按顺序、原样包含==**解题过程**==和==**最终答案**==，不具备3分格式条件，最高顶格判定为2分"}}
{{"score_0_3":3,"judge_reason":"已解析伪标签并拆分待测结构化模块，按顺序、原样完整包含==**解题过程**==和==**最终答案**==，格式完全符合唯一约定；核心答案与推理逻辑正确完整，无0分、1分硬伤且无任何可感知瑕疵，满足3分全部条件，判定为3分"}}"""


REFERENCE_PROMPT_TEMPLATE = """你是 K12 作业解题模型专业评分裁判，必须同时严格遵守两类规则：
第一类：历史全局约束规则（必须逐条严格执行，不得简化、不得省略）
第二类：严格逐字遵循《作业解题模型评测方案》3.1、3.2 完整 0～3 分评分量表所有条款，逐条匹配判分，不自行增减、不放宽、不主观变通。
==================== 第一部分 全局固定约束规则（逐条强制执行） ====================
0. 最高优先级字段非空强制拦截规则
所有输入字段：原题图片、题目学段、题目学科、题目题型、伪标签原文、待测模型完整输出原文，
任意一个字段为空、无有效内容、缺失，无需后续任何判分、不自行从图片读取答案补标答、不做内容比对，直接强制判 0 分，不再执行后续所有评分逻辑。
本规则优先级高于所有读图、内容比对、量表打分规则，无条件优先执行。
原图使用约束
仅把原题图片用来读取题干、已知条件、设问、自动识别大题包含的小题数量、学段、学科、题型；
严格忽略图片内所有手写笔迹、演算过程、草稿批注、勾选答案、人为标注、解题痕迹；
禁止参考图片里任何手写答案、解题步骤作为评分依据；
仅以模块 B 产出的伪标签作为唯一标准答案依据。
答案与推理一致性比对约束
评判模型答案、推理是否一致时，以下所有差异全部忽略，不作为分歧、扣分依据：
有无小题题号、序号括号编号、多余文字、空格换行排版、表述语序、句式措辞、答案详略长短、等价写法格式差异；
步骤拆分粗细不同、有的分多步有的合并简写、有无过渡连接词、先给答案后推导 / 先推导后答案、排版分段缩进差异，全部忽略；
只以核心答案内容等价性、核心知识点、解题思路主线、关键推理逻辑、核心推导步骤作为唯一评判标准。
开放型主观试题专属完整约束
适用：语文阅读理解、赏析简答、主观论述、表达类、作文类无唯一标准答案题目；
1）开放题无唯一标准答案，允许多维度、多角度、不同行文结构、不同话术表述、不同篇幅长短；
2）不要求文字完全雷同、表述角度完全一致、句式结论一模一样；
3）评分仅考核三项：核心立意答题主旨贴合题意、关键采分点核心答题要点覆盖完整、逻辑通顺观点合理价值观合规论证自洽；
4）只要满足以上三项，仅话术、行文结构、切入角度、篇幅不同，一律不扣分、不压低分数；
5）禁止用客观题唯一标准答案标准生硬卡死开放题文字表述；
6）开放题同样必须有明确收尾作答观点与最终结论，仅有分析推理无最终明确观点结论，同样按规则判 0 分。
待测模型内容解析完整规则
1）无需依靠固定标题拆分内容，通读全文自主甄别解题推理内容与最终答案；无论排版、行文结构自由与否，均可正常参与 0-3 全档位打分，不因无结构化降档。
2）硬性零分强约束（无条件执行）：
无论是否有结构化标题，只要仅有深度思考、推理、解析过程，没有给出任何明确最终答案、没有收尾作答结论，直接强制判 0 分，不按过程酌情给分。
通用输出约束
1）严禁改写、润色、优化、修正、二次创作伪标签内容与待测模型输出内容，只做解析、识别、比对、按规则客观打分；
2）模型标注「无法作答」、明确说明题型限制无法作答、无任何有效解题内容，统一纳入 0 分判定；
3）打分只看内容合规性、逻辑正确性、完整度、量表条款，不因输出格式不标准无故压分；
4）一图多题场景：由裁判自行阅读原题图片，自动识别本大题包含的小题总数量，自行判断漏答占比，不依赖外部传入小题总数字段。
==================== 第二部分 官方评测方案 3.1 统一评分原则 ====================
采用同一套 0～3 分可用性量表，0～3 分仅用于待测模型输出；
模块 A、模块 B 不打分，仅模块 C 评分裁判对入集题的待测返回进行打分；
主指标：2 分率 = 待测输出得分≥2 的题数 / 有效评测题数；
风险指标：0 分率 = 待测输出得分 = 0 的题数 / 有效评测题数；
总分官方口径：
0 分：数理 / 逻辑硬伤，完全不可用
1 分：核心任务有瑕疵，基本不可用或勉强可参考
2 分：核心任务无问题（数理 / 逻辑正确），用户可看懂、可使用
3 分：在 2 分基础上，解题内容完整严谨、全文无各类细微瑕疵
==================== 第三部分 官方评测方案 3.2 完整 0～3 分评分量表 ====================
一、0 分 完全不可用（满足任意一条即判 0 分）
推理不可用：解题 / 分析类推理步骤可用占比不足 50%，不论最终答案是否正确；
答非所问：内容与题目无关，或严重理解错题意；
胡言乱语：输出无意义、乱码、重复堆砌、无实质解题内容；
漏答严重：一图多题时，裁判自行阅读原题图片识别小题总数，漏答小题数量占比超过 50%；
输入异常仍硬答：题图缺失 / 不可读仍强行输出作答；
仅有思考推理解析过程，无任何明确最终答案、无收尾作答结论；
模型输出无法作答、无有效解题推理与答案内容。
二、1 分 基本不可用（不满足 0 分条件，满足任意一条即判 1 分）
过程尚可但答案错：推理步骤可用占比超过 50%，但最终答案错误；
答案对但缺关键过程：最终答案正确，但缺少必要推理关键步骤，无法支撑学习使用；
选对但过程错：选择 / 填空类题目，解题过程明显错误，答案碰巧正确；
严重啰嗦：超过 30% 内容为废话、车轱辘话、无意义重复表述；
漏答较多：一图多题时，裁判自行阅读原题图片识别小题总数，仅少数小题答对，漏答占比未超 50% 但缺陷明显；
幻觉多题：无中生有编造额外小题，原题部分小题答对；
逻辑对结果错：推理逻辑链条完整严谨，但最终结果错误或幻觉错误；
公式硬伤：数学公式存在明显原理性错误，单纯文字笔误不判 1 分。
三、2 分 核心可用（本期主达标线，不满足 0 分、1 分所有条件）
核心数理、逻辑、主旨立意无问题，与伪标答 /consensus 无实质性冲突；
关键推理步骤、最终答案正确完整，用户可看懂、可正常使用；
允许非致命轻微瑕疵：LaTeX 渲染小问题、缺少 $ 符号、少量错别字、表述简略；
一图多题场景，裁判自行阅读原题图片识别小题总数，漏答小题占比不足 50%，无严重漏答；
凡解题解析略有简略、少量步骤省略、表述口语化、可选模块内容偏敷衍、存在轻微可感知不严谨瑕疵，均属于正常可用范围，统一判定为 2 分，不晋级 3 分。
四、3 分 完全可用（严格高标准，缺一不可，门槛抬高）
必须同时满足以下所有严苛条件，才可评定 3 分，任意一条不满足一律最高只给 2 分：
先完全达到 2 分所有标准，无 0 分、1 分任何硬伤；
解析推理步骤完整饱满，无跳步、无关键步骤省略、无过度简略，逻辑推导严谨闭环；
表述专业规范、逻辑严谨，无口语化随意表述、无逻辑断层、无解释敷衍；
无任何可感知瑕疵：无轻微逻辑疏漏、无简略省略、无表述不严谨、无冗余废话、无公式小瑕疵、无错别字；
无严重啰嗦、无幻觉多题、无较多漏答、无公式原理硬伤；
只要存在解析简略、跳步骤、表述随意、考点分析敷衍、轻微不严谨任一情况，直接取消 3 分资格，顶格判定 2 分。
==================== 第四部分 伪标签固定解析规则 ====================
伪标签固定标准格式：
Reasoning：标准解题推理全过程
Answer：标准最终答案
严格按前缀拆分解析：
Reasoning：后面全部内容 ＝ 标准推理过程
Answer：后面全部内容 ＝ 标准答案
==================== 第五部分 输入输出规范 ====================
输入信息
原题图片：{image}
题目学段：{grade_band}
题目学科：{subject}
题目题型：{question_type}
伪标签原文：{pseudo_label_text}
待测模型完整输出原文：{model_full_text}
输出强制要求
仅返回纯净 JSON，无多余解释、无多余文字、无 Markdown；
固定两个字段：
score_0_3：只能为 0、1、2、3 整数
judge_reason：必须完整写明：
① 伪标签是否正常解析出 Reasoning 与标准答案；
② 通读待测全文自主拆分识别推理、答案内容；
③ 明确说明是否识别到有效最终作答结论，无结论需明确说明并引用 0 分规则；
④ 严格对标官方量表，写明触发 0 分 / 1 分 / 2 分 / 3 分的具体条款依据；
⑤ 说明通过原题图片识别到的小题数量、推理可用占比、漏答比例、啰嗦程度、公式是否有硬伤；
输出示例
{{"score_0_3":0,"judge_reason":"已正常解析伪标签 Reasoning 与 Answer，通读待测全文仅输出思考推理过程，无任何明确最终作答结论，按全局约束及官方 0 分规则判定为 0 分"}}
{{"score_0_3":1,"judge_reason":"已解析标准伪标签内容，通读待测全文识别推理与答案；推理可用占比超 50% 但最终答案错误，符合官方 1 分「过程尚可但答案错」条款，判定为 1 分"}}
{{"score_0_3":2,"judge_reason":"已解析伪标签并通读识别待测作答内容，无 0 分、1 分任一违规情形，核心逻辑与答案正确，存在轻微表述瑕疵，符合 2 分判定标准，判定为 2 分"}}
{{"score_0_3":3,"judge_reason":"已解析伪标签并通读识别待测推理、答案内容，核心答案与推理逻辑正确，无 0 分 1 分硬伤、全文无各类细微瑕疵，满足 3 分全部条件，判定为 3 分"}}
"""


def build_prompt(grade_band, subject, question_type, pseudo_label_text, model_full_output):
    """构建小度评分裁判提示词。小度评分逻辑保持原样。"""
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
    """将本地图片文件转为 base64 data URL，超限时自动压缩到限制内。

    上游按 base64 编码后的字节数限制 5MB，而 base64 会把体积放大约 4/3，
    因此原图阈值取 5MB*3/4，避免出现 IMAGE_SIZE_EXCEEDED。
    """
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
    max_base64_bytes = 5 * 1024 * 1024
    max_bytes = max_base64_bytes * 3 // 4

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
                timeout=200,
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
    score_header = 'score_0_3' if targets == 'target' else '小度_score_0_3'
    ws.cell(row=1, column=COL_SCORE_0_3, value=score_header)
    ws.cell(row=1, column=COL_JUDGE_REASON, value='judge_reason')
    for target in REFERENCE_MODEL_SCORE_TARGETS:
        ws.cell(row=1, column=target['score_col'], value=target['score_header'])
        ws.cell(row=1, column=target['reason_col'], value=target['reason_header'])


def parse_args():
    parser = argparse.ArgumentParser(description="待测模型与五参考模型评分裁判")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录，默认使用 AUTO_EVAL_WORK_DIR 或脚本目录")
    parser.add_argument("--concurrency", type=int, default=8, help="评分请求级并发数，默认 8")
    parser.add_argument("--targets", choices=("all", "target", "references"), default="all", help="评分目标：全部、仅待测模型或仅五参考模型")
    parser.add_argument("--invalid-retries", type=int, default=2, help="裁判响应无法解析为合法评分时的额外重试次数，默认 2")
    parser.add_argument("--only-invalid", action="store_true", help="只补跑分数非法或理由为空的评分目标")
    return parser.parse_args()


def process_excel(work_dir=None, concurrency=8, targets="all", invalid_retries=2, only_invalid=False):
    """主流程：读取 evaluation.xlsx，以单个评分请求为粒度并发评分。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    global API_SEMAPHORE
    concurrency = max(1, int(concurrency or 8))
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
    stage.info(f"共 {total_rows - 1} 行数据，开始处理（仅confidence=high，评分请求级并发={concurrency}）...")

    row_indices = list(range(2, total_rows + 1))
    stats = {"processed": 0, "skipped": 0, "failed": 0, "xiaodu_scored": 0, "score_tasks": 0, "score_task_failed": 0}
    row_states = {}
    score_tasks = []

    def mark_row_failed(row_idx, score_col, reason_col, reason):
        with EXCEL_LOCK:
            ws.cell(row=row_idx, column=score_col, value='')
            ws.cell(row=row_idx, column=reason_col, value=reason)
        row_states[row_idx] = {'failed': True, 'xiaodu_scored': False, 'summaries': [reason]}

    for row_idx in row_indices:
        confidence = (ws.cell(row=row_idx, column=COL_CONFIDENCE).value or '').strip().lower()
        if confidence != 'high':
            stats['skipped'] += 1
            message = f"[行{row_idx}] 跳过: confidence={confidence or '空'}"
            print(message)
            stage.info(message)
            continue

        row_states[row_idx] = {'failed': False, 'xiaodu_scored': False, 'summaries': []}
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
                row_states[row_idx]['xiaodu_scored'] = True
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
                row_states[row_idx]['xiaodu_scored'] = True
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
                    'is_xiaodu': True,
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
                    'is_xiaodu': False,
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
            state = row_states.setdefault(row_idx, {'failed': False, 'xiaodu_scored': False, 'summaries': []})

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
                    if item['is_xiaodu'] and score != '':
                        state['xiaodu_scored'] = True
                    message = f"[行{row_idx}] {label}评分完成: {score}，耗时 {elapsed}s"

            print(message)
            stage.info(message)

    for row_idx, state in sorted(row_states.items()):
        summary = ', '.join(state['summaries'])
        if state['xiaodu_scored']:
            stats['processed'] += 1
            stats['xiaodu_scored'] += 1
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
        f"failed={stats['failed']}, xiaodu_scored={stats['xiaodu_scored']}, "
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

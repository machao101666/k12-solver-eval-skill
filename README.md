# K12 解题模型自动化评测流水线

给「拍照解题」类多模态模型做端到端质量评测的一套脚本。

输入一批试卷图片，自动完成切题、清洗分类、多模型交叉解题、共识伪标签、待测模型解题、LLM 裁判评分，最后产出一份可直接给人看的 Markdown 评测报告。

解决的核心问题是**没有标准答案**：K12 作业题图没有现成 ground truth，人工标注成本高。这里的做法是让 5 个不同厂商的强模型独立解题，只有它们互相印证、答案收敛的题才被采纳为「高置信题」并生成伪标准答案，再拿这批题去评测目标模型。低置信题直接丢弃，宁可少评也不用脏标签。

---

## 目录

- [它能做什么](#它能做什么)
- [流水线](#流水线)
- [快速开始](#快速开始)
- [两种运行模式](#两种运行模式)
- [配置](#配置)
- [提示词](#提示词)
- [评分标准](#评分标准)
- [产出物](#产出物)
- [Excel 列结构](#excel-列结构)
- [换成你自己的模型](#换成你自己的模型)
- [常见问题](#常见问题)
- [目录结构](#目录结构)
- [License](#license)

---

## 它能做什么

- **无标注评测**：用多模型共识生成伪标准答案，不需要人工标注
- **0–3 分四档评分**：由 LLM 裁判逐题打分并给出中文理由，理由完整落表，可复核
- **横向对比**：同一套标准同时给待测模型和 5 个参考模型打分，得到可比基线
- **断点续跑**：每个批次独立目录，任意阶段失败可从该阶段重跑，已完成的行不会重复请求
- **两种入口**：从零跑全流程，或者对已有题集只评一个新模型的结果
- **全链路日志**：每阶段结构化日志 + `progress.jsonl` 进度事件，失败可定位到具体题和具体请求

---

## 流水线

```
CSV(图片URL)
    │
    ├─ Stage 1  OCR 切题          整页试卷 → 单题图                  image/ → image_ocr/
    ├─ Stage 2  清洗分类          剔除无效图，打学段/学科/题型标签    → Excel B–E
    ├─ Stage 3  五模型解题        5 个参考模型各自独立解题            → Excel F–Y
    ├─ Stage 4  共识裁判          交叉比对，产出置信度 + 伪标准答案   → Excel Z–AD
    ├─ Stage 5  图床上传          本地题图 → 公网 URL                 → Excel AE
    ├─ Stage 6  待测模型解题      只跑 confidence=high 的题           → Excel AF–AH
    ├─ Stage 7  评分裁判          待测模型 + 5 参考模型，0–3 分       → Excel AI–AT
    └─ Stage 8  报告生成          指标、分组表现、低分清单、横向对比  → report.md
```

关键设计：

- **Stage 4 是质量闸门**。只有 `confidence=high` 的题会进入 Stage 6/7。这个阈值口径决定了整份评测的可信度。
- **Stage 5 存在的唯一原因**是很多解题接口只接受图片 URL 而不接受 base64。把 `providers.target_model.input_mode` 改成 `image` 就能直接传本地图片，跳过这一阶段，也不用配 BOS。
- **单题失败不阻断**。除 Stage 5 图床是硬依赖外，某个模型某道题失败只会把错误写进表格，流水线继续。
- **Stage 7 的裁判不应与待测模型同源**，否则同源偏好会抬高分数。

---

## 快速开始

需要 Python 3.10 及以上。

```bash
git clone https://github.com/machao101666/k12-solver-eval-skill.git
cd k12-solver-eval-skill

pip3 install -r requirements.txt

cp .env.example .env
# 编辑 .env，填入你的 API 凭证

python3 config.py            # 配置自检：告诉你还缺哪些凭证、提示词能否加载
```

`python3 config.py` 输出示例：

```
[凭证与接入点]
  ✔ providers.openai_compat.base_url  https://...
  ✖ providers.openai_compat.api_key   未设置环境变量 OPENAI_COMPAT_API_KEY
[提示词]
  ✔ solve            1885 字符
  ✔ score_target     5086 字符
[评测口径]
  只评置信度 : high
  达标线     : >= 2 分
  3分格式要求: ==**解题过程**== + ==**最终答案**==
```

自检通过后跑全流程：

```bash
python3 run_pipeline.py --csv input.csv
```

`input.csv` 只需要一列图片 URL。产物在 `runs/run_<时间戳>/`，报告是其中的 `report.md`。

先小样本验证再跑全量。评测是花钱的：每道高置信题至少要 5 次参考模型解题 + 1 次共识 + 1 次待测解题 + 6 次评分。

---

## 两种运行模式

### 模式一：完整流水线

从图片跑到报告，适合首次建立评测集。

```bash
python3 run_pipeline.py --csv input.csv

# 断点续跑与运维
python3 run_pipeline.py --run-dir runs/run_xxx --from-stage 5   # 从第 5 阶段起重跑
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 8   # 只重新生成报告
python3 run_pipeline.py --latest                               # 最近一次批次概况
python3 run_pipeline.py --status runs/run_xxx
python3 run_pipeline.py --watch  runs/run_xxx                  # 实时跟随进度
```

### 模式二：在已有题集上评一个新模型

评测集建好之后，换模型只需要重跑评分。Stage 1–6 全部跳过，直接导入模型输出。

```bash
python3 evaluate_existing_results.py \
  --xlsx        evaluation.xlsx \
  --images-dir  题图目录 \
  --results-jsonl result_a.jsonl result_b.jsonl \
  --output-root  ./out
```

- 按 JSONL 的 `images[0].path` 与 Excel B 列做路径映射，不依赖行序
- 只保留 `confidence=high` 的行，把 `response` 写进 AF，只评待测模型（AI/AJ）
- 每个 JSONL 各自输出独立目录，互不干扰
- **源 Excel、源 JSONL、源题图全程只读**
- 评分不完整时会在生成报告前直接报错，不会产出半成品报告

JSONL 每行格式：

```json
{"response": "模型的完整解题输出", "images": [{"path": "evaluation/xxx.png"}]}
```

中断后续跑：

```bash
python3 evaluate_existing_results.py --resume-run out/xxx --invalid-retries 2
```

---

## 配置

**所有可调项集中在 `config.yaml`，凭证集中在 `.env`，提示词集中在 `prompts/`。正常使用不需要改任何 Python 代码。**

| 我想改…… | 改哪里 |
| --- | --- |
| API 凭证 | `.env` |
| 接入点 / 网关地址 | `config.yaml` → `providers.*.base_url` |
| 待测模型换成别的接口 | `config.yaml` → `providers.target_model.api_style` |
| 用哪些模型 | `config.yaml` → `models.*` |
| 提示词 | `prompts/*.md` |
| 并发、超时、重试 | `config.yaml` → `runtime.*` |
| 置信度口径、达标线 | `config.yaml` → `evaluation.*` |
| Excel 列位置 | `config.yaml` → `excel.columns` |

### 配置分层

```
.env            凭证（已 gitignore，绝不提交）
   ↓ ${ENV_VAR}
config.yaml     所有非密钥配置
   ↓
命令行参数       单次覆盖，如 --concurrency
```

也可以完全不用 `.env`，直接 export 环境变量；或用 `AUTO_EVAL_CONFIG` 指向你自己的配置文件：

```bash
export AUTO_EVAL_CONFIG=/path/to/my-config.yaml
```

### 主要配置块

- **`providers`** — 服务接入点，每个块都带了推荐值。`baidu_ocr`（切题）、`qianfan`、`openai_compat`（OpenAI 兼容网关）、`bos`（图床）、`target_model`（待测模型，含 `api_style` / `input_mode`）
- **`models`** — `classify`（分类）、`reference_models`（参考模型列表，顺序即 M1…M5）、`consensus_judge`（共识裁判）、`score_judge`（评分裁判）
- **`runtime`** — `solve_concurrency`、`consensus_concurrency`、`score_concurrency`、`request_timeout`、`max_retries`、`max_image_bytes`
- **`evaluation`** — `only_confidence`（默认 `high`）、`pass_score`（默认 2）、`invalid_retries`、`required_output_headings`（3 分要求的结构化标题）
- **`excel`** — 列映射，`columns` + `reference_score_columns`

### 在代码里读配置

```python
from config import cfg, secret, prompt, column, require_secrets

cfg("runtime.score_concurrency")            # 普通配置
secret("providers.qianfan.api_key")         # 凭证，缺失时报错并提示要设哪个环境变量
prompt("solve")                             # 提示词文件内容
column("target_score")                      # Excel 列号
require_secrets("providers.bos.bucket")     # 入口处批量前置校验
```

---

## 提示词

6 份提示词全部是 `prompts/` 下的独立 Markdown 文件，改文件即生效：

- `stage2_clean_classify.md` — 图片有效性判定 + 学段/学科/题型分类
- `stage3_solve.md` — 参考模型解题（5 个模型共用，保证横向可比）
- `stage4_consensus.md` — 共识伪标签与置信度判定
- `stage6_target_solve.md` — 待测模型解题，仅 `api_style: openai_compat` 时使用（`baidu_solve` 是产品化接口，只收图片不收提示词）
- `stage7_score_target.md` — 待测模型评分
- `stage7_score_reference.md` — 参考模型横向评分

占位符用 Python `str.format()` 语法，**正文里的字面花括号要写成 `{{` `}}`**。每份文件的占位符清单、输出契约和「改了会影响什么」见 [`prompts/README.md`](prompts/README.md)。

改提示词前请注意：

1. 每份提示词的输出格式都有对应的 Python 解析代码，只改提示词不改解析器会整批失败
2. 改了评分提示词，新旧分数不可比，不要混在同一张横向表里
3. 先用小样本试跑验证解析正常

---

## 评分标准

裁判模型对每道题输出 0–3 分和中文理由：

| 分数 | 含义 |
| --- | --- |
| 3 | 完全可用：答案正确、解析正确完整，**且**结构化输出格式完全合规 |
| 2 | 基本可用：答案正确，解析基本正确，格式不合规或有小瑕疵 |
| 1 | 部分可用：答案正确但解析有明显问题，或解析对但答案错 |
| 0 | 不可用：答案错误，或答非所问、拒答、乱码 |

`pass_score`（默认 2）是达标线，报告里的「≥2 分率」是核心指标。

**3 分有格式门槛**：待测模型输出必须按顺序原样包含 `config.yaml` 里 `evaluation.required_output_headings` 定义的标题（默认 `==**解题过程**==` 和 `==**最终答案**==`）。内容再对，格式不合规最高 2 分。参考模型不受此门槛约束——它们没有约定输出格式。这是刻意的不对称，看横向表时要意识到这一点：参考模型的 3 分率与待测模型的 3 分率口径不同，`≥2 分率`才是完全可比的指标。

改了 `required_output_headings` 必须同步改 `prompts/stage7_score_target.md`，两处描述必须一致。

---

## 产出物

```
runs/run_20260731_120000/
├── input.csv              输入副本
├── evaluation.xlsx        全量数据，54 列，逐题可追溯
├── image/                 原始下载图
├── image_ocr/             切题结果
├── image_ocr_clean/       清洗后有效题图
├── logs/                  分阶段结构化日志
├── progress.jsonl         进度事件流
├── report.md              评测报告
└── run.json               批次元信息与阶段状态
```

`report.md` 章节：

- **核心指标汇总与结论**
- **一、高置信题集概览** — 学段 × 学科 × 题型分布
- **二、待测模型评分结果** — 整体指标、分学段学科表现、分题型 ≥2 分率、耗时、**完整低分题清单**（含完整评分理由与模型原始输出，不截断）
- **三、解题模型横向评分** — 待测模型与 5 个参考模型的分档对比，按 Excel 实际数据计算

单独重新生成报告：

```bash
python3 stage8_generate_report.py evaluation.xlsx --output report.md --title "评测报告"
python3 stage8_generate_report.py evaluation.xlsx --cross-model-xlsx baseline.xlsx   # 横向表用另一份 Excel
```

---

## Excel 列结构

`evaluation.xlsx` 是流水线的中间状态与最终数据源，54 列，每阶段只写自己那几列，因此可以断点续跑。

| 列 | 内容 | 写入方 |
| --- | --- | --- |
| A–B | 题目 id、题图相对路径 | Stage 1 |
| C–E | 学段、学科、题型 | Stage 2 |
| F–Y | 5 个参考模型 × (耗时/模型名/解析/答案) | Stage 3 |
| Z–AD | 置信度、选中模型、共识答案、伪标签、过滤原因 | Stage 4 |
| AE | 题图公网 URL | Stage 5 |
| AF–AH | 待测模型完整输出、耗时、状态 | Stage 6 |
| AI–AJ | 待测模型分数、评分理由 | Stage 7 |
| AK–AT | 5 个参考模型的分数与理由 | Stage 7 |

列号可在 `config.yaml` 的 `excel.columns` 与 `excel.reference_score_columns` 调整。完整字段说明见 [`harness/references/excel-schema.md`](harness/references/excel-schema.md)。

---

## 换成你自己的模型

### 换参考模型（只改配置）

编辑 `config.yaml` 的 `models.reference_models`。增减模型时必须同步调整 `excel_columns` 和 `excel.reference_score_columns`，两个列表长度必须一致，否则启动时会直接报错。

```yaml
models:
  reference_models:
    - label: M1(my-model)
      provider: openai_compat        # 复用已配置的接入点
      model: your-model-id
      temperature: 0.1
      enable_thinking: true
      excel_columns: { time: 6, model: 7, reasoning: 8, answer: 9 }
```

### 换网关（只改配置）

任何 OpenAI 兼容服务（OpenRouter、one-api、自建 vLLM）都只需要改 `providers.openai_compat.base_url` 和对应的环境变量。

### 换待测模型（只改配置）

待测模型接口由 `providers.target_model.api_style` 决定，内置两种风格：

- `baidu_solve`（默认，推荐）——百度智能云 OCR [拍照解题](https://ai.baidu.com/tech/ocr/solve)，[API 文档](https://cloud.baidu.com/doc/OCR/s/Mmqsveul4)。鉴权方式和 Stage 1 一样（API Key + Secret Key 换 `access_token`），公测期个人实名有 200 次免费额度。只填 `TARGET_SOLVE_API_KEY` / `TARGET_SOLVE_SECRET_KEY` 就能跑。
- `openai_compat`——任意 OpenAI 兼容的多模态 chat 接口。填 `TARGET_MODEL_BASE_URL` / `TARGET_MODEL_API_KEY` / `TARGET_MODEL_NAME`，提示词在 `prompts/stage6_target_solve.md`。

```yaml
providers:
  target_model:
    api_style: openai_compat   # baidu_solve | openai_compat
    input_mode: url            # url = 用 AE 列公网 URL；image = 直接读 B 列本地图片
```

`input_mode: image` 时题图直接 base64 上传，**Stage 5 图床上传可以整段跳过**，也就不需要配 BOS。

如果你的待测模型不属于这两种风格（比如自家网关有专属签名），有两条路：在 `stage6_target_solve.py` 里按 `_SOLVERS` 的形状加一个 solver 函数；或者自己批量跑一遍生成 JSONL，用**模式二**只做评分——后者完全不用碰代码。

---

## 常见问题

**为什么只评 high 置信度的题？**
共识不收敛说明这道题本身有歧义、图不清、或超出模型能力，此时伪标准答案不可信，拿它评分只会引入噪声。宁可样本量小，也不要脏标签。想放宽改 `evaluation.only_confidence`，但要清楚代价。

**为什么参考模型也要评分？**
只有待测模型的绝对分数没有参照意义——不知道 78% 的 3 分率是高还是低。用同一套裁判、同一批题给 5 个强模型打分，才能判断待测模型处在什么水位。

**裁判返回的分数解析失败怎么办？**
Stage 7 会自动定向重试 `evaluation.invalid_retries` 次。裁判有时会返回多个 JSON（比如自己「复核更正」），解析器取最后一个合法对象。仍然失败的行会在报告生成前被拦截并报错。

**评测集能复用吗？**
可以，这是推荐做法。Stage 1–5 建好的 `evaluation.xlsx` 是资产，之后每换一个待测模型只跑模式二即可。

**为什么图片路径是相对路径？**
安全考虑。绝对路径和 `..` 会被直接拒绝，防止 JSONL 里的路径穿越到题集目录之外。

**Excel 里 `==**解题过程**==` 被识别成公式了？**
以 `=` 开头的字符串会被 Excel 当公式。写入时已强制 `data_type='s'`，如果你自己改代码写 AF 列，记得也这么做。

**跑起来很慢 / 触发限流？**
调低 `runtime.*_concurrency`。Stage 6 的待测模型并发单独由 `providers.target_model.concurrency` 控制，默认 1，因为 ToB 接口通常有 QPS 限制。

---

## 目录结构

```
.
├── config.yaml                    统一配置（唯一事实来源）
├── config.py                      配置加载器 + 自检入口
├── .env.example                   凭证模板
├── prompts/                       6 份提示词 + 说明
├── run_pipeline.py                完整流水线编排
├── evaluate_existing_results.py   已有题集独立评测入口
├── stage1_ocr_cut_questions.py    ┐
├── stage2_clean_classify.py       │
├── stage3_multi_model_solve.py    │
├── stage4_consensus_judge.py      ├ 八个阶段，均可单独运行
├── stage5_upload_bos_urls.py      │
├── stage6_target_solve.py         │
├── stage7_score_judge.py          │
├── stage8_generate_report.py      ┘
├── logger.py                      结构化日志
├── evaluation.xlsx                空白表模板
├── harness/                       给 AI Agent 用的工程文件
│   ├── AGENT.md                   Agent 行为准则
│   ├── SCENARIOS.md               意图路由
│   ├── skills/                    按意图拆分的能力入口
│   ├── references/                列结构、评分规则、报告结构、playbook
│   └── memory/                    长期偏好与架构决策
└── evals/evals.json               Agent 路由回归用例
```

`harness/` 是可选的：这套脚本本身可以纯命令行使用。如果你在 AI 编程助手（Claude Code、Comate、Cursor 等）里用它，`harness/` 让助手知道该在什么场景下调哪个脚本、口径是什么、失败了怎么查。不需要就整个删掉，不影响脚本运行。

---

## License

MIT © [machao101666](https://github.com/machao101666)

本仓库不含任何模型权重和 API 凭证，所有外部服务需要你自行申请并在 `.env` 中配置。

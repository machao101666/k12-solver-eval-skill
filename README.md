# K12 拍照解题模型自动化评测流水线

面向 K12 作业「拍照解题」大模型的自动化评测工具包。输入试卷题图，自动完成 OCR 切题、有效性分类、多模型解题、共识伪标签、待测模型解题、专业评分裁判，并产出可追溯的评测报告。

> 模型负责理解与判断，脚本负责数据与执行，文档负责流程与规范。本仓库不含任何真实密钥，所有凭证由使用者通过 `.env` 提供。

## 功能特性

- 一条命令跑完整评测：从图片 URL 的 CSV 到最终 Markdown 报告（Stage 1-8）。
- 已有结果独立评测：跳过前置阶段，仅对已有 `evaluation.xlsx` + 题图 + 待测结果 JSONL 打分并各自出报告。
- 每次运行生成独立、可追溯、可续跑的产物包（`runs/run_xxx/`）。
- 支持失败续跑、单阶段运行、报告生成、结果分析。
- 凭证完全外置，源码零硬编码密钥，可安全公开与分发。

## 快速开始

### 1. 环境要求

- Python 3.10+
- 安装依赖：

```
pip3 install -r requirements.txt
```

BOS 上传依赖百度智能云官方 SDK（`bce-python-sdk` + `pycryptodome`），已写入 `requirements.txt`。

### 2. 配置凭证（只需一次）

复制模板并填入你自己的凭证：

```
cp .env.example .env
```

然后编辑 `.env`，等号后直接写值。凭证按用到的阶段划分：

- `BAIDU_OCR_API_KEY` / `BAIDU_OCR_SECRET_KEY`：Stage 1 百度 OCR 切题
- `QIANFAN_API_TOKEN`：Stage 2 有效性分类、Stage 3 五模型解题（只填 token，不带 Bearer 前缀）
- `HUAYAN_API_KEY`：Stage 4 共识裁判必填；Stage 7 选用 huayan 裁判时需要
- `BOS_ACCESS_KEY` / `BOS_SECRET_KEY`：Stage 5 题图上传取 URL
- `XIAODU_APP_KEY` / `XIAODU_APP_SECRET`：Stage 6 小度待测模型解题
- `ONE_API_TOKEN`：Stage 7 默认裁判（oneapi-opus-5）

说明：不同阶段调用不同服务，因此需要多组凭证，但只在 `.env` 一处集中填写。缺失时脚本会给出明确中文提示，指出缺哪个变量、用于哪个阶段。真实环境变量（export）优先级高于 `.env`。API 地址、BOS Endpoint/Bucket 等非敏感项已内置默认值，一般无需改动。`.env` 已被 `.gitignore` 忽略，请勿提交或外传。

### 3. 运行

完整流水线（从图片 URL 的 CSV 开始）：

```
python3 run_pipeline.py --csv <input.csv>
```

已有结果独立评测（跳过 Stage 1-6，只评 high 并出报告）：

```
python3 evaluate_existing_results.py \
  --xlsx <evaluation.xlsx> \
  --images-dir <题图目录> \
  --results-jsonl <result_a.jsonl> [result_b.jsonl ...] \
  --output-root <输出根目录> \
  --concurrency 8
```

常用运维命令：

```
python3 run_pipeline.py --run-dir runs/run_xxx --from-stage 5   # 失败续跑
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 8   # 只跑某阶段
python3 run_pipeline.py --latest                                # 定位最近一次 run
python3 run_pipeline.py --status runs/run_xxx                   # 查看状态
python3 run_pipeline.py --watch runs/run_xxx                    # 实时跟踪
```

## 流水线阶段

- Stage 1 `stage1_ocr_cut_questions.py`：百度 OCR 试卷切题。
- Stage 2 `stage2_clean_classify.py`：有效性判定、学段/学科/题型分类，写入 B-E 列。
- Stage 3 `stage3_multi_model_solve.py`：五个参考模型解题，写入 F-Y 列。
- Stage 4 `stage4_consensus_judge.py`：共识裁判，生成伪标签，写入 Z-AD 列。
- Stage 5 `stage5_upload_bos_urls.py`：题图上传百度 BOS 取 URL，写入 AE 列（BOS 失败会停止流程）。
- Stage 6 `stage6_xiaodu_solve.py`：调用小度待测模型解题，写入 AF/AG 列。
- Stage 7 `stage7_score_judge.py`：待测模型与五参考模型评分，写入 AI-AT 列；默认裁判 oneapi-opus-5。
- Stage 8 `stage8_generate_report.py`：生成 Markdown 评测报告。

## 目录结构

```
.
├── run_pipeline.py              # 总控入口
├── evaluate_existing_results.py # 已有结果独立评测入口
├── stage1_..._solve.py          # Stage 1-8 阶段工具
├── config.py                    # 统一读取凭证（环境变量 / .env）
├── .env.example                 # 凭证模板，复制为 .env 使用
├── logger.py                    # 结构化日志
├── SKILL.md                     # skill 说明（供 Comate/Cursor 等识别）
├── harness/                     # Agent 工程：AGENT.md / SCENARIOS / skills / references / memory
├── runs/                        # 每次运行的产物（xlsx / report.md / jsonl / logs）
├── shared_question_images/      # 题图唯一副本（去重后，161 张）
└── 图片路径说明.md               # 题图去重后的路径映射说明
```

## 运行产物

每次运行生成独立批次目录，核心产物：`report.md`（评测报告）+ `evaluation.xlsx`（明细表格）+ `run.json`（状态）+ `logs/`（结构化日志）。

## 题图说明

为避免每个批次重复保存同一套题图，题图已去重为唯一副本，统一存放于 `shared_question_images/`（按 学段/学科 分类，161 张）。历史 `runs/` 中各批次的 `evaluation/` 题图目录已移除，路径映射见根目录《图片路径说明.md》。

## 安全说明

- 源码不含任何真实密钥，全部凭证由 `.env` 提供，且 `.env` 已被 `.gitignore` 忽略。
- API Key、Token、BOS 密钥应视为配置，不写入报告或对外文档。
- 本包不内置大模型权重；规划与意图理解由外部 LLM 完成，脚本调用配置好的外部模型/API 服务。

## 版本历史

- v2.1：新增较完整的 README。
- v2.0：凭证外置化（`config.py` + `.env`，源码零硬编码密钥）；题图去重为唯一副本 `shared_question_images/`，新增《图片路径说明.md》。
- v1.0：K12 解题模型自动化评测流水线首次开源。


---
name: auto-evaluation
description: >
  K12 作业解题模型自动化评测 Agent 包。既支持从图片 URL CSV 运行 OCR、五模型解题、共识、待测模型解题、评分和报告的完整流水线，也支持对已有 evaluation.xlsx、题图目录和待测结果 JSONL 跳过前置阶段，仅筛选 high 数据、评分待测模型并生成独立报告。用户提到现有解题结果评测、JSONL 打分、只评 high、跳过 OCR/五模型直接评分、双轮结果独立报告时，应优先使用本 skill；同时支持失败续跑、单阶段运行、报告生成和结果分析。
---

# 自动化评测 Agent 包

## 首次使用：配置凭证（只需一次）

本包源码不含任何真实密钥，所有 Token/Key 从环境变量或包根目录的 `.env` 文件读取。使用者拿到本包后，只需复制模板并填入自己的凭证即可运行：

```bash
cp .env.example .env
# 编辑 .env，填入下列凭证（等号后直接写值）
```

需要的凭证（按用到的阶段）：

- `BAIDU_OCR_API_KEY` / `BAIDU_OCR_SECRET_KEY`：Stage 1 百度 OCR 切题。
- `QIANFAN_API_TOKEN`：Stage 2 分类、Stage 3 五模型解题（只填 token，不带 `Bearer `）。
- `HUAYAN_API_KEY`：Stage 4 共识裁判必填；Stage 7 选 huayan 裁判时需要。
- `BOS_ACCESS_KEY` / `BOS_SECRET_KEY`：Stage 5 题图上传取 URL。
- `XIAODU_APP_KEY` / `XIAODU_APP_SECRET`：Stage 6 小度待测模型解题。
- `ONE_API_TOKEN`：Stage 7 默认裁判（oneapi-opus-5）。

说明：不同阶段用到不同服务，因此需要多个凭证，但只需在 `.env` 一处集中填写。缺失时脚本会给出明确的中文提示，指出缺哪个变量、用于哪个阶段。真实环境变量（`export`）优先级高于 `.env`。API 地址、BOS Endpoint/Bucket 等非敏感项已内置默认值，一般无需改动（如需覆盖见 `.env.example` 末尾）。`.env` 已被 `.gitignore` 忽略，请勿提交或外传。

## 推荐入口

```bash
python3 run_pipeline.py --csv <input.csv>
```

已有解题集独立评测入口：

```bash
python3 evaluate_existing_results.py \
  --xlsx <evaluation.xlsx> \
  --images-dir <题图目录> \
  --results-jsonl <result_a.jsonl> [result_b.jsonl ...] \
  --output-root <输出根目录> \
  --concurrency 8
```

该模式跳过 Stage 1-6，按 `images[0].path` 映射 Excel B 列，只保留 `confidence=high` 行，将 JSONL 的 `response` 写入 AF，仅评分待测模型并生成每个 JSONL 各自独立的 `evaluation.xlsx`、筛选后 JSONL、题图、日志和 `report.md`。源 Excel、源 JSONL 和源题图不会被修改。

运行时会按阶段输出 `[PROGRESS] {...}` 进度事件，并写入 `progress.jsonl`。运行后会生成独立批次目录：

```text
runs/run_<YYYYMMDD_HHMMSS>/
├── input.csv
├── evaluation.xlsx
├── image/
├── image_ocr/
├── image_ocr_clean/
├── logs/
├── progress.jsonl
├── report.md
└── run.json
```

最终输出会明确给出报告和表格路径：`report.md` + `evaluation.xlsx`。除 Stage 5 BOS 这类硬依赖外，单个模型调用失败只写入表格，不作为全过程失败。

## 常用命令

```bash
python3 run_pipeline.py --csv <input.csv>
python3 evaluate_existing_results.py --xlsx <evaluation.xlsx> --images-dir <题图目录> --results-jsonl <result.jsonl>
python3 evaluate_existing_results.py --xlsx <evaluation.xlsx> --images-dir <题图目录> --results-jsonl <result_a.jsonl> <result_b.jsonl>
python3 evaluate_existing_results.py --resume-run <已有独立评测目录> --invalid-retries 2
python3 run_pipeline.py --run-dir runs/run_xxx --from-stage 5
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 8
python3 run_pipeline.py --latest
python3 run_pipeline.py --status runs/run_xxx
python3 run_pipeline.py --tail runs/run_xxx
python3 run_pipeline.py --watch runs/run_xxx
```

## 阶段说明

- Stage 1：`stage1_ocr_cut_questions.py`，OCR 试卷切题。
- Stage 2：`stage2_clean_classify.py`，有效性判定、分类并写入 B-E 列；每次模型调用最多重试 1 次，总共 2 次，每次限时 200 秒。
- Stage 3：`stage3_multi_model_solve.py`，五个参考模型解题并写入 F-Y 列；每次模型调用最多重试 1 次，总共 2 次，每次限时 200 秒。
- Stage 4：`stage4_consensus_judge.py`，共识裁判并写入 Z-AD 列；每次模型调用最多重试 1 次，总共 2 次，每次限时 200 秒。
- Stage 5：`stage5_upload_bos_urls.py`，仅使用百度 BOS 上传题图并写入 AE 列；BOS 失败会直接停止流程。
- Stage 6：`stage6_xiaodu_solve.py`，调用小度待测模型并写入 AF/AG 列；单题失败写入表格，不阻断后续流程。
- Stage 7：`stage7_score_judge.py`，待测模型和五参考模型评分并写入 AI-AT 列；默认 `--targets all` 保持完整流水线行为，独立评测使用 `--targets target` 只写 AI/AJ；按单个评分请求并发，默认并发数 8；裁判多 JSON 响应取最后一个合法结果，非法结果按 `--invalid-retries` 定向重试。
- Stage 8：`stage8_generate_report.py`，生成 Markdown 评测报告。

## Agent 工程文件

- `harness/AGENT.md`：Agent 单一事实来源，说明身份、决策流、fallback 和输出规范。
- `harness/SCENARIOS.md`：用户意图与场景路由。
- `harness/skills/`：按意图拆分的能力入口。
- `harness/references/`：Excel 列映射、评分规则、报告结构、API 速查和 playbook。
- `harness/memory/`：长期偏好、架构决策、评测政策和已知问题。

## 运行依赖

至少需要 Python 3.10，并安装脚本用到的依赖：

```bash
pip3 install -r requirements.txt
```

BOS 上传依赖百度智能云官方 Python SDK，安装包名为 `bce-python-sdk`。官方文档要求先安装 `pycrypto`，如果安装失败可使用兼容替代 `pycryptodome`；本包的 `requirements.txt` 默认使用 `pycryptodome` + `bce-python-sdk`。

本包不内置大模型权重；规划和意图理解由 Comate 或外部 LLM 完成，脚本会调用配置好的外部模型/API 服务。
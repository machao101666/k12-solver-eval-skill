---
name: auto-evaluation
description: >
  K12 作业解题模型自动化评测 Agent 包。既支持从试卷图片 URL CSV 运行 OCR、五模型解题、共识、待测模型解题、评分和报告的完整流水线，也支持对已有 evaluation.xlsx、题图目录和待测结果 JSONL 跳过前置阶段，仅筛选 high 数据、评分待测模型并生成独立报告。用户提到现有解题结果评测、JSONL 打分、只评 high、跳过 OCR/五模型直接评分、双轮结果独立报告时，应优先使用本 skill；同时支持失败续跑、单阶段运行、报告生成和结果分析。
---

# 自动化评测 Agent 包

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

该模式跳过 Stage 1-6，按 `images[0].path` 映射 Excel B 列，只保留 `confidence=high` 行（置信度口径由 `config.yaml` 的 `evaluation.only_confidence` 决定），将 JSONL 的 `response` 写入 AF，仅评分待测模型并生成每个 JSONL 各自独立的 `evaluation.xlsx`、筛选后 JSONL、题图、日志和 `report.md`。源 Excel、源 JSONL 和源题图不会被修改。

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

## 配置

所有接入点、模型、并发、阈值和提示词集中在统一配置层，不需要改脚本：

```bash
python3 config.py          # 配置自检：打印缺失的凭证和提示词加载情况
```

- `config.yaml`：单一事实来源。providers（接入点）、models（classify / reference_models / consensus_judge / score_judge）、prompts 文件映射、runtime（solve_concurrency、consensus_concurrency、score_concurrency、request_timeout、max_retries、max_image_bytes）、ocr 轮询参数、evaluation（only_confidence、pass_score、invalid_retries、required_output_headings）、paths、excel 列映射（`excel.columns` + `excel.reference_score_columns`）、logging。
- `config.py`：配置加载器，提供 `cfg("dotted.path")`、`secret(path)`、`soft_secret(path)`、`require_secrets(...)`、`prompt(name)`、`column(name)`。
- `.env`：所有密钥只放这里（模板见 `.env.example`），不写入 `config.yaml`。
- `prompts/`：六个提示词模板已外置为独立文件 `stage2_clean_classify.md`、`stage3_solve.md`、`stage4_consensus.md`、`stage6_target_solve.md`（仅 `api_style: openai_compat` 时使用）、`stage7_score_target.md`、`stage7_score_reference.md`，各文件作用与占位符见 `prompts/README.md`。

换模型或端点改 `config.yaml`；改提示词改 `prompts/`；配密钥改 `.env`。换待测模型改 `providers.target_model.api_style`（内置 `baidu_solve` / `openai_compat`）；只有接口风格不在内置两种之内时，才需要在 `stage6_target_solve.py` 的 `_SOLVERS` 里加一个 solver 函数。

## 阶段说明

- Stage 1：`stage1_ocr_cut_questions.py`，OCR 试卷切题。
- Stage 2：`stage2_clean_classify.py`，有效性判定、分类并写入 B-E 列；每次模型调用最多重试 1 次，总共 2 次，每次限时 200 秒。
- Stage 3：`stage3_multi_model_solve.py`，五个参考模型解题并写入 F-Y 列；每次模型调用最多重试 1 次，总共 2 次，每次限时 200 秒。
- Stage 4：`stage4_consensus_judge.py`，共识裁判并写入 Z-AD 列；每次模型调用最多重试 1 次，总共 2 次，每次限时 200 秒。
- Stage 5：`stage5_upload_bos_urls.py`，仅使用百度 BOS 上传题图并写入 AE 列；BOS 失败会直接停止流程。`providers.target_model.input_mode` 设为 `image` 时本阶段可整段跳过，也不需要配 BOS。
- Stage 6：`stage6_target_solve.py`，调用待测模型解题并写入 AF/AG 列；接口风格由 `providers.target_model.api_style` 决定（`baidu_solve` 为默认，百度智能云 OCR「拍照解题」；`openai_compat` 为任意 OpenAI 兼容多模态接口），题图来源由 `input_mode` 决定（`url` 读 AE 列，`image` 读 B 列本地图片并 base64 上传）；单题失败写入表格，不阻断后续流程。
- Stage 7：`stage7_score_judge.py`，待测模型和五参考模型评分并写入 AI-AT 列；默认 `--targets all` 保持完整流水线行为，独立评测使用 `--targets target` 只写 AI/AJ；按单个评分请求并发，`--concurrency` 默认取 `config.yaml` 的 `runtime.score_concurrency`；裁判多 JSON 响应取最后一个合法结果，非法结果按 `--invalid-retries` 定向重试，默认取 `config.yaml` 的 `evaluation.invalid_retries`。
- Stage 8：`stage8_generate_report.py`，生成 Markdown 评测报告；`--cross-model-xlsx` 可指定横向评分来源 Excel。报告最后一章“三、解题模型横向评分”按 Excel 实际评分列计算，待测模型取 AI 列，五参考模型取 AK/AM/AO/AQ/AS 列，评分列缺失或全空的模型不出现在表中。

上述重试次数和 200 秒超时是 `config.yaml` 的 `runtime.max_retries`、`runtime.request_timeout` 默认值，只处理 `confidence=high` 的口径由 `evaluation.only_confidence` 决定。

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
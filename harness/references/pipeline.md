# Pipeline Reference

## 阶段顺序

```text
Stage 1 -> Stage 2 -> Stage 3 -> Stage 4 -> Stage 5 -> Stage 6 -> Stage 7 -> Stage 8
```

## 推荐入口

```bash
python3 config.py                      # 先做配置自检
python3 run_pipeline.py --csv <input.csv>
```

续跑或单阶段：

```bash
python3 run_pipeline.py --run-dir runs/run_xxx --from-stage 5
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 8
```

现有解题集独立评测：

```bash
python3 evaluate_existing_results.py \
  --xlsx <evaluation.xlsx> \
  --images-dir <题图目录> \
  --results-jsonl <result_a.jsonl> [result_b.jsonl ...] \
  --output-root <输出根目录>
```

独立模式不执行 Stage 1-6。它按图片路径把 JSONL `response` 导入 AF，只保留 high 行（置信度口径见 `config.yaml` 的 `evaluation.only_confidence`），调用 Stage 7 的 `--targets target`，校验每条评分合法后执行 Stage 8。每个 JSONL 生成独立子目录。已有独立目录出现空分或非法评分时，执行 `python3 evaluate_existing_results.py --resume-run <run_dir>`，只补跑异常 high 行。

## 批次产物

每次运行保存到：

```text
runs/run_<YYYYMMDD_HHMMSS>/
```

核心产物：

- `input.csv`
- `evaluation.xlsx`
- `image/`
- `image_ocr/`
- `image_ocr_clean/`
- `logs/`
- `progress.jsonl`
- `report.md`
- `run.json`

## 实时进度事件

`run_pipeline.py` 会在运行时输出 `[PROGRESS] {...}` JSON 事件，并同时追加写入 `runs/run_xxx/progress.jsonl`。

事件包括：

- `run_started`
- `stage_started`
- `stage_completed`
- `stage_failed`
- `stage_partial`
- `run_completed`
- `run_partial`
- `run_failed`

最终输出必须包含：

- 报告：`runs/run_xxx/report.md`
- 表格：`runs/run_xxx/evaluation.xlsx`

## 模型调用失败策略

- Stage 2、Stage 3、Stage 4、Stage 6、Stage 7 的每一次模型/API 网络调用最多重试 1 次，总共 2 次（`runtime.max_retries`）。
- Stage 7 对网络调用成功但无法解析为合法 0-3 分和非空理由的裁判响应，可通过 `--invalid-retries` 做任务级定向重试，默认取 `config.yaml` 的 `evaluation.invalid_retries`。
- 每次网络调用限时由 `runtime.request_timeout` 指定（默认 200 秒），不在自动化流程中动态放宽超时。
- 单题、单模型、单评分请求失败时，把错误写入 `evaluation.xlsx` 对应列，其他成功结果照常保留并继续后续流程。
- Stage 5 BOS 上传在 `providers.target_model.input_mode: url` 时是硬依赖；SDK 不可用、上传失败或对象校验失败时直接停止流程。`input_mode: image` 时 Stage 5 可整段跳过，也不需要配 BOS。

## 执行模式

- 根目录只保留单份 `run_pipeline.py` 和 `stage*.py`。
- run 目录不复制脚本，只保存本次数据和产物。
- 总控脚本向每个 stage 显式传入 `--work-dir <run_dir>`。
- `AUTO_EVAL_WORK_DIR` 作为日志和旧调用兼容兜底。

## 进度查看

```bash
python3 run_pipeline.py --latest
python3 run_pipeline.py --status runs/run_xxx
python3 run_pipeline.py --tail runs/run_xxx
python3 run_pipeline.py --watch runs/run_xxx
```

- `--latest`：查看最近一次 run 的状态。
- `--status`：查看指定 run 的阶段状态、产物和最新日志位置；不传 run 目录时默认最近一次 run。
- `--tail`：输出指定 run 的最近日志；不传 run 目录时默认最近一次 run。
- `--watch`：持续刷新状态和新增日志，直到状态变为 `success`、`failed`、`partial` 或 `dry_run`。

## Stage 7 并发粒度

Stage 7 以单个评分请求为并发粒度，而不是以 Excel 行为粒度。默认 `--targets all` 时，一道 `confidence=high` 题拆成 6 个请求：待测模型评分 + 五参考模型评分；现有解题集独立模式使用 `--targets target`，每题只产生 1 个待测评分请求。所有请求进入同一个线程池，并发数默认取 `config.yaml` 的 `runtime.score_concurrency`，可用 `--concurrency` 覆盖。

总控可调整：

```bash
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 7 --stage7-concurrency <n>
```

直接运行脚本可调整：

```bash
python3 stage7_score_judge.py --work-dir runs/run_xxx --concurrency <n> --targets all
python3 stage7_score_judge.py --work-dir runs/run_xxx --concurrency <n> --targets target --invalid-retries <n>
```

## 手动阶段命令

```bash
python3 stage1_ocr_cut_questions.py <csv> --work-dir runs/run_xxx
python3 stage2_clean_classify.py --work-dir runs/run_xxx
python3 stage3_multi_model_solve.py --work-dir runs/run_xxx
python3 stage4_consensus_judge.py --work-dir runs/run_xxx
python3 stage5_upload_bos_urls.py runs/run_xxx/image_ocr_clean --write-excel --work-dir runs/run_xxx
python3 stage6_target_solve.py --work-dir runs/run_xxx
python3 stage7_score_judge.py --work-dir runs/run_xxx --targets all
python3 stage7_score_judge.py --work-dir runs/run_xxx --targets target
python3 stage8_generate_report.py --work-dir runs/run_xxx
```

Stage 6 脚本是 `stage6_target_solve.py`，其职责是调用待测模型解题；接口风格由 `config.yaml` 的 `providers.target_model.api_style` 决定（`baidu_solve` 默认 / `openai_compat`），题图来源由 `input_mode` 决定（`url` 读 AE 列，`image` 读 B 列本地图片）。换待测模型只改 `api_style`；只有接口风格不在内置两种之内时，才需要在 `stage6_target_solve.py` 的 `_SOLVERS` 里加一个 solver 函数。Stage 6 未显式传入 `--concurrency` 时取 `providers.target_model.concurrency`（默认 1）；Stage 7 未显式传入 `--concurrency`、`--invalid-retries` 时，取 `config.yaml` 的 `runtime.score_concurrency` 和 `evaluation.invalid_retries`。

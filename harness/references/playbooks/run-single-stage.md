# 单阶段运行

## 适用场景

只重跑一个阶段，例如只重新生成报告，或只重跑评分。

## 命令

```bash
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 8
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 7
```

## 注意

- 单阶段运行前确认该阶段依赖的上游产物存在。
- Stage 8 依赖 `evaluation.xlsx`。
- Stage 7 依赖 B-E、Z、AC、AF 以及图片路径可读；`--concurrency`、`--invalid-retries` 未显式指定时取 `config.yaml` 的 `runtime.score_concurrency` 和 `evaluation.invalid_retries`。
- Stage 6（`stage6_target_solve.py`，调用待测模型解题）依赖 Z=high，以及 `providers.target_model.input_mode: url` 时的 AE URL 或 `input_mode: image` 时的 B 列本地题图。
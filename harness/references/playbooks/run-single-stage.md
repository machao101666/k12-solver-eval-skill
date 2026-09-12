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
- Stage 7 依赖 B-E、Z、AC、AF 以及图片路径可读。
- Stage 6 依赖 Z=high 和 AE URL。
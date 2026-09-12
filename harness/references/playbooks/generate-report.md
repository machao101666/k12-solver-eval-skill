# 生成报告

## 命令

```bash
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 8
```

或直接运行：

```bash
python3 stage8_generate_report.py runs/run_xxx/evaluation.xlsx --output runs/run_xxx/report.md
```

## 报告内容

- 核心指标汇总与结论。
- 候选题集概览。
- 多模型解题与共识入集。
- 待测模型评分结果。
- 解题模型横向评分。
- 横向评分低分题清单。

## 明确不包含

- 不生成第五大点“自动优化建议”。
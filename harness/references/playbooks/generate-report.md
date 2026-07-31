# 生成报告

## 命令

```bash
python3 run_pipeline.py --run-dir runs/run_xxx --only-stage 8
```

或直接运行：

```bash
python3 stage8_generate_report.py runs/run_xxx/evaluation.xlsx --output runs/run_xxx/report.md
```

横向评分需要用另一份 Excel 时：

```bash
python3 stage8_generate_report.py runs/run_xxx/evaluation.xlsx --output runs/run_xxx/report.md --cross-model-xlsx <other.xlsx>
```

## 报告内容

- 核心指标汇总与结论。
- 一、高置信题集概览。
- 二、待测模型评分结果，含分学段学科、分题型、待测模型耗时和低分题清单。
- 三、解题模型横向评分：按 Excel 实际评分列计算，待测模型取 AI 列，参考模型取 AK/AM/AO/AQ/AS 列；评分列缺失或全为空的模型不出现在表中。

## 明确不包含

- 不生成第五大点“自动优化建议”。
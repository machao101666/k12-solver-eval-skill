---
name: generate-report
description: 从 evaluation.xlsx 或同结构历史 Excel 生成 Markdown 评测报告。
---

# Generate Report

## 触发意图

用户希望根据表格生成报告、重新生成报告、查看统计结论或只跑 Stage 8。

## 默认入口

已有 run：

```bash
python3 run_pipeline.py --run-dir <run_dir> --only-stage 8
```

直接指定 Excel：

```bash
python3 stage8_generate_report.py <evaluation.xlsx> --output <report.md>
```

## 报告规则

- 支持 `evaluation.xlsx` 和历史同结构 `报告.xlsx`。
- 报告按输入 Excel 中的 high 评测集统计，包含样本分布、待测模型评分、完整低分题原因与原始输出，以及固定六模型横向评分表。
- 不生成“自动优化建议”章节。
- 低分题原因和原始待测输出必须完整保留，不截断或摘要化。
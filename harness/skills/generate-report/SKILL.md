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

横向评分来自另一份 Excel 时追加 `--cross-model-xlsx <xlsx>`。

## 报告规则

- 支持 `evaluation.xlsx` 和历史同结构 `报告.xlsx`。
- 报告按输入 Excel 中的 high 评测集统计，包含样本分布、待测模型评分、完整低分题原因与原始输出，以及最后一章“三、解题模型横向评分”。
- 横向评分表按 Excel 实际评分列计算：待测模型取 AI 列，五个参考模型取 AK/AM/AO/AQ/AS 列；评分列缺失或全为空的模型不出现在表中，不存在固定模型清单。
- 不生成“自动优化建议”章节。
- 低分题原因和原始待测输出必须完整保留，不截断或摘要化。
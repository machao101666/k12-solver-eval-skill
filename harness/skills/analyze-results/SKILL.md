---
name: analyze-results
description: 分析 evaluation.xlsx、report.md 或 run 产物中的评测结果和横向模型表现。
---

# Analyze Results

## 触发意图

用户希望解释结果、比较模型、查看低分题、统计入集率、分析某学段学科题型表现。

## 数据源优先级

1. `runs/run_xxx/evaluation.xlsx`：最权威的结构化明细。
2. `runs/run_xxx/report.md`：可读汇总。
3. `runs/run_xxx/run.json`：运行状态与阶段耗时。
4. `logs/`：失败或异常细节。

## 分析边界

- 待测模型评分 AI/AJ 与五参考模型评分 AK-AT 分开解释。
- 不把非 `confidence=high` 行纳入待测模型评分口径（置信度口径见 `config.yaml` 的 `evaluation.only_confidence`）。
- 不更改评分规则，只解释当前结果。
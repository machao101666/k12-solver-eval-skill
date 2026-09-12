---
name: diagnose-failure
description: 根据 run.json、结构化日志和 Excel 列定位 OCR、BOS、小度、评分或报告失败原因。
---

# Diagnose Failure

## 触发意图

用户询问为什么失败、报错含义、某阶段没有输出、某列为空、接口是否异常。

## 默认排查顺序

1. `runs/INDEX.md`：确认目标 run 和总体状态。
2. `runs/run_xxx/run.json`：确认失败阶段、命令、returncode。
3. `runs/run_xxx/logs/metadata.json` 和 `logs/stages/*.json`：定位阶段内错误。
4. `runs/run_xxx/evaluation.xlsx`：检查关键列是否满足阶段依赖。
5. 必要时检查图片目录、BOS URL 可访问性和接口返回。

## Playbooks

- OCR：`harness/references/playbooks/diagnose-ocr-error.md`
- 小度：`harness/references/playbooks/diagnose-xiaodu-error.md`
- 续跑：`harness/references/playbooks/retry-failed-stage.md`

## 输出

优先给出根因、影响范围和下一步修复命令。
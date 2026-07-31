# 失败后续跑

## 适用场景

某个 run 中途失败，需要从失败阶段或后续阶段继续。

## 步骤

1. 查看 `runs/INDEX.md` 找到失败 run。
2. 打开 `runs/run_xxx/run.json` 确认失败 stage。
3. 查看 `runs/run_xxx/logs/metadata.json` 和对应 `logs/stages/*.json`。
4. 修复问题后从失败阶段继续：

```bash
python3 run_pipeline.py --run-dir runs/run_xxx --from-stage <failed_stage>
```

## 判断是否能跳过前置阶段

- 如果 `evaluation.xlsx` 已有当前阶段需要的输入列，可跳过前置阶段。
- 如果图片目录或 Excel 被删除，不要跳过依赖它们的阶段。
- 如果 Stage 4 的 `confidence` 变化，Stage 5-8 建议全部重跑。
---
name: run-evaluation
description: 从 CSV 输入运行完整 K12 自动化评测流水线，生成独立 run 产物包和报告。
---

# Run Evaluation

## 触发意图

用户希望跑完整评测、批量评测、从图片 URL CSV 生成 `evaluation.xlsx` 和 `report.md`。

## 默认入口

```bash
python3 run_pipeline.py --csv <input.csv>
```

## 行为规则

- 每次新运行必须创建 `runs/run_<timestamp>/`。
- run 目录只保存输入、Excel、图片、日志、报告和 `run.json`，不复制脚本。
- 阶段脚本由根目录单份代码执行，并通过 `--work-dir <run_dir>` 明确指定产物目录。
- 执行后刷新 `runs/INDEX.md`。

## 输出

向用户说明 run 目录、最终状态、`evaluation.xlsx`、`report.md` 和日志路径。
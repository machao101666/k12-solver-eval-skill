---
name: retry-run
description: 诊断已有 run 的失败阶段，并从失败阶段或指定阶段续跑。
---

# Retry Run

## 触发意图

用户希望继续上次运行、从失败处续跑、重跑后续阶段或恢复中断 run。

## 默认入口

```bash
python3 run_pipeline.py --run-dir <run_dir> --from-stage <stage>
```

也可只跑单阶段：

```bash
python3 run_pipeline.py --run-dir <run_dir> --only-stage <stage>
```

## 处理步骤

1. 读取 `runs/INDEX.md` 或用户指定的 `run.json`。
2. 判断失败阶段和已有产物是否满足续跑条件。
3. 必要时先修复输入、配置、URL、Excel 列或日志中暴露的问题。
4. 使用总控脚本续跑，不直接复制脚本到 run 目录。

## 输出

说明续跑起点、依赖检查结果、命令和最新状态。
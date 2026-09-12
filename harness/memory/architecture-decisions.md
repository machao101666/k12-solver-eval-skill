# 架构决策

- 采用单份代码 + `runs/run_xxx/` 批次隔离的 Agent 工程结构。
- `run_pipeline.py` 是默认总控入口，负责创建/续跑 run、调用阶段脚本、写 `run.json`、刷新 `runs/INDEX.md`。
- 长期目标是所有 Stage 支持 `--work-dir`，由根目录脚本直接读写指定 run 目录，不复制脚本。
- `harness/` 存放 Agent 规则、SOP、参考资料和长期经验；不存运行产物。
- 每次运行产物统一存放在 `runs/run_xxx/`：`input.csv`、`evaluation.xlsx`、`image/`、`image_ocr/`、`image_ocr_clean/`、`logs/`、`report.md`、`run.json`。

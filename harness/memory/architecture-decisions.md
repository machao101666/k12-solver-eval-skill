# 架构决策

- 采用单份代码 + `runs/run_xxx/` 批次隔离的 Agent 工程结构。
- `run_pipeline.py` 是默认总控入口，负责创建/续跑 run、调用阶段脚本、写 `run.json`、刷新 `runs/INDEX.md`。
- 长期目标是所有 Stage 支持 `--work-dir`，由根目录脚本直接读写指定 run 目录，不复制脚本。
- `harness/` 存放 Agent 规则、SOP、参考资料和长期经验；不存运行产物。
- 采用统一配置层：`config.yaml` 是接入点、模型、提示词映射、并发和评测口径的单一事实来源，`config.py` 负责加载与自检，凭证只放 `.env`，提示词只放 `prompts/`。
- Stage 6 待测模型接口抽象为 `providers.target_model.api_style`（`baidu_solve` / `openai_compat`）+ `input_mode`（`url` / `image`）：换待测模型只改配置，新增接口风格才在 `stage6_target_solve.py` 的 `_SOLVERS` 里注册新的 solver 函数。
- 每次运行产物统一存放在 `runs/run_xxx/`：`input.csv`、`evaluation.xlsx`、`image/`、`image_ocr/`、`image_ocr_clean/`、`logs/`、`report.md`、`run.json`。

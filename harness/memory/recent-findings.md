# 近期发现

- 当前已新增 Stage 8 报告生成工具，并删除报告第五大点“自动优化建议”。
- 当前已新增 AG 列记录待测模型单次请求耗时。
- 当前已新增 Stage 7 五模型横向评分列 AK-AT。
- 当前已引入统一配置层：`config.yaml`（接入点、模型、提示词映射、并发、评测口径、列映射）、`config.py`（加载与配置自检）、`.env`（凭证）和 `prompts/`（外置提示词）。
- 当前 Stage 8 的“三、解题模型横向评分”按 Excel 实际评分列计算，评分列缺失或全空的模型不出现在表中。
- 当前 Stage 6 已重写为 `stage6_target_solve.py`：接口风格按 `providers.target_model.api_style` 分发（默认 `baidu_solve`，即百度智能云 OCR「拍照解题」；另一种是 `openai_compat`），题图来源按 `input_mode` 选择 `url` 或 `image`，新增提示词 `prompts/stage6_target_solve.md`（仅 `openai_compat` 使用），提示词总数变为 6 份。
- 当前正在将工程改造为 Agent Harness：补充 `harness/`，并推动 `run_pipeline.py` 采用单份代码 + run 产物目录模式。

# 评测口径

- Stage 4 只在 `confidence=high` 时输出可用于后续评测的 `pseudo_label`。
- 入集置信度口径由 `config.yaml` 的 `evaluation.only_confidence` 配置，默认 `high`。
- Stage 5 默认只给 `confidence=high` 的行写入 AE 列 BOS URL；`providers.target_model.input_mode: image` 时不需要 AE。
- Stage 6 只对 `confidence=high` 的行调用待测模型，AF 写完整输出，AG 写单次请求耗时；接口风格由 `providers.target_model.api_style` 决定，不影响评测口径。
- Stage 7 待测模型评分逻辑保持不变：AF 是待测模型输出，AI 写待测模型分数，AJ 写待测模型评分理由。
- Stage 7 同时对五个参考模型做横向评分：AK/AL ernie，AM/AN doubao，AO/AP qwen，AQ/AR claude，AS/AT gemini；列号见 `config.yaml` 的 `excel.reference_score_columns`。
- 达标线、合法分值和非法评分重试次数由 `evaluation.pass_score`、`evaluation.valid_scores`、`evaluation.invalid_retries` 配置。
- Stage 8 同时兼容 `ernie_score_0_3` 和历史拼写 `ernire_score_0_3`。

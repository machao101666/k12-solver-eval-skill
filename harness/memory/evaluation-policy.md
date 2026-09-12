# 评测口径

- Stage 4 只在 `confidence=high` 时输出可用于后续评测的 `pseudo_label`。
- Stage 5 默认只给 `confidence=high` 的行写入 AE 列 BOS URL。
- Stage 6 只对 `confidence=high` 的行调用小度，AF 写完整输出，AG 写单次请求耗时。
- Stage 7 小度评分逻辑保持不变：AF 是待测模型输出，AI 写小度分数，AJ 写小度评分理由。
- Stage 7 同时对五个参考模型做横向评分：AK/AL ernie，AM/AN doubao，AO/AP qwen，AQ/AR claude，AS/AT gemini。
- Stage 8 同时兼容 `ernie_score_0_3` 和历史拼写 `ernire_score_0_3`。

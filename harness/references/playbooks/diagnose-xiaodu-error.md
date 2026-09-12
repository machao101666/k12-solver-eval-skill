# 小度错误诊断

## 常见问题

- AE 为空：Stage 5 未写入 URL，或该行不是 `confidence=high`。
- URL 不可访问：BOS Bucket 权限或 URL 编码问题。
- SSE 为空：小度接口异常、签名问题或图片无法访问。
- 超时：单题耗时超过配置的 `TIMEOUT`。

## 排查步骤

1. 检查 `evaluation.xlsx` 的 Z、AE、AF、AG 列。
2. 只允许 `confidence=high` 行进入 Stage 6。
3. 打开 AE URL 确认图片可访问。
4. 查看 `logs/stages/*xiaodu*.json` 或 `pipeline.log`。
5. 失败时 AF 写 `ERROR:`，AG 写失败前已消耗时间。
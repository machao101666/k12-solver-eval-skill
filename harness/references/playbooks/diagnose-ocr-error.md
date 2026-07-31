# OCR 错误诊断

## 常见错误

- `216630 recognize error`：识别错误，重试；持续出现需提交工单。
- `error_code=18`：QPS/限流，退避重试。
- `error_code=60` 或 `status=running`：任务处理中，继续轮询。
- 没有 `task_id`：检查是否已同步返回 `result.qus_result`。

## 排查步骤

1. 查看 Stage 1 日志：`logs/stages/01_batch.json` 或 `logs/stages/*batch*.json`。
2. 确认原图是否成功下载到 `image/`。
3. 确认 OCR 返回是否包含 `result.qus_result` 或 `task_id`。
4. 如果 `qus_result_num=0`，检查原图质量和接口参数。
5. 如果坐标异常，检查返回 coordinate 格式。
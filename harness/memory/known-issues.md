# 已知问题

- OCR `216630 recognize error` 是识别错误，官方建议重试；持续出现需要提交工单。
- OCR `error_code=60` 或 `status=running/processing/pending` 不是失败，应继续轮询。
- OCR create_task 在 `only_split=true` 时可能直接同步返回 `result.qus_result`，不一定有 `task_id`。
- OCR 坐标可能是 `[x1, y1, x2, y2]`，也可能是 `{x,y,w,h}` 或 left/right 形式，Stage 1 已兼容。
- 小度接口输入必须是可访问 URL，本地路径不能直接传入。
- BOS URL 写入 AE 前应确认 Bucket 公开读或 URL 可被小度访问。

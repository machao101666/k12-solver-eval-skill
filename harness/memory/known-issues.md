# 已知问题

- OCR `216630 recognize error` 是识别错误，官方建议重试；持续出现需要提交工单。
- OCR `error_code=60` 或 `status=running/processing/pending` 不是失败，应继续轮询。
- OCR create_task 在 `only_split=true` 时可能直接同步返回 `result.qus_result`，不一定有 `task_id`。
- OCR 坐标可能是 `[x1, y1, x2, y2]`，也可能是 `{x,y,w,h}` 或 left/right 形式，Stage 1 已兼容。
- 待测模型 `input_mode: url` 时只接受可访问 URL，本地路径不能直接传入；需要直传本地图片就把 `providers.target_model.input_mode` 改成 `image`。
- BOS URL 写入 AE 前应确认 Bucket 公开读或 URL 可被待测模型访问（`input_mode: url` 时才需要 BOS）。
- `api_style: baidu_solve` 的失败信息在响应的 `error_code` / `error_msg` 里；`access_token` 有效期 30 天，脚本内已做进程内缓存，跨进程会重新换取。
- 凭证或端点未配置时脚本会直接报错，先执行 `python3 config.py` 自检，再补齐 `.env`。

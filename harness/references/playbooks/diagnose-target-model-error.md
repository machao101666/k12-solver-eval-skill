# 待测模型错误诊断

## 先确认接口形态

- `config.yaml` 的 `providers.target_model.api_style`：`baidu_solve`（默认，百度智能云 OCR「拍照解题」）或 `openai_compat`。
- `providers.target_model.input_mode`：`url` 走 AE 列公网 URL，`image` 直接读 B 列本地图片。

## 常见问题

- AE 为空：`input_mode: url` 时说明 Stage 5 未写入 URL，或该行不是 `confidence=high`（入集置信度由 `config.yaml` 的 `evaluation.only_confidence` 决定）；`input_mode: image` 时 AE 为空是正常的。
- URL 不可访问：BOS Bucket 权限或 URL 编码问题。
- 本地图片读不到：`input_mode: image` 时 B 列相对路径基于 run 目录解析，文件不存在会直接报「图片文件不存在」；超过 `runtime.max_image_bytes` 的大图会自动转 JPEG 逐步压缩后再上传（与 Stage 4 / Stage 7 同一套策略）。
- 输出为空：待测模型接口异常、鉴权失败或图片无法解析。`api_style: baidu_solve` 时先看响应里的 `error_code` / `error_msg`；`stream_switch: true` 时确认 `result.data_block[].answer_content` 是否有内容。
- 超时：单题耗时超过配置的 `providers.target_model.timeout`。
- 接口地址或密钥缺失：执行 `python3 config.py` 做配置自检。`api_style: baidu_solve` 需要 `TARGET_SOLVE_API_KEY`、`TARGET_SOLVE_SECRET_KEY`；`api_style: openai_compat` 需要 `TARGET_MODEL_BASE_URL`、`TARGET_MODEL_API_KEY`、`TARGET_MODEL_NAME`。

## 排查步骤

1. 检查 `evaluation.xlsx` 的 Z、AE、AF、AG 列（`input_mode: image` 时看 B 列）。
2. 只允许 `confidence=high` 行进入 Stage 6。
3. 确认题图可访问：`input_mode: url` 打开 AE URL；`input_mode: image` 确认 B 列本地图片存在可读。
4. 查看待测模型解题日志 `logs/stages/*target_solve*.json` 或 `pipeline.log`。
5. 失败时 AF 写 `ERROR:`，AG 写失败前已消耗时间。

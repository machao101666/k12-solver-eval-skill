# API Cheatsheet

所有接入点和凭证都由统一配置层提供：端点、模型、并发写在 `config.yaml`，密钥写在 `.env`（见 `.env.example`）。本文只记录配置路径和协议要点，不写具体地址、Bucket 名或密钥。改接入点时改 `config.yaml`，不要改脚本。

## Stage 1 OCR 切题

- 服务：百度 OCR 多模态试卷切题 `paper_cut_edu_vlm`
- 端点：由 `config.yaml` 的 `providers.baidu_ocr.token_url` / `create_task_url` / `get_result_url` 指定（默认值见 config.yaml，均为百度智能云公开 API）
- 官方文档：https://console.bce.baidu.com/ai/#/ai/ocr/overview/index
- 凭证：环境变量 `BAIDU_OCR_API_KEY`、`BAIDU_OCR_SECRET_KEY`（见 .env.example）
- 常见状态：`60/running` 继续轮询，`216630` 识别错误需重试。

## Stage 2/3 千帆

- 端点：由 `config.yaml` 的 `providers.qianfan.base_url` 指定（默认值见 config.yaml）
- 凭证：环境变量 `QIANFAN_API_KEY`（见 .env.example）
- 用途：题目有效性分类、ernie/qwen 解题。

## Stage 3/4/7 OpenAI 兼容网关

- OpenAI 兼容端点：由 `config.yaml` 的 `providers.openai_compat.base_url` 指定，默认留空（`${OPENAI_COMPAT_BASE_URL}`），必须自己填一个，否则 Stage 3/4/7 无法请求
- Anthropic 风格调用：模型配置 `models.reference_models[*].api_style: claude` 时走 Anthropic 分支，使用同一网关的 `/v1/messages` 路径
- 凭证：环境变量 `OPENAI_COMPAT_BASE_URL`、`OPENAI_COMPAT_API_KEY`（见 .env.example）
- 用途：doubao/gemini/claude 解题、共识裁判、评分裁判。
- `models.reference_models[*].api_style` 只接受 `qianfan`、`openai_compat`、`claude` 三种取值。
- 换 OpenRouter、one-api、自建 vLLM 等任意 OpenAI 兼容服务时，只改 `base_url` 和该环境变量。

## Stage 5 BOS

- 用途：将本地题图上传为待测模型可访问 URL；仅在 `providers.target_model.input_mode: url` 时需要，改成 `image` 时整段跳过。
- 配置：`providers.bos.endpoint`、`providers.bos.bucket`、`providers.bos.local_root_name`；凭证为环境变量 `BOS_*`（见 .env.example）。文档中不记录实际 Bucket 名。
- 官方 Python SDK 安装包：`bce-python-sdk`。
- 官方依赖说明：安装 SDK 前先装 `pycrypto`；如安装失败，使用 `pycryptodome`。
- 关键导入：`BceClientConfiguration`、`BceCredentials`、`BosClient`。
- 上传方法：`client.put_object_from_file(bucket, object_key, local_path)`。
- 上传后校验：使用 `client.list_objects(bucket, prefix=object_key)` 并精确匹配返回对象的 key，避免依赖不同 SDK 版本可能不存在的 `get_object_meta`。
- 路径规则：`image_ocr_clean/高中/数学/1-选择题.png` -> `<批次时间>/高中/数学/1-选择题.png`。
- 脚本策略：只使用 BOS SDK；SDK 不可用、上传失败或上传后对象校验失败时 Stage 5 直接失败退出，不使用其他命令行工具降级，不手动写 AE。
- 时间目录规则：单次 Stage 5 只生成一个 `upload_batch_dir`，上传 object_key 和写入 AE 的 URL 都基于这个同一个 object_key，避免上传与读取使用不同时间导致 `NoSuchKey`。

## Stage 6 待测模型

- 接口风格：由 `config.yaml` 的 `providers.target_model.api_style` 决定，内置 `baidu_solve`（默认、推荐）和 `openai_compat` 两种。
- 题图来源：由 `providers.target_model.input_mode` 决定，`url` 读 Excel AE 列公网 URL（需先跑 Stage 5），`image` 直接读 Excel B 列本地图片并 base64 上传（可跳过 Stage 5，也不需要 BOS）。
- 换成自己的待测模型时只改 `providers.target_model.api_style`；只有接口风格不在内置两种之内时，才需要在 `stage6_target_solve.py` 的 `_SOLVERS` 里加一个 solver 函数。

### api_style: baidu_solve（百度智能云 OCR「拍照解题」）

- 端点：由 `config.yaml` 的 `providers.target_model.baidu_solve.token_url` / `solve_url` 指定（默认值见 config.yaml，均为百度智能云公开 API）
- 产品页：https://ai.baidu.com/tech/ocr/solve
- API 文档：https://cloud.baidu.com/doc/OCR/s/Mmqsveul4
- 认证：API Key + Secret Key，`grant_type=client_credentials` 换 `access_token`（有效期 30 天，脚本内做进程内缓存），调用时以 query 参数 `access_token` 下发
- 请求：`Content-Type: application/json`，body 中 `image`（纯 base64）与 `url` 二选一，另有 `disable_preprocess`（图片矫正与增强）、`stream_switch`（false=一次性返回完整结果，true=SSE 分块返回）
- 响应：`log_id`、`error_code`、`error_msg`；非流式取 `result.content`，流式取 `result.data_block[].answer_content` 并由脚本自动拼接
- 凭证：环境变量 `TARGET_SOLVE_API_KEY`、`TARGET_SOLVE_SECRET_KEY`（见 .env.example）
- 配额：公测期个人实名 200 次、企业实名 500 次免费额度

### api_style: openai_compat（任意 OpenAI 兼容多模态 chat 接口）

- 端点与模型：由 `providers.target_model.openai_compat.base_url` / `model` 指定
- 凭证：环境变量 `TARGET_MODEL_BASE_URL`、`TARGET_MODEL_API_KEY`、`TARGET_MODEL_NAME`（见 .env.example）
- 提示词：`prompts/stage6_target_solve.md`（无占位符，只喂图片和输出格式要求）
- 响应：OpenAI 兼容 chat completion，取首个 choice 的 message content

## 调用可靠性策略

- Stage 2/3/4/6/7 每个模型调用最多重试 1 次，总共 2 次（`runtime.max_retries`）。
- 每次调用限时由 `runtime.request_timeout` 指定（默认 200 秒），不能在自动化流程中自动增加重试次数或修改时间阈值。
- 模型调用失败写入 Excel 对应输出/原因列，不代表全过程失败；其余模型和后续阶段继续写入。
- 接入点或凭证异常时，先执行 `python3 config.py` 做配置自检，确认缺失项后再排查接口。

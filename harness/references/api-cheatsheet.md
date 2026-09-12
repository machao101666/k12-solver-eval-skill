# API Cheatsheet

## Stage 1 OCR 切题

- 服务：百度 OCR 多模态试卷切题 `paper_cut_edu_vlm`
- create_task：`/rest/2.0/ocr/v1/paper_cut_edu_vlm/create_task`
- get_result：`/rest/2.0/ocr/v1/paper_cut_edu_vlm/get_result`
- 常见状态：`60/running` 继续轮询，`216630` 识别错误需重试。

## Stage 2/3 千帆

- 端点：`https://qianfan.baidubce.com/v2/chat/completions`
- 用途：题目有效性分类、ernie/qwen 解题。

## Stage 3/4/7 华研

- OpenAI 兼容端点：`https://www.huayanapi.com/v1/chat/completions`
- Anthropic 原生端点：`https://www.huayanapi.com/v1/messages`
- 用途：doubao/gemini/claude 解题、共识裁判、评分裁判。

## Stage 5 BOS

- 用途：将本地题图上传为小度可访问 URL。
- 官方 Python SDK 安装包：`bce-python-sdk`。
- 官方依赖说明：安装 SDK 前先装 `pycrypto`；如安装失败，使用 `pycryptodome`。
- 关键导入：`BceClientConfiguration`、`BceCredentials`、`BosClient`。
- 上传方法：`client.put_object_from_file(bucket, object_key, local_path)`。
- 上传后校验：使用 `client.list_objects(bucket, prefix=object_key)` 并精确匹配返回对象的 key，避免依赖不同 SDK 版本可能不存在的 `get_object_meta`。
- 路径规则：`image_ocr_clean/高中/数学/1-选择题.png` -> `<批次时间>/高中/数学/1-选择题.png`。
- 脚本策略：只使用 BOS SDK；SDK 不可用、上传失败或上传后对象校验失败时 Stage 5 直接失败退出，不使用 dodo_cli 降级，不手动写 AE。
- 时间目录规则：单次 Stage 5 只生成一个 `upload_batch_dir`，上传 object_key 和写入 AE 的 URL 都基于这个同一个 object_key，避免上传与读取使用不同时间导致 `NoSuchKey`。

## Stage 6 小度

- 端点：`https://xiaodu.baidu.com/edu-gateway/gateway/openapi/homework-tob/solve`
- 认证：AppKey + AppSecret HMAC-SHA256
- 响应：SSE，需拼接完整 content。

## 调用可靠性策略

- Stage 2/3/4/6/7 每个模型调用最多重试 1 次，总共 2 次。
- 每次调用限时 200 秒，不能在自动化流程中自动增加重试次数或修改时间阈值。
- 模型调用失败写入 Excel 对应输出/原因列，不代表全过程失败；其余模型和后续阶段继续写入。

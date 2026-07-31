# 五模型横向评分

## 目标

Stage 7 除待测模型外，还对 Stage 3 的五个参考模型输出进行横向评分。

## 输入列

- ernie：H reasoning + I answer
- doubao：L reasoning + M answer
- qwen：P reasoning + Q answer
- claude：T reasoning + U answer
- gemini：X reasoning + Y answer

## 输出列

- AK/AL：ernie score / reason
- AM/AN：doubao score / reason
- AO/AP：qwen score / reason
- AQ/AR：claude score / reason
- AS/AT：gemini score / reason

## 注意

- 待测模型评分的提示词和输入输出逻辑不变，见 `prompts/stage7_score_target.md`。
- 五模型评分使用独立评分提示词 `prompts/stage7_score_reference.md`。
- 输出列号在 `config.yaml` 的 `excel.reference_score_columns` 中定义。
- 评分只对 `confidence=high` 行执行，置信度口径由 `evaluation.only_confidence` 配置。
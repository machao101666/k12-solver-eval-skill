# 五模型横向评分

## 目标

Stage 7 除小度外，还对 Stage 3 的五个参考模型输出进行横向评分。

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

- 小度评分提示词和输入输出逻辑不变。
- 五模型评分使用用户提供的独立评分提示词。
- 评分只对 `confidence=high` 行执行。
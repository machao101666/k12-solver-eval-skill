# Excel Schema

`evaluation.xlsx` 是流水线主数据总线。列号为 1-based。

| 列 | 字母 | 内容 | 写入阶段 |
|---|---|---|---|
| 2 | B | 图片路径 | Stage 2 |
| 3 | C | 学段 | Stage 2 |
| 4 | D | 学科 | Stage 2 |
| 5 | E | 题型 | Stage 2 |
| 6-9 | F-I | 模型1 time/model/reasoning/answer | Stage 3 |
| 10-13 | J-M | 模型2 time/model/reasoning/answer | Stage 3 |
| 14-17 | N-Q | 模型3 time/model/reasoning/answer | Stage 3 |
| 18-21 | R-U | 模型4 time/model/reasoning/answer | Stage 3 |
| 22-25 | V-Y | 模型5 time/model/reasoning/answer | Stage 3 |
| 26 | Z | confidence | Stage 4 |
| 27 | AA | selected_ref_id | Stage 4 |
| 28 | AB | consensus_answer | Stage 4 |
| 29 | AC | pseudo_label | Stage 4 |
| 30 | AD | filter_reason | Stage 4 |
| 31 | AE | 待测模型输入 URL | Stage 5 |
| 32 | AF | 待测模型完整输出；完整流水线中为小度，独立模式中来自结果 JSONL | Stage 6 / 独立评测导入 |
| 33 | AG | 待测模型单次请求耗时（秒）；JSONL 未提供时为空 | Stage 6 |
| 35 | AI | 待测模型 score_0_3 | Stage 7 |
| 36 | AJ | 待测模型 judge_reason | Stage 7 |
| 37-46 | AK-AT | 五模型横向评分与理由 | Stage 7 |

## 关键规则

- Stage 6/7 默认只处理 `confidence=high` 行。
- Stage 8 读取小度评分和五模型横向评分生成报告。
- Stage 8 兼容 `ernie_score_0_3` 和历史 `ernire_score_0_3`。

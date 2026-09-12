---
name: evaluate-existing-results
description: 对已有 evaluation.xlsx、题图目录和一个或多个待测结果 JSONL 单独执行高置信题评分并生成独立报告，也支持在已有独立评测目录中只补跑空分或非法评分。用户提到现有解题结果评测、JSONL 打分、只评 high、跳过 OCR/五模型解题直接评分、两轮结果分别出报告、只补跑异常评分时使用；不要路由到完整 CSV 流水线。
---

# Evaluate Existing Results

## 触发意图

用户已有以下产物，希望跳过 OCR、分类、五模型解题、共识、BOS 和待测模型解题，只对现有待测答案评分并生成报告：

- `evaluation.xlsx`：包含 B-E、Z、AC 等评分输入；
- 题图目录：Excel B 列和 JSONL `images[0].path` 可映射到实际图片；
- 一个或多个结果 JSONL：每条记录包含非空 `response` 和 `images[0].path`。

## 默认入口

```bash
python3 evaluate_existing_results.py \
  --xlsx <evaluation.xlsx> \
  --images-dir <题图目录> \
  --results-jsonl <result_a.jsonl> [result_b.jsonl ...] \
  --output-root <输出根目录> \
  --concurrency 8
```

已有独立评测目录只补跑异常评分：

```bash
python3 evaluate_existing_results.py \
  --resume-run <已有独立评测目录> \
  --concurrency 8 \
  --invalid-retries 2
```

## 默认行为

- `--resume-run` 只请求 AI/AJ 分数非法或理由为空的 high 行，已有合法评分不重复调用。
- 只处理 `confidence=high`，不把 medium/low 加入评测分母。
- 按图片路径映射，不依赖 JSONL 行号或 Excel 行号相同。
- 每个 JSONL 生成独立子目录，互不覆盖。
- 只评待测模型，Stage 7 使用 `--targets target`，不重复评分五参考模型。
- 3 分格式只认可按顺序、原样出现 `==**解题过程**==` 和 `==**最终答案**==`。
- 裁判返回多个 JSON 时取最后一个完整评分；非法评分自动定向重试。
- 161 条 high 必须得到 161 条合法 0-3 分和非空理由后才生成报告。
- 不修改源 Excel、源 JSONL 和源图片。

## 输出

每个 JSONL 的独立目录包含：

```text
<result_stem>_high_<count>/
├── evaluation.xlsx
├── <result_stem>_high_<count>.jsonl
├── evaluation/
├── logs/
├── report.md
└── run.json
```

## 不适用场景

- 用户提供图片 URL CSV 并要求从 OCR 开始：使用 `run-evaluation`。
- 用户已有评分，只要求重新生成报告：使用 `generate-report`。
- 用户要求恢复完整 run 的失败阶段：使用 `retry-run`。

详细检查和异常处理见 `harness/references/playbooks/evaluate-existing-results.md`。

# 评测现有解题结果

## 适用场景

已有 `evaluation.xlsx`、本地题图和待测结果 JSONL，希望只评 high 数据并直接生成报告，不重复运行 Stage 1-6。

## 输入检查

1. Excel 至少包含 B-E、Z、AC、AF、AI、AJ 列。
2. Z 列至少有一条 `high`，每条 high 的 B 和 AC 非空。
3. JSONL 每条记录包含非空 `response` 和 `images[0].path`。
4. Excel B 与 JSONL 图片路径允许一个带 `evaluation/` 前缀、另一个不带；规范化后必须一一对应。
5. 题图目录中必须存在每条 high 对应的图片。

## 单轮命令

```bash
python3 evaluate_existing_results.py \
  --xlsx <evaluation.xlsx> \
  --images-dir <evaluation目录> \
  --results-jsonl <result.jsonl> \
  --output-root <output目录> \
  --concurrency 8
```

## 多轮命令

```bash
python3 evaluate_existing_results.py \
  --xlsx <evaluation.xlsx> \
  --images-dir <evaluation目录> \
  --results-jsonl <result_1.jsonl> <result_2.jsonl> \
  --output-root <output目录> \
  --concurrency 8
```

每个 JSONL 按文件名生成独立子目录；不要把不同轮次写进同一个 `evaluation.xlsx`。

## 只检查映射

```bash
python3 evaluate_existing_results.py \
  --xlsx <evaluation.xlsx> \
  --images-dir <evaluation目录> \
  --results-jsonl <result.jsonl> \
  --output-root <output目录> \
  --prepare-only
```

该模式不调用评分接口，也不生成报告，用于确认 high 数量、JSONL 映射和图片完整性。

## 只补跑异常评分

```bash
python3 evaluate_existing_results.py \
  --resume-run <已有独立评测目录> \
  --concurrency 8 \
  --invalid-retries 2
```

该模式扫描 high 行，只为 AI 非 0-3 整数或 AJ 为空的题目创建请求；已有合法评分保持不变。补跑后重新校验全部 high，并覆盖生成该目录的 `report.md`。

## 评分行为

- 独立模式调用 Stage 7 的 `targets=target`，每题只产生一个待测评分请求。
- 默认并发 8。
- API 网络错误沿用单次调用重试策略。
- 返回缺少合法 `score_0_3` 或 `judge_reason` 时，按 `--invalid-retries` 额外重试。
- 一个响应含多个评分 JSON 时，采用最后一个完整 JSON，适配裁判“复核更正”的返回。
- 最终任一 high 行不是 0-3 整数或理由为空，整轮停止，不生成误导性报告。

## 输出检查

1. `evaluation.xlsx` 行数等于 high 数量加表头。
2. AF 全部是普通文本，不能被 Excel 识别成公式。
3. AI 全部为 0、1、2、3，AJ 全部非空。
4. 题图数量等于 high 数量。
5. 筛选后 JSONL 数量等于 high 数量。
6. `run.json` 的状态为 `success`。
7. `report.md` 存在且统计分母等于 high 数量。

## 路由边界

- 从 CSV 和 OCR 开始的完整评测不使用本入口。
- 已有 AI/AJ、仅重新生成报告时直接执行 Stage 8。
- 仅分析现有结果而不执行评分时使用 `analyze-results`。

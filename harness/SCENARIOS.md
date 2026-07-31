# 场景与提问说明书

本文用于把用户自然语言请求路由到稳定的 Agent 行为。先判断用户意图，再选择对应 Skill、Playbook 和脚本；不要只按关键词机械匹配。

## 场景一：完整评测

用户表达：跑完整流程、批量评测、从 CSV 生成报告、帮我跑一下评测。

需要信息：输入 CSV 路径；如未提供，先确认文件。

默认动作：使用 `harness/skills/run-evaluation/SKILL.md`，执行 `python3 run_pipeline.py --csv <csv>`。

输出给用户：run 目录、当前状态、关键产物路径、失败阶段或最终报告路径。

## 场景二：现有解题集独立评测

用户表达：已有解题结果 JSONL 要评分、只评 high、跳过 OCR 和五模型解题、对两轮结果分别生成报告、在现有解题集上单独执行待测模型评测、只补跑已有目录中的空分或非法评分。

需要信息：`evaluation.xlsx` 路径、题图目录、一个或多个结果 JSONL；输出目录可选。

默认动作：使用 `harness/skills/evaluate-existing-results/SKILL.md`，执行 `python3 evaluate_existing_results.py --xlsx <xlsx> --images-dir <dir> --results-jsonl <jsonl...>`。

输出给用户：每轮独立目录、high 数量、评分完整性、`evaluation.xlsx`、筛选后 JSONL、日志和 `report.md` 路径。

不得回退到 CSV 完整流水线，不重复执行 Stage 1-6，不把五参考模型评分任务混入待测模型独立评测。

## 场景三：失败续跑

用户表达：继续跑、从失败处继续、重跑后续、查看失败后怎么处理。

需要信息：run 目录；如果未指定，先查 `runs/INDEX.md` 找最近失败 run。

默认动作：查看 `run.json` 和日志，定位失败阶段，修复后执行 `python3 run_pipeline.py --run-dir <run_dir> --from-stage <stage>`。

输出给用户：失败原因、续跑起点、是否覆盖已有产物。

## 场景四：单阶段运行

用户表达：只切题、只清洗分类、只上传 BOS、只跑待测模型解题、只评分、只生成报告。

需要信息：run 目录和阶段号；如果是 Stage 1 还需要 CSV。

默认动作：优先用 `run_pipeline.py --only-stage <n>`，不要绕过总控，除非用户明确要求直接运行脚本。

输出给用户：该阶段依赖是否满足、执行命令、更新的 Excel 列或文件。

## 场景五：只生成报告

用户表达：根据表格生成报告、重新生成报告、报告不要某个部分。

需要信息：Excel 路径或 run 目录。

默认动作：使用 `harness/skills/generate-report/SKILL.md`，执行 Stage 8。报告不得生成“自动优化建议”章节。

输出给用户：报告路径、统计口径、如缺少评分列则说明降级统计。

## 场景六：OCR 诊断

用户表达：切题失败、OCR 报错、216630、没有切出图片、坐标异常。

默认动作：使用 `harness/references/playbooks/diagnose-ocr-error.md`。先看 Stage 1 日志和原图，再判断是否重试或需要接口侧排查。

输出给用户：错误码含义、复现行、下一步处理。

## 场景七：待测模型诊断

用户表达：待测模型没输出、AE/AF/AG 列异常、URL 不可访问、返回内容为空、耗时问题。

默认动作：使用 `harness/references/playbooks/diagnose-target-model-error.md`。先确认 `providers.target_model.api_style` 和 `input_mode`，再检查 confidence、AE URL（`input_mode: url` 时）或 B 列本地图片（`input_mode: image` 时）、AF 输出、AG 耗时和 Stage 6 日志。

输出给用户：是否为输入缺失、接口失败、超时或非 high 行跳过。

## 场景八：横向评分分析

用户表达：其他五个模型有没有评分、横向对比、谁表现更好、低分题有哪些。

默认动作：确认 Stage 7 已写 AK-AT，再用 Stage 8 报告或直接读 Excel 汇总。

输出给用户：待测模型评分与五参考模型评分分开说明，不混用口径。

## 场景九：历史 run 查询

用户表达：查看之前运行、最近一次在哪里、哪些产物可以下载、现在进度如何。

默认动作：查看 `runs/INDEX.md`、目标 run 的 `run.json`、`logs/metadata.json`。

输出给用户：run 状态、完成阶段、核心产物、失败点。

## 暂不支持或需降级

- 没有 CSV/Excel/图片输入时，不能凭空执行评测。
- API 密钥无效、配额不足、BOS 权限不可用时，只能定位问题，不能绕过接口；先用 `python3 config.py` 自检配置。
- 评分规则、提示词、入集口径变更必须由用户明确确认；确认后改 `config.yaml` 或 `prompts/`，不改 stage 脚本。
- 非 `confidence=high` 的样本默认不进入待测模型解题和评分，该口径由 `config.yaml` 的 `evaluation.only_confidence` 配置。
# K12 自动化评测 Agent

## 版本与定位

- 版本：2026-06-17
- 定位：这是 K12 作业解题模型评测的 Harness Agent 工程说明，也是本目录的单一事实来源。
- 核心原则：模型负责理解与判断，工具负责数据与执行，文档负责流程与规范。

Harness 不是大模型本身，也不是某一个脚本；它是让大模型稳定调度 Stage 1-8、沉淀规则、管理产物、诊断失败和生成报告的外层工程结构。

## 目标

- 从 CSV 图片 URL 自动完成切题、清洗分类、五模型解题、共识伪标签、BOS URL、小度解题、评分和报告生成。
- 对已有 `evaluation.xlsx`、题图和待测结果 JSONL 跳过 Stage 1-6，只评 high 待测输出并为每个 JSONL 生成独立报告。
- 每次运行形成独立、可追溯、可续跑的产物包。
- 让用户可以用自然语言触发完整评测、单阶段运行、失败续跑、报告生成和结果分析。
- 保持评分口径、Excel 列约定和报告结构稳定。

## 目录职责

```text
harness/
├── AGENT.md                  # 本文件，Agent SSOT
├── SCENARIOS.md              # 用户意图与问题场景路由
├── skills/                   # 按用户意图拆分的能力入口
├── references/               # 稳定事实、列映射、API 速查、报告/评分规则
│   └── playbooks/            # 复杂 SOP、失败诊断和续跑手册
└── memory/                   # 长期偏好、架构决策、已知问题和近期发现
```

根目录脚本仍是确定性工具层：

- `run_pipeline.py`：总控入口。
- `stage1_ocr_cut_questions.py` 到 `stage8_generate_report.py`：阶段工具。
- `logger.py`：结构化日志工具。
- `evaluation.xlsx`：默认模板。
- `runs/`：每次运行的产物目录。

## 运行前置

- 完整流水线输入 CSV 第一列必须是原始图片 URL，首次运行必须存在 `evaluation.xlsx` 模板。
- 现有解题集独立评测不要求 CSV，但必须提供已有 `evaluation.xlsx`、题图目录和至少一个结果 JSONL。
- Stage 6/7 和现有解题集独立评测默认只处理 `confidence=high` 行。
- API Key、Token、BOS 密钥应视为配置，不应写入报告或对外文档。
- 所有凭证从环境变量或包根目录 `.env` 读取（模板见 `.env.example`，加载逻辑见 `config.py`）；源码不含真实密钥。首次使用需 `cp .env.example .env` 并填入凭证，缺失时脚本报错会指明缺哪个变量、用于哪个阶段。
- run 目录不复制脚本；脚本从根目录单份代码执行，产物通过 `--work-dir <run_dir>` 和 `AUTO_EVAL_WORK_DIR` 定位。

## 主决策流

1. 识别用户意图：先读 `SCENARIOS.md`，再选择 `harness/skills/<intent>/SKILL.md`。
2. 判断是否需要执行工具：能通过脚本完成的，不只停留在解释。
3. 选择入口：完整/续跑/单阶段优先使用 `run_pipeline.py`；已有 Excel、题图和结果 JSONL 时使用 `evaluate_existing_results.py`；只有用户明确要求或调试时才直接调用 stage 脚本。
4. 检查依赖：确认上游产物、Excel 列和图片目录满足当前阶段。
5. 执行并记录：每个 run 更新 `run.json`、`logs/` 和 `runs/INDEX.md`。
6. 输出结论：告诉用户 run 目录、完成阶段、核心产物、失败原因和下一步。

## 意图到执行映射

- 完整评测：`harness/skills/run-evaluation/SKILL.md`，执行 `python3 run_pipeline.py --csv <csv>`。
- 现有解题集独立评测：`harness/skills/evaluate-existing-results/SKILL.md`，执行 `python3 evaluate_existing_results.py --xlsx <xlsx> --images-dir <dir> --results-jsonl <jsonl...>`。
- 失败续跑：`harness/skills/retry-run/SKILL.md`，执行 `python3 run_pipeline.py --run-dir <run_dir> --from-stage <stage>`。
- 单阶段运行：参考 `harness/references/playbooks/run-single-stage.md`，执行 `--only-stage <stage>`。
- 只生成报告：`harness/skills/generate-report/SKILL.md`，执行 Stage 8。
- 故障诊断：`harness/skills/diagnose-failure/SKILL.md`，先读 `run.json` 和日志，再看 Excel。
- 结果分析：`harness/skills/analyze-results/SKILL.md`，优先读取 `evaluation.xlsx`。

## Fallback 规则

- 完整评测缺输入 CSV：先让用户提供或定位 CSV；如果已有 Excel、题图和结果 JSONL，应切换到现有解题集独立评测，不要求 CSV。
- 独立评测缺任一输入：明确指出缺少 Excel、题图目录或结果 JSONL，不回退到完整流水线。
- 缺 run 目录：先查 `runs/INDEX.md` 找最近 run。
- 缺上游产物：不要强行跳到下游阶段，提示需要补跑的阶段。
- 接口错误：先看错误码、日志和输入可访问性，再决定重试或停止。
- Excel 列不全：按 `references/excel-schema.md` 检查列名和列号，必要时只生成降级报告。
- 评分规则不明：保持现有规则，不自行更改提示词或打分口径。

## 输出规范

- 对用户报告执行结果时，优先给出：状态、run 目录、关键产物、失败点。
- 引用文件使用路径，例如 `runs/run_xxx/run.json`、`runs/run_xxx/report.md`。
- 诊断问题时先说根因，再说证据和修复方式。
- 报告生成不得包含“自动优化建议”章节。
- 小度评分和五参考模型横向评分必须分开说明。

## 参考文档索引

- 流水线：`harness/references/pipeline.md`
- Excel 列映射：`harness/references/excel-schema.md`
- API 速查：`harness/references/api-cheatsheet.md`
- 评分规则：`harness/references/score-rubric.md`
- 报告结构：`harness/references/report-schema.md`
- 完整评测 SOP：`harness/references/playbooks/run-full-pipeline.md`
- 失败续跑 SOP：`harness/references/playbooks/retry-failed-stage.md`
- OCR 诊断：`harness/references/playbooks/diagnose-ocr-error.md`
- 小度诊断：`harness/references/playbooks/diagnose-xiaodu-error.md`

## Memory 系统

`harness/memory/` 保存不适合写死到脚本、但会影响后续判断的信息：

- 用户偏好：例如 run 目录不复制脚本、报告不生成第五大点。
- 架构决策：例如单份代码 + work-dir 产物隔离。
- 评测政策：例如只处理 `confidence=high`，小度计时口径。
- 已知问题：例如接口错误码、历史容易误判的环节。
- 近期发现：短期仍有参考价值、但后续可能过期的信息。

Memory 不是日志，不存每次运行明细；运行明细属于 `runs/run_xxx/`。

## Skill 清单

- `run-evaluation`：从 CSV 开始的完整评测。
- `evaluate-existing-results`：对已有 Excel、题图和结果 JSONL 只评 high 待测输出并生成报告。
- `retry-run`：失败续跑和指定阶段恢复。
- `diagnose-failure`：接口、Excel、日志和阶段失败诊断。
- `generate-report`：报告生成。
- `analyze-results`：结果解释和横向对比。

## 架构决策进度

- 已采用：根目录单份脚本 + `runs/run_xxx/` 产物隔离。
- 已采用：`run_pipeline.py` 作为统一调度入口。
- 已采用：Stage 8 报告生成工具化。
- 已采用：Stage 7 默认同时评分待测模型与五个参考模型，也可用 `--targets target` 只评待测模型。
- 已采用：现有解题集独立评测入口按图片路径导入 JSONL、只筛 high、自动补跑非法裁判响应并独立出报告。
- 已采用：Agent 文件按 `AGENT.md`、`SCENARIOS.md`、`skills/`、`references/`、`memory/` 拆分。
- 待持续优化：API 密钥配置化，避免硬编码在脚本里。

## 修订记录

- 2026-07-30：新增现有解题集独立评测模式、JSONL 路径映射、目标模型单独评分和非法裁判响应补跑。
- 2026-06-17：补齐 Harness Agent SSOT、场景路由、意图 Skill、references/playbooks 结构和 work-dir 执行原则。
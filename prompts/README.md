# 提示词目录

流水线全部提示词外置在这里。**改文件即生效，不需要动 Python 代码。**
逻辑名到文件名的映射在 `config.yaml` 的 `prompts.files`，想换文件名改那里即可。

占位符用 Python `str.format()` 语法。**因此正文里的字面花括号必须写成 `{{` 和 `}}`**，
否则会抛 `KeyError`。改完建议跑一次 `python3 config.py` 确认能正常加载。

## 文件清单

### stage2_clean_classify.md — 图片审核与分类

- 用在：Stage 2，对每张切好的题图判定是否为「有效完整单题」，并打上学段/学科/题型标签
- 占位符：无
- 输出契约：固定五行 `有效性判定/学段/学科/题型/无效原因`，Stage 2 按行解析后写入 Excel B–E 列。**改输出格式必须同步改 `stage2_clean_classify.py` 的解析逻辑**
- 想改什么：可选学科（默认只收数学/语文/英语）、题型清单、无效判定标准

### stage3_solve.md — 参考模型解题

- 用在：Stage 3，五个参考模型共用同一份提示词，保证横向可比
- 占位符：`{学段}` `{学科}` `{题型}`
- 输出契约：解题过程与最终答案，Stage 4 会读取
- 想改什么：解题风格、步骤粒度、干扰项（手写痕迹/已有答案）过滤规则、无法解题时的兜底话术

### stage4_consensus.md — 共识伪标签

- 用在：Stage 4，对比五个模型的解题结果，产出高置信伪标准答案
- 占位符：`{学段}` `{学科}` `{题型}`，以及 `{M1推理过程}` `{M1最终答案}` … `{M5推理过程}` `{M5最终答案}`
- 输出契约：JSON，含 `confidence`（high/medium/low）、`selected_ref_id`（M1–M5）、`consensus_answer`、`pseudo_label`、`filter_reason`，写入 Excel Z–AD 列
- 想改什么：置信度判定的严格程度。**这是整个评测的地基**——放松了会让低质题混进评测集，收紧了会让 high 样本量掉下来

### stage6_target_solve.md — 待测模型解题

- 用在：Stage 6，**且仅当 `config.yaml` 的 `providers.target_model.api_style` 设为 `openai_compat` 时**。默认的 `baidu_solve`（百度「拍照解题」）是产品化接口，只收图片、不接受提示词，因此不读这个文件
- 占位符：无。刻意只给图片 + 输出格式要求，不喂学段/学科/题型，与「拍照解题」只收图片的口径保持一致，避免待测模型比参考模型多拿信息
- 输出契约：必须按顺序原样输出 `==**解题过程**==` 和 `==**最终答案**==` 两个标题，否则 Stage 7 最高只能给 2 分
- 想改什么：解题风格、详略程度。**两个标题不要改**，它们同时写在 `config.yaml` 的 `evaluation.required_output_headings` 和 `stage7_score_target.md` 里

### stage7_score_target.md — 待测模型评分

- 用在：Stage 7，给待测模型的输出打 0–3 分
- 占位符：`{image}` `{grade_band}` `{subject}` `{question_type}` `{pseudo_label_text}` `{model_full_output}`
- 输出契约：JSON，含 `score_0_3`（整数 0–3）和 `judge_reason`，写入 Excel AI/AJ 列
- 想改什么：评分档位定义、格式合规门槛。**3 分要求的结构化标题同时写在 `config.yaml` 的 `evaluation.required_output_headings`，两处必须一致**

### stage7_score_reference.md — 参考模型横向评分

- 用在：Stage 7，用同一套标准给五个参考模型打分，得到横向对比基线
- 占位符：`{image}` `{grade_band}` `{subject}` `{question_type}` `{pseudo_label_text}` `{model_full_text}`
- 输出契约：同上，写入 Excel AK–AT 列
- 与 target 版的区别：参考模型输出没有固定结构化格式要求，因此不做格式合规门槛判定。**如果你想让横向对比完全同口径，就把两个文件改成一致**

## 改提示词的注意事项

1. **输出契约是硬约束**。上面每个「输出契约」都有对应的 Python 解析代码，只改提示词不改解析器会导致整批数据解析失败。
2. **改了评分提示词，历史数据不可比**。评分口径变化后，旧的 Excel 分数和新的不能混在一张横向表里。
3. **先小样本试跑**。用 `--limit 5` 之类的方式验证解析正常，再跑全量。
4. 保留一份原始版本，方便回滚和对比。

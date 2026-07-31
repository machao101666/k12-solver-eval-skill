# 运行完整评测

## 适用场景

用户提供 CSV，希望从原始图片 URL 跑到最终报告。

## 命令

```bash
python3 config.py                      # 先做配置自检
python3 run_pipeline.py --csv <input.csv>
```

## 执行阶段

1. Stage 1 OCR 切题。
2. Stage 2 清洗分类并写 B-E。
3. Stage 3 五模型解题并写 F-Y。
4. Stage 4 共识裁判并写 Z-AD。
5. Stage 5 上传 BOS 并写 AE（`providers.target_model.input_mode: image` 时可跳过）。
6. Stage 6 待测模型解题并写 AF/AG（接口风格由 `providers.target_model.api_style` 决定）。
7. Stage 7 待测模型和五模型评分并写 AI-AT。
8. Stage 8 生成 `report.md`。

## 检查点

- `run.json` 的 `status` 应为 `success`。
- `evaluation.xlsx` 应存在。
- `report.md` 应存在。
- `logs/` 应存在并包含阶段日志。
- `runs/INDEX.md` 应包含本次 run。
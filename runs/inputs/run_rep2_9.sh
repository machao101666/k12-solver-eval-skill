#!/bin/bash
# 顺序评测 rep2-rep9：每个 rep 独立输出目录，失败后自动补跑一次并重新出报告
cd "$(dirname "$0")/../.." || exit 1
XLSX="/Users/machao/Desktop/拍照解题/【评测&自动化】拍照解题/0730vis解题评测/evaluation.xlsx"
IMAGES="/Users/machao/Desktop/拍照解题/【评测&自动化】拍照解题/0730vis解题评测/evaluation"

for n in 2 3 4 5 6 7 8 9; do
  label="amuL20（FP16tp2x2 rep${n}）"
  root="runs/existing_eval_oneapi_opus5_rep${n}_20260907"
  echo "=========== rep${n} 开始 $(date '+%H:%M:%S') ==========="
  python3 evaluate_existing_results.py \
    --xlsx "$XLSX" --images-dir "$IMAGES" \
    --results-jsonl "runs/inputs/result_rep${n}.jsonl" \
    --output-root "$root" --concurrency 8 --invalid-retries 3 \
    --report-title "${label}高置信度161题评测报告（裁判 OneAPI Opus 5）"
  status=$?
  run_dir="${root}/result_rep${n}_high_161"
  if [ $status -ne 0 ]; then
    echo "--- rep${n} 首轮未通过校验，补跑无效行 ---"
    python3 evaluate_existing_results.py --resume-run "$run_dir" --invalid-retries 3 \
      --report-title "${label}高置信度161题评测报告（裁判 OneAPI Opus 5）"
    status=$?
  fi
  if [ $status -eq 0 ]; then
    python3 stage8_generate_report.py evaluation.xlsx --work-dir "$run_dir" \
      --output report.md --title "${label}高置信度161题评测报告（裁判 OneAPI Opus 5）" \
      --target-label "$label" >/dev/null
    echo "=========== rep${n} 完成 $(date '+%H:%M:%S') ==========="
  else
    echo "=========== rep${n} 失败，需人工检查 ==========="
  fi
done
echo "全部 rep2-rep9 处理结束 $(date '+%H:%M:%S')"

#!/bin/bash
# 顺序评测三个版本各 rep1-3（共 9 个），裁判 OneAPI Opus 5；失败自动补跑并重出报告
cd "$(dirname "$0")/../.." || exit 1
XLSX="/Users/machao/Desktop/拍照解题/【评测&自动化】拍照解题/0730vis解题评测/evaluation.xlsx"
IMAGES="/Users/machao/Desktop/拍照解题/【评测&自动化】拍照解题/0730vis解题评测/evaluation"
DATE=20260910

run_one() {
  local tag="$1" n="$2" label="$3"
  local root="runs/existing_eval_oneapi_opus5_${tag}_rep${n}_${DATE}"
  local run_dir="${root}/result_${tag}_rep${n}_high_161"
  echo "=========== ${tag} rep${n} 开始 $(date '+%H:%M:%S') ==========="
  python3 evaluate_existing_results.py \
    --xlsx "$XLSX" --images-dir "$IMAGES" \
    --results-jsonl "runs/inputs/result_${tag}_rep${n}.jsonl" \
    --output-root "$root" --concurrency 8 --invalid-retries 3 \
    --report-title "${label}高置信度161题评测报告（裁判 OneAPI Opus 5）"
  local status=$?
  if [ $status -ne 0 ]; then
    echo "--- ${tag} rep${n} 首轮未通过，补跑无效行 ---"
    python3 evaluate_existing_results.py --resume-run "$run_dir" --invalid-retries 3 \
      --report-title "${label}高置信度161题评测报告（裁判 OneAPI Opus 5）"
    status=$?
  fi
  if [ $status -eq 0 ]; then
    python3 stage8_generate_report.py evaluation.xlsx --work-dir "$run_dir" \
      --output report.md --title "${label}高置信度161题评测报告（裁判 OneAPI Opus 5）" \
      --target-label "$label" >/dev/null
    echo "=========== ${tag} rep${n} 完成 $(date '+%H:%M:%S') ==========="
  else
    echo "=========== ${tag} rep${n} 失败，需人工检查 ==========="
  fi
}

for n in 1 2 3; do run_one v0.1.1.0 "$n" "amu v0.1.1.0（L20 FP16tp2 rep${n}）"; done
for n in 1 2 3; do run_one v0.1.2.0 "$n" "amu v0.1.2.0（L20 FP16tp2 rep${n}）"; done
for n in 1 2 3; do run_one H100 "$n" "amu v0.0.9.4（H100 FP16tp2x2 rep${n}）"; done
echo "全部 9 个 rep 处理结束 $(date '+%H:%M:%S')"

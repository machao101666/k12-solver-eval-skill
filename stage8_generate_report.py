#!/usr/bin/env python3
"""
根据 evaluation/报告 Excel 自动生成评测 Markdown 报告。

用法：
  python3 stage8_generate_report.py evaluation.xlsx
  python3 stage8_generate_report.py 报告.xlsx --output reports/report.md
"""

from __future__ import annotations

import argparse
import math
import os
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

try:
    import pandas as pd
except ImportError as exc:
    raise SystemExit("缺少依赖 pandas，请先执行：pip3 install pandas openpyxl") from exc

from config import cfg


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_DIR / cfg("paths.excel_file", "evaluation.xlsx")
DEFAULT_REPORT_DIR = BASE_DIR / "reports"

GRADE_ORDER = ["小学", "初中", "高中"]
SUBJECT_ORDER = ["语文", "数学", "英语"]
SCORE_LABELS = [0, 1, 2, 3]
PASS_SCORE = int(cfg("evaluation.pass_score", 2))
ONLY_CONFIDENCE = str(cfg("evaluation.only_confidence", "high")).strip().lower()

# 待测模型评分列的候选表头。stage7 写入的表头随版本变化，这里全部兼容。
TARGET_MODEL_LABEL = "M0(待测模型)"
TARGET_SCORE_COLS = ["target_score_0_3", "score_0_3"]


def _build_model_specs() -> list[dict[str, Any]]:
    """参考模型在报告中的列映射。

    pandas 读 Excel 时重复表头会自动加 .1/.2 后缀，因此这里用「候选表头列表」
    做兼容匹配：既支持 config.yaml 里配置的标签，也保留历史表头别名。
    """
    legacy = [
        {"score": ["ernie_score_0_3", "ernire_score_0_3"], "ref_keys": ["ernie", "ernire"]},
        {"score": ["doubao_score_0_3"], "ref_keys": ["doubao", "seed"]},
        {"score": ["qwen_score_0_3"], "ref_keys": ["qwen"]},
        {"score": ["claude_score_0_3"], "ref_keys": ["claude"]},
        {"score": ["gemini_score_0_3"], "ref_keys": ["gemini"]},
    ]
    ref_models = cfg("models.reference_models", [])
    score_columns = cfg("excel.reference_score_columns", [])
    specs: list[dict[str, Any]] = []
    for index in range(max(len(ref_models), len(legacy))):
        legacy_item = legacy[index] if index < len(legacy) else {"score": [], "ref_keys": []}
        name = ref_models[index]["label"] if index < len(ref_models) else f"M{index + 1}"
        key = None
        if index < len(score_columns):
            key = score_columns[index].get("label")
        score_cols = ([f"{key}_score_0_3"] if key else []) + legacy_item["score"]
        reason_cols = ([f"{key}_judge_reason"] if key else []) + [
            f"{legacy_item['ref_keys'][0]}_judge_reason" if legacy_item["ref_keys"] else "",
            f"judge_reason.{index + 1}",
        ]
        specs.append({
            "name": name,
            "time_col": f"time{index + 1}",
            "answer_col": "answer" if index == 0 else f"answer.{index}",
            "score_cols": [col for col in dict.fromkeys(score_cols) if col],
            "reason_cols": [col for col in dict.fromkeys(reason_cols) if col],
            "ref_keys": list(dict.fromkeys(([key.lower()] if key else []) + legacy_item["ref_keys"])),
        })
    return specs


MODEL_SPECS = _build_model_specs()


def value_exists(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, float) and math.isnan(value):
        return False
    text = str(value).strip()
    return text != "" and text.lower() not in {"nan", "none", "null"}


def is_success_answer(value: Any) -> bool:
    if not value_exists(value):
        return False
    text = str(value).strip().lower()
    failed_markers = ["无法作答", "无法正常解题", "error", "timeout", "exception", "failed", "失败"]
    return not any(marker in text for marker in failed_markers)


def normalize_score(value: Any) -> int | None:
    if not value_exists(value):
        return None
    try:
        score = int(float(str(value).strip()))
    except ValueError:
        return None
    return score if score in SCORE_LABELS else None


def pct(numerator: int | float, denominator: int | float, digits: int = 1) -> str:
    if not denominator:
        return "—"
    return f"{numerator / denominator * 100:.{digits}f}%"


def fmt_seconds(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    try:
        return f"{float(value):.1f}s"
    except (TypeError, ValueError):
        return "—"


def md_table(headers: list[str], rows: Iterable[Iterable[Any]]) -> str:
    lines = ["|" + "|".join(headers) + "|", "|" + "|".join(["-"] * len(headers)) + "|"]
    for row in rows:
        cells = []
        for cell in row:
            text = "—" if cell is None or cell == "" else str(cell)
            text = text.replace("|", "\\|").replace("\n", "<br/>")
            trailing_backslashes = len(text) - len(text.rstrip("\\"))
            if trailing_backslashes % 2 == 1:
                text += "&#8203;"
            cells.append(text)
        lines.append("|" + "|".join(cells) + "|")
    return "\n".join(lines)


def sorted_group_keys(df: pd.DataFrame, cols: list[str]) -> list[tuple[Any, ...]]:
    keys = list(df.groupby(cols, dropna=False).size().index)

    def rank(key: tuple[Any, ...]) -> tuple[int, int, str]:
        grade = str(key[0]) if len(key) > 0 else ""
        subject = str(key[1]) if len(key) > 1 else ""
        return (
            GRADE_ORDER.index(grade) if grade in GRADE_ORDER else 99,
            SUBJECT_ORDER.index(subject) if subject in SUBJECT_ORDER else 99,
            " - ".join(str(x) for x in key),
        )

    return sorted(keys, key=rank)


def first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def score_series(df: pd.DataFrame, score_col: str) -> pd.Series:
    if score_col not in df.columns:
        return pd.Series([None] * len(df), index=df.index)
    return df[score_col].apply(normalize_score)


def valid_scored_df(df: pd.DataFrame, score_col: str) -> pd.DataFrame:
    if "confidence" in df.columns:
        base = df[df["confidence"].astype(str).str.lower().str.strip() == "high"].copy()
    else:
        base = df.copy()
    scores = score_series(base, score_col)
    return base[scores.notna()].assign(_score=scores[scores.notna()].astype(int))


def summarize_score_distribution(scored: pd.DataFrame) -> dict[int, int]:
    counter = Counter(scored["_score"].tolist()) if len(scored) else Counter()
    return {score: counter.get(score, 0) for score in SCORE_LABELS}


def grade_subject_score_table(scored: pd.DataFrame) -> str:
    rows = []
    for grade, subject in sorted_group_keys(scored, ["学段", "学科"]):
        part = scored[(scored["学段"] == grade) & (scored["学科"] == subject)]
        total = len(part)
        counts = summarize_score_distribution(part)
        pass_count = counts[2] + counts[3]
        rows.append([
            f"{grade} - {subject}", total, counts[0], counts[1], counts[2], counts[3], f"**{pct(pass_count, total)}**",
        ])
    total = len(scored)
    counts = summarize_score_distribution(scored)
    rows.append(["**合计**", f"**{total}**", f"**{counts[0]}**", f"**{counts[1]}**", f"**{counts[2]}**", f"**{counts[3]}**", f"**{pct(counts[2] + counts[3], total)}**"])
    return md_table(["学段 - 学科", "总数", "0分", "1分", "2分", "3分", "2+分率"], rows)


def subject_grade_sample_table(df: pd.DataFrame) -> str:
    rows = []
    for subject in SUBJECT_ORDER:
        part = df[df["学科"] == subject]
        if part.empty:
            continue
        total = len(part)
        counts_by_grade = {grade: len(part[part["学段"] == grade]) for grade in GRADE_ORDER}
        type_counts = part["题型"].fillna("未标注").astype(str).value_counts()
        type_desc = "<br/>".join([f"* {name}【{pct(count, total)}】" for name, count in type_counts.items()])
        type_count_desc = "<br/>".join([f"* 共{count}" for _, count in type_counts.items()])
        rows.append([subject, total, counts_by_grade["小学"], counts_by_grade["初中"], counts_by_grade["高中"], type_desc, type_count_desc])
    total_row = ["**合计**", f"**{len(df)}**"] + [f"**{len(df[df['学段'] == grade])}**" for grade in GRADE_ORDER] + ["", ""]
    rows.append(total_row)
    return md_table(["学科", "合计", "小学", "初中", "高中", "题型分布 & 各题型占比", "题型对应数量"], rows)


def model_success_table(df: pd.DataFrame) -> str:
    rows = []
    high_mask = df["confidence"].astype(str).str.lower().str.strip() == ONLY_CONFIDENCE if "confidence" in df.columns else pd.Series([True] * len(df), index=df.index)
    for spec in MODEL_SPECS:
        answer_col = spec["answer_col"]
        time_col = spec["time_col"]
        if answer_col not in df.columns:
            continue
        success_mask = df[answer_col].apply(is_success_answer)
        total = len(df)
        success_count = int(success_mask.sum())
        times_all = pd.to_numeric(df[time_col], errors="coerce") if time_col in df.columns else pd.Series(dtype=float)
        times_success = pd.to_numeric(df.loc[success_mask, time_col], errors="coerce") if time_col in df.columns else pd.Series(dtype=float)
        times_high = pd.to_numeric(df.loc[high_mask & success_mask, time_col], errors="coerce") if time_col in df.columns else pd.Series(dtype=float)
        rows.append([
            spec["name"],
            f"{pct(success_count, total)}（{success_count}/{total}）",
            fmt_seconds(times_all.quantile(0.5)), fmt_seconds(times_all.quantile(0.9)), fmt_seconds(times_all.quantile(0.99)),
            fmt_seconds(times_success.quantile(0.5)), fmt_seconds(times_success.quantile(0.9)), fmt_seconds(times_success.quantile(0.99)),
            fmt_seconds(times_high.mean()), fmt_seconds(times_high.max()), fmt_seconds(times_high.min()),
        ])
    return md_table([
        "模型", "解题成功率", "P50(全部)", "P90(全部)", "P99(全部)",
        "P50(成功)", "P90(成功)", "P99(成功)", "平均耗时(high)", "最大耗时(high)", "最小耗时(high)",
    ], rows)


def model_success_by_group_table(df: pd.DataFrame) -> str:
    rows = []
    for grade, subject in sorted_group_keys(df, ["学段", "学科"]):
        part = df[(df["学段"] == grade) & (df["学科"] == subject)]
        row = [f"{grade} - {subject}"]
        for spec in MODEL_SPECS:
            answer_col = spec["answer_col"]
            if answer_col not in df.columns:
                row.append("—")
            else:
                row.append(pct(int(part[answer_col].apply(is_success_answer).sum()), len(part), 0))
        rows.append(row)
    return md_table(["学段 - 学科", "ernie", "doubao", "qwen", "claude", "gemini"], rows)


def confidence_table(df: pd.DataFrame) -> str:
    if "confidence" not in df.columns:
        return "未找到 confidence 列。"
    counts = df["confidence"].fillna("空").astype(str).str.strip().replace("", "空").value_counts()
    rows = [[name, count, pct(count, len(df))] for name, count in counts.items()]
    return md_table(["置信度", "题数", "占比"], rows)


def qtype_score_tables(scored: pd.DataFrame) -> str:
    sections = []
    for subject in SUBJECT_ORDER:
        subject_df = scored[scored["学科"] == subject]
        if subject_df.empty:
            continue
        rows = []
        qtypes = list(subject_df["题型"].fillna("未标注").astype(str).value_counts().index)
        for qtype in qtypes:
            row = [qtype]
            qtype_df = subject_df[subject_df["题型"].fillna("未标注").astype(str) == qtype]
            total_pass = 0
            total_count = 0
            for grade in GRADE_ORDER:
                part = qtype_df[qtype_df["学段"] == grade]
                count = len(part)
                pass_count = int((part["_score"] >= 2).sum())
                total_pass += pass_count
                total_count += count
                row.append("—" if count == 0 else f"{pass_count}/{count} {pct(pass_count, count)}")
            row.append("—" if total_count == 0 else f"{total_pass}/{total_count} {pct(total_pass, total_count)}")
            rows.append(row)
        sections.append(f"{subject}\n\n" + md_table(["题型", "小学", "初中", "高中", "合计"], rows))
    return "\n\n".join(sections) if sections else "无可统计题型数据。"


def tested_model_time_table(scored: pd.DataFrame) -> str:
    time_col = first_existing_col(scored, ["time", "target_time"])
    if not time_col:
        return "未找到待测模型 time 列。"
    times = pd.to_numeric(scored[time_col], errors="coerce").dropna()
    rows = [
        ["有效数据", f"{len(times)}行"],
        ["P50", fmt_seconds(times.quantile(0.5))],
        ["P90", fmt_seconds(times.quantile(0.9))],
        ["P99", fmt_seconds(times.quantile(0.99))],
        ["平均值", fmt_seconds(times.mean())],
        ["最小值", fmt_seconds(times.min())],
        ["最大值", fmt_seconds(times.max())],
    ]
    return md_table(["指标", "数值"], rows)


def score_overview_table(scored: pd.DataFrame) -> str:
    total = len(scored)
    counts = summarize_score_distribution(scored)
    rows = [
        ["有效评测题数（confidence=high 且评分有效）", f"{total} 题"],
        ["0分率", f"{pct(counts[0], total)}（{counts[0]}/{total}）"],
        ["1分率", f"{pct(counts[1], total)}（{counts[1]}/{total}）"],
        ["2分率（主达标线，>=2）", f"**{pct(counts[2] + counts[3], total)}**（{counts[2] + counts[3]}/{total}）"],
        ["3分率（优质率）", f"{pct(counts[3], total)}（{counts[3]}/{total}）"],
    ]
    return md_table(["指标", "值"], rows)


def score_detail_table(scored: pd.DataFrame) -> str:
    total = len(scored)
    counts = summarize_score_distribution(scored)
    rows = [[f"{score}分", counts[score], pct(counts[score], total)] for score in SCORE_LABELS]
    rows.append(["**合计**", f"**{total}**", "**100%**" if total else "—"])
    return md_table(["分数", "题数", "占比"], rows)


def cross_model_score_table(scored: pd.DataFrame) -> str:
    """按 Excel 实际评分数据计算待测模型与参考模型的横向对比表。

    待测模型分数取 AI 列，参考模型分数取 AK/AM/AO/AQ/AS 列。
    某个模型的评分列不存在或全为空时，该行不出现在表中。
    """
    specs = [{"name": TARGET_MODEL_LABEL, "score_cols": TARGET_SCORE_COLS}] + MODEL_SPECS
    rows = []
    for spec in specs:
        score_col = first_existing_col(scored, spec["score_cols"])
        if not score_col:
            continue
        scores = score_series(scored, score_col).dropna().astype(int)
        total = len(scores)
        if total == 0:
            continue
        counts = {score: int((scores == score).sum()) for score in SCORE_LABELS}
        pass_count = counts[2] + counts[3]
        rows.append([
            spec["name"], total,
            counts[0], counts[1], counts[2], counts[3],
            pct(pass_count, total), pct(counts[3], total),
        ])
    if not rows:
        return "当前 Excel 尚未写入任何有效评分列，无法生成横向评分表。"
    return md_table(["模型", "总有效", "0分", "1分", "2分", "3分", "≥2分率", "3分率"], rows)


def cross_model_low_score_table(scored: pd.DataFrame, max_cases: int = 20) -> str:
    rows = []
    specs = [{"name": TARGET_MODEL_LABEL, "score_cols": TARGET_SCORE_COLS, "reason_cols": ["judge_reason"]}] + MODEL_SPECS
    path_col = "题目路径" if "题目路径" in scored.columns else "图片路径" if "图片路径" in scored.columns else None
    for spec in specs:
        score_col = first_existing_col(scored, spec["score_cols"])
        if not score_col:
            continue
        reason_col = first_existing_col(scored, spec.get("reason_cols", []))
        scores = score_series(scored, score_col)
        low = scored[scores.isin([0, 1])].copy()
        for idx, row in low.head(max_cases).iterrows():
            rows.append([
                spec["name"], idx + 2, row.get("学段", ""), row.get("学科", ""), row.get("题型", ""),
                int(normalize_score(row.get(score_col))), row.get(path_col, "") if path_col else "",
                row.get(reason_col, "") if reason_col else "",
            ])
    if not rows:
        return "横向评分中未发现 0/1 分低分题，或当前 Excel 尚未写入五模型评分列。"
    return md_table(["模型", "Excel行", "学段", "学科", "题型", "分数", "题目路径", "评分原因"], rows)


def low_score_cases(scored: pd.DataFrame, max_cases: int = 12) -> str:
    low = scored[scored["_score"] <= 1].copy()
    if low.empty:
        return "未发现 0/1 分低分题。"
    path_col = "题目路径" if "题目路径" in low.columns else "图片路径" if "图片路径" in low.columns else None
    reason_col = "judge_reason" if "judge_reason" in low.columns else None
    output_col = "model_output" if "model_output" in low.columns else None
    rows = []
    for idx, row in low.head(max_cases).iterrows():
        rows.append([
            idx + 2,
            row.get("学段", ""), row.get("学科", ""), row.get("题型", ""), int(row["_score"]),
            row.get(path_col, "") if path_col else "",
            row.get(reason_col, "") if reason_col else "",
            row.get(output_col, "") if output_col else "",
        ])
    return md_table(["Excel行", "学段", "学科", "题型", "分数", "题目路径", "评分原因", "原始待测模型解题结果"], rows)


def generate_conclusions(df: pd.DataFrame, scored: pd.DataFrame) -> list[str]:
    total = len(df)
    high = len(df[df["confidence"].astype(str).str.lower().str.strip() == "high"]) if "confidence" in df.columns else len(scored)
    valid = len(scored)
    counts = summarize_score_distribution(scored)
    pass_rate = (counts[2] + counts[3]) / valid if valid else 0
    zero_one = counts[0] + counts[1]
    conclusions = [
        f"本次共采用 **{total} 道高置信题**。",
        f"高置信且评分有效样本 **{valid} 道**，待测模型 **2+分率为 {pct(counts[2] + counts[3], valid)}**，3分率为 **{pct(counts[3], valid)}**。",
        f"低分题共 **{zero_one} 道**，其中 0分 {counts[0]} 道、1分 {counts[1]} 道，可优先作为后续问题复盘样本。",
    ]
    if not scored.empty:
        group_stats = []
        for grade, subject in sorted_group_keys(scored, ["学段", "学科"]):
            part = scored[(scored["学段"] == grade) & (scored["学科"] == subject)]
            rate = float((part["_score"] >= 2).mean()) if len(part) else 0
            group_stats.append((rate, len(part), f"{grade} - {subject}"))
        if group_stats:
            best = sorted(group_stats, key=lambda x: (-x[0], -x[1], x[2]))[0]
            weak = sorted(group_stats, key=lambda x: (x[0], -x[1], x[2]))[0]
            conclusions.append(f"分学段学科看，表现最好的是 **{best[2]}**（2+分率 {pct(best[0], 1)}，样本 {best[1]} 道）；当前短板是 **{weak[2]}**（2+分率 {pct(weak[0], 1)}，样本 {weak[1]} 道）。")
    if pass_rate >= 0.95:
        conclusions.append("整体达到较高可用水平，后续优化重点建议放在低分题归因、输出格式稳定性和难题解析完整性上。")
    elif pass_rate >= 0.85:
        conclusions.append("整体基本可用，但仍需针对低分高发学科/题型做专项优化后再扩大使用范围。")
    else:
        conclusions.append("整体可用率偏低，建议先收敛样本质量和解题输出稳定性，再进入更大规模评测。")
    return conclusions


def build_report(df: pd.DataFrame, source_path: Path, title: str, cross_model_df: pd.DataFrame | None = None) -> str:
    score_col = first_existing_col(df, TARGET_SCORE_COLS) or "score_0_3"
    scored = valid_scored_df(df, score_col)
    cross_source = cross_model_df if cross_model_df is not None else df
    cross_score_col = first_existing_col(cross_source, TARGET_SCORE_COLS) or "score_0_3"
    cross_scored = valid_scored_df(cross_source, cross_score_col)
    conclusions = generate_conclusions(df, scored)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    total = len(df)
    high = len(df[df["confidence"].astype(str).str.lower().str.strip() == "high"]) if "confidence" in df.columns else len(scored)
    counts = summarize_score_distribution(scored)
    pass_count = counts[2] + counts[3]

    parts = [
        title,
        "",
        f"> 生成时间：{now}",
        f"> 数据来源：{source_path.name}",
        "",
        "# 核心指标汇总与结论",
        f"本次共采用 **{total} 道高置信题**，开展 K12 作业解题模型自动化评测。",
        "",
        "* 核心指标",
        f"    * **2分率（主指标）**：{pct(pass_count, len(scored))}（{pass_count}/{len(scored)}）",
        f"    * **0分率（风险指标）**：{pct(counts[0], len(scored))}（{counts[0]}/{len(scored)}）",
        f"    * **1分率（风险指标）**：{pct(counts[1], len(scored))}（{counts[1]}/{len(scored)}）",
        f"    * **3分率（优质率）**：{pct(counts[3], len(scored))}（{counts[3]}/{len(scored)}）",
        "",
        "* 自动结论",
    ]
    parts.extend([f"    * {line}" for line in conclusions])
    parts.extend([
        "",
        grade_subject_score_table(scored) if not scored.empty else "暂无有效评分数据。",
        "",
        "# 一、高置信题集概览",
        "## 1.1 样本分布",
        f"本次共采用 **{total} 道高置信题**，按学科、学段和题型分布如下。", 
        "",
        subject_grade_sample_table(df),
        "",
        "# 二、待测模型评分结果",
        "## 2.1 整体评分指标",
        score_overview_table(scored),
        "",
        "分数段明细：",
        "",
        score_detail_table(scored),
        "",
        "## 2.2 分学段 × 学科评分表现",
        grade_subject_score_table(scored) if not scored.empty else "暂无有效评分数据。",
        "",
        "## 2.3 分题型 2+ 分率",
        qtype_score_tables(scored),
        "",
        "## 2.4 待测模型耗时",
        tested_model_time_table(scored),
        "",
        "## 2.5 低分题清单",
        low_score_cases(scored),
        "",
        "# 三、解题模型横向评分",
        cross_model_score_table(cross_scored),
        "",

    ])
    return "\n".join(parts)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="根据评测 Excel 自动生成 Markdown 报告")
    parser.add_argument("xlsx", nargs="?", default=None, help="输入 Excel 路径，默认 <work-dir>/evaluation.xlsx")
    parser.add_argument("--output", "-o", help="输出 Markdown 路径，默认 reports/report_<timestamp>.md")
    parser.add_argument("--title", default="解题模型评测报告", help="报告标题")
    parser.add_argument("--cross-model-xlsx", dest="cross_model_xlsx", help="横向评分来源 Excel；未指定时使用主输入 Excel")
    parser.add_argument("--work-dir", default=os.environ.get("AUTO_EVAL_WORK_DIR"), help="运行产物目录；未指定 xlsx/output 时用于默认路径")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    work_dir = Path(args.work_dir).expanduser() if args.work_dir else BASE_DIR
    if not work_dir.is_absolute():
        work_dir = BASE_DIR / work_dir
    work_dir = work_dir.resolve()

    xlsx_path = Path(args.xlsx).expanduser() if args.xlsx else work_dir / "evaluation.xlsx"
    if not xlsx_path.is_absolute():
        xlsx_path = work_dir / xlsx_path
    if not xlsx_path.exists():
        raise SystemExit(f"输入 Excel 不存在：{xlsx_path}")

    df = pd.read_excel(xlsx_path)
    required = ["学段", "学科", "题型"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise SystemExit(f"Excel 缺少必要列：{', '.join(missing)}")

    cross_model_df = None
    if args.cross_model_xlsx:
        cross_path = Path(args.cross_model_xlsx).expanduser()
        if not cross_path.is_absolute():
            cross_path = work_dir / cross_path
        if not cross_path.exists():
            raise SystemExit(f"横向评分 Excel 不存在：{cross_path}")
        cross_model_df = pd.read_excel(cross_path)

    report = build_report(df, xlsx_path, args.title, cross_model_df=cross_model_df)

    if args.output:
        output_path = Path(args.output).expanduser()
        if not output_path.is_absolute():
            output_path = work_dir / output_path
    else:
        report_dir = work_dir / "reports"
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = report_dir / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(report, encoding="utf-8")
    print(f"报告已生成：{output_path}")


if __name__ == "__main__":
    main()

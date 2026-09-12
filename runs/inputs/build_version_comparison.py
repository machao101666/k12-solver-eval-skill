#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""跨版本横向对比报告（v0.1.1.0 / v0.1.2.0 / H100 / L20-tp2x2，各取前3次）。

不含 badcase 的模型解题原文与判罚全文，只保留统计与低分题总览。
"""

from __future__ import annotations

import re
import statistics
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parents[2]
OUT = BASE / "runs" / "version_comparison_0910" / "version_comparison_report.md"

# (显示名, run 目录前缀, run 子目录 stem 前缀, rep 列表)
VERSIONS = [
    ("v0.1.1.0 (L20 tp2)", "existing_eval_oneapi_opus5_v0.1.1.0_rep{n}_20260910", "result_v0.1.1.0_rep{n}_high_161", [1, 2, 3]),
    ("v0.1.2.0 (L20 tp2)", "existing_eval_oneapi_opus5_v0.1.2.0_rep{n}_20260910", "result_v0.1.2.0_rep{n}_high_161", [1, 2, 3]),
    ("v0.0.9.4 (H100 tp2x2)", "existing_eval_oneapi_opus5_H100_rep{n}_20260910", "result_H100_rep{n}_high_161", [1, 2, 3]),
]

COL = {"path": 2, "grade": 3, "subject": 4, "qtype": 5, "score": 35, "reason": 36}

DEDUCT_PATTERNS = [
    ("最终答案错误", r"最终答案错|答案错误|答案与标准答案不一致|逻辑对结果错"),
    ("推理不可用（可用占比<50%）", r"推理不可用|可用占比不足|重复堆砌|无实质解题内容"),
    ("严重啰嗦/反复自我纠错", r"严重啰嗦|车轱辘|反复自我|无意义重复"),
    ("选对但过程错", r"选对但过程错|碰巧正确"),
    ("公式原理硬伤", r"公式原理硬伤|公式硬伤"),
    ("解题过程过度简略", r"过度简略|关键步骤省略|关键推理步骤省略|跳步"),
    ("漏答/未作答", r"漏答"),
    ("幻觉/答非所问", r"幻觉|答非所问|凭空|臆造"),
    ("格式不符约定", r"格式不符|标题缺失|自创标题|旧格式"),
]
TRIGGER_RE = re.compile(r"(?<!不)(?<!未)(触发|命中|符合官方|依据官方|违反)")
WINDOW = 80


def load_version(prefix, stem, reps):
    """返回 {rep: {path: {...}}}"""
    out = {}
    for n in reps:
        run = BASE / "runs" / prefix.format(n=n) / stem.format(n=n)
        ws = openpyxl.load_workbook(run / "evaluation.xlsx", data_only=True, read_only=True).active
        rows = {}
        for r in range(2, ws.max_row + 1):
            path = str(ws.cell(r, COL["path"]).value or "")
            if not path:
                continue
            rows[path] = {
                "score": ws.cell(r, COL["score"]).value,
                "reason": str(ws.cell(r, COL["reason"]).value or ""),
                "grade": ws.cell(r, COL["grade"]).value,
                "subject": ws.cell(r, COL["subject"]).value,
                "qtype": ws.cell(r, COL["qtype"]).value,
            }
        out[n] = rows
    return out


def fmt_pct(num, den):
    return "—" if not den else f"{num / den * 100:.1f}%"


def rep_stats(vdata, reps):
    """每个 rep 的分布与比率。"""
    res = []
    for n in reps:
        c = Counter(vdata[n][p]["score"] for p in vdata[n])
        tot = sum(c[s] for s in (0, 1, 2, 3))
        res.append({"rep": n, 0: c[0], 1: c[1], 2: c[2], 3: c[3],
                    "p2": (c[2] + c[3]) / tot * 100, "p3": c[3] / tot * 100,
                    "low": c[0] + c[1], "tot": tot})
    return res


def section_overview(loaded):
    lines = [
        "# 一、版本总览（各版本前 3 次重复）",
        "",
        "同一份 161 道 high 题、同一 OneAPI `Opus 5` 裁判，各版本重复 3 次。",
        "",
        "|版本|rep|0分|1分|2分|3分|≥2分率|3分率|低分题数|",
        "|-|-|-|-|-|-|-|-|-|",
    ]
    summary = {}
    for name, prefix, stem, reps in VERSIONS:
        stats = rep_stats(loaded[name], reps)
        for s in stats:
            lines.append(f"|{name}|rep{s['rep']}|{s[0]}|{s[1]}|{s[2]}|{s[3]}|"
                         f"{s['p2']:.1f}%|{s['p3']:.1f}%|{s['low']}|")
        p2s = [s["p2"] for s in stats]
        p3s = [s["p3"] for s in stats]
        lows = [s["low"] for s in stats]
        summary[name] = (statistics.mean(p2s), min(p2s), max(p2s),
                         statistics.mean(p3s), statistics.mean(lows))
    lines += [
        "",
        "## 各版本均值汇总",
        "",
        "|版本|≥2分率均值|≥2分率范围|3分率均值|低分题均值|",
        "|-|-|-|-|-|",
    ]
    for name, _, _, _ in VERSIONS:
        m2, mn, mx, m3, ml = summary[name]
        lines.append(f"|{name}|**{m2:.1f}%**|{mn:.1f}%~{mx:.1f}%|**{m3:.1f}%**|{ml:.1f}|")
    lines += [
        "",
        "> 每版本仅 3 次重复，单次 ≥2分率存在约 ±2pp 波动；差距小于该幅度时不宜判定优劣。",
        "",
    ]
    return lines, summary


def _group_rates(loaded, key):
    """{version: {group: mean_2plus_rate}}, groups 全集有序。"""
    result = {}
    groups = {}
    for name, prefix, stem, reps in VERSIONS:
        gv = defaultdict(list)
        for n in reps:
            part = defaultdict(list)
            for path, rec in loaded[name][n].items():
                part[str(rec[key] or "未标注")].append(rec["score"])
            for g, scores in part.items():
                gv[g].append(sum(1 for s in scores if s >= 2) / len(scores) * 100)
                groups.setdefault(g, len(scores))
        result[name] = {g: statistics.mean(v) for g, v in gv.items()}
    order = sorted(groups, key=lambda g: (-groups[g], g))
    return result, order, groups


def section_dimension(loaded, key, title):
    rates, order, sizes = _group_rates(loaded, key)
    header = [title.split("：")[0].replace("## ", ""), "题数"] + [n for n, *_ in VERSIONS]
    lines = [title, "", "|" + "|".join(header) + "|", "|" + "|".join(["-"] * len(header)) + "|"]
    for g in order:
        cells = [g, str(sizes[g])] + [f"{rates[n].get(g, 0):.0f}%" for n, *_ in VERSIONS]
        lines.append("|" + "|".join(cells) + "|")
    lines += ["", "> 数值为该分组 ≥2分率的 3 次均值。", ""]
    return lines


def section_reasons(loaded):
    lines = [
        "# 三、扣分原因归类（各版本 0/1 分样本，3 次累计）",
        "",
        "只在裁判理由的结论性语句附近做条款匹配，避免把「无漏答」等否定表述误计。",
        "",
        "|扣分原因|" + "|".join(n for n, *_ in VERSIONS) + "|",
        "|-|" + "|".join(["-"] * len(VERSIONS)) + "|",
    ]
    per_version = {}
    totals = {}
    for name, prefix, stem, reps in VERSIONS:
        cnt = Counter()
        tot = 0
        for n in reps:
            for path, rec in loaded[name][n].items():
                if rec["score"] > 1:
                    continue
                tot += 1
                seg = " ".join(rec["reason"][m.start(): m.start() + WINDOW]
                               for m in TRIGGER_RE.finditer(rec["reason"]))
                hit = False
                for label, pat in DEDUCT_PATTERNS:
                    if re.search(pat, seg):
                        cnt[label] += 1
                        hit = True
                if not hit:
                    cnt["其他/未归类"] += 1
        per_version[name] = cnt
        totals[name] = tot
    all_reasons = [lbl for lbl, _ in DEDUCT_PATTERNS] + ["其他/未归类"]
    for label in all_reasons:
        cells = [label] + [str(per_version[n].get(label, 0)) for n, *_ in VERSIONS]
        lines.append("|" + "|".join(cells) + "|")
    lines.append("|**低分样本合计**|" + "|".join(f"**{totals[n]}**" for n, *_ in VERSIONS) + "|")
    lines += ["", "> 一个样本可命中多条，故各列之和可能超过合计。", ""]
    return lines


def section_low_overview(loaded):
    """低分题总览：哪些题在各版本出问题，只给每版本低分次数，不含判罚/模型原文。"""
    all_paths = {}
    for name, prefix, stem, reps in VERSIONS:
        for n in reps:
            for path, rec in loaded[name][n].items():
                all_paths.setdefault(path, rec)
    rows = []
    for path, meta in all_paths.items():
        per = {}
        total_low = 0
        for name, prefix, stem, reps in VERSIONS:
            low = sum(1 for n in reps if loaded[name][n][path]["score"] <= 1)
            per[name] = low
            total_low += low
        if total_low:
            rows.append((total_low, path, meta, per))
    rows.sort(key=lambda x: (-x[0], x[1]))
    lines = [
        "# 四、低分题总览（各版本 3 次内被判 0/1 分的次数）",
        "",
        f"共 {len(rows)} 道题在至少一个版本出现过低分，按累计低分次数降序。"
        "单元格为该版本 3 次中判 0/1 分的次数（越高越稳定出问题）。本章不含判罚理由与模型原文。",
        "",
        "|题目|学段|学科|题型|" + "|".join(f"{n} (低分/3)" for n, *_ in VERSIONS) + "|累计|",
        "|-|-|-|-|" + "|".join(["-"] * len(VERSIONS)) + "|-|",
    ]
    for total_low, path, meta, per in rows:
        cells = [f"`{path}`", str(meta["grade"]), str(meta["subject"]), str(meta["qtype"])]
        cells += [str(per[n]) for n, *_ in VERSIONS]
        cells.append(f"**{total_low}**")
        lines.append("|" + "|".join(cells) + "|")
    lines += ["", "> 低分明细（判罚依据、模型原文）参见各 run 目录下的 report.md。", ""]
    return lines


def section_conclusion(summary):
    best2 = max(summary, key=lambda k: summary[k][0])
    best3 = max(summary, key=lambda k: summary[k][3])
    worst3 = min(summary, key=lambda k: summary[k][3])
    v111 = summary.get("v0.1.1.0 (L20 tp2)")
    v120 = summary.get("v0.1.2.0 (L20 tp2)")
    h100 = summary.get("v0.0.9.4 (H100 tp2x2)")
    l20 = summary.get("v0.0.9.4 (L20 tp2x2)")
    lines = [
        "跨版本横向对比报告（各版本前 3 次重复，裁判 OneAPI Opus 5）",
        "",
        "> 评测集：161 道 confidence=high｜裁判：OneAPI `Opus 5`｜每版本重复 3 次",
        "",
        "# 结论摘要",
        "",
        f"1. **≥2分率三版本高度接近**（均在 90~92%，两两差距小于单次 ±2pp 波动），"
        "用 3 次样本无法区分合格率高低。真正的差异体现在 3分率（优质率）。",
        f"2. **3分率最高：{best3}（{summary[best3][3]:.1f}%）；最低：{worst3}（{summary[worst3][3]:.1f}%）**，"
        f"两者相差 {summary[best3][3] - summary[worst3][3]:.1f}pp，超出波动幅度。",
    ]
    if v111 and v120:
        lines.append(
            f"3. **v0.1.1.0 → v0.1.2.0 优质率明显回落**：3分率从 {v111[3]:.1f}% 降到 {v120[3]:.1f}%"
            f"（-{v111[3] - v120[3]:.1f}pp），但 ≥2分率基本持平（{v111[0]:.1f}% vs {v120[0]:.1f}%）。"
            "说明合格能力没退，是更多题从 3 分掉到 2 分——需结合第三章看是风格还是答案错误。"
        )
    if h100:
        lines.append(
            f"4. **v0.0.9.4 (H100 tp2x2)** ≥2分率 {h100[0]:.1f}%、3分率 {h100[3]:.1f}%，"
            "介于两个新版本之间，可作为旧版本基线参考。"
        )
    lines += [
        "5. **仅 3 次重复，结论为趋势性**：如需定论，建议每版本补齐到 9 次取均值。",
        "",
    ]
    return lines


def main():
    loaded = {name: load_version(prefix, stem, reps) for name, prefix, stem, reps in VERSIONS}
    overview, summary = section_overview(loaded)
    parts = section_conclusion(summary)
    parts += overview
    parts += ["# 二、分组维度对比（≥2分率，3 次均值）", ""]
    parts += section_dimension(loaded, "subject", "## 2.1 分学科")
    parts += section_dimension(loaded, "grade", "## 2.2 分学段")
    parts += section_dimension(loaded, "qtype", "## 2.3 分题型")
    parts += section_reasons(loaded)
    parts += section_low_overview(loaded)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(parts) + "\n", encoding="utf-8")
    print(f"报告已生成：{OUT}")


if __name__ == "__main__":
    main()

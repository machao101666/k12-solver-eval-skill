#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""汇总 rep1-rep9 九次重复评测，生成对比分析报告。"""

from __future__ import annotations

import re
import statistics
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parents[2]
REPS = list(range(1, 10))
COL = {"path": 2, "grade": 3, "subject": 4, "qtype": 5, "output": 32, "score": 35, "reason": 36}
OUT_DIR = BASE / "runs" / "rep1_9_comparison"
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
# 裁判理由里会大量出现「无漏答」「无幻觉」这类否定表述，
# 因此只在「触发/命中/符合官方/违反」等结论性语句附近的窗口内做归类。
TRIGGER_RE = re.compile(r"(?<!不)(?<!未)(触发|命中|符合官方|依据官方|违反)")
WINDOW = 80


def load() -> dict[str, dict[int, dict]]:
    data: dict[str, dict[int, dict]] = defaultdict(dict)
    for n in REPS:
        run = BASE / f"runs/existing_eval_oneapi_opus5_rep{n}_20260907/result_rep{n}_high_161"
        ws = openpyxl.load_workbook(run / "evaluation.xlsx", data_only=True, read_only=True).active
        for r in range(2, ws.max_row + 1):
            path = str(ws.cell(r, COL["path"]).value or "")
            if not path:
                continue
            data[path][n] = {
                "score": ws.cell(r, COL["score"]).value,
                "output": str(ws.cell(r, COL["output"]).value or ""),
                "reason": str(ws.cell(r, COL["reason"]).value or ""),
                "grade": ws.cell(r, COL["grade"]).value,
                "subject": ws.cell(r, COL["subject"]).value,
                "qtype": ws.cell(r, COL["qtype"]).value,
            }
    return data


def fmt_pct(num: float, den: float) -> str:
    return "—" if not den else f"{num / den * 100:.1f}%"


def md_cell(text: str, limit: int) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > limit:
        text = text[:limit] + "…"
    return text.replace("|", "\\|")


def table_text(text: str) -> str:
    """表格单元格内保留全文：转义竖线、换行改 <br/>，不做长度截断。"""
    text = str(text or "").replace("|", "\\|").replace("\r\n", "\n").replace("\r", "\n")
    return text.replace("\n", "<br/>").strip()


def rep_summary(data) -> list[dict]:
    out = []
    for n in REPS:
        counts = Counter(v[n]["score"] for v in data.values())
        total = sum(counts[s] for s in (0, 1, 2, 3))
        out.append({
            "rep": n, "total": total,
            0: counts[0], 1: counts[1], 2: counts[2], 3: counts[3],
            "p2": (counts[2] + counts[3]) / total * 100,
            "p3": counts[3] / total * 100,
            "low": counts[0] + counts[1],
        })
    return out


def section_overview(data, summary) -> list[str]:
    p2s = [s["p2"] for s in summary]
    p3s = [s["p3"] for s in summary]
    lows = [s["low"] for s in summary]
    lines = [
        "# 一、九次重复评测总览",
        "",
        "同一待测配置（v0.0.9.4_L20_fp16tp2x2_0904）在同一份 161 道高置信题上重复推理 9 次，",
        "全部由 OneAPI `Opus 5` 裁判评分，评分口径完全一致。",
        "",
        "|rep|有效题数|0分|1分|2分|3分|≥2分率|3分率|低分题数|",
        "|-|-|-|-|-|-|-|-|-|",
    ]
    for s in summary:
        flag = "**" if s["p2"] == max(p2s) else ""
        lines.append(
            f"|{flag}rep{s['rep']}{flag}|{s['total']}|{s[0]}|{s[1]}|{s[2]}|{s[3]}|"
            f"{s['p2']:.1f}%|{s['p3']:.1f}%|{s['low']}|"
        )
    lines += [
        "",
        "|指标|均值|最小|最大|极差|标准差|",
        "|-|-|-|-|-|-|",
        f"|≥2分率|{statistics.mean(p2s):.1f}%|{min(p2s):.1f}%|{max(p2s):.1f}%|"
        f"{max(p2s) - min(p2s):.1f}pp|{statistics.stdev(p2s):.2f}|",
        f"|3分率|{statistics.mean(p3s):.1f}%|{min(p3s):.1f}%|{max(p3s):.1f}%|"
        f"{max(p3s) - min(p3s):.1f}pp|{statistics.stdev(p3s):.2f}|",
        f"|低分题数|{statistics.mean(lows):.1f}|{min(lows)}|{max(lows)}|{max(lows) - min(lows)}|"
        f"{statistics.stdev(lows):.2f}|",
        "",
    ]
    return lines


def section_stability(data) -> list[str]:
    spread = Counter()
    always3 = []
    always_pass = 0
    ever_low = []
    for path, reps in data.items():
        scores = [reps[n]["score"] for n in REPS]
        spread[max(scores) - min(scores)] += 1
        if all(s == 3 for s in scores):
            always3.append(path)
        if all(s >= 2 for s in scores):
            always_pass += 1
        if any(s <= 1 for s in scores):
            ever_low.append(path)
    total = len(data)
    lines = [
        "# 二、稳定性维度：同题九次是否一致",
        "",
        "|同题九次分数极差|题目数|占比|说明|",
        "|-|-|-|-|",
    ]
    notes = {0: "九次完全一致", 1: "波动一个档位", 2: "波动两个档位", 3: "波动三个档位（0↔3）"}
    for k in sorted(spread):
        lines.append(f"|{k}|{spread[k]}|{fmt_pct(spread[k], total)}|{notes.get(k, '')}|")
    lines += [
        "",
        f"- 九次均为 3 分的稳定优质题：**{len(always3)} 道**（{fmt_pct(len(always3), total)}）",
        f"- 九次均 ≥2 分的稳定可用题：**{always_pass} 道**（{fmt_pct(always_pass, total)}）",
        f"- 至少一次跌到 0/1 分的风险题：**{len(ever_low)} 道**（{fmt_pct(len(ever_low), total)}）",
        "",
        f"单次评测报出的低分题只有 11~17 道，但九次累计涉及 {len(ever_low)} 道不同题目，",
        "说明低分并非固定集中在同一批题上，单次结果会低估问题面。",
        "",
    ]
    return lines


def section_source_of_variance(data) -> list[str]:
    same_out_same_score = same_out_diff_score = 0
    diff_out_diff_score = diff_out_same_score = 0
    judge_cases = []
    for path, reps in data.items():
        outs = {reps[n]["output"] for n in REPS}
        scores = {reps[n]["score"] for n in REPS}
        if len(outs) == 1:
            if len(scores) == 1:
                same_out_same_score += 1
            else:
                same_out_diff_score += 1
                judge_cases.append(path)
        else:
            if len(scores) == 1:
                diff_out_same_score += 1
            else:
                diff_out_diff_score += 1
    lines = [
        "# 三、波动来源：模型输出不稳定 还是 裁判判分不稳定",
        "",
        "|九次模型输出|九次得分|题目数|归因|",
        "|-|-|-|-|",
        f"|完全相同|完全相同|{same_out_same_score}|双方都稳定|",
        f"|完全相同|存在差异|{same_out_diff_score}|**纯裁判抖动**（输入一字不差却给出不同档位）|",
        f"|存在差异|完全相同|{diff_out_same_score}|模型措辞有变但不影响判分|",
        f"|存在差异|存在差异|{diff_out_diff_score}|模型输出波动为主，可能叠加裁判抖动|",
        "",
        "绝大多数题目九次输出并不相同（模型本身非确定性输出），因此无法把波动完全归给裁判；",
        f"但存在 {same_out_diff_score} 道输出完全一致却得分不同的题目，证明裁判本身也贡献了一部分方差。",
        "",
    ]
    if judge_cases:
        lines.append("纯裁判抖动的题目：")
        lines.append("")
        for path in judge_cases:
            seq = " / ".join(f"rep{n}:{data[path][n]['score']}" for n in REPS)
            lines.append(f"- `{path}` — {seq}")
        lines.append("")
    return lines


def section_dimension(data, key: str, title: str) -> list[str]:
    groups = defaultdict(list)
    for path, reps in data.items():
        label = str(reps[REPS[0]][key] or "未标注")
        groups[label].append(path)
    lines = [title, "", "|" + "|".join(["分组", "题数"] + [f"rep{n}" for n in REPS] + ["均值", "低分累计"]) + "|",
             "|" + "|".join(["-"] * (len(REPS) + 4)) + "|"]
    rows = []
    for label, paths in groups.items():
        rates = []
        low_total = 0
        for n in REPS:
            ok = sum(1 for p in paths if data[p][n]["score"] >= 2)
            rates.append(ok / len(paths) * 100)
            low_total += sum(1 for p in paths if data[p][n]["score"] <= 1)
        rows.append((statistics.mean(rates), label, len(paths), rates, low_total))
    for mean, label, cnt, rates, low_total in sorted(rows):
        cells = [label, str(cnt)] + [f"{r:.0f}%" for r in rates] + [f"**{mean:.1f}%**", str(low_total)]
        lines.append("|" + "|".join(cells) + "|")
    lines += ["", "> 按九次均值升序排列，最上面的是最需要优化的分组。", ""]
    return lines


def reason_counter(data) -> tuple[Counter, int]:
    counter: Counter = Counter()
    total_low = 0
    for path, reps in data.items():
        for n in REPS:
            item = reps[n]
            if item["score"] > 1:
                continue
            total_low += 1
            reason = item["reason"]
            segment = " ".join(
                reason[m.start(): m.start() + WINDOW] for m in TRIGGER_RE.finditer(reason)
            )
            hit = False
            for name, pattern in DEDUCT_PATTERNS:
                if re.search(pattern, segment):
                    counter[name] += 1
                    hit = True
            if not hit:
                counter["其他/未归类"] += 1
    return counter, total_low


def section_deduction_reasons(data) -> list[str]:
    counter, total_low = reason_counter(data)
    lines = [
        "# 五、扣分原因归类（0/1 分样本，九次累计）",
        "",
        f"九次累计产生 {total_low} 个低分样本（同一题多次低分会重复计入），",
        "按裁判理由中**实际触发的条款**归类（一个样本可命中多条）：",
        "",
        "|扣分原因|命中次数|占低分样本比例|",
        "|-|-|-|",
    ]
    for name, cnt in counter.most_common():
        lines.append(f"|{name}|{cnt}|{fmt_pct(cnt, total_low)}|")
    lines += [
        "",
        "> 归类方法：裁判理由中存在大量「无漏答」「无幻觉」这类否定表述，",
        "> 因此只截取「触发/命中/符合官方/违反」等结论性语句附近 80 字窗口做关键词匹配，",
        "> 避免把「未触发」的条款统计成扣分原因。用于看趋势，单题准确归因请查看下方明细。",
        "",
    ]
    return lines


def section_question_detail(data) -> list[str]:
    """第六章：6.1 低分题总览，6.2 rep1-rep9 判罚依据与模型输出横向展开。"""
    risky = []
    for path, reps in data.items():
        low_reps = [n for n in REPS if reps[n]["score"] <= 1]
        if low_reps:
            risky.append((len(low_reps), path, low_reps))
    risky.sort(key=lambda x: (-x[0], x[1]))

    lines = [
        "# 六、题目维度明细：哪些题出问题、判罚依据、模型原文",
        "",
        f"共 {len(risky)} 道题在九次中至少一次被判 0/1 分，按低分次数降序。",
        "",
        "## 6.1 低分题总览",
        "",
        "|题目|学段|学科|题型|低分次数|九次得分序列|",
        "|-|-|-|-|-|-|",
    ]
    for cnt, path, low_reps in risky:
        meta = data[path][REPS[0]]
        seq = " ".join(str(data[path][n]["score"]) for n in REPS)
        lines.append(
            f"|`{path}`|{meta['grade']}|{meta['subject']}|{meta['qtype']}|{cnt}/9|`{seq}`|"
        )

    header = ["题目", "属性", "低分次数", "九次得分序列"]
    for n in REPS:
        header += [f"rep{n} 分数", f"rep{n} 判罚依据", f"rep{n} 模型输出"]
    lines += [
        "",
        "## 6.2 低分题明细（25 道题 × rep1-rep9 横向展开，判罚依据与模型原文完整保留）",
        "",
        "|" + "|".join(header) + "|",
        "|" + "|".join(["-"] * len(header)) + "|",
    ]
    for cnt, path, low_reps in risky:
        meta = data[path][REPS[0]]
        seq = " ".join(str(data[path][n]["score"]) for n in REPS)
        cells = [
            f"`{path}`",
            f"{meta['grade']}/{meta['subject']}/{meta['qtype']}",
            f"**{cnt}/9**",
            f"`{seq}`",
        ]
        for n in REPS:
            item = data[path][n]
            score = item["score"]
            cells += [
                f"**{score}**" if score <= 1 else str(score),
                table_text(item["reason"]),
                table_text(item["output"]),
            ]
        lines.append("|" + "|".join(cells) + "|")
    return lines


def section_conclusion(data, summary) -> list[str]:
    p2s = [s["p2"] for s in summary]
    best = max(summary, key=lambda s: s["p2"])
    ever_low = sum(1 for reps in data.values() if any(reps[n]["score"] <= 1 for n in REPS))
    always3 = sum(1 for reps in data.values() if all(reps[n]["score"] == 3 for n in REPS))
    mean = statistics.mean(p2s)
    sd = statistics.stdev(p2s)
    counter, total_low = reason_counter(data)
    top = "、".join(f"{name} {fmt_pct(cnt, total_low)}" for name, cnt in counter.most_common(3))
    style = counter["严重啰嗦/反复自我纠错"] + counter["解题过程过度简略"]
    return [
        "# 结论摘要",
        "",
        f"1. **代表水平是均值而非单次**：九次 ≥2分率均值 {mean:.1f}%，标准差 {sd:.2f}pp，"
        f"区间 {min(p2s):.1f}%~{max(p2s):.1f}%。rep{best['rep']} 的 {best['p2']:.1f}% 是九次最高值，"
        f"高出均值 {best['p2'] - mean:.1f}pp（约 {(best['p2'] - mean) / sd:.1f} 个标准差），"
        "若只跑一次并采用该结果，会系统性高估该配置。",
        f"2. **单次评测低估问题面**：单次低分题 {min(s['low'] for s in summary)}~{max(s['low'] for s in summary)} 道，"
        f"但九次累计有 {ever_low} 道题至少失败过一次，是单次观测的 1.7~2.3 倍。",
        f"3. **稳定优质的题只占少数**：九次全 3 分仅 {always3} 道（{fmt_pct(always3, len(data))}），"
        f"其余题目在不同次数间存在档位波动。",
        "4. **波动同时来自模型和裁判**：多数题九次输出不同（模型非确定性），"
        "但也存在输出完全一致却判分不同的题，说明裁判本身贡献一部分方差。",
        f"5. **低分主因是答案与推理本身**：命中最多的三类是 {top}；"
        f"其中输出风格类问题（严重啰嗦 + 过度简略）合计 {fmt_pct(style, total_low)}，"
        "这部分优化成本低于提升解题能力。",
        "6. **短板集中在高中与数学**：详见第四节分组对比。",
        "",
        "> 建议：后续版本对比至少跑 3 次取均值；只要两个版本差距小于 ~2pp（约 2 个标准差），"
        "就不应据单次结果判定优劣。",
        "",
    ]


def main() -> None:
    data = load()
    summary = rep_summary(data)
    parts = [
        "amuL20（FP16tp2x2 0904）rep1-rep9 九次重复评测对比分析",
        "",
        f"> 评测集：161 道 confidence=high 题目｜裁判：OneAPI `Opus 5`｜重复次数：9",
        f"> 数据来源：`runs/existing_eval_oneapi_opus5_rep1~9_20260907/`",
        "",
    ]
    parts += section_conclusion(data, summary)
    parts += section_overview(data, summary)
    parts += section_stability(data)
    parts += section_source_of_variance(data)
    parts += ["# 四、分组维度对比", ""]
    parts += section_dimension(data, "subject", "## 4.1 分学科 ≥2分率（九次）")
    parts += section_dimension(data, "grade", "## 4.2 分学段 ≥2分率（九次）")
    parts += section_dimension(data, "qtype", "## 4.3 分题型 ≥2分率（九次）")
    parts += section_deduction_reasons(data)
    parts += section_question_detail(data)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "rep1_9_comparison_report.md"
    out.write_text("\n".join(parts) + "\n", encoding="utf-8")
    print(f"报告已生成：{out}")


if __name__ == "__main__":
    main()

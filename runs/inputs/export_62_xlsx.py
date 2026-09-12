#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把对比报告 6.2 低分题明细导出为 xlsx（25 行 × 31 列，全文不截断）。"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_rep_comparison import REPS, load  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "rep1_9_comparison" / "rep1-9-badcase汇总.xlsx"


def main() -> None:
    data = load()
    risky = []
    for path, reps in data.items():
        low = [n for n in REPS if reps[n]["score"] <= 1]
        if low:
            risky.append((len(low), path, low))
    risky.sort(key=lambda x: (-x[0], x[1]))

    wb = Workbook()
    ws = wb.active
    ws.title = "6.2 低分题明细"

    header = ["题目", "属性", "低分次数", "九次得分序列"]
    for n in REPS:
        header += [f"rep{n} 分数", f"rep{n} 判罚依据", f"rep{n} 模型输出"]
    ws.append(header)

    for cnt, path, low in risky:
        meta = data[path][REPS[0]]
        row = [
            path,
            f"{meta['grade']}/{meta['subject']}/{meta['qtype']}",
            f"{cnt}/9",
            " ".join(str(data[path][n]["score"]) for n in REPS),
        ]
        for n in REPS:
            item = data[path][n]
            row += [item["score"], item["reason"], item["output"]]
        ws.append(row)

    head_fill = PatternFill("solid", fgColor="DDEBF7")
    low_fill = PatternFill("solid", fgColor="FFC7CE")

    # 模型输出以 "==**解题过程**==" 开头，openpyxl 默认把以 = 开头的字符串当成公式写入，
    # 会导致 Excel 打开后显示为错误公式、且程序读取时取到 None，这里强制标记为文本。
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.data_type = "s"

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)

    widths = {1: 34, 2: 22, 3: 10, 4: 18}
    for idx, name in enumerate(header, start=1):
        if idx in widths:
            width = widths[idx]
        elif name.endswith("分数"):
            width = 8
        else:
            width = 60
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for idx, name in enumerate(header, start=1):
            if name.endswith("分数"):
                cell = row[idx - 1]
                if isinstance(cell.value, int) and cell.value <= 1:
                    cell.fill = low_fill
                    cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical="top", horizontal="center")
        ws.row_dimensions[row[0].row].height = 120

    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{ws.max_row}"
    wb.save(OUT)
    print(f"已导出：{OUT}")
    print(f"行数（含表头）：{ws.max_row}｜列数：{len(header)}")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""识别结果写入 Excel 工作簿（XlsxWriter / RustPy-XlsxWriter）。"""
from __future__ import annotations

import re
from typing import Dict, List, Tuple

from .fast_excel import CellRange, FastWorkbook, FastWorksheet

# 疑似数字的文本（含千分位、百分号、负号）转数值，便于在 Excel 里校验求和
_NUM_RE = re.compile(r"^-?\d{1,3}(,\d{3})*(\.\d+)?$")
_NUM_PLAIN_RE = re.compile(r"^-?\d+(\.\d+)?$")
_PCT_RE = re.compile(r"^(-?\d+(?:\.\d+)?)%$")


def coerce_number(text: str):
    """纯数字文本转 int/float，其余原样返回（百分数按数值写入）。"""
    t = (text or "").strip()
    if not t:
        return t
    m = _PCT_RE.match(t)
    if m:
        return round(float(m.group(1)) / 100.0, 6)
    if _NUM_PLAIN_RE.match(t):
        return int(t) if re.match(r"^-?\d+$", t) else float(t)
    if _NUM_RE.match(t):
        return int(t.replace(",", "")) if "." not in t else float(t.replace(",", ""))
    return t


def safe_sheet_name(name: str, used: Dict[str, int]) -> str:
    """规范化 sheet 名（Excel 限制 31 字符与部分字符），冲突自动加序号。"""
    cleaned = re.sub(r"[\\/*?:\[\]]", "_", (name or "Sheet")).strip() or "Sheet"
    cleaned = cleaned[:31]
    if cleaned not in used:
        used[cleaned] = 1
        return cleaned
    used[cleaned] += 1
    suffix = f"({used[cleaned]})"
    return cleaned[:31 - len(suffix)] + suffix


def write_workbook(out_path: str, pages: List[Dict]) -> None:
    """写入新工作簿，保留识别值、标题和合并区；以行优先写入。"""
    workbook = FastWorkbook()
    used: Dict[str, int] = {}
    for page in pages:
        rows: List[List[str]] = page.get("rows") or []
        merges: List[Tuple[int, int, int, int]] = page.get("merges") or []
        title = (page.get("title") or "").strip()
        sheet_rows: List[List] = []
        merge_ranges: List[CellRange] = []
        offset = 1 if title else 0

        if title:
            n_cols = max((len(row) for row in rows), default=1)
            sheet_rows.append([title] + [None] * (n_cols - 1))
            if n_cols > 1:
                merge_ranges.append(CellRange(1, 1, 1, n_cols))

        for row in rows:
            sheet_rows.append([coerce_number(str(value)) for value in row])

        for r, c, rspan, cspan in merges:
            if rspan > 1 or cspan > 1:
                merge_ranges.append(CellRange(
                    r + 1 + offset, c + 1, r + rspan + offset, c + cspan))

        name = safe_sheet_name(page.get("name") or "Sheet", used)
        workbook.add_sheet(FastWorksheet(name, sheet_rows, merge_ranges))

    if not workbook.sheetnames:
        workbook.add_sheet(FastWorksheet("Sheet"))
    workbook.save(out_path)

# -*- coding: utf-8 -*-
"""xlsx 模板：用命名区域定义"结构区/数字区"的固定格式表格模板。

设计（用户主导）：
  - 模板就是一个普通 .xlsx 文件，可用 Excel 直接打开查看与编辑；
  - 命名区域约定（支持多区域，逗号分隔的联合区域也支持）：
      数字区域 / NUM*   → 每月需要 OCR 并填入的格子
      结构区域 / STRUCT*→ 冻结文本格（用于和识别结果做结构匹配/校验）
      中缝表头 / MID*   → 表内插入的第二段表头所在行（整行冻结）
  - 勾稽关系：额外的"校验"工作表里写普通 Excel 公式（如
      =数据!G8-SUM(数据!G3:G7)），套用模板导出时公式原样保留
      （Excel 打开即算），本地也做一次轻量求值，非零即告警。
  - "多页"= 工作簿里的多个数据工作表（校验表除外）。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.utils import (column_index_from_string, get_column_letter,
                            range_boundaries)
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

Cell = Tuple[int, int]      # (row, col) 0-based
Rect = Tuple[int, int, int, int]   # (r0, c0, r1, c1) 闭区间 0-based

CHECK_SHEET = "校验"
GEO_SHEET = "几何"      # 隐藏表：存每页行列比例（程序用，用户无需关心）
BALANCE_SHEET = "勾稽"  # 逐行勾稽差额公式（用户在差额列加条件格式即可）

_NUM_NAME = re.compile(r"^(数字|NUM)", re.I)
_STRUCT_NAME = re.compile(r"^(结构|STRUCT)", re.I)
_MID_NAME = re.compile(r"^(中缝|MID)", re.I)


# ----------------------------------------------------------------------
# 命名区域解析 / 生成
# ----------------------------------------------------------------------
def _parse_refs(attr_text: str, default_sheet: str) -> List[Tuple[str, Rect]]:
    """解析命名区域的 attr_text（可能多段逗号分隔），返回 [(sheet, rect)]。"""
    out: List[Tuple[str, Rect]] = []
    for part in str(attr_text or "").split(","):
        part = part.strip().replace("$", "")
        if not part:
            continue
        sheet = default_sheet
        m = re.match(r"^'?([^'!]+)'?!(.+)$", part)
        if m:
            sheet, part = m.group(1), m.group(2)
        try:
            c0, r0, c1, r1 = range_boundaries(part)
        except Exception:
            continue
        if None in (r0, c0, r1, c1):
            continue
        out.append((sheet, (int(r0) - 1, int(c0) - 1, int(r1) - 1, int(c1) - 1)))
    return out


def _rect_cells(rect: Rect) -> Set[Cell]:
    r0, c0, r1, c1 = rect
    return {(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)}


def cells_to_rects(cells: Set[Cell]) -> List[Rect]:
    """把单元格集合贪心合并为矩形列表（供写命名区域用）。"""
    if not cells:
        return []
    remaining = set(cells)
    rects: List[Rect] = []
    while remaining:
        r0, c0 = min(remaining)
        # 向右扩展
        c1 = c0
        while (r0, c1 + 1) in remaining:
            c1 += 1
        # 向下扩展（要求整行都还在）
        r1 = r0
        while all((r1 + 1, c) in remaining for c in range(c0, c1 + 1)):
            r1 += 1
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                remaining.discard((r, c))
        rects.append((r0, c0, r1, c1))
    return sorted(rects)


def _rects_to_attr(sheet: str, rects: List[Rect]) -> str:
    parts = []
    for (r0, c0, r1, c1) in rects:
        q = f"'{sheet}'" if re.search(r"[^A-Za-z0-9_\u4e00-\u9fff]", sheet) else sheet
        parts.append(f"{q}!${get_column_letter(c0 + 1)}${r0 + 1}:"
                     f"${get_column_letter(c1 + 1)}${r1 + 1}")
    return ",".join(parts)


# ----------------------------------------------------------------------
# 模板对象
# ----------------------------------------------------------------------
@dataclass
class XlsxSheetPage:
    """工作簿里的一张数据表（= 模板的一页）。"""

    page_name: str
    rows: List[List[str]] = field(default_factory=list)
    num_cells: Set[Cell] = field(default_factory=set)
    struct_cells: Set[Cell] = field(default_factory=set)
    mid_rows: Set[int] = field(default_factory=set)
    merges: List[List[int]] = field(default_factory=list)   # [r,c,rspan,cspan] 0 基
    # 几何比例（建模板时从校对页写入隐藏表）：套用时不依赖内部细线检测，
    # 只用表格外框（最稳）+ 比例铺格。空 = 无几何信息，退回按检测网格。
    col_fracs: List[float] = field(default_factory=list)
    row_fracs: List[float] = field(default_factory=list)

    # 命名区域信息（载入时填）：{"数字区域": [rect,...], ...}
    named_ranges: Dict[str, List[Rect]] = field(default_factory=dict)

    @property
    def n_rows(self) -> int:
        return len(self.rows)

    @property
    def n_cols(self) -> int:
        return max((len(r) for r in self.rows), default=0)

    @property
    def header_rows(self) -> int:
        first_num = min((r for r, _c in self.num_cells), default=self.n_rows)
        return max(0, min(first_num, self.n_rows))

    @property
    def value_cols(self) -> List[int]:
        return sorted({c for _r, c in self.num_cells})

    def _header_text(self, r: int, c: int) -> str:
        """表头文本（合并格取锚点文本并传播到被跨列）。"""
        for m in self.merges:
            try:
                mr, mc, mrs, mcs = (int(x) for x in m)
            except (TypeError, ValueError):
                continue
            if mr <= r < mr + mrs and mc <= c < mc + mcs:
                if r < len(self.rows) and mc < len(self.rows[r]):
                    return str(self.rows[r][mc] or "").strip()
                return ""
        if r < len(self.rows) and c < len(self.rows[r]):
            return str(self.rows[r][c] or "").strip()
        return ""

    def col_labels(self) -> List[str]:
        """逐列显示名：表头行 0..header_rows-1 逐层拼接（合并格展开）
        + 中缝段补充（组·收方/付方）。"""
        n = self.n_cols
        hr = self.header_rows
        labels: List[str] = []
        for c in range(n):
            parts: List[str] = []
            for r in range(hr):
                t = self._header_text(r, c)
                if t and (not parts or parts[-1] != t):
                    parts.append(t)
            labels.append("·".join(parts))
        # 中缝段标签：叶子行（收方/付方）与上半段组名拼为 "上期余额·收方"
        if self.mid_rows:
            mr = max(self.mid_rows)          # 叶子行（第二行表头）
            for c in range(n):
                leaf = self._header_text(mr, c)
                if leaf and leaf not in ("科目代码", "科目名称"):
                    group = labels[c].split("·")[0] if labels[c] else ""
                    labels[c] = (f"{group}·{leaf}" if group and group != leaf
                                 else leaf)
        return labels

    def struct_texts(self) -> Dict[Cell, str]:
        """结构区文本（用于与识别结果做匹配校验）。"""
        out = {}
        for (r, c) in self.struct_cells:
            if r < len(self.rows) and c < len(self.rows[r]):
                t = str(self.rows[r][c] or "").strip()
                if t:
                    out[(r, c)] = t
        return out


@dataclass
class XlsxTemplate:
    """一个模板文档（xlsx），可含多张数据表（多页）。"""

    name: str
    path: Path
    pages: List[XlsxSheetPage] = field(default_factory=list)
    # 校验表公式：[{"cell": "B2", "raw": "=数据!G8-SUM(数据!G3:G7)"}]
    check_formulas: List[Dict[str, str]] = field(default_factory=list)
    # 命名区域是否由启发式推断（手搓且未定义区域）——载入时提示用户
    inferred_regions: bool = False

    @property
    def rows(self) -> List[List[str]]:
        return self.pages[0].rows if self.pages else []

    @property
    def value_cols(self) -> List[int]:
        return self.pages[0].value_cols if self.pages else []


# ----------------------------------------------------------------------
# 读取
# ----------------------------------------------------------------------
def _cell_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def load_xlsx_template(path: Path) -> XlsxTemplate:
    """读取 xlsx 模板：工作簿内每个非校验表 = 一页；命名区域定义结构/数字区。"""
    path = Path(path)
    wb = load_workbook(path, data_only=False)

    # 命名区域 → 每张表的数字/结构中缝单元格集合
    num_by_sheet: Dict[str, Set[Cell]] = {}
    struct_by_sheet: Dict[str, Set[Cell]] = {}
    mid_by_sheet: Dict[str, Set[int]] = {}

    def _add_region(nm: str, default_sheet: str, attr_text: str) -> None:
        for sheet, rect in _parse_refs(attr_text, default_sheet):
            if not sheet:
                continue
            if _NUM_NAME.match(nm):
                num_by_sheet.setdefault(sheet, set()).update(_rect_cells(rect))
            elif _STRUCT_NAME.match(nm):
                struct_by_sheet.setdefault(sheet, set()).update(_rect_cells(rect))
            elif _MID_NAME.match(nm):
                mid_by_sheet.setdefault(sheet, set()).update(
                    range(rect[0], rect[2] + 1))

    # ① 工作簿级命名区域（工具生成/Excel 里跨表定义）
    for dn in wb.defined_names.values():
        _add_region(dn.name or "", "", str(dn.attr_text or ""))
    # ② 工作表级命名区域（用户在 Excel 里按单表作用域定义；名字可各表重复，
    #    这对"一本 xlsx 多张表"的用法更自然）
    for ws in wb.worksheets:
        try:
            sheet_defs = getattr(ws, "defined_names", None) or {}
        except Exception:
            sheet_defs = {}
        for dn in list(sheet_defs.values()):
            _add_region(dn.name or "", ws.title, str(dn.attr_text or ""))

    # 每张表的命名区域明细（界面/校验共用同一事实来源）
    ranges_by_sheet: Dict[str, Dict[str, List[Rect]]] = {}
    for dn in wb.defined_names.values():
        nm = dn.name or ""
        for sheet, rect in _parse_refs(dn.attr_text, ""):
            if not sheet:
                continue
            ranges_by_sheet.setdefault(sheet, {}).setdefault(nm, []).append(rect)

    tpl = XlsxTemplate(name=path.stem, path=path)
    inferred = False
    for ws in wb.worksheets:
        if ws.title in (CHECK_SHEET, GEO_SHEET, BALANCE_SHEET):
            continue
        rows: List[List[str]] = []
        n_cols = 0
        for r in ws.iter_rows():
            vals = [_cell_text(c.value) for c in r]
            if any(vals):
                n_cols = max(n_cols, len(vals))
            rows.append(vals)
        while rows and not any(rows[-1]):
            rows.pop()
        if not rows:
            continue
        for r in rows:
            r += [""] * (n_cols - len(r))
        mid = mid_by_sheet.get(ws.title, set())
        rows = [r for i, r in enumerate(rows) if i not in mid or True]
        num = {c for c in num_by_sheet.get(ws.title, set())
               if c[0] < len(rows)}
        struct = {c for c in struct_by_sheet.get(ws.title, set())
                  if c[0] < len(rows)}
        if not num:
            # 手工制作的 xlsx 未写命名区域：启发式推断数字区（并标记）
            inferred = True
            hr = 0
            for r in rows[:4]:
                if any(_looks_numeric(v) for v in r):
                    break
                hr += 1
            mid_rows = set()
            for r in range(hr, len(rows)):
                head = "".join(rows[r])
                if "科目代码" in head and "科目名称" in head:
                    mid_rows.add(r)
                    mid_rows.add(r + 1)      # 段内子表头行（借/贷 或 收/付）
            num = {(r, c) for r in range(len(rows))
                   for c in range(n_cols)
                   if r >= hr and r not in mid_rows
                   and c >= 2 and _looks_numeric(rows[r][c])}
            if not num:
                # 启发式第二档：数字尚未填的空白模板——取数据区除前两列
                # 与末列（常见为重复的代码列）之外的格
                num = {(r, c) for r in range(hr, len(rows))
                       for c in range(2, max(2, n_cols - 1))
                       if r not in mid_rows}
        merges: List[List[int]] = []
        for mr in ws.merged_cells.ranges:
            c0, r0, c1, r1 = mr.min_col, mr.min_row, mr.max_col, mr.max_row
            if r1 - r0 + 1 > 1 or c1 - c0 + 1 > 1:
                merges.append([r0 - 1, c0 - 1, r1 - r0 + 1, c1 - c0 + 1])
        tpl.pages.append(XlsxSheetPage(
            page_name=ws.title, rows=rows, num_cells=num,
            struct_cells=struct, mid_rows=set(mid), merges=merges,
            named_ranges=ranges_by_sheet.get(ws.title, {})))

    tpl.inferred_regions = inferred
    if CHECK_SHEET in wb.sheetnames:
        ws = wb[CHECK_SHEET]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    tpl.check_formulas.append({"cell": cell.coordinate,
                                               "raw": v})
    # 读取几何比例（隐藏表）
    if GEO_SHEET in wb.sheetnames:
        try:
            geo = json.loads(str(wb[GEO_SHEET]["A1"].value or "{}"))
        except Exception:
            geo = {}
        for pg in tpl.pages:
            g = geo.get(pg.page_name) or {}
            pg.col_fracs = [float(x) for x in g.get("col_fracs", [])]
            pg.row_fracs = [float(x) for x in g.get("row_fracs", [])]
    wb.close()
    return tpl


def _looks_numeric(v) -> bool:
    s = str(v or "").strip().replace(",", "")
    if not s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


# ----------------------------------------------------------------------
# 生成（从一页已校对的识别结果）
# ----------------------------------------------------------------------
def save_sheet_to_workbook(path: Path, sheet_name: str, rows: List[List[str]],
                           value_cols: List[int], header_rows: int,
                           merges: List[List[int]],
                           mid_rows: Optional[Set[int]] = None,
                           replace: bool = False,
                           col_fracs: Optional[List[float]] = None,
                           row_fracs: Optional[List[float]] = None) -> None:
    """向 xlsx 模板写入/追加一页（表 + 命名区域）。

    replace=True 时重建整个工作簿（首个版本）；否则追加一张数据表。
    """
    mid_rows = set(mid_rows or set())
    path = Path(path)
    if replace or not path.is_file():
        wb = Workbook()
        wb.remove(wb.active)
    else:
        wb = load_workbook(path, data_only=False)
    # 表名去重
    name = sheet_name or "第1页"
    base = name
    i = 2
    while name in wb.sheetnames:
        name = f"{base}_{i}"
        i += 1
    ws = wb.create_sheet(title=name)

    n_rows = len(rows)
    n_cols = max((len(r) for r in rows), default=0)
    for r in range(n_rows):
        for c in range(n_cols):
            v = rows[r][c] if c < len(rows[r]) else ""
            ws.cell(row=r + 1, column=c + 1,
                    value=(float(v.replace(",", "")) if _looks_numeric(v)
                           and r >= header_rows and r not in mid_rows else v))
    for m in merges or []:
        try:
            r, c, rspan, cspan = (int(x) for x in m)
        except (TypeError, ValueError):
            continue
        if rspan > 1 or cspan > 1:
            try:
                ws.merge_cells(start_row=r + 1, start_column=c + 1,
                               end_row=r + rspan, end_column=c + cspan)
            except Exception:
                pass

    num_cells = {(r, c) for r in range(n_rows) for c in value_cols
                 if r >= header_rows and r not in mid_rows and c < n_cols}
    all_cells = {(r, c) for r in range(n_rows) for c in range(n_cols)}
    struct_cells = all_cells - num_cells

    def add_name(nm: str, cells: Set[Cell]) -> None:
        """把本页单元格并入命名区域（多页共用同一名称，逗号分隔联合区域）。

        注意必须"读旧值 + 追加"，否则第二页会整体覆盖第一页的区域。
        """
        rects = cells_to_rects(cells)
        if not rects:
            return
        parts = []
        existing = wb.defined_names.get(nm)
        if existing is not None and existing.attr_text:
            parts.append(str(existing.attr_text))
        parts.append(_rects_to_attr(name, rects))
        attr = ",".join(x for x in parts if x)
        try:
            if existing is not None:
                del wb.defined_names[nm]
            wb.defined_names.add(DefinedName(nm, attr_text=attr))
        except Exception:
            pass

    # 几何比例存隐藏表（套用时按外框+比例铺格，摆脱内部细线检测）
    if col_fracs and row_fracs:
        if GEO_SHEET not in wb.sheetnames:
            ws_geo = wb.create_sheet(GEO_SHEET)
            ws_geo.sheet_state = "veryHidden"
        else:
            ws_geo = wb[GEO_SHEET]
        try:
            geo = json.loads(str(ws_geo["A1"].value or "{}"))
        except Exception:
            geo = {}
        geo[name] = {"col_fracs": [round(float(x), 5) for x in col_fracs],
                     "row_fracs": [round(float(y), 5) for y in row_fracs]}
        ws_geo["A1"] = json.dumps(geo, ensure_ascii=False)

    add_name("数字区域", num_cells)
    add_name("结构区域", struct_cells)
    # 段内表头（中缝）行已包含在"结构区域"中，不再单独建区域

    wb.save(path)
    wb.close()


# ----------------------------------------------------------------------
# 匹配 / 套用
# ----------------------------------------------------------------------
def _match_text(a: str, b: str) -> bool:
    """结构区文本匹配：去空白后相同，或数字型代码逐步比对。"""
    a = re.sub(r"\s+", "", a or "")
    b = re.sub(r"\s+", "", b or "")
    if not a or not b:
        return False
    if a == b:
        return True
    if _looks_numeric(a) and _looks_numeric(b):
        return a.replace(",", "") == b.replace(",", "")
    # 中文长文本容忍零星误读
    if len(a) >= 4 and len(b) >= 4:
        same = sum(1 for x, y in zip(a, b) if x == y)
        return same / max(len(a), len(b)) >= 0.8
    return False


def geometry_page_cells(structure, pg: XlsxSheetPage):
    """按模板几何（外框 + 比例）生成单元格 {（r,c): (x, y, w, h)}。

    表格外框用检测到的第一条/最后一条横线与竖线（比内部细线稳得多），
    内部按模板比例铺格——这是 xlsx 模板"稳"的关键：内部线检测不稳时
    仍然能切出正确行列。
    """
    xs, ys = structure.xs, structure.ys
    if len(xs) < 2 or len(ys) < 2 or not pg.col_fracs or not pg.row_fracs:
        return {}
    left, right = xs[0], xs[-1]
    top, bottom = ys[0], ys[-1]
    # 外框并上"文字内容框"：列/行线检测被截断（只检出左半）时，
    # 用 det 文字的包围盒兜底，避免铺格整体压缩
    items = getattr(structure, "det_items", None) or []
    if items:
        tx0 = min(min(pt[0] for pt in b) for b, _t, _s in items)
        ty0 = min(min(pt[1] for pt in b) for b, _t, _s in items)
        tx1 = max(max(pt[0] for pt in b) for b, _t, _s in items)
        ty1 = max(max(pt[1] for pt in b) for b, _t, _s in items)
        left, top = min(left, tx0), min(top, ty0)
        right, bottom = max(right, tx1), max(bottom, ty1)
    col_x = [int(round(left + f * (right - left))) for f in pg.col_fracs]
    row_y = [int(round(top + f * (bottom - top))) for f in pg.row_fracs]
    cells = {}
    for r in range(min(pg.n_rows, len(row_y) - 1)):
        for c in range(min(pg.n_cols, len(col_x) - 1)):
            cells[(r, c)] = (col_x[c], row_y[r],
                             col_x[c + 1] - col_x[c], row_y[r + 1] - row_y[r])
    return cells


def _texts_in_cells(structure, cells) -> Dict[Cell, str]:
    """det 文本按中心点落入几何格：{（r,c): 拼接文本}。"""
    merged, _pos = texts_and_centers_in_cells(structure, cells)
    return merged


def texts_and_centers_in_cells(structure, cells):
    """同上，另返回每格首个文本条的中心像素 {(r,c): (cx, cy)}（供锚点对齐）。"""
    out: Dict[Cell, List] = {}
    for box, text, _s in (structure.det_items or []):
        cx = sum(p[0] for p in box) / len(box)
        cy = sum(p[1] for p in box) / len(box)
        for (r, c), (x, y, w, h) in cells.items():
            if x <= cx < x + w and y <= cy < y + h:
                ys_ = [p[1] for p in box]
                xs_ = [p[0] for p in box]
                out.setdefault((r, c), []).append(
                    (sum(ys_) / len(ys_), sum(xs_) / len(xs_), str(text)))
                break
    merged: Dict[Cell, str] = {}
    centers: Dict[Cell, Tuple[float, float]] = {}
    for key, rows in out.items():
        rows.sort(key=lambda e: (e[0], e[1]))
        merged[key] = "".join(t for _, _, t in rows)
        centers[key] = (rows[0][1], rows[0][0])
    return merged, centers


def refine_cells_by_anchors(pg: XlsxSheetPage, cells, centers) -> Dict:
    """用"匹配上的结构区文本"做行列线性校正（文字锚点对齐）。

    纯比例铺格假设检测外边界与模板一致；实际检测的表格外框可能少含
    多含边缘列（相差一行/列），导致整表平移。这里以匹配成功（文本一致）
    的格子为锚点，最小二乘拟合"模板像素位置 → 实际像素位置"的线性映射
    （吸收平移与缩放），校正后再铺格。
    """
    if not cells or not centers:
        return cells
    xs0, xs1, ys0, ys1 = [], [], [], []
    for (r, c), (dcx, dcy) in centers.items():
        box = cells.get((r, c))
        if box is None:
            continue
        x, y, w, h = box
        xs0.append(x + w / 2.0)
        xs1.append(float(dcx))
        ys0.append(y + h / 2.0)
        ys1.append(float(dcy))
    if len(xs0) < 4:
        return cells
    try:
        import numpy as np
        ax, bx = np.polyfit(xs0, xs1, 1)
        ay, by = np.polyfit(ys0, ys1, 1)
        if not (0.7 <= ax <= 1.4 and 0.7 <= ay <= 1.4):
            return cells          # 拟合异常，保守用原铺格
    except Exception:
        return cells
    out = {}
    for key, (x, y, w, h) in cells.items():
        nx0 = ax * x + bx
        nx1 = ax * (x + w) + bx
        ny0 = ay * y + by
        ny1 = ay * (y + h) + by
        out[key] = (int(round(nx0)), int(round(ny0)),
                    int(round(nx1 - nx0)), int(round(ny1 - ny0)))
    return out


def match_xlsx_template(templates: List[XlsxTemplate],
                        structure,
                        n_rows: int, n_cols: int,
                        force_doc: str = "",
                        force_page: str = ""):
    """在 xlsx 模板（文档×表）里选最佳匹配。

    匹配用"结构区文本命中率"：
      1) 几何映射（优先）：表格外框 + 模板比例铺格，把 det 文本按位置
         落格后比对结构区——内部细线检测不稳也能匹配，尺寸差异不再是
         硬门槛（模板决定行列数）；
      2) 尺寸精确匹配：无几何信息的模板退回按检测网格比对（严格尺寸）。
    返回 (文档, 页, 命中率, 备注, cells)；cells 为命中的铺格（几何）或
    检测网格 cells 映射，供套用阶段使用。
    """
    if not templates:
        return None, None, 0.0, "无模板", {}
    best = (None, None, 0.0, {})
    notes = []
    for doc in templates:
        if force_doc and doc.name != force_doc:
            continue
        for pg in doc.pages:
            if force_page and pg.page_name != force_page:
                continue
            texts = pg.struct_texts()
            if not texts:
                continue
            # ---- 几何路径 ----
            if pg.col_fracs and pg.row_fracs and structure.det_items:
                cells = geometry_page_cells(structure, pg)
                if cells:
                    mapped, centers = texts_and_centers_in_cells(structure, cells)
                    hit = sum(1 for cell, t in texts.items()
                              if _match_text(t, mapped.get(cell, "")))
                    score = hit / len(texts)
                    if score >= 0.25:
                        # 文字锚点校正：匹配成功的格子拟合线性映射后重铺
                        good = {k: v for k, v in centers.items()
                                if k in texts and _match_text(texts[k],
                                                              mapped.get(k, ""))}
                        cells = refine_cells_by_anchors(pg, cells, good)
                        mapped2, _c2 = texts_and_centers_in_cells(structure, cells)
                        score2 = sum(1 for cell, t in texts.items()
                                     if _match_text(t, mapped2.get(cell, ""))) / len(texts)
                        if score2 > score:
                            score = score2
                    if score > best[2]:
                        best = (doc, pg, score, cells)
                    continue
            # ---- 尺寸精确路径（无几何信息）----
            if (pg.n_rows, pg.n_cols) != (n_rows, n_cols):
                notes.append(f"{doc.name}·{pg.page_name} "
                             f"{pg.n_rows}×{pg.n_cols}≠{n_rows}×{n_cols}")
                continue
            det_cells = {(c.row, c.col): (c.x, c.y, c.w, c.h)
                         for c in structure.cells}
            mapped = _texts_in_cells(structure, det_cells)
            hit = sum(1 for cell, t in texts.items()
                      if _match_text(t, mapped.get(cell, "")))
            score = hit / len(texts)
            if score > best[2]:
                best = (doc, pg, score, det_cells)
    doc, pg, score, cells = best
    if doc is None:
        note = ("尺寸不符：" + "；".join(notes[:3])) if notes else             ("指定页未找到" if force_page else "无候选")
        return None, None, 0.0, note, {}
    if score < 0.55:
        if force_page:
            # 用户显式指定了页：仍套用，但明确告警命中率低
            return doc, pg, score, f"结构区命中率仅 {score:.0%}（按指定页填充，请重点核对）", cells
        return None, None, score, f"结构区命中率仅 {score:.0%}", {}
    return doc, pg, score, "", cells


def _key_code(t: str) -> str:
    return "".join(ch for ch in (t or "") if ch.isdigit() or ch.isalpha())


def _key_match(a: str, b: str) -> bool:
    """科目键匹配（包含式，容忍 OCR 把前两列粘成一格）。

    OCR 常把"170"与"中央预算支出"识别进同一格（要么代码格含名称，
    要么名称格含代码）。这里对纯代码做子串判定：短代码是长串的子串
    且长度差不过分，即视为同一行（"170" ∈ "170中央预算支出" ✓）。
    """
    ra, rb = _key_code(a), _key_code(b)
    if not ra or not rb:
        return False
    if ra == rb:
        return True
    # 纯数字代码：子串匹配（要求至少 3 位，避免 "1" 命中一切）
    da = "".join(ch for ch in ra if ch.isdigit())
    db = "".join(ch for ch in rb if ch.isdigit())
    if da and db:
        if len(da) >= 3 and da in db and len(db) - len(da) <= len(db) * 0.6:
            return True
        if len(db) >= 3 and db in da and len(da) - len(db) <= len(da) * 0.6:
            return True
    # 宽松近似（容忍零星误读）
    if abs(len(ra) - len(rb)) > 1:
        return False
    n = min(len(ra), len(rb))
    same = sum(1 for i in range(n) if ra[i] == rb[i])
    return same / max(len(ra), len(rb)) >= 0.75


def _row_key_match(tpl_code: str, tpl_name: str,
                   ocr_code: str, ocr_name: str) -> bool:
    """整行键匹配：代码包含式命中，或名称互相包含，或代码+名称合起来命中。

    针对 OCR 把"170 中央预算支出"粘进一格的情况：
      · 模板代码 "170" 出现在 OCR 文本里，且模板名称也出现在 OCR 文本里 → 命中
      · 或名称互相包含（中央预算支出 ∈ "170中央预算支出"）
    """
    combined = f"{ocr_code}{ocr_name}"
    # ① 代码对"整行文本"的包含式命中（170 ∈ "170中央预算支出"）
    if tpl_code and _key_match(tpl_code, combined):
        if tpl_code in combined or len(_key_code(tpl_code)) >= 3:
            return True
    # ② 名称对"整行文本"的包含式命中（中央预算支出 ∈ "170中央预算支出"）
    if tpl_name and len(tpl_name) >= 3 and (
            tpl_name in combined or combined in tpl_name):
        return True
    # ③ 常规分列匹配（代码/名称各在其列）
    if _key_match(tpl_code, ocr_code):
        return True
    if tpl_name and _match_text(tpl_name, ocr_name):
        return True
    return False


def build_row_mapping(tpl_page: XlsxSheetPage,
                      geo_keys: Dict[int, Tuple[str, str]]):
    """VLOOKUP 式行对齐：把模板数据行映射到照片检出的行。

    geo_keys: {几何行号: (科目代码, 科目名称)}（来自照片上的 det 文本）。
    返回 (row_map, unmatched_tpl, extra_geo)：
      row_map      {模板行 -> 几何行}（代码/名称匹配上的）
      unmatched_tpl[(模板行, 代码, 名称)]  照片里找不到的模板行
      extra_geo   [(几何行, 代码)]         照片里多出的行
    合计行等空代码行不参与匹配（由调用方按邻近行偏移对齐）。
    """
    num_rows = {r for (r, _c) in tpl_page.num_cells}
    tpl_rows = []
    for r in range(tpl_page.n_rows):
        # 数据行 = "数字区域"覆盖的行；段内表头（科目代码/科目名称行、
        # 中缝表头）没有数字格，自动排除，不依赖单独的中缝区域
        if r < tpl_page.header_rows or (num_rows and r not in num_rows):
            continue
        code = str(tpl_page.rows[r][0] if tpl_page.rows[r] else "").strip()
        name = str(tpl_page.rows[r][1] if len(tpl_page.rows[r]) > 1 else "").strip()
        if _key_code(code):
            tpl_rows.append((r, code, name))

    used_geo = set()
    row_map: Dict[int, int] = {}
    # 第一轮：整行键匹配（含"170中央预算支出"粘格场景）
    for (tr, code, name) in tpl_rows:
        for gr in sorted(geo_keys):
            if gr in used_geo:
                continue
            gcode, gname = geo_keys[gr]
            if _row_key_match(code, name, gcode, gname):
                row_map[tr] = gr
                used_geo.add(gr)
                break
    # 第二轮：名称匹配（代码误读严重的行）
    for (tr, code, name) in tpl_rows:
        if tr in row_map or not name:
            continue
        for gr in sorted(geo_keys):
            if gr in used_geo:
                continue
            gcode, gname = geo_keys[gr]
            if gname and _match_text(name, gname):
                row_map[tr] = gr
                used_geo.add(gr)
                break
    unmatched_tpl = [(tr, code, name) for (tr, code, name) in tpl_rows
                     if tr not in row_map]
    extra_geo = []
    tpl_codes = [c for (_r, c, _n) in tpl_rows]
    for gr in sorted(geo_keys):
        if gr in used_geo:
            continue
        gcode, gname = geo_keys[gr]
        # 段内表头行（科目代码/科目名称）不是数据行
        if ("科目代码" in (gcode or "") or "科目名称" in (gname or "")
                or "科目名称" in (gcode or "") or "科目代码" in (gname or "")):
            continue
        # 合计行（代码空/名称含"合计"）由邻近偏移对齐，不算"多出的行"
        if "合计" in (gname or "") or "合计" in (gcode or ""):
            continue
        if not _key_code(gcode) and not (gname or "").strip():
            continue
        if not _key_code(gcode):
            continue
        # 该代码在模板里已存在（说明是照片底部多检出的细行/重复行，
        # 文字与相邻行重复），不算"模板未包含"
        if any(_key_match(tc, gcode) for tc in tpl_codes):
            continue
        extra_geo.append((gr, gcode))
    return row_map, unmatched_tpl, extra_geo


def apply_xlsx_template(page, structure, tpl_page: XlsxSheetPage,
                        engine, doc_name: str = "",
                        cells: Optional[Dict[Cell, Tuple[int, int, int, int]]] = None,
                        fill_mode: str = "vlookup"):
    """套用 xlsx 模板页：结构区取模板文本，数字区逐格 OCR 填入。

    fill_mode：
      "vlookup"（默认，稳）——按科目代码/名称做行级 VLOOKUP 对齐后填入，
          数值只会落到"科目匹配上的那一行"；未匹配的行留空并给出明确提示，
          照片里多出的行忽略。专防"数值正确但填错行"。
      "position"——按网格位置直接填入（旧行为；用户确认后可选）。
    cells：匹配阶段给出的铺格坐标（几何优先）。缺省时用检测网格。
    返回 (warnings, mismatch)；mismatch 供界面提示用户选择。
    """
    warnings: List[str] = []
    mismatch: Dict = {}
    if cells is None:
        cells = {(c.row, c.col): (c.x, c.y, c.w, c.h)
                 for c in structure.cells}
    n_rows, n_cols = tpl_page.n_rows, tpl_page.n_cols
    out_rows: List[List[str]] = [list(r) for r in tpl_page.rows]

    H = structure.image.shape[0]
    W = structure.image.shape[1]

    # 整表 det 文本按铺格落位（避免按格裁切切穿数字）
    det_by_cell: Dict[Cell, List] = {}
    for box, text, score in (structure.det_items or []):
        if not str(text).strip():
            continue
        cx = sum(pt[0] for pt in box) / len(box)
        cy = sum(pt[1] for pt in box) / len(box)
        for (r2, c2), (x, y, w, h) in cells.items():
            if x <= cx < x + w and y <= cy < y + h:
                det_by_cell.setdefault((r2, c2), []).append(
                    (sum(pt[1] for pt in box) / len(box), cx, str(text), float(score)))
                break

    def cell_text(gr: int, gc: int):
        """几何行 gr、列 gc 的文本（det 优先，缺失时裁切识别）。"""
        got = det_by_cell.get((gr, gc))
        if got:
            got.sort(key=lambda e: (e[0], e[1]))
            return ("".join(t for _cy, _cx, t, _s in got),
                    min(s for _cy, _cx, _t, s in got))
        box = cells.get((gr, gc))
        if box is None:
            return "", 1.0
        x, y, w, h = box
        crop = structure.image[max(0, y - 2):min(H, y + h + 2),
                               max(0, x - 2):min(W, x + w + 2)]
        if crop.size == 0 or engine.is_blank(crop):
            return "", 1.0
        return engine.recognize_cell_numeric(crop)

    # ---- 行级 VLOOKUP 对齐 ----
    data_geo_rows = [r for r in range(n_rows)
                     if r >= tpl_page.header_rows and r not in tpl_page.mid_rows]
    geo_keys: Dict[int, Tuple[str, str]] = {}
    for gr in data_geo_rows:
        code = cell_text(gr, 0)[0]
        name = cell_text(gr, 1)[0]
        geo_keys[gr] = (code, name)

    if fill_mode == "vlookup" and tpl_page.n_rows:
        row_map, unmatched_tpl, extra_geo = build_row_mapping(tpl_page, geo_keys)
        # 未匹配行（合计等空代码行）：按同段邻近匹配行的偏移对齐
        mid = max(tpl_page.mid_rows) if tpl_page.mid_rows else -1

        def section_of(r: int) -> int:
            return 0 if (mid < 0 or r <= mid) else 1

        offsets_within = {}
        for tr, gr in row_map.items():
            offsets_within.setdefault(section_of(tr), []).append(gr - tr)
        for r in data_geo_rows:
            if r in row_map:
                continue
            code = str(tpl_page.rows[r][0] if tpl_page.rows[r] else "").strip()
            name = str(tpl_page.rows[r][1] if len(tpl_page.rows[r]) > 1 else "").strip()
            if _key_code(code):
                continue          # 有代码但没匹配上 → 保持未填充
            seg = offsets_within.get(section_of(r))
            if seg:
                from statistics import median as _med
                gr = r + int(round(_med(seg)))
                if 0 <= gr < n_rows:
                    row_map[r] = gr
        mismatch = {
            "mode": "vlookup",
            "matched": len(row_map),
            "unmatched": [[tr, code, name] for (tr, code, name) in unmatched_tpl],
            "extra": [[gr, gc] for (gr, gc) in extra_geo],
        }
        if unmatched_tpl:
            head = "；".join(f"第{tr + 1}行[{code} {name}]"
                            for (tr, code, name) in unmatched_tpl[:4])
            warnings.append(f"{len(unmatched_tpl)} 行科目未在照片中找到对应"
                            f"（已留空）：{head}"
                            + ("…" if len(unmatched_tpl) > 4 else ""))
        if extra_geo:
            head = "；".join(f"第{gr + 1}行[{code}]" for (gr, code) in extra_geo[:4])
            warnings.append(f"照片中有 {len(extra_geo)} 行模板未包含（已忽略）：{head}"
                            + ("…" if len(extra_geo) > 4 else ""))
    else:
        # 按位置填充（旧行为）；仍记录键不匹配的行以提示
        row_map = {r: r for r in data_geo_rows}
        if fill_mode == "position":
            _, unmatched_tpl, extra_geo = build_row_mapping(tpl_page, geo_keys)
            mismatch = {"mode": "position", "matched": len(row_map) - len(unmatched_tpl),
                        "unmatched": [[tr, code, name] for (tr, code, name) in unmatched_tpl],
                        "extra": [[gr, gc] for (gr, gc) in extra_geo]}
            if unmatched_tpl or extra_geo:
                warnings.append(
                    f"按位置强制填充：{len(unmatched_tpl)} 行科目对不上、"
                    f"{len(extra_geo)} 行照片多出，数值可能填错行，请重点核对")

    # ---- 填值：以"数字区域"为唯一依据逐格填 ----
    # 数字区域之外的格（表头/中缝表头/合计标签等冻结文本）永不覆盖。
    for (tr, tc) in sorted(tpl_page.num_cells):
        if tr >= len(out_rows) or tc >= len(out_rows[tr]):
            continue
        gr = row_map.get(tr)
        if gr is None:
            out_rows[tr][tc] = ""    # 未匹配行：留空，不按位置猜
            continue
        text, score = cell_text(gr, tc)
        out_rows[tr][tc] = text
        if text:
            page.scores[f"{tr},{tc}"] = score
            page.min_score = min(page.min_score, score)

    # 覆盖预览：数字区绿框（高亮实际取数的几何格）
    import cv2
    overlay = structure.image.copy()
    for (tr, tc) in sorted(tpl_page.num_cells):
        gr = row_map.get(tr)
        if gr is None:
            continue
        box = cells.get((gr, tc))
        if box is not None:
            x, y, w, h = box
            cv2.rectangle(overlay, (x, y), (x + w, y + h), (80, 200, 80), 2)

    page.mode = "table"
    page.borderless = False
    page.rows = out_rows
    page.merges = [[int(v) for v in m] for m in tpl_page.merges
                   if int(m[0]) < n_rows]
    page.n_rows, page.n_cols = n_rows, n_cols
    page.template = f"{doc_name}·{tpl_page.page_name}" if doc_name         else tpl_page.page_name
    page.fill_mode = fill_mode
    page.row_mismatch = mismatch
    from .service import _encode_jpeg, _scale_for
    sc = _scale_for(structure.image)
    page.cell_boxes = {
        f"{r},{c}": [int(x * sc), int(y * sc), int(w * sc), int(h * sc)]
        for (r, c), (x, y, w, h) in cells.items()}
    page.overlay_jpeg = _encode_jpeg(overlay)
    return warnings, mismatch


# ----------------------------------------------------------------------
# 勾稽：填充导出 + 公式求值
# ----------------------------------------------------------------------
def fill_template_workbook(tpl_path: Path, sheet_fills: Dict[str, object],
                           out_path: Path) -> None:
    """基于模板 xlsx 生成填充后的工作簿（公式/命名区域/格式原样保留）。

    sheet_fills: {sheet_name: page}；page 需有 .rows 与模板页尺寸一致。
    """
    wb = load_workbook(Path(tpl_path), data_only=False)
    for sheet_name, page in sheet_fills.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        rows = getattr(page, "rows", [])
        merged_children = set()
        for mr in ws.merged_cells.ranges:
            for rr in range(mr.min_row, mr.max_row + 1):
                for cc in range(mr.min_col, mr.max_col + 1):
                    if (rr, cc) != (mr.min_row, mr.min_col):
                        merged_children.add((rr, cc))
        for r, row in enumerate(rows, start=1):
            for c, v in enumerate(row, start=1):
                if (r, c) in merged_children:
                    continue          # 合并格非锚点不可写
                try:
                    cell = ws.cell(row=r, column=c)
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        continue      # 不覆盖公式
                    cell.value = (float(str(v).replace(",", ""))
                                  if _looks_numeric(v) else v)
                except Exception:
                    continue
    wb.save(Path(out_path))
    wb.close()


def capture_geometry(path: Path, sheet_name: str,
                     xs: List[int], ys: List[int]) -> bool:
    """把一次成功匹配的检测网格折算为比例，写入隐藏"几何"表。

    手搓的 xlsx 模板没有几何信息，匹配只能要求"照片网格与模板尺寸严格
    一致"；首次成功套用后自动回写几何比例，之后拍摄规格略有差异也能稳定
    套用（外框+比例+文字锚点）。文件被 Excel 占用等异常时静默返回 False。
    """
    try:
        if len(xs) < 2 or len(ys) < 2:
            return False
        path = Path(path)
        wb = load_workbook(path, data_only=False)
        if GEO_SHEET not in wb.sheetnames:
            ws = wb.create_sheet(GEO_SHEET)
            ws.sheet_state = "veryHidden"
        else:
            ws = wb[GEO_SHEET]
        try:
            geo = json.loads(str(ws["A1"].value or "{}"))
        except Exception:
            geo = {}
        w = max(1, xs[-1] - xs[0])
        h = max(1, ys[-1] - ys[0])
        geo[sheet_name] = {
            "col_fracs": [round((x - xs[0]) / w, 5) for x in xs],
            "row_fracs": [round((y - ys[0]) / h, 5) for y in ys],
        }
        ws["A1"] = json.dumps(geo, ensure_ascii=False)
        wb.save(path)
        wb.close()
        return True
    except Exception:
        return False


def _instantiate_cf(formula: str, arow: int, acol: int, r: int, c: int) -> str:
    """把条件格式公式按"锚点单元格 -> 目标单元格"实例化。

    CF 公式里的相对引用（行列号前无 $）相对于规则范围左上角，对范围里
    每一格都要换算成该格的实际坐标（如 B3:B25 的 $C3 -> 第4行时 $C4）。
    """
    def repl(m):
        colp, rowp = m.group(1), m.group(2)
        col_abs, row_abs = colp.startswith("$"), rowp.startswith("$")
        if col_abs:
            outc = colp.lstrip("$")
        else:
            outc = get_column_letter(c + (column_index_from_string(colp) - acol))
        outr = int(rowp.lstrip("$")) if row_abs else \
            r + (int(rowp.lstrip("$")) - arow)
        return ("$" if col_abs else "") + outc + str(outr)

    return re.sub(r"(\$?[A-Za-z]{1,3})(\$?\d+)", repl, formula)


def evaluate_checks(tpl_path: Path, filled_path: Path,
                    sheets: Optional[List[str]] = None) -> List[str]:
    """评估模板数据表上的**条件格式**勾稽规则（用户自建，不平标红）。

    读取每个数据工作表的 expression 型条件格式，逐格实例化公式并求值：
    命中（返回真 = 该格本该标红）的行在应用内报"勾稽不平"，与 Excel 中
    看到的红色标记一致——不用打开 Excel 也能第一时间发现。

    sheets 指定时只评估这些工作表（本次实际填充的页）；缺省评估全部
    数据表——多页模板里其它页保留的是模板原值，通常不该一起判。
    """
    out: List[str] = []
    try:
        wb = load_workbook(Path(filled_path), data_only=False)
    except Exception:
        return out

    cache = {}

    def _sumargs(*a):
        vals = a[0] if len(a) == 1 and isinstance(a[0], (list, tuple)) else a
        return float(sum(x for x in vals if isinstance(x, (int, float))))

    def _minargs(*a):
        vals = a[0] if len(a) == 1 and isinstance(a[0], (list, tuple)) else a
        nums = [x for x in vals if isinstance(x, (int, float))]
        return min(nums) if nums else 0.0

    def _maxargs(*a):
        vals = a[0] if len(a) == 1 and isinstance(a[0], (list, tuple)) else a
        nums = [x for x in vals if isinstance(x, (int, float))]
        return max(nums) if nums else 0.0

    def cell_val(sheet, coord):
        key = (sheet, coord)
        if key in cache:
            return cache[key]
        cache[key] = 0.0
        ws = wb[sheet] if sheet in wb.sheetnames else None
        if ws is None:
            return 0.0
        v = ws[coord].value
        if isinstance(v, str) and v.startswith("="):
            v = _eval_formula(v, sheet, 0)
        if isinstance(v, str):
            s = v.strip().replace(",", "")
            try:
                v = float(s) if s else 0.0
            except ValueError:
                pass
        elif v is None:
            v = 0.0
        cache[key] = v
        return v

    def _expand_sum(expr, sheet):
        def repl(m):
            arg = m.group(1).strip()
            if ":" not in arg:
                return m.group(0)
            a, b = arg.split(":", 1)
            sh = sheet
            if "!" in a:
                sh, a = a.split("!", 1)
                sh = sh.strip("'")
            if "!" in b:
                b = b.split("!", 1)[1]
            vals = []
            c0, r0, c1, r1 = range_boundaries(a + ":" + b)
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    vals.append(cell_val(sh, get_column_letter(cc) + str(rr)))
            return repr(_sumargs(*vals))
        return re.sub(r"(?i)SUM\(\s*([^()]+?)\s*\)", repl, expr)

    def _eval_formula(raw, sheet, depth):
        if depth > 8:
            return 0.0
        expr = raw.lstrip("=").strip()
        m = re.match(r"(?is)^IF\((.*)\)$", expr)
        if m:
            inner = m.group(1)
            parts, dp, cur = [], 0, ""
            for ch in inner:
                if ch == "(":
                    dp += 1
                elif ch == ")":
                    dp -= 1
                if ch == "," and dp == 0:
                    parts.append(cur)
                    cur = ""
                else:
                    cur += ch
            parts.append(cur)
            if len(parts) == 3:
                cond = _eval_formula("=" + parts[0], sheet, depth + 1)
                branch = parts[1] if cond else parts[2]
                return _eval_formula("=" + branch, sheet, depth + 1)
        expr = _expand_sum(expr, sheet)

        def ref_sub(m):
            sh = (m.group(1) or sheet).strip("'")
            return str(cell_val(sh, m.group(2).replace("$", "")))

        expr = re.sub(
            "([A-Za-z0-9_\u4e00-\u9fff']*)!\\$?([A-Z]{1,3}\\$?\\d+)",
            ref_sub, expr)
        expr = re.sub(
            "(?<![A-Za-z0-9_'\"\\)])\\$?([A-Z]{1,3})\\$?(\\d+)",
            lambda m: str(cell_val(sheet, m.group(1) + m.group(2))),
            expr)
        # 顺序关键：先 <> 再单独 =，避免把 != 变成 !==
        expr = expr.replace("<>", "!=")
        expr = re.sub(r"(?<![<>!=])=(?!=)", "==", expr)
        expr = re.sub(r"(?i)\bSUM\(", "sumargs(", expr)
        expr = re.sub(r"(?i)\bMIN\(", "minargs(", expr)
        expr = re.sub(r"(?i)\bMAX\(", "maxargs(", expr)
        expr = re.sub(r"(?i)\bABS\(", "abs(", expr)
        expr = re.sub(r"(?i)\bROUND\(", "round(", expr)
        allowed = {"abs": abs, "round": round, "sumargs": _sumargs,
                   "minargs": _minargs, "maxargs": _maxargs}
        return eval(expr, {"__builtins__": {}}, allowed)  # noqa: S307

    try:
        from openpyxl.worksheet.cell_range import CellRange
    except Exception:
        wb.close()
        return out

    for ws in wb.worksheets:
        if ws.title in (CHECK_SHEET, GEO_SHEET, BALANCE_SHEET):
            continue
        if sheets and ws.title not in sheets:
            continue
        try:
            cf_list = list(ws.conditional_formatting)
        except Exception:
            continue
        for cf in cf_list:
            for rule in cf.rules:
                try:
                    if rule.type != "expression" or not rule.formula:
                        continue
                    tpl_formula = str(rule.formula[0])
                    refs = str(cf.sqref).split()
                    if not refs:
                        continue
                    anchor = CellRange(refs[0])
                    arow, acol = anchor.min_row, anchor.min_col
                    bad = []
                    for rng in refs:
                        cr = CellRange(rng)
                        for rr in range(cr.min_row, cr.max_row + 1):
                            for cc in range(cr.min_col, cr.max_col + 1):
                                inst = _instantiate_cf(tpl_formula, arow, acol,
                                                       rr, cc)
                                try:
                                    res = _eval_formula("=" + inst, ws.title, 0)
                                except Exception:
                                    continue
                                if res is True:
                                    bad.append((rr, cc))
                    if bad:
                        labels = []
                        for (rr, _cc) in bad[:6]:
                            code = ws.cell(row=rr, column=1).value
                            name = ws.cell(row=rr, column=2).value
                            tag = " ".join(str(x) for x in (code, name) if x)
                            labels.append("第%d行" % rr
                                          + ("[%s]" % tag if tag else ""))
                        out.append("%s 勾稽不平 %d 行：%s%s（模板条件格式标红处，请核对）"
                                   % (ws.title, len(bad), "；".join(labels),
                                      "…" if len(bad) > 6 else ""))
                except Exception:
                    continue
    wb.close()
    return out

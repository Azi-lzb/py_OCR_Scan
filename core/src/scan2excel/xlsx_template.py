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
from openpyxl.utils import get_column_letter, range_boundaries
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
    for dn in wb.defined_names.values():
        nm = dn.name or ""
        refs = _parse_refs(dn.attr_text, "")
        for sheet, rect in refs:
            if not sheet:
                continue
            if _NUM_NAME.match(nm):
                num_by_sheet.setdefault(sheet, set()).update(_rect_cells(rect))
            elif _STRUCT_NAME.match(nm):
                struct_by_sheet.setdefault(sheet, set()).update(_rect_cells(rect))
            elif _MID_NAME.match(nm):
                mid_by_sheet.setdefault(sheet, set()).update(
                    range(rect[0], rect[2] + 1))

    tpl = XlsxTemplate(name=path.stem, path=path)
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
            # 手工制作的 xlsx 未写命名区域：启发式推断数字区
            hr = 0
            for r in rows[:4]:
                if any(_looks_numeric(v) for v in r):
                    break
                hr += 1
            mid_rows = set()
            for r in range(hr, len(rows)):
                row = rows[r]
                head = "".join(row)
                if "科目代码" in head and "科目名称" in head:
                    mid_rows.add(r)
            num = {(r, c) for r in range(len(rows))
                   for c in range(n_cols)
                   if r >= hr and r not in mid_rows
                   and c >= 2 and _looks_numeric(rows[r][c])}
        merges: List[List[int]] = []
        for mr in ws.merged_cells.ranges:
            c0, r0, c1, r1 = mr.min_col, mr.min_row, mr.max_col, mr.max_row
            if r1 - r0 + 1 > 1 or c1 - c0 + 1 > 1:
                merges.append([r0 - 1, c0 - 1, r1 - r0 + 1, c1 - c0 + 1])
        tpl.pages.append(XlsxSheetPage(
            page_name=ws.title, rows=rows, num_cells=num,
            struct_cells=struct, mid_rows=set(mid), merges=merges))

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
    if CHECK_SHEET not in wb.sheetnames:
        ws_check = wb.create_sheet(CHECK_SHEET)
        ws_check["A1"] = "勾稽校验（在此写 Excel 公式，如 B2 = =数据!G8-SUM(数据!G3:G7)；结果非 0 会在套用时告警）"
        ws_check["A1"].font = ws_check["A1"].font.copy(bold=True)
        ws_check.column_dimensions["A"].width = 96

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
    if mid_rows:
        add_name("中缝表头", {(r, 0) for r in mid_rows} | {(r, 1) for r in mid_rows})

    wb.save(path)
    wb.close()


# ----------------------------------------------------------------------
# 勾稽公式生成（用户口径：上期借-上期贷+本期发生借-本期发生贷 = 本期借-本期贷）
# ----------------------------------------------------------------------
def generate_balance_checks(path: Path, sheet_name: str,
                            pg: XlsxSheetPage) -> int:
    """为模板页生成逐行勾稽差额公式，写入「勾稽」表。

    差额 = (上期借-上期贷) + (本期发生借-本期发生贷) - (本期借-本期贷)，
    正常应为 0。公式全部写全表名引用，用户对该列加条件格式（<>0）即可
    可视化告警；有中缝段（收/付方）时按段分别生成。
    返回生成的公式行数。
    """
    # 兼容两种页对象：XlsxSheetPage（有 num_cells/col_labels()）与
    # template_mode.TemplatePage（有 value_cols/col_labels 列表）
    raw_labels = pg.col_labels() if callable(getattr(pg, "col_labels", None))         else list(getattr(pg, "col_labels", []) or [])
    labels = [str(x or "") for x in raw_labels]

    def _seg_num_cells(r0: int, r1: int):
        nc = getattr(pg, "num_cells", None)
        if nc:
            return {(r, c) for (r, c) in nc if r0 <= r < r1}
        cols = list(getattr(pg, "value_cols", []) or [])
        mids = getattr(pg, "mid_rows", None)
        mids = set(mids) if mids else set()
        mr = int(getattr(pg, "mid_section_row", -1) or -1)
        if mr >= 0:
            mids.add(mr)
        return {(r, c) for r in range(r0, r1) if r not in mids for c in cols}

    def find_col(include, exclude=()):
        for i, lab in enumerate(labels):
            if not lab:
                continue
            if all(k in lab for k in include) and not any(n in lab for n in exclude):
                return i
        return -1

    # 段落：[（名称, 起始行含, 结束行不含, 借方关键字, 贷方关键字）]
    segs = []
    mids = getattr(pg, "mid_rows", None)
    mids = set(mids) if mids else set()
    mr = int(getattr(pg, "mid_section_row", -1) or -1)
    if mr >= 0:
        mids.add(mr)
    if mids:
        mid_lo, mid_hi = min(mids), max(mids)   # 中缝表头可能占两行
        segs.append(("上段", pg.header_rows, mid_lo, "借", "贷"))
        segs.append(("中缝段", mid_hi + 1, pg.n_rows, "收", "付"))
    else:
        segs.append(("全表", pg.header_rows, pg.n_rows, "借", "贷"))

    q = f"'{sheet_name}'" if re.search(r"[^A-Za-z0-9_一-鿿]", sheet_name)         else sheet_name
    wb = load_workbook(Path(path), data_only=False)
    if BALANCE_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(BALANCE_SHEET)
        ws["A1"] = "模板页"
        ws["B1"] = "行"
        ws["C1"] = "科目"
        ws["D1"] = "勾稽差额（0=平；可对本列加条件格式 <>0 告警）"
        ws["E1"] = "段落"
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["C"].width = 22
        ws.column_dimensions["D"].width = 60
        ws.column_dimensions["E"].width = 12
    ws = wb[BALANCE_SHEET]
    row_out = ws.max_row + 1
    made = 0
    for seg_name, r0, r1, dkw, ckw in segs:
        open_d = find_col(["上期", dkw])
        open_c = find_col(["上期", ckw])
        per_d = find_col(["发生", dkw])
        per_c = find_col(["发生", ckw])
        close_d = find_col(["余额", dkw], exclude=("上期",))
        close_c = find_col(["余额", ckw], exclude=("上期",))
        if min(open_d, open_c, per_d, per_c, close_d, close_c) < 0:
            # 兜底：表头标签不全（组标题行未被识别）时按"数值列从左到右
            # 借/贷交替成对"推断——三对依次为 上期/发生/期末
            seg_cols = sorted({c for (r, c) in _seg_num_cells(r0, r1)})
            if len(seg_cols) >= 6:
                open_d, open_c, per_d, per_c, close_d, close_c = seg_cols[:6]
            else:
                continue      # 列数不足以配对，跳过（不硬编）
        L = get_column_letter
        for r in range(r0, r1):
            code = str(pg.rows[r][0] or pg.rows[r][1] or "").strip()
            rn = r + 1
            formula = (f"={q}!{L(open_d + 1)}{rn}-{q}!{L(open_c + 1)}{rn}"
                       f"+{q}!{L(per_d + 1)}{rn}-{q}!{L(per_c + 1)}{rn}"
                       f"-({q}!{L(close_d + 1)}{rn}-{q}!{L(close_c + 1)}{rn})")
            ws.cell(row=row_out, column=1, value=sheet_name)
            ws.cell(row=row_out, column=2, value=rn)
            ws.cell(row=row_out, column=3, value=code)
            ws.cell(row=row_out, column=4, value=formula)
            ws.cell(row=row_out, column=5, value=seg_name)
            row_out += 1
            made += 1
    wb.save(Path(path))
    wb.close()
    return made


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
                        force_doc: str = ""):
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
        note = ("尺寸不符：" + "；".join(notes[:3])) if notes else "无候选"
        return None, None, 0.0, note, {}
    if score < 0.55:
        return None, None, score, f"结构区命中率仅 {score:.0%}", {}
    return doc, pg, score, "", cells


def apply_xlsx_template(page, structure, tpl_page: XlsxSheetPage,
                        engine, doc_name: str = "",
                        cells: Optional[Dict[Cell, Tuple[int, int, int, int]]] = None):
    """套用 xlsx 模板页：结构区取模板文本，数字区逐格 OCR 填入。

    cells：匹配阶段给出的铺格坐标（几何优先）。缺省时用检测网格。
    """
    warnings: List[str] = []
    if cells is None:
        cells = {(c.row, c.col): (c.x, c.y, c.w, c.h)
                 for c in structure.cells}
    n_rows, n_cols = tpl_page.n_rows, tpl_page.n_cols
    out_rows: List[List[str]] = [list(r) for r in tpl_page.rows]

    H = structure.image.shape[0]
    W = structure.image.shape[1]

    # 数字填充优先用"整表 det 的整体文本按中心落格"——det 把完整数字识别
    # 成一条，避免按格裁切把数字切穿（如 1,983,687,343.16 → "…343." + "16"）；
    # 该格没有 det 文本时才退回按格裁切识别
    det_by_cell = {}
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

    for r in range(n_rows):
        for c in range(n_cols):
            if (r, c) not in tpl_page.num_cells:
                continue
            got = det_by_cell.get((r, c))
            if got:
                got.sort(key=lambda e: (e[0], e[1]))
                text = "".join(t for _cy, _cx, t, _s in got)
                score = min(s for _cy, _cx, _t, s in got)
                out_rows[r][c] = text
                if text:
                    page.scores[f"{r},{c}"] = score
                    page.min_score = min(page.min_score, score)
                continue
            box = cells.get((r, c))
            if box is None:
                continue
            x, y, w, h = box
            crop = structure.image[max(0, y - 2):min(H, y + h + 2),
                                   max(0, x - 2):min(W, x + w + 2)]
            if crop.size == 0:
                continue
            if engine.is_blank(crop):
                out_rows[r][c] = ""
                continue
            text, score = engine.recognize_cell_numeric(crop)
            out_rows[r][c] = text
            if text:
                page.scores[f"{r},{c}"] = score
                page.min_score = min(page.min_score, score)

    import cv2
    overlay = structure.image.copy()
    for (r, c) in sorted(tpl_page.num_cells):
        box = cells.get((r, c))
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
    # 铺格坐标转预览坐标（界面联动/高亮用）
    from .service import _encode_jpeg, _scale_for
    sc = _scale_for(structure.image)
    page.cell_boxes = {
        f"{r},{c}": [int(x * sc), int(y * sc), int(w * sc), int(h * sc)]
        for (r, c), (x, y, w, h) in cells.items()}
    page.overlay_jpeg = _encode_jpeg(overlay)
    return warnings


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


def evaluate_checks(tpl_path: Path, filled_path: Path) -> List[str]:
    """在填充后的工作簿上求值校验表公式；结果非 0/非空/False 即告警。

    轻量求值器：支持 单元格引用(可跨表)、SUM/ABS/MIN/MAX/ROUND、四则、
    比较运算与 IF；不支持的公式跳过并提示在 Excel 中查看。
    """
    out: List[str] = []
    try:
        wb = load_workbook(Path(filled_path), data_only=False)
    except Exception:
        return out
    if CHECK_SHEET not in wb.sheetnames:
        wb.close()
        return out

    cache: Dict[Tuple[str, str], object] = {}

    def cell_val(sheet: str, coord: str):
        key = (sheet, coord)
        if key in cache:
            return cache[key]
        cache[key] = 0.0
        ws = wb[sheet] if sheet in wb.sheetnames else None
        if ws is None:
            return 0.0
        v = ws[coord].value
        if isinstance(v, str) and v.startswith("="):
            v = _eval_formula(v, sheet)
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

    def _refs(expr: str, sheet: str, depth: int):
        # 展开 A1 与 SUM(A1:B2) 里的引用为值
        def sum_range(ab):
            a, b = ab
            sh = sheet
            if "!" in a:
                sh, a = a.split("!", 1)
                sh = sh.strip("'")
            if "!" in b:
                b = b.split("!", 1)[1]
            vals = []
            c0, r0, c1, r1 = range_boundaries(f"{a}:{b}")
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    vals.append(cell_val(sh, f"{get_column_letter(cc)}{rr}"))
            return str(sum(v for v in vals if isinstance(v, (int, float))))

        expr = re.sub(r"(?i)SUM\(\s*([^()]+?)\s*\)",
                      lambda m: sum_range(_split_sum_args(m.group(1))), expr)
        def ref_sub(m):
            sh = m.group(1) or sheet
            return str(cell_val(sh.strip("'"), m.group(2)))
        # 带表名前缀的引用
        expr = re.sub(r"([A-Za-z0-9_\u4e00-\u9fff']+)!\$?([A-Z]{1,3}\$?\d+)",
                      ref_sub, expr)
        # 裸引用（同表）
        expr = re.sub(r"(?<![A-Za-z0-9_'\"])\$?([A-Z]{1,3})\$?(\d+)",
                      lambda m: str(cell_val(sheet, m.group(1) + m.group(2))),
                      expr)
        return expr

    def _split_sum_args(arg: str):
        parts = arg.split(":")
        if len(parts) == 2:
            return parts[0].strip(), parts[1].strip()
        return arg, arg

    def _eval_formula(raw: str, sheet: str, depth: int = 0):
        if depth > 6:
            return 0.0
        expr = raw.lstrip("=").strip()
        # IF(cond,a,b) → (a if cond else b)
        m = re.match(r"(?is)^IF\((.*)\)$", expr)
        if m:
            inner = m.group(1)
            parts, depth_par, cur = [], 0, ""
            for ch in inner:
                if ch == "(":
                    depth_par += 1
                elif ch == ")":
                    depth_par -= 1
                if ch == "," and depth_par == 0:
                    parts.append(cur); cur = ""
                else:
                    cur += ch
            parts.append(cur)
            if len(parts) == 3:
                cond = _eval_formula("=" + parts[0], sheet, depth + 1)
                branch = parts[1] if cond else parts[2]
                return _eval_formula("=" + branch, sheet, depth + 1)
        expr = _refs(expr, sheet, depth)
        expr = expr.replace("<>", "!=").replace("=", "==")
        expr = re.sub(r"(?<![<>!])(?<![=!])=(?!=)", "==", expr)
        expr = re.sub(r"\bABS\(", "abs((", expr)
        # 补右括号由 eval 报错兜底（简单公式足够）
        if not re.fullmatch(r"[-+*/%().,\s\d.:'\"A-Za-z_!=\u4e00-\u9fff]*", expr):
            raise ValueError("unsupported")
        allowed = {"abs": abs, "min": min, "max": max, "round": round}
        return eval(expr.replace("\\", ""), {"__builtins__": {}}, allowed)  # noqa: S307

    ws_check = wb[CHECK_SHEET]
    for row in ws_check.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            label = str(ws_check.cell(row=cell.row, column=1).value or
                        f"公式 {cell.coordinate}").strip()
            try:
                res = _eval_formula(v, CHECK_SHEET)
            except Exception:
                out.append(f"{label}：公式较复杂，请在 Excel 中查看（{v}）")
                continue
            bad = False
            if isinstance(res, str):
                bad = bool(res.strip()) and res.strip() not in ("True", "OK")
            elif isinstance(res, (int, float)):
                bad = abs(float(res)) > 1e-6
            elif isinstance(res, bool):
                bad = not res
            if bad:
                out.append(f"{label}：勾稽不符（{v} → {res}）")

    # 逐行勾稽表：统计不平行（差额 = 上期借-贷 + 本期发生借-贷 - 期末借-贷）
    if BALANCE_SHEET in wb.sheetnames:
        ws_bal = wb[BALANCE_SHEET]
        bad_rows = []
        for row in ws_bal.iter_rows(min_row=2):
            formula = row[3].value if len(row) > 3 else None
            if not (isinstance(formula, str) and formula.startswith("=")):
                continue
            try:
                res = _eval_formula(formula, BALANCE_SHEET)
            except Exception:
                continue
            if isinstance(res, (int, float)) and abs(float(res)) > 0.01:
                code = str(row[2].value if len(row) > 2 else "")
                bad_rows.append((code, float(res)))
        if bad_rows:
            head = "；".join(f"{c or '合计/空'} 差{abs(v):,.2f}"
                             for c, v in bad_rows[:4])
            out.append(f"逐行勾稽不平 {len(bad_rows)} 行：{head}"
                       + ("…" if len(bad_rows) > 4 else "")
                       + "（见「勾稽」表，可对差额列加条件格式）")
    wb.close()
    return out

# -*- coding: utf-8 -*-
"""月计表专用模板识别。

固定格式表格（如银行会计月计表）：行列结构、科目代码/名称每月不变，
只有数值变化。模板把整页的行列结构"写死"：
  - 页面几何：列边界与行边界（相对表格宽高的比例，由模板页采样）
  - 各行文本：科目代码、科目名称、表头文字全部冻结，不受 OCR 错误影响
  - 数值列：每月实际 OCR 的格子（借方/贷方 × 上期余额/本期发生额/本期余额…）
识别时只用检测到的表格外边界（最稳的线）做定位，内部按模板比例铺格，
再用科目代码列做锚定校验（容错匹配，只报警不覆盖模板文本）。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import cv2
import numpy as np

_NUM_RE = re.compile(r"^[-+]?[\d,]+(\.\d+)?%?$")
_DIGITS = re.compile(r"\d+")


def _is_numeric_text(t: str) -> bool:
    t = (t or "").strip()
    return bool(t) and bool(_NUM_RE.match(t))


def _norm_code(t: str) -> str:
    return "".join(ch for ch in (t or "") if ch.isdigit() or ch.isalpha())


def _codes_match(a: str, b: str) -> bool:
    """科目代码模糊匹配：容忍零星 OCR 误读（数字 6/9、缺字等）。"""
    a, b = _norm_code(a), _norm_code(b)
    if not a or not b:
        return a == b
    if a == b:
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    # 逐位比较取相同位置吻合率
    n = min(len(a), len(b))
    same = sum(1 for i in range(n) if a[i] == b[i])
    return same / max(len(a), len(b)) >= 0.7


@dataclass
class TemplatePage:
    """模板文档中的一页（一种版式）：行列几何 + 冻结文本 + 数值列。"""

    page_name: str = "第1页"
    rows: List[List[str]] = field(default_factory=list)
    col_fracs: List[float] = field(default_factory=list)
    row_fracs: List[float] = field(default_factory=list)
    value_cols: List[int] = field(default_factory=list)
    code_col: int = 0
    header_rows: int = 0

    @property
    def n_cols(self) -> int:
        return max(0, len(self.col_fracs) - 1)

    @property
    def n_rows(self) -> int:
        return len(self.rows)

    def to_dict(self) -> Dict:
        return {"page_name": self.page_name, "rows": self.rows,
                "col_fracs": [round(f, 5) for f in self.col_fracs],
                "row_fracs": [round(f, 5) for f in self.row_fracs],
                "value_cols": self.value_cols, "code_col": self.code_col,
                "header_rows": self.header_rows}

    @staticmethod
    def from_dict(d: Dict) -> "TemplatePage":
        return TemplatePage(
            page_name=str(d.get("page_name", "第1页")),
            rows=[list(r) for r in d.get("rows", [])],
            col_fracs=[float(f) for f in d.get("col_fracs", [])],
            row_fracs=[float(f) for f in d.get("row_fracs", [])],
            value_cols=[int(c) for c in d.get("value_cols", [])],
            code_col=int(d.get("code_col", 0)),
            header_rows=int(d.get("header_rows", 0)),
        )

    def codes(self) -> List[str]:
        """锚定匹配用的代码列（数据区）。"""
        if self.code_col >= self.n_cols:
            return []
        vals = []
        for r in self.rows[self.header_rows:]:
            v = (r[self.code_col] if self.code_col < len(r) else "") or ""
            if _DIGITS.search(v):
                vals.append(v)
        return vals


@dataclass
class TableTemplate:
    """模板文档：一个固定格式报表（可含多页，各页版式不同）。"""

    name: str
    pages: List[TemplatePage] = field(default_factory=list)

    # ---- 兼容旧单页字段的快捷访问（指向首页） ----
    @property
    def rows(self) -> List[List[str]]:
        return self.pages[0].rows if self.pages else []

    @property
    def col_fracs(self) -> List[float]:
        return self.pages[0].col_fracs if self.pages else []

    @property
    def row_fracs(self) -> List[float]:
        return self.pages[0].row_fracs if self.pages else []

    @property
    def value_cols(self) -> List[int]:
        return self.pages[0].value_cols if self.pages else []

    @property
    def header_rows(self) -> int:
        return self.pages[0].header_rows if self.pages else 0

    @property
    def code_col(self) -> int:
        return self.pages[0].code_col if self.pages else 0

    @property
    def n_rows(self) -> int:
        return self.pages[0].n_rows if self.pages else 0

    @property
    def n_cols(self) -> int:
        return self.pages[0].n_cols if self.pages else 0

    # ------------------------------------------------------------------ #
    def to_dict(self) -> Dict:
        return {"name": self.name,
                "pages": [pg.to_dict() for pg in self.pages]}

    @staticmethod
    def from_dict(d: Dict) -> "TableTemplate":
        if d.get("pages"):
            return TableTemplate(name=str(d["name"]),
                                 pages=[TemplatePage.from_dict(pg)
                                        for pg in d["pages"]])
        # 旧版单页格式：整份文档即一页
        return TableTemplate(name=str(d["name"]),
                             pages=[TemplatePage.from_dict(d)])

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=1),
                        encoding="utf-8")

    @staticmethod
    def load(path: Path) -> "TableTemplate":
        return TableTemplate.from_dict(
            json.loads(Path(path).read_text(encoding="utf-8")))


# ---------------------------------------------------------------------- #
def build_page_from_result(page_name: str, rows: List[List[str]],
                           xs: List[int], ys: List[int]) -> TemplatePage:
    """从一页已校对好的识别结果生成模板页。

    rows 为整页文本；xs/ys 为该页的列/行边界像素坐标（表格内）。
    """
    n_cols = max(len(r) for r in rows)
    rows = [list(r) + [""] * (n_cols - len(r)) for r in rows]
    w = max(1, xs[-1] - xs[0])
    h = max(1, ys[-1] - ys[0])
    col_fracs = [(x - xs[0]) / w for x in xs]
    row_fracs = [(y - ys[0]) / h for y in ys]

    code_values = [(r[0] or "").strip() for r in rows]

    def is_dup_code_col(c: int) -> bool:
        vals = [(i, (r[c] or "").strip()) for i, r in enumerate(rows)
                if (r[c] or "").strip()]
        if not vals:
            return False
        dup = sum(1 for i, v in vals
                  if code_values[i] and _codes_match(v, code_values[i]))
        return dup / len(vals) >= 0.7

    header_rows = 0
    for r in rows:
        code = _norm_code((r[0] or ""))
        if code and not code.isdigit():
            header_rows += 1
        else:
            break

    def looks_texty(v: str) -> bool:
        if _is_numeric_text(v):
            return False
        cjk = sum(1 for ch in v if '一' <= ch <= '鿿')
        return cjk >= 2 and cjk / max(1, len(v)) >= 0.5

    value_cols: List[int] = []
    for c in range(1, n_cols):
        if is_dup_code_col(c):
            continue
        data_vals = [(r[c] or "").strip() for r in rows[header_rows:]
                     if (r[c] or "").strip()]
        if data_vals and sum(1 for v in data_vals if looks_texty(v)) / len(data_vals) >= 0.6:
            continue
        value_cols.append(c)

    return TemplatePage(page_name=page_name, rows=rows, col_fracs=col_fracs,
                        row_fracs=row_fracs, value_cols=value_cols,
                        code_col=0, header_rows=header_rows)


# ---------------------------------------------------------------------- #
def match_template(structure, templates: List[TableTemplate],
                   force_doc: str = "") -> Tuple[Optional[TableTemplate], Optional[TemplatePage], float]:
    """按"模板页科目代码在 det 文本中的命中率"在文档×页两级里选最佳。

    force_doc 指定后只在本文档的各页里选（页级路由）。
    返回 (文档, 页, 命中率)；命中率 < 0.5 视为未匹配。
    """
    texts = set()
    if structure.det_items:
        for _b, t, _s in structure.det_items:
            texts.add(_norm_code(str(t)))
    if not texts or not templates:
        return None, None, 0.0
    best_doc = best_page = None
    best_score = 0.0
    for tpl in templates:
        if force_doc and tpl.name != force_doc:
            continue
        for pg in tpl.pages:
            codes = pg.codes()
            if not codes:
                continue
            hit = sum(1 for c in codes if any(_codes_match(c, t) for t in texts))
            score = hit / len(codes)
            if score > best_score:
                best_doc, best_page, best_score = tpl, pg, score
    if best_score < 0.5:
        return None, None, best_score
    return best_doc, best_page, best_score


# ---------------------------------------------------------------------- #
def apply_template(page, structure, tpl_page: TemplatePage, engine,
                   doc_name: str = "") -> List[str]:
    """按模板页识别一页：只 OCR 数值列 + 代码列校验，其余文本用模板冻结值。

    返回警告列表（代码校验不符等）；page.rows 已填好（模板文本 + 数值）。
    """
    warnings: List[str] = []
    xs, ys = structure.xs, structure.ys
    if len(xs) < 2 or len(ys) < 2:
        raise ValueError("网格边界不足，无法套用模板")
    left, right = xs[0], xs[-1]
    top, bottom = ys[0], ys[-1]
    W = structure.image.shape[1]
    H = structure.image.shape[0]

    def bx(frac: float) -> int:
        return int(round(left + frac * (right - left)))

    def by(frac: float) -> int:
        return int(round(top + frac * (bottom - top)))

    col_x = [bx(f) for f in tpl_page.col_fracs]
    row_y = [by(f) for f in tpl_page.row_fracs]

    n_rows = tpl_page.n_rows
    n_cols = tpl_page.n_cols
    out_rows: List[List[str]] = [list(r) for r in tpl_page.rows]

    for r in range(n_rows):
        if r < tpl_page.header_rows:
            continue
        for c in tpl_page.value_cols:
            if c >= n_cols or r >= len(out_rows):
                continue
            y0, y1 = row_y[r], row_y[r + 1]
            x0, x1 = col_x[c], col_x[c + 1]
            if x1 - x0 < 4 or y1 - y0 < 4:
                continue
            crop = structure.image[max(0, y0 - 2):min(H, y1 + 2),
                                   max(0, x0 - 2):min(W, x1 + 2)]
            if engine.is_blank(crop):
                out_rows[r][c] = ""
                continue
            text, score = engine.recognize_cell_numeric(crop)
            out_rows[r][c] = text
            if text:
                page.scores[f"{r},{c}"] = score
                page.min_score = min(page.min_score, score)
        if tpl_page.code_col < n_cols:
            y0, y1 = row_y[r], row_y[r + 1]
            x0, x1 = col_x[tpl_page.code_col], col_x[tpl_page.code_col + 1]
            expect = (tpl_page.rows[r][tpl_page.code_col] or "").strip()
            if expect and x1 - x0 > 4 and y1 - y0 > 4:
                crop = structure.image[max(0, y0 - 2):min(H, y1 + 2),
                                       max(0, x0 - 2):min(W, x1 + 2)]
                if not engine.is_blank(crop):
                    got, _sc = engine.recognize_cell(crop)
                    if got and not _codes_match(got, expect):
                        warnings.append(f"第{r + 1}行代码：模板[{expect}] 疑似[{got}]")

    overlay = structure.image.copy()
    for x in col_x:
        cv2.line(overlay, (x, max(0, top - 4)), (x, min(H - 1, bottom + 4)),
                 (80, 200, 80), 2)
    for y in row_y:
        cv2.line(overlay, (max(0, left - 4), y), (min(W - 1, right + 4), y),
                 (80, 200, 80), 2)

    page.mode = "table"
    page.borderless = False
    page.rows = out_rows
    page.merges = []
    page.n_rows, page.n_cols = n_rows, n_cols
    page.template = (f"{doc_name}·{tpl_page.page_name}" if doc_name
                     else tpl_page.page_name)
    page.overlay_jpeg = _encode(overlay)
    return warnings


def _encode(img: np.ndarray, max_side: int = 1100, quality: int = 80) -> bytes:
    h, w = img.shape[:2]
    if max(h, w) > max_side:
        sc = max_side / max(h, w)
        img = cv2.resize(img, (int(w * sc), int(h * sc)), interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, quality])
    return buf.tobytes() if ok else b""

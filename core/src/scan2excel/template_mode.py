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
class TableTemplate:
    """一页固定格式表格的模板。"""

    name: str
    rows: List[List[str]]                 # 整页文本（表头+数据行，全部冻结）
    col_fracs: List[float]                # n_cols+1 个列边界（0~1）
    row_fracs: List[float]                # n_rows+1 个行边界（0~1）
    value_cols: List[int] = field(default_factory=list)   # 需要逐月 OCR 的列
    code_col: int = 0                     # 锚定校验用的科目代码列
    header_rows: int = 0                  # 前 N 行视为表头（整行冻结）

    # ------------------------------------------------------------------ #
    @property
    def n_cols(self) -> int:
        return len(self.col_fracs) - 1

    @property
    def n_rows(self) -> int:
        return len(self.rows)

    def to_dict(self) -> Dict:
        return {"name": self.name, "rows": self.rows,
                "col_fracs": [round(f, 5) for f in self.col_fracs],
                "row_fracs": [round(f, 5) for f in self.row_fracs],
                "value_cols": self.value_cols, "code_col": self.code_col,
                "header_rows": self.header_rows}

    @staticmethod
    def from_dict(d: Dict) -> "TableTemplate":
        return TableTemplate(
            name=str(d["name"]), rows=[list(r) for r in d["rows"]],
            col_fracs=[float(f) for f in d["col_fracs"]],
            row_fracs=[float(f) for f in d["row_fracs"]],
            value_cols=[int(c) for c in d.get("value_cols", [])],
            code_col=int(d.get("code_col", 0)),
            header_rows=int(d.get("header_rows", 0)),
        )

    def save(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=1),
                        encoding="utf-8")

    @staticmethod
    def load(path: Path) -> "TableTemplate":
        return TableTemplate.from_dict(
            json.loads(Path(path).read_text(encoding="utf-8")))


# ---------------------------------------------------------------------- #
def build_template_from_page(name: str, rows: List[List[str]],
                             xs: List[int], ys: List[int],
                             ys_page: Optional[List[int]] = None) -> TableTemplate:
    """从一页已校对好的识别结果生成模板。

    rows 为整页文本；xs/ys 为该页的列/行边界像素坐标（表格内）。
    数值列自动判定：跳过代码列后，列内非空值 ≥70% 为数字的列。
    """
    n_cols = max(len(r) for r in rows)
    rows = [list(r) + [""] * (n_cols - len(r)) for r in rows]
    w = max(1, xs[-1] - xs[0])
    h = max(1, ys[-1] - ys[0])
    col_fracs = [(x - xs[0]) / w for x in xs]
    row_fracs = [(y - ys[0]) / h for y in ys]
    # 数值列判定（排除法，比"多数为数字"稳健——借方/贷方列本月可能
    # 多为空，稀疏数字列也必须纳入）：
    #   排除 ①代码列 ②与代码列重复的列（左右双代码）③数据区以中文文本
    #   为主的列（科目名称）；其余（含空列）全部作为数值列逐月 OCR。
    code_values = [(r[0] or "").strip() for r in rows]

    def is_dup_code_col(c: int) -> bool:
        vals = [(i, (r[c] or "").strip()) for i, r in enumerate(rows)
                if (r[c] or "").strip()]
        if not vals:
            return False
        dup = sum(1 for i, v in vals
                  if code_values[i] and _codes_match(v, code_values[i]))
        return dup / len(vals) >= 0.7

    # 表头行：开头连续、科目代码列为非数字的行（有数值的行必属数据区）
    header_rows = 0
    for r in rows:
        code = _norm_code((r[0] or ""))
        if code and not code.isdigit():
            header_rows += 1
        else:
            break

    def looks_texty(v: str) -> bool:
        """中文文本为主的单元格（科目名称类）；纯数字/代号不算。"""
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
            continue                       # 科目名称类文本列，冻结不 OCR
        value_cols.append(c)

    return TableTemplate(name=name, rows=rows, col_fracs=col_fracs,
                         row_fracs=row_fracs, value_cols=value_cols,
                         code_col=0, header_rows=header_rows)


# ---------------------------------------------------------------------- #
def match_template(structure, templates: List[TableTemplate]) -> Tuple[Optional[TableTemplate], float]:
    """按"模板科目代码在 det 文本中的命中率"挑选最匹配的模板。

    返回 (最佳模板, 命中率)；命中率 < 0.5 时视为未匹配（返回 None）。
    """
    texts = set()
    if structure.det_items:
        for _b, t, _s in structure.det_items:
            texts.add(_norm_code(str(t)))
    if not texts or not templates:
        return None, 0.0
    best, best_score = None, 0.0
    for tpl in templates:
        codes = [r[tpl.code_col] for r in tpl.rows
                 if tpl.code_col < len(r) and _DIGITS.search(r[tpl.code_col] or "")]
        codes = codes[tpl.header_rows:] if len(codes) > tpl.header_rows else codes
        if not codes:
            continue
        hit = sum(1 for c in codes
                  if any(_codes_match(c, t) for t in texts))
        score = hit / len(codes)
        if score > best_score:
            best, best_score = tpl, score
    if best_score < 0.5:
        return None, best_score
    return best, best_score


# ---------------------------------------------------------------------- #
def apply_template(page, structure, tpl: TableTemplate, engine,
                   value_min_score: float = 0.0) -> List[str]:
    """按模板识别一页：只 OCR 数值列 + 代码列校验，其余文本用模板冻结值。

    返回警告列表（代码校验不符等），页面 rows 已填好（模板文本 + 数值）。
    """
    warnings: List[str] = []
    xs, ys = structure.xs, structure.ys
    if len(xs) < 2 or len(ys) < 2:
        raise ValueError("网格边界不足，无法套用模板")
    left, right = xs[0], xs[-1]
    top, bottom = ys[0], ys[-1]
    W = img_w = structure.image.shape[1]
    H = structure.image.shape[0]

    def bx(frac: float) -> int:
        return int(round(left + frac * (right - left)))

    def by(frac: float) -> int:
        return int(round(top + frac * (bottom - top)))

    col_x = [bx(f) for f in tpl.col_fracs]
    row_y = [by(f) for f in tpl.row_fracs]

    n_rows = tpl.n_rows
    out_rows: List[List[str]] = []
    for r in range(n_rows):
        out_rows.append(list(tpl.rows[r]))

    # 数值列 OCR
    n_cols = tpl.n_cols
    for r in range(n_rows):
        if r < tpl.header_rows:
            continue                      # 表头整行冻结
        for c in tpl.value_cols:
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
        # 代码列校验（不改文本，只报警）
        if tpl.code_col < n_cols:
            y0, y1 = row_y[r], row_y[r + 1]
            x0, x1 = col_x[tpl.code_col], col_x[tpl.code_col + 1]
            expect = (tpl.rows[r][tpl.code_col] or "").strip()
            if expect and x1 - x0 > 4 and y1 - y0 > 4:
                crop = structure.image[max(0, y0 - 2):min(H, y1 + 2),
                                       max(0, x0 - 2):min(W, x1 + 2)]
                if not engine.is_blank(crop):
                    got, _sc = engine.recognize_cell(crop)
                    if got and not _codes_match(got, expect):
                        warnings.append(f"第{r + 1}行代码：模板[{expect}] 疑似[{got}]")

    # 预览叠加：画出模板网格（便于肉眼确认套用是否准确）
    overlay = structure.image.copy()
    for x in col_x:
        cv2.line(overlay, (x, max(0, top - 4)), (x, min(H - 1, bottom + 4)),
                 (80, 200, 80), 2)
    for y in row_y:
        cv2.line(overlay, (max(0, left - 4), y), (min(W - 1, right + 4), y),
                 (80, 200, 80), 2)

    # 供界面"点击单元格 ⇄ 图片高亮"联动（预览图坐标）
    sc_preview = min(1.0, 1100.0 / max(W, H))
    for r in range(n_rows):
        for c in range(n_cols):
            if c < len(col_x) - 1 and r < len(row_y) - 1:
                page.cell_boxes[f"{r},{c}"] = [
                    int(col_x[c] * sc_preview), int(row_y[r] * sc_preview),
                    int((col_x[c + 1] - col_x[c]) * sc_preview),
                    int((row_y[r + 1] - row_y[r]) * sc_preview)]

    page.mode = "table"
    page.borderless = False
    page.rows = out_rows
    page.merges = []
    page.n_rows, page.n_cols = n_rows, n_cols
    page.template = tpl.name
    page.overlay_jpeg = _encode(overlay)
    return warnings


def _encode(img: np.ndarray, max_side: int = 1100, quality: int = 80) -> bytes:
    h, w = img.shape[:2]
    if max(h, w) > max_side:
        sc = max_side / max(h, w)
        img = cv2.resize(img, (int(w * sc), int(h * sc)), interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, quality])
    return buf.tobytes() if ok else b""

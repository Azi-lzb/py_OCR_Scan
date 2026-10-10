# -*- coding: utf-8 -*-
"""业务编排：一张照片 → 表格结构 → 逐格 OCR → 行列数据。

纯业务类，不依赖任何 UI，可被 WebApi 或 CLI 直接调用。
"""
from __future__ import annotations

import os
import re
import threading
import time
from pathlib import Path
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

import cv2
import numpy as np

from .ocr_engine import OcrEngine
from .scan_doc import enhance_scan, flatten_page
from .table_extractor import (ImageReadError, TableStructure, extract_table,
                              imread_unicode, preprocess)

StepCallback = Callable[[str], None]

# 仅把全角数字/字母转半角（OCR 常见把数字字母识别成全角形态）。
# 全角标点（，：；（）等）是中文文本的正常用法，保留不转。
_FULLWIDTH_ALNUM = {}
for _i in range(10):
    _FULLWIDTH_ALNUM[chr(0xFF10 + _i)] = chr(0x30 + _i)      # ０-９
for _i in range(26):
    _FULLWIDTH_ALNUM[chr(0xFF21 + _i)] = chr(0x41 + _i)      # Ａ-Ｚ
    _FULLWIDTH_ALNUM[chr(0xFF41 + _i)] = chr(0x61 + _i)      # ａ-ｚ

# 中文语境的半角标点转全角：OCR 对拍照图常输出半角形态（扫描图输出
# 全角），统一成中文排版习惯；前后均为拉丁/数字（如 "1,200" "A, B"）
# 时保留半角。
_HALF_PUNC_TO_FULL = {",": "，", ":": "：", ";": "；", "!": "！", "?": "？",
                      "(": "（", ")": "）"}


def _is_cjk(ch: str) -> bool:
    return bool(ch) and ("\u4e00" <= ch <= "\u9fff" or ch in "，。；：！？（）")


def _nearest_non_space(chars, i: int, step: int) -> str:
    j = i + step
    while 0 <= j < len(chars) and chars[j] in " \u3000":
        j += step
    return chars[j] if 0 <= j < len(chars) else ""


def normalize_text(t: str) -> str:
    if not t:
        return t or ""
    chars = [_FULLWIDTH_ALNUM.get(c, c) for c in t]
    out = []
    for i, c in enumerate(chars):
        if c in _HALF_PUNC_TO_FULL:
            prev = _nearest_non_space(chars, i, -1)
            nxt = _nearest_non_space(chars, i, 1)
            if _is_cjk(prev) or _is_cjk(nxt):
                c = _HALF_PUNC_TO_FULL[c]
        out.append(c)
    text = "".join(out)
    # 全角标点旁不应有空格（OCR 分段拼接常残留，如 '明细 （2026）'）
    text = re.sub(r"\s+([，。；：！？）】》])", r"\1", text)
    text = re.sub(r"\s+([（【《])", r"\1", text)
    text = re.sub(r"([（【《])\s+", r"\1", text)
    return text


@dataclass
class TablePage:
    """一张照片的完整识别结果。"""
    path: str
    name: str
    mode: str = "table"                  # table=表格 | text=整页文字
    title: str = ""                      # 表格上方的大字标题（可能为空）
    rows: List[List[str]] = field(default_factory=list)
    merges: List[List[int]] = field(default_factory=list)  # [r, c, rspan, cspan]
    scores: Dict[str, float] = field(default_factory=dict)  # {"r,c": 置信度}
    # {"r,c": [x,y,w,h]}：单元格在预览图（最长边 1100 的 JPEG）上的像素框，
    # 供界面做"点击单元格 ↔ 高亮图中区域"联动
    cell_boxes: Dict[str, List[int]] = field(default_factory=dict)
    n_rows: int = 0
    n_cols: int = 0
    elapsed: float = 0.0
    error: Optional[str] = None
    warped: bool = False
    borderless: bool = False        # 无框线表格（模型结构识别）
    template: str = ""              # 命中的月计表模板名（模板模式）
    template_file: str = ""         # xlsx 模板文件路径（导出时基于它填值）
    template_page: str = ""         # 命中的模板工作表名
    fill_mode: str = "vlookup"      # 填充模式：vlookup=按科目对齐 / position=按位置
    geometry_saved: bool = False    # 本次是否为新模板回写了几何比例
    checks: List[str] = field(default_factory=list)   # 勾稽不平提示（条件格式求值）
    check_rows: List[int] = field(default_factory=list)  # 勾稽不平的页行号（0 基，界面整行标红）
    header_rows: int = 1            # 表头行数（模板页来自模板；普通页=1）
    row_mismatch: Dict = field(default_factory=dict)  # 行对齐结果（供界面提示）
    warning: str = ""               # 模板校验等提示信息
    xs: List[int] = field(default_factory=list)   # 列/行边界（有框线模式的网格）
    ys: List[int] = field(default_factory=list)
    min_score: float = 1.0          # 全表最低识别置信度，供界面提示
    preview_jpeg: Optional[bytes] = None   # 原图预览（校正后）
    overlay_jpeg: Optional[bytes] = None   # 框线叠加预览

    def to_dict(self) -> Dict:
        return {
            "path": self.path, "name": self.name, "mode": self.mode,
            "title": self.title,
            "rows": self.rows, "merges": self.merges,
            "scores": {k: round(v, 3) for k, v in self.scores.items()},
            "cell_boxes": self.cell_boxes,
            "n_rows": self.n_rows, "n_cols": self.n_cols,
            "elapsed": round(self.elapsed, 2), "error": self.error,
            "warped": self.warped, "borderless": self.borderless,
            "template": self.template, "template_page": self.template_page,
            "fill_mode": self.fill_mode, "row_mismatch": self.row_mismatch,
            "checks": self.checks, "header_rows": self.header_rows,
            "warning": self.warning,
            "min_score": round(self.min_score, 3),
        }


class Scan2ExcelService:
    """识别管线。cancel_event 由调用方持有，worker 内周期检查。"""

    def __init__(self) -> None:
        self._server_rec = False   # 高精度档（由 process_image 每次设置）

    def _engine(self):
        return OcrEngine.instance(server_rec=self._server_rec)

    def process_image(self, path: str, on_step: Optional[StepCallback] = None,
                      cancel_event: Optional[threading.Event] = None,
                      force_text: bool = False,
                      server_rec: bool = False,
                      ignore_regions: Optional[List[Dict]] = None,
                      templates: Optional[List] = None,
                      template_name: str = "",
                      template_page: str = "",
                      fill_mode: str = "vlookup",
                      auto_rotate: bool = True) -> TablePage:
        name = os.path.splitext(os.path.basename(path))[0]
        page = TablePage(path=path, name=name)
        self._server_rec = server_rec
        regions = ignore_regions or []
        start = time.time()
        try:
            self._step(on_step, f"读取图片：{os.path.basename(path)}")
            img = imread_unicode(path)   # 失败抛 ImageReadError，真实原因随异常上报

            # 方向纠正可按设置关闭（照片已摆正时省去每张数秒的方向探测）
            if auto_rotate and not force_text:
                img = self._fix_page_rotation(img, on_step)

            if force_text:
                # 用户显式点"识别文字"：无论有无表格线都按整页文字处理
                self._step(on_step, "整页文字识别 ...")
                self._process_text(page, img, on_step, cancel_event, regions)
            else:
                self._step(on_step, "透视校正 / 检测表格线")
                structure: Optional[TableStructure] = extract_table(
                    img, engine=self._engine())
                if structure is None:
                    # 没有表格框线：先试无框线表格模型，不行再转整页文字
                    if not self._process_borderless(page, img, on_step, cancel_event, regions):
                        self._step(on_step, "按整页文字识别")
                        self._process_text(page, img, on_step, cancel_event, regions)
                else:
                    page.mode = "table"
                    page.warped = structure.warped
                    page.n_rows, page.n_cols = structure.n_rows, structure.n_cols
                    page.xs, page.ys = list(structure.xs), list(structure.ys)

                    page.preview_jpeg = _encode_jpeg(structure.image)
                    page.overlay_jpeg = _encode_jpeg(structure.overlay)
                    sc = _scale_for(structure.image)
                    page.cell_boxes = {
                        f"{c.row},{c.col}": [int(c.x * sc), int(c.y * sc),
                                             int(c.w * sc), int(c.h * sc)]
                        for c in structure.cells}

                    # 首遍整表 OCR：既用于通用路径，也供 xlsx 模板的
                    # 结构区文本匹配（先把文字拿到手再决定套哪个模板）
                    # 首遍整表 OCR：既用于通用路径，也供 xlsx 模板的
                    # 结构区文本匹配（先把文字拿到手再决定套哪个模板）。
                    # extract_table 已在校正图上跑过完整识别（det_items，
                    # 坐标系同为 structure.image），直接复用不再识别一遍
                    self._step(on_step, "整表 OCR ...")
                    items = self._filter_regions(
                        list(structure.det_items)
                        if structure.det_items is not None
                        else self._engine().recognize_full(structure.image),
                        regions, structure.image.shape)
                    per_cell = self._per_cell_texts(structure, items)

                    applied = False
                    if templates:
                        applied = self._try_template(page, structure, templates,
                                                     template_name, on_step,
                                                     per_cell, fill_mode,
                                                     template_page)
                    if not applied:
                        self._recognize_title(page, structure, on_step)
                        self._build_grid(page, structure)
                        self._fill_from_first_pass(page, per_cell)
                        self._ocr_cells(page, structure, on_step, cancel_event,
                                        regions)
        except _Cancelled:
            page.error = "已取消"
            raise
        except Exception as exc:  # noqa: BLE001 —— 单图失败不中断整批
            page.error = str(exc)
        finally:
            page.elapsed = time.time() - start
        return page

    # ------------------------------------------------------------------ #
    # 整页方向纠正（横放/倒放的照片）
    # ------------------------------------------------------------------ #
    def _fix_page_rotation(self, img, on_step: Optional[StepCallback]):
        """探测整页是否横放/倒放（90/180/270 度），需要时转正。

        两个信号（缩小图探测）：
        - 文字框形状：90/270 横放时文字框普遍"高瘦"（tall_ratio>0.5），
          只有这种情况才需要全方向比较（90 与 270 无法从形状分辨）；
        - 关闭 cls 方向分类后的识别分相对比较：180 倒置不改变文字框
          形状，且干净大字倒置也能 rec 出绝对高分（实测 mean≈0.75 仍
          通过绝对门槛）——绝对分不可信，只信"正立显著优于倒置"
          （≥1.35 倍）的相对比较，因此正立/倒置页固定补测一个 180°。
        """
        try:
            q0 = self._probe_quality(img)
        except Exception:
            return img
        best_k, best_q = 0, q0["q"]
        if q0["tall_ratio"] > 0.5:
            # 横放（90/270）：全方向比较
            for k in (1, 2, 3):
                self._check_cancel(None)
                try:
                    qk = self._probe_quality(np.rot90(img, k))
                except Exception:
                    continue
                if qk["q"] > best_q:
                    best_k, best_q = k, qk["q"]
        else:
            # 正立或倒置：只补测 180°（省 2 次探测，相对比较兜住倒置）
            try:
                q2 = self._probe_quality(np.rot90(img, 2))
            except Exception:
                q2 = None
            if q2 and q2["q"] > best_q:
                best_k, best_q = 2, q2["q"]
        if best_k != 0 and best_q >= max(12.0, q0["q"] * 1.35):
            self._step(on_step, f"已自动转正页面方向（旋转 {best_k * 90}°）")
            return np.rot90(img, best_k)
        return img

    def _probe_quality(self, img) -> Dict:
        """小图探测，返回 q（综合质量分）与 tall_ratio（高瘦文字框占比）。"""
        h, w = img.shape[:2]
        scale = 600.0 / max(h, w)
        small = cv2.resize(img, (int(w * scale), int(h * scale)),
                           interpolation=cv2.INTER_AREA) if scale < 1.0 else img
        items = self._engine().recognize_full(small, use_cls=False)
        if not items:
            return {"q": 0.0, "tall_ratio": 0.0}
        chars = sum(len(str(t)) for _, t, _ in items)
        scores = [float(s) for _, _, s in items]
        mean_score = sum(scores) / len(scores)
        tall = sum(1 for box, _, _ in items
                   if (max(p[1] for p in box) - min(p[1] for p in box))
                   > 1.2 * (max(p[0] for p in box) - min(p[0] for p in box)))
        tall_ratio = tall / len(items)
        q = chars * mean_score * (0.5 if tall_ratio > 0.5 else 1.0)
        return {"q": q, "tall_ratio": tall_ratio}

    # ------------------------------------------------------------------ #
    # 无框线表格（SLANet-Plus 结构识别）
    # ------------------------------------------------------------------ #
    def _process_borderless(self, page: TablePage, img,
                            on_step: Optional[StepCallback],
                            cancel_event: Optional[threading.Event],
                            regions: List[Dict]) -> bool:
        """无框线表格识别；成功返回 True（page 已填充），失败/不像表格返回 False。"""
        try:
            from .table_structure import TableStructureEngine, structure_gate
        except ImportError:
            return False
        proc = preprocess(img)
        self._step(on_step, "无框线表格结构识别 ...")
        items = self._engine().recognize_full(proc)
        self._check_cancel(cancel_event)
        items = self._filter_regions(items, regions, proc.shape)
        if not items:
            return False
        result = TableStructureEngine.instance().recognize(
            proc, [(b, t, s) for b, t, s in items])
        if not result or not structure_gate(result):
            return False
        rows = [[normalize_text(str(c)) for c in r] for r in result["rows"]]
        merges = [list(m) for m in result["merges"]]
        # 模型常把标题当成首行的全宽合并格：提炼为 page.title，与
        # 有框线表格"标题在表格之外"的形态对齐
        score_offset = 0
        if rows and len(rows[0]) >= 2:
            first_nonempty = [c for c in rows[0] if str(c).strip()]
            covers_full = any(m[0] == 0 and m[1] == 0 and m[3] >= len(rows[0])
                              for m in merges)
            if covers_full and len(first_nonempty) == 1:
                if not page.title:
                    page.title = first_nonempty[0]
                score_offset = len(rows[0])
                rows = rows[1:]
                merges = [[m[0] - 1, m[1], m[2], m[3]] for m in merges if m[0] > 0]
        page.mode = "table"
        page.borderless = True
        page.rows = rows
        page.merges = merges
        page.n_rows = len(rows)
        page.n_cols = max(len(r) for r in rows)
        page.min_score = min(float(s) for _, _, s in items)
        self._borderless_scores(page, result.get("cell_bboxes"), items,
                                proc.shape, score_offset)
        self._borderless_boxes(page, result.get("cell_bboxes"),
                               proc, score_offset)

        page.preview_jpeg = _encode_jpeg(proc)
        overlay = _borderless_overlay(proc, result.get("cell_bboxes"))
        if overlay is not None:
            page.overlay_jpeg = _encode_jpeg(overlay)
        self._borderless_title(page, proc, result.get("cell_bboxes"), on_step)
        return True

    def _borderless_scores(self, page: TablePage, bboxes, items, shape,
                           offset: int = 0) -> None:
        """把 OCR 文字条的置信度按位置归到结构网格的每个单元格。

        offset：标题提炼删掉的单元格数（bboxes 按原表行优先展开）。
        """
        if bboxes is None:
            return
        bboxes = np.asarray(bboxes, dtype=np.float32)
        h, w = shape[:2]
        # 单元格按行优先展开（与 HTML 解析的 cells 顺序一致）
        idx = offset
        for r, row in enumerate(page.rows):
            for c in range(len(row)):
                if idx >= len(bboxes):
                    return
                bbox = bboxes[idx]
                idx += 1
                cx = float(bbox[:, 0].mean()) / w
                cy = float(bbox[:, 1].mean()) / h
                scores = [float(s) for box, _, s in items
                          if _box_center_in(box, cx, cy, shape)]
                if scores:
                    page.scores[f"{r},{c}"] = min(scores)

    def _borderless_boxes(self, page: TablePage, bboxes, proc,
                          offset: int = 0) -> None:
        """结构模型的单元格 bbox（含标题 offset）转预览图坐标存入 cell_boxes。"""
        if bboxes is None:
            return
        bboxes = np.asarray(bboxes, dtype=np.float32)
        sc = _scale_for(proc)
        idx = offset
        for r, row in enumerate(page.rows):
            for c in range(len(row)):
                if idx >= len(bboxes):
                    return
                bb = bboxes[idx]
                idx += 1
                x0, y0 = float(bb[:, 0].min()), float(bb[:, 1].min())
                x1, y1 = float(bb[:, 0].max()), float(bb[:, 1].max())
                page.cell_boxes[f"{r},{c}"] = [int(x0 * sc), int(y0 * sc),
                                               int((x1 - x0) * sc), int((y1 - y0) * sc)]

    def _borderless_title(self, page: TablePage, proc,
                          bboxes, on_step: Optional[StepCallback]) -> None:
        """无框线表格的标题：取最高单元格上方的窄条里字号最大的文字。"""
        if bboxes is None:
            return
        bboxes = np.asarray(bboxes, dtype=np.float32)
        top = float(bboxes[:, :, 1].min())
        bottom = float(bboxes[:, :, 1].max())
        left = float(bboxes[:, :, 0].min())
        right = float(bboxes[:, :, 0].max())
        strip_h = min(top, (bottom - top) * 0.5)
        if top < 14 or strip_h < 12:
            return
        x0 = max(0, int(left) - 40)
        x1 = min(proc.shape[1], int(right) + 40)
        crop = proc[int(top - strip_h):int(top), x0:x1]
        if crop.size == 0:
            return
        best_text, best_h = "", 0.0
        for box, text, _ in self._engine().recognize_full(crop):
            text = str(text).strip()
            if not text:
                continue
            bh = max(p[1] for p in box) - min(p[1] for p in box)
            if bh > best_h:
                best_text, best_h = text, bh
        if best_text:
            page.title = normalize_text(best_text)

    # ------------------------------------------------------------------ #
    # 忽略区域过滤（页眉/水印/印章等不需要识别的区域）
    # ------------------------------------------------------------------ #
    @staticmethod
    def _filter_regions(items, regions: List[Dict], shape) -> List:
        """丢弃中心点落在忽略区域内的 OCR 文字条。区域坐标为归一化 0~1。"""
        if not regions:
            return list(items)
        h, w = shape[:2]
        out = []
        for box, text, score in items:
            cx = sum(p[0] for p in box) / len(box) / w
            cy = sum(p[1] for p in box) / len(box) / h
            if _in_regions(cx, cy, regions):
                continue
            out.append((box, text, score))
        return out

    # ------------------------------------------------------------------ #
    # 整页文字模式（无表格线的文档）
    # ------------------------------------------------------------------ #
    def _process_text(self, page: TablePage, img,
                      on_step: Optional[StepCallback],
                      cancel_event: Optional[threading.Event],
                      regions: List[Dict] = None) -> None:
        page.mode = "text"
        proc = preprocess(img)

        self._step(on_step, "整页 OCR ...")
        engine = self._engine()
        items = engine.recognize_full(proc)
        self._check_cancel(cancel_event)
        items = self._filter_regions(items, regions or [], proc.shape)

        # 按中心 y 聚行（相邻条目高度 60% 以内算同一行），行内按 x 排序
        entries = []
        for box, text, score in items:
            text = str(text).strip()
            if not text:
                continue
            ys = [p[1] for p in box]
            xs = [p[0] for p in box]
            h = max(ys) - min(ys)
            entries.append((sum(ys) / len(ys), sum(xs) / len(xs), h, text,
                            float(score), [min(xs), min(ys), max(xs), max(ys)]))
        entries.sort(key=lambda e: e[0])
        lines: List[List] = []
        for cy, cx, h, text, score, bbox in entries:
            if lines and abs(cy - lines[-1][0][0]) < max(6.0, 0.6 * max(h, lines[-1][0][2])):
                lines[-1].append((cy, cx, h, text, score, bbox))
            else:
                lines.append([(cy, cx, h, text, score, bbox)])

        page.rows = []
        sc = _scale_for(proc)
        for i, line in enumerate(lines):
            line.sort(key=lambda e: e[1])
            page.rows.append([normalize_text(" ".join(e[3] for e in line))])
            page.scores[f"{i},0"] = min(e[4] for e in line)
            page.min_score = min(page.min_score, min(e[4] for e in line))
            x0 = min(e[5][0] for e in line); y0 = min(e[5][1] for e in line)
            x1 = max(e[5][2] for e in line); y1 = max(e[5][3] for e in line)
            page.cell_boxes[f"{i},0"] = [int(x0 * sc), int(y0 * sc),
                                         int((x1 - x0) * sc), int((y1 - y0) * sc)]
        page.n_rows, page.n_cols = len(page.rows), 1
        page.merges = []

        page.preview_jpeg = _encode_jpeg(proc)
        page.overlay_jpeg = _encode_jpeg(_text_overlay(proc, items))

    # ------------------------------------------------------------------
    @staticmethod
    def _step(on_step: Optional[StepCallback], text: str) -> None:
        if on_step:
            on_step(text)

    @staticmethod
    def _per_cell_texts(structure: TableStructure, items) -> Dict:
        """整表 OCR 结果按中心点分配到格子：{(r,c): (text, score)}。"""
        per: Dict[Tuple[int, int], List] = {}
        for box, text, score in items:
            cell = Scan2ExcelService._locate(structure, box)
            if cell is None:
                continue
            ys = [p[1] for p in box]
            xs = [p[0] for p in box]
            per.setdefault((cell.row, cell.col), []).append(
                (sum(ys) / len(ys), sum(xs) / len(xs), str(text), float(score)))
        out: Dict[Tuple[int, int], Tuple[str, float]] = {}
        for key, rows in per.items():
            out[key] = Scan2ExcelService._merge_fragments(rows)
        return out

    @staticmethod
    def _fill_from_first_pass(page: TablePage, per_cell: Dict) -> None:
        for (r, c), (text, score) in per_cell.items():
            if r < len(page.rows) and c < len(page.rows[r]):
                page.rows[r][c] = normalize_text(text)
                page.scores[f"{r},{c}"] = score
                page.min_score = min(page.min_score, score)

    def _try_template(self, page: TablePage, structure: TableStructure,
                      templates: List, template_name: str,
                      on_step: Optional[StepCallback],
                      per_cell: Optional[Dict] = None,
                      fill_mode: str = "vlookup",
                      template_page: str = "") -> bool:
        """月计表模板模式：命中模板则只 OCR 数值列（行列与科目文本冻结）。

        指定了 template_name 时强制使用该模板（仍校验，不匹配只告警）；
        否则按科目代码命中率自动匹配，低于阈值返回 False 走通用识别。
        """
        if len(structure.xs) < 2 or len(structure.ys) < 2:
            return False

        # ---- xlsx 模板：尺寸严格匹配 + 结构区文本校验（用户主导的稳健方案）----
        xlsx_tpls = [t for t in templates
                     if getattr(t, "pages", None) is not None
                     and hasattr(t, "path")]
        if xlsx_tpls:
            from .xlsx_template import apply_xlsx_template, match_xlsx_template
            doc, tpl_page, score, note, cells = match_xlsx_template(
                xlsx_tpls, structure,
                len(structure.ys) - 1, len(structure.xs) - 1,
                force_doc=template_name, force_page=template_page)
            if doc is not None:
                self._step(on_step, f"套用 xlsx 模板：{doc.name}·{tpl_page.page_name}"
                                     f"（结构区命中 {score:.0%}，"
                                     + ("按科目对齐填充" if fill_mode == "vlookup"
                                        else "按位置填充") + "）")
                warns, mismatch = apply_xlsx_template(
                    page, structure, tpl_page, self._engine(),
                    doc_name=doc.name, cells=cells, fill_mode=fill_mode)
                # 模板路径同样识别表格上方标题（此前被跳过，界面"表格标题"
                # 一直为空；标题也可用于按标题命名导出）
                self._recognize_title(page, structure, on_step)
                # 立即求值勾稽（条件格式规则）——识别完就能看到不平的行，
                # 不必等到导出；结果存 page.checks 供界面与导出共用，
                # page.check_rows 供界面把整行标红
                checks, check_rows = self._run_check_formulas(doc, page)
                page.checks = checks
                page.check_rows = check_rows
                page.header_rows = max(1, int(getattr(tpl_page, "header_rows", 1) or 1))
                page.template_file = str(doc.path)
                page.template_page = tpl_page.page_name
                # 手搓模板（无几何信息）：首次成功套用后回写几何比例，
                # 之后匹配不再要求照片网格与模板尺寸严格一致
                if (not getattr(tpl_page, "col_fracs", None) and cells
                        and len(cells) == tpl_page.n_rows * tpl_page.n_cols):
                    from .xlsx_template import capture_geometry
                    if capture_geometry(Path(doc.path), tpl_page.page_name,
                                        structure.xs, structure.ys):
                        page.geometry_saved = True
                        self._step(on_step, "已为模板页「%s」记录几何比例"
                                            "（后续匹配更稳）" % tpl_page.page_name)
                if mismatch and (mismatch.get("unmatched") or mismatch.get("extra")):
                    n_bad = (len(mismatch.get("unmatched") or [])
                             + len(mismatch.get("extra") or []))
                    self._step(on_step, f"行对齐：{mismatch.get('matched', 0)} 行匹配，"
                                         f"{n_bad} 行需人工确认")
                base = [w for w in warns if w]
                if note:      # 指定页低命中率等告警
                    base.append(note)
                if checks:
                    base.extend(checks)
                if base:
                    # 多条提示一行一条（界面 pre-line 展示，长列举不再糊成一行）
                    page.warning = "\n".join(base[:8])
                    self._step(on_step, "校验提示：" + page.warning.replace("\n", " ⏎ "))
                return True
            if note:
                self._step(on_step, f"xlsx 模板未匹配（{note}）")
                # 原因挂到警告：图片列表的 ⚠ 直接可见（尺寸不符/命中率不足），
                # 免得"未匹配模板"三个字让人猜是哪里不对
                page.warning = f"未匹配模板：{note}"

        # ---- 旧版 JSON 模板（沿用比例几何方案）----
        from .template_mode import apply_template, match_template
        legacy = [t for t in templates if not hasattr(t, "path")]
        if not legacy:
            return False
        templates = legacy
        # 指定了模板文档名时只在它的各页里路由；否则全库自动匹配
        doc, tpl_page, score = match_template(structure, templates,
                                              force_doc=template_name)
        if doc is None:
            if template_name:
                self._step(on_step, f"模板「{template_name}」各页均未匹配"
                                     f"（最高命中率 {score:.0%}），按通用模式识别")
            else:
                self._step(on_step, f"未匹配到月计表模板（命中率 {score:.0%}），"
                                     "按通用模式识别")
            return False
        self._step(on_step, f"套用月计表模板：{doc.name} · {tpl_page.page_name}")
        warnings = apply_template(page, structure, tpl_page, self._engine(),
                                  doc_name=doc.name)
        if warnings:
            page.warning = "；".join(warnings[:6])
            self._step(on_step, "模板校验提示：" + page.warning)
        return True

    @staticmethod
    def _run_check_formulas(doc, page):
        """把当前页填充值写入模板副本并求值条件格式（勾稽关系）。

        返回 (提示消息列表, 不平的页行号 0 基列表)——行号供界面整行标红。
        """
        try:
            import tempfile
            from pathlib import Path
            from .xlsx_template import evaluate_checks, fill_template_workbook
            tpl_path = Path(doc.path)
            if not tpl_path.is_file():
                return [], []
            page_name = getattr(page, "template_page", "") or                 (doc.pages[0].page_name if doc.pages else "")
            with tempfile.TemporaryDirectory() as td:
                out = Path(td) / "filled.xlsx"
                fill_template_workbook(tpl_path, {page_name: page}, out)
                return evaluate_checks(tpl_path, out, sheets=[page_name])
        except Exception:
            return [], []

    _TITLE_STOP = ("科目代码", "科目名称", "上期余额", "本期发生额", "本期余额",
                   "借方", "贷方", "合计", "编制日期", "编制单位", "制表",
                   "审核", "复核", "共", "第", "页", "单位")

    @staticmethod
    def _title_like(text: str) -> bool:
        """剔除表头词/数字/标点后仍剩实质汉字才像标题。

        透视裁剪会把表框外的真标题裁掉，候选区时常只剩"上期余额"这类
        表头片段——宁要空标题也不要表头冒充。
        """
        t = str(text or "")
        for w in Scan2ExcelService._TITLE_STOP:
            t = t.replace(w, "")
        t = re.sub(r"[\s0-9（）()：:．.、·—\-‐–]+", "", t)
        return len(t) >= 2

    def render_scan(self, path: str, mode: str = "enhanced",
                    on_step: Optional[StepCallback] = None,
                    cancel_event: Optional[threading.Event] = None) -> np.ndarray:
        """拍照图 → 扫描件：整页方向转正（复用 OCR 探测）→ 透视拉平 → 增强。

        与识别管线互相独立，不影响模板匹配结果。
        """
        self._step(on_step, f"读取图片：{os.path.basename(path)}")
        img = imread_unicode(path)   # EXIF 方向已在此处理
        if img is None:
            raise ImageReadError(
                f"无法读取图片（{os.path.splitext(path)[1]}）：解码失败")
        self._step(on_step, "方向探测 / 转正 ...")
        img = self._fix_page_rotation(img, on_step)
        self._check_cancel(cancel_event)
        self._step(on_step, "透视拉平 ...")
        flat = flatten_page(img)
        self._check_cancel(cancel_event)
        self._step(on_step, f"扫描增强（{mode}）...")
        return enhance_scan(flat, mode)

    def _recognize_title(self, page: TablePage, structure: TableStructure,
                         on_step: Optional[StepCallback]) -> None:
        """识别表格上方的标题：取"字号最大的像标题的文字"。

        首选透视前图上、沿 quad 上边缘向外切出的条带（透视拉正后 OCR）：
        标题在表框之外，透视裁剪必然把它裁掉，A/B 候选区时常只剩表头
        片段（A 捡到组表头、B 的横线检测在畸变图上漂到表格中部）；A/B
        区只作补充（quad 缺失时的唯一来源）。表头词一律不参与评选。
        """
        engine = self._engine()
        candidates: List[Tuple[float, str]] = []   # (文字框高度, 文本)

        quad = getattr(structure, "quad", None)
        pre = structure.pre_image
        if quad is not None and pre is not None and len(quad) == 4:
            try:
                q = np.asarray(quad, dtype=np.float32)
                top_edge = float(np.linalg.norm(q[1] - q[0]))
                band = float(np.clip(top_edge * 0.12, 48.0, 320.0))
                src = np.float32([q[0] - (0, band), q[1] - (0, band), q[1], q[0]])
                w = max(int(top_edge), 40)
                strip = cv2.warpPerspective(
                    pre,
                    cv2.getPerspectiveTransform(
                        src, np.float32([[0, 0], [w, 0],
                                         [w, int(band)], [0, int(band)]])),
                    (w, int(band)))
                for box, text, _s in engine.recognize_full(strip):
                    text = str(text).strip()
                    if not text:
                        continue
                    bh = max(p[1] for p in box) - min(p[1] for p in box)
                    candidates.append((bh * 1.15, text))   # 条带内候选优先
            except Exception:
                pass    # 条带构建失败退回 A/B 区

        det_items = list(structure.det_items or [])
        for tag, x, y, w, h in structure.title_zones:
            src = structure.image if tag == "A" else structure.pre_image
            if src is None:
                continue
            if tag == "A" and det_items:
                # A 区复用整表 det 结果按中心点落区，省一次 OCR
                for box, text, _s in det_items:
                    cx = sum(p[0] for p in box) / len(box)
                    cy = sum(p[1] for p in box) / len(box)
                    if x <= cx < x + w and y <= cy < y + h:
                        bh = max(p[1] for p in box) - min(p[1] for p in box)
                        candidates.append((bh, str(text).strip()))
                continue
            crop = src[y:y + h, x:x + w]
            if crop.size == 0:
                continue
            for box, text, _s in engine.recognize_full(crop):
                bh = max(p[1] for p in box) - min(p[1] for p in box)
                candidates.append((bh, str(text).strip()))

        best, best_h = "", 0.0
        for bh, text in candidates:
            text = text.strip()
            if not text or not self._title_like(text):
                continue
            if bh > best_h:
                best, best_h = text, bh
        if best:
            page.title = normalize_text(best)

    @staticmethod
    def _check_cancel(cancel_event: Optional[threading.Event]) -> None:
        if cancel_event is not None and cancel_event.is_set():
            raise _Cancelled()

    def _build_grid(self, page: TablePage, structure: TableStructure) -> None:
        page.rows = [[""] * structure.n_cols for _ in range(structure.n_rows)]
        page.merges = []
        for cell in structure.cells:
            if cell.row_span > 1 or cell.col_span > 1:
                page.merges.append([cell.row, cell.col, cell.row_span, cell.col_span])

    def _ocr_cells(self, page: TablePage, structure: TableStructure,
                   on_step: Optional[StepCallback],
                   cancel_event: Optional[threading.Event],
                   regions: List[Dict] = None) -> None:
        engine = self._engine()

        # 第一步：整表一次识别（det 精裁文本条，质量最高），按中心点分配到格子。
        # extract_table 已顺带跑过 det（供文字掩模用）时直接复用，不重复识别
        self._step(on_step, "整表 OCR ...")
        if structure.det_items is not None:
            items = list(structure.det_items)
        else:
            items = engine.recognize_full(structure.image)
        items = self._filter_regions(items, regions, structure.image.shape)
        per_cell: Dict[Tuple[int, int], List] = {}
        for box, text, score in items:
            cell = self._locate(structure, box)
            if cell is None:
                continue
            ys = [p[1] for p in box]
            xs = [p[0] for p in box]
            per_cell.setdefault((cell.row, cell.col), []).append(
                (sum(ys) / len(ys), sum(xs) / len(xs), str(text), float(score)))
        for (r, c), rows in per_cell.items():
            text, score = self._merge_fragments(rows)
            page.rows[r][c] = normalize_text(text)
            page.scores[f"{r},{c}"] = score
            page.min_score = min(page.min_score, score)

        # 第二步：仍为空但非空白的格子，逐格完整模式补漏。
        # 拍照件的阴影格可能整批被判"非空"导致补漏风暴（每格一次 det+rec，
        # 数十秒），仅当格内 OCR 前景确实可见时才有价值——用归一化背景的
        # 文字密度把关：密度极低的阴影/噪点格直接视为空白跳过。
        # 巨格（>图宽/高 60%）是竖/横线漏检的病态切分产物（如整行数字
        # 粘成一格），逐格识别又慢又无意义，直接跳过。
        h_img, w_img = structure.image.shape[:2]
        holes = [cell for cell in structure.cells
                 if not page.rows[cell.row][cell.col]
                 and cell.w <= w_img * 0.6 and cell.h <= h_img * 0.6
                 and not _cell_in_regions(cell, regions, structure.image.shape)]
        done = 0
        for cell in holes:
            self._check_cancel(cancel_event)
            crop = structure.image[cell.y:cell.y + cell.h, cell.x:cell.x + cell.w]
            if engine.is_blank(crop):
                continue
            text, score = engine.recognize_cell(crop)
            if text:
                page.rows[cell.row][cell.col] = normalize_text(text)
                page.scores[f"{cell.row},{cell.col}"] = score
                page.min_score = min(page.min_score, score)
            done += 1
            if on_step and (done % 10 == 0 or done == len(holes)):
                self._step(on_step, f"补漏识别 {done}/{len(holes)} 格")

    @staticmethod
    def _locate(structure: TableStructure, box) -> Optional:
        """按文字条中心点找到所属单元格。"""
        cx = sum(p[0] for p in box) / len(box)
        cy = sum(p[1] for p in box) / len(box)
        for cell in structure.cells:
            if cell.x <= cx < cell.x + cell.w and cell.y <= cy < cell.y + cell.h:
                return cell
        return None

    @staticmethod
    def _merge_fragments(rows: List) -> Tuple[str, float]:
        """同一格子的多条文字按位置排序拼接。

        行内直接连接（不加空格）：单元格内容多为编码/数字/中文词，
        det 分段拼接时插入空格反而破坏内容（如 WL-1003 → 'WL-1 1003'）。
        """
        if not rows:
            return "", 1.0
        if len(rows) == 1:
            return rows[0][2], rows[0][3]
        lines: List[List] = []
        for cy, cx, text, score in sorted(rows, key=lambda r: (r[0], r[1])):
            if lines and abs(cy - lines[-1][0][0]) < 8.0:
                lines[-1].append((cy, cx, text, score))
            else:
                lines.append([(cy, cx, text, score)])
        texts = ["".join(t for _, _, t, _ in sorted(l, key=lambda r: r[1])) for l in lines]
        scores = [s for *_, s in rows]
        return "".join(texts), min(scores)


class _Cancelled(Exception):
    pass


PREVIEW_MAX_SIDE = 1100


def _scale_for(img: np.ndarray, max_side: int = PREVIEW_MAX_SIDE) -> float:
    """预览缩放系数（cell_boxes 坐标与预览 JPEG 必须同一变换）。"""
    h, w = img.shape[:2]
    side = max(h, w)
    return (max_side / side) if side > max_side else 1.0


def _encode_jpeg(img: np.ndarray, max_side: int = PREVIEW_MAX_SIDE, quality: int = 80) -> bytes:
    """预览图编码为 JPEG（限尺寸，供界面 base64 展示）。"""
    h, w = img.shape[:2]
    if max(h, w) > max_side:
        scale = max_side / max(h, w)
        img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, quality])
    return buf.tobytes() if ok else b""


def _text_overlay(img: np.ndarray, items) -> np.ndarray:
    """文字模式预览：每条识别结果画绿框，便于肉眼核对。"""
    vis = img.copy()
    for box, text, score in items:
        pts = np.array(box, dtype=np.int32)
        cv2.polylines(vis, [pts], True, (80, 200, 80), 2)
    return vis


def _in_regions(cx: float, cy: float, regions: List[Dict]) -> bool:
    """归一化坐标 (cx, cy) 是否落在任一忽略区域 [x, y, w, h] 内。"""
    return any(r["x"] <= cx <= r["x"] + r["w"] and r["y"] <= cy <= r["y"] + r["h"]
               for r in regions)


def _box_center_in(box, cx: float, cy: float, shape) -> bool:
    """OCR 文字条中心（归一化）是否落在单元格 bbox（归一化 cx, cy）附近。

    用于把 OCR 置信度归属到无框线表格的单元格：比较文字条中心与
    单元格中心的归一化距离。
    """
    bx = sum(p[0] for p in box) / len(box)
    by = sum(p[1] for p in box) / len(box)
    h, w = shape[:2]
    return abs(bx / w - cx) < 0.05 and abs(by / h - cy) < 0.05


def _cell_in_regions(cell, regions: List[Dict], shape) -> bool:
    """单元格中心（像素）是否落在任一忽略区域内。"""
    if not regions:
        return False
    h, w = shape[:2]
    return _in_regions((cell.x + cell.w / 2) / w, (cell.y + cell.h / 2) / h, regions)


def _borderless_overlay(img: np.ndarray, bboxes) -> Optional[np.ndarray]:
    """无框线表格预览：画结构模型输出的单元格框。"""
    if bboxes is None:
        return None
    vis = img.copy()
    for bbox in np.asarray(bboxes, dtype=np.int32):
        cv2.polylines(vis, [bbox], True, (80, 200, 80), 2)
    return vis


def thumb_from_jpeg(base: bytes, max_side: int = 96, quality: int = 72) -> bytes:
    """从预览 JPEG 解码出更小的列表缩略图；失败返回空 bytes。"""
    try:
        arr = np.frombuffer(base, np.uint8)
        img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
        if img is None:
            return b""
        h, w = img.shape[:2]
        side = max(h, w)
        if side > max_side:
            sc = max_side / side
            img = cv2.resize(img, (int(w * sc), int(h * sc)),
                             interpolation=cv2.INTER_AREA)
        ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, quality])
        return buf.tobytes() if ok else b""
    except Exception:
        return b""

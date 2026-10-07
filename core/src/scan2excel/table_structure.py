# -*- coding: utf-8 -*-
"""无框线表格结构识别（RapidTable / SLANet-Plus）。

针对"没有完整打印框线"的表格：先用 OCR 拿到文字条及其位置，
再由序列预测模型还原行/列结构（含合并单元格），输出与
table_extractor.TableStructure 同构的行列数据。

模型 slanet-plus.onnx 约 7.4MB，CPU 单图约 0.12s，首次使用时自动
下载到包内缓存。
"""
from __future__ import annotations

import threading
from html.parser import HTMLParser
from typing import Dict, List, Optional, Tuple

import numpy as np


class _TableHTMLParser(HTMLParser):
    """解析 pred_html 里的 <table>，产出 (cells, n_rows, n_cols)。

    cells 与 pred_html 单元格同序（行优先，合并单元格记在锚点），
    同时收集 rowspan/colspan 供合并还原。
    """

    def __init__(self) -> None:
        super().__init__()
        self.cells: List[List[str]] = []
        self.spans: List[Tuple[int, int, int, int]] = []  # (r, c, rspan, cspan)
        self._row: List[str] = []
        self._in_cell = False
        self._buf: List[str] = []
        self._cell_attrs: Dict[str, int] = {}
        self._row_idx = 0
        self._col_cursor = 0
        self._occupied: Dict[Tuple[int, int], bool] = {}
        self._max_col = 0

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
            self._col_cursor = 0
        elif tag in ("td", "th"):
            self._in_cell = True
            self._buf = []
            self._cell_attrs = dict(attrs)
            # 跳过被 rowspan/colspan 占用的位置
            while (self._row_idx, self._col_cursor) in self._occupied:
                self._col_cursor += 1
            self._cell_attrs["_col"] = self._col_cursor

    def handle_data(self, data):
        if self._in_cell:
            self._buf.append(data)

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._in_cell:
            self._in_cell = False
            r = self._row_idx
            c = self._cell_attrs.pop("_col")
            rspan = int(self._cell_attrs.get("rowspan", 1) or 1)
            cspan = int(self._cell_attrs.get("colspan", 1) or 1)
            text = "".join(self._buf).replace("\u00a0", " ").strip()
            self._row.append(text)
            if rspan > 1 or cspan > 1:
                self.spans.append((r, c, rspan, cspan))
                for rr in range(r, r + rspan):
                    for cc in range(c, c + cspan):
                        self._occupied[(rr, cc)] = True
            self._col_cursor = c + cspan
            self._max_col = max(self._max_col, c + cspan)
        elif tag == "tr" and self._row:
            self.cells.append(self._row)
            self._row_idx += 1


class TableStructureEngine:
    _instance: Optional["TableStructureEngine"] = None
    _lock = threading.Lock()

    @classmethod
    def instance(cls) -> "TableStructureEngine":
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = cls()
        return cls._instance

    def __init__(self) -> None:
        from rapid_table import RapidTable
        from rapid_table.utils.typings import ModelType, RapidTableInput

        # use_ocr=True 但始终传入自带 OCR 结果（get_ocr_results 优先用
        # ocr_results，不会调用内置 rapidocr）；pred_html 只在该分支生成
        self._engine = RapidTable(
            RapidTableInput(model_type=ModelType.SLANETPLUS, use_ocr=True))

    def recognize(self, img: np.ndarray, ocr_items: List) -> Optional[Dict]:
        """输入整图与 OCR 结果，返回 {rows, merges, cell_bboxes}；失败返回 None。

        ocr_items: [[box, text, score], ...]（与 OcrEngine.recognize_full 输出一致）。
        注意 rapid_table 的 ocr_results 是"每张图一个元组"：(全部框(N,4,2),
        全部文字, 全部分数)。
        """
        if not ocr_items:
            return None
        boxes = np.asarray([np.asarray(box, dtype=np.float32)
                            for box, _, _ in ocr_items], dtype=np.float32)
        texts = tuple(str(t) for _, t, _ in ocr_items)
        scores = tuple(float(s) for _, _, s in ocr_items)
        try:
            out = self._engine(img, ocr_results=[(boxes, texts, scores)])
        except Exception:
            return None
        html = (out.pred_htmls or [""])[0]
        if "<table" not in html:
            return None
        parser = _TableHTMLParser()
        try:
            parser.feed(html)
        except Exception:
            return None
        rows = parser.cells
        if not rows:
            return None
        bboxes = out.cell_bboxes[0] if out.cell_bboxes else None
        if bboxes is not None:
            # 模型输出压平的 (N,8)：四点 x1,y1,x2,y2,...，重排为 (N,4,2)
            bboxes = np.asarray(bboxes, dtype=np.float32).reshape(-1, 4, 2)
        return {"rows": rows, "merges": parser.spans, "cell_bboxes": bboxes}


def structure_gate(result: Dict) -> bool:
    """判定无框线识别结果是否可信（防纯文字页被误判成表格）。

    要求至少 2 列且大部分文字确实落在网格里；纯文档只会得到
    单列或空壳结构，过不了这道门。
    """
    rows = result.get("rows") or []
    if len(rows) < 2:
        return False
    n_cols = max(len(r) for r in rows)
    if n_cols < 2:
        return False
    filled = sum(1 for r in rows for c in r if str(c).strip())
    total = sum(len(r) for r in rows)
    if total == 0 or filled / total < 0.5:
        return False
    # 每列平均至少要有一定数量的非空格（避免 1 行 x N 列的碎结构）
    col_fill = [sum(1 for r in rows if c < len(r) and str(r[c]).strip())
                for c in range(n_cols)]
    if sum(1 for f in col_fill if f >= 2) < 2:
        return False
    return True

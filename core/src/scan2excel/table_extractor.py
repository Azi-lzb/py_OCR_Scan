# -*- coding: utf-8 -*-
"""表格框线检测与单元格切分。

针对"打印表格拍照"场景：表格带完整框线，用形态学检测横竖线切出
每个单元格，附带透视校正，容忍拍照倾斜。纯 OpenCV，不依赖深度模型。
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import cv2
import numpy as np

# 处理图片时的长边上限，超出则等比缩小（兼顾速度与识别率）
MAX_SIDE = 2200
# 一条"线"的最小长度（占整宽/整高的比例）
MIN_LINE_RATIO = 0.12
# 聚类间隔：像素距离小于该值的两条线并作一条
CLUSTER_GAP = 10


@dataclass
class TableCell:
    """一个单元格在校正后图像上的位置与网格坐标。"""
    row: int
    col: int
    row_span: int
    col_span: int
    x: int
    y: int
    w: int
    h: int


@dataclass
class TableStructure:
    cells: List[TableCell] = field(default_factory=list)
    n_rows: int = 0
    n_cols: int = 0
    image: Optional[np.ndarray] = None   # 校正后的 BGR 图
    overlay: Optional[np.ndarray] = None  # 画了检测结果的预览图
    warped: bool = False                  # 是否做了透视校正
    xs: List[int] = field(default_factory=list)  # 竖线 x 坐标
    ys: List[int] = field(default_factory=list)  # 横线 y 坐标
    # 标题定位：透视校正会丢弃表格线框外的内容（标题常在框外），
    # 保存校正前图像与表格上边界，供上层单独识别标题。
    pre_image: Optional[np.ndarray] = None   # 转正（未透视）图
    quad: Optional[np.ndarray] = None        # 表格外框四点（转正图坐标系）
    # 候选标题区域列表 [(tag, x, y, w, h)]，tag: 'A'=处理图 'B'=透视前图
    title_zones: List[Tuple[str, int, int, int, int]] = field(default_factory=list)

    def build_title_zones(self) -> None:
        """生成候选标题区域（表格首条横线之上的窄条）。

        'A' 处理图（透视后）：表格必然平正，quad 若是"整图/纸面"级四边形，
            标题会被带进处理后图内，位于首条横线上方；
        'B' 透视前图：quad 若是表格线框本身，标题只在透视前图上可见。
        两种情况互补，由上层分别 OCR 后择优。
        """
        self.title_zones = []
        # ---- zone A：处理图上首条横线之上 ----
        if self.image is not None and len(self.ys) >= 2 and len(self.xs) >= 2:
            top = self.ys[0]
            ih, iw = self.image.shape[:2]
            table_h = self.ys[-1] - self.ys[0]
            y0 = max(0, top - int(table_h * 0.5))
            # 处理图四周有 28px 白 pad，zone 必须显著高于它才可能含标题
            if top - y0 >= 48:
                x0 = max(0, self.xs[0] - 40)
                x1 = min(iw, self.xs[-1] + 40)
                self.title_zones.append(("A", x0, y0, x1 - x0, top - y0))
        # ---- zone B：透视前图上首条"真"横线之上 ----
        if self.pre_image is not None and self.pre_image is not self.image:
            ph, pw = self.pre_image.shape[:2]
            bin_img = _binarize(self.pre_image)
            h_mask = _line_mask(bin_img, horizontal=True)
            ys_pre = _collect_line_positions(h_mask, horizontal=True, total=pw)
            if len(ys_pre) >= 3:
                top = _first_true_line(ys_pre)
                if top is not None and top >= 14:
                    bottom = ys_pre[-1]
                    table_h = bottom - top
                    y0 = max(0, top - int(table_h * 0.5))
                    if top - y0 >= 12:
                        if self.quad is not None:
                            left = int(min(self.quad[0][0], self.quad[3][0]))
                            right = int(max(self.quad[1][0], self.quad[2][0]))
                        else:
                            left, right = 0, pw
                        x0 = max(0, left - 40)
                        x1 = min(pw, right + 40)
                        self.title_zones.append(("B", x0, y0, x1 - x0, top - y0))


def imread_unicode(path: str) -> Optional[np.ndarray]:
    """读取图片（兼容中文路径与 EXIF 旋转），返回 BGR 数组。"""
    try:
        from PIL import Image, ImageOps
        img = Image.open(path)
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        return cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2BGR)
    except Exception:
        return None


def _limit_side(img: np.ndarray) -> np.ndarray:
    h, w = img.shape[:2]
    side = max(h, w)
    if side <= MAX_SIDE:
        return img
    scale = MAX_SIDE / side
    return cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)


def _order_quad(pts: np.ndarray) -> np.ndarray:
    """将 4 个点排成 左上/右上/右下/左下。"""
    pts = pts.reshape(4, 2).astype(np.float32)
    s = pts.sum(axis=1)
    d = np.diff(pts, axis=1).ravel()
    return np.array([
        pts[np.argmin(s)],   # 左上：x+y 最小
        pts[np.argmin(d)],   # 右上：y-x 最小
        pts[np.argmax(s)],   # 右下：x+y 最大
        pts[np.argmax(d)],   # 左下：y-x 最大
    ], dtype=np.float32)


def find_table_quad(img: np.ndarray) -> Optional[np.ndarray]:
    """在图中找表格/纸张外框的四边形，返回排序后的 4 点；找不到返回 None。"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blur, 50, 150)
    edges = cv2.dilate(edges, np.ones((3, 3), np.uint8), iterations=1)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    img_area = img.shape[0] * img.shape[1]
    candidates = sorted(contours, key=cv2.contourArea, reverse=True)[:5]
    for cnt in candidates:
        area = cv2.contourArea(cnt)
        if area < img_area * 0.25:   # 外框应至少覆盖画面 1/4
            continue
        peri = cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, 0.02 * peri, True)
        if len(approx) == 4 and cv2.isContourConvex(approx):
            return _order_quad(approx)
    return None


def warp_perspective(img: np.ndarray, quad: np.ndarray) -> np.ndarray:
    (tl, tr, br, bl) = quad
    width = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
    height = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
    width, height = max(width, 32), max(height, 32)
    dst = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype=np.float32)
    m = cv2.getPerspectiveTransform(quad, dst)
    out = cv2.warpPerspective(img, m, (width, height))
    # 透视插值在画布最外圈可能留下背景色残条，会被误检成表格线，置白消除
    out[:2, :] = 255
    out[-2:, :] = 255
    out[:, :2] = 255
    out[:, -2:] = 255
    return out


def _binarize(img: np.ndarray) -> np.ndarray:
    """自适应二值化（INV：线条与文字为白），前置 CLAHE 缓解拍照光照不均。"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    block = max(15, min(71, (gray.shape[1] // 20) | 1))
    return cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                 cv2.THRESH_BINARY_INV, block, 12)


def _line_mask(bin_img: np.ndarray, horizontal: bool) -> np.ndarray:
    """形态学提取长横线/长竖线。"""
    h, w = bin_img.shape[:2]
    if horizontal:
        klen = max(20, int(w // 25))
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (klen, 1))
        link = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 1))
    else:
        klen = max(20, int(h // 25))
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, klen))
        link = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 9))
    mask = cv2.morphologyEx(bin_img, cv2.MORPH_OPEN, kernel)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, link)  # 连接轻微断裂的线
    return mask


def _collect_line_positions(mask: np.ndarray, horizontal: bool, total: int) -> List[int]:
    """从线 mask 中提取连通线段，两轮聚类合并出整线坐标列表。

    第一轮按固定小间距聚类；第二轮按"典型网格间距"的比例合并贴边双线
    （透视校正后图像边缘灰边与表格边框会被检出为两条近邻线）。
    """
    n, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    centers: List[float] = []
    min_len = max(25.0, total * MIN_LINE_RATIO)
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if horizontal:
            length, thick, center = w, h, y + h / 2.0
        else:
            length, thick, center = h, w, x + w / 2.0
        if length >= min_len and thick <= 14:
            centers.append(center)
    if not centers:
        return []
    centers.sort()
    merged = _cluster_1d(centers, CLUSTER_GAP)
    # 第二轮：近邻间距远小于典型间距的视为同一条线
    if len(merged) >= 3:
        gaps = [merged[i + 1] - merged[i] for i in range(len(merged) - 1)]
        # 去掉最大的一段（可能是误检 outlier）后取中位数更稳
        gaps_sorted = sorted(gaps)[:-1] if len(gaps) >= 3 else gaps
        typical = median(gaps_sorted)
        adaptive = max(CLUSTER_GAP, typical * 0.35)
        merged = _cluster_1d(merged, adaptive)
    return [int(round(c)) for c in merged]


def _cluster_1d(values: List[float], gap: float) -> List[float]:
    """一维聚类：间距 <= gap 的相邻值并作一组，取组内均值。"""
    groups: List[List[float]] = [[values[0]]]
    for v in values[1:]:
        if v - groups[-1][-1] <= gap:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def _grid_cells(xs: List[int], ys: List[int]) -> List[TableCell]:
    """纯网格法：每行×每列一个格子（不含合并单元格信息）。"""
    cells = []
    for r in range(len(ys) - 1):
        for c in range(len(xs) - 1):
            cells.append(TableCell(r, c, 1, 1, xs[c], ys[r], xs[c + 1] - xs[c], ys[r + 1] - ys[r]))
    return cells


def _contour_cells(grid_mask: np.ndarray, xs: List[int], ys: List[int]) -> List[TableCell]:
    """轮廓法：封闭格子即单元格，天然支持合并单元格（跨行/跨格的大封闭区）。"""
    contours, hierarchy = cv2.findContours(grid_mask, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    if hierarchy is None:
        return []
    min_cell = 100  # 至少 10x10 像素
    img_area = grid_mask.shape[0] * grid_mask.shape[1]
    cells = []
    for i, cnt in enumerate(contours):
        if hierarchy[0][i][3] == -1:      # 只取内轮廓（表格线围出的封闭区）
            continue
        x, y, w, h = cv2.boundingRect(cnt)
        if w * h < min_cell:
            continue
        # 覆盖大半张图的"格子"是轮廓闭合失败产生的假象，丢弃
        if w * h > img_area * 0.6:
            continue
        c0 = _snap_index(x, xs)
        c1 = _snap_index(x + w, xs)
        r0 = _snap_index(y, ys)
        r1 = _snap_index(y + h, ys)
        if c0 is None or c1 is None or r0 is None or r1 is None:
            continue
        if c1 <= c0 or r1 <= r0 or c1 > len(xs) - 1 or r1 > len(ys) - 1:
            continue
        cells.append(TableCell(r0, c0, r1 - r0, c1 - c0, x, y, w, h))
    # 去重（同一格子偶尔出现嵌套轮廓）：同 (row,col) 保留面积最大的
    best = {}
    for cell in cells:
        key = (cell.row, cell.col)
        if key not in best or cell.w * cell.h > best[key].w * best[key].h:
            best[key] = cell
    return sorted(best.values(), key=lambda c: (c.row, c.col))


def _snap_index(v: int, lines: List[int]) -> Optional[int]:
    """把像素坐标吸附到最近的网格线序号（容差为半条网格间距）。"""
    if not lines:
        return None
    gaps = [lines[i + 1] - lines[i] for i in range(len(lines) - 1)]
    tol = max(6, (median(gaps) if gaps else 20) // 2)
    best_i, best_d = None, None
    for i, x in enumerate(lines):
        d = abs(x - v)
        if best_d is None or d < best_d:
            best_i, best_d = i, d
    return best_i if best_d is not None and best_d <= tol else None


def median(values: List[int]) -> int:
    if not values:
        return 0
    s = sorted(values)
    return s[len(s) // 2]


def _draw_overlay(img: np.ndarray, structure: TableStructure) -> np.ndarray:
    vis = img.copy()
    # 检测到的表格线：绿色
    for x in structure.xs:
        cv2.line(vis, (x, 0), (x, vis.shape[0] - 1), (80, 200, 80), 2)
    for y in structure.ys:
        cv2.line(vis, (0, y), (vis.shape[1] - 1, y), (80, 200, 80), 2)
    # 每个单元格：红色描边
    for cell in structure.cells:
        cv2.rectangle(vis, (cell.x + 1, cell.y + 1),
                      (cell.x + cell.w - 1, cell.y + cell.h - 1), (60, 60, 230), 2)
        if cell.row_span > 1 or cell.col_span > 1:
            cv2.putText(vis, f"{cell.row_span}x{cell.col_span}",
                        (cell.x + 4, cell.y + 20), cv2.FONT_HERSHEY_SIMPLEX,
                        0.55, (60, 60, 230), 2)
    return vis


def _estimate_skew(img: np.ndarray) -> Optional[float]:
    """用 Hough 直线估计图像整体小角度歪斜（扫描件/ADF 常见 1~3 度）。

    横竖线被旋转后不再严格水平/垂直，形态学检测会失效，必须先转正。
    返回角度（度，绕图像中心顺时针为正）；无法可靠估计时返回 None。
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(cv2.GaussianBlur(gray, (3, 3), 0), 50, 150)
    h, w = img.shape[:2]
    min_len = int(min(h, w) * 0.3)
    lines = cv2.HoughLinesP(edges, 1, np.pi / 360, threshold=150,
                            minLineLength=min_len, maxLineGap=12)
    if lines is None:
        return None
    segments = np.asarray(lines).reshape(-1, 4)
    angles: List[float] = []
    for x1, y1, x2, y2 in segments:
        if x2 == x1 and y2 == y1:
            continue
        ang = math.degrees(math.atan2(y2 - y1, x2 - x1))
        if abs(ang) <= 12:            # 近水平线族
            angles.append(ang)
        elif abs(ang) >= 78:          # 近垂直线族：归一到相对 90/-90 的偏差
            angles.append(ang - 90 if ang > 0 else ang + 90)
    if len(angles) < 4:
        return None
    return float(np.median(angles))


def _rotate(img: np.ndarray, angle: float) -> np.ndarray:
    h, w = img.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), angle, 1.0)
    return cv2.warpAffine(img, m, (w, h), flags=cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))


def _strip_dark_borders(img: np.ndarray) -> np.ndarray:
    """抹掉与图像边缘连通的深色区域（扫描件盖板黑边 / ADF 边缘暗条）。

    黑边残留会被线检测误认为表格线，或与贴边的表格线粘连成粗条。
    只处理"接触图像边界"的暗区，表格内部内容不受影响。
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    dark = (gray < 120).astype(np.uint8)
    n, labels, _, _ = cv2.connectedComponentsWithStats(dark, connectivity=4)
    if n <= 1:
        return img
    border_ids = set(np.unique(np.concatenate([
        labels[0, :], labels[-1, :], labels[:, 0], labels[:, -1],
    ])))
    border_ids.discard(0)
    if not border_ids:
        return img
    mask = np.isin(labels, list(border_ids))
    if mask.mean() > 0.35:   # "黑边"占画面 1/3 以上，多半是深色背景整图，不动
        return img
    out = img.copy()
    out[mask] = (255, 255, 255)
    return out


def _first_true_line(ys: List[int]) -> Optional[int]:
    """从横线列表中找第一条"真"表格线：与下一条的间距必须接近
    典型行距（排除扫描边缘残留、页眉装饰线等离群假线）。"""
    gaps = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
    if not gaps:
        return None
    med = median(sorted(gaps)[:max(3, len(gaps) * 3 // 4)])  # 偏小的行距更可信
    for i, y in enumerate(ys[:-1]):
        gap = ys[i + 1] - y
        if 0.45 * med <= gap <= 2.0 * med:
            return y
    return ys[0]


def preprocess(img_bgr: np.ndarray) -> np.ndarray:
    """通用前处理：限制尺寸、清理扫描边缘黑边、纠正整体小角度歪斜。

    表格识别与整页文字识别共用。
    """
    img = _limit_side(img_bgr)
    img = _strip_dark_borders(img)
    skew = _estimate_skew(img)
    if skew is not None and 0.3 <= abs(skew) <= 10.0:
        img = _rotate(img, skew)
    return img


def extract_table(img_bgr: np.ndarray) -> Optional[TableStructure]:
    """主入口：输入 BGR 图，输出表格结构；检测不到表格线时返回 None。"""
    if img_bgr is None or img_bgr.size == 0:
        return None
    img = preprocess(img_bgr)

    warped = False
    quad = find_table_quad(img)
    pre_image = None          # warp 前的转正图（标题在表格线框外，warp 会丢弃）
    if quad is not None:
        candidate = warp_perspective(img, quad)
        # 校正后的图必须明显更接近矩形表格，否则保守用原图
        if candidate.shape[0] > 60 and candidate.shape[1] > 60:
            pre_image = img
            img = candidate
            warped = True

    # 加白色边框：让表格线不贴图像边缘（贴边会导致自适应二值化不稳、
    # 边框被检出为双线、以及最外侧横/竖线漏检）
    pad = 28
    img = cv2.copyMakeBorder(img, pad, pad, pad, pad,
                             cv2.BORDER_CONSTANT, value=(255, 255, 255))

    bin_img = _binarize(img)
    h_mask = _line_mask(bin_img, horizontal=True)
    v_mask = _line_mask(bin_img, horizontal=False)

    xs = _collect_line_positions(v_mask, horizontal=False, total=img.shape[0])
    ys = _collect_line_positions(h_mask, horizontal=True, total=img.shape[1])
    if len(xs) < 2 or len(ys) < 2:
        return None

    grid_mask = cv2.bitwise_or(h_mask, v_mask)
    cells = _contour_cells(grid_mask, xs, ys)
    # 轮廓法覆盖不完整（线断裂漏格）时退回纯网格法，保证行列完整
    expect = (len(ys) - 1) * (len(xs) - 1)
    covered = sum(c.row_span * c.col_span for c in cells)
    if covered < expect * 0.9:
        cells = _grid_cells(xs, ys)

    structure = TableStructure(
        cells=cells, n_rows=len(ys) - 1, n_cols=len(xs) - 1,
        image=img, xs=xs, ys=ys, warped=warped,
        pre_image=(pre_image if warped else img),
        quad=(quad if warped else None),
    )
    structure.build_title_zones()
    structure.overlay = _draw_overlay(img, structure)
    return structure

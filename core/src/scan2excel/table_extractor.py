# -*- coding: utf-8 -*-
"""表格框线检测与单元格切分。

针对"打印表格拍照"场景：表格带完整框线，用形态学检测横竖线切出
每个单元格，附带透视校正，容忍拍照倾斜。纯 OpenCV，不依赖深度模型。
"""
from __future__ import annotations

import math
import os
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

    # 整页 det 文本条（可选，由调用方传入 engine 时在 extract_table 内
    # 识别一次并复用：既做文字掩模辅助检线，也供上层直接分配到格子）
    det_items: Optional[List] = None

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


def imread_unicode(path: str) -> np.ndarray:
    """读取图片（兼容中文路径与 EXIF 旋转），返回 BGR 数组。

    失败时抛 ImageReadError，错误文本带真实原因（扩展名 + 底层异常），
    不再静默返回 None——「文件损坏」必须能分辨是 HEIC、超大像素还是真损坏。
    """
    _register_extra_decoders()
    ext = os.path.splitext(path)[1].lower() or "(无扩展名)"
    if ext in (".heic", ".heif") and not _heif_available:
        raise ImageReadError(
            f"无法读取图片（{ext}）：缺少 HEIC 解码组件 pillow-heif，"
            "请重拍为 JPG 或安装该组件")
    try:
        from PIL import Image, ImageOps
        img = Image.open(path)
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        return cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2BGR)
    except Exception as exc:
        raise ImageReadError(
            f"无法读取图片（{ext}）：{type(exc).__name__}: {exc}") from exc


class ImageReadError(ValueError):
    """图片解码失败，message 含真实底层原因。"""


_heif_registered = False
_heif_available = False


def _register_extra_decoders() -> None:
    """注册扩展解码器；只成功一次，缺失时静默跳过。"""
    global _heif_registered, _heif_available
    if _heif_registered:
        return
    _heif_registered = True
    try:
        import pillow_heif
        pillow_heif.register_heif_opener()
        _heif_available = True
    except ImportError:
        pass
    # 手机 2 亿像素模式超过 PIL 默认 1.79 亿像素上限会直接拒绝解码；
    # 本地离线工具处理的是用户自己的文件，解除该限制。
    try:
        from PIL import Image
        Image.MAX_IMAGE_PIXELS = None
    except Exception:
        pass


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


def _expand_quad(quad: np.ndarray, margin: float = 16.0) -> np.ndarray:
    """四角向外扩 margin 像素：透视四边形常略微切进表格，
    导致最左列的行首数字被裁掉。"""
    c = quad.mean(axis=0)
    d = quad - c
    n = np.maximum(np.linalg.norm(d, axis=1, keepdims=True), 1e-6)
    return (quad + d / n * margin).astype(np.float32)


def warp_perspective(img: np.ndarray, quad: np.ndarray) -> np.ndarray:
    (tl, tr, br, bl) = quad
    width = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
    height = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
    width, height = max(width, 32), max(height, 32)
    dst = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype=np.float32)
    m = cv2.getPerspectiveTransform(quad, dst)
    out = cv2.warpPerspective(img, m, (width, height),
                              borderMode=cv2.BORDER_REPLICATE)
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


def _darker_mask(gray: np.ndarray, contrast: int = 12) -> np.ndarray:
    """每个像素是否比左右各 3px 的均值暗（返回形状 (h, w-6)，x 偏移 +3）。"""
    g = gray.astype(np.float32)
    sides = (g[:, :-6] + g[:, 6:]) / 2.0
    return (sides - g[:, 3:-3]) > contrast


def _shear_correct(img: np.ndarray, darker: np.ndarray) -> np.ndarray:
    """按近垂直 Hough 段的中位斜率做剪切校正，扶直残余微斜的竖线。

    手机拍照经透视校正后竖线仍可能带 1~3 度的倾斜/弯曲，固定 x 列的
    检测会因此漏线；此步把这类线扶直成真正的竖直线。
    darker 传入文字已抹掉的暗度掩模——文字笔画的竖段会污染斜率投票，
    把真线剪散。
    """
    d = darker.astype(np.uint8) * 255
    h, w = d.shape
    segs = cv2.HoughLinesP(d, 1, np.pi / 720, threshold=int(h * 0.10),
                           minLineLength=int(h * 0.18), maxLineGap=int(h * 0.04))
    if segs is None:
        return img
    slopes = []
    for x1, y1, x2, y2 in np.asarray(segs).reshape(-1, 4):
        if abs(x2 - x1) <= h * 0.12 and abs(y2 - y1) >= h * 0.18:
            slopes.append((x2 - x1) / float(y2 - y1 + 1e-6))
    if len(slopes) < 3:
        return img
    slope = float(np.median(slopes))
    if not (0.012 <= abs(slope) <= 0.06):
        return img
    # 质量门控：只有剪切后"竖向浓度"明显提升才动手——文字行驱动的
    # 假斜率会把真线剪散（浓度=各列暗像素数的平方和，线越集中越大）
    def concentration(mat):
        counts = mat.sum(axis=0, dtype=np.float64)
        return float((counts ** 2).sum())
    before = concentration(darker)
    h_, w_ = img.shape[:2]
    extra = int(abs(slope) * h_) + 2
    m = np.float32([[1.0, -slope, slope * h_ / 2.0 + extra / 2.0],
                    [0.0, 1.0, 0.0]])
    sheared = cv2.warpAffine(d, m, (w_ + extra, h_),
                             flags=cv2.INTER_NEAREST,
                             borderMode=cv2.BORDER_CONSTANT, borderValue=0)
    after = concentration(sheared)
    if after < before * 1.10:
        return img
    return cv2.warpAffine(img, m, (w_ + extra, h_), flags=cv2.INTER_LINEAR,
                          borderMode=cv2.BORDER_REPLICATE)


def _max_run(col: np.ndarray) -> int:
    idx = np.flatnonzero(col)
    if idx.size == 0:
        return 0
    segs = np.split(idx, np.flatnonzero(np.diff(idx) > 1) + 1)
    return max(len(s) for s in segs)


def _collect_v_lines(gray: np.ndarray, h_lines: List[int],
                     contrast: int = 12, gap: int = 25,
                     min_cov: float = 0.12, min_run_frac: float = 0.15) -> List[int]:
    """弱竖线检测（灰度图包装版，供调试）：见 _v_axes_from_mask。"""
    darker = _darker_mask(gray, contrast)
    if h_lines:
        m = np.zeros(darker.shape[0], bool)
        for y in h_lines:
            m[max(0, y - 2):y + 3] = True
        darker &= ~m[:, None]
    return _v_axes_from_mask(darker, gap, min_cov, min_run_frac)


def _v_axes_from_mask(darker: np.ndarray, gap: int = 25,
                      min_cov: float = 0.12,
                      min_run_frac: float = 0.15) -> List[int]:
    """从"比两侧暗"掩模里检出竖直线 x 坐标：累计覆盖率 + 间隙闭合
    后的最长连续段双重门槛，并做近距去重（双边界/阴影线）。"""
    closed = cv2.morphologyEx(darker.astype(np.uint8), cv2.MORPH_CLOSE,
                              np.ones((gap, 1), np.uint8))
    cov = closed.mean(axis=0)
    min_run = min_run_frac * darker.shape[0]
    cand = cov >= min_cov
    xs, covs = [], []
    x = 0
    n = len(cand)
    while x < n:
        if cand[x]:
            j = x
            while j + 1 < n and cand[j + 1]:
                j += 1
            if j - x > 14:      # 组过宽 = 斜线/弥散噪声，不是竖线
                x = j + 1
                continue
            col = closed[:, x:j + 1].max(axis=1)
            if _max_run(col) >= min_run:
                xs.append((x + j) // 2 + 3)   # 抵消 _darker_mask 的 +3 偏移
                covs.append(float(cov[x:j + 1].max()))
            x = j + 1
        else:
            x += 1
    # 近距去重：右边界与页面阴影线常被检成两条（相距几十像素），
    # 保留覆盖率高的那条。正常表格列不可能窄于图宽的 2%。
    min_gap = max(30, int(darker.shape[1] * 0.025))
    out_x, out_c = [], []
    for cx, cv_ in zip(xs, covs):
        if out_x and cx - out_x[-1] < min_gap:
            if cv_ > out_c[-1]:
                out_x[-1], out_c[-1] = cx, cv_
        else:
            out_x.append(cx)
            out_c.append(cv_)
    return out_x


def _synth_grid(shape: Tuple[int, int], xs: List[int], ys: List[int]) -> np.ndarray:
    """按检出的线坐标合成网格掩模（弱线也能参与轮廓切格）。"""
    m = np.zeros(shape, np.uint8)
    for y in ys:
        y0, y1 = max(0, y - 1), min(shape[0], y + 2)
        m[y0:y1, :] = 255
    for x in xs:
        x0, x1 = max(0, x - 1), min(shape[1], x + 2)
        m[:, x0:x1] = 255
    return m


def extract_table(img_bgr: np.ndarray,
                  engine: Optional[object] = None) -> Optional[TableStructure]:
    """主入口：输入 BGR 图，输出表格结构；检测不到表格线时返回 None。

    engine（OcrEngine）可选：传入时对处理图做一次整页 det，文本框用于
    构造文字掩模（把文字抹掉后再检线，从根上排除文字笔画链的假线），
    det 结果存入 structure.det_items 供上层复用，避免重复识别。"""
    if img_bgr is None or img_bgr.size == 0:
        return None
    img = preprocess(img_bgr)

    warped = False
    quad = find_table_quad(img)
    pre_image = None          # warp 前的转正图（标题在表格线框外，warp 会丢弃）
    expanded = 0              # 透视外扩量（外扩会带进纸张边缘台阶，检测后需过滤）
    if quad is not None:
        expanded = 16
        candidate = warp_perspective(img, _expand_quad(quad, margin=expanded))
        # 校正后的图必须明显更接近矩形表格，否则保守用原图
        if candidate.shape[0] > 60 and candidate.shape[1] > 60:
            pre_image = img
            img = candidate
            warped = True

    # 剪切校正：斜率只在"文字已抹掉"的暗度掩模上估计（文字笔画的
    # 竖段会污染斜率投票，把真线剪散），门控通过才真正剪切
    gray0 = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    darker_all0 = _darker_mask(gray0, contrast=10)    # 宽度 = w-6（两侧各去 3px）
    img = _shear_correct(img, darker_all0)

    # 加白色边框：让表格线不贴图像边缘（贴边会导致自适应二值化不稳、
    # 边框被检出为双线、以及最外侧横/竖线漏检）
    pad = 28
    img = cv2.copyMakeBorder(img, pad, pad, pad, pad,
                             cv2.BORDER_CONSTANT, value=(255, 255, 255))
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    bin_img = _binarize(img)

    # 整页 det：文本框用于交叉否决（横跨候选线的中文文本 ≥2 ⇒ 该处
    # 不存在竖线），结果存入结构供上层复用，避免重复识别
    det_items = None
    if engine is not None:
        det_items = engine.recognize_full(img)

    ys = _collect_line_positions(_line_mask(bin_img, horizontal=True),
                                 horizontal=True, total=img.shape[1])

    # 竖线：形态学优先（印刷清晰/合成图最稳）。若结果疑似漏线
    # （列间距中出现 ≥2.8 倍于中位间距的大空档），切换到弱线检测器
    # （背景归一化消光照渐变 + 覆盖率/连续段容忍断续浅印），两套取长。
    def _uniform(axes):
        if len(axes) < 3:
            return False
        gaps_ = [axes[i + 1] - axes[i] for i in range(len(axes) - 1)]
        med_ = median(sorted(gaps_)[:max(1, len(gaps_) // 2)])
        return not (med_ > 0 and max(gaps_) >= med_ * 2.8)

    v_mask = _line_mask(bin_img, horizontal=False)
    xs = _collect_line_positions(v_mask, horizontal=False, total=img.shape[0])
    xs = [x for x in xs if img.shape[1] * 0.01 <= x <= img.shape[1] * 0.99]
    if len(xs) < 5 or not _uniform(xs):
        g8 = gray.astype(np.uint8)
        bg = cv2.morphologyEx(g8, cv2.MORPH_CLOSE,
                              cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (51, 51)))
        norm = cv2.divide(g8, bg, scale=255).astype(np.float32)
        darker = _darker_mask(norm, contrast=10)      # 宽度 = w-6（两侧各去 3px）
        xs2 = _v_axes_from_mask(darker, gap=25)
        xs2 = [x for x in xs2 if img.shape[1] * 0.02 <= x <= img.shape[1] * 0.98]
        merged = sorted(set(xs) | set(xs2))
        dedup = []
        for a in merged:
            if dedup and a - dedup[-1] < 20:
                continue
            dedup.append(a)
        xs = dedup

    # 中文文本穿越否决：横跨候选线的中文文本框 ≥2 个（不同行次的长名
    # 被同一 x 贯穿）⇒ 该处照片里不存在竖线，判为文字笔画链。
    # 数字右对齐紧贴分隔线属正常溢出，不计入；单个框（标题/组表头/
    # 超长名）也不否决——合并单元格本就允许文字跨列。
    if det_items:
        def has_cjk(t):
            return any('一' <= ch <= '鿿' for ch in t)

        def iou(b1, b2):
            ix0, iy0 = max(b1[0], b2[0]), max(b1[1], b2[1])
            ix1, iy1 = min(b1[2], b2[2]), min(b1[3], b2[3])
            if ix1 <= ix0 or iy1 <= iy0:
                return 0.0
            inter = (ix1 - ix0) * (iy1 - iy0)
            u1 = (b1[2] - b1[0]) * (b1[3] - b1[1])
            u2 = (b2[2] - b2[0]) * (b2[3] - b2[1])
            return inter / float(u1 + u2 - inter)

        boxes = []
        for b, t, _s in det_items:
            p = np.asarray(b)
            boxes.append((int(p[:, 0].min()), int(p[:, 1].min()),
                          int(p[:, 0].max()), int(p[:, 1].max()), has_cjk(str(t))))

        kept = []
        for x in xs:
            # 深穿越（两侧各越出 6px）的中文框：数字右对齐紧贴分隔线的
            # 1~3px 溢出不算；同一文字的重复 det 框按 IoU 去重只算一次；
            # ≥2 条不同中文文字横跨 ⇒ 该处照片里不存在竖线
            cross = [bb for bb in boxes
                     if bb[4] and bb[0] < x - 6 and bb[2] > x + 6]
            uniq = []
            for bb in cross:
                if not any(iou(bb, u) > 0.4 for u in uniq):
                    uniq.append(bb)
            if len(uniq) >= 4:
                continue
            kept.append(x)
        xs = kept

    # 合理性门槛：横竖线至少各 3 条才像表格（文字页/噪声页常凑出
    # 1x1、2x2 的假网格）；不足则交还上层转文字模式
    if expanded:
        # 外扩带来的纸张边缘台阶在贴框处被误检成线：过滤 pad 附近的伪线
        # （真表格边框在 pad+expanded 处，不会被误伤）
        lim = pad + max(4, expanded // 2)
        xs = [x for x in xs if lim < x < img.shape[1] - lim]
        ys = [y for y in ys if lim < y < img.shape[0] - lim]
    if len(xs) < 3 or len(ys) < 3:
        return None

    # 按检出坐标合成网格掩模，弱线同样参与轮廓切格
    # 轮廓切格优先用原始线掩模（保留合并单元格的"无线"信息）；
    # 覆盖不足（线断裂漏格）再用合成网格兜底，最后退回纯网格法
    expect = (len(ys) - 1) * (len(xs) - 1)
    real_mask = cv2.bitwise_or(_line_mask(bin_img, True),
                               _line_mask(bin_img, False))
    cells = _contour_cells(real_mask, xs, ys)
    covered = sum(c.row_span * c.col_span for c in cells)
    if covered < expect * 0.9:
        grid_mask = _synth_grid(img.shape[:2], xs, ys)
        cells2 = _contour_cells(grid_mask, xs, ys)
        covered2 = sum(c.row_span * c.col_span for c in cells2)
        cells = cells2 if covered2 > covered else cells
        covered = max(covered, covered2)
    if covered < expect * 0.9:
        cells = _grid_cells(xs, ys)

    structure = TableStructure(
        cells=cells, n_rows=len(ys) - 1, n_cols=len(xs) - 1,
        image=img, xs=xs, ys=ys, warped=warped,
        pre_image=(pre_image if warped else img),
        quad=(quad if warped else None),
        det_items=det_items,
    )
    structure.build_title_zones()
    structure.overlay = _draw_overlay(img, structure)
    return structure

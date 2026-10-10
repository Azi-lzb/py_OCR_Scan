# -*- coding: utf-8 -*-
"""拍照图片 → 扫描件（扫描全能王式增强）。

管线参考主流文档扫描工具（CamScanner / OpenScan 的经典套路）：
方向转正（上层复用整页探测）→ 表格四边形透视拉平（外扩外框把
表头/标题留在画内）→ 背景归一化（除以大核模糊，消阴影/光照渐变）
→ 按模式出图：
  enhanced 彩色增强：白化系数作用到彩色三通道（消阴影保色）+ L 通道
           拉伸提对比 + 轻锐化——对应扫描全能王的「增强并锐化」
  bw       黑白：归一化图自适应二值化，纸面纯白、笔画黑、中值去噪
  gray     灰度：归一化后按 1%/99% 分位线性拉伸
  origin   只拉平不增强
"""
from __future__ import annotations

import cv2
import numpy as np

from .table_extractor import (_expand_quad, find_table_quad, preprocess,
                              warp_perspective)

SCAN_MODES = ("enhanced", "bw", "gray", "origin")
SCAN_MODE_NAMES = {"enhanced": "彩色增强", "bw": "黑白", "gray": "灰度",
                   "origin": "原图拉平"}


def _paper_quad(proc: np.ndarray):
    """纸张轮廓检测：亮纸面 vs 深色桌面，取最大四边形（主流文档扫描
    工具的做法）。斜着摆放/带透视的纸张也能找到边界；找不到返回 None
    （交给上层退回表格四边形）。"""
    h, w = proc.shape[:2]
    gray = cv2.cvtColor(proc, cv2.COLOR_BGR2GRAY)
    # 纸面亮、背景暗：用大核模糊后的自适应边缘把纸轮廓描出来
    blur = cv2.GaussianBlur(gray, (0, 0), 3)
    _, mask = cv2.threshold(blur, 0, 255,
                            cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    mask = cv2.morphologyEx(
        mask, cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (max(9, w // 60),) * 2))
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts:
        return None
    best, best_area = None, 0.0
    for c in cnts:
        area = cv2.contourArea(c)
        if area < h * w * 0.25:          # 纸张至少占画面 1/4
            continue
        if area > best_area:
            best, best_area = c, area
    if best is None:
        return None
    for eps in (0.02, 0.05, 0.08):       # 由紧到松逼近四边形
        approx = cv2.approxPolyDP(best, eps * cv2.arcLength(best, True), True)
        if len(approx) == 4:
            pts = approx.reshape(4, 2).astype(np.float32)
            return [(float(p[0]), float(p[1])) for p in pts]
    return None


def flatten_page(img_bgr: np.ndarray) -> np.ndarray:
    """透视拉平：优先纸张轮廓四边形（斜摆/带透视的纸张也适用），
    退回表格四边形（上边缘沿纸张亮度扩到纸界保住标题）。"""
    proc = preprocess(img_bgr)
    h, w = proc.shape[:2]
    quad = _paper_quad(proc)
    if quad is not None:
        pts = _order_quad_np(quad)
        margin = max(12, int(max(h, w) * 0.008))
        src = np.float32([pts[0], pts[1], pts[2], pts[3]])
        wid = int(max(np.linalg.norm(src[1] - src[0]),
                      np.linalg.norm(src[2] - src[3])))
        hei = int(max(np.linalg.norm(src[3] - src[0]),
                      np.linalg.norm(src[2] - src[1])))
        if wid >= 200 and hei >= 200:
            dst = np.float32([[0, 0], [wid, 0], [wid, hei], [0, hei]])
            return cv2.warpPerspective(
                proc, cv2.getPerspectiveTransform(src, dst), (wid, hei))
    # 退回：表格四边形（识别管线同款），上边缘按亮度扩到纸界保标题
    quad = find_table_quad(proc)
    if quad is None:
        return proc
    gray = cv2.cvtColor(proc, cv2.COLOR_BGR2GRAY)
    q = np.asarray(quad, dtype=np.float32)
    x0 = max(0, int(q[:, 0].min()))
    x1 = min(w, int(q[:, 0].max()))
    top = int(min(q[0][1], q[1][1]))
    if top >= 14 and x1 > x0:
        bright = (gray[:, x0:x1] > 165).mean(axis=1)
        y = top - 4
        while y > 0 and bright[y] > 0.25:
            y -= 1
        lift = float(top - max(0, y + 2))
        if lift > 0:
            q = q.copy()
            q[0][1] = max(0.0, q[0][1] - lift)
            q[1][1] = max(0.0, q[1][1] - lift)
    margin = max(16, int(max(h, w) * 0.012))
    return warp_perspective(proc, _expand_quad(q, margin=margin))


def _order_quad_np(quad):
    """四点排序为 左上/右上/右下/左下（按 x+y 与 x-y）。"""
    pts = np.asarray(quad, dtype=np.float32)
    s = pts.sum(axis=1)
    d = np.diff(pts, axis=1).ravel()
    return [pts[np.argmin(s)], pts[np.argmin(d)],
            pts[np.argmax(s)], pts[np.argmax(d)]]


def _bg_normalized(gray: np.ndarray) -> np.ndarray:
    """打灯式背景估计：膨胀去字 + 大核中值得到"无字纸面"光照图，
    再相除——阴影/光照渐变被整体抬平（等效把灯打均匀后扫描）。"""
    k = int(max(21, min(61, (min(gray.shape[:2]) // 40) | 1)))
    if k % 2 == 0:
        k += 1
    dil = cv2.dilate(gray, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5)))
    bg = cv2.medianBlur(dil, k)
    return cv2.divide(gray, bg, scale=255)


def enhance_scan(flat: np.ndarray, mode: str = "enhanced") -> np.ndarray:
    """对已拉平的页面做扫描件增强。mode ∈ SCAN_MODES。"""
    if mode not in SCAN_MODES:
        mode = "enhanced"
    if mode == "origin":
        return flat
    gray = cv2.cvtColor(flat, cv2.COLOR_BGR2GRAY)
    norm = _bg_normalized(gray)                  # 打灯式去阴影 → 白底
    if mode == "bw":
        bw = cv2.adaptiveThreshold(norm, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 31, 10)
        bw = cv2.medianBlur(bw, 3)               # 去孤立噪点
        return cv2.cvtColor(bw, cv2.COLOR_GRAY2BGR)
    if mode == "gray":
        lo, hi = np.percentile(norm, 1), np.percentile(norm, 99)
        g = np.clip((norm - lo) * 255.0 / max(1.0, float(hi - lo)),
                    0, 255).astype(np.uint8)
        return cv2.cvtColor(g, cv2.COLOR_GRAY2BGR)
    # enhanced：白化系数作用到彩色三通道（消阴影同时保色），再提对比 + 轻锐化
    f = np.clip(norm.astype(np.float32) / np.maximum(gray.astype(np.float32), 1.0),
                0.0, 4.0)
    out = np.clip(flat.astype(np.float32) * f[..., None], 0, 255).astype(np.uint8)
    lab = cv2.cvtColor(out, cv2.COLOR_BGR2LAB)
    l_ch, a_ch, b_ch = cv2.split(lab)
    lo, hi = np.percentile(l_ch, 1), np.percentile(l_ch, 99)
    l_ch = np.clip((l_ch.astype(np.float32) - lo) * 255.0
                   / max(1.0, float(hi - lo)), 0, 255).astype(np.uint8)
    out = cv2.cvtColor(cv2.merge([l_ch, a_ch, b_ch]), cv2.COLOR_LAB2BGR)
    soft = cv2.GaussianBlur(out, (0, 0), 2.0)
    return cv2.addWeighted(out, 1.35, soft, -0.35, 0)

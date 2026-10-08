# -*- coding: utf-8 -*-
"""RapidOCR 封装：懒加载单例 + 单元格小图识别。

rapidocr-onnxruntime 的模型打包在 wheel 内，加载即离线可用。
"""
from __future__ import annotations

import sys
import threading
from pathlib import Path
from typing import List, Optional, Tuple

import cv2
import numpy as np

# 格子小图放大到的最小高度（短边太小会明显掉识别率）
MIN_CELL_HEIGHT = 44
# 整图识别时，短边低于该值的图先放大再识别（低分辨率扫描件里
# 文字行高常不足 15px，det 找框与 rec 识别都会退化，典型症状是
# 连字符"-"被误读成"1"）
MIN_FULL_SIDE = 1000
UPSCALE = 1.5
# 前景像素占比低于该值视为空格子（拍照件有阴影噪声，阈值需容忍
# 少量噪点，否则大片空格触发逐格补漏，整表耗时暴涨）
EMPTY_FILL_RATIO = 0.015

# 高精度档：server rec 模型（识别更准，速度约慢 2~3 倍）。
# 本地不存在时从 ModelScope 下载一次（约 85MB），之后离线可用。
SERVER_REC_URL = ("https://www.modelscope.cn/models/RapidAI/RapidOCR/"
                  "resolve/v3.3.0/onnx/PP-OCRv4/rec/ch_PP-OCRv4_rec_server_infer.onnx")
_SERVER_REC_NAME = "ch_PP-OCRv4_rec_server_infer.onnx"


class OcrEngine:
    _instances: dict = {}
    _lock = threading.Lock()

    @classmethod
    def instance(cls, server_rec: bool = False) -> "OcrEngine":
        """获取 OCR 引擎单例。server_rec=True 返回高精度档实例。"""
        key = "server" if server_rec else "mobile"
        if key not in cls._instances:
            with cls._lock:
                if key not in cls._instances:
                    cls._instances[key] = cls(server_rec=server_rec)
        return cls._instances[key]

    def __init__(self, server_rec: bool = False) -> None:
        from rapidocr_onnxruntime import RapidOCR  # 延迟导入，加快启动
        self.server_rec = server_rec
        kwargs = {}
        if server_rec:
            path = ensure_server_rec_model()
            if path:
                kwargs["rec_model_path"] = str(path)
            else:
                # 模型不可得：静默降级 mobile，上层已提示
                self.server_rec = False
        self._engine = RapidOCR(**kwargs)

    @staticmethod
    def is_server_ready() -> bool:
        return ensure_server_rec_model(download=False) is not None

    @staticmethod
    def download_server_model() -> bool:
        return ensure_server_rec_model(download=True) is not None

    # ------------------------------------------------------------------
    def recognize_full(self, img: np.ndarray, use_cls: bool = True,
                       upscale: bool = True) -> List:
        """整表识别一次（det+cls+rec），返回 [[box, text, score], ...]。

        det 阶段会精确裁出每条文字，识别质量显著高于把整个格子塞给
        rec；一次调用覆盖全图所有文字，供上层按坐标分配到单元格。
        低分辨率输入先放大（box 坐标同步换算回原图坐标系）。
        use_cls=False 用于方向探测（倒立文字分数会明显下降）。
        """
        scale = 1.0
        h, w = img.shape[:2]
        if upscale and min(h, w) < MIN_FULL_SIDE:
            scale = UPSCALE
            img = cv2.resize(img, (int(w * scale), int(h * scale)),
                             interpolation=cv2.INTER_CUBIC)
        result, _ = self._engine(img, use_cls=use_cls)
        if not result or scale == 1.0:
            return [item for item in (result or []) if str(item[1]).strip()]
        out = []
        for box, text, score in result:
            if not str(text).strip():
                continue
            out.append([[[p / scale for p in pt] for pt in box], text, score])
        return out

    def recognize_cell_numeric(self, cell_img: np.ndarray):
        """数值格专用识别：不做线消除（数字+千分位逗号构成的横向条带会被
        误判成线而抹掉）、直接整行 rec（数字单行识别的最优路径），
        去除拆散空格。用于模板模式的数值列，快且准。"""
        if cell_img is None or cell_img.size == 0:
            return "", 1.0
        h, w = cell_img.shape[:2]
        pad = 4
        if h > 2 * pad + 1 and w > 2 * pad + 1:
            cell_img = cell_img[pad:h - pad, pad:w - pad]
        if self._is_blank(cell_img):
            return "", 1.0
        h, w = cell_img.shape[:2]
        if h < 52:
            sc = 52.0 / h
            cell_img = cv2.resize(cell_img, (max(1, int(w * sc)), 52),
                                  interpolation=cv2.INTER_CUBIC)
        result, _ = self._engine(cell_img, use_det=False, use_cls=False,
                                 use_rec=True)
        if not result:
            return "", 1.0
        text = "".join(str(t) for t, _s in result if str(t).strip())
        scores = [float(s) for _t, s in result
                  if str(_t).strip() and s is not None]
        # 清除拆散空格与全角逗号；保留半角逗号/小数点
        text = text.replace(" ", "").replace("　", "").replace("，", ",")
        return text.strip(), (min(scores) if scores else 1.0)

    def is_blank(self, cell_img: np.ndarray) -> bool:
        return self._is_blank(cell_img)

    def recognize_cell(self, cell_img: np.ndarray) -> Tuple[str, float]:
        """补漏路径：单个格子完整模式识别（det+cls+rec）。"""
        if cell_img is None or cell_img.size == 0:
            return "", 1.0
        h, w = cell_img.shape[:2]
        pad = 3
        if h > 2 * pad + 1 and w > 2 * pad + 1:
            cell_img = cell_img[pad:h - pad, pad:w - pad]
            h, w = cell_img.shape[:2]
        if self._is_blank(cell_img):
            return "", 1.0
        cell_img = self._remove_lines(cell_img)
        if h < MIN_CELL_HEIGHT:
            scale = MIN_CELL_HEIGHT / h
            cell_img = cv2.resize(cell_img, (max(1, int(w * scale)), MIN_CELL_HEIGHT),
                                  interpolation=cv2.INTER_CUBIC)
        result, _ = self._engine(cell_img)
        return self._join(result)

    # ------------------------------------------------------------------
    @staticmethod
    def _remove_lines(img: np.ndarray) -> np.ndarray:
        """抹掉格子内残留的横/竖线段，防止被识别成'電/時/間'等字形。

        只删除长度接近贯穿格子的直线（表格残线的特征），并以前景损失率
        兜底——若"删线"吃掉过多前景像素，说明误伤了文字笔画，放弃处理。
        """
        h, w = img.shape[:2]
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        bin_img = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                        cv2.THRESH_BINARY_INV, 15, 12)
        # 贯穿性长线：横线超过半宽、竖线超过半高；汉字笔画达不到这个长度
        h_k = max(15, int(w * 0.5))
        v_k = max(15, int(h * 0.5))
        h_mask = cv2.morphologyEx(bin_img, cv2.MORPH_OPEN,
                                  cv2.getStructuringElement(cv2.MORPH_RECT, (h_k, 1)))
        v_mask = cv2.morphologyEx(bin_img, cv2.MORPH_OPEN,
                                  cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_k)))
        lines = cv2.dilate(cv2.bitwise_or(h_mask, v_mask), np.ones((3, 3), np.uint8))
        n_lines = np.count_nonzero(lines)
        n_fg = np.count_nonzero(bin_img)
        if n_lines == 0:
            return img
        # 残线只占前景的一小部分；吃掉超过 40% 前景说明误伤文字，回退原图
        if n_fg == 0 or n_lines > n_fg * 0.4:
            return img
        out = img.copy()
        out[lines > 0] = (255, 255, 255)
        return out

    @staticmethod
    def _is_blank(img: np.ndarray) -> bool:
        # 裁掉四周边框线（补漏传入的格子裁切含边框，是强前景，
        # 不裁的话所有格子都被判"非空"，补漏风暴）
        h, w = img.shape[:2]
        if h > 8 and w > 8:
            img = img[3:h - 3, 3:w - 3]
        # 背景归一化后再二值：拍照件的灰色阴影格归一化后接近白纸，
        # 不会被误判为有内容
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        bg = cv2.morphologyEx(gray, cv2.MORPH_CLOSE,
                              cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (25, 25)))
        norm = cv2.divide(gray, bg, scale=255)
        # 固定阈值而非 Otsu：平坦的阴影格（灰度起伏仅十几级）会让 Otsu
        # 退化到把噪声分半；归一化后真实笔墨必然显著暗于背景
        bin_img = (norm < 200).astype(np.uint8) * 255
        return float(np.count_nonzero(bin_img)) / bin_img.size < EMPTY_FILL_RATIO

    @staticmethod
    def _join(items) -> Tuple[str, float]:
        """把 [[box, text, score], ...] 按位置排序拼接为一段文本。

        rec-only 模式（use_det=False）返回的 box 为 None：条目按原顺序拼接。
        """
        if not items:
            return "", 1.0
        if len(items[0]) == 2:   # rec-only 模式返回 [text, score]
            texts, scores = [], []
            for text, score in items:
                t = str(text).strip()
                if t:
                    texts.append(t)
                    scores.append(float(score))
            return "".join(texts), (min(scores) if scores else 1.0)
        rows = []
        for box, text, score in items:
            ys = [p[1] for p in box]
            xs = [p[0] for p in box]
            rows.append((sum(ys) / len(ys), sum(xs) / len(xs), str(text), float(score)))
        rows.sort(key=lambda r: (r[0], r[1]))
        # 按 y 聚成行（阈值取行高的 60%）
        lines: List[List] = []
        for item in rows:
            if lines and abs(item[0] - lines[-1][0][0]) < max(6.0, (item[0] - lines[-1][0][0]) * 0.6):
                lines[-1].append(item)
            else:
                lines.append([item])
        texts = []
        for line in lines:
            line.sort(key=lambda r: r[1])
            texts.append(" ".join(t for _, _, t, _ in line if t))
        text = " ".join(t for t in texts if t).replace("\u3000", " ").strip()
        scores = [s for _, _, _, s in rows if s > 0]
        return text, (min(scores) if scores else 1.0)


def _bundled_server_rec() -> Optional[Path]:
    """随程序分发的内置模型：打包态在 _MEIPASS/models，开发态在包内 models/。"""
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(Path(getattr(sys, "_MEIPASS", ".")) / "models" / _SERVER_REC_NAME)
    candidates.append(Path(__file__).resolve().parent / "models" / _SERVER_REC_NAME)
    for cand in candidates:
        try:
            if cand.is_file() and cand.stat().st_size > 50 * 1024 * 1024:
                return cand
        except OSError:
            continue
    return None


def ensure_server_rec_model(download: bool = True) -> Optional[Path]:
    """确保高精度 server rec 模型就位，返回模型路径；不可得返回 None。

    查找顺序：用户目录缓存 → 程序内置（打包/开发均随包分发，离线可用）
    → 联网下载（仅程序未内置时的兜底，约 85MB，下载到用户目录）。
    """
    model_dir = Path.home() / ".scan2excel" / "models"
    path = model_dir / _SERVER_REC_NAME
    if path.is_file() and path.stat().st_size > 10 * 1024 * 1024:
        return path
    bundled = _bundled_server_rec()
    if bundled is not None:
        return bundled
    if not download:
        return None
    try:
        import urllib.request

        model_dir.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".onnx.part")
        print(f"下载高精度识别模型（约 85MB）…")
        urllib.request.urlretrieve(SERVER_REC_URL, tmp)
        tmp.replace(path)
        return path
    except Exception as exc:  # noqa: BLE001 —— 离线环境静默失败
        print(f"高精度模型下载失败：{exc}")
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
        return None

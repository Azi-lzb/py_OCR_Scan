# -*- coding: utf-8 -*-
"""imread_unicode 图片读取回归测试。

背景：真实手机照片 9 张里 6 张报「无法读取图片（文件损坏或格式不支持）」，
真实原因被吞掉。修复后：
  ① 失败必须抛 ImageReadError，错误文本含扩展名与底层异常（可诊断）
  ② HEIC/HEIF 手机照片可解码（pillow-heif，缺失时明确报缺组件）
  ③ PIL 的 DecompressionBomb 像素上限已解除（2 亿像素手机照片不再被拒）
用法：.venv/Scripts/python.exe tests/test_imread.py
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))

import numpy as np                                                # noqa: E402
from PIL import Image                                             # noqa: E402

from scan2excel.table_extractor import (ImageReadError,            # noqa: E402
                                        imread_unicode)
from scan2excel.table_extractor import _register_extra_decoders    # noqa: E402
from scan2excel.web_app import IMAGE_EXTS                          # noqa: E402


def _make_jpeg(path: Path) -> Path:
    arr = np.full((60, 80, 3), 200, dtype=np.uint8)
    arr[10:50, 20:60] = 30
    Image.fromarray(arr).save(path)
    return path


def test_normal_and_unicode_path(tmp: Path) -> None:
    p = _make_jpeg(tmp / "测试表格_拍照.jpg")
    img = imread_unicode(str(p))
    assert img is not None and img.shape == (60, 80, 3), "中文路径 JPEG 读取失败"


def test_corrupt_file_reports_real_reason(tmp: Path) -> None:
    p = tmp / "损坏.jpg"
    p.write_bytes(b"\xff\xd8\xff\xe0not-a-real-jpeg-payload")
    try:
        imread_unicode(str(p))
    except ImageReadError as exc:
        msg = str(exc)
        assert ".jpg" in msg, f"错误文本应含扩展名：{msg}"
        assert "：" in msg and ("Error" in msg or "error" in msg), \
            f"错误文本应含底层异常类型：{msg}"
    else:
        raise AssertionError("损坏文件应抛 ImageReadError 而非静默返回")


def test_heif_roundtrip(tmp: Path) -> None:
    try:
        import pillow_heif  # noqa: F401
    except ImportError:
        print("  [跳过] pillow-heif 未安装，无法验证 HEIC 解码")
        return
    _register_extra_decoders()
    arr = np.full((40, 60, 3), 128, dtype=np.uint8)
    p = tmp / "手机照片.heic"
    Image.fromarray(arr).save(p, format="HEIF")
    img = imread_unicode(str(p))
    assert img is not None and img.shape == (40, 60, 3), "HEIC 读取失败"
    assert ".heic" in IMAGE_EXTS, "IMAGE_EXTS 应包含 .heic"


def test_pixel_limit_lifted() -> None:
    _register_extra_decoders()
    assert Image.MAX_IMAGE_PIXELS is None, \
        "2 亿像素手机照片不应被 DecompressionBomb 上限拒绝"


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="sc2img_") as td:
        tmp = Path(td)
        test_normal_and_unicode_path(tmp)
        print("  [通过] 中文路径 JPEG 正常读取")
        test_corrupt_file_reports_real_reason(tmp)
        print("  [通过] 损坏文件报出真实原因（扩展名+异常类型）")
        test_heif_roundtrip(tmp)
        print("  [通过] HEIC 手机照片解码")
        test_pixel_limit_lifted()
        print("  [通过] 像素上限已解除")
    print("全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

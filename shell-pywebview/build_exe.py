# -*- coding: utf-8 -*-
"""pywebview 外壳打包（PyInstaller onedir）。

用法：python build_exe.py [--onefile]
产物：dist/OCR_Tools/ —— 文件夹版（推荐）：启动免解压、杀软误报低；
    OCR_Tools.exe + _internal（依赖与模型）+ core/data（config/模板种子）。
    core/data 里的 config.xlsx 与 国库报表.xlsx 是种子：打包时"不存在才复制"，
    用户后续修改不会被重建覆盖；加 --reset-data 可强制重置为打包内置版本。
--onefile 仍可打单文件版（启动慢、误报高，备用）。
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CORE = ROOT.parent / "core"
if not (CORE / "src" / "scan2excel").is_dir():
    raise SystemExit(f"[错误] 未找到业务核心：{CORE}\\src\\scan2excel")
if not (CORE / "frontend" / "web" / "index.html").is_file():
    raise SystemExit("[错误] 未找到前端 index.html")
NAME = "OCR_Tools"
SEP = ";" if sys.platform == "win32" else ":"

# 打包内置的种子数据（相对 dist 应用目录）：不存在才复制，不覆盖用户修改
SEED_FILES = [
    (CORE / "data" / "config" / "config.xlsx",          "core/data/config/config.xlsx"),
    (CORE / "data" / "transforms" / "国库报表.xlsx",     "core/data/transforms/国库报表.xlsx"),
    (CORE / "data" / "templates" / "Template.xlsx",     "core/data/templates/Template.xlsx"),
]


def seed_data(app_dir: Path, reset: bool = False) -> None:
    """把种子文件复制进发行目录 core/data（已存在则跳过，保留用户修改）。"""
    for src, rel in SEED_FILES:
        if not src.is_file():
            continue
        dst = app_dir / rel
        if dst.exists() and not reset:
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        print(f"  种子: {rel}")


def build() -> int:
    onefile = "--onefile" in sys.argv
    args = [
        "--onedir" if not onefile else "--onefile", "--noconsole", "--clean",
        "--noconfirm",                   # 覆盖旧 dist 不询问，一路打到底
        "--name", NAME,
        "--paths", str(CORE / "src"),
        # 冻结态前端页面从 _MEIPASS/web/index.html 读取
        "--add-data", f"{CORE / 'frontend' / 'web' / 'index.html'}{SEP}web",
        # 高精度 server rec 模型随包分发（离线可用，无需首次联网下载）
        "--add-data", f"{ROOT.parent / 'core/src/scan2excel/models/ch_PP-OCRv4_rec_server_infer.onnx'}{SEP}models",
        # RapidOCR 的 .onnx 模型在包内，必须随包收集
        "--collect-all", "rapidocr_onnxruntime",
        "--collect-all", "webview",
        # 无框线表格模型（SLANet-Plus）与 PDF 渲染引擎
        "--collect-all", "rapid_table",
        "--collect-all", "pypdfium2",
        # HEIC/HEIF 手机照片解码（libheif 原生 dll 需一并收集）
        "--collect-all", "pillow_heif",
        "--hidden-import", "webview.platforms.edgechromium",
        # Tk 对话框回退必须保留：冻结态 pywebview 原生对话框若异常，
        # 回退路径 import tkinter 缺失会表现为"点选择没反应"
        "--hidden-import", "tkinter",
        # 未使用的重模块一律剔除，控制体积
        "--exclude-module", "torch",
        "--exclude-module", "pandas",
        "--exclude-module", "matplotlib",
        "--exclude-module", "PyQt5", "--exclude-module", "PySide2",
        "--exclude-module", "PyQt6", "--exclude-module", "PySide6",
        "--exclude-module", "webview.platforms.android",
        "--exclude-module", "webview.platforms.cocoa",
        "--exclude-module", "webview.platforms.gtk",
        "--exclude-module", "webview.platforms.qt",
        "--exclude-module", "pytest", "--exclude-module", "unittest",
        "--exclude-module", "pydoc", "--exclude-module", "doctest",
        # 本工具不做 COM/Excel 自动化，剔除 pywin32 全家
        "--exclude-module", "win32com",
        "--exclude-module", "pythoncom",
        "--exclude-module", "pywintypes",
        "--exclude-module", "win32api",
        "--exclude-module", "win32clipboard",
        "--exclude-module", "win32con",
        "--exclude-module", "win32print",
        "--exclude-module", "pywin32",
        "--distpath", str(ROOT / "dist"),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"),
        str(ROOT / "run.py"),
    ]
    import PyInstaller.__main__
    PyInstaller.__main__.run(args)

    # onedir：把种子数据复制进发行目录（不存在才复制，不覆盖用户修改）
    if not onefile:
        app_dir = ROOT / "dist" / NAME
        if app_dir.is_dir():
            seed_data(app_dir)
            print(f"种子数据已装配：{app_dir / 'core/data'}")
    report_size(ROOT / "dist" / NAME if not onefile else ROOT / "dist" / f"{NAME}.exe")
    return 0


def report_size(path: Path) -> None:
    if path.is_dir():
        total = sum(f.stat().st_size for f in path.rglob("*") if f.is_file())
        print(f"\n打包完成（onedir）：{path}")
        print(f"成品大小：{total / 1048576:.1f} MB")
    elif path.is_file():
        print(f"\n打包完成（onefile）：{path}")
        print(f"成品大小：{path.stat().st_size / 1048576:.1f} MB")


if __name__ == "__main__":
    sys.exit(build())

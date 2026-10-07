# -*- coding: utf-8 -*-
"""Flask 外壳打包（PyInstaller onefile）。

用法：python build_exe.py
产物：dist/ScanToExcel-Flask.exe —— 单文件；双击后启动本地服务并
自动打开浏览器（控制台窗口保留，显示服务地址，Ctrl+C 退出）。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CORE = ROOT.parent / "core"
if not (CORE / "src" / "scan2excel").is_dir():
    raise SystemExit(f"[错误] 未找到业务核心：{CORE}\\src\\scan2excel")
if not (CORE / "frontend" / "web" / "index.html").is_file():
    raise SystemExit("[错误] 未找到前端 index.html")
NAME = "ScanToExcel-Flask"
SEP = ";" if sys.platform == "win32" else ":"


def build() -> int:
    args = [
        "--onefile", "--clean",            # Flask 壳保留控制台便于看服务状态
        "--name", NAME,
        "--paths", str(CORE / "src"),
        # 冻结态前端页面从 _MEIPASS/web/index.html 读取
        "--add-data", f"{CORE / 'frontend' / 'web' / 'index.html'}{SEP}web",
        # 高精度 server rec 模型随包分发（离线可用，无需首次联网下载）
        "--add-data", f"{ROOT.parent / 'core/src/scan2excel/models/ch_PP-OCRv4_rec_server_infer.onnx'}{SEP}models",
        # RapidOCR 的 .onnx 模型在包内，必须随包收集
        "--collect-all", "rapidocr_onnxruntime",
        "--collect-all", "flask",
        # 无框线表格模型（SLANet-Plus）与 PDF 渲染引擎
        "--collect-all", "rapid_table",
        "--collect-all", "pypdfium2",
        "--hidden-import", "tkinter.filedialog",
        "--exclude-module", "torch",
        "--exclude-module", "pandas",
        "--exclude-module", "matplotlib",
        "--exclude-module", "webview",
        "--exclude-module", "PyQt5", "--exclude-module", "PySide2",
        "--exclude-module", "PyQt6", "--exclude-module", "PySide6",
        "--exclude-module", "pytest", "--exclude-module", "unittest",
        "--exclude-module", "pydoc", "--exclude-module", "doctest",
        "--distpath", str(ROOT / "dist"),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"),
        str(ROOT / "run.py"),
    ]
    import PyInstaller.__main__
    PyInstaller.__main__.run(args)
    report_size(ROOT / "dist" / f"{NAME}.exe")
    return 0


def report_size(path: Path) -> None:
    if not path.is_file():
        return
    print(f"\n打包完成：{path}")
    print(f"成品大小：{path.stat().st_size / 1048576:.1f} MB")


if __name__ == "__main__":
    sys.exit(build())

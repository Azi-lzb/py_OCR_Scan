# -*- coding: utf-8 -*-
"""Flask 外壳打包（PyInstaller onedir）。

用法：python build_exe.py
产物：dist/OCR_Tools-Flask/ —— 文件夹版（与 pywebview 壳一致）：
    OCR_Tools-Flask.exe + _internal（依赖与模型）+ core/data（config/模板种子）。
    core/data 里的 config.xlsx 与 国库报表.xlsx 是种子：打包时"不存在才复制"，
    用户后续修改不会被重建覆盖。双击 exe 启动本地服务并自动开浏览器。
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
NAME = "OCR_Tools-Flask"
SEP = ";" if sys.platform == "win32" else ":"

# 打包内置的种子数据（相对发行目录）：不存在才复制，不覆盖用户修改
SEED_FILES = [
    (CORE / "data" / "config" / "config.xlsx",          "core/data/config/config.xlsx"),
    (CORE / "data" / "transforms" / "国库报表.xlsx",     "core/data/transforms/国库报表.xlsx"),
    (CORE / "data" / "templates" / "Template.xlsx",     "core/data/templates/Template.xlsx"),
]


def seed_data(app_dir: Path) -> None:
    for src, rel in SEED_FILES:
        if not src.is_file():
            continue
        dst = app_dir / rel
        if dst.exists():
            continue                    # 保留用户修改
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        print(f"  种子: {rel}")


def build() -> int:
    args = [
        "--onedir", "--clean", "--noconfirm",  # 覆盖旧 dist 不询问；Flask 壳保留控制台便于看服务状态
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

    # 种子装配：config/模板不存在才复制（不覆盖用户修改）
    app_dir = ROOT / "dist" / NAME
    if app_dir.is_dir():
        seed_data(app_dir)
        print(f"种子数据已装配：{app_dir / 'core/data'}")
    report_size(app_dir)
    return 0


def report_size(path: Path) -> None:
    total = 0
    if path.is_dir():
        for f in path.rglob("*"):
            if f.is_file():
                total += f.stat().st_size
    print(f"\n打包完成（onedir）：{path}")
    print(f"成品大小：{total / 1048576:.1f} MB")


if __name__ == "__main__":
    sys.exit(build())

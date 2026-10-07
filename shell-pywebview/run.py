# -*- coding: utf-8 -*-
"""pywebview 外壳启动器：定位 core/、注入 sys.path、单实例、开桌面窗口。"""
from __future__ import annotations

import multiprocessing
import sys
from pathlib import Path


def resolve_paths() -> Path:
    if getattr(sys, "frozen", False):
        # 打包态：scan2excel 包与前端页面已内嵌进 exe，无需外部 core/
        return Path(sys.executable).resolve().parent
    root = Path(__file__).resolve().parent
    core = root / "core"
    if not (core / "src" / "scan2excel").is_dir():
        core = root.parent / "core"   # 开发布局：shell 与 core 同级
    if not (core / "src" / "scan2excel").is_dir():
        raise SystemExit("未找到 core 目录，请保持目录结构完整")
    src = core / "src"
    if str(src) not in sys.path:
        sys.path.insert(0, str(src))
    return core


def main() -> int:
    core = resolve_paths()
    from scan2excel.single_instance import acquire, show_already_running_message
    if not acquire():
        show_already_running_message()
        return 0
    from scan2excel.web_app import launch_web
    launch_web(core)
    return 0


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(main())

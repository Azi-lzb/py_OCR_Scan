# -*- coding: utf-8 -*-
"""Flask 外壳启动器：端口探测 + 防重复启动 + 自动打开浏览器。"""
from __future__ import annotations

import atexit
import os
import socket
import sys
import threading
import webbrowser
from pathlib import Path

ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) \
    else Path(__file__).resolve().parent
# 打包态：exe 自身目录即工作根；开发态：core/ 与 shell 同级
if getattr(sys, "frozen", False):
    # 发行包布局：exe 旁有 core\ 就用它（数据落 core/data）；
    # 否则用 exe 自身目录（模板等数据随 exe 走）
    CORE = ROOT / "core" if (ROOT / "core").is_dir() else ROOT
elif (ROOT / "core" / "src").is_dir():
    CORE = ROOT / "core"
else:
    CORE = ROOT.parent / "core"
PID_FILE = CORE / "data" / "app.pid"
BASE_PORT = 8750


def can_bind(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


def first_available_port(base: int) -> int:
    for port in range(base, base + 20):
        if can_bind(port):
            return port
    raise RuntimeError("8750~8769 端口均被占用")


def write_pid(port: int) -> None:
    PID_FILE.parent.mkdir(parents=True, exist_ok=True)
    PID_FILE.write_text(f'{{"pid": {os.getpid()}, "port": {port}}}', encoding="utf-8")
    atexit.register(lambda: PID_FILE.exists() and PID_FILE.unlink())


def open_browser(url: str) -> None:
    webbrowser.open(url)


def main() -> int:
    no_browser = "--no-browser" in sys.argv
    port = first_available_port(BASE_PORT)
    write_pid(port)
    from app import app
    url = f"http://127.0.0.1:{port}/"
    if not no_browser:
        threading.Timer(1.0, lambda: open_browser(url)).start()
    print(f"ScanToExcel (Flask) 服务已启动：{url}")
    app.run(host="127.0.0.1", port=port, debug=False)
    return 0


if __name__ == "__main__":
    sys.exit(main())

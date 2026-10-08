# -*- coding: utf-8 -*-
"""前端 JS 语法自检：从 index.html 提取 <script> 用 node --check 校验。

改完前端必跑：Python 里写 JS 字符串极易把 \\n 转义成真换行，
导致整段脚本静默不执行（页面点不动、S 未定义等）。
用法：.venv/Scripts/python.exe tests/check_frontend_js.py
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent  # tests/ 的上一级即项目根
HTML = ROOT / "core" / "frontend" / "web" / "index.html"


def main() -> int:
    html = HTML.read_text(encoding="utf-8")
    scripts = re.findall(r"<script>(.*?)</script>", html, re.S)
    if not scripts:
        print("FAIL: 未找到 <script> 块")
        return 1
    js = max(scripts, key=len)
    tmp = ROOT / "_check_frontend.js"
    tmp.write_text(js, encoding="utf-8")
    try:
        r = subprocess.run(["node", "--check", str(tmp)],
                           capture_output=True, text=True)
    except FileNotFoundError:
        print("SKIP: 未安装 node，无法做 JS 语法校验")
        tmp.unlink(missing_ok=True)
        return 0
    tmp.unlink(missing_ok=True)
    if r.returncode != 0:
        print("FAIL: 前端 JS 语法错误")
        print(r.stderr[:800])
        return 1
    # 附带检查：危险的字面量换行（Python 补丁常见的破坏形态）
    if re.search(r'"[^"\n]*\n[^"]*"', js):
        print("WARN: JS 中疑似存在跨行字符串字面量")
    print("OK: 前端 JS 语法通过（%d 字符）" % len(js))
    return 0


if __name__ == "__main__":
    sys.exit(main())

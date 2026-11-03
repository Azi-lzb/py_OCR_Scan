# -*- coding: utf-8 -*-
"""Build an amd64 Debian package on a glibc 2.28 Linux host."""
from __future__ import annotations

import argparse
import io
import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "OCR国库会计工具-Flask"
PACKAGE = "ocr-guoku-accounting"
INSTALL_DIR = Path("opt/ocr-guoku")
SYSTEM_DEPENDS = (
    "libc6 (>= 2.28), libstdc++6, libglib2.0-0, libgl1, libgomp1, "
    "libssl1.1, libffi6, libreadline7, libtinfo6, libgnutls30, libgmp10, "
    "libunistring2, libidn2-0, libhogweed4, libnettle6, libp11-kit0, "
    "libtasn1-6, libx11-6, libxau6, libxext6, libice6, libsm6, "
    "libbz2-1.0, liblzma5, libuuid1, zlib1g, xdg-utils"
)


def require_target_host() -> None:
    libc, version = platform.libc_ver()
    machine = platform.machine().lower()
    if libc != "glibc" or version != "2.28":
        raise SystemExit(
            f"请在 glibc 2.28 环境构建；当前检测到 {libc or 'unknown'} {version or 'unknown'}")
    if machine not in ("x86_64", "amd64"):
        raise SystemExit(f"当前构建只支持 amd64，检测到架构：{machine}")


def _write_wrapper(path: Path) -> None:
    script = f'''#!/bin/sh
set -eu
APP_DIR=/opt/ocr-guoku
USER_CORE="${{XDG_DATA_HOME:-$HOME/.local/share}}/ocr-guoku/core"
mkdir -p "$USER_CORE/data"
if [ -d "$APP_DIR/core/data" ]; then
    cp -an "$APP_DIR/core/data/." "$USER_CORE/data/"
fi
export OCR_TOOLS_CORE_PATH="$USER_CORE"
exec "$APP_DIR/{NAME}" "$@"
'''
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(script, encoding="utf-8")
    path.chmod(0o755)


def _write_desktop_file(path: Path) -> None:
    content = f'''[Desktop Entry]
Type=Application
Name=OCR 国库会计工具
Comment=离线 OCR 与会计报表处理
Exec=ocr-guoku-accounting
Terminal=false
Categories=Office;Utility;
StartupNotify=true
'''
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def _tar_gz(root: Path, names: list[str]) -> bytes:
    buffer = io.BytesIO()
    with tarfile.open(fileobj=buffer, mode="w:gz", format=tarfile.PAX_FORMAT) as archive:
        for name in names:
            item = root / name
            archive.add(item, arcname=name, recursive=True)
    return buffer.getvalue()


def _ar_member(name: str, payload: bytes) -> bytes:
    if len(name) > 15:
        raise ValueError(f"ar 成员名过长：{name}")
    header = (
        f"{name + '/':<16}"
        f"{int(time.time()):<12}"
        f"{0:<6}{0:<6}{0o100644:<8o}{len(payload):<10}`\n"
    ).encode("ascii")
    if len(header) != 60:
        raise ValueError("生成的 ar 成员头长度错误")
    return header + payload + (b"\n" if len(payload) % 2 else b"")


def _write_control(stage: Path, version: str) -> None:
    package_root = stage / "DEBIAN"
    package_root.mkdir(parents=True, exist_ok=True)
    content = (
        f"Package: {PACKAGE}\n"
        f"Version: {version}\n"
        "Section: utils\n"
        "Priority: optional\n"
        "Architecture: amd64\n"
        f"Depends: {SYSTEM_DEPENDS}\n"
        "Maintainer: OCR Tools <noreply@example.invalid>\n"
        "Description: Offline OCR and accounting spreadsheet tool\n"
        " Bundles the Python runtime, OCR models, and application libraries.\n"
    )
    control = package_root / "control"
    control.write_text(content, encoding="utf-8")
    control.chmod(0o644)


def _make_deb(stage: Path, output: Path) -> None:
    # Debian's ar member should be named "control" inside the control archive.
    control_buffer = io.BytesIO()
    with tarfile.open(fileobj=control_buffer, mode="w:gz", format=tarfile.PAX_FORMAT) as archive:
        archive.add(stage / "DEBIAN" / "control", arcname="control")
    control_data = control_buffer.getvalue()
    data_names = ["opt", "usr"]
    data_data = _tar_gz(stage, data_names)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_suffix(output.suffix + ".tmp")
    with temp.open("wb") as stream:
        stream.write(b"!<arch>\n")
        stream.write(_ar_member("debian-binary", b"2.0\n"))
        stream.write(_ar_member("control.tar.gz", control_data))
        stream.write(_ar_member("data.tar.gz", data_data))
    os.replace(temp, output)


def build_package(version: str, build_app: bool = True) -> Path:
    require_target_host()
    if build_app:
        subprocess.run([sys.executable, str(ROOT / "build_exe.py")],
                       cwd=ROOT, check=True)
    app_dir = ROOT / "dist" / NAME
    executable = app_dir / NAME
    if not executable.is_file():
        raise SystemExit(f"PyInstaller 目录版不存在：{executable}")

    stage = ROOT / "build" / "deb-root"
    resolved_stage = stage.resolve()
    if ROOT.joinpath("build").resolve() not in resolved_stage.parents:
        raise RuntimeError("Debian staging path escaped the build directory")
    if stage.exists():
        shutil.rmtree(stage)
    install_path = stage / INSTALL_DIR
    install_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(app_dir, install_path)
    _write_wrapper(stage / "usr" / "bin" / PACKAGE)
    _write_desktop_file(stage / "usr" / "share" / "applications"
                        / f"{PACKAGE}.desktop")
    _write_control(stage, version)

    size = sum(f.stat().st_size for f in install_path.rglob("*") if f.is_file())
    control = stage / "DEBIAN" / "control"
    text = control.read_text(encoding="utf-8")
    control.write_text(text + f"Installed-Size: {(size + 1023) // 1024}\n",
                       encoding="utf-8")
    output = ROOT / "dist" / f"{PACKAGE}_{version}_amd64.deb"
    _make_deb(stage, output)
    print(f"Debian 包已生成：{output}")
    print(f"应用文件大小：{size / 1048576:.1f} MiB")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", default="1.0.0")
    parser.add_argument("--no-build", action="store_true",
                        help="使用已有的 PyInstaller onedir 目录")
    args = parser.parse_args()
    if not re.fullmatch(r"[0-9]+(?:\.[0-9]+){1,3}", args.version):
        parser.error("版本号格式应为例如 1.0.0")
    build_package(args.version, build_app=not args.no_build)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

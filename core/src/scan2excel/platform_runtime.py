"""Environment for launching system applications from a frozen Linux bundle."""
from __future__ import annotations

import os
import sys
from pathlib import Path, PureWindowsPath


def packaged_target_path(value: str, core: Path) -> str:
    """Relocate only the known Windows seed target on Linux."""
    if not sys.platform.startswith("linux") or not value:
        return value
    path = PureWindowsPath(value)
    if not path.is_absolute():
        return value
    tail = tuple(part.casefold() for part in path.parts[-4:])
    if tail == ("core", "data", "transforms", "国库报表.xlsx"):
        target = core / "data" / "transforms" / path.name
        if target.is_file():
            return str(target)
    raise RuntimeError("目标工作簿仍为 Windows 路径，请在规则配置中改为 Linux 本机路径：" + value)


def system_app_environment() -> dict[str, str]:
    env = os.environ.copy()
    if sys.platform.startswith("linux") and getattr(sys, "frozen", False):
        original = env.pop("LD_LIBRARY_PATH_ORIG", None)
        if original is None:
            env.pop("LD_LIBRARY_PATH", None)
        else:
            env["LD_LIBRARY_PATH"] = original
    return env

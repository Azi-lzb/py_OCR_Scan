# -*- coding: utf-8 -*-
"""Windows 命名互斥体，防止同一程序开多个实例。非 Windows 平台直接放行。"""
from __future__ import annotations

import sys

_MUTEX_NAME = "Local\\ScanToExcel_SingleInstance_v1"
_handle = None


def acquire() -> bool:
    """尝试获得单实例锁；成功返回 True（本进程是唯一实例）。"""
    global _handle
    if sys.platform != "win32":
        return True
    try:
        import ctypes
        ERROR_ALREADY_EXISTS = 183
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, _MUTEX_NAME)
        if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
            return False
        _handle = handle
        return True
    except Exception:
        return True  # 拿不到系统 API 时宁可放行


def show_already_running_message() -> None:
    try:
        import ctypes
        MB_OK = 0x0
        MB_ICONINFORMATION = 0x40
        ctypes.windll.user32.MessageBoxW(
            None, "OCR工具 已在运行中（请查看任务栏）。",
            "已打开", MB_OK | MB_ICONINFORMATION)
    except Exception:
        pass

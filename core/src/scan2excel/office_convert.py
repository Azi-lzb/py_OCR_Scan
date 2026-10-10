# -*- coding: utf-8 -*-
"""批量格式转换（移植 pytools 4-2/4-3）：
- Excel 系：xls/xlsx/xlsm/csv/et… 互转，引擎依次试 Excel.Application / WPS 表格(ket)
- Word 系：doc/docx 互转，引擎依次试 Word.Application / WPS 文字(kwps)
输出到 源目录/<目标格式>/ 子文件夹，同名自动 _1/_2 递增；单引擎实例跑整批。
调用方须在已 CoInitialize 的线程里执行（web_app 的 _COM 执行器负责）。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List

EXCEL_FMT_MAP = {"xls": 56, "xlsx": 51, "xlsm": 52, "csv": 6,
                 "xlt": 17, "xltx": 54, "xltm": 53, "xlsb": 50}
WORD_FMT_MAP = {"doc": 0, "docx": 12}
EXCEL_LIKE = (".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".et", ".ett", ".csv")
WORD_LIKE = (".doc", ".docx")


def _dispatch_first(progids: List[str]):
    import win32com.client
    last_err: Any = None
    for pid in progids:
        try:
            return win32com.client.DispatchEx(pid), pid
        except Exception as e:
            last_err = e
    raise RuntimeError(f"无法创建 COM 应用（已尝试: {', '.join(progids)}）: {last_err}")


def _start_excel_app(prefer_wps: bool = False):
    progids = (["ket.Application", "KET.Application", "Excel.Application"]
               if prefer_wps else
               ["Excel.Application", "ket.Application", "KET.Application"])
    app, engine = _dispatch_first(progids)
    app.Visible = False
    for attr, val in (("DisplayAlerts", False), ("ScreenUpdating", False),
                      ("EnableEvents", False), ("AskToUpdateLinks", False),
                      ("AlertBeforeOverwriting", False)):
        try:
            setattr(app, attr, val)
        except Exception:
            pass
    try:
        app.Calculation = -4135          # xlCalculationManual
    except Exception:
        pass
    return app, engine


def _start_word_app():
    app, engine = _dispatch_first(["Word.Application", "kwps.Application",
                                   "KWPS.Application", "wps.Application"])
    app.Visible = False
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    return app, engine


def _next_path(p: Path) -> Path:
    if not p.exists():
        return p
    i = 1
    while True:
        cand = p.with_name(f"{p.stem}_{i}{p.suffix}")
        if not cand.exists():
            return cand
        i += 1


def _plan(files: List[Path], target_ext: str, like: tuple) -> tuple:
    """预筛：扩展名合法、源≠目标；输出 = 源目录/<target>/同名。"""
    plan: List[tuple] = []
    skipped: List[Dict[str, str]] = []
    for p in files:
        p = Path(p).resolve()          # COM 服务进程按自己的 CWD 解析相对路径，必须绝对
        if not p.is_file() or p.suffix.lower() not in like:
            skipped.append({"src": str(p), "reason": "不是可转换的文件类型"})
            continue
        out_dir = p.parent / target_ext
        out_dir.mkdir(parents=True, exist_ok=True)   # 目标子文件夹必须先建，Excel 才能存
        out = out_dir / f"{p.stem}.{target_ext}"
        if out.resolve() == p.resolve():
            skipped.append({"src": str(p), "reason": "源文件已是目标格式"})
            continue
        if out.exists():
            out = _next_path(out)
        plan.append((p, out))
    return plan, skipped


def run_excel_convert(files: List[Path], target_ext: str) -> Dict[str, Any]:
    ext = (target_ext or "xlsx").lower().lstrip(".")
    if ext not in EXCEL_FMT_MAP:
        ext = "xlsx"
    plan, skipped = _plan(files, ext, EXCEL_LIKE)
    stats: Dict[str, Any] = {"target": ext, "ok": 0, "skip": len(skipped),
                             "skipped": skipped, "done": [], "engine": ""}
    if not plan:
        return stats
    prefer_wps = any(src.suffix.lower() in (".et", ".ett") for src, _ in plan)
    app, engine = _start_excel_app(prefer_wps=prefer_wps)
    stats["engine"] = engine
    try:
        for src, out in plan:
            wb = app.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)
            try:
                wb.SaveAs(str(out), FileFormat=EXCEL_FMT_MAP[ext], CreateBackup=False)
                stats["ok"] += 1
                stats["done"].append({"src": str(src), "out": str(out)})
            except Exception as e:
                stats["skip"] += 1
                stats["skipped"].append({"src": str(src), "reason": f"转换失败：{e}"})
            finally:
                wb.Close(SaveChanges=False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return stats


def run_word_convert(files: List[Path], target_ext: str) -> Dict[str, Any]:
    ext = (target_ext or "docx").lower().lstrip(".")
    if ext not in WORD_FMT_MAP:
        ext = "docx"
    plan, skipped = _plan(files, ext, WORD_LIKE)
    stats: Dict[str, Any] = {"target": ext, "ok": 0, "skip": len(skipped),
                             "skipped": skipped, "done": [], "engine": ""}
    if not plan:
        return stats
    app, engine = _start_word_app()
    stats["engine"] = engine
    try:
        for src, out in plan:
            doc = app.Documents.Open(str(src), ConfirmConversions=False,
                                     ReadOnly=True, AddToRecentFiles=False)
            try:
                try:
                    doc.SaveAs2(str(out), FileFormat=WORD_FMT_MAP[ext])
                except Exception:
                    doc.SaveAs(str(out), FileFormat=WORD_FMT_MAP[ext])
                stats["ok"] += 1
                stats["done"].append({"src": str(src), "out": str(out)})
            except Exception as e:
                stats["skip"] += 1
                stats["skipped"].append({"src": str(src), "reason": f"转换失败：{e}"})
            finally:
                doc.Close(SaveChanges=False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return stats

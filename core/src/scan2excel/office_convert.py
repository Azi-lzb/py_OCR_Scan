# -*- coding: utf-8 -*-
"""批量格式转换（移植 pytools 4-2/4-3，强制计算对齐 pytools-core 的 pywin32 worker）：
- Excel 系：xls/xlsx/xlsm/csv/et… 互转，引擎依次试 Excel.Application / WPS 表格(ket)
- Word 系：doc/docx 互转，引擎依次试 Word.Application / WPS 文字(kwps)
输出到 源目录/<目标格式>/ 子文件夹，同名自动 _1/_2 递增；单引擎实例跑整批。
调用方须在已 CoInitialize 的线程里执行（web_app 的 _COM 执行器负责）。

强制计算：SaveAs 前调用 Application.CalculateFullRebuild（该方法是 Application
级，Workbook 上没有），重算完先给源簿拍快照，转出候选文件后重新打开逐格比对，
不一致就不发布——避免把过期公式缓存或结构丢失的文件当成转换成功。
"""
from __future__ import annotations

import math
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

EXCEL_FMT_MAP = {"xls": 56, "xlsx": 51, "xlsm": 52, "csv": 6,
                 "xlt": 17, "xltx": 54, "xltm": 53, "xlsb": 50}
WORD_FMT_MAP = {"doc": 0, "docx": 12}
EXCEL_LIKE = (".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".et", ".ett", ".csv")
WORD_LIKE = (".doc", ".docx")
# WPS 原生表格容器：.et/.ett 是 CFB 文件，Excel 会当成改名 .xls 打开并报出看似
# 合理的 FileFormat，所以必须先确认是 WPS 自己的格式再发布结果。
WPS_SPREADSHEET_SUFFIXES = (".et", ".ett")
WPS_NATIVE_FILE_FORMAT_MIN = 65521
_OOXML_BOOK_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")
_CALCPR_RE = re.compile(r"<calcPr\b([^>]*?)/>", re.S)


# ---------------------------------------------------------------- COM 启动
def _dispatch_empty_application(progids: List[str], collection_name: str,
                                document_alerts: bool = False):
    """只创建能被证明是空白的 COM 实例，绝不复用用户已打开文档的会话。

    DispatchEx 本应新建进程，但 WPS 等实现可能回传已有实例；若其中已有文档，
    后续 Quit 会连带关掉用户未保存的工作，因此宁可跳过该 ProgID。"""
    import win32com.client

    errors: List[str] = []
    for pid in progids:
        try:
            app = win32com.client.DispatchEx(pid)
        except Exception as e:
            errors.append(f"{pid}: {e}")
            continue
        try:
            opened = int(getattr(app, collection_name).Count)
        except Exception as e:
            # 无法确认文档数量，只释放我们的引用，不对未知会话调用 Quit。
            errors.append(f"{pid}: 无法确认实例中的文档数量：{e}")
            continue
        if opened != 0:
            errors.append(f"{pid}: 返回了含 {opened} 个已打开文档的实例，为保护用户数据已跳过")
            continue
        try:
            app.Visible = False
            if document_alerts:
                app.DisplayAlerts = 0          # Word 的 WdAlertLevel
            else:
                app.DisplayAlerts = False
                app.ScreenUpdating = False
                app.EnableEvents = False
            try:
                app.AutomationSecurity = 3     # msoAutomationSecurityForceDisable
            except Exception:
                pass
            return app, pid
        except Exception as e:
            try:
                app.Quit()
            except Exception:
                pass
            errors.append(f"{pid}: 初始化空实例失败：{e}")
    raise RuntimeError("无法创建 COM 应用（已尝试: " + "；".join(errors) + "）")


def _start_excel_app(prefer_wps: bool = False):
    progids = (["ket.Application", "KET.Application", "Excel.Application"]
               if prefer_wps else
               ["Excel.Application", "ket.Application", "KET.Application"])
    app, engine = _dispatch_empty_application(progids, "Workbooks")
    for attr, val in (("AskToUpdateLinks", False), ("AlertBeforeOverwriting", False),
                      ("DisplayRecentFiles", False), ("CutCopyMode", False)):
        try:
            setattr(app, attr, val)
        except Exception:
            pass
    return app, engine


def _start_word_app():
    app, engine = _dispatch_empty_application(
        ["Word.Application", "kwps.Application", "KWPS.Application", "wps.Application"],
        "Documents", document_alerts=True)
    return app, engine


# ---------------------------------------------------------------- 强制计算 + 校验
def _force_recalc(app) -> Tuple[bool, str]:
    """转换前强制全量重算，返回 (是否成功, 实际调用的方法名)。

    源文件若存成 calcMode="manual" 且带有效 calcId，Excel 打开时不会重算，
    SaveAs 会把过期的公式缓存值原样写进输出文件（实测 999 就是 999）。
    CalculateFullRebuild 重建依赖链并重算全部公式，比 CalculateFull/Calculate
    更彻底；它只在 Application 上，Workbook 上没有该成员。"""
    for name in ("CalculateFullRebuild", "CalculateFull", "Calculate"):
        try:
            getattr(app, name)()
            return True, name
        except Exception:
            continue
    return False, ""


def _matrix(value: Any, rows: int, columns: int) -> List[Any]:
    """COM 批量读取返回值：单格是标量，单行/单列是一维，其余是二维。"""
    if rows == 1 and columns == 1:
        return [[value]]
    if rows == 1:
        return [list(value)]
    if columns == 1:
        return [[item] for item in value]
    return [list(row) for row in value]


def _column_name(column: int) -> str:
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _cell_name(address: Tuple[int, int]) -> str:
    return f"{_column_name(address[1])}{address[0]}"


def _values_equal(left: Any, right: Any) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        if left == right:
            return True
        return (math.isfinite(left) and math.isfinite(right)
                and left != 0 and right != 0
                and abs(left - right) <= 1e-15 * max(abs(left), abs(right)))
    return type(left) is type(right) and left == right


def _snapshot(workbook) -> List[Dict[str, Any]]:
    """用 UsedRange 批量读取工作表快照：公式文本 + 非公式值。"""
    sheets: List[Dict[str, Any]] = []
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        worksheet = workbook.Worksheets(index)
        used = worksheet.UsedRange
        rows = int(used.Rows.Count)
        columns = int(used.Columns.Count)
        first_row = int(used.Row)
        first_column = int(used.Column)
        formulas = _matrix(used.Formula, rows, columns)
        values = _matrix(used.Value2, rows, columns)
        formula_cells: Dict[Tuple[int, int], str] = {}
        value_cells: Dict[Tuple[int, int], Any] = {}
        for row_offset in range(rows):
            for column_offset in range(columns):
                address = (first_row + row_offset, first_column + column_offset)
                formula = formulas[row_offset][column_offset]
                value = values[row_offset][column_offset]
                if isinstance(formula, str) and formula.startswith("="):
                    formula_cells[address] = formula
                elif value is not None:
                    value_cells[address] = value
        sheets.append({"name": str(worksheet.Name), "formulas": formula_cells,
                       "values": value_cells})
    return sheets


def _compare(source: Sequence[Dict[str, Any]],
             output: Sequence[Dict[str, Any]]) -> List[str]:
    """逐格比对源快照与转出快照，返回差异说明（空列表表示一致）。"""
    problems: List[str] = []
    source_names = [sheet["name"] for sheet in source]
    output_names = [sheet["name"] for sheet in output]
    if source_names != output_names:
        return [f"工作表数量、名称或顺序发生变化：{source_names} -> {output_names}"]
    for source_sheet, output_sheet in zip(source, output):
        sheet_name = source_sheet["name"]
        for kind, label, equals in (
            ("formulas", "公式", lambda a, b: str(a).strip() == str(b).strip()),
            ("values", "非公式数据", _values_equal),
        ):
            before = source_sheet[kind]
            after = output_sheet[kind]
            for address in sorted(set(before) | set(after)):
                if address not in before or address not in after:
                    action = "缺失" if address in before else "多出"
                    problems.append(
                        f"工作表 {sheet_name!r} 的{label}单元格 {_cell_name(address)} {action}")
                elif not equals(before[address], after[address]):
                    problems.append(
                        f"工作表 {sheet_name!r} 的{label}在 {_cell_name(address)} 不一致："
                        f"{before[address]!r} -> {after[address]!r}")
                if len(problems) >= 8:
                    return problems
    return problems


def _verify_candidate(app, candidate: Path,
                      source_snapshot: Sequence[Dict[str, Any]]) -> List[str]:
    """重新打开转出文件再拍一次快照，与源快照比对。"""
    reopened = app.Workbooks.Open(str(candidate), UpdateLinks=0, ReadOnly=True)
    try:
        output_snapshot = _snapshot(reopened)
    finally:
        reopened.Close(SaveChanges=False)
    return _compare(source_snapshot, output_snapshot)


def _mark_full_calc_on_load(path: Path) -> None:
    """给 OOXML 输出打 fullCalcOnLoad 标记，让 Excel/WPS 下次打开时再全量重算。

    这不代替转换时的 CalculateFullRebuild：只是给外部引擎（含本工具后续用
    calamine 读取）一个二次保险，缓存值仍以转换时重算结果为准。"""
    if path.suffix.lower() not in _OOXML_BOOK_SUFFIXES:
        return
    try:
        with zipfile.ZipFile(path) as archive:
            items = [(info, archive.read(info.filename)) for info in archive.infolist()]
        patched: List[Tuple[zipfile.ZipInfo, bytes]] = []
        changed = False
        for info, data in items:
            if info.filename == "xl/workbook.xml":
                text = data.decode("utf-8")
                match = _CALCPR_RE.search(text)
                if match:
                    attrs = re.sub(r'\s(fullCalcOnLoad|forceFullCalc)="[^"]*"', "",
                                   match.group(1))
                    text = (text[:match.start()] + f'<calcPr{attrs} fullCalcOnLoad="1"/>'
                            + text[match.end():])
                elif "</workbook>" in text:
                    text = text.replace("</workbook>",
                                        '<calcPr fullCalcOnLoad="1"/></workbook>')
                data = text.encode("utf-8")
                changed = True
            patched.append((info, data))
        if not changed:
            return
        temp = path.with_name(f".{path.name}.calc.tmp")
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
            for info, data in patched:
                archive.writestr(info, data)
        temp.replace(path)
    except Exception:
        # 标记失败不影响已重算的缓存值，也不应让整批转换失败。
        return


# ---------------------------------------------------------------- 路径规划
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


# ---------------------------------------------------------------- 批量转换
def run_excel_convert(files: List[Path], target_ext: str) -> Dict[str, Any]:
    ext = (target_ext or "xlsx").lower().lstrip(".")
    if ext not in EXCEL_FMT_MAP:
        ext = "xlsx"
    plan, skipped = _plan(files, ext, EXCEL_LIKE)
    stats: Dict[str, Any] = {"target": ext, "ok": 0, "skip": len(skipped),
                             "skipped": skipped, "done": [], "engine": "",
                             "recalc": 0, "recalc_failed": 0, "recalc_method": "",
                             "verified": 0, "verify_failed": 0, "verify_detail": []}
    if not plan:
        return stats
    prefer_wps = any(src.suffix.lower() in WPS_SPREADSHEET_SUFFIXES for src, _ in plan)
    app, engine = _start_excel_app(prefer_wps=prefer_wps)
    stats["engine"] = engine
    try:
        for src, out in plan:
            # 先写临时候选，校验通过再改名到最终路径，失败不留半成品。
            candidate = out.with_name(f".{out.stem}.converting{out.suffix}")
            try:
                if candidate.exists():
                    candidate.unlink()
                wb = app.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)
                try:
                    if src.suffix.lower() in WPS_SPREADSHEET_SUFFIXES:
                        file_format = int(wb.FileFormat)
                        if file_format < WPS_NATIVE_FILE_FORMAT_MIN:
                            raise ValueError(
                                f"扩展名是 {src.suffix}，但实际打开为 FileFormat "
                                f"{file_format}，不是 WPS 原生表格")
                    snapshot = None
                    if ext != "csv":          # CSV 没有公式，不做重算与逐格比对
                        ok, method = _force_recalc(app)
                        if ok:
                            stats["recalc"] += 1
                            stats["recalc_method"] = method
                        else:
                            stats["recalc_failed"] += 1
                        snapshot = _snapshot(wb)
                    wb.SaveAs(str(candidate), FileFormat=EXCEL_FMT_MAP[ext],
                              CreateBackup=False)
                finally:
                    wb.Close(SaveChanges=False)

                if snapshot is not None:
                    problems = _verify_candidate(app, candidate, snapshot)
                    if problems:
                        stats["verify_failed"] += 1
                        stats["verify_detail"].append({"src": str(src),
                                                       "problems": problems[:8]})
                        raise ValueError("转换后校验不一致：" + "；".join(problems[:3]))
                    stats["verified"] += 1
                _mark_full_calc_on_load(candidate)
                candidate.replace(out)
                stats["ok"] += 1
                stats["done"].append({"src": str(src), "out": str(out),
                                      "verified": snapshot is not None})
            except Exception as e:
                stats["skip"] += 1
                stats["skipped"].append({"src": str(src), "reason": f"转换失败：{e}"})
                try:
                    if candidate.exists():
                        candidate.unlink()
                except OSError:
                    pass
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

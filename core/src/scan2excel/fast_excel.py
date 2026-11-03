# -*- coding: utf-8 -*-
"""Fast Excel I/O based on Calamine, OOXML and XlsxWriter.

Reads .xls/.xlsx with python-calamine and rewrites .xlsx files with XlsxWriter.
XLSX formulas, names, merge ranges, hidden sheets and expression checks are
carried through; visual formatting is intentionally minimal.
"""
from __future__ import annotations

import os
import posixpath
import tempfile
import threading
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple


BACKEND_MODES = ("auto", "fast", "compat")
_BACKEND_MODE = "auto"
_BACKEND_LOCK = threading.RLock()
_FALLBACK_EVENTS: List[str] = []


def get_backend_mode() -> str:
    with _BACKEND_LOCK:
        return _BACKEND_MODE


def set_backend_mode(mode: str) -> str:
    """Select the Excel engine used by later read/write calls."""
    normalized = str(mode or "").strip().lower()
    if normalized not in BACKEND_MODES:
        raise ValueError("Excel 引擎只能是 auto、fast 或 compat")
    global _BACKEND_MODE
    with _BACKEND_LOCK:
        _BACKEND_MODE = normalized
    return normalized


def consume_fallback_events() -> List[str]:
    """Return and clear automatic compatibility fallbacks for the UI log."""
    with _BACKEND_LOCK:
        events = list(_FALLBACK_EVENTS)
        _FALLBACK_EVENTS.clear()
    return events


def _record_fallback(message: str) -> None:
    with _BACKEND_LOCK:
        _FALLBACK_EVENTS.append(message)


@dataclass(frozen=True)
class CellRange:
    min_row: int
    min_col: int
    max_row: int
    max_col: int


@dataclass(frozen=True)
class DefinedName:
    name: str
    formula: str
    local_sheet: Optional[str] = None


@dataclass(frozen=True)
class ConditionalFormat:
    ranges: str
    formula: str
    kind: str = "expression"


class _MergedCells:
    def __init__(self, ranges: Iterable[CellRange] = ()) -> None:
        self.ranges = list(ranges)


class FastCell:
    """Small mutable cell facade for code that expects ``cell.value``."""

    def __init__(self, sheet: "FastWorksheet", row: int, column: int) -> None:
        self._sheet = sheet
        self._row = row
        self._column = column
        self.number_format: Any = None
        self.font: Any = None
        self.fill: Any = None

    @property
    def value(self) -> Any:
        return self._sheet._get(self._row, self._column)

    @value.setter
    def value(self, value: Any) -> None:
        self._sheet._set(self._row, self._column, value)


class FastWorksheet:
    def __init__(self, title: str, rows: Optional[Sequence[Sequence[Any]]] = None,
                 merged_ranges: Iterable[CellRange] = (), visible: bool = True,
                 state: Optional[str] = None) -> None:
        self.title = title
        self.rows: List[List[Any]] = [list(row) for row in (rows or [])]
        self.merged_cells = _MergedCells(merged_ranges)
        self.visible = visible
        self.state = state or ("visible" if visible else "hidden")
        self.formula_cells: Dict[Tuple[int, int], str] = {}
        self.conditional_formats: List[ConditionalFormat] = []

    @property
    def max_row(self) -> int:
        return len(self.rows)

    @property
    def max_column(self) -> int:
        return max((len(row) for row in self.rows), default=0)

    def _get(self, row: int, column: int) -> Any:
        if row < 1 or column < 1 or row > len(self.rows):
            return None
        current = self.rows[row - 1]
        return current[column - 1] if column <= len(current) else None

    def _set(self, row: int, column: int, value: Any) -> None:
        if row < 1 or column < 1:
            raise ValueError("Excel 行列号从 1 开始")
        while len(self.rows) < row:
            self.rows.append([])
        current = self.rows[row - 1]
        if len(current) < column:
            current.extend([None] * (column - len(current)))
        current[column - 1] = value

    def cell(self, row: int, column: int, value: Any = None) -> FastCell:
        cell = FastCell(self, row, column)
        if value is not None:
            cell.value = value
        return cell

    def iter_rows(self, min_row: int = 1, max_row: Optional[int] = None,
                  min_col: int = 1, max_col: Optional[int] = None,
                  values_only: bool = False) -> Iterator[Tuple[Any, ...]]:
        last_row = self.max_row if max_row is None else max_row
        last_col = self.max_column if max_col is None else max_col
        for row in range(min_row, last_row + 1):
            values = tuple(self._get(row, col) for col in range(min_col, last_col + 1))
            yield values if values_only else tuple(
                self.cell(row, col) for col in range(min_col, last_col + 1))


class FastWorkbook:
    def __init__(self) -> None:
        self.worksheets: List[FastWorksheet] = []
        self.defined_names: List[DefinedName] = []

    @property
    def sheetnames(self) -> List[str]:
        return [sheet.title for sheet in self.worksheets]

    def __contains__(self, title: str) -> bool:
        return title in self.sheetnames

    def __getitem__(self, title: str) -> FastWorksheet:
        for sheet in self.worksheets:
            if sheet.title == title:
                return sheet
        raise KeyError(title)

    def add_sheet(self, sheet: FastWorksheet, index: Optional[int] = None) -> FastWorksheet:
        if sheet.title in self.sheetnames:
            raise ValueError(f"工作表名称重复：{sheet.title}")
        if index is None:
            self.worksheets.append(sheet)
        else:
            self.worksheets.insert(index, sheet)
        return sheet

    def create_sheet(self, title: str, index: Optional[int] = None) -> FastWorksheet:
        return self.add_sheet(FastWorksheet(title), index)

    def remove(self, sheet: FastWorksheet) -> None:
        self.worksheets.remove(sheet)

    def __delitem__(self, title: str) -> None:
        self.remove(self[title])

    def close(self) -> None:
        return None

    def save(self, path: Path) -> None:
        write_workbook(path, self)


_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _ooxml_metadata(path: Path) -> Tuple[Dict[str, dict], List[DefinedName]]:
    """Read formulas, names, visibility and conditional rules from XLSX XML."""
    sheet_meta: Dict[str, dict] = {}
    names: List[DefinedName] = []
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return sheet_meta, names
    try:
        with zipfile.ZipFile(path) as archive:
            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            rels = {item.attrib["Id"]: item.attrib["Target"]
                    for item in rel_root.findall(f"{{{_PKG_REL_NS}}}Relationship")}
            sheet_paths: Dict[str, str] = {}
            for index, item in enumerate(workbook_root.findall(
                    f"{{{_MAIN_NS}}}sheets/{{{_MAIN_NS}}}sheet")):
                name = item.attrib.get("name", "")
                rel_id = item.attrib.get(f"{{{_REL_NS}}}id", "")
                target = rels.get(rel_id, "")
                if target.startswith("/"):
                    target = target.lstrip("/")
                else:
                    target = posixpath.normpath(posixpath.join("xl", target))
                state = item.attrib.get("state", "visible")
                sheet_paths[name] = target
                sheet_meta[name] = {"path": target, "state": state,
                                    "index": index, "formulas": {},
                                    "conditional_formats": []}

            defined = workbook_root.find(f"{{{_MAIN_NS}}}definedNames")
            if defined is not None:
                index_to_name = {v["index"]: k for k, v in sheet_meta.items()}
                for item in defined.findall(f"{{{_MAIN_NS}}}definedName"):
                    raw_local = item.attrib.get("localSheetId")
                    local = index_to_name.get(int(raw_local)) if raw_local is not None else None
                    names.append(DefinedName(item.attrib.get("name", ""),
                                             item.text or "", local))

            for name, metadata in sheet_meta.items():
                try:
                    root = ET.fromstring(archive.read(metadata["path"]))
                except (KeyError, ET.ParseError):
                    continue
                shared_bases: Dict[str, Tuple[Tuple[int, int], str]] = {}
                deferred_shared: List[Tuple[Tuple[int, int], str]] = []
                for cell in root.findall(f".//{{{_MAIN_NS}}}c"):
                    formula = cell.find(f"{{{_MAIN_NS}}}f")
                    if formula is None:
                        continue
                    coord = _a1_to_rc(cell.attrib.get("r", ""))
                    if coord is None:
                        continue
                    text = (formula.text or "").strip()
                    if formula.attrib.get("t") == "shared":
                        share_id = formula.attrib.get("si", "")
                        if text:
                            shared_bases[share_id] = (coord, text)
                            metadata["formulas"][coord] = "=" + text
                        else:
                            deferred_shared.append((coord, share_id))
                    elif text:
                        metadata["formulas"][coord] = "=" + text
                for coord, share_id in deferred_shared:
                    shared = shared_bases.get(share_id)
                    if shared:
                        base_cell, base_formula = shared
                        metadata["formulas"][coord] = "=" + _translate_formula(
                            base_formula, base_cell, coord)
                for cf in root.findall(f"{{{_MAIN_NS}}}conditionalFormatting"):
                    ranges = cf.attrib.get("sqref", "")
                    for rule in cf.findall(f"{{{_MAIN_NS}}}cfRule"):
                        formula = rule.find(f"{{{_MAIN_NS}}}formula")
                        if formula is not None and formula.text:
                            metadata["conditional_formats"].append(
                                ConditionalFormat(ranges, formula.text.strip(),
                                                  rule.attrib.get("type", "expression")))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError, ValueError):
        return {}, []
    return sheet_meta, names


def _a1_to_rc(coord: str) -> Optional[Tuple[int, int]]:
    import re
    match = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", coord or "", re.I)
    if not match:
        return None
    col = 0
    for char in match.group(1).upper():
        col = col * 26 + ord(char) - 64
    return int(match.group(2)), col


def _translate_formula(formula: str, base: Tuple[int, int],
                       target: Tuple[int, int]) -> str:
    """Translate relative A1 references in a shared formula to another cell."""
    import re

    row_delta = target[0] - base[0]
    col_delta = target[1] - base[1]
    string_parts = re.split(r'("(?:[^"]|"")*")', formula)
    ref_pattern = re.compile(r"(?<![A-Za-z0-9_.])(?P<ca>\$?)(?P<col>[A-Z]{1,3})"
                             r"(?P<ra>\$?)(?P<row>\d+)(?![A-Za-z0-9_])", re.I)

    def translate(part: str) -> str:
        def repl(match):
            if match.end() < len(part) and part[match.end()] == "(":
                return match.group(0)
            col = _a1_to_rc(match.group("col") + "1")[1]
            row = int(match.group("row"))
            if not match.group("ca"):
                col += col_delta
            if not match.group("ra"):
                row += row_delta
            if row < 1 or col < 1:
                return "#REF!"
            return (match.group("ca") + column_letter(col)
                    + match.group("ra") + str(row))
        return ref_pattern.sub(repl, part)

    return "".join(part if index % 2 else translate(part)
                   for index, part in enumerate(string_parts))


def _read_fast(path: Path, include_metadata: bool) -> FastWorkbook:
    """Read values with Calamine and optional XLSX semantics from OOXML."""
    from python_calamine import CalamineWorkbook, SheetVisibleEnum

    source_path = Path(path)
    if include_metadata:
        sheet_metadata, defined_names = _ooxml_metadata(source_path)
    else:
        sheet_metadata, defined_names = {}, []
    result = FastWorkbook()
    result.defined_names = defined_names
    with CalamineWorkbook.from_path(source_path) as wb:
        for name in wb.sheet_names:
            source = wb.get_sheet_by_name(name)
            rows = source.to_python(skip_empty_area=False)
            raw_merges = source.merged_cell_ranges or []
            merges = [CellRange(start[0] + 1, start[1] + 1,
                                end[0] + 1, end[1] + 1)
                      for start, end in raw_merges]
            metadata = sheet_metadata.get(name, {})
            state = metadata.get("state")
            if state is None:
                raw_visibility = next((meta.visible for meta in wb.sheets_metadata
                                       if meta.name == name), SheetVisibleEnum.Visible)
                if raw_visibility == SheetVisibleEnum.VeryHidden:
                    state = "veryHidden"
                elif raw_visibility == SheetVisibleEnum.Hidden:
                    state = "hidden"
                else:
                    state = "visible"
            sheet = FastWorksheet(name, rows, merges, state == "visible", state)
            sheet.formula_cells = metadata.get("formulas", {})
            sheet.conditional_formats = metadata.get("conditional_formats", [])
            for row, column in sheet.formula_cells:
                sheet._set(row, column, sheet._get(row, column))
            result.add_sheet(sheet)
    return result


def _read_compat(path: Path, include_metadata: bool) -> FastWorkbook:
    """Legacy reader: OpenPyXL for XLSX, xlrd for XLS."""
    suffix = path.suffix.lower()
    if suffix == ".xls":
        try:
            import xlrd
        except ImportError as exc:
            raise RuntimeError("兼容引擎读取 .xls 需要 xlrd，请安装 xlrd>=2.0.1") from exc
        from xlrd import XL_CELL_BLANK, XL_CELL_DATE, XL_CELL_EMPTY, XL_CELL_ERROR

        source = xlrd.open_workbook(str(path), on_demand=True)
        result = FastWorkbook()
        try:
            visibility_values = getattr(
                source, "sheet_visibility",
                getattr(source, "_sheet_visibility", ()))
            for index in range(source.nsheets):
                old = source.sheet_by_index(index)
                rows = []
                for row_index in range(old.nrows):
                    values = []
                    for cell in old.row(row_index):
                        value = cell.value
                        if cell.ctype in (XL_CELL_EMPTY, XL_CELL_BLANK):
                            value = None
                        elif cell.ctype == XL_CELL_DATE:
                            value = xlrd.xldate.xldate_as_datetime(
                                value, source.datemode)
                        elif cell.ctype == XL_CELL_ERROR:
                            value = xlrd.error_text_from_code.get(
                                int(value), str(value))
                        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                            value = bool(value)
                        elif cell.ctype == xlrd.XL_CELL_NUMBER and float(value).is_integer():
                            value = int(value)
                        values.append(value)
                    rows.append(values)
                merges = [CellRange(rlo + 1, clo + 1, rhi, chi)
                          for rlo, rhi, clo, chi in old.merged_cells]
                visibility = (visibility_values[index]
                              if index < len(visibility_values) else 0)
                state = "visible" if visibility == 0 else (
                    "veryHidden" if visibility == 2 else "hidden")
                result.add_sheet(FastWorksheet(old.name, rows, merges,
                                               state == "visible", state))
            return result
        finally:
            source.release_resources()

    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError("兼容引擎读取 .xlsx 需要 openpyxl>=3.1") from exc
    old_book = load_workbook(path, data_only=False, read_only=False)
    result = FastWorkbook()
    try:
        sheet_indices = {name: index for index, name in enumerate(old_book.sheetnames)}
        for old_sheet in old_book.worksheets:
            rows = []
            sheet = FastWorksheet(old_sheet.title)
            sheet.state = old_sheet.sheet_state
            sheet.visible = old_sheet.sheet_state == "visible"
            for row_index, old_row in enumerate(old_sheet.iter_rows(), start=1):
                values = []
                for col_index, old_cell in enumerate(old_row, start=1):
                    value = old_cell.value
                    values.append(value)
                    if isinstance(value, str) and value.startswith("="):
                        sheet.formula_cells[(row_index, col_index)] = value
                rows.append(values)
            sheet.rows = rows
            sheet.merged_cells.ranges = [
                CellRange(rg.min_row, rg.min_col, rg.max_row, rg.max_col)
                for rg in old_sheet.merged_cells.ranges]
            if include_metadata:
                for old_cf in old_sheet.conditional_formatting:
                    for old_rule in old_cf.rules:
                        formulas = getattr(old_rule, "formula", None) or []
                        for formula in formulas:
                            sheet.conditional_formats.append(ConditionalFormat(
                                str(old_cf.sqref), str(formula),
                                str(getattr(old_rule, "type", "expression"))))
            result.add_sheet(sheet)

        for old_name in old_book.defined_names.values():
            local_index = getattr(old_name, "localSheetId", None)
            local_sheet = (old_book.sheetnames[int(local_index)]
                           if local_index is not None
                           and int(local_index) < len(old_book.sheetnames) else None)
            result.defined_names.append(DefinedName(
                str(old_name.name or ""), str(old_name.attr_text or ""),
                local_sheet))
        if include_metadata:
            for old_sheet in old_book.worksheets:
                local_names = getattr(old_sheet, "defined_names", None) or {}
                for old_name in local_names.values():
                    result.defined_names.append(DefinedName(
                        str(old_name.name or ""), str(old_name.attr_text or ""),
                        old_sheet.title))
        return result
    finally:
        old_book.close()


def read_workbook(path: Path | str, include_metadata: bool = False,
                  backend: Optional[str] = None) -> FastWorkbook:
    """Read using auto fallback, forced fast, or legacy-compatible libraries."""
    selected = str(backend or get_backend_mode()).lower()
    if selected == "auto":
        try:
            return _read_fast(Path(path), include_metadata)
        except Exception as fast_error:
            try:
                workbook = _read_compat(Path(path), include_metadata)
            except Exception as compat_error:
                raise RuntimeError(
                    f"自动模式读取失败：快速引擎错误（{fast_error}）；"
                    f"兼容引擎错误（{compat_error}）") from compat_error
            _record_fallback(
                f"自动模式：{Path(path).name} 快速读取失败，已回退兼容引擎：{fast_error}")
            return workbook
    if selected == "fast":
        return _read_fast(Path(path), include_metadata)
    if selected == "compat":
        return _read_compat(Path(path), include_metadata)
    raise ValueError(f"未知 Excel 引擎：{selected}")


def _write_xlsxwriter(path: Path, workbook: FastWorkbook) -> None:
    import xlsxwriter

    has_merges = any(sheet.merged_cells.ranges for sheet in workbook.worksheets)
    writer = xlsxwriter.Workbook(
        str(path),
        {"constant_memory": not has_merges,
         "strings_to_formulas": False,
         "strings_to_urls": False,
         "default_date_format": "yyyy-mm-dd"},
    )
    try:
        sheets = workbook.worksheets or [FastWorksheet("Sheet1")]
        for data in sheets:
            ws = writer.add_worksheet(data.title)
            if data.state == "veryHidden":
                ws.very_hidden()
            elif not data.visible:
                ws.hide()

            merged_cells: set[Tuple[int, int]] = set()
            for rg in data.merged_cells.ranges:
                r0, c0 = rg.min_row - 1, rg.min_col - 1
                r1, c1 = rg.max_row - 1, rg.max_col - 1
                value = data._get(rg.min_row, rg.min_col)
                ws.merge_range(r0, c0, r1, c1, "" if value is None else value)
                for r in range(r0, r1 + 1):
                    for c in range(c0, c1 + 1):
                        merged_cells.add((r, c))

            for r, row in enumerate(data.rows):
                for c, value in enumerate(row):
                    if (r, c) in merged_cells:
                        continue
                    formula = data.formula_cells.get((r + 1, c + 1))
                    if formula:
                        ws.write_formula(r, c, formula)
                    elif value is None or value == "":
                        continue
                    elif isinstance(value, datetime):
                        ws.write_datetime(r, c, value)
                    elif isinstance(value, date):
                        ws.write_datetime(r, c, datetime.combine(value, datetime.min.time()))
                    elif isinstance(value, bool):
                        ws.write_boolean(r, c, value)
                    elif isinstance(value, (int, float)):
                        ws.write_number(r, c, value)
                    elif isinstance(value, str):
                        ws.write_string(r, c, value)
                    else:
                        ws.write_string(r, c, str(value))

            for rule in data.conditional_formats:
                if rule.kind != "expression":
                    continue
                # User-facing template formatting is intentionally minimal; retain
                # the formula rule that drives in-app reconciliation checks.
                fmt = writer.add_format({"font_color": "#CC0000"})
                ranges = rule.ranges.split()
                if not ranges:
                    continue
                options = {"type": "formula",
                           "criteria": "=" + rule.formula.lstrip("="),
                           "format": fmt}
                if len(ranges) > 1:
                    options["multi_range"] = " ".join(ranges)
                ws.conditional_format(ranges[0], options)

        for item in workbook.defined_names:
            if not item.name or not item.formula:
                continue
            name = item.name
            if item.local_sheet:
                name = f"'{item.local_sheet.replace(chr(39), chr(39) * 2)}'!{name}"
            try:
                writer.define_name(name, item.formula if item.formula.startswith("=")
                                   else "=" + item.formula)
            except Exception:
                # Built-in or external workbook names may not be accepted by
                # XlsxWriter; they do not affect template region semantics.
                continue
    finally:
        writer.close()


def _write_openpyxl(path: Path, workbook: FastWorkbook) -> None:
    """Legacy-compatible XLSX writer; visual formatting is not rebuilt."""
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Font
    from openpyxl.workbook.defined_name import DefinedName as OpenpyxlDefinedName

    old_book = Workbook()
    old_book.remove(old_book.active)
    try:
        for data in workbook.worksheets or [FastWorksheet("Sheet1")]:
            sheet = old_book.create_sheet(title=data.title)
            if data.state in ("hidden", "veryHidden"):
                sheet.sheet_state = data.state
            for row_index, row in enumerate(data.rows, start=1):
                for col_index, value in enumerate(row, start=1):
                    formula = data.formula_cells.get((row_index, col_index))
                    if formula:
                        sheet.cell(row_index, col_index).value = formula
                    elif value not in (None, ""):
                        sheet.cell(row_index, col_index).value = value
            for rg in data.merged_cells.ranges:
                sheet.merge_cells(start_row=rg.min_row, start_column=rg.min_col,
                                  end_row=rg.max_row, end_column=rg.max_col)
            for rule in data.conditional_formats:
                if rule.kind != "expression" or not rule.formula:
                    continue
                try:
                    sheet.conditional_formatting.add(
                        rule.ranges,
                        FormulaRule(formula=[rule.formula.lstrip("=")],
                                    font=Font(color="FFCC0000")))
                except Exception:
                    continue

        for item in workbook.defined_names:
            if not item.name or not item.formula:
                continue
            local_index = (old_book.sheetnames.index(item.local_sheet)
                           if item.local_sheet in old_book.sheetnames else None)
            defined_name = OpenpyxlDefinedName(
                item.name, attr_text=item.formula.lstrip("="),
                localSheetId=local_index)
            try:
                old_book.defined_names.add(defined_name)
            except Exception:
                continue
        old_book.save(path)
    finally:
        old_book.close()


def _write_fast(path: Path, workbook: FastWorkbook) -> None:
    """Write with XlsxWriter."""
    _write_xlsxwriter(path, workbook)


def write_workbook(path: Path | str, workbook: FastWorkbook,
                   backend: Optional[str] = None) -> None:
    """Atomically write with auto fallback, forced fast, or compat engine.

    The fast writer is XlsxWriter; compat mode retains the OpenPyXL fallback.
    """
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(
        prefix=target.stem + ".", suffix=".tmp.xlsx", dir=str(target.parent))
    os.close(handle)
    temp = Path(temp_name)
    try:
        selected = str(backend or get_backend_mode()).lower()
        if selected == "auto":
            try:
                _write_fast(temp, workbook)
            except Exception as fast_error:
                temp.unlink(missing_ok=True)
                try:
                    _write_openpyxl(temp, workbook)
                except Exception as compat_error:
                    raise RuntimeError(
                        f"自动模式写入失败：快速引擎错误（{fast_error}）；"
                        f"兼容引擎错误（{compat_error}）") from compat_error
                _record_fallback(
                    f"自动模式：{target.name} 快速写入失败，已回退兼容引擎：{fast_error}")
        elif selected == "compat":
            _write_openpyxl(temp, workbook)
        elif selected == "fast":
            _write_fast(temp, workbook)
        else:
            raise ValueError(f"未知 Excel 引擎：{selected}")
        os.replace(temp, target)
    finally:
        temp.unlink(missing_ok=True)


def column_letter(index: int) -> str:
    """Convert a 1-based Excel column index to letters."""
    if index < 1:
        raise ValueError("列号必须大于 0")
    parts: List[str] = []
    while index:
        index, rem = divmod(index - 1, 26)
        parts.append(chr(65 + rem))
    return "".join(reversed(parts))

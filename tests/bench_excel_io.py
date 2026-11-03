# -*- coding: utf-8 -*-
"""Reproducible, synthetic before/after Excel I/O benchmark.

Run: .venv/Scripts/python.exe tests/bench_excel_io.py --rows 12000 --repeat 3
The benchmark creates temporary workbooks and removes them on exit.
"""
from __future__ import annotations

import argparse
import json
import statistics
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))

import openpyxl
import xlsxwriter
import xlrd
import xlwt
from openpyxl import Workbook, load_workbook
from python_calamine import CalamineWorkbook

from scan2excel import fast_excel


def matrix(rows: int, cols: int):
    headers = [f"列{i:02d}" for i in range(cols)]
    yield headers
    for r in range(1, rows + 1):
        yield [f"ID-{r:06d}", f"名称-{r}", f"地区-{r % 31:02d}"] + [
            round(r * (c + 1) / 7.0, 2) for c in range(cols - 3)]


def make_xlsx(path: Path, rows: int, cols: int) -> None:
    writer = xlsxwriter.Workbook(str(path), {"constant_memory": True})
    sheet = writer.add_worksheet("数据")
    for r, row in enumerate(matrix(rows, cols)):
        for c, value in enumerate(row):
            if isinstance(value, (int, float)):
                sheet.write_number(r, c, value)
            else:
                sheet.write_string(r, c, value)
    writer.close()


def make_xls(path: Path, rows: int, cols: int) -> None:
    book = xlwt.Workbook()
    sheet = book.add_sheet("数据")
    for r, row in enumerate(matrix(rows, cols)):
        for c, value in enumerate(row):
            sheet.write(r, c, value)
    book.save(str(path))


def make_names_xlsx(path: Path, count: int) -> None:
    writer = xlsxwriter.Workbook(str(path))
    sheet = writer.add_worksheet("模板页")
    sheet.write(0, 0, "科目代码")
    for i in range(count):
        writer.define_name(f"区域_{i:04d}", f"='模板页'!$A$1:$A${i + 1}")
    writer.close()


def timed(fn, repeat: int):
    values = []
    result = None
    for _ in range(repeat):
        start = time.perf_counter()
        result = fn()
        values.append(time.perf_counter() - start)
    return statistics.median(values), result


def read_calamine(path: Path):
    with CalamineWorkbook.from_path(path) as workbook:
        sheet = workbook.get_sheet_by_name("数据")
        rows = sheet.to_python(skip_empty_area=False)
    return len(rows), len(rows[0]) if rows else 0, rows[-1][-1] if rows else None


def read_openpyxl(path: Path):
    book = load_workbook(path, read_only=True, data_only=True)
    sheet = book["数据"]
    last = None
    rows = 0
    cols = 0
    for row in sheet.iter_rows(values_only=True):
        rows += 1
        cols = len(row)
        last = row[-1] if row else None
    book.close()
    return rows, cols, last


def read_xlrd(path: Path):
    book = xlrd.open_workbook(str(path), on_demand=True)
    sheet = book.sheet_by_name("数据")
    rows = sheet.nrows
    cols = sheet.ncols
    last = sheet.cell_value(rows - 1, cols - 1) if rows and cols else None
    book.release_resources()
    return rows, cols, last


def read_names_xml(path: Path) -> int:
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    names = root.find(f"{{{ns}}}definedNames")
    return len(names) if names is not None else 0


def read_names_openpyxl(path: Path) -> int:
    book = load_workbook(path, read_only=True, data_only=False)
    result = len(book.defined_names)
    book.close()
    return result


def read_names_app_metadata(path: Path) -> tuple[int, int]:
    sheets, names = fast_excel._ooxml_metadata(path)
    formulas = sum(len(sheet["formulas"]) for sheet in sheets.values())
    return len(names), formulas


def write_openpyxl(path: Path, rows: int, cols: int, write_only: bool):
    book = Workbook(write_only=write_only)
    sheet = book.create_sheet("数据") if write_only else book.active
    sheet.title = "数据"
    for row in matrix(rows, cols):
        sheet.append(row)
    book.save(path)


def write_xlsxwriter(path: Path, rows: int, cols: int):
    fast_book = fast_excel.FastWorkbook()
    fast_sheet = fast_excel.FastWorksheet("数据")
    fast_sheet.rows = list(matrix(rows, cols))
    fast_book.add_sheet(fast_sheet)
    fast_excel.write_workbook(path, fast_book, backend="fast")


def write_backend(path: Path, rows: int, cols: int, backend: str):
    fast_book = fast_excel.FastWorkbook()
    fast_book.add_sheet(fast_excel.FastWorksheet("数据", list(matrix(rows, cols))))
    fast_excel.write_workbook(path, fast_book, backend=backend)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=12000,
                        help="data rows excluding header (default: 12000)")
    parser.add_argument("--cols", type=int, default=18)
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--names", type=int, default=250)
    parser.add_argument("--temp-root", type=Path, default=HERE,
                        help="temporary output directory")
    args = parser.parse_args()
    row_count, col_count = args.rows + 1, args.cols
    results = {"rows": row_count, "cols": col_count, "repeat": args.repeat}

    args.temp_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ocr-excel-bench-",
                                     dir=args.temp_root) as td:
        root = Path(td)
        xlsx = root / "wide.xlsx"
        xls = root / "wide.xls"
        names = root / "names.xlsx"
        make_xlsx(xlsx, args.rows, args.cols)
        make_xls(xls, args.rows, args.cols)
        make_names_xlsx(names, args.names)

        # Correctness guard: every engine sees the same dimensions and final value.
        cal_xlsx = read_calamine(xlsx)
        open_xlsx = read_openpyxl(xlsx)
        if cal_xlsx != open_xlsx:
            raise AssertionError(("xlsx read mismatch", cal_xlsx, open_xlsx))
        cal_xls = read_calamine(xls)
        old_xls = read_xlrd(xls)
        if cal_xls != old_xls:
            raise AssertionError(("xls read mismatch", cal_xls, old_xls))

        measurements = {}
        measurements["read_xlsx_calamine_s"], _ = timed(
            lambda: read_calamine(xlsx), args.repeat)
        measurements["read_xlsx_openpyxl_read_only_s"], _ = timed(
            lambda: read_openpyxl(xlsx), args.repeat)
        measurements["read_xls_calamine_s"], _ = timed(
            lambda: read_calamine(xls), args.repeat)
        measurements["read_xls_xlrd_s"], _ = timed(
            lambda: read_xlrd(xls), args.repeat)
        measurements["read_xlsx_fast_backend_s"], _ = timed(
            lambda: fast_excel.read_workbook(xlsx, backend="fast"), args.repeat)
        measurements["read_xlsx_compat_backend_s"], _ = timed(
            lambda: fast_excel.read_workbook(xlsx, backend="compat"), args.repeat)
        measurements["read_xls_fast_backend_s"], _ = timed(
            lambda: fast_excel.read_workbook(xls, backend="fast"), args.repeat)
        measurements["read_xls_compat_backend_s"], _ = timed(
            lambda: fast_excel.read_workbook(xls, backend="compat"), args.repeat)

        expected_names = args.names
        if read_names_xml(names) != expected_names:
            raise AssertionError("OOXML name count mismatch")
        if read_names_openpyxl(names) != expected_names:
            raise AssertionError("OpenPyXL name count mismatch")
        if read_names_app_metadata(names)[0] != expected_names:
            raise AssertionError("application metadata name count mismatch")
        measurements["read_names_ooxml_only_s"], _ = timed(
            lambda: read_names_xml(names), args.repeat)
        measurements["read_names_openpyxl_s"], _ = timed(
            lambda: read_names_openpyxl(names), args.repeat)
        measurements["read_template_metadata_ooxml_s"], _ = timed(
            lambda: read_names_app_metadata(names), args.repeat)

        writer_specs = [
            ("write_xlsxwriter_s", write_xlsxwriter),
            ("write_fast_auto_s", lambda p, r, c:
             write_backend(p, r, c, "fast")),
            ("write_compat_backend_s", lambda p, r, c:
             write_backend(p, r, c, "compat")),
            ("write_openpyxl_normal_s",
             lambda p, r, c: write_openpyxl(p, r, c, False)),
            ("write_openpyxl_write_only_s",
             lambda p, r, c: write_openpyxl(p, r, c, True)),
        ]
        output_files = {}
        for key, writer_fn in writer_specs:
            out = root / f"{key}.xlsx"
            measurements[key], _ = timed(
                lambda out=out, writer_fn=writer_fn:
                writer_fn(out, args.rows, args.cols), args.repeat)
            output_files[key] = out
            if read_calamine(out) != cal_xlsx:
                raise AssertionError(f"writer output mismatch: {key}")
        measurements["output_sizes_bytes"] = {
            name: path.stat().st_size for name, path in output_files.items()}
        results["measurements"] = measurements
        print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

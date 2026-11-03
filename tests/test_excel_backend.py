# -*- coding: utf-8 -*-
"""Fast/compat Excel backend parity and settings persistence checks."""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))

import xlsxwriter

from scan2excel.fast_excel import (CellRange, FastWorkbook, FastWorksheet,
                                   consume_fallback_events, get_backend_mode,
                                   read_workbook, set_backend_mode,
                                   write_workbook)
from scan2excel.web_app import WebApi, _excel_io_operation


def _make_template(path: Path) -> None:
    book = xlsxwriter.Workbook(str(path))
    data = book.add_worksheet("模板页")
    data.merge_range("A1:B1", "月计表")
    data.write("A2", "指标")
    data.write_formula("B2", "=SUM(B3:B4)")
    data.write_number("A3", 1001)
    data.write_number("B3", 2.5)
    data.write_number("A4", 1002)
    data.write_number("B4", 3.5)
    data.conditional_format("B3:B4", {
        "type": "formula", "criteria": "=$B3<0",
        "format": book.add_format({"font_color": "#CC0000"}),
    })
    geo = book.add_worksheet("几何")
    geo.write("A1", '{"模板页":{"col_fracs":[0,1],"row_fracs":[0,1]}}')
    geo.very_hidden()
    book.define_name("数字区域", "='模板页'!$B$3:$B$4")
    book.close()


def main() -> int:
    original_mode = get_backend_mode()
    try:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source = root / "source.xlsx"
            _make_template(source)

            snapshots = {}
            for mode in ("auto", "fast", "compat"):
                set_backend_mode(mode)
                workbook = read_workbook(source, include_metadata=True)
                page = workbook["模板页"]
                assert page.cell(1, 1).value == "月计表"
                assert page.cell(3, 2).value == 2.5
                assert page.formula_cells[(2, 2)] == "=SUM(B3:B4)"
                assert page.merged_cells.ranges
                assert any(name.name == "数字区域" for name in workbook.defined_names)
                assert page.conditional_formats
                assert workbook["几何"].state == "veryHidden"
                output = root / f"{mode}.xlsx"
                write_workbook(output, workbook, backend=mode,
                               writer="xlsxwriter" if mode in ("auto", "fast") else None)
                saved = read_workbook(output, include_metadata=True,
                                      backend="fast")
                snapshots[mode] = (
                    saved["模板页"].cell(1, 1).value,
                    saved["模板页"].formula_cells[(2, 2)],
                    saved["模板页"].cell(4, 2).value,
                    saved["几何"].state,
                    [name.formula for name in saved.defined_names
                     if name.name == "数字区域"],
                )
            assert snapshots["auto"] == snapshots["fast"] == snapshots["compat"], snapshots

            plain_book = FastWorkbook()
            plain_book.add_sheet(FastWorksheet(
                "普通数据", rows=[["代码", "金额"], ["1001", 12.5]]))
            auto_output = root / "auto_fast_success.xlsx"
            consume_fallback_events()
            write_workbook(auto_output, plain_book, backend="auto")
            auto_saved = read_workbook(auto_output, backend="auto")
            assert auto_saved["普通数据"].cell(2, 2).value == 12.5
            assert not consume_fallback_events(), "自动模式成功时不应触发兼容回退"

            consume_fallback_events()
            with patch("scan2excel.fast_excel._read_fast",
                       side_effect=RuntimeError("模拟快速读取异常")):
                fallback = read_workbook(source, include_metadata=True,
                                         backend="auto")
            assert fallback["模板页"].cell(1, 1).value == "月计表"
            read_notes = consume_fallback_events()
            assert any("已回退兼容引擎" in note for note in read_notes), read_notes

            fallback_book = FastWorkbook()
            merged = FastWorksheet("回退写入", merged_ranges=[CellRange(1, 1, 1, 2)])
            merged.cell(1, 1, "兼容写入成功")
            fallback_book.add_sheet(merged)
            fallback_output = root / "auto_fallback.xlsx"
            with patch("scan2excel.fast_excel._write_xlsxwriter",
                       side_effect=RuntimeError("模拟快速写入异常")):
                write_workbook(fallback_output, fallback_book, backend="auto")
            saved_fallback = read_workbook(fallback_output, backend="compat")
            assert saved_fallback["回退写入"].cell(1, 1).value == "兼容写入成功"
            write_notes = consume_fallback_events()
            assert any("已回退兼容引擎" in note for note in write_notes), write_notes

            api = WebApi(root / "api")
            logged_read = _excel_io_operation(
                lambda instance: read_workbook(source, include_metadata=True,
                                               backend="auto"))
            with patch("scan2excel.fast_excel._read_fast",
                       side_effect=RuntimeError("模拟快速读取异常")):
                logged_read(api)
            assert any("已回退兼容引擎" in entry["text"]
                       for entry in api.state["log"]), api.state["log"]

        with tempfile.TemporaryDirectory() as td:
            WebApi._LAST_PATHS = None
            first = WebApi(Path(td))
            result = first.set_excel_io_backend("auto")
            assert result["excel_io_backend"] == "auto"
            WebApi._LAST_PATHS = None
            second = WebApi(Path(td))
            assert second.state["excel_io_backend"] == "auto"
            result = second.set_excel_io_backend("compat")
            assert result["excel_io_backend"] == "compat"
            WebApi._LAST_PATHS = None
            third = WebApi(Path(td))
            assert third.state["excel_io_backend"] == "compat"
            third.set_excel_io_backend("auto")
        print("✅ auto/fast/compat 公式、命名区域、条件格式、合并、隐藏几何往返一致")
        print("✅ 自动模式成功时直接使用快速读写路径")
        print("✅ 自动模式读取/写入失败会回退兼容引擎并记录日志")
        print("✅ 三档切换可保存并在 WebApi 重启后恢复")
    finally:
        set_backend_mode(original_mode)
        WebApi._LAST_PATHS = None
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

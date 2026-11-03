# -*- coding: utf-8 -*-
"""xlsx 模板加载/命名区域/几何回写 单元测试（纯 xlsx，不含 OCR）。

覆盖"什么样的 xlsx 可以放进模板库"的三种写法：
  ① 工作表级命名区域（每个 sheet 各自定义 数字区域/结构区域）——推荐
  ② 工作簿级命名区域（跨表 union 引用）
  ③ 完全没定义 —— 启发式推断（并标记 inferred_regions，载入时提示）
另验证 capture_geometry 回写隐藏"几何"表。
用法：.venv/Scripts/python.exe tests/test_template_xlsx.py
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))

import xlsxwriter                                                  # noqa: E402
from scan2excel.fast_excel import read_workbook, set_backend_mode   # noqa: E402

from scan2excel.xlsx_template import (GEO_SHEET, capture_geometry,  # noqa: E402
                                      load_xlsx_template)

HEADER = ["序号", "物料编码", "物料名称", "规格型号", "单位", "数量"]
DATA = [[str(i), f"WL-100{i}", "件" * i, f"M8x{i}", "个", str(i * 10)]
        for i in range(1, 11)]
ROWS = [HEADER] + DATA


def _create_workbook(path: Path, sheets, names=()):
    wb = xlsxwriter.Workbook(str(path))
    for title, rows in sheets:
        ws = wb.add_worksheet(title)
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                ws.write(r, c, value)
    for name, formula in names:
        wb.define_name(name, formula)
    wb.close()


def case_sheet_scoped(path: Path) -> None:
    _create_workbook(path, [("第1页", ROWS), ("第2页", ROWS)], [
        ("'第1页'!数字区域", "='第1页'!$F$2:$F$11"),
        ("'第1页'!结构区域", "='第1页'!$A$1:$E$11"),
        ("'第2页'!数字区域", "='第2页'!$F$2:$F$11"),
        ("'第2页'!结构区域", "='第2页'!$A$1:$E$11"),
    ])

    tpl = load_xlsx_template(path)
    assert len(tpl.pages) == 2, tpl.pages
    for pg in tpl.pages:
        assert len(pg.num_cells) == 10, (pg.page_name, len(pg.num_cells))
        assert (1, 5) in pg.num_cells               # F2（0 基）
        assert (0, 0) in pg.struct_cells            # A1
        assert pg.header_rows == 1
        assert pg.value_cols == [5]
    print("✅ ① 工作表级命名区域：两页各自解析正确")


def case_workbook_scoped(path: Path) -> None:
    _create_workbook(path, [("第1页", ROWS)], [
        ("数字区域", "='第1页'!$F$2:$F$11"),
        ("结构区域", "='第1页'!$A$1:$E$11"),
    ])

    tpl = load_xlsx_template(path)
    pg = tpl.pages[0]
    assert len(pg.num_cells) == 10 and pg.value_cols == [5]
    print("✅ ② 工作簿级命名区域（union 写法）解析正确")


def case_heuristic(path: Path) -> None:
    _create_workbook(path, [("第1页", ROWS)])

    tpl = load_xlsx_template(path)
    pg = tpl.pages[0]
    assert tpl.inferred_regions is True, "应标记为启发式推断"
    assert len(pg.num_cells) >= 10, len(pg.num_cells)
    assert pg.value_cols == [5], pg.value_cols    # 仅"数量"列有数字内容
    print("✅ ③ 无命名区域：启发式推断（inferred_regions=True）")


def case_geometry(path: Path) -> None:
    _create_workbook(path, [("第1页", ROWS)], [
        ("'第1页'!数字区域", "='第1页'!$F$2:$F$11"),
        ("'第1页'!结构区域", "='第1页'!$A$1:$E$11"),
    ])

    tpl = load_xlsx_template(path)
    assert not tpl.pages[0].col_fracs, "初始不应有几何"
    ok = capture_geometry(path, "第1页", list(range(0, 601, 100)),
                          list(range(0, 111, 10)))
    assert ok, "几何回写失败"
    tpl2 = load_xlsx_template(path)
    pg2 = tpl2.pages[0]
    assert len(pg2.col_fracs) == 7 and len(pg2.row_fracs) == 12
    assert abs(pg2.col_fracs[-1] - 1.0) < 1e-6
    wb2 = read_workbook(path, include_metadata=True)
    assert GEO_SHEET in wb2.sheetnames
    assert wb2[GEO_SHEET].state == "veryHidden"
    print("✅ ④ capture_geometry：几何比例写入隐藏表并可回读")


def main() -> int:
    for backend in ("auto", "fast", "compat"):
        set_backend_mode(backend)
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            case_sheet_scoped(d / "a.xlsx")
            case_workbook_scoped(d / "b.xlsx")
            case_heuristic(d / "c.xlsx")
            case_geometry(d / "d.xlsx")
        print(f"✅ {backend} 引擎模板回归通过")
    set_backend_mode("auto")
    print("\n总体结论：✅ 三种 Excel 模式模板加载全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())

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

from openpyxl import Workbook, load_workbook                      # noqa: E402
from openpyxl.workbook.defined_name import DefinedName            # noqa: E402

from scan2excel.xlsx_template import (GEO_SHEET, capture_geometry,  # noqa: E402
                                      load_xlsx_template)

HEADER = ["序号", "物料编码", "物料名称", "规格型号", "单位", "数量"]
DATA = [[str(i), f"WL-100{i}", "件" * i, f"M8x{i}", "个", str(i * 10)]
        for i in range(1, 11)]
ROWS = [HEADER] + DATA


def _fill(ws, rows):
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)


def _add_name(ws_or_wb, name, attr):
    dn = DefinedName(name, attr_text=attr)
    try:
        ws_or_wb.defined_names[name] = dn      # openpyxl >= 3.1 dict 形式
    except Exception:
        ws_or_wb.defined_names.add(dn)         # 旧式 add 形式


def case_sheet_scoped(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "第1页"
    _fill(ws, ROWS)
    _add_name(ws, "数字区域", "$F$2:$F$11")
    _add_name(ws, "结构区域", "$A$1:$E$11")
    ws2 = wb.create_sheet("第2页")
    _fill(ws2, ROWS)
    _add_name(ws2, "数字区域", "$F$2:$F$11")     # 同名，工作表级
    _add_name(ws2, "结构区域", "$A$1:$E$11")
    wb.save(path)

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
    wb = Workbook()
    ws = wb.active
    ws.title = "第1页"
    _fill(ws, ROWS)
    _add_name(wb, "数字区域", "'第1页'!$F$2:$F$11")
    _add_name(wb, "结构区域", "'第1页'!$A$1:$E$11")
    wb.save(path)

    tpl = load_xlsx_template(path)
    pg = tpl.pages[0]
    assert len(pg.num_cells) == 10 and pg.value_cols == [5]
    print("✅ ② 工作簿级命名区域（union 写法）解析正确")


def case_heuristic(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "第1页"
    _fill(ws, ROWS)                              # 无任何命名区域
    wb.save(path)

    tpl = load_xlsx_template(path)
    pg = tpl.pages[0]
    assert tpl.inferred_regions is True, "应标记为启发式推断"
    assert len(pg.num_cells) >= 10, len(pg.num_cells)
    assert pg.value_cols == [5], pg.value_cols    # 仅"数量"列有数字内容
    print("✅ ③ 无命名区域：启发式推断（inferred_regions=True）")


def case_geometry(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "第1页"
    _fill(ws, ROWS)
    _add_name(ws, "数字区域", "$F$2:$F$11")
    _add_name(ws, "结构区域", "$A$1:$E$11")
    wb.save(path)

    tpl = load_xlsx_template(path)
    assert not tpl.pages[0].col_fracs, "初始不应有几何"
    ok = capture_geometry(path, "第1页", list(range(0, 601, 100)),
                          list(range(0, 111, 10)))
    assert ok, "几何回写失败"
    tpl2 = load_xlsx_template(path)
    pg2 = tpl2.pages[0]
    assert len(pg2.col_fracs) == 7 and len(pg2.row_fracs) == 12
    assert abs(pg2.col_fracs[-1] - 1.0) < 1e-6
    wb2 = load_workbook(path)
    assert GEO_SHEET in wb2.sheetnames
    wb2.close()
    print("✅ ④ capture_geometry：几何比例写入隐藏表并可回读")


def main() -> int:
    with tempfile.TemporaryDirectory() as td:
        d = Path(td)
        case_sheet_scoped(d / "a.xlsx")
        case_workbook_scoped(d / "b.xlsx")
        case_heuristic(d / "c.xlsx")
        case_geometry(d / "d.xlsx")
    print("\n总体结论：✅ xlsx 模板加载全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())

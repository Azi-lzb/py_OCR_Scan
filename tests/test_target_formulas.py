# -*- coding: utf-8 -*-
"""目标工作簿追加：必须就地打补丁，公式/批注/样式不能因整本重建而丢失。"""
from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))

import xlsxwriter
from openpyxl.styles import Side

from scan2excel.data_process import append_to_target
from scan2excel.fast_excel import read_workbook


def _make_target(path: Path) -> None:
    """造一本带公式主表 + 数据表的簿（模拟 transforms/国库报表.xlsx）。

    同时带批注、合并区域和样式，用来验证写入走"就地补丁"而不是整本重建——
    重建会丢批注、主题，样式表也会大幅缩水。"""
    book = xlsxwriter.Workbook(str(path))
    main = book.add_worksheet("主表")
    style = book.add_format({"bold": True, "bg_color": "#DDEBF7",
                             "border": 1, "num_format": "#,##0.00"})
    main.write("A1", "指标代码", style)
    main.write("B1", "指标名称", style)
    main.write("C1", "惠州市", style)
    main.write("D1", "市本部", style)
    main.write_number("A2", 11001)
    main.write("B2", "一、资产合计")
    main.write_number("C2", 100.0)
    main.write_number("D2", 200.0)
    main.write_comment("B2", "这是批注，重建会丢")
    # 主表公式：合计与校验列
    main.write("E1", "合计")
    main.write_formula("E2", "=C2+D2", None, 300.0)
    main.write_formula("F2", '=IF(E2=300,"平","不平")', None, "平")
    main.merge_range("G1:H1", "校验区")

    data = book.add_worksheet("源数据时序")
    header = ["工作簿名", "工作表名", "数据日期", "行头路径",
              "本期余额_借方", "科目代码_科目代码"]
    for c, name in enumerate(header):
        data.write(0, c, name)
    data.write_row(1, 0, ["旧工作簿", "旧表", "2026-09-01", "甲.乙", 5.0, "11001"])
    book.close()


def main() -> int:
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        target = root / "国库报表.xlsx"
        _make_target(target)

        before = read_workbook(target, include_metadata=True)
        before_formulas = dict(before["主表"].formula_cells)
        before_values = [[before["主表"].cell(r, c).value
                          for c in range(1, before["主表"].max_column + 1)]
                         for r in range(1, before["主表"].max_row + 1)]
        assert len(before_formulas) == 2, before_formulas

        header = ["工作簿名", "工作表名", "数据日期", "行头路径",
                  "本期余额_借方", "科目代码_科目代码"]
        row = ["新工作簿", "新表", "2026-10-01", "甲.乙", 9.0, "11001"]

        before_parts = set(zipfile.ZipFile(target).namelist())
        before_styles = zipfile.ZipFile(target).getinfo("xl/styles.xml").file_size

        stat = append_to_target(target, "源数据时序", header, [row],
                                ["数据日期", "行头路径"],
                                ["工作簿名", "工作表名", "数据日期", "行头路径"])
        assert stat["written"] and stat["added"] == 1, stat
        # 必须用 openpyxl 追加，不能整本重建
        assert stat.get("mode") == "openpyxl", stat

        # 样式表必须被保留（整本重建会从 20KB 缩到 1KB）
        after_styles = zipfile.ZipFile(target).getinfo("xl/styles.xml").file_size
        assert after_styles >= before_styles * 0.9, (
            f"样式表大幅缩水：{before_styles} -> {after_styles}")

        # 公式缓存会被 openpyxl 清空，因此必须带 fullCalcOnLoad 让 Excel
        # 打开时自动重算；否则用户会看到空白的主表。
        with zipfile.ZipFile(target) as archive:
            workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        assert 'fullCalcOnLoad="1"' in workbook_xml, workbook_xml

        after = read_workbook(target, include_metadata=True)
        main_sheet = after["主表"]
        assert main_sheet.formula_cells == before_formulas, (
            "主表公式被写成了静态值", main_sheet.formula_cells)

        after_values = [[main_sheet.cell(r, c).value
                         for c in range(1, main_sheet.max_column + 1)]
                        for r in range(1, main_sheet.max_row + 1)]
        # 公式单元格的缓存值被 openpyxl 清空，其余（非公式）数据必须逐格不变
        formula_cells = set(before_formulas)
        for r_idx, row_values in enumerate(before_values, 1):
            for c_idx, value in enumerate(row_values, 1):
                if (r_idx, c_idx) in formula_cells:
                    continue
                assert after_values[r_idx - 1][c_idx - 1] == value, (
                    f"非公式单元格 ({r_idx},{c_idx}) 被改动：{value!r} -> "
                    f"{after_values[r_idx-1][c_idx-1]!r}")

        # 公式缓存值会被 openpyxl 清空（已知代价，靠 fullCalcOnLoad 恢复），
        # 但公式文本本身和合并区域必须原样保留。
        assert main_sheet.merged_cells.ranges, "合并区域丢失"

        # 追加只发生在数据表上
        assert after["源数据时序"].max_row == 3, after["源数据时序"].max_row
        assert after["源数据时序"].rows[-1][2] == "2026-10-01"

        # 再追加一次同样的键：去重生效，主表仍不动
        again = append_to_target(target, "源数据时序", header, [row],
                                 ["数据日期", "行头路径"],
                                 ["工作簿名", "工作表名", "数据日期", "行头路径"])
        assert again["added"] == 0, again
        final = read_workbook(target, include_metadata=True)
        assert final["主表"].formula_cells == before_formulas

        _check_cleared_target(root)
        _check_code_columns_text(root)

    print("test_target_formulas OK")
    return 0


def _check_cleared_target(root: Path) -> None:
    """用户清空目标表后应能正常写入。

    两个曾经踩过的坑：
    1) 整表清空（连表头）时 old_header 为空 → 误报 header_mismatch，
       用户清空目标表本来就是为了重新灌数据。
    2) 只清内容、保留边框时 openpyxl 的 max_row 仍算到格式行（实测 50），
       按 max_row+1 会把数据追加到第 51 行，看起来"显示新增了行但表里没有"。
    """
    from openpyxl import Workbook, load_workbook

    header = ["工作簿名", "工作表名", "数据日期", "行头路径",
              "本期余额_借方", "科目代码_科目代码"]
    required = ["工作簿名", "工作表名", "数据日期", "行头路径"]
    dedup = ["数据日期", "行头路径"]
    rows = [[f"wb{i}", "sh", "2026-10-01", f"p{i}", i, f"1100{i}"]
            for i in range(3)]

    # 场景 1：完全空表（连表头都没有）
    blank = root / "blank_target.xlsx"
    book = Workbook()
    book.active.title = "源数据时序"
    book.save(blank)
    stat = append_to_target(blank, "源数据时序", header, rows, dedup, required)
    assert stat["written"] and stat["added"] == 3, stat
    sheet = load_workbook(blank)["源数据时序"]
    assert sheet.max_row == 4, sheet.max_row
    assert sheet.cell(1, 1).value == "工作簿名", "空表必须先写表头"
    assert sheet.cell(2, 1).value == "wb0", "数据应从第 2 行开始"

    # 场景 2：数据被清空但保留格式（openpyxl max_row 仍很大）
    formatted = root / "formatted_target.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "源数据时序"
    for column, value in enumerate(header, 1):
        sheet.cell(1, column, value)
    for row_index in range(2, 51):
        sheet.cell(row_index, 1, f"old{row_index}")
        sheet.cell(row_index, 1).border = Side(style="thin")
    book.save(formatted)
    book = load_workbook(formatted)
    sheet = book["源数据时序"]
    for row_index in range(2, 51):
        sheet.cell(row_index, 1).value = None     # 只清值，保留边框
    book.save(formatted)
    assert load_workbook(formatted)["源数据时序"].max_row >= 50, "前提：格式行仍占位"

    stat = append_to_target(formatted, "源数据时序", header, rows, dedup, required)
    assert stat["written"] and stat["added"] == 3, stat
    sheet = load_workbook(formatted)["源数据时序"]
    written_at = [r for r in range(1, sheet.max_row + 1)
                  if sheet.cell(r, 1).value == "wb0"]
    assert written_at == [2], f"数据被写到格式行之后：第 {written_at} 行"


def _check_code_columns_text(root: Path) -> None:
    """「行头路径」「科目代码」必须存成文本。

    源表同一列常混着文本和数值（'11347' 与 11344.0）。若写成数值，Excel 显示
    11344，读回来与 '11344' 对不上，去重和跨月比对都会失效。"""
    from openpyxl import Workbook, load_workbook

    header = ["工作簿名", "数据日期", "行头路径", "本期余额_借方",
              "科目代码_科目代码"]
    required = ["工作簿名", "数据日期", "行头路径"]
    path = root / "code_text.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "源数据时序"
    for column, value in enumerate(header, 1):
        sheet.cell(1, column, value)
    sheet.cell(2, 1, "old")
    sheet.cell(2, 2, "2026-09-01")
    sheet.cell(2, 3, "999")          # 已有行保持原样
    book.save(path)

    rows = [
        ["wb", "2026-11-01", 11344.0, 1.0, 11344.0],   # 数值代码
        ["wb", "2026-11-01", "11347", 2.0, "11347"],   # 文本代码
        ["wb", "2026-11-01", "1123P", 3.0, "1123P"],   # 字母结尾
    ]
    stat = append_to_target(path, "源数据时序", header, rows,
                            ["数据日期", "行头路径"], required)
    assert stat["written"] and stat["added"] == 3, stat

    sheet = load_workbook(path)["源数据时序"]
    for row_index, expected in ((3, "11344"), (4, "11347"), (5, "1123P")):
        for column in (3, 5):
            cell = sheet.cell(row_index, column)
            assert isinstance(cell.value, str), (
                f"({row_index},{column}) 不是文本：{cell.value!r}")
            assert cell.number_format == "@", (
                f"({row_index},{column}) 缺少文本格式：{cell.number_format}")
        assert sheet.cell(row_index, 3).value == expected, (
            f"行头路径未归一：{sheet.cell(row_index, 3).value!r} != {expected!r}")
    # 数值 11344.0 必须落成 '11344'，不能是 '11344.0'
    assert sheet.cell(3, 3).value == "11344", sheet.cell(3, 3).value
    assert sheet.cell(3, 5).value == "11344", sheet.cell(3, 5).value


if __name__ == "__main__":
    raise SystemExit(main())

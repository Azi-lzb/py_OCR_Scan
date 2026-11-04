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
        target = Path(td) / "国库报表.xlsx"
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

    print("test_target_formulas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

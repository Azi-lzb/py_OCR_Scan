"""Public package seeds must preserve structure and remove business values."""
import io
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "shell-flask"))
from prepare_public_release import clean_workbook


def payload(wb):
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "一页"
    ws.append(["科目", "名称", "余额"])
    ws.append([272, "科目名", 12345.67])
    wb.defined_names.add(DefinedName("数字区域", attr_text="一页!$C$2:$C$3"))
    wb.create_sheet("几何")["A1"] = '{"一页":{"row_fracs":[0,0.5,1]}}'
    clean = load_workbook(io.BytesIO(clean_workbook(payload(wb), "templates/Template.xlsx")))
    assert clean["一页"].max_row == 2, "Cleaning must not extend the template geometry"
    assert clean["一页"]["C2"].value is None
    assert clean["一页"]["A2"].value == 272
    assert clean["几何"]["A1"].value == wb["几何"]["A1"].value
    assert clean.defined_names["数字区域"].attr_text == wb.defined_names["数字区域"].attr_text

    wb = Workbook()
    wb.active.title = "源数据时序"
    wb.active.append(["工作簿名", "工作表名", "数据日期", "行头路径", "余额"])
    wb.active.append(["private", "private", "2026-01-01", "272", 100])
    wb.create_sheet("私人汇总")["A1"] = 999
    clean = load_workbook(io.BytesIO(clean_workbook(payload(wb), "transforms/国库报表.xlsx")))
    assert clean.sheetnames == ["源数据时序"] and clean.active.max_row == 1
    assert list(clean.active.values) == [tuple(c.value for c in wb["源数据时序"][1])]

    wb = Workbook()
    wb.active["A1"] = r"C:\Users\private\core\data\transforms\国库报表.xlsx"
    clean = load_workbook(io.BytesIO(clean_workbook(payload(wb), "config/config.xlsx")))
    assert clean.active["A1"].value == r"C:\OCR_Tools\core\data\transforms\国库报表.xlsx"
    wb.active["A1"] = r"D:\private\other.xlsx"
    try:
        clean_workbook(payload(wb), "config/config.xlsx")
    except ValueError:
        pass
    else:
        raise AssertionError("Unrecognized personal paths must block public preparation")
    print("Public seed structure, blank amounts, empty targets and path checks passed")


if __name__ == "__main__":
    main()

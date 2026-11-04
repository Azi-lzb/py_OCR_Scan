# -*- coding: utf-8 -*-
"""批量 Excel 转换：转换前必须强制全量重算，不能把过期公式缓存值带进输出。

需要 Windows + Excel/WPS COM；非 Windows 或没有 Office 时跳过。
"""
from __future__ import annotations

import re
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))


def _make_manual_mode_stale(path: Path, good: Path, app) -> None:
    """用 Excel 存一份带 calcChain 的簿，再改写成"手动计算 + 过期缓存 999"。"""
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "S"
    ws.Range("A1").Value = 1
    ws.Range("A2").Value = 2
    ws.Range("A3").Formula = "=A1+A2"
    wb.SaveAs(str(good))
    wb.Close(SaveChanges=False)

    with zipfile.ZipFile(good) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                text = data.decode("utf-8")
                text = text.replace("<calcPr ", '<calcPr calcMode="manual" ')
                text = text.replace('fullCalcOnLoad="1"', 'fullCalcOnLoad="0"')
                data = text.encode("utf-8")
            elif item.filename.startswith("xl/worksheets/sheet"):
                text = data.decode("utf-8")
                data = re.sub(r'(<c r="A3"[^>]*>.*?<v>)3(</v>)',
                              r"\g<1>999\g<2>", text, flags=re.S).encode("utf-8")
            zout.writestr(item, data)


def main() -> int:
    if sys.platform != "win32":
        print("skip：批量 Office 转换只在 Windows 上可用")
        return 0

    try:
        import pythoncom
        import win32com.client
        from openpyxl import load_workbook
        from scan2excel.office_convert import _force_recalc, run_excel_convert
    except Exception as exc:
        print(f"skip：{exc}")
        return 0

    root = Path(tempfile.mkdtemp(prefix="conv_recalc_"))
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            print(f"skip：无法创建 Excel COM（{exc}）")
            return 0
        try:
            app.Visible = False
            app.DisplayAlerts = False

            stale = root / "stale.xlsx"
            _make_manual_mode_stale(stale, root / "good.xlsx", app)
            before = load_workbook(stale, data_only=True)["S"]["A3"].value
            assert before == 999, before   # 前提：源文件确实是过期缓存

            recalc_wb = app.Workbooks.Open(str(stale), ReadOnly=True)
            recalculated, method = _force_recalc(app)
            assert recalculated, "CalculateFullRebuild 应可用"
            assert method == "CalculateFullRebuild", method
            recalc_wb.Close(SaveChanges=False)

            # Calculate() 必须不在降级链里：它只重算"脏"单元格，刚打开的簿没有
            # 脏标记，会静默返回而不计算，还会骗过逐格比对。
            import scan2excel.office_convert as oc
            assert "Calculate" not in oc._RECALC_METHODS, oc._RECALC_METHODS
            assert oc._RECALC_METHODS == ("CalculateFullRebuild", "CalculateFull")
        finally:
            try:
                app.Quit()
            except Exception:
                pass

        res = run_excel_convert([stale], "xlsx")
        assert res["ok"] == 1, res
        assert res["recalc"] == 1, res
        assert res["recalc_failed"] == 0, res
        assert res["recalc_method"] == "CalculateFullRebuild", res
        assert res["verified"] == 1, res
        assert res["verify_failed"] == 0, res
        out = Path(res["done"][0]["out"])
        after = load_workbook(out, data_only=True)["S"]["A3"].value
        assert after == 3, f"输出仍带过期缓存：{after}"

        # 转出的 xlsx 必须是"已计算好"的版本：缓存值就是重算结果，
        # 下游 calamine 之类按缓存读取也能拿到正确数字。
        from python_calamine import CalamineWorkbook
        with CalamineWorkbook.from_path(out) as raw:
            cached_rows = raw.get_sheet_by_name("S").to_python(skip_empty_area=False)
        assert cached_rows[2][0] == 3, f"calamine 读到过期缓存：{cached_rows[2]}"

        # 输出应带 fullCalcOnLoad 标记，外部引擎再次打开时会全量重算
        with zipfile.ZipFile(out) as archive:
            workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        assert 'fullCalcOnLoad="1"' in workbook_xml, workbook_xml

        # CSV 也存公式显示值：源是手动计算模式时若不重算，会把陈旧缓存导出去
        # （实测应为 3 却写出 999）。CSV 无单元格结构，只重算、不做逐格比对。
        csv_res = run_excel_convert([stale], "csv")
        assert csv_res["ok"] == 1, csv_res
        assert csv_res["recalc"] == 1, csv_res
        csv_out = Path(csv_res["done"][0]["out"])
        csv_rows = csv_out.read_text(encoding="utf-8-sig").splitlines()
        assert csv_rows[2] == "3", f"CSV 导出了陈旧缓存：{csv_rows}"

        # 校验不通过时不丢弃文件：照常另存并给出警告，公式文本保留，
        # 用户打开工作簿即可重算（fullCalcOnLoad 已写入）。
        from unittest.mock import patch
        import scan2excel.office_convert as oc
        with patch.object(oc, "_compare", return_value=["模拟：S!A3 不一致"]):
            warned = run_excel_convert([stale], "xlsx")
        assert warned["ok"] == 1, warned
        assert warned["skip"] == 0, warned
        assert warned["verify_failed"] == 1, warned
        warned_item = warned["done"][0]
        assert warned_item["warning"] is True, warned_item
        assert warned_item["verified"] is False, warned_item
        assert warned_item["problems"], warned_item
        warned_out = Path(warned_item["out"])
        assert warned_out.is_file(), "校验不通过也必须保留输出文件"
        # 不留临时候选，目录里只有正式输出
        leftovers = [p.name for p in warned_out.parent.iterdir() if ".converting" in p.name]
        assert not leftovers, leftovers
        with zipfile.ZipFile(warned_out) as archive:
            warned_xml = archive.read("xl/workbook.xml").decode("utf-8")
        assert 'fullCalcOnLoad="1"' in warned_xml, "警告文件也必须带重算标记"

        # 重算失败 + 表内有公式 → 必须标记不可信（不能靠逐格比对放行：
        # 源与候选读到同一份陈旧缓存，_compare 必然"一致"）。
        from unittest.mock import patch as _patch
        with _patch.object(oc, "_force_recalc", return_value=(False, "")), \
             _patch.object(oc, "_compare", return_value=[]):
            untrusted = run_excel_convert([stale], "xlsx")
        assert untrusted["recalc_failed"] == 1, untrusted
        assert untrusted["recalc_untrusted"] == 1, untrusted
        assert untrusted["verified"] == 0, untrusted
        item = untrusted["done"][0]
        assert item["warning"] is True, item
        assert any("重算失败" in p for p in item["problems"]), item
    finally:
        pythoncom.CoUninitialize()

    print("test_convert_recalc OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

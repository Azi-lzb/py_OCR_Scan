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

        # 输出应带 fullCalcOnLoad 标记，外部引擎再次打开时会全量重算
        with zipfile.ZipFile(out) as archive:
            workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        assert 'fullCalcOnLoad="1"' in workbook_xml, workbook_xml
    finally:
        pythoncom.CoUninitialize()

    print("test_convert_recalc OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

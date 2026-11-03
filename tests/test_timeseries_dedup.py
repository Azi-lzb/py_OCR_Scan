# -*- coding: utf-8 -*-
"""时序表重复指标行合并：旧版数字代码 11001.0 ≠ 11001 造成整块副本追加。"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "core" / "src"))

from scan2excel.fast_excel import FastWorkbook, FastWorksheet, read_workbook, write_workbook
from scan2excel.timeseries import (_code_text, _merge_duplicate_rows, _normalize_code_column,
                                   _snapshot, _write_back, load_config)


def _book(rows_by_sheet):
    book = FastWorkbook()
    for title, rows in rows_by_sheet.items():
        book.add_sheet(FastWorksheet(title, rows))
    return book


def main() -> int:
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)

        # 首行 1-3 月有数，表尾副本行只有 2026.10（旧版追加的形态）
        rows = [["指标代码", "指标名称", "行次号", "2026.01", "2026.02", "2026.03", "2026.10"],
                [11001, "一、资产合计", 1, 100.0, 200.0, 300.0, None],
                [11002, "（一）贵金属", 2, 0, 0, 0, None],
                ["1123P", "7.中央应急流动性借款", 227, 0, 0, 0, None],
                ["11001", "一、资产合计", "1", None, None, None, 999.0],
                ["11002", "（一）贵金属", "2", None, None, None, 0],
                ["1123P", "7.中央应急流动性借款", "227", None, None, None, 0],
                ["11632", "（二十六）调平项（县级专用）", "1159", None, None, None, 0]]
        book = _book({"惠州市": rows, "校验结果": [["月份"], ["2026.10"]]})

        ws = book["惠州市"]
        cols, indicators = _snapshot(ws)
        assert len(indicators) == 4, indicators.keys()
        assert indicators["11001"]["dups"] == [5]
        assert indicators["11001"]["vals"][7] == 999.0
        assert indicators["11001"]["own"][4] == 100.0
        assert indicators["11001"]["name"] == "一、资产合计"

        orig_vals = {c: dict(i.get("own", i["vals"])) for c, i in indicators.items()}
        _write_back(book, "惠州市", indicators, cols, dict(cols), orig_vals, [])

        assert book["惠州市"].max_row == 5, book["惠州市"].rows
        merged = book["惠州市"].rows
        # 已有行的代码统一成文本（补录后代码不再变数值），副本行的非空格合并进来
        assert merged[1] == ["11001", "一、资产合计", 1, 100.0, 200.0, 300.0, 999.0], merged[1]
        assert merged[2] == ["11002", "（一）贵金属", 2, 0, 0, 0, 0], merged[2]
        assert merged[3] == ["1123P", "7.中央应急流动性借款", 227, 0, 0, 0, 0], merged[3]
        assert merged[4][0] == "11632" and merged[4][3:] == [None, None, None, 0], merged[4]

        # 未参与本次归集的 sheet 也要清理（只给了一个地区源文件的情况）
        book2 = _book({"惠州市": [["指标代码", "指标名称", "行次号", "2026.01"]],
                       "市本部": [["指标代码", "指标名称", "行次号", "2026.01", "2026.10"],
                                 [11001, "一、资产合计", 1, 5.0, None],
                                 ["11001", "一、资产合计", "1", None, 7.0]],
                       "校验结果": [["月份"]]})
        cleaned = _merge_duplicate_rows(book2, {})
        assert cleaned == ["市本部"], cleaned
        assert book2["市本部"].max_row == 2
        assert book2["市本部"].rows[1] == ["11001", "一、资产合计", 1, 5.0, 7.0]

        # 落盘后读回：代码是文本、无重复、值不变（两种引擎各验一次）
        out = root / "dedup.xlsx"
        write_workbook(out, book, backend="compat")
        for backend in ("compat", "fast"):
            try:
                saved = read_workbook(out, backend=backend)
            except Exception as exc:  # fast 引擎需要 calamine 元数据，缺失时跳过
                print(f"skip {backend}: {exc}")
                continue
            saved_rows = list(saved["惠州市"].iter_rows(values_only=True))
            codes = [str(row[0]).strip() for row in saved_rows[1:]]
            assert len(codes) == len(set(codes)), codes
            assert saved_rows[1][6] == 999.0, saved_rows[1]

        # 指标代码列：补录/归集后必须仍是文本，并挂 "@" 列格式
        text_book = _book({"惠州市": [["指标代码", "指标名称", "行次号", "2026.10"],
                                     [11001, "一、资产合计", 1, 5.0],
                                     ["1123P", "7.中央应急流动性借款", 227, 0]]})
        done = _normalize_code_column(text_book)
        assert done == ["惠州市"], done
        assert text_book["惠州市"].text_columns == {1}
        assert all(isinstance(text_book["惠州市"].cell(r, 1).value, str)
                   for r in (2, 3))

        text_out = root / "text_codes.xlsx"
        for backend in ("compat", "fast"):
            write_workbook(text_out, text_book, backend=backend)
            from python_calamine import CalamineWorkbook
            with CalamineWorkbook.from_path(text_out) as raw:
                raw_rows = raw.get_sheet_by_name("惠州市").to_python(skip_empty_area=False)
            assert raw_rows[1][0] == "11001", raw_rows[1]
            assert raw_rows[2][0] == "1123P", raw_rows[2]

        # 配置表里的代码被用户设成数值时，白名单/补录映射/父项都要照样命中
        cfg_book = FastWorkbook()
        cfg_book.add_sheet(FastWorksheet("归集配置", [
            ["配置项", "值"], ["容差", 0.01], ["补录父项", "11364;11001"]]))
        cfg_book.add_sheet(FastWorksheet("补录映射", [
            ["指标代码", "余额表代码", "取数列"], [11367, 512, 7]]))
        cfg_book.add_sheet(FastWorksheet("白名单", [
            ["国库配置项", "国库白名单", "", "会计配置项", "会计白名单"],
            ["", 11600, "", "", 11348],
            ["", "11362", "", "", 11001]]))
        cfg_path = root / "config_numeric.xlsx"
        write_workbook(cfg_path, cfg_book, backend="compat")
        cfg = load_config(cfg_path)
        assert cfg["whitelist"]["国库"] == {"11600", "11362"}, cfg["whitelist"]["国库"]
        assert cfg["whitelist"]["会计"] == {"11348", "11001"}, cfg["whitelist"]["会计"]
        assert cfg["backfill_map"] == [("11367", "512", 7)], cfg["backfill_map"]
        assert cfg["backfill_parents"] == "11364;11001", cfg["backfill_parents"]

        for raw, want in ((11600, "11600"), (11600.0, "11600"), ("11600", "11600"),
                          ("11600.0", "11600"), (" 11600 ", "11600"),
                          (None, ""), ("1123P", "1123P")):
            assert _code_text(raw) == want, (raw, _code_text(raw))

    print("test_timeseries_dedup OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

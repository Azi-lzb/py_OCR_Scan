# -*- coding: utf-8 -*-
"""用 openpyxl 向既有 .xlsx 目标表追加明细行。

为什么不整本重建（XlsxWriter）：重建会丢读不到的东西——实测目标簿丢批注
（xl/comments1.xml + vmlDrawing）、自定义文档属性、主题，样式表从 20KB 缩到
1KB。用户自己的报表不能这么洗。

openpyxl 会保留公式、批注、样式、主题，只重写自己认识的部分。唯一代价是保存
时把公式缓存值清空（<f>+E3</f><v></v>）：这些缓存是给人看的显示值，openpyxl
同时会写 fullCalcOnLoad="1"，Excel/WPS 打开时自动全量重算（实测 E2 从空恢复
为 33804983557.14）。因此不再做"抓原缓存再补回"的正则补丁——那种改动容易出错，
且本工具从不按缓存读取目标簿主表。

注意：openpyxl 重写的是整包，不是只动目标 sheet；但在语义上只新增了目标表的
明细行，其余内容逐项保留。
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, List, Optional, Sequence, Tuple

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

#: 表头名包含这些关键字的列按"代码"处理，写成文本并挂 "@" 格式。
_CODE_COLUMN_KEYS = ("行头路径", "科目代码", "指标代码")


def _code_as_text(value: Any) -> str:
    """11344.0 / 11344 / '11344.0' 一律收敛成 '11344'（同 data_process._code_norm）。"""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if re.fullmatch(r"[-+]?\d+\.0+", text):
        return text.split(".")[0]
    return text


def _force_text_columns(sheet, header: Sequence[Any],
                        row_indices: Sequence[int]) -> None:
    """把表头里属于"代码"的列，在指定行上统一成文本并挂 "@" 格式。

    「行头路径」「科目代码」写的是指标/科目代码：存成数值后 Excel 显示 11344，
    再读回来与 '11344' 对不上，去重和跨月比对都会失效。"""
    for column_index, name in enumerate(header, 1):
        label = str(name or "")
        if not any(key in label for key in _CODE_COLUMN_KEYS):
            continue
        for row_index in row_indices:
            cell = sheet.cell(row_index, column_index)
            if cell.value is not None and cell.value != "":
                cell.value = _code_as_text(cell.value)
            cell.number_format = "@"


def _last_content_row(sheet) -> int:
    """最后一个有实际内容（非空）的行号，1 基；整表无内容返回 0。

    不能用 sheet.max_row：openpyxl 会把"只剩格式、内容已清空"的行也算进去。
    用户清空目标表内容但保留边框时，max_row 仍是 50，直接写 max_row+1 会把
    新数据追加到第 51 行——看起来像"显示新增了行但表里没有"。"""
    for row_index in range(sheet.max_row, 0, -1):
        for column_index in range(1, sheet.max_column + 1):
            value = sheet.cell(row_index, column_index).value
            if value is not None and value != "":
                return row_index
    return 0


def append_rows_openpyxl(path: Path, sheet_name: str,
                         rows: Sequence[Sequence[Any]],
                         extend_header: Optional[Sequence[Any]] = None,
                         header: Optional[Sequence[Any]] = None,
                         columns: Optional[Sequence[Any]] = None
                         ) -> Tuple[bool, str]:
    """用 openpyxl 向目标工作表追加明细行，返回 (是否成功, 说明)。

    extend_header：目标表需要新增列时传入完整表头，用来补写第 1 行空位。
    header：目标表现在完全空白时传入，先写表头再写数据（第 2 行起）。
    columns：本次写入行对应的列名，用来定位必须存文本的「行头路径」等列。"""
    path = Path(path)
    if not path.is_file():
        return False, "目标文件不存在"

    try:
        from openpyxl import load_workbook
    except ImportError:
        return False, "openpyxl 不可用"

    try:
        book = load_workbook(path)          # data_only=False：必须保留公式
    except Exception as e:
        return False, f"openpyxl 打开失败：{e}"
    try:
        if sheet_name not in book.sheetnames:
            return False, f"目标工作表不存在：{sheet_name}"
        sheet = book[sheet_name]
        last = _last_content_row(sheet)
        if header:
            # 空表：表头写第 1 行，数据从第 2 行开始
            for column_index, value in enumerate(header, 1):
                if value not in (None, ""):
                    sheet.cell(1, column_index, value)
            start = 2
        else:
            start = last + 1
        if rows:
            for offset, values in enumerate(rows):
                for col, value in enumerate(values, 1):
                    if value is None or value == "":
                        continue
                    sheet.cell(start + offset, col, value)
        # 「行头路径」「科目代码」这类列几乎都是指标/科目代码，必须保持文本：
        # 写成数值后 Excel 会显示成 11344，再读回来与 '11344' 对不上，去重和
        # 跨月比对都会失效。按表头名定位这些列，把本次写入的值和时间统一改文本，
        # 并挂 "@" 格式，防止后续手工输入被 Excel 自动转成数值。
        if rows:
            _force_text_columns(sheet, columns or header or extend_header or (),
                                list(range(start, start + len(rows))))
        if extend_header:
            header_row = 1
            for col, value in enumerate(extend_header, 1):
                if value in (None, ""):
                    continue
                cell = sheet.cell(header_row, col)
                if cell.value in (None, ""):
                    cell.value = value
        book.save(path)
    except PermissionError:
        return False, "目标工作簿被占用，请先关闭 Excel 后再重试"
    except Exception as e:
        return False, f"openpyxl 写入失败：{e}"
    finally:
        try:
            book.close()
        except Exception:
            pass

    prefix = "空表重建表头并写入" if header else "追加"
    return True, f"openpyxl {prefix} {len(rows)} 行（第 {start} 行起）"

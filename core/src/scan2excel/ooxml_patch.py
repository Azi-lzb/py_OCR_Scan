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
import zipfile
from pathlib import Path
from typing import Any, List, Optional, Sequence, Tuple

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def append_rows_openpyxl(path: Path, sheet_name: str,
                         rows: Sequence[Sequence[Any]],
                         extend_header: Optional[Sequence[Any]] = None
                         ) -> Tuple[bool, str]:
    """用 openpyxl 向目标工作表追加明细行，返回 (是否成功, 说明)。

    extend_header：目标表需要新增列时传入完整表头，用来补写第 1 行空位。"""
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
        start = sheet.max_row + 1
        if rows:
            for offset, values in enumerate(rows):
                for col, value in enumerate(values, 1):
                    if value is None or value == "":
                        continue
                    sheet.cell(start + offset, col, value)
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

    return True, f"openpyxl 追加 {len(rows)} 行（第 {start} 行起）"

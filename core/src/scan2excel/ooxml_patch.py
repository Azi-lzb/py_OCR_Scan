# -*- coding: utf-8 -*-
"""OOXML 就地补丁：直接改既有 .xlsx 的目标工作表，其余部件逐字节保留。

为什么需要它：XlsxWriter / OpenPyXL 都是"整本重建"，即使把值、公式、合并都
读出来再写回，重建仍会丢掉读不到的东西——实测目标簿会丢批注
（xl/comments1.xml + vmlDrawing）、自定义文档属性、主题，样式表从 20KB 缩到
1KB。用户自己的报表被这样洗一遍是不可接受的。

所以写用户既有工作簿时走这里：只把新行插进目标 sheet 的 <sheetData>，zip 里
其他条目原样带过（连压缩内容和顺序都不改），批注/样式/主题/公式全部不受影响。

字符串一律写 inlineStr（<c t="inlineStr"><is><t>…</t></is></c>），这样完全不
用碰 sharedStrings.xml，也就不必重排既有共享字符串索引。
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# 兼容默认命名空间与带前缀（x:sheetData）两种写法
_SHEETDATA_RE = re.compile(r"(<([A-Za-z0-9]*:)?sheetData\b[^>]*>)(.*?)(</\2?sheetData>)",
                           re.S)
_DIMENSION_RE = re.compile(r"<((?:[A-Za-z0-9]*:)?dimension)\b[^>]*?/>")
_ROW_INDEX_RE = re.compile(r"<(?:(?:[A-Za-z0-9]*:)?row)\b[^>]*?\br=\"(\d+)\"")


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _column_name(column: int) -> str:
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _column_index(letters: str) -> int:
    index = 0
    for ch in letters.upper():
        index = index * 26 + (ord(ch) - 64)
    return index


def _cell_xml(ref: str, value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (value != value or value in (
                float("inf"), float("-inf"))):
            return ""          # NaN/Inf 无法写成数值，留空
        if isinstance(value, float) and value.is_integer():
            return f'<c r="{ref}"><v>{int(value)}</v></c>'
        return f'<c r="{ref}"><v>{value}</v></c>'
    text = _xml_escape(str(value))
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _row_xml(row_index: int, values: Sequence[Any]) -> str:
    cells = "".join(_cell_xml(f"{_column_name(c)}{row_index}", v)
                    for c, v in enumerate(values, 1))
    return f'<row r="{row_index}">{cells}</row>'


def _rels_path_for(part: str) -> str:
    directory, name = posixpath.split(part)
    return posixpath.join(directory, "_rels", name + ".rels")


def resolve_sheet_part(archive: zipfile.ZipFile, sheet_name: str) -> Optional[str]:
    """工作表名 → 包内 sheet XML 路径（如 xl/worksheets/sheet2.xml）。"""
    try:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        rels_xml = archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    except KeyError:
        return None
    rels = {item.attrib["Id"]: item.attrib["Target"]
            for item in _parse_rels(rels_xml)}
    for item in _parse_sheets(workbook_xml):
        if item.get("name") != sheet_name:
            continue
        rid = item.get("{%s}id" % _REL_NS) or item.get("id")
        target = rels.get(rid or "", "")
        if not target:
            return None
        target = target.lstrip("/")
        return target if target.startswith("xl/") else posixpath.join("xl", target)
    return None


def _parse_rels(xml: str):
    import xml.etree.ElementTree as ET
    root = ET.fromstring(xml)
    return root.findall(f"{{{_PKG_REL_NS}}}Relationship")


def _parse_sheets(xml: str):
    import xml.etree.ElementTree as ET
    root = ET.fromstring(xml)
    return root.findall(f"{{{_MAIN_NS}}}sheets/{{{_MAIN_NS}}}sheet")


def _formula_cache_from_parts(infos: Sequence[Tuple[zipfile.ZipInfo, bytes]],
                              sheet_parts: Sequence[str]) -> Dict[str, Dict[str, str]]:
    """从原始包里摘出 {sheet 部件: {"E2": "缓存值"}}，供 openpyxl 写完补回。"""
    cache: Dict[str, Dict[str, str]] = {}
    for part in sheet_parts:
        data = next((d for i, d in infos if i.filename == part), None)
        if data is None:
            continue
        xml = data.decode("utf-8")
        cells: Dict[str, str] = {}
        # 必须显式区分自闭合 <c .../> 与成对 <c ...>...</c>：否则空单元格
        # <c r="D2" s="11"/> 会把后面 E2 的内容一起吞进来，E 列的公式缓存
        # 就抓不到（实测 191 个只抓到 182 个）。
        for match in re.finditer(
                r'<c\b[^>]*?\br="([A-Z]+\d+)"[^>]*?(?:/>|>(.*?)</c>)', xml, re.S):
            ref, body = match.group(1), match.group(2)
            if not body or "<f" not in body:
                continue
            value = re.search(r"<v>(.*?)</v>", body, re.S)
            if value:
                cells[ref] = value.group(1)
        if cells:
            cache[part] = cells
    return cache


def _restore_formula_cache(path: Path,
                           cache: Dict[str, Dict[str, str]]) -> int:
    """把 openpyxl 保存时清空的公式缓存值写回，返回补回的单元格数。"""
    if not cache:
        return 0
    try:
        with zipfile.ZipFile(path) as archive:
            infos = [(info, archive.read(info.filename))
                     for info in archive.infolist()]
    except (zipfile.BadZipFile, OSError):
        return 0
    restored = 0
    patched: List[Tuple[zipfile.ZipInfo, bytes]] = []
    for info, data in infos:
        refs = cache.get(info.filename)
        if not refs:
            patched.append((info, data))
            continue
        xml = data.decode("utf-8")
        for ref, cached in refs.items():
            if not cached:
                continue
            # 只用正则限定在目标单元格内替换空的 <v></v>
            cell_re = re.compile(
                r'(<c\b[^>]*?\br="%s"[^>]*>(?:(?!</c>).)*?<v>)</v>'
                % re.escape(ref), re.S)
            new_xml, n = cell_re.subn(r"\g<1>%s</v>" % cached, xml, count=1)
            if n:
                xml = new_xml
                restored += 1
        patched.append((info, xml.encode("utf-8")))
    if not restored:
        return 0
    temp = path.with_name(f".{path.name}.cache.tmp")
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
            for info, data in patched:
                archive.writestr(info, data)
        temp.replace(path)
    except Exception:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass
        return 0
    return restored


def append_rows_openpyxl(path: Path, sheet_name: str,
                         rows: Sequence[Sequence[Any]],
                         extend_header: Optional[Sequence[Any]] = None
                         ) -> Tuple[bool, str]:
    """用 openpyxl 就地追加行，返回 (是否成功, 说明)。

    相比整本重建（XlsxWriter），openpyxl 会保留公式、批注、样式、主题等所有
    既有内容，只重写自己认识的部分；代价是保存时把公式缓存值清空，因此这里
    先抓一份原缓存，写完补回去（否则 calamine 之类按缓存读取会读到空值）。

    extend_header：目标表需要新增列时传入完整表头，用来补写第 1 行空位。"""
    path = Path(path)
    if not path.is_file():
        return False, "目标文件不存在"
    try:
        with zipfile.ZipFile(path) as archive:
            parts = [resolve_sheet_part(archive, sheet_name) or ""]
            parts += [resolve_sheet_part(archive, n) or ""
                      for n in _sheet_names(archive)]
            infos = [(info, archive.read(info.filename))
                     for info in archive.infolist()]
    except (zipfile.BadZipFile, KeyError, OSError) as e:
        return False, f"无法读取目标工作簿：{e}"
    cache = _formula_cache_from_parts(infos, [p for p in parts if p])

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

    restored = _restore_formula_cache(path, cache)
    note = f"openpyxl 追加 {len(rows)} 行（第 {start} 行起）"
    if cache:
        note += f"，补回公式缓存 {restored} 格"
    return True, note


def _sheet_names(archive: zipfile.ZipFile) -> List[str]:
    try:
        xml = archive.read("xl/workbook.xml").decode("utf-8")
    except KeyError:
        return []
    return [item.get("name", "") for item in _parse_sheets(xml) if item.get("name")]


def append_rows_in_place(path: Path, sheet_name: str,
                         rows: Sequence[Sequence[Any]]) -> Tuple[bool, str]:
    """向既有 .xlsx 的目标工作表追加行，返回 (是否成功, 说明)。

    只重写目标 sheet 的 XML；zip 内其他条目按原顺序、原压缩方式写回。"""
    path = Path(path)
    if not path.is_file() or not rows:
        return False, "目标文件不存在或没有要追加的行"
    try:
        with zipfile.ZipFile(path) as archive:
            sheet_part = resolve_sheet_part(archive, sheet_name)
            if not sheet_part or sheet_part not in archive.namelist():
                return False, f"目标工作表不存在：{sheet_name}"
            sheet_xml = archive.read(sheet_part).decode("utf-8")
            infos = [(info, archive.read(info.filename))
                     for info in archive.infolist()]
    except (zipfile.BadZipFile, KeyError, OSError) as e:
        return False, f"无法读取目标工作簿：{e}"

    match = _SHEETDATA_RE.search(sheet_xml)
    if not match:
        return False, "目标工作表没有 sheetData 节点"
    prefix = match.group(2) or ""
    body = match.group(3)
    existing_rows = [int(x) for x in _ROW_INDEX_RE.findall(body)]
    next_row = (max(existing_rows) if existing_rows else 0) + 1

    new_xml = []
    for offset, values in enumerate(rows):
        new_xml.append(_row_xml(next_row + offset, values))
    last_row = next_row + len(rows) - 1

    # 插到 sheetData 末尾（行必须按升序，追加天然满足）
    rebuilt = (sheet_xml[:match.start(3)] + body + "".join(new_xml)
               + sheet_xml[match.end(3):])

    # 更新 dimension 的行范围，列范围保持不变
    def _fix_dimension(dim: re.Match) -> str:
        tag = dim.group(1)
        tag_text = dim.group(0)
        ref_match = re.search(r'ref="([^"]*)"', tag_text)
        if not ref_match:
            return tag_text
        ref = ref_match.group(1)
        if ":" in ref:
            start, end = ref.split(":", 1)
        else:
            start = end = ref
        col_match = re.match(r"([A-Za-z]*)(\d*)", end) or None
        end_col = col_match.group(1) if col_match else "A"
        try:
            end_row = int(col_match.group(2)) if col_match and col_match.group(2) else 1
        except ValueError:
            end_row = 1
        new_ref = f"{start}:{end_col}{max(end_row, last_row)}"
        return tag_text.replace(f'ref="{ref}"', f'ref="{new_ref}"', 1)

    rebuilt = _DIMENSION_RE.sub(_fix_dimension, rebuilt, count=1)

    temp = path.with_name(f".{path.name}.append.tmp")
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
            for info, data in infos:
                if info.filename == sheet_part:
                    data = rebuilt.encode("utf-8")
                archive.writestr(info, data)
        temp.replace(path)
    except Exception as e:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass
        return False, f"写入失败：{e}"
    return True, f"就地追加 {len(rows)} 行（第 {next_row}~{last_row} 行）"

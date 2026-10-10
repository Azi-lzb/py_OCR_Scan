# -*- coding: utf-8 -*-
"""「数据处理」引擎：把 OCR 导出的多 sheet 工作簿摊平为长表。

四类能力：
  load_rules_v2     读 config.xlsx「时序提取规则」（pytools 兼容列名，列支持字母）
  wide_summary      宽表规则汇总（1-4 同款）：行=行头，列=列头路径，直接可 SUMIFS
  dedup_append      去重追加到目标（2-5 同款）：源各 sheet 追加到目标工作簿同名 sheet
  pages_from_export 读导出文件的各 sheet 为页面数据

自家导出格式的内置默认规则：两行表头（组名/借贷方）、A/B 为行头、数据自第 3 行起。
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

DEFAULT_RULE = {
    "enabled": True,
    "name": "OCR导出表",
    "wb_keyword": "",
    "sheet_keyword": "",
    "header_rows": "1,2",
    "row_header_cols": "1,2",
    "data_row_start": 3,
    "data_col_start": 3,
    "data_col_end": 8,          # 本期余额贷（H）；空=到最右
    "skip_keywords": "",
    "skip_nonnumeric": "是",    # 是=只取数值格（滤掉右缘代码列等文本噪声）
    "set_items": "",            # 别名=地址，分号分隔，如 数据日期=C4
}

_LONG_HEADERS = ["源文件", "工作表", "规则", "固定项", "行头", "列头", "数值"]


def _csv_ints(s: str) -> List[int]:
    out = []
    for part in str(s or "").split(","):
        part = part.strip()
        if part.isdigit():
            out.append(int(part))
    return out


def _norm(s) -> str:
    return str(s or "").strip()


def load_rules(config_xlsx: Optional[Path]) -> List[Dict[str, Any]]:
    """读 config.xlsx「时序规则」sheet；没有则返回内置默认规则。"""
    rules: List[Dict[str, Any]] = []
    if config_xlsx and Path(config_xlsx).is_file():
        try:
            wb = load_workbook(Path(config_xlsx), data_only=True)
            if "时序规则" in wb.sheetnames:
                ws = wb["时序规则"]
                headers = [_norm(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
                for r in range(2, ws.max_row + 1):
                    row = {h: _norm(ws.cell(r, c).value)
                           for c, h in enumerate(headers, start=1) if h}
                    if not any(row.values()):
                        continue
                    rule = dict(DEFAULT_RULE)
                    rule.update({k: v for k, v in row.items() if v})
                    rule["enabled"] = str(row.get("启用", "是")).strip() not in ("否", "0", "false")
                    for k in ("data_row_start", "data_col_start"):
                        rule[k] = int(row[k]) if str(row.get(k, "")).strip().isdigit() else DEFAULT_RULE[k]
                    if str(row.get("数据结束列", "")).strip().isdigit():
                        rule["data_col_end"] = int(row["数据结束列"])
                    rules.append(rule)
            wb.close()
        except Exception:
            rules = []
    if not rules:
        rules = [dict(DEFAULT_RULE)]
    return rules


def _fill_merged_forward(ws, row: int, c_start: int, c_end: int) -> List[str]:
    """行内取值：合并单元格取锚点值并向后顺延（组表头跨列）。"""
    vals: List[str] = []
    last = ""
    ranges = [(rg.min_row, rg.min_col, rg.max_col)
              for rg in ws.merged_cells.ranges if rg.min_row <= row <= rg.max_row]
    for c in range(c_start, c_end + 1):
        anchor = next((r0 for r0, c0, _ in ranges if c0 <= c <= _), None)
        v = _norm(ws.cell(anchor, c).value if anchor else ws.cell(row, c).value)
        if v and c > c_start and anchor == next(
                (r0 for r0, c0, _ in ranges if c0 <= c - 1 <= _), None):
            last = v                      # 同一合并区内顺延
        elif v:
            last = v
        vals.append(last)
    return vals


def timeline_extract(source_path: Path, rules: List[Dict[str, Any]],
                     out_path: Path) -> Dict[str, Any]:
    """按规则把源文件所有匹配 sheet 摊平为长表并写出。"""
    wb = load_workbook(Path(source_path), data_only=True)
    src_name = Path(source_path).stem
    out_rows: List[List[Any]] = []
    rule_stat: Dict[str, int] = {}
    for ws in wb.worksheets:
        for rule in rules:
            if not rule.get("enabled", True):
                continue
            wb_kw = _norm(rule.get("wb_keyword"))
            sh_kw = _norm(rule.get("sheet_keyword"))
            if wb_kw and wb_kw not in Path(source_path).name:
                continue
            if sh_kw and sh_kw not in ws.title:
                continue
            header_rows = _csv_ints(rule.get("header_rows", "1,2"))
            row_cols = _csv_ints(rule.get("row_header_cols", "1,2"))
            d_start = int(rule.get("data_row_start", 3) or 3)
            d_col = int(rule.get("data_col_start", 3) or 3)
            d_end = int(rule["data_col_end"]) if str(rule.get("data_col_end", "")).strip().isdigit() else ws.max_column
            skip_kw = [k.strip() for k in _norm(rule.get("skip_keywords")).split(";") if k.strip()]
            if not header_rows or not row_cols:
                continue
            col_parts = [_fill_merged_forward(ws, h, 1, ws.max_column)
                         for h in header_rows]
            col_paths = []
            for c in range(1, ws.max_column + 1):
                parts = []
                for pr in col_parts:
                    pv = pr[c - 1] if c - 1 < len(pr) else ""
                    if pv and (not parts or pv != parts[-1]):
                        parts.append(pv)
                col_paths.append("_".join(parts) or f"列{c}")
            n = 0
            for r in range(max(d_start, max(h for h in header_rows) + 1), ws.max_row + 1):
                row_head = " | ".join(_norm(ws.cell(r, c).value)
                                      for c in row_cols if c <= ws.max_column)
                cells = []
                for c in range(d_col, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    cells.append(v)
                if all(v in (None, "") for v in cells):
                    continue
                if any(k in row_head for k in skip_kw):
                    continue
                only_num = str(rule.get("skip_nonnumeric", "")).strip() in ("是", "1", "true")
                for c, v in enumerate(cells, start=d_col):
                    if c > d_end:
                        break
                    if v in (None, ""):
                        continue
                    if only_num:
                        s = str(v).strip().replace(",", "")
                        try:
                            float(s)
                        except ValueError:
                            continue
                    cp = col_paths[c - 1] if c - 1 < len(col_paths) else f"列{c}"
                    out_rows.append([src_name, ws.title, rule.get("name", ""),
                                     "", row_head, cp,
                                     float(v) if isinstance(v, (int, float)) else _norm(v)])
                    n += 1
            if n:
                rule_stat[rule.get("name", "")] = rule_stat.get(rule.get("name", ""), 0) + n
    wb.close()

    out = Workbook()
    sheet = out.active
    sheet.title = "汇总长表"
    sheet.append(_LONG_HEADERS)
    for row in out_rows:
        sheet.append(row)
    for c, wd in zip("ABCDEFG", (18, 26, 12, 12, 30, 22, 18)):
        sheet.column_dimensions[c].width = wd
    sheet.freeze_panes = "A2"
    out.save(Path(out_path))
    return {"rows": len(out_rows), "rule_stat": rule_stat, "path": str(out_path)}


def pages_from_export(source_path: Path) -> List[Dict[str, Any]]:
    """读导出文件的每个 sheet 为页面数据（名称=工作表名，行=值矩阵）。"""
    wb = load_workbook(Path(source_path), data_only=True)
    out: List[Dict[str, Any]] = []
    for ws in wb.worksheets:
        rows: List[List] = []
        for r in range(1, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if any(v not in (None, "") for v in row):
                rows.append(row)
        if rows:
            out.append({"name": ws.title, "title": ws.title, "rows": rows})
    wb.close()
    return out


def write_default_config(path: Path) -> None:
    """生成默认规则 config.xlsx（用户可在此基础上加规则）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "时序规则"
    headers = ["启用", "规则名称", "工作簿关键词", "工作表关键词", "表头行",
               "行头列", "数据起始行", "数据结束列", "数据起始列", "跳过关键词", "仅数值"]
    ws.append(headers)
    ws.append(["是", DEFAULT_RULE["name"], "", "", DEFAULT_RULE["header_rows"],
               DEFAULT_RULE["row_header_cols"], DEFAULT_RULE["data_row_start"],
               DEFAULT_RULE["data_col_end"], DEFAULT_RULE["data_col_start"],
               "", DEFAULT_RULE["skip_nonnumeric"]])
    ws.append(["说明", "一行一条规则；工作簿/工作表关键词留空=匹配全部；"
               "表头行=列组名所在行；行头列=行标识所在列；多列/多行用英文逗号"])
    for c, wd in zip("ABCDEFGHIJ", (6, 16, 16, 16, 8, 8, 10, 10, 16, 8)):
        ws.column_dimensions[c].width = wd
    wb.save(Path(path))
    wb.close()


# ------------------------------------------------------------------
# pytools 兼容：规则加载（「时序提取规则」sheet）+ 宽表汇总 + 去重追加
# ------------------------------------------------------------------

def _col_letter_to_index(s: str) -> int:
    """列字母转 1 基序号；数字直接用；空/非法返回 0。"""
    s = str(s or "").strip().upper()
    if not s:
        return 0
    if s.isdigit():
        return int(s)
    n = 0
    for ch in s:
        if not ("A" <= ch <= "Z"):
            return 0
        n = n * 26 + (ord(ch) - 64)
    return n


def load_rules_v2(config_xlsx: Optional[Path]) -> List[Dict[str, Any]]:
    """读 config.xlsx「时序提取规则」sheet（pytools 兼容列名，列支持字母）。"""
    rules: List[Dict[str, Any]] = []
    cfg = Path(config_xlsx) if config_xlsx else None
    if cfg and cfg.is_file():
        try:
            wb = load_workbook(cfg, data_only=True)
            if "时序提取规则" in wb.sheetnames:
                ws = wb["时序提取规则"]
                headers = [_norm(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]

                def cellv(r: int, name: str) -> str:
                    if name in headers:
                        return _norm(ws.cell(r, headers.index(name) + 1).value)
                    return ""

                for r in range(2, ws.max_row + 1):
                    if cellv(r, "是否启用") in ("", "否", "0", "false"):
                        continue
                    name = cellv(r, "规则名称") or "规则%d" % (r - 1)
                    rules.append({
                        "enabled": True,
                        "name": name,
                        "wb_keyword": cellv(r, "工作簿关键字"),
                        "sheet_keyword": cellv(r, "工作表关键字"),
                        "row_header_col": _col_letter_to_index(cellv(r, "行头列")) or 1,
                        "header_rows": _csv_ints(cellv(r, "列表头行")) or [1],
                        "data_row_start": int(cellv(r, "数据起始行") or 0),
                        "data_row_end": int(cellv(r, "数据结束行") or 0),
                        "data_col_end": _col_letter_to_index(cellv(r, "数据结束列")),
                        "data_col_start": _col_letter_to_index(cellv(r, "数据起始列")),
                        "skip_keywords": cellv(r, "跳过关键字"),
                        "skip_nonnumeric": "否",
                        "target_wb": cellv(r, "目标工作簿路径"),
                        "target_sheet": cellv(r, "目标工作表"),
                        # 启用目标写入：留空默认启用（pytools 同款）
                        "target_write": _norm(cellv(r, "启用目标写入")).lower()
                        not in ("n", "否", "0", "false", "no"),
                        "target_dedup": cellv(r, "目标去重列"),
                    })
            wb.close()
        except Exception:
            rules = []
    if not rules:
        # 内置默认规则（typed）：两行表头、A/B 行头、数据 C3:H8
        rules = [{
            "enabled": True,
            "name": DEFAULT_RULE["name"],
            "wb_keyword": "",
            "sheet_keyword": "",
            "row_header_col": _col_letter_to_index("A") or 1,
            "header_rows": _csv_ints("1,2"),
            "data_row_start": 3,
            "data_col_start": 3,
            "data_col_end": 8,
            "skip_keywords": "",
            "skip_nonnumeric": DEFAULT_RULE.get("skip_nonnumeric", "否"),
        }]
    return rules


def _match_all_kw(text: str, kw: str) -> bool:
    """工作簿/工作表关键字（pytools 语义）：;/，/, 分割，全部包含才算匹配；空关键字恒匹配。"""
    if not kw:
        return True
    return all(k in text for k in re.split(r"[;，,]", kw) if k)


def _col_paths_for(ws, header_rows: List[int],
                   c_start: int = 1, c_end: int = 0) -> List[str]:
    """列头路径（pytools 同款）：表头格取合并区左上锚点值，非空段用 _ 拼接；
    全空列用 列_字母 兜底；同名路径按出现顺序加 _1/_2 消歧。
    返回 c_start..c_end（含）的路径列表，下标 0 对应 c_start。"""
    ce = c_end or ws.max_column
    anchor: Dict[Tuple[int, int], Any] = {}
    for mg in ws.merged_cells.ranges:
        av = ws.cell(mg.min_row, mg.min_col).value
        for rr in range(mg.min_row, mg.max_row + 1):
            for cc in range(mg.min_col, mg.max_col + 1):
                anchor[(rr, cc)] = av
    raw: Dict[int, str] = {}
    for c in range(c_start, ce + 1):
        parts = []
        for h in header_rows:
            v = ws.cell(h, c).value
            if v in (None, ""):
                v = anchor.get((h, c))
            pv = _norm(v)
            if pv:
                parts.append(pv)
        raw[c] = "_".join(parts)
    dup = {p for p in raw.values() if p and list(raw.values()).count(p) > 1}
    out: List[str] = []
    seen: Dict[str, int] = {}
    for c in range(c_start, ce + 1):
        p = raw[c]
        if not p:
            out.append("列_" + get_column_letter(c))
        elif p in dup:
            seen[p] = seen.get(p, 0) + 1
            out.append("%s_%d" % (p, seen[p]))
        else:
            out.append(p)
    return out


def _resolve_dedup_cols(cfg: str, header: List[str]) -> List[str]:
    """目标去重列解析：列名优先，也支持字母列号（B→第2列）；未配置返回空。"""
    out: List[str] = []
    for c in [x.strip() for x in _norm(cfg).split(";") if x.strip()]:
        if c in header:
            out.append(c)
            continue
        idx = _col_letter_to_index(c)
        if idx and 1 <= idx <= len(header):
            out.append(header[idx - 1])
    return out


def _text_key(v: Any) -> str:
    """去重键文本化（pytools _to_text_key_df 口径）：370.0 与 370 同键；
    日期取日期部分（目标文件读回 datetime、内存是 date，须统一）。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def append_to_target(target_wb: Path, target_sheet: str, header: List[str],
                     rows: List[List[Any]], dedup_cols: List[str],
                     required_prefix: List[str]) -> Dict[str, Any]:
    """去重追加到目标簿/表（pytools append_to_target 同款）。

    批内按键去重（保留首条）→ 与目标已有键比对过滤 → 只写新增行；
    目标簿/表不存在则新建；目标表头缺固定前缀列则保护性跳过。
    """
    stat: Dict[str, Any] = {"input": len(rows), "batch": 0, "added": 0,
                            "written": False, "reason": ""}
    if not rows:
        return stat
    ki = [header.index(c) for c in dedup_cols]
    seen = set()
    batch = []
    for row in rows:
        k = tuple(_text_key(row[i]) for i in ki)
        if k in seen:
            continue
        seen.add(k)
        batch.append(row)
    stat["batch"] = len(batch)

    tw = Path(target_wb)
    if tw.is_file():
        try:
            twb = load_workbook(tw)
        except Exception as e:
            raise RuntimeError(
                f"目标工作簿无法读取（文件损坏或格式不对）：{tw} —— {e}")
        if target_sheet in twb.sheetnames:
            tws = twb[target_sheet]
            old_header = [_norm(c.value) for c in tws[1]]
            if any(c not in old_header for c in required_prefix):
                twb.close()
                stat["reason"] = "header_mismatch"
                return stat
            # 列对齐：旧列在前，新列顺延（pytools allow_extend 同款）
            merged_cols = list(old_header) + [c for c in header if c not in old_header]
            old_rows = []
            for r in tws.iter_rows(min_row=2, values_only=True):
                if all(v in (None, "") for v in r):
                    continue
                old_rows.append(list(r) + [""] * (len(merged_cols) - len(r)))
            ok_i = [merged_cols.index(c) for c in dedup_cols]
            existing = {tuple(_text_key(row[i]) for i in ok_i) for row in old_rows}
            new_rows = []
            for row in batch:
                mrow = [row[header.index(c)] if c in header else "" for c in merged_cols]
                k = tuple(_text_key(mrow[i]) for i in ok_i)
                if k in existing:
                    continue
                existing.add(k)
                new_rows.append(mrow)
            stat["added"] = len(new_rows)
            if not new_rows:
                twb.close()
                return stat
            idx = twb.sheetnames.index(target_sheet)
            twb.remove(tws)
            tws = twb.create_sheet(target_sheet, idx)
            tws.append(merged_cols)
            for row in old_rows + new_rows:
                tws.append(row)
        else:
            stat["added"] = len(batch)
            tws = twb.create_sheet(target_sheet)
            tws.append(header)
            for row in batch:
                tws.append(row)
    else:
        tw.parent.mkdir(parents=True, exist_ok=True)
        twb = Workbook()
        tws = twb.active
        tws.title = target_sheet
        tws.append(header)
        for row in batch:
            tws.append(row)
        stat["added"] = len(batch)
    hdr_vals = [_norm(x.value) for x in tws[1]]
    if "数据日期" in hdr_vals:
        di = hdr_vals.index("数据日期") + 1
        for rr in range(2, tws.max_row + 1):
            cell = tws.cell(rr, di)
            if isinstance(cell.value, (date, datetime)):
                cell.number_format = "yyyy-mm-dd"
    try:
        twb.save(tw)
    except PermissionError:
        raise RuntimeError(f"目标工作簿被占用，请先关闭 Excel 里的 {tw.name} 再重试")
    stat["written"] = True
    return stat


def wide_summary(source_path: Path, rules: List[Dict[str, Any]],
                 out_path: Path) -> Dict[str, Any]:
    """宽表规则汇总（pytools 3.9.8 同款，行头不加后缀）：行=源数据行，列=列头路径。

    行头按合并格锚点解析（第二页中缝表头的从属行也能取到行头，各占一行）；
    同名行头不压并（每个源行独立一行）；数据日期从源文件名解析（YYYYMMDD）。
    """
    src = Path(source_path)
    wb = load_workbook(src, data_only=True)
    wb_name = src.name
    data_date: Any = ""
    m = re.search(r"(20\d{2})[_.-]?(\d{2})[_.-]?(\d{2})", src.stem)
    if m:
        try:
            data_date = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            data_date = ""
    if data_date == "":
        try:
            data_date = datetime.fromtimestamp(src.stat().st_mtime).date()
        except OSError:
            pass
    buckets: Dict[str, Dict[Tuple[str, str, str, int], Dict[str, Any]]] = {}
    rule_cols: Dict[str, Tuple[List[str], set]] = {}
    col_order: List[str] = []        # 宽表 sheet 用：跨规则首见并集
    col_seen: set = set()
    for rule in rules:
        if not rule.get("enabled", True):
            continue
        rname = rule.get("name") or "规则"
        rb = buckets.setdefault(rname, {})
        rorder, rseen = rule_cols.setdefault(rname, ([], set()))
        wb_kw = _norm(rule.get("wb_keyword"))
        sh_kw = _norm(rule.get("sheet_keyword"))
        header_rows = rule.get("header_rows") or [1]
        rh_col = int(rule.get("row_header_col", 1) or 1)
        d_start = int(rule.get("data_row_start", 0) or 0)
        d_end = int(rule.get("data_col_end", 0) or 0)
        d_start_c = int(rule.get("data_col_start", 0) or 0)
        skip_kw = [k.strip() for k in _norm(rule.get("skip_keywords")).split(";") if k.strip()]
        for ws in wb.worksheets:
            if not _match_all_kw(src.name, wb_kw):
                continue
            if not _match_all_kw(ws.title, sh_kw):
                continue
            cs = max(d_start_c, 1)
            ce = d_end or ws.max_column
            col_paths = _col_paths_for(ws, header_rows, cs, ce)
            for cp in col_paths:      # 列序 = 数据列范围左到右（全部列，非首见）
                if cp not in rseen:
                    rseen.add(cp)
                    rorder.append(cp)
                if cp not in col_seen:
                    col_seen.add(cp)
                    col_order.append(cp)
            # 合并格 → 锚点值映射（行头解析：中缝表头等合并从属格取锚点文本）
            anchor: Dict[Tuple[int, int], Any] = {}
            for mg in ws.merged_cells.ranges:
                av = ws.cell(mg.min_row, mg.min_col).value
                for rr in range(mg.min_row, mg.max_row + 1):
                    for cc in range(mg.min_col, mg.max_col + 1):
                        if (rr, cc) != (mg.min_row, mg.min_col):
                            anchor[(rr, cc)] = av
            d_row_end = int(rule.get("data_row_end", 0) or 0)
            row_max = min(ws.max_row, d_row_end) if d_row_end else ws.max_row
            for r in range(max(d_start, 1), row_max + 1):
                hv = ws.cell(r, rh_col).value
                if hv in (None, ""):
                    hv = anchor.get((r, rh_col))
                row_head = _norm(hv)
                if not row_head or any(k in row_head for k in skip_kw):
                    continue
                vals = {}
                for c in range(cs, ce + 1):
                    v = ws.cell(r, c).value
                    if v in (None, ""):
                        continue
                    cp = col_paths[c - cs] if 0 <= c - cs < len(col_paths) else ("列_" + get_column_letter(c))
                    vals[cp] = v
                if not vals:
                    continue
                # key 带源行号：同名行头（如中缝表头两行）各占一行，不互相覆盖
                key = (wb_name, ws.title, row_head, r)
                bucket = rb.setdefault(key, {})
                for cp, v in vals.items():
                    bucket.setdefault(cp, v)
    wb.close()

    out = Workbook()
    sheet = out.active
    sheet.title = "宽表"
    headers = ["工作簿名", "工作表名", "数据日期", "行头路径"] + col_order
    sheet.append(headers)
    # 按源顺序输出：源工作簿的 sheet 顺序 × 源行号顺序（pytools 同款，不排序）
    total = 0
    for rname, rb in buckets.items():
        for (wbk, shk, rhk, _rn), vals in rb.items():
            sheet.append([wbk, shk, data_date, rhk] + [vals.get(cp, "") for cp in col_order])
            total += 1
    for c, wd in zip("ABCD", (22, 26, 12, 20)):
        sheet.column_dimensions[c].width = wd
    for c in range(5, len(headers) + 1):
        sheet.column_dimensions[get_column_letter(c)].width = 16
    sheet.freeze_panes = "E2"
    out.save(Path(out_path))

    # 写目标簿（pytools 同款）：每条启用目标写入的规则，把自己的行去重追加到目标表
    targets: List[Dict[str, Any]] = []
    for rule in rules:
        if not rule.get("enabled", True):
            continue
        rname = rule.get("name") or "规则"
        rb = buckets.get(rname) or {}
        if not rb:
            continue
        tw = _norm(rule.get("target_wb"))
        ts = _norm(rule.get("target_sheet"))
        if not rule.get("target_write", True) or not (tw and ts):
            continue
        rorder, _ = rule_cols.get(rname, ([], set()))
        header = ["工作簿名", "工作表名", "数据日期", "行头路径"] + rorder
        rows = [[wbk, shk, data_date, rhk] + [vals.get(cp, "") for cp in rorder]
                for (wbk, shk, rhk, _rn), vals in rb.items()]
        dedup_cols = _resolve_dedup_cols(rule.get("target_dedup"), header) \
            or ["数据日期", "行头路径"]
        stat = append_to_target(Path(tw), ts, header, rows, dedup_cols,
                                ["工作簿名", "工作表名", "数据日期", "行头路径"])
        stat.update({"rule": rname, "wb": tw, "sheet": ts})
        targets.append(stat)
    return {"rows": total, "cols": len(col_order), "targets": targets,
            "path": str(out_path), "source": str(source_path), "date": data_date}



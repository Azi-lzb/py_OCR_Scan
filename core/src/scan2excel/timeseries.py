# -*- coding: utf-8 -*-
"""时序归集校验（移植 ref-VBA_A1402_CHECK 的 PBCManual 国库链路）。

输入：多份 `YYYY.MM 地区名.xls` 指标表（Sheet0，5 列：指标代码/指标名称/
行序号/余额/上期－余额）+ 一本时序工作簿（每地区一 sheet，行=指标，列=月份）。
行为：归集进时序 → 总分校验（市级 vs Σ县区）→ 不应有数校验 →
时序表内写「校验结果」（13 列+颜色）与「log」sheet。
全部规则读 config.xlsx：归集配置/地区规则/白名单 三 sheet（键见 _CFG 常量）。
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill

# ---------------------------------------------------------------- 配色（VBA 同款 RGB）
_FILL = {
    "backfill_src": ("FFFFCC", "000000"),   # 会计补录来源：浅黄
    "fee_adjust":   ("E2EFDA", "006100"),   # 费用调整：淡绿
    "whitelist":    ("92D050", "000000"),   # 白名单：绿
    "error":        ("FFC7CE", "9C0006"),   # 不平/不符/异常/应为0：红
    "adjusted":     ("F2FFF2", "008000"),   # 已补录/已调整：浅绿
}
_RESULT_HEADERS = ["月份", "区域/来源", "指标代码", "指标名称", "异常说明",
                   "当前值1", "调整前值2", "地区合计值3", "差异(1-2)", "差异(1-3)",
                   "业务类型", "源文件名称", "处理时间"]
_EXCLUDE_SHEETS = {"校验结果", "log", "校验页", "config"}


def _num(v: Any) -> float:
    """VBA Val 口径：空/None/错误 → 0；字符串取前导数字。"""
    if v is None or isinstance(v, str) and not v.strip():
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        m = re.match(r"[-+]?\d*\.?\d+", str(v).strip())
        return float(m.group(0)) if m else 0.0


# ---------------------------------------------------------------- 配置读取
def load_config(config_xlsx: Path) -> Dict[str, Any]:
    """读 config.xlsx 的 归集配置/地区规则/白名单 三 sheet，缺省值兜底。"""
    cfg: Dict[str, Any] = {
        "value_col": "余额",            # 源表取数列（按表头名匹配）
        "tolerance": 0.01,
        "month_fmt": "YYYY.MM",
        "base_region": "惠州市",
        "backfill_parents": "11364;11001;11628;11371",
        "regions": [                     # 地区名, 文件名关键字, 角色(base/member), 不应有数
            {"name": "惠州市", "kw": "惠州市", "role": "base", "no_data": False},
            {"name": "市本部", "kw": "市本部", "role": "member", "no_data": False},
            {"name": "惠东县", "kw": "惠东", "role": "member", "no_data": False},
            {"name": "惠阳区", "kw": "惠阳", "role": "member", "no_data": True},
            {"name": "博罗县", "kw": "博罗", "role": "member", "no_data": True},
            {"name": "龙门县", "kw": "龙门", "role": "member", "no_data": True},
        ],
        "whitelist": {"国库": set(), "会计": set()},
    }
    if not Path(config_xlsx).is_file():
        return cfg
    wb = load_workbook(config_xlsx, data_only=True)
    try:
        if "归集配置" in wb.sheetnames:
            ws = wb["归集配置"]
            kvmap = {"取数列": "value_col", "容差": "tolerance", "市级基准": "base_region",
                     "月份格式": "month_fmt", "补录父项": "backfill_parents"}
            for r in range(2, ws.max_row + 1):
                k = str(ws.cell(r, 1).value or "").strip()
                v = ws.cell(r, 2).value
                if k in kvmap and v not in (None, ""):
                    cfg[kvmap[k]] = v
            try:
                cfg["tolerance"] = float(cfg["tolerance"])
            except (TypeError, ValueError):
                cfg["tolerance"] = 0.01
        if "地区规则" in wb.sheetnames:
            ws = wb["地区规则"]
            regions = []
            for r in range(2, ws.max_row + 1):
                name = str(ws.cell(r, 1).value or "").strip()
                if not name:
                    continue
                regions.append({
                    "name": name,
                    "kw": str(ws.cell(r, 2).value or "").strip() or name,
                    "role": str(ws.cell(r, 3).value or "成员").strip(),
                    "no_data": str(ws.cell(r, 4).value or "").strip() in ("是", "1", "y", "Y"),
                })
            if regions:
                cfg["regions"] = regions
        if "白名单" in wb.sheetnames:
            ws = wb["白名单"]
            for r in range(2, ws.max_row + 1):
                for col, key in ((2, "国库"), (5, "会计")):   # B 列国库 / E 列会计（VBA 布局）
                    v = ws.cell(r, col).value
                    v = str(v).strip() if v is not None else ""
                    if v:
                        cfg["whitelist"][key].add(v)
    finally:
        wb.close()
    return cfg


# ---------------------------------------------------------------- 文件名解析
def normalize_period(raw: str) -> str:
    """VBA NormalizeDate：去掉 -./，取 YYYY+MM 拼 'YYYY.MM'。"""
    clean = re.sub(r"[-./]", "", str(raw)).strip()
    if len(clean) >= 6:
        return clean[:4] + "." + clean[4:6].zfill(2)
    return str(raw)


def parse_file_name(name: str, regions: List[Dict[str, Any]]) -> Dict[str, str]:
    """`2026.09 惠州市.xls` → {period, region}（VBA ParseFileInfo 口径）。"""
    stem = name.rsplit(".", 1)[0]
    parts = stem.split(" ")
    if not parts or not parts[0]:
        return {"period": "", "region": "未知区域"}
    period = normalize_period(parts[0])
    rest = "".join(parts[1:]) or stem
    for rg in regions:
        if rg["kw"] in rest:
            return {"period": period, "region": rg["name"]}
    return {"period": period, "region": "未知区域"}


# ---------------------------------------------------------------- 源表读取
def read_source_table(path: Path, value_col_name: str) -> List[Dict[str, Any]]:
    """读源指标表（.xls 用 xlrd / .xlsx 用 openpyxl），返回
    [{code, name, rowidx, value}]；取数列按表头名匹配，找不到退回第 4 列。"""
    rows: List[Dict[str, Any]] = []
    if path.suffix.lower() == ".xls":
        import xlrd
        wb = xlrd.open_workbook(str(path))
        ws = wb.sheet_by_index(0)
        headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
        vcol = headers.index(value_col_name) if value_col_name in headers else 3
        for r in range(1, ws.nrows):
            code = str(ws.cell_value(r, 0)).strip()
            if not code:
                continue
            rows.append({
                "code": code,
                "name": str(ws.cell_value(r, 1)).strip(),
                "rowidx": str(ws.cell_value(r, 2)).strip(),
                "value": _num(ws.cell_value(r, vcol)),
            })
    else:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        headers = [str(h).strip() if h is not None else "" for h in next(it)]
        vcol = headers.index(value_col_name) if value_col_name in headers else 3
        for row in it:
            code = str(row[0]).strip() if row and row[0] is not None else ""
            if not code:
                continue
            def cell(i: int) -> Any:
                return row[i] if i < len(row) else None
            rows.append({"code": code, "name": str(cell(1) or "").strip(),
                         "rowidx": str(cell(2) or "").strip(),
                         "value": _num(cell(vcol))})
        wb.close()
    return rows


# ---------------------------------------------------------------- 归集 + 校验
def run_treasury(timeseries_path: Path, source_paths: List[Path],
                 config_xlsx: Path) -> Dict[str, Any]:
    """国库链路：归集 → 总分/不应有数校验 → 校验结果+log 落时序表。返回统计。"""
    cfg = load_config(config_xlsx)
    tol = cfg["tolerance"]
    base = cfg["base_region"]
    regions = cfg["regions"]
    wl = cfg["whitelist"].get("国库", set())
    ts_path = Path(timeseries_path)
    ts = load_workbook(ts_path) if ts_path.is_file() else Workbook()

    issues: List[List[Any]] = []
    logs: List[List[Any]] = []
    stats = {"files": 0, "regions_new": [], "cells": 0, "updated": 0,
             "total_bad": 0, "no_data_bad": 0, "wl_hit": 0}

    def month_col(ws, period: str) -> int:
        """找月份列（首行精确匹配），没有则在末尾追加。"""
        for c in range(4, ws.max_column + 1):
            if str(ws.cell(1, c).value).strip() == period:
                return c
        c = max(ws.max_column, 3) + 1
        cell = ws.cell(1, c, period)
        cell.font = Font(bold=True)
        return c

    def indicator_row(ws, code: str, name: str, rowidx: str) -> int:
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(r, 1).value).strip() == code:
                return r
        r = ws.max_row + 1 if ws.max_row > 1 else 2
        ws.cell(r, 1, code)
        ws.cell(r, 2, name)
        ws.cell(r, 3, rowidx)
        return r

    def add_issue(period: str, region: str, code: str, name: str, detail: str,
                  v1: Any, v2: Any, v3: Any, src: str) -> None:
        issues.append([period, region, code, name, detail, v1, v2, v3,
                       _num(v1) - _num(v2), _num(v1) - _num(v3),
                       "国库校验", src, datetime.now()])

    # ---- 逐源文件归集 ----
    for src in source_paths:
        src = Path(src)
        if not src.is_file():
            continue
        info = parse_file_name(src.name, regions)
        period, region = info["period"], info["region"]
        if not period or region == "未知区域":
            logs.append([src.name, region or "?", period, "[跳过] 文件名无法解析地区/月份",
                         datetime.now()])
            continue
        ws = ts[region] if region in ts.sheetnames else None
        if ws is None:
            ws = ts.create_sheet(region)
            for c, h in enumerate(("指标代码", "指标名称", "行次号"), 1):
                ws.cell(1, c, h).font = Font(bold=True)
            stats["regions_new"].append(region)
        col = month_col(ws, period)
        rows = read_source_table(src, cfg["value_col"])
        changed = 0
        for row in rows:
            r = indicator_row(ws, row["code"], row["name"], row["rowidx"])
            old = ws.cell(r, col).value
            if _num(old) != row["value"]:
                changed += 1
            ws.cell(r, col, row["value"])
            stats["cells"] += 1
        stats["files"] += 1
        logs.append([src.name, region, period,
                     "[数据更新]" if changed else "[数据导入]", datetime.now()])

    # ---- 总分校验：基准 sheet 每指标×每月份列 vs Σ成员 ----
    if base in ts.sheetnames:
        bws = ts[base]
        members = [rg for rg in regions if rg["name"] != base]
        last_col = max((bws.max_column,), default=4)
        for col in range(4, bws.max_column + 1):
            period = str(bws.cell(1, col).value or "").strip()
            if not period:
                continue
            for r in range(2, bws.max_row + 1):
                code = str(bws.cell(r, 1).value or "").strip()
                if not code:
                    continue
                name = str(bws.cell(r, 2).value or "")
                city = _num(bws.cell(r, col).value)
                total = 0.0
                for rg in members:
                    if rg["name"] in ts.sheetnames:
                        mws = ts[rg["name"]]
                        for rr in range(2, mws.max_row + 1):
                            if str(mws.cell(rr, 1).value or "").strip() == code:
                                total += _num(mws.cell(rr, col).value)
                                break
                if abs(city - total) > tol:
                    stats["total_bad"] += 1
                    if code in wl:
                        stats["wl_hit"] += 1
                    add_issue(period, "总分校验", code, name,
                              "总分不平 (白名单指标)" if code in wl else "总分校验不平",
                              city, "", total, ts_path.name)

    # ---- 不应有数校验 ----
    for rg in regions:
        if not rg["no_data"] or rg["name"] not in ts.sheetnames:
            continue
        ws = ts[rg["name"]]
        for col in range(4, ws.max_column + 1):
            period = str(ws.cell(1, col).value or "").strip()
            for r in range(2, ws.max_row + 1):
                v = _num(ws.cell(r, col).value)
                if abs(v) > tol:
                    stats["no_data_bad"] += 1
                    add_issue(period, rg["name"], str(ws.cell(r, 1).value or "").strip(),
                              str(ws.cell(r, 2).value or ""), "不应有数", v, 0, "", ts_path.name)

    # ---- 校验结果 + log 落表 ----
    _write_result_sheet(ts, issues)
    _write_log_sheet(ts, logs)
    ts.save(ts_path)
    stats["issues"] = len(issues)
    stats["result_path"] = str(ts_path)
    return stats


def _color_for(detail: str, region: str) -> Optional[str]:
    if region in ("会计补录", "会计父项补录"):
        return "backfill_src"
    if "未调整或调整后仍为0" in detail or "调整后仍为0" in detail:
        return "error"
    if any(k in detail for k in ("不平", "不符", "异常", "应为0")):
        return "error"
    if "已补录" in detail or "已调整" in detail:
        return "adjusted"
    if "费用调整" in detail:
        return "fee_adjust"
    if "白名单" in detail:
        return "whitelist"
    return None


def _write_result_sheet(ts: Workbook, issues: List[List[Any]]) -> None:
    if "校验结果" in ts.sheetnames:
        del ts["校验结果"]
    ws = ts.create_sheet("校验结果")
    for c, h in enumerate(_RESULT_HEADERS, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9D9D9")
    for r, row in enumerate(issues, 2):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
        key = _color_for(str(row[4]), str(row[1]))
        if key:
            bg, fg = _FILL[key]
            fill = PatternFill("solid", fgColor=bg)
            font = Font(color=fg)
            for c in range(1, 14):
                ws.cell(r, c).fill = fill
                ws.cell(r, c).font = font


def _write_log_sheet(ts: Workbook, logs: List[List[Any]]) -> None:
    if "log" in ts.sheetnames:
        del ts["log"]
    ws = ts.create_sheet("log")
    for c, h in enumerate(("源文件", "地区", "月份", "动作", "时间"), 1):
        ws.cell(1, c, h).font = Font(bold=True)
    for r, row in enumerate(logs, 2):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)

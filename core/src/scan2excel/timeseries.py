# -*- coding: utf-8 -*-
"""时序归集校验（移植 ref-VBA_A1402_CHECK 的 PBCManual 国库/会计链路）。

输入：多份 `YYYY.MM 地区名.xls` 指标表（Sheet0，5 列：指标代码/指标名称/
行序号/余额/上期－余额）+ 一本时序工作簿（每地区一 sheet，行=指标，列=月份）。
行为：归集进时序 → 总分校验（市级 vs Σ县区）→ 不应有数校验 →
时序表内写「校验结果」（13 列+颜色）与「log」sheet。
全部规则读 config.xlsx：归集配置/地区规则/白名单/补录映射 四 sheet。

性能关键：所有读/算都走内存快照（_snapshot/_merge），只在最后把**变更过的
格子**写回 openpyxl——openpyxl 的 max_row/cell 属性在循环里是 O(全表)，
逐格访问会让 1291 指标 × 12 月的簿跑出几十秒。
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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
    """读 config.xlsx 的 归集配置/地区规则/白名单/补录映射 四 sheet，缺省值兜底。"""
    cfg: Dict[str, Any] = {
        "value_col": "余额",
        "tolerance": 0.01,
        "month_fmt": "YYYY.MM",
        "base_region": "惠州市",
        "backfill_parents": "11364;11001;11628;11371",
        "regions": [
            {"name": "惠州市", "kw": "惠州市", "role": "base", "no_data": False},
            {"name": "市本部", "kw": "市本部", "role": "member", "no_data": False},
            {"name": "惠东县", "kw": "惠东", "role": "member", "no_data": False},
            {"name": "惠阳区", "kw": "惠阳", "role": "member", "no_data": True},
            {"name": "博罗县", "kw": "博罗", "role": "member", "no_data": True},
            {"name": "龙门县", "kw": "龙门", "role": "member", "no_data": True},
        ],
        "backfill_regions": ["惠州市"],
        "backfill_map": [("11367", "512", 7), ("11368", "513", 7), ("11369", "514", 7),
                         ("11370", "515", 7), ("11631", "502", 8)],
        "whitelist": {"国库": set(), "会计": set()},
    }
    if not Path(config_xlsx).is_file():
        return cfg
    wb = load_workbook(config_xlsx, data_only=True)
    try:
        if "归集配置" in wb.sheetnames:
            ws = wb["归集配置"]
            kvmap = {"取数列": "value_col", "容差": "tolerance", "市级基准": "base_region",
                     "月份格式": "month_fmt", "补录父项": "backfill_parents",
                     "补录地区": "backfill_regions"}
            for r in range(2, ws.max_row + 1):
                k = str(ws.cell(r, 1).value or "").strip()
                v = ws.cell(r, 2).value
                if k in kvmap and v not in (None, ""):
                    cfg[kvmap[k]] = v
            try:
                cfg["tolerance"] = float(cfg["tolerance"])
            except (TypeError, ValueError):
                cfg["tolerance"] = 0.01
            if isinstance(cfg["backfill_regions"], str):
                cfg["backfill_regions"] = [x.strip() for x in
                                           cfg["backfill_regions"].split(";") if x.strip()]

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
        if "补录映射" in wb.sheetnames:
            ws = wb["补录映射"]
            bmap = []
            for r in range(2, ws.max_row + 1):
                code = str(ws.cell(r, 1).value or "").strip()
                fee_code = str(ws.cell(r, 2).value or "").strip()
                fee_col = ws.cell(r, 3).value
                if code and fee_code:
                    bmap.append((code, fee_code, int(fee_col) if fee_col else 7))
            if bmap:
                cfg["backfill_map"] = bmap
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
    """`2026.09 惠州市.xls` → {period, region}（关键字最长优先防歧义）。"""
    stem = name.rsplit(".", 1)[0]
    parts = stem.split(" ")
    if not parts or not parts[0]:
        return {"period": "", "region": "未知区域"}
    period = normalize_period(parts[0])
    rest = "".join(parts[1:]) or stem
    for rg in sorted(regions, key=lambda x: -len(x["kw"])):
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
                "src_row": r,
                "prev_raw": ws.cell_value(r, 4) if ws.ncols > 4 else "",
            })
    else:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        headers = [str(h).strip() if h is not None else "" for h in next(it)]
        vcol = headers.index(value_col_name) if value_col_name in headers else 3
        r0 = 1
        for row in it:
            r0 += 1
            code = str(row[0]).strip() if row and row[0] is not None else ""
            if not code:
                continue
            def cell(i: int) -> Any:
                return row[i] if i < len(row) else None
            rows.append({"code": code, "name": str(cell(1) or "").strip(),
                         "rowidx": str(cell(2) or "").strip(),
                         "value": _num(cell(vcol)),
                         "src_row": r0 - 1,
                         "prev_raw": cell(4)})
        wb.close()
    return rows


# ---------------------------------------------------------------- 内存快照模型
def _snapshot(ws) -> Tuple[Dict[int, str], Dict[str, Dict[str, Any]]]:
    """一次性读时序 sheet →
    cols: {列号: 'YYYY.MM'}（数值月份归一）
    indicators: {code: {"row": 行号, "name": 名, "rowidx": 行次, "vals": {列号: 值}}}"""
    cols: Dict[int, str] = {}
    indicators: Dict[str, Dict[str, Any]] = {}
    for r, row in enumerate(ws.iter_rows(values_only=True), 1):
        if r == 1:
            for c, v in enumerate(row, 1):
                if c >= 4 and v not in (None, ""):
                    # 数值月份（如 2026.09 存成 float）必须保留两位月份
                    cols[c] = (normalize_period(f"{v:.2f}")
                               if isinstance(v, float) else normalize_period(str(v).strip()))
            continue
        code = str(row[0]).strip() if row and row[0] is not None else ""
        if not code:
            continue
        vals: Dict[int, float] = {}
        for c, v in enumerate(row, 1):
            if c >= 4 and v not in (None, ""):
                vals[c] = _num(v)
        indicators[code] = {"row": r,
                            "name": str(row[1]) if len(row) > 1 and row[1] is not None else "",
                            "rowidx": str(row[2]) if len(row) > 2 and row[2] is not None else "",
                            "vals": vals}
    return cols, indicators


# ---------------------------------------------------------------- 校验结果 / log 落表
def _color_for(detail: str, region: str) -> Optional[str]:
    """VBA 原始判定顺序：白名单/费用调整优先于"不平"，否则白名单行会误标红。"""
    if region in ("会计补录", "会计父项补录"):
        return "backfill_src"
    if "白名单" in detail:
        return "whitelist"          # 白名单优先：带"- 费用调整"后缀也按绿色
    if "费用调整" in detail:
        return "fee_adjust"
    if "未调整或调整后仍为0" in detail or "调整后仍为0" in detail:
        return "error"
    if any(k in detail for k in ("不平", "不符", "异常", "应为0")):
        return "error"
    if "已补录" in detail or "已调整" in detail:
        return "adjusted"
    return None


def _write_result_sheet(ts: Workbook, issues: List[List[Any]], biz: str) -> None:
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


# ---------------------------------------------------------------- 写回（只写变更格）
def _write_back(ts: Workbook, region: str,
                model: Dict[str, Dict[str, Any]],
                cols: Dict[int, str],
                orig_cols: Dict[int, str],
                orig_vals: Dict[str, Dict[int, Any]],
                new_rows: List[str]) -> None:
    """把内存模型写回 region sheet：新月份列表头、新增指标行、值有变化的格。"""
    ws = ts[region] if region in ts.sheetnames else ts.create_sheet(region)
    for c, period in cols.items():
        if c not in orig_cols:
            cell = ws.cell(1, c, period)
            cell.font = Font(bold=True)
    r_next = ws.max_row + 1
    for code in new_rows:
        ind = model[code]
        ws.cell(r_next, 1, code)
        ws.cell(r_next, 2, ind["name"])
        ws.cell(r_next, 3, ind["rowidx"])
        ind["row"] = r_next
        r_next += 1
    for code, ind in model.items():
        r = ind.get("row")
        if r is None:
            continue
        ovals = orig_vals.get(code, {})
        for c, v in ind["vals"].items():
            if c not in orig_cols or _num(ovals.get(c)) != v:
                ws.cell(r, c, v)


# ---------------------------------------------------------------- 国库链路
def run_treasury(timeseries_path: Path, source_paths: List[Path],
                 config_xlsx: Path) -> Dict[str, Any]:
    """国库链路：归集 → 总分/不应有数校验 → 校验结果+log 落时序表。"""
    cfg = load_config(config_xlsx)
    tol = cfg["tolerance"]
    base = cfg["base_region"]
    regions = cfg["regions"]
    wl = cfg["whitelist"].get("国库", set())
    ts_path = Path(timeseries_path)
    ts = load_workbook(ts_path) if ts_path.is_file() else Workbook()

    issues: List[List[Any]] = []
    logs: List[List[Any]] = []
    stats: Dict[str, Any] = {"files": 0, "regions_new": [], "cells": 0,
                             "total_bad": 0, "no_data_bad": 0, "wl_hit": 0}
    models: Dict[str, Tuple[Dict[int, str], Dict[str, Dict[str, Any]],
                            Dict[int, str], Dict[str, Dict[int, Any]], List[str]]] = {}

    def model_for(region: str):
        if region in models:
            return models[region]
        ws = ts[region] if region in ts.sheetnames else None
        cols, indicators = _snapshot(ws) if ws is not None else ({}, {})
        orig_cols = dict(cols)
        orig_vals = {code: dict(ind["vals"]) for code, ind in indicators.items()}
        new_rows: List[str] = []
        if ws is None:
            stats["regions_new"].append(region)
        models[region] = (cols, indicators, orig_cols, orig_vals, new_rows)
        return models[region]

    for src in map(Path, source_paths):
        if not src.is_file():
            continue
        info = parse_file_name(src.name, regions)
        period, region = info["period"], info["region"]
        if not period or region == "未知区域":
            logs.append([src.name, region or "?", period,
                         "[跳过] 文件名无法解析地区/月份", datetime.now()])
            continue
        cols, indicators, orig_cols, orig_vals, new_rows = model_for(region)
        col = next((c for c, p in cols.items() if p == period), None)
        if col is None:
            col = max(cols, default=3) + 1
            cols[col] = period
        rows = read_source_table(src, cfg["value_col"])
        changed = 0
        for row in rows:
            ind = indicators.get(row["code"])
            if ind is None:
                indicators[row["code"]] = {"row": None, "name": row["name"],
                                           "rowidx": row["rowidx"],
                                           "vals": {col: row["value"]}}
                new_rows.append(row["code"])
                changed += 1
                stats["cells"] += 1
                continue
            old = ind["vals"].get(col)
            if old is None or old != row["value"]:
                changed += 1
            ind["vals"][col] = row["value"]
            stats["cells"] += 1
        stats["files"] += 1
        logs.append([src.name, region, period,
                     "[数据更新]" if changed else "[数据导入]", datetime.now()])

    # ---- 总分校验（纯内存） ----
    if base in models:
        bcols, bmodel, _oc, _ov, _nr = models[base]
        members = [rg for rg in regions if rg["name"] != base]
        member_models = {rg["name"]: models[rg["name"]][1]
                         for rg in members if rg["name"] in models}
        for code, ind in bmodel.items():
            for c, period in bcols.items():
                city = _num(ind["vals"].get(c))
                total = 0.0
                for mmodel in member_models.values():
                    m = mmodel.get(code)
                    if m:
                        total += _num(m["vals"].get(c))
                if abs(city - total) > tol:
                    stats["total_bad"] += 1
                    detail = "总分校验不平"
                    if code in wl:
                        detail = "总分不平 (白名单指标)"
                        stats["wl_hit"] += 1
                    issues.append([period, "总分校验", code, ind["name"], detail,
                                   city, "", total, "", city - total,
                                   "国库校验", ts_path.name, datetime.now()])

    # ---- 不应有数校验（纯内存） ----
    for rg in regions:
        if not rg["no_data"] or rg["name"] not in models:
            continue
        rcols, rmodel, _oc, _ov, _nr = models[rg["name"]]
        for code, ind in rmodel.items():
            for c, period in rcols.items():
                v = _num(ind["vals"].get(c))
                if abs(v) > tol:
                    stats["no_data_bad"] += 1
                    issues.append([period, rg["name"], code, ind["name"], "不应有数",
                                   v, 0, "", v, v, "国库校验", ts_path.name,
                                   datetime.now()])

    for region, (cols, indicators, orig_cols, orig_vals, new_rows) in models.items():
        _write_back(ts, region, indicators, cols, orig_cols, orig_vals, new_rows)
    if "log" in ts.sheetnames:
        del ts["log"]
    _write_result_sheet(ts, issues, "国库校验")
    ts.save(ts_path)
    stats["issues"] = len(issues)
    stats["rows"] = _issue_rows(issues, wl)
    stats["result_path"] = str(ts_path)
    return stats


def _issue_rows(issues: List[List[Any]], wl: set) -> List[Dict[str, Any]]:
    """校验结果行 → 界面明细（含是否白名单）。"""
    out = []
    for it in issues:
        detail = str(it[4] if len(it) > 4 else "")
        out.append({"period": str(it[0]), "region": str(it[1]),
                    "code": str(it[2]), "name": str(it[3]),
                    "detail": detail,
                    "v1": it[5] if len(it) > 5 else "",
                    "v2": it[6] if len(it) > 6 else "",
                    "v3": it[7] if len(it) > 7 else "",
                    "wl": "白名单" in detail or it[2] in wl})
    return out


# ---------------------------------------------------------------- 会计补录（阶段二）
def _fee_period(ws) -> str:
    """余额表 A2 形如 `2026年9月…` → `YYYY-MM`（VBA GetFeePeriod 口径）。"""
    if hasattr(ws, "cell_value"):
        raw = str(ws.cell_value(1, 0) or "").strip()
    else:
        raw = str(ws["A2"] or "").strip()
    m = re.search(r"(\d{4})年(\d{1,2})月", raw)
    if not m:
        return "ERR"
    return "%s-%02d" % (m.group(1), int(m.group(2)))


def _read_fee_values(fee_path: Path, bmap: List[Any]) -> Dict[str, float]:
    """余额表：A 列找余额表代码，值在映射指定的 1 基列号 → {指标代码: 值}。"""
    import xlrd
    wb = xlrd.open_workbook(str(fee_path))
    ws = wb.sheet_by_index(0)
    out: Dict[str, float] = {}
    for code_m, code_f, col_f in bmap:
        for r in range(ws.nrows):
            if str(ws.cell_value(r, 0)).strip() == code_f:
                out[code_m] = _num(ws.cell_value(r, col_f - 1))
                break
    return out


def _needs_adjustment(rows: List[Dict[str, Any]]) -> bool:
    codes = {"11367", "11368", "11369", "11370", "11631"}
    return sum(abs(r["value"]) for r in rows if r["code"] in codes) < 0.01


def _save_xls_copy(src: Path, overrides: Dict[Tuple[int, int], Any]) -> Path:
    """另存 *_已补录* 副本：复制源簿保留全部原值/格式，仅覆盖 overrides
    （0 基 (行, 列) → 新值）。"""
    out = src.parent / f"{src.stem}_已补录{src.suffix}"
    i = 1
    while out.exists():
        out = src.parent / f"{src.stem}_已补录_{i}{src.suffix}"
        i += 1
    if src.suffix.lower() == ".xls":
        import xlrd
        from xlutils.copy import copy as xl_copy
        rb = xlrd.open_workbook(str(src), formatting_info=True)
        w = xl_copy(rb)
        ws = w.get_sheet(0)
        for (r0, c0), v in overrides.items():
            ws.write(r0, c0, v)
        w.save(str(out))
    else:
        from openpyxl import load_workbook
        wb = load_workbook(src)
        ws = wb[wb.sheetnames[0]]
        for (r0, c0), v in overrides.items():
            ws.cell(r0 + 1, c0 + 1, v)
        wb.save(str(out))
    return out


def run_accounting(timeseries_path: Path, source_paths: List[Path],
                   config_xlsx: Path, fee_path: Path,
                   do_backfill: bool) -> Dict[str, Any]:
    """会计链路：惠州市表 5 费用指标全 0 时从余额表补录（do_backfill=True 执行，
    False 返回预览待前端确认），随后与国库同口径归集校验进会计时序表。"""
    cfg = load_config(config_xlsx)
    tol = cfg["tolerance"]
    base = cfg["base_region"]
    regions = cfg["regions"]
    wl = cfg["whitelist"].get("会计", set())
    bmap = cfg["backfill_map"]
    parents = [x for x in str(cfg["backfill_parents"]).split(";") if x]
    fee_codes = {c for c, _f, _col in bmap}
    zero_watch = fee_codes | set(parents)
    ts_path = Path(timeseries_path)
    ts = load_workbook(ts_path) if ts_path.is_file() else Workbook()

    issues: List[List[Any]] = []
    logs: List[List[Any]] = []
    stats: Dict[str, Any] = {"files": 0, "regions_new": [], "cells": 0,
                             "total_bad": 0, "no_data_bad": 0, "wl_hit": 0,
                             "backfilled": 0, "copies": []}
    adjustments: Dict[str, float] = {}
    pre_adjust: Dict[str, float] = {}

    fee_vals: Dict[str, float] = {}
    fee_period = ""
    if do_backfill:
        fp = Path(fee_path) if fee_path else None
        if not fp or not fp.is_file():
            raise RuntimeError("请先选择费用余额表")
        import xlrd
        fwb = xlrd.open_workbook(str(fp))
        fws = fwb.sheet_by_index(0)
        fee_period = _fee_period(fws)
        fee_vals = _read_fee_values(fp, bmap)

    def add_issue(period: str, region: str, code: str, name: str, detail: str,
                  v1: Any, v2: Any, v3: Any, src: str) -> None:
        issues.append([period, region, code, name, detail, v1, v2, v3,
                       _num(v1) - _num(v2), _num(v1) - _num(v3),
                       "会计校验", src, datetime.now()])

    models: Dict[str, Tuple[Dict[int, str], Dict[str, Dict[str, Any]],
                            Dict[int, str], Dict[str, Dict[int, Any]], List[str]]] = {}

    def model_for(region: str):
        if region in models:
            return models[region]
        ws = ts[region] if region in ts.sheetnames else None
        cols, indicators = _snapshot(ws) if ws is not None else ({}, {})
        orig_cols = dict(cols)
        orig_vals = {code: dict(ind["vals"]) for code, ind in indicators.items()}
        new_rows: List[str] = []
        if ws is None:
            stats["regions_new"].append(region)
        models[region] = (cols, indicators, orig_cols, orig_vals, new_rows)
        return models[region]

    for src in map(Path, source_paths):
        if not src.is_file():
            continue
        info = parse_file_name(src.name, regions)
        period, region = info["period"], info["region"]
        if not period or region == "未知区域":
            logs.append([src.name, region or "?", period,
                         "[跳过] 文件名无法解析地区/月份", datetime.now()])
            continue
        rows = read_source_table(src, cfg["value_col"])

        if do_backfill and region in cfg["backfill_regions"] and _needs_adjustment(rows):
            if fee_period != period.replace(".", "-"):
                logs.append([src.name, region, period,
                             f"[余额表日期不符未补录] 余额表 {fee_period}",
                             datetime.now()])
            else:
                by_code = {r["code"]: r for r in rows}
                overrides: Dict[Tuple[int, int], Any] = {}
                deltas = {"11364": 0.0, "11628": 0.0}   # 费用组/增加值组
                for code_m, code_f, col_f in bmap:
                    if code_m not in by_code or code_m not in fee_vals:
                        continue
                    row = by_code[code_m]
                    old = row["value"]
                    val_f = fee_vals[code_m]
                    pre_adjust[f"{period}|{code_m}"] = old
                    adjustments[f"{period}|{code_m}"] = adjustments.get(
                        f"{period}|{code_m}", 0.0) + (val_f - old)
                    row["value"] = val_f
                    overrides[(row["src_row"], 3)] = val_f
                    grp = "11628" if code_m == "11631" else "11364"
                    deltas[grp] += (val_f - old)
                    add_issue(period, "会计补录", code_m, row["name"],
                              "已补录费用数据", val_f, old, 0, src.name)
                # 符合校验：11367+11368+11369+11370 应等于 11631（增加值）
                s4 = sum(by_code.get(x, {}).get("value", 0.0)
                         for x in ("11367", "11368", "11369", "11370") if x in by_code)
                drow11631 = by_code.get("11631")
                v11631 = drow11631["value"] if drow11631 else 0.0
                if abs(s4 - v11631) > cfg["tolerance"]:
                    add_issue(period, base, "11631",
                              (drow11631 or {}).get("name", ""),
                              "增加值(11631)与四项费用(11367~11370)之和不符",
                              v11631, s4, "", src.name)
                # 父项联动：11364/11001 ← 四项费用差额；11628/11371 ← 11631 差额
                for pcode in parents:
                    prow = by_code.get(pcode)
                    if prow is None:
                        continue
                    diff = deltas.get("11364", 0.0) if pcode in ("11364", "11001") \
                        else deltas.get("11628", 0.0)
                    if abs(diff) < 0.01:
                        continue
                    old = prow["value"]
                    pre_adjust[f"{period}|{pcode}"] = old
                    adjustments[f"{period}|{pcode}"] = adjustments.get(
                        f"{period}|{pcode}", 0.0) + diff
                    prow["value"] = old + diff
                    overrides[(prow["src_row"], 3)] = prow["value"]
                    add_issue(period, "会计父项补录", pcode, prow["name"],
                              "父项已同步补录", prow["value"], old, 0, src.name)
                stats["copies"].append(str(_save_xls_copy(src, overrides)))
                stats["backfilled"] += 1

        cols, indicators, orig_cols, orig_vals, new_rows = model_for(region)
        col = next((c for c, p in cols.items() if p == period), None)
        if col is None:
            col = max(cols, default=3) + 1
            cols[col] = period
        changed = 0
        for row in rows:
            ind = indicators.get(row["code"])
            if ind is None:
                indicators[row["code"]] = {"row": None, "name": row["name"],
                                           "rowidx": row["rowidx"],
                                           "vals": {col: row["value"]}}
                new_rows.append(row["code"])
                changed += 1
                stats["cells"] += 1
                continue
            old = ind["vals"].get(col)
            if old is None or old != row["value"]:
                changed += 1
            ind["vals"][col] = row["value"]
            stats["cells"] += 1
            if region == base and row["code"] in zero_watch and abs(row["value"]) < 0.01:
                add_issue(period, region, row["code"], row["name"],
                          "未调整或调整后仍为0", row["value"],
                          row["value"] - adjustments.get(f"{period}|{row['code']}", 0.0),
                          "", src.name)
        if region in ("惠阳区", "博罗县", "龙门县"):
            for row in rows:
                if abs(row["value"]) > tol:
                    stats["no_data_bad"] += 1
                    add_issue(period, region, row["code"], row["name"],
                              "不应有数", row["value"], 0, "", src.name)
        stats["files"] += 1
        logs.append([src.name, region, period,
                     "[数据更新]" if changed else "[数据导入]", datetime.now()])

    # ---- 总分校验（内存 + 会计白名单/调整前值/费用调整标注） ----
    if base in models:
        bcols, bmodel, _oc, _ov, _nr = models[base]
        members = [rg for rg in regions if rg["name"] != base]
        member_models = {rg["name"]: models[rg["name"]][1]
                         for rg in members if rg["name"] in models}
        for code, ind in bmodel.items():
            for c, period in bcols.items():
                city = _num(ind["vals"].get(c))
                total = 0.0
                for mmodel in member_models.values():
                    m = mmodel.get(code)
                    if m:
                        total += _num(m["vals"].get(c))
                if abs(city - total) > tol:
                    stats["total_bad"] += 1
                    pre = pre_adjust.get(f"{period}|{code}", city)
                    detail = "总分校验不平"
                    if code in wl:
                        detail = "总分不平 (白名单指标)"
                        stats["wl_hit"] += 1
                    if abs(city - pre) > tol:
                        detail += " - 费用调整"
                    issues.append([period, "总分校验", code, ind["name"], detail,
                                   city, pre, total, city - pre, city - total,
                                   "会计校验", ts_path.name, datetime.now()])

    # ---- 不应有数校验（按地区规则表） ----
    for rg in regions:
        if not rg["no_data"] or rg["name"] not in models:
            continue
        rcols, rmodel, _oc, _ov, _nr = models[rg["name"]]
        for code, ind in rmodel.items():
            for c, period in rcols.items():
                v = _num(ind["vals"].get(c))
                if abs(v) > tol:
                    stats["no_data_bad"] += 1
                    issues.append([period, rg["name"], code, ind["name"], "不应有数",
                                   v, 0, "", v, v, "会计校验", ts_path.name,
                                   datetime.now()])

    for region, (cols, indicators, orig_cols, orig_vals, new_rows) in models.items():
        _write_back(ts, region, indicators, cols, orig_cols, orig_vals, new_rows)
    if "log" in ts.sheetnames:
        del ts["log"]
    _write_result_sheet(ts, issues, "会计校验")
    ts.save(ts_path)
    stats["issues"] = len(issues)
    stats["rows"] = _issue_rows(issues, wl)
    stats["result_path"] = str(ts_path)
    stats["fee_period"] = fee_period
    return stats


def _fee_names(fee_path: Path) -> Dict[str, str]:
    """余额表：代码 → 科目名称（B 列），供补录预览显示。"""
    import xlrd
    wb = xlrd.open_workbook(str(fee_path))
    ws = wb.sheet_by_index(0)
    out: Dict[str, str] = {}
    for r in range(1, ws.nrows):
        code = str(ws.cell_value(r, 0)).strip()
        if code:
            out[code] = str(ws.cell_value(r, 1)).strip()
    return out

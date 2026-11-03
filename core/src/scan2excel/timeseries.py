# -*- coding: utf-8 -*-
"""时序归集校验（移植 ref-VBA_A1402_CHECK 的 PBCManual 国库/会计链路）。

输入：多份 `YYYY.MM 地区名.xls` 指标表（Sheet0，5 列：指标代码/指标名称/
行序号/余额/上期－余额）+ 一本时序工作簿（每地区一 sheet，行=指标，列=月份）。
行为：归集进时序 → 总分校验（市级 vs Σ县区）→ 不应有数校验 →
时序表内写「校验结果」（13 列）。
全部规则读 config.xlsx：归集配置/地区规则/白名单/补录映射 四 sheet。

性能关键：输入用 python-calamine 读取，运算在内存快照完成，输出整簿用
XlsxWriter 重建。
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .fast_excel import FastWorkbook, FastWorksheet, read_workbook

# ---------------------------------------------------------------- 配色（VBA 同款 RGB）
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
        text = str(v).strip().replace(",", "").replace("，", "")
        text = re.sub(r"[\s\u00a0\u3000]+", "", text)
        try:
            return float(text)
        except (TypeError, ValueError):
            m = re.match(r"[-+]?\d*\.?\d+", text)
        return float(m.group(0)) if m else 0.0


def _cell_text(value: Any) -> str:
    """Convert integral Excel numeric identifiers to stable text."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


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
    wb = read_workbook(config_xlsx)
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


def _header_period(value: Any) -> str:
    """标准化时序表头；数值 2026.1 有一月/十月歧义，保留原样待人工核实。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        raw = str(value).strip()
        if re.fullmatch(r"\d{4}\.1", raw):
            return raw
        raw = f"{value:.2f}"
    else:
        raw = str(value).strip()
    return normalize_period(raw)


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
    """读源指标表（.xls/.xlsx 均用 python-calamine），返回
    [{code, name, rowidx, value}]；取数列按表头名匹配，找不到退回第 4 列。"""
    rows: List[Dict[str, Any]] = []
    wb = read_workbook(path)
    if not wb.worksheets:
        return rows
    ws = wb.worksheets[0]
    matrix = ws.rows
    if not matrix:
        return rows
    headers = [str(h).strip() if h is not None else "" for h in matrix[0]]
    vcol = headers.index(value_col_name) if value_col_name in headers else 3

    def cell(row: Sequence[Any], index: int) -> Any:
        return row[index] if index < len(row) else None

    for r, row in enumerate(matrix[1:], 1):
        code = _cell_text(cell(row, 0))
        if not code:
            continue
        rows.append({"code": code, "name": _cell_text(cell(row, 1)),
                     "rowidx": _cell_text(cell(row, 2)),
                     "value": _num(cell(row, vcol)),
                     "src_row": r,
                     "prev_raw": cell(row, 4)})
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
                    cols[c] = _header_period(v)
            continue
        code = _cell_text(row[0]) if row else ""
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
def _write_result_sheet(ts: FastWorkbook, issues: List[List[Any]], biz: str) -> None:
    if "校验结果" in ts.sheetnames:
        del ts["校验结果"]
    ts.add_sheet(FastWorksheet("校验结果", [_RESULT_HEADERS] + issues))


# ---------------------------------------------------------------- 写回（只写变更格）
def _write_back(ts: FastWorkbook, region: str,
                model: Dict[str, Dict[str, Any]],
                cols: Dict[int, str],
                orig_cols: Dict[int, str],
                orig_vals: Dict[str, Dict[int, Any]],
                new_rows: List[str]) -> None:
    """把内存模型写回 region sheet：新月份列表头、新增指标行、值有变化的格。"""
    ws = ts[region] if region in ts.sheetnames else ts.create_sheet(region)
    for c, period in cols.items():
        cell = ws.cell(1, c)
        cell.value = period
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
        # 统一保存为文本，避免数字代码在不同读取引擎中出现 .0 差异。
        ws.cell(r, 1, code)
        ovals = orig_vals.get(code, {})
        for c, v in ind["vals"].items():
            if c not in orig_cols or _num(ovals.get(c)) != v:
                ws.cell(r, c, v)


def _verify_written_values(wb: FastWorkbook, expected: Dict[Tuple[str, str, str], Tuple[str, float, str]],
                           tolerance: float, biz_label: str,
                           ts_name: str) -> List[List[Any]]:
    """重新打开已保存时序后，逐项核对本次导入的最终源值。"""
    issues: List[List[Any]] = []
    indexes: Dict[str, Tuple[Dict[int, str], Dict[str, Dict[str, Any]]]] = {}
    for region, period, code in expected:
        if region not in indexes and region in wb.sheetnames:
            indexes[region] = _snapshot(wb[region])

    for (region, period, code), (name, expected_value, source_name) in expected.items():
        detail = "写入核对不符"
        actual: Optional[float] = None
        indexed = indexes.get(region)
        if indexed is None:
            detail += "：缺少地区工作表"
        else:
            cols, indicators = indexed
            col = next((c for c, p in cols.items() if p == period), None)
            if col is None:
                detail += "：缺少月份列"
            else:
                indicator = indicators.get(code)
                if indicator is None:
                    detail += "：缺少指标行"
                else:
                    raw = indicator["vals"].get(col)
                    if raw is None or isinstance(raw, str) and not raw.strip():
                        detail += "：时序值为空"
                    else:
                        try:
                            actual = float(str(raw).replace(",", "").strip())
                            if actual != actual or actual in (float("inf"), float("-inf")):
                                raise ValueError("non-finite")
                        except (TypeError, ValueError):
                            actual = None
                            detail += "：时序值异常"
                        if actual is not None:
                            if abs(expected_value - actual) > tolerance + 1e-9:
                                detail += "：金额不一致"
                            else:
                                continue
        delta = expected_value - actual if actual is not None else ""
        issues.append([period, region, code, name, detail,
                       expected_value, "", actual if actual is not None else "",
                       "", delta, biz_label, source_name or ts_name, datetime.now()])
    return issues


def _check_saved_totals(wb: FastWorkbook, base: str, regions: List[Dict[str, Any]],
                        periods: set, tolerance: float, whitelist: set,
                        biz_label: str, ts_name: str,
                        pre_adjust: Optional[Dict[str, float]] = None
                        ) -> Tuple[List[List[Any]], int, int]:
    """仅从已保存的时序工作表读取本期数据，按各地区自己的月份列校验。"""
    if not periods or base not in wb.sheetnames:
        return [], 0, 0
    base_cols, base_rows = _snapshot(wb[base])
    region_data = {}
    for rg in regions:
        name = rg["name"]
        if name != base and name in wb.sheetnames:
            region_data[name] = _snapshot(wb[name])
    ordered_periods = sorted(periods)
    total_cols = {
        period: next((col for col, p in base_cols.items() if p == period), None)
        for period in ordered_periods
    }
    region_period_cols = {
        name: {period: next((col for col, p in cols.items() if p == period), None)
               for period in ordered_periods}
        for name, (cols, _rows) in region_data.items()
    }
    issues: List[List[Any]] = []
    whitelist_hits = 0
    adjustments = pre_adjust or {}
    for code, indicator in base_rows.items():
        for period, base_col in total_cols.items():
            if base_col is None:
                continue
            city = _num(indicator["vals"].get(base_col))
            total = 0.0
            for name, (_cols, rows) in region_data.items():
                region_col = region_period_cols[name][period]
                row = rows.get(code)
                if region_col is not None and row is not None:
                    total += _num(row["vals"].get(region_col))
            if abs(city - total) <= tolerance:
                continue
            pre = adjustments.get(f"{period}|{code}", city)
            detail = "总分校验不平"
            if code in whitelist:
                detail = "总分不平 (白名单指标)"
                whitelist_hits += 1
            if pre_adjust and abs(city - pre) > tolerance:
                detail += " - 费用调整"
            issues.append([period, "总分校验", code, indicator["name"], detail,
                           city, pre if pre_adjust else "", total,
                           city - pre if pre_adjust else "", city - total,
                           biz_label, ts_name, datetime.now()])
    return issues, len(issues), whitelist_hits


# ---------------------------------------------------------------- 国库链路
def run_treasury(timeseries_path: Path, source_paths: List[Path],
                 config_xlsx: Path) -> Dict[str, Any]:
    """国库链路：归集写回 → 回读核对 → 时序总分/不应有数校验。"""
    cfg = load_config(config_xlsx)
    tol = cfg["tolerance"]
    base = cfg["base_region"]
    regions = cfg["regions"]
    wl = cfg["whitelist"].get("国库", set())
    ts_path = Path(timeseries_path)
    ts = read_workbook(ts_path) if ts_path.is_file() else FastWorkbook()

    issues: List[List[Any]] = []
    logs: List[List[Any]] = []
    processed_periods: set = set()
    expected_writes: Dict[Tuple[str, str, str], Tuple[str, float, str]] = {}
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
            expected_writes[(region, period, row["code"])] = (
                row["name"], _num(row["value"]), src.name)
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
        processed_periods.add(period)
        logs.append([src.name, region, period,
                     "[数据更新]" if changed else "[数据导入]", datetime.now()])

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
    ts.save(ts_path)

    # 先从磁盘重新读取实际保存结果，再做写入核对和总分校验。
    saved_ts = read_workbook(ts_path)
    try:
        issues.extend(_verify_written_values(
            saved_ts, expected_writes, tol, "国库校验", ts_path.name))
        total_issues, stats["total_bad"], stats["wl_hit"] = _check_saved_totals(
            saved_ts, base, regions, processed_periods, tol, wl,
            "国库校验", ts_path.name)
        issues.extend(total_issues)
    finally:
        saved_ts.close()

    _write_result_sheet(ts, issues, "国库校验")
    ts.save(ts_path)
    stats["write_checks"] = len(expected_writes)
    stats["write_mismatches"] = sum(
        1 for issue in issues if str(issue[4]).startswith("写入核对不符"))
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
                    "wl": "白名单" in detail})
    return out


# ---------------------------------------------------------------- 会计补录（阶段二）
def _fee_period(ws) -> str:
    """余额表 A2 形如 `2026年9月…` → `YYYY-MM`（VBA GetFeePeriod 口径）。"""
    raw = str(ws.cell(2, 1).value or "").strip()
    m = re.search(r"(\d{4})年(\d{1,2})月", raw)
    if not m:
        return "ERR"
    return "%s-%02d" % (m.group(1), int(m.group(2)))


def _read_fee_values(fee_path: Path, bmap: List[Any]) -> Dict[str, float]:
    """余额表：A 列找余额表代码，值在映射指定的 1 基列号 → {指标代码: 值}。"""
    wb = read_workbook(fee_path)
    ws = wb.worksheets[0] if wb.worksheets else FastWorksheet("空")
    return _fee_values_from_sheet(ws, bmap)


def _fee_values_from_sheet(ws: FastWorksheet, bmap: List[Any]) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for code_m, code_f, col_f in bmap:
        for r in range(ws.max_row):
            if _cell_text(ws.cell(r + 1, 1).value) == code_f:
                out[code_m] = _num(ws.cell(r + 1, col_f).value)
                break
    return out


def _needs_adjustment(rows: List[Dict[str, Any]]) -> bool:
    codes = {"11367", "11368", "11369", "11370", "11631"}
    return sum(abs(r["value"]) for r in rows if r["code"] in codes) < 0.01


def _save_xls_copy(src: Path, overrides: Dict[Tuple[int, int], Any]) -> Path:
    """另存 *_已补录.xlsx* 值副本，覆盖 0 基 (行, 列) 指定的补录值。"""
    wb = read_workbook(src)
    if not wb.worksheets:
        raise RuntimeError(f"源文件没有工作表：{src.name}")
    ws = wb.worksheets[0]
    for (r0, c0), value in overrides.items():
        ws.cell(r0 + 1, c0 + 1, value)

    out = src.parent / f"{src.stem}_已补录.xlsx"
    i = 1
    while out.exists():
        out = src.parent / f"{src.stem}_已补录_{i}.xlsx"
        i += 1
    wb.save(out)
    return out


def run_accounting(timeseries_path: Path, source_paths: List[Path],
                   config_xlsx: Path, fee_path: Path,
                   do_backfill: bool) -> Dict[str, Any]:
    """会计链路：费用及父项调整 → 写回 → 回读核对和时序总分校验。"""
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
    ts = read_workbook(ts_path) if ts_path.is_file() else FastWorkbook()

    issues: List[List[Any]] = []
    logs: List[List[Any]] = []
    processed_periods: set = set()
    expected_writes: Dict[Tuple[str, str, str], Tuple[str, float, str]] = {}
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
        fwb = read_workbook(fp)
        if not fwb.worksheets:
            raise RuntimeError(f"费用余额表没有工作表：{fp.name}")
        fws = fwb.worksheets[0]
        fee_period = _fee_period(fws)
        fee_vals = _fee_values_from_sheet(fws, bmap)

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
            expected_writes[(region, period, row["code"])] = (
                row["name"], _num(row["value"]), src.name)
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
        processed_periods.add(period)
        logs.append([src.name, region, period,
                     "[数据更新]" if changed else "[数据导入]", datetime.now()])

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
    ts.save(ts_path)

    saved_ts = read_workbook(ts_path)
    try:
        issues.extend(_verify_written_values(
            saved_ts, expected_writes, tol, "会计校验", ts_path.name))
        total_issues, stats["total_bad"], stats["wl_hit"] = _check_saved_totals(
            saved_ts, base, regions, processed_periods, tol, wl,
            "会计校验", ts_path.name, pre_adjust)
        issues.extend(total_issues)
    finally:
        saved_ts.close()

    _write_result_sheet(ts, issues, "会计校验")
    ts.save(ts_path)
    stats["write_checks"] = len(expected_writes)
    stats["write_mismatches"] = sum(
        1 for issue in issues if str(issue[4]).startswith("写入核对不符"))
    stats["issues"] = len(issues)
    stats["rows"] = _issue_rows(issues, wl)
    stats["result_path"] = str(ts_path)
    stats["fee_period"] = fee_period
    return stats


def _fee_names(fee_path: Path) -> Dict[str, str]:
    """余额表：代码 → 科目名称（B 列），供补录预览显示。"""
    wb = read_workbook(fee_path)
    if not wb.worksheets:
        return {}
    ws = wb.worksheets[0]
    out: Dict[str, str] = {}
    for r in range(2, ws.max_row + 1):
        code = _cell_text(ws.cell(r, 1).value)
        if code:
            out[code] = _cell_text(ws.cell(r, 2).value)
    return out

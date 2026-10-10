# -*- coding: utf-8 -*-
"""WebApi：业务核心与前端之间的桥接层。

同一份 WebApi 同时服务两个外壳：
  - shell-pywebview：js_api 直调 + pywebview 原生文件对话框
  - shell-flask    ：POST /api/<method> 分发 + Tk 专用线程对话框
所有方法 JSON 进出；前端通过 get_state() 轮询获得全部状态。
"""
from __future__ import annotations

import base64
import os
import queue
import re
import threading
import time
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path

import numpy as np

from openpyxl.utils import get_column_letter
from typing import Any, Dict, List, Optional

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp",
              ".heic", ".heif"}   # HEIC 需 pillow-heif，缺失时读取错误会给出真实原因


class TkRunner:
    """所有 Tk 调用固定在同一个后台线程执行（Tcl 不允许跨线程），
    Flask 多线程环境下弹文件对话框不会崩。"""

    def __init__(self) -> None:
        self._tasks: "queue.Queue" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.Lock()

    def _loop(self) -> None:
        while True:
            fn, result_q = self._tasks.get()
            try:
                result_q.put(("ok", fn()))
            except Exception as exc:  # noqa: BLE001
                result_q.put(("err", exc))
            self._tasks.task_done()

    def run(self, fn, timeout: float = 600.0):
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._loop, daemon=True,
                                                name="tk-dialog")
                self._thread.start()
        result_q: "queue.Queue" = queue.Queue()
        self._tasks.put((fn, result_q))
        status, value = result_q.get(timeout=timeout)
        if status == "err":
            raise value
        return value


_TK = TkRunner()
_COM = TkRunner()


def _com_call(fn):
    """在 COM 专用线程执行并 CoInitialize（COM 不允许跨线程直呼）。"""
    def wrapped():
        import pythoncom
        pythoncom.CoInitialize()
        try:
            return fn()
        finally:
            pythoncom.CoUninitialize()
    return _COM.run(wrapped)


class WebApi:
    """前端可调用的全部方法（方法名即 Flask 端点名）。"""

    def __init__(self, project_root) -> None:
        self.root = Path(project_root)
        self.state: Dict[str, Any] = {
            "app_title": "OCR工具",
            "busy": False,
            "status": "就绪，请先选择照片",
            "images": [],          # [{path,name,status,mode,n_rows,n_cols,elapsed,error,warped,borderless,min_score,ignore_regions}]
            "current": -1,
            "table": None,         # 当前图 {mode,title,rows,merges,scores}
            "preview_mode": "photo",   # photo | overlay
            "has_result": False,
            "high_accuracy": False,    # 高精度识别档（server rec 模型）
            "templates": [],           # 月计表模板名列表
            "scan_mode": "enhanced",   # 扫描件输出模式（enhanced/bw/gray/origin）
            "file_picker_mode": "自动",  # 文件对话框：自动/系统原生/Tk 对话框/浏览器内置
            "scan_busy": False,        # 扫描件生成中
            "scan_done": 0,            # 已生成的张数（进度）
            "scan_images": [],         # 扫描王的照片列表（与 OCR 列表相互独立）
            "dp": {"source": "", "config": "", "template": ""},  # 数据处理页的文件选择
            "dp_mode": "宽表汇总",     # 数据处理类型：宽表汇总/国库数据校验归集/会计数据补录校验
            "ts": {"timeseries": "", "sources": []},   # 国库归集：时序表 + 多个源文件
            "ac": {"timeseries": "", "fee": "", "sources": []},  # 会计补录归集
            "tool": {"excel_files": [], "word_files": [],
                     "excel_fmt": "xlsx", "word_fmt": "docx"},  # 工具页批量转换
            "template_auto": True,     # 自动匹配模板（设置可关：只用每图手动指定的 sheet）
            "auto_rotate": True,       # 自动纠正页面方向（设置可关：照片已摆正时省数秒/张）
            "log": [],
            "version": "25.10.8.0",
        }
        self._pages: List = []     # TablePage 对象（含预览图字节，不进 state）
        self._previews: Dict[str, bytes] = {}   # path → 加图即生成的原图预览 JPEG
        self._thumbs: Dict[str, bytes] = {}     # path → 列表缩略图（更小）
        self._scans: List = []                 # index → 扫描件 JPEG 字节
        self._scan_previews: Dict[tuple, bytes] = {}   # (index, mode, n) → 预览 JPEG
        self._templates: List = []              # 已加载的月计表模板
        self._templates_dir = self.root / "data" / "templates"
        self._exports_dir = self.root / "data" / "exports"     # 识别结果默认存放处
        self._config_xlsx = self.root / "data" / "config" / "config.xlsx"
        self._guoku_tpl = self.root / "data" / "transforms" / "国库模板.xlsx"
        try:
            self._templates_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        self.state["templates_dir"] = str(self._templates_dir)
        self.state["exports_dir"] = str(self._exports_dir)
        self.state["dp"]["config"] = str(self._config_xlsx)   # UI 显示用绝对路径
        self._load_templates()
        self._cancel = threading.Event()
        self._window = None        # pywebview 窗口引用（attach_window 注入）
        self._log("程序启动")
        self._restore_last_paths()

    # ------------------------------------------------------------------ #
    # 查询
    # ------------------------------------------------------------------ #
    def get_state(self) -> Dict[str, Any]:
        return self.state

    def set_scan_mode(self, mode: str) -> Dict[str, Any]:
        """扫描件输出模式：enhanced（彩色增强）/ bw（黑白）/ gray / origin。"""
        if mode in ("enhanced", "bw", "gray", "origin"):
            self.state["scan_mode"] = mode
            self._log("扫描件模式：" + mode)
        return self.state

    def render_scans(self) -> Dict[str, Any]:
        """后台线程：把列表里每张照片渲染成扫描件（当前模式）。"""
        if self.state.get("scan_busy"):
            return self.state
        if not self.state["scan_images"]:
            return self.state
        self.state["scan_busy"] = True
        self.state["scan_done"] = 0

        def worker() -> None:
            try:
                mode = self.state.get("scan_mode", "enhanced")
                n = len(self.state["scan_images"])
                self._scans = [None] * n
                self._scan_previews.clear()
                from .service import Scan2ExcelService
                svc = Scan2ExcelService()
                for idx in range(n):
                    if self._cancel.is_set():
                        break
                    info = self.state["scan_images"][idx]
                    self._log(f"扫描件 {idx + 1}/{n}：{info['name']}")
                    try:
                        arr = svc.render_scan(info["path"], mode)
                        import cv2
                        ok, buf = cv2.imencode(".jpg", arr,
                                               [cv2.IMWRITE_JPEG_QUALITY, 92])
                        self._scans[idx] = buf.tobytes() if ok else None
                    except Exception as exc:  # noqa: BLE001
                        self._scans[idx] = None
                        self._log(f"扫描件失败：{info['name']} —— {exc}")
                    self.state["scan_done"] = idx + 1
                done = sum(1 for b in self._scans if b)
                self.state["status"] = f"扫描件生成完成：{done}/{n}"
                self._log(self.state["status"])
            finally:
                self.state["scan_busy"] = False
        threading.Thread(target=worker, daemon=True).start()
        return self.state

    def get_scan_preview(self, index: int) -> str:
        """单张扫描件预览（dataURL，长边压到 1400 内）。"""
        if not (0 <= index < len(self._scans)) or not self._scans[index]:
            return ""
        key = (index, self.state.get("scan_mode", ""), len(self._scans))
        cached = self._scan_previews.get(key)
        if not cached:
            import cv2
            arr = cv2.imdecode(
                np.frombuffer(self._scans[index], dtype=np.uint8),
                cv2.IMREAD_COLOR)
            if arr is None:
                return ""
            h, w = arr.shape[:2]
            scale = 1400.0 / max(h, w)
            if scale < 1.0:
                arr = cv2.resize(arr, (int(w * scale), int(h * scale)),
                                 interpolation=cv2.INTER_AREA)
            ok, buf = cv2.imencode(".jpg", arr, [cv2.IMWRITE_JPEG_QUALITY, 88])
            if not ok:
                return ""
            cached = buf.tobytes()
            self._scan_previews[key] = cached
        import base64
        return "data:image/jpeg;base64," + base64.b64encode(cached).decode("ascii")

    def export_scan_jpg(self) -> Dict[str, Any]:
        """扫描件逐张存 JPG：第一张照片旁的「扫描件_时间戳」文件夹。"""
        outs = [b for b in self._scans if b]
        if not outs:
            raise RuntimeError("还没有扫描件——先点「生成扫描件」")
        first = (self.state["scan_images"][0]["path"]
                 if self.state["scan_images"] else ".")
        out_dir = Path(first).parent / f"扫描件_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        out_dir.mkdir(parents=True, exist_ok=True)
        n = 0
        for idx, buf in enumerate(self._scans):
            if not buf:
                continue
            name = Path(self.state["scan_images"][idx]["name"]).stem
            (out_dir / f"{idx + 1:02d}_{name}.jpg").write_bytes(buf)
            n += 1
        self._log(f"已导出 {n} 张扫描件 JPG：{out_dir}")
        return {"ok": True, "dir": str(out_dir), "count": n}

    def export_scan_pdf(self, path: str = "") -> Dict[str, Any]:
        """扫描件合并成一个 PDF（系统对话框选保存位置）。"""
        outs = [b for b in self._scans if b]
        if not outs:
            raise RuntimeError("还没有扫描件——先点「生成扫描件」")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = path or self._dialog_save_file(f"扫描件_{stamp}.pdf", kind="pdf")
        if not out:
            return {"ok": False, "canceled": True}
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
        from io import BytesIO
        from PIL import Image
        imgs = []
        for buf in outs:
            im = Image.open(BytesIO(buf))
            if im.mode != "RGB":
                im = im.convert("RGB")
            imgs.append(im)
        imgs[0].save(out, save_all=True, append_images=imgs[1:],
                     resolution=150.0)
        self._log(f"已导出扫描件 PDF（{len(imgs)} 页）：{out}")
        return {"ok": True, "path": str(out), "count": len(imgs)}

    def dp_pick(self, which: str = "source", path: str = "") -> Dict[str, Any]:
        """数据处理页：选择 源文件/规则配置/国库模板。

        显式 path（浏览器内置模式上传落盘后）直接采用；否则按对话框模式弹窗。
        """
        if not path:
            mode = self._picker_mode()
            if mode == "系统原生" and self._window is not None:
                try:
                    paths = _sta_open_dialog(
                        ("Excel 工作簿 (*.xlsx;*.xlsm)",), allow_multiple=False)
                    if isinstance(paths, (list, tuple)):
                        paths = paths[0] if paths else ""
                    path = str(paths) if paths else ""
                except Exception as e:
                    self._log(f"原生对话框不可用：{e}")
                    raise RuntimeError(
                        f"原生文件对话框不可用（{e}）。"
                        "可在设置页把「文件对话框」切换为「浏览器内置」") from e
            elif mode == "Tk 对话框" or self._window is None:
                path = _tk_open_xlsx("选择文件")
            else:
                return self.state   # 浏览器内置：前端应先走 upload_files
            if not path:
                return self.state
        self.state["dp"][which] = str(path)
        self._log(f"数据处理：已选择{ {'source': '源文件', 'config': '规则配置', 'template': '国库模板', 'target': '目标文件'}.get(which, which)}：{path}")
        return self.state

    def dp_run(self) -> Dict[str, Any]:
        """宽表规则汇总（pytools 1-4 同款）：源文件 → 宽表 xlsx + 按规则去重追加到目标簿。"""
        from .data_process import load_rules_v2, wide_summary
        src = self._dp_source()
        cfg = self._config_xlsx
        if not cfg.is_file():
            raise RuntimeError(f"未找到规则配置：{cfg}")
        rules = load_rules_v2(cfg)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = src.parent / f"宽表汇总_{stamp}.xlsx"
        res = wide_summary(src, rules, Path(out))
        self._log(f"宽表汇总完成：{res['rows']} 行 x {res['cols']} 列 → {out}")
        for t in res.get("targets", []):
            if t.get("reason") == "header_mismatch":
                self._log(f"目标写入跳过（表头不匹配）：{t['wb']}::{t['sheet']}")
            elif t.get("written"):
                self._log(f"目标写入：{t['wb']}::{t['sheet']} 新增 {t['added']} 行"
                          f"（输入 {t['input']}，批内去重后 {t['batch']}）")
            else:
                self._log(f"目标写入：{t['wb']}::{t['sheet']} 无新增行")
        return {"ok": True, "path": res["path"], "rows": res["rows"],
                "cols": res["cols"], "date": str(res["date"] or ""),
                "targets": res.get("targets", [])}

    def _latest_export(self) -> Optional[Path]:
        """导出目录里最新的「识别结果」xlsx（没有则 None）。

        只认 识别结果_*.xlsx 命名——宽表/长表等生成产物不会被误当源文件。
        """
        if not self._exports_dir.is_dir():
            return None
        files = [f for f in self._exports_dir.glob("识别结果_*.xlsx")
                 if f.is_file()]
        return max(files, key=lambda f: f.stat().st_mtime) if files else None

    def _dp_source(self) -> Path:
        """数据处理源文件：只认手动选择。"""
        src = self.state["dp"].get("source", "")
        if src and Path(src).is_file():
            return Path(src)
        raise RuntimeError("请先点击「选择源文件」选择识别结果 xlsx")

    def get_summary(self) -> Dict[str, Any]:
        """汇总检查页数据：每张照片的勾稽（条件格式）触发一览。

        bad 携带不平行的行号与科目（1 基行号），供前端免翻页检查；
        未套模板/未识别的页如实标注（它们没有勾稽结论）。
        """
        items: List[Dict[str, Any]] = []
        for idx, info in enumerate(self.state["images"]):
            pg = self._pages[idx] if idx < len(self._pages) else None
            bad: List[Dict[str, Any]] = []
            if pg is not None and getattr(pg, "check_rows", None):
                rows = pg.rows or []
                for r in pg.check_rows:
                    row = rows[r] if 0 <= r < len(rows) else None
                    bad.append({
                        "row": r + 1,
                        "code": str(row[0]) if row and len(row) > 0 else "",
                        "name": str(row[1]) if row and len(row) > 1 else "",
                    })
            items.append({
                "index": idx,
                "name": info.get("name", ""),
                "status": info.get("status", ""),
                "n_rows": info.get("n_rows", 0),
                "n_cols": info.get("n_cols", 0),
                "elapsed": info.get("elapsed", 0),
                "template": str(info.get("template") or ""),
                "title": str(getattr(pg, "title", "") or "") if pg else "",
                "bad": bad,
                "mismatch": info.get("row_mismatch") or {},
            })
        n_ok = sum(1 for it in items
                   if it["status"] == "完成" and it["template"] and not it["bad"])
        return {
            "items": items,
            "stats": {
                "total": len(items),
                "templated": sum(1 for it in items if it["template"]),
                "bad": sum(1 for it in items if it["bad"]),
                "ok": n_ok,
            },
        }

    def get_preview(self) -> str:
        """当前图的预览（dataURL，photo/overlay 由 preview_mode 决定）。

        识别前没有识别产物预览，回退到加图时生成的原图预览——
        选图即应看到照片，这是主流工具的默认行为。
        """
        idx = self.state["current"]
        if not (0 <= idx < len(self.state["images"])):
            return ""
        mode = self.state["preview_mode"]
        page = self._pages[idx] if idx < len(self._pages) else None
        data = None
        if page is not None:
            data = page.overlay_jpeg if mode == "overlay" else page.preview_jpeg
            if not data and mode == "overlay":
                data = page.preview_jpeg     # 叠加图需识别后才有，回退原图
        if not data:
            data = self._previews.get(self.state["images"][idx]["path"])
        if not data:
            return ""
        return "data:image/jpeg;base64," + base64.b64encode(data).decode("ascii")

    # ------------------------------------------------------------------ #
    # 月计表模板：从校对好的一页生成 / 列表 / 删除 / 指定
    # ------------------------------------------------------------------ #
    @staticmethod
    def _tpl_safe(name: str) -> str:
        return re.sub(r'[\/:*?"<>|]+', "_", name or "模板")

    def _load_templates(self) -> None:
        from .template_mode import TableTemplate
        from .xlsx_template import load_xlsx_template
        self._templates = []
        if self._templates_dir.is_dir():
            for f in sorted(self._templates_dir.glob("*.xlsx")):
                try:
                    t = load_xlsx_template(f)
                    self._templates.append(t)
                    if getattr(t, "inferred_regions", False):
                        self._log(f"模板「{t.name}」未定义命名区域，已按启发式"
                                  "推断数字区（建议在 Excel 中用名称管理器"
                                  "定义 数字区域/结构区域）")
                except Exception as exc:
                    self._log(f"xlsx 模板损坏已跳过：{f.name}（{exc}）")
            for f in sorted(self._templates_dir.glob("*.json")):
                try:
                    self._templates.append(TableTemplate.load(f))
                except Exception:
                    self._log(f"模板文件损坏已跳过：{f.name}")
        self.state["templates"] = [t.name for t in self._templates]

    def save_template(self, index: int, name: str) -> Dict[str, Any]:
        """把当前页（已校对）存为 xlsx 模板（命名区域定义结构/数字区）。

        同名文档：追加为新的数据工作表（多页）；首次：新建工作簿。
        """
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法保存模板")
        if not (0 <= index < len(self._pages)) or self._pages[index] is None:
            raise RuntimeError("该页尚无识别结果")
        page = self._pages[index]
        if not page.rows or not page.xs or not page.ys:
            raise RuntimeError("该页缺少网格信息，无法生成模板（仅支持有框线表格页）")
        name = (name or "").strip() or f"模板{len(self._templates) + 1}"
        safe = self._tpl_safe(name)
        self._templates_dir.mkdir(parents=True, exist_ok=True)
        path = self._templates_dir / f"{safe}.xlsx"

        from .template_mode import build_page_from_result
        tpl_page = build_page_from_result("第1页", page.rows, page.xs, page.ys,
                                          merges=page.merges)
        from .xlsx_template import save_sheet_to_workbook
        mid = ({tpl_page.mid_section_row} if tpl_page.mid_section_row >= 0
               else set())
        sheet_name = f"第{len(self._templates[0].pages) + 1}页"             if any(t.name == name for t in self._templates) else "第1页"
        if any(t.name == name for t in self._templates):
            doc0 = next(t for t in self._templates if t.name == name)
            sheet_name = f"第{len(doc0.pages) + 1}页"                 if hasattr(doc0, "pages") else sheet_name
        save_sheet_to_workbook(path, sheet_name, page.rows,
                               tpl_page.value_cols, tpl_page.header_rows,
                               page.merges, mid_rows=mid,
                               replace=not path.is_file(),
                               col_fracs=tpl_page.col_fracs,
                               row_fracs=tpl_page.row_fracs)
        # 逐行勾稽公式（上期借-贷 + 本期发生借-贷 - 期末借-贷 = 0），
        # 写入「勾稽」表；用户对差额列加条件格式即可可视化告警
        try:
            from .xlsx_template import generate_balance_checks
            n_checks = generate_balance_checks(path, sheet_name, tpl_page)
        except Exception as exc:
            n_checks = 0
            self._log(f"勾稽公式生成跳过：{exc}")
        self._load_templates()
        self._log(f"已保存 xlsx 模板：{name} · {sheet_name}"
                  f"（{tpl_page.n_rows}行 x {tpl_page.n_cols}列，"
                  f"{len(tpl_page.value_cols)}个数值列；命名区域：数字区域/结构区域"
                  + (f"；已生成 {n_checks} 行勾稽公式" if n_checks else "") + "）")
        return self.state

    def get_template_detail(self, name: str) -> Dict[str, Any]:
        """模板完整内容（库页查看/细调用）。

        xlsx 模板：每个数据工作表 = 一页，附带命名区域与校验公式信息；
        旧版 json 模板：沿用原结构。
        """
        tpl = next((t for t in self._templates if t.name == name), None)
        if tpl is None:
            raise RuntimeError(f"模板不存在：{name}")
        if hasattr(tpl, "path"):          # xlsx
            # 命名区域是模板的事实来源：数字区域=每月 OCR 的格；
            # 结构区域=冻结文本格（用于与照片匹配/校验）；中缝表头=段界线。
            d = {"name": tpl.name, "format": "xlsx",
                 "file": str(tpl.path), "pages": []}
            for pg in tpl.pages:
                # 单元格区域映射（只读预览着色）：数字区域=绿、结构区域=黄
                region_map = {f"{r},{c}": "num" for (r, c) in pg.num_cells}
                for (r, c) in pg.struct_cells:
                    region_map.setdefault(f"{r},{c}", "struct")
                d["pages"].append({
                    "page_name": pg.page_name,
                    "rows": pg.rows,
                    "merges": pg.merges,
                    "n_rows": pg.n_rows, "n_cols": pg.n_cols,
                    "regions": region_map,
                    "checks": self._cf_rules(pg.page_name),
                })
            return d
        # ---- 旧版 json（预览已废弃：模板库以 sheet 为唯一基准）----
        d = tpl.to_dict()
        d["format"] = "json"
        return d

    def update_template(self, name: str, patch_data: Dict[str, Any]) -> Dict[str, Any]:
        """模板库细调保存：按页编辑（单元格文本/数值列/表头行数/页名）+ 改名。

        模板冻结的科目名与代码若有 OCR 错字，必须能在这里改掉——
        否则错误会每月重复。patch_data.page_index 指定编辑哪一页（默认 0）。
        """
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法修改模板")
        doc = next((t for t in self._templates if t.name == name), None)
        if doc is None or not doc.pages:
            raise RuntimeError(f"模板不存在：{name}")
        if hasattr(doc, "path"):
            return self._update_xlsx_template(doc, patch_data)
        patch_data = patch_data or {}
        pidx = int(patch_data.get("page_index", 0))
        if not (0 <= pidx < len(doc.pages)):
            raise RuntimeError("页码超出范围")
        pg = doc.pages[pidx]

        rows = patch_data.get("rows")
        if isinstance(rows, list) and rows:
            n_cols = max(len(r) for r in rows)
            pg.rows = [[str(c) for c in (list(r) + [""] * n_cols)[:n_cols]]
                       for r in rows]
        if isinstance(patch_data.get("value_cols"), list):
            vc = []
            for c in patch_data["value_cols"]:
                try:
                    ic = int(c)
                except (TypeError, ValueError):
                    continue
                if 0 <= ic < pg.n_cols and ic not in vc:
                    vc.append(ic)
            pg.value_cols = sorted(vc)
        if patch_data.get("header_rows") is not None:
            hr = int(patch_data["header_rows"])
            pg.header_rows = max(0, min(hr, max(0, pg.n_rows - 1)))
        if patch_data.get("code_col") is not None:
            cc = int(patch_data["code_col"])
            pg.code_col = max(0, min(cc, pg.n_cols - 1))
        if isinstance(patch_data.get("col_labels"), list):
            labels = [str(x) for x in patch_data["col_labels"]]
            pg.col_labels = (labels + [""] * pg.n_cols)[:pg.n_cols]
        if patch_data.get("mid_section_row") is not None:
            mr = int(patch_data["mid_section_row"])
            pg.mid_section_row = max(-1, min(mr, pg.n_rows - 1))
        new_page_name = str(patch_data.get("page_name") or "").strip()
        if new_page_name and new_page_name != pg.page_name:
            old_pv = self._templates_dir / f"{self._tpl_safe(doc.name)}_{pg.page_name}.jpg"
            if old_pv.is_file():
                try:
                    old_pv.replace(self._templates_dir
                                   / f"{self._tpl_safe(doc.name)}_{new_page_name}.jpg")
                except OSError:
                    pass
            pg.page_name = new_page_name

        old_name = doc.name
        new_name = str(patch_data.get("new_name") or "").strip()
        if new_name and new_name != old_name:
            old_base = self._tpl_safe(old_name)
            new_base = self._tpl_safe(new_name)
            doc.name = new_name
            # 预览图（主图 + 各页图）随改名迁移
            for f in list(self._templates_dir.glob(f"{old_base}*.jpg")):
                suffix = f.stem[len(old_base):]
                try:
                    f.replace(self._templates_dir / f"{new_base}{suffix}.jpg")
                except OSError:
                    pass
            old_json = self._templates_dir / f"{old_base}.json"
            if old_json.is_file():
                try:
                    old_json.unlink()
                except OSError:
                    pass
            for im in self.state["images"]:
                if im.get("template_name") == old_name:
                    im["template_name"] = new_name
        doc.save(self._templates_dir / f"{self._tpl_safe(doc.name)}.json")
        self._load_templates()
        self._log(f"模板已更新：{doc.name} · {pg.page_name}（{pg.n_rows}行 x {pg.n_cols}列"
                  f" · 数值列 {len(pg.value_cols)} 个 · 表头 {pg.header_rows} 行）")
        return self.state

    def _cf_rules(self, sheet_name: str, name: str = "") -> List[Dict[str, Any]]:
        """读取模板某表上的条件格式规则（勾稽/告警规则，供界面展示）。"""
        doc = next((t for t in self._templates
                    if not name or t.name == name), None)
        if doc is None or not hasattr(doc, "path"):
            return []
        try:
            from openpyxl import load_workbook
            wb = load_workbook(doc.path, data_only=False)
            if sheet_name not in wb.sheetnames:
                wb.close()
                return []
            ws = wb[sheet_name]
            rules = []
            for cf in ws.conditional_formatting:
                for rule in cf.rules:
                    rules.append({
                        "range": str(cf.sqref),
                        "type": rule.type,
                        "formula": (str(rule.formula[0])
                                    if getattr(rule, "formula", None) else ""),
                        "desc": (rule.dxf.fill.bgColor.rgb
                                 if getattr(rule, "dxf", None)
                                 and rule.dxf and rule.dxf.fill
                                 and rule.dxf.fill.bgColor else ""),
                    })
            wb.close()
            return rules
        except Exception:
            return []

    def _update_xlsx_template(self, doc, patch_data: Dict[str, Any]) -> Dict[str, Any]:
        """xlsx 模板的库内编辑：重建该表（文本/数值列/表头/中缝/页名）。

        以"整表重写该工作表 + 重命名区域"实现，校验表与其它页不受影响。
        """
        from openpyxl import load_workbook
        from .xlsx_template import save_sheet_to_workbook
        patch_data = patch_data or {}
        pidx = int(patch_data.get("page_index", 0))
        if not (0 <= pidx < len(doc.pages)):
            raise RuntimeError("页码超出范围")
        pg = doc.pages[pidx]
        rows = patch_data.get("rows") or pg.rows
        n_cols = max(len(r) for r in rows)
        rows = [[str(c) for c in (list(r) + [""] * n_cols)[:n_cols]]
                for r in rows]
        header_rows = int(patch_data.get("header_rows", pg.header_rows))
        mid_row = int(patch_data.get("mid_section_row",
                                     min(pg.mid_rows) if pg.mid_rows else -1))
        vc = sorted({int(c) for c in patch_data.get("value_cols", pg.value_cols)})
        new_page_name = str(patch_data.get("page_name") or pg.page_name or "第1页")

        path = doc.path
        wb = load_workbook(path, data_only=False)
        old_name = pg.page_name
        # 删旧表、建新表（保留位置尽量靠前）
        if old_name in wb.sheetnames:
            del wb[old_name]
        tmp_path = path.with_suffix(".tmp.xlsx")
        wb.save(tmp_path)
        wb.close()
        save_sheet_to_workbook(tmp_path, new_page_name, rows, vc, header_rows,
                               pg.merges,
                               mid_rows=({mid_row} if mid_row >= 0 else set()),
                               replace=False)
        tmp_path.replace(path)

        # 改名（文档级）
        new_name = str(patch_data.get("new_name") or "").strip()
        if new_name and new_name != doc.name:
            new_base = self._tpl_safe(new_name)
            old_jpg = self._templates_dir / f"{self._tpl_safe(doc.name)}.jpg"
            path.replace(self._templates_dir / f"{new_base}.xlsx")
            if old_jpg.is_file():
                try:
                    old_jpg.replace(self._templates_dir / f"{new_base}.jpg")
                except OSError:
                    pass
            for im in self.state["images"]:
                if im.get("template_name") == doc.name:
                    im["template_name"] = new_name
        self._load_templates()
        self._log(f"xlsx 模板已更新：{new_name or doc.name} · {new_page_name}"
                  f"（{len(rows)}行 x {n_cols}列 · 数值列 {len(vc)} 个）")
        return self.state

    def set_all_templates(self, name: str) -> Dict[str, Any]:
        """批量指定：把**所有照片**统一绑定到一个模板文档（各图自动路由到对应页）。

        name="" 恢复全部自动匹配；"__none__" 全部禁用模板。
        """
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法修改模板绑定")
        name = str(name or "")
        if name and name != "__none__" and                 not any(t.name == name for t in self._templates):
            raise RuntimeError(f"模板不存在：{name}")
        n = 0
        for im in self.state["images"]:
            if name == "" :
                im["template_name"] = ""
            else:
                im["template_name"] = name
            n += 1
        desc = {"": "自动匹配", "__none__": "不使用模板"}.get(name, f"模板「{name}」")
        self._log(f"已为全部 {n} 张照片指定：{desc}")
        return self.state

    def set_template_auto(self, on: bool) -> Dict[str, Any]:
        """自动匹配模板开关（设置页）。关闭后仅用每张图显式指定的 sheet。"""
        self.state["template_auto"] = bool(on)
        self._log("自动匹配模板：" + ("开启" if on else "关闭（仅用每张图手动指定的 sheet）"))
        return self.state

    def set_auto_rotate(self, on: bool) -> Dict[str, Any]:
        """自动旋转开关（设置页）。关闭后跳过方向探测，要求照片方向已摆正。"""
        self.state["auto_rotate"] = bool(on)
        self._log("自动旋转：" + ("开启" if on else "关闭（跳过方向探测，请确保照片方向正确）"))
        return self.state

    def set_page_template(self, index: int, name: str) -> Dict[str, Any]:
        """给某页指定模板：""=自动匹配，"__none__"=不用，"文档::页名"=指定页。"""
        if 0 <= index < len(self.state["images"]):
            name = str(name or "")
            doc_name, page_name = name, ""
            if "::" in name:
                doc_name, page_name = name.split("::", 1)
            if doc_name and doc_name != "__none__" and                     not any(t.name == doc_name for t in self._templates):
                doc_name, page_name = "", ""
            im = self.state["images"][index]
            changed = (im.get("template_name") != doc_name
                       or im.get("template_page") != page_name)
            im["template_name"] = doc_name
            im["template_page"] = page_name
            if changed and im.get("status") in ("完成", "失败"):
                # 模板已变：标记为待识别，由"开始识别"统一批量重跑
                im["status"] = "等待"
                im["template"] = ""
                im["warning"] = ""
            if doc_name:
                desc = "不使用" if doc_name == "__none__" else (
                    f"{doc_name}·{page_name}" if page_name else doc_name)
                suffix = "（待识别，点「开始识别」执行）" if changed else ""
                self._log(f"已为「{im['name']}」指定模板：{desc}{suffix}")
        return self.state

    def get_template_pages(self, name: str) -> List[str]:
        """模板文档的所有页名（界面"指定到具体页"用）。"""
        tpl = next((t for t in self._templates if t.name == name), None)
        if tpl is None:
            return []
        pages = getattr(tpl, "pages", []) or []
        return [getattr(p, "page_name", "") for p in pages]

    def get_thumb(self, index: int, target: str = "ocr") -> str:
        """列表缩略图（约 96px）dataURL；按路径缓存，复用原图预览解码。"""
        img_list = self._img_list(target)
        if not (0 <= index < len(img_list)):
            return ""
        path = img_list[index]["path"]
        data = self._thumbs.get(path)
        if data:
            return "data:image/jpeg;base64," + base64.b64encode(data).decode("ascii")
        base = self._previews.get(path)
        if base is None:
            try:
                from .service import imread_unicode, _encode_jpeg
                img = imread_unicode(path)
                base = _encode_jpeg(img)
                self._previews[path] = base
            except Exception:
                return ""
        from .service import thumb_from_jpeg
        thumb = thumb_from_jpeg(base)
        if not thumb:
            return ""
        self._thumbs[path] = thumb
        return "data:image/jpeg;base64," + base64.b64encode(thumb).decode("ascii")

    # ------------------------------------------------------------------ #
    # 照片列表
    # ------------------------------------------------------------------ #
    def _img_list(self, target: str) -> List[Dict[str, Any]]:
        """按 target 取对应照片列表：ocr / scan（两套独立）。"""
        return (self.state["scan_images"] if target == "scan"
                else self.state["images"])

    def choose_images(self, target: str = "ocr",
                      paths: Optional[List[str]] = None) -> int:
        """选照片追加进对应列表：显式路径（浏览器上传）优先，否则弹对话框。"""
        if not paths:
            paths = self._dialog_open_images()
        return self.add_image_paths(paths, target=target)

    def add_image_paths(self, paths: Optional[List[str]],
                        target: str = "ocr") -> int:
        if not paths:
            return 0
        img_list = self._img_list(target)
        expanded: List[str] = []
        for p in paths:
            p = str(p)
            if p.lower().endswith(".pdf") and os.path.isfile(p):
                expanded.extend(self._render_pdf(p))    # PDF 逐页转图片
                continue
            expanded.append(p)
        added = 0
        for p in expanded:
            ext = os.path.splitext(p)[1].lower()
            if ext not in IMAGE_EXTS or not os.path.isfile(p):
                self._log(f"跳过非图片文件：{os.path.basename(p)}")
                continue
            if any(img["path"] == p for img in img_list):
                continue
            img_list.append({
                "path": p, "name": os.path.basename(p),
                "status": "等待", "mode": "table", "n_rows": 0, "n_cols": 0,
                "elapsed": 0.0, "error": "", "warped": False, "borderless": False,
                "min_score": 1.0, "ignore_regions": [],
                "template_name": "", "template_page": "",
                "template": "", "warning": "",
                "fill_mode": "vlookup", "row_mismatch": {}, "checks": [],
            })
            added += 1
            if p not in self._previews:
                try:
                    from .service import imread_unicode, _encode_jpeg
                    img = imread_unicode(p)
                    if img is not None:
                        self._previews[p] = _encode_jpeg(img)
                except Exception as exc:
                    # 预览失败不影响加入列表，但真实原因要立刻可见，
                    # 免得到识别阶段只剩一句笼统的"文件损坏"。
                    self._log(f"预览生成失败：{os.path.basename(p)} → {exc}")
        if added:
            self._log(f"已添加 {added} 张照片（{target}）")
            if target == "ocr":
                if self.state["current"] < 0:
                    self.set_current(0)
                else:
                    self._refresh_status()
        return added

    @staticmethod
    def _render_pdf(pdf_path: str) -> List[str]:
        """PDF 每页渲染成 PNG（约 200DPI），返回图片路径列表。"""
        try:
            import pypdfium2 as pdfium
        except ImportError:
            raise RuntimeError("未安装 pypdfium2，无法识别 PDF")
        out_dir = Path(pdf_path).parent / "_pdf_pages"
        out_dir.mkdir(exist_ok=True)
        stem = Path(pdf_path).stem
        paths = []
        pdf = pdfium.PdfDocument(pdf_path)
        try:
            for i in range(len(pdf)):
                bitmap = pdf[i].render(scale=200 / 72)   # 200DPI
                pil = bitmap.to_pil()
                if pil.mode != "RGB":
                    pil = pil.convert("RGB")
                out = out_dir / f"{stem}_p{i + 1}.png"
                pil.save(out)
                paths.append(str(out))
        finally:
            pdf.close()
        return paths

    def remove_image(self, index: int, target: str = "ocr") -> Dict[str, Any]:
        img_list = self._img_list(target)
        if target == "ocr" and self.state["busy"]:
            raise RuntimeError("识别进行中，无法删除照片")
        if target == "scan" and self.state.get("scan_busy"):
            raise RuntimeError("扫描件生成中，无法删除照片")
        if 0 <= index < len(img_list):
            img_list.pop(index)
            if target == "ocr":
                self._sync_pages()
                if self.state["current"] >= len(self.state["images"]):
                    self.set_current(len(self.state["images"]) - 1)
        self._refresh_status()
        return self.state

    def clear_images(self) -> Dict[str, Any]:
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法清空列表")
        self.state["images"].clear()
        self._pages.clear()
        self._previews.clear()
        self._thumbs.clear()
        self.state["current"] = -1
        self.state["table"] = None
        self.state["has_result"] = False
        self._refresh_status()
        self._log("已清空照片列表（OCR）")
        return self.state

    def clear_scans(self) -> Dict[str, Any]:
        """清空扫描王的照片列表与扫描结果（与 OCR 列表相互独立）。"""
        if self.state.get("scan_busy"):
            raise RuntimeError("扫描件生成中，无法清空列表")
        self.state["scan_images"].clear()
        self._scans = []
        self._scan_previews.clear()
        self.state["scan_done"] = 0
        self._log("已清空照片列表（扫描王）")
        return self.state

    def set_current(self, index: int) -> Dict[str, Any]:
        if not (0 <= index < len(self.state["images"])):
            return self.state
        self.state["current"] = index
        self.state["table"] = None
        self.state["has_result"] = False
        if index < len(self._pages):
            page = self._pages[index]
            if page.error:
                self.state["table"] = None
            elif page.rows:
                self.state["table"] = {
                    "mode": page.mode, "title": page.title,
                    "rows": page.rows, "merges": page.merges,
                    "scores": {k: round(v, 3) for k, v in page.scores.items()},
                    "cell_boxes": page.cell_boxes,
                    "borderless": page.borderless,
                    "checks": list(page.checks or []),
                    "check_rows": list(page.check_rows or []),
                    "header_rows": max(1, int(page.header_rows or 1)),
                }
                self.state["has_result"] = True
        return self.state

    def update_title(self, text: str) -> bool:
        """校对修正：改写当前表的标题。"""
        table = self.state.get("table")
        if not table:
            return False
        idx = self.state["current"]
        if 0 <= idx < len(self._pages) and self._pages[idx] is not None:
            self._pages[idx].title = str(text or "")
        table["title"] = str(text or "")
        return True

    def set_preview_mode(self, mode: str) -> Dict[str, Any]:
        if mode in ("photo", "overlay"):
            self.state["preview_mode"] = mode
        return self.state

    def update_cell(self, r: int, c: int, text: str) -> bool:
        """校对修正：直接改写当前表格单元格。"""
        table = self.state.get("table")
        if not table:
            return False
        rows = table["rows"]
        if not (0 <= r < len(rows)) or not (0 <= c < len(rows[r])):
            return False
        rows[r][c] = str(text)
        return True

    # ------------------------------------------------------------------ #
    # 识别（后台线程）
    # ------------------------------------------------------------------ #
    def start_ocr(self) -> bool:
        if self.state["busy"]:
            return False
        pending = [i for i, img in enumerate(self.state["images"])
                   if img["status"] != "完成"]
        if not pending:
            self._log("所有照片均已识别完成")
            return False
        self._cancel.clear()
        self.state["busy"] = True
        self.state["status"] = "正在识别，请稍候……"
        threading.Thread(target=self._worker, args=(pending, False), daemon=True,
                         name="ocr-worker").start()
        return True

    def start_text_ocr(self, index: int = -1) -> bool:
        """对指定照片（默认当前页）强制按"整页文字"识别，忽略表格线。"""
        if self.state["busy"]:
            return False
        idx = index if 0 <= index < len(self.state["images"]) else self.state["current"]
        if not (0 <= idx < len(self.state["images"])):
            self._log("请先选择要识别文字的照片")
            return False
        self._cancel.clear()
        self.state["busy"] = True
        self.state["status"] = "正在识别文字，请稍候……"
        threading.Thread(target=self._worker, args=([idx], True), daemon=True,
                         name="text-ocr").start()
        return True

    def reload_templates(self) -> Dict[str, Any]:
        """重新扫描模板目录（手动放入/替换 xlsx 后无需重启程序）。"""
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法重载模板")
        n0 = len(self._templates)
        self._load_templates()
        n1 = len(self._templates)
        self._log(f"模板目录已重扫：共 {n1} 个模板"
                  + (f"（较之前 +{n1 - n0}）" if n1 > n0 else ""))
        return self.state

    def set_page_fill_mode(self, index: int, mode: str) -> Dict[str, Any]:
        """设置某页的填充模式：vlookup=按科目对齐（默认）/ position=按位置强制。

        用户在对齐提示中选"继续填充"时调用 position，然后重识别该页。
        """
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法修改填充模式")
        mode = "position" if str(mode) == "position" else "vlookup"
        if 0 <= index < len(self.state["images"]):
            self.state["images"][index]["fill_mode"] = mode
            self._log(f"填充模式已设为："
                      + ("按位置强制填充" if mode == "position" else "按科目对齐（VLOOKUP）"))
        return self.state

    def reprocess_page(self, index: int = -1) -> bool:
        """重新识别指定页（默认当前页）。

        套用模板指定、忽略区域、精度档的最新设置——用于：改模板、画完
        忽略区域、切换精度档后重跑单页。"""
        if self.state["busy"]:
            return False
        idx = index if 0 <= index < len(self.state["images"]) else self.state["current"]
        if not (0 <= idx < len(self.state["images"])):
            return False
        self._cancel.clear()
        self.state["busy"] = True
        self.state["status"] = "正在重新识别，请稍候……"
        self._log(f"重新识别：{self.state['images'][idx]['name']}")
        threading.Thread(target=self._worker, args=([idx], False), daemon=True,
                         name="ocr-reprocess").start()
        return True

    def cancel_ocr(self) -> bool:
        if self.state["busy"]:
            self._cancel.set()
            self._log("已请求取消……")
        return True

    def _worker(self, pending: List[int], force_text: bool = False) -> None:
        from .service import Scan2ExcelService
        service = Scan2ExcelService()
        try:
            for idx in pending:
                if self._cancel.is_set():
                    break
                info = self.state["images"][idx]
                info["status"] = "识别中"
                self._log(f"开始识别：{info['name']}" + ("（文字模式）" if force_text else ""))

                def on_step(text: str, _info=info) -> None:
                    self._log(text, detail=True)

                # 模板解析：图片显式指定的 sheet 优先；"__none__" = 本页不用模板；
                # 未指定且设置开着自动匹配才全库匹配（关掉后只用手动指定的，
                # 用户预先把每张图的 sheet 选好可省去匹配开销）
                tname = str(info.get("template_name") or "")
                if tname == "__none__":
                    tpl_list, tname = None, ""
                elif tname:
                    tpl_list = list(self._templates)
                else:
                    tpl_list = (list(self._templates)
                                if self.state.get("template_auto", True) else None)
                try:
                    page = service.process_image(
                        info["path"], on_step=on_step,
                        cancel_event=self._cancel,
                        force_text=force_text,
                        server_rec=bool(self.state.get("high_accuracy")),
                        ignore_regions=info.get("ignore_regions") or [],
                        templates=tpl_list,
                        template_name=tname,
                        template_page=info.get("template_page", ""),
                        fill_mode=info.get("fill_mode", "vlookup"),
                        auto_rotate=bool(self.state.get("auto_rotate", True)))
                except Exception as exc:  # noqa: BLE001
                    # process_image 内部已兜住单图错误，只有"用户取消"会向外抛；
                    # 其余异常如实报告（此前一律记成"已取消"，掩盖真实原因）
                    if self._cancel.is_set():
                        info["status"] = "等待"
                        self._log("已取消识别")
                        break
                    info["status"] = "失败"
                    info["error"] = f"{type(exc).__name__}: {exc}"
                    self._log(f"识别失败：{info['name']} —— {info['error']}")
                    self._log(traceback.format_exc(limit=3))
                    continue
                self._store_page(idx, page)
                if getattr(page, "geometry_saved", False):
                    # 几何比例已写入模板文件：重载，使同批后续图片用上新几何
                    self._load_templates()
                # 大图识别会累积可观的中间数组（onnx/opencv 分配），
                # 每张处理完回收一次，避免批量越跑越慢（低内存机器更明显）
                import gc
                gc.collect()
                info.update({
                    "status": "完成" if not page.error else "失败",
                    "mode": page.mode,
                    "n_rows": page.n_rows, "n_cols": page.n_cols,
                    "elapsed": round(page.elapsed, 1),
                    "error": page.error or "",
                    "warped": page.warped,
                    "borderless": page.borderless,
                    "template": page.template,
                    "warning": page.warning,
                    "fill_mode": page.fill_mode,
                    "row_mismatch": page.row_mismatch or {},
                    "checks": list(page.checks or []),
                    "min_score": round(page.min_score, 3),
                })
                if page.error:
                    self._log(f"识别失败：{info['name']} —— {page.error}")
                else:
                    shape = (f"文字 {page.n_rows} 行" if page.mode == "text"
                             else f"{page.n_rows}行x{page.n_cols}列")
                    extra = "（无框线结构识别）" if page.borderless else ""
                    self._log(f"识别完成：{info['name']}（{shape}{extra}，"
                              f"{page.elapsed:.1f}s）")
            if not self._cancel.is_set():
                done = sum(1 for i in self.state["images"] if i["status"] == "完成")
                fail = sum(1 for i in self.state["images"] if i["status"] == "失败")
                self.state["status"] = (
                    f"识别完成：成功 {done} 张" if fail == 0
                    else f"识别结束：成功 {done} 张，失败 {fail} 张")
                self._log(self.state["status"])
        except Exception as exc:  # noqa: BLE001
            self.state["status"] = f"发生错误：{exc}"
            self._log("内部错误：" + traceback.format_exc(limit=3))
        finally:
            self.state["busy"] = False
            cur = self.state["current"]
            if cur >= 0:
                self.set_current(cur)

    def _store_page(self, idx: int, page) -> None:
        # _pages 按 images 下标对齐存放（None 占位）
        while len(self._pages) < len(self.state["images"]):
            self._pages.append(None)
        self._pages[idx] = page

    def _sync_pages(self) -> None:
        by_path = {id(pg): pg for pg in self._pages if pg is not None}
        alive = []
        for info in self.state["images"]:
            match = next((pg for pg in by_path.values() if pg.path == info["path"]), None)
            alive.append(match)
        self._pages = alive

    # ------------------------------------------------------------------ #
    # 导出
    # ------------------------------------------------------------------ #
    def export_excel(self, path: str = "", naming: str = "file") -> Dict[str, Any]:
        """所有识别结果保存进一张工作簿（每张照片一个 Sheet）。

        套用了模板的页从模板复制 sheet——**保留条件格式（勾稽标红）、
        合并格与命名区域**，打开 Excel 即见规则；未套模板的页按普通
        方式生成。
        naming: "file"=sheet 名用文件名（默认）；"title"=优先用表标题
        （识别出的表格标题，其次模板页名）。
        每次导出生成全新文件：选择已有文件会整体覆盖，绝不追加。
        """
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法导出")
        pages = [pg for pg in self._pages if pg is not None and pg.rows]
        if not pages:
            raise RuntimeError("还没有可导出的识别结果")
        if not pages:
            raise RuntimeError("还没有可导出的识别结果")
        from .xlsx_template import build_export_workbook, evaluate_checks

        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        # 默认目标路径 = 模板目录：导出的 Excel 可直接当模板
        #（定义好命名区域后点模板库「重新扫描」即可）
        self._exports_dir.mkdir(parents=True, exist_ok=True)
        # 默认存放：data/exports/（数据处理页源文件默认取这里最新的）
        out = path or str(self._exports_dir / f"识别结果_{stamp}.xlsx")
        if not out:
            return {"ok": False, "paths": [], "canceled": True}
        if not out.lower().endswith(".xlsx"):
            out += ".xlsx"

        naming = "title" if str(naming) == "title" else "file"
        sheet_titles = build_export_workbook(Path(out), pages, naming=naming)
        # 「汇总检查」sheet（插在最前）：与「OCR识别结果」页同源的勾稽一览
        from .xlsx_template import append_summary_sheet
        append_summary_sheet(Path(out), pages, sheet_titles)

        # 勾稽：对导出文件中"套了模板的 sheet"评估其条件格式规则
        tpl_sheets = [t for t, pg in zip(sheet_titles, pages)
                      if getattr(pg, "template_file", "")]
        checks: List[str] = []
        if tpl_sheets:
            checks, _bad_rows = evaluate_checks(None, Path(out), sheets=tpl_sheets)
        n_tpl = sum(1 for pg in pages if getattr(pg, "template_file", ""))
        self._log(f"已导出 Excel：{out}（{len(pages)} 个 sheet"
                  + (f"，其中 {n_tpl} 个按模板导出并保留条件格式" if n_tpl else "")
                  + "）")
        for c in checks[:5]:
            self._log("勾稽提示：" + c)
        if checks:
            self.state["status"] = f"导出完成，勾稽不平 {len(checks)} 处（见日志/Excel 标红）"
        self._maybe_register_template(out)
        return {"ok": True, "paths": [out], "canceled": False,
                "checks": checks, "sheets": sheet_titles}

    def export_word(self, path: str = "") -> Dict[str, Any]:
        """所有识别结果保存进一个 Word 文档（每页一节，分页符隔开）。

        表格页渲染为 Word 表格（含合并单元格），文字页为逐行段落。
        每次导出生成全新文件，不追加。
        """
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法导出")
        pages = [pg for pg in self._pages if pg is not None and pg.rows]
        if not pages:
            raise RuntimeError("还没有可导出的识别结果")
        from .word_writer import write_document

        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        out = path or self._dialog_save_file(f"识别结果_{stamp}.docx", kind="docx")
        if not out:
            return {"ok": False, "paths": [], "canceled": True}
        if not out.lower().endswith(".docx"):
            out += ".docx"
        write_document(out, [{"name": pg.name, "title": pg.title, "mode": pg.mode,
                              "rows": pg.rows, "merges": pg.merges}
                             for pg in pages])
        self._log(f"已导出 Word：{out}")
        return {"ok": True, "paths": [out], "canceled": False}

    def _maybe_register_template(self, out: str) -> bool:
        """导出文件落在模板目录时自动重扫，使其立即可作模板使用。"""
        try:
            if Path(out).resolve().parent == self._templates_dir.resolve():
                self._load_templates()
                self._log("导出文件在模板目录内，已自动重新扫描（可作模板使用；"
                          "如需自定义数字/结构区域请在 Excel 名称管理器里定义）")
                return True
        except OSError:
            pass
        return False

    def open_path(self, path: str) -> bool:
        """用系统默认程序打开文件；打不开则退回其所在文件夹。"""
        try:
            if os.path.isfile(path) and path.lower().endswith((".xlsx", ".docx")):
                os.startfile(path)  # noqa: S606 —— Windows 本地工具
            elif os.path.isdir(path):
                os.startfile(path)  # noqa: S606
            elif os.path.isfile(path):
                os.startfile(str(Path(path).parent))  # noqa: S606
            return True
        except Exception:
            return False

    # ------------------------------------------------------------------ #
    # 多格式导出（CSV=当前页；TXT/Markdown/JSON=全部页）
    # ------------------------------------------------------------------ #
    def export_data(self, fmt: str, path: str = "") -> Dict[str, Any]:
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法导出")
        pages = [pg for pg in self._pages if pg is not None and pg.rows]
        if not pages:
            raise RuntimeError("还没有可导出的识别结果")
        fmt = fmt.lower()
        if fmt not in ("csv", "txt", "md", "json"):
            raise ValueError(f"不支持的格式：{fmt}")

        ext = fmt
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        out = path or self._dialog_save_file(f"识别结果_{stamp}.{ext}", kind=ext)
        if not out:
            return {"ok": False, "paths": [], "canceled": True}
        if not out.lower().endswith("." + ext):
            out += "." + ext
        content = {
            "csv": lambda: self._to_csv(self._current_page(pages)),
            "txt": lambda: self._to_txt(pages),
            "md": lambda: self._to_md(pages),
            "json": lambda: self._to_json(pages),
        }[fmt]()
        # csv 用 newline=""：避免 \r\n 被再次翻译成 \r\r\n（Excel 里出空行）
        with open(out, "w", encoding="utf-8-sig" if fmt == "csv" else "utf-8",
                  newline="" if fmt == "csv" else None) as f:
            f.write(content)
        self._log(f"已导出 {fmt.upper()}：{out}")
        return {"ok": True, "paths": [out], "canceled": False}

    def _current_page(self, pages):
        """CSV 导出当前页；当前页无结果时退回第一个有结果的页。"""
        idx = self.state["current"]
        if 0 <= idx < len(self._pages) and self._pages[idx] is not None \
                and self._pages[idx].rows:
            return self._pages[idx]
        return pages[0]

    @staticmethod
    def _to_csv(page) -> str:
        import csv as _csv
        import io
        buf = io.StringIO()
        writer = _csv.writer(buf)
        for row in page.rows:
            writer.writerow(row)
        return buf.getvalue()

    @staticmethod
    def _to_txt(pages) -> str:
        parts = []
        for i, pg in enumerate(pages):
            if i:
                parts.append("")
            parts.append(f"===== {pg.title or pg.name} =====")
            for row in pg.rows:
                parts.append("  ".join(str(c) for c in row))
        return "\n".join(parts)

    @staticmethod
    def _to_md(pages) -> str:
        parts = []
        for i, pg in enumerate(pages):
            if i:
                parts.append("")
            parts.append(f"## {pg.title or pg.name}")
            parts.append("")
            rows = pg.rows
            if not rows:
                continue
            n_cols = max(len(r) for r in rows)

            def fmt_row(row):
                cells = [str(row[c]).replace("|", "\\|") if c < len(row) else ""
                         for c in range(n_cols)]
                return "| " + " | ".join(cells) + " |"

            parts.append(fmt_row(rows[0]))
            parts.append("|" + "---|" * n_cols)
            for row in rows[1:]:
                parts.append(fmt_row(row))
        return "\n".join(parts)

    @staticmethod
    def _to_json(pages) -> str:
        import json
        data = [{"name": pg.name, "title": pg.title, "mode": pg.mode,
                 "rows": pg.rows, "merges": pg.merges,
                 "scores": {k: round(v, 3) for k, v in pg.scores.items()}}
                for pg in pages]
        return json.dumps(data, ensure_ascii=False, indent=1)

    # ------------------------------------------------------------------ #
    # 忽略区域（归一化坐标 x/y/w/h，识别时跳过区域内的文字）
    # ------------------------------------------------------------------ #
    def add_ignore_region(self, index: int, x: float, y: float,
                          w: float, h: float) -> Dict[str, Any]:
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法修改忽略区域")
        if 0 <= index < len(self.state["images"]):
            img = self.state["images"][index]
            img.setdefault("ignore_regions", []).append(
                {"x": round(float(x), 4), "y": round(float(y), 4),
                 "w": round(float(w), 4), "h": round(float(h), 4)})
            self._log(f"已添加忽略区域（共 {len(img['ignore_regions'])} 个，"
                      "重新识别后生效）")
        return self.state

    def clear_ignore_regions(self, index: int) -> Dict[str, Any]:
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法修改忽略区域")
        if 0 <= index < len(self.state["images"]):
            self.state["images"][index]["ignore_regions"] = []
            self._log("已清空忽略区域")
        return self.state

    # ------------------------------------------------------------------ #
    # 高精度识别档（server rec 模型，更准但更慢）
    # ------------------------------------------------------------------ #
    def set_high_accuracy(self, on: bool) -> Dict[str, Any]:
        on = bool(on)
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法切换精度档")
        from .ocr_engine import OcrEngine
        if on and not OcrEngine.is_server_ready():
            # 后台下载模型（约 85MB），完成后档位自动生效
            self.state["status"] = "正在下载高精度模型（约 85MB，仅首次）……"
            self.state["busy"] = True
            threading.Thread(target=self._download_server_worker, daemon=True,
                             name="model-download").start()
            return self.state
        self.state["high_accuracy"] = on
        self._log("已切换到" + ("高精度" if on else "标准") + "识别档")
        return self.state

    def _download_server_worker(self) -> None:
        from .ocr_engine import OcrEngine
        try:
            ok = OcrEngine.download_server_model()
            if ok:
                self.state["high_accuracy"] = True
                self._log("高精度模型就绪，后续识别将使用高精度档")
                self.state["status"] = "高精度模型就绪"
            else:
                self._log("高精度模型下载失败（离线？），已保持标准档")
                self.state["status"] = "高精度模型下载失败，已保持标准档"
                self.state["high_accuracy"] = False
        finally:
            self.state["busy"] = False

    def check_high_accuracy_ready(self) -> bool:
        from .ocr_engine import OcrEngine
        return OcrEngine.is_server_ready()

    # ------------------------------------------------------------------ #
    # 平台对话框（pywebview 优先，回退 Tk）
    # ------------------------------------------------------------------ #
    def attach_window(self, window) -> None:
        self._window = window

    # 文件对话框模式（设置页可选；参考 open-data-audit 的三模式 + 自动档）
    PICKER_MODES = ("自动", "系统原生", "Tk 对话框", "浏览器内置")

    # ---- 数据处理：三模式 + 路径记忆 ----------------------------------
    DP_MODES = ("宽表汇总", "国库数据校验归集", "会计数据补录校验")
    _LAST_PATHS = None   # config/last_paths.json 的内容缓存

    def set_dp_mode(self, mode: str) -> Dict[str, Any]:
        if mode in self.DP_MODES:
            self.state["dp_mode"] = mode
            self._save_last_paths()
            self._log("数据处理类型：" + mode)
        return self.state

    def _load_last_paths(self) -> Dict[str, Any]:
        import json
        p = self.root / "data" / "config" / "last_paths.json"
        if self._LAST_PATHS is None:
            try:
                self._LAST_PATHS = json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                self._LAST_PATHS = {}
        return self._LAST_PATHS

    def _save_last_paths(self) -> None:
        import json
        data = {
            "dp_mode": self.state.get("dp_mode"),
            "宽表汇总": {"source": self.state["dp"].get("source", "")},
            "国库数据校验归集": dict(self.state.get("ts") or {}),
            "会计数据补录校验": dict(self.state.get("ac") or {}),
            "工具": dict(self.state.get("tool") or {}),
        }
        try:
            p = self.root / "data" / "config" / "last_paths.json"
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                         encoding="utf-8")
            self._LAST_PATHS = data
        except Exception:
            pass

    def _restore_last_paths(self) -> None:
        """启动时把上次的选择恢复到 state（文件不存在则忽略该项）。"""
        import os
        data = self._load_last_paths()
        if data.get("dp_mode") in self.DP_MODES:
            self.state["dp_mode"] = data["dp_mode"]
        src = (data.get("宽表汇总") or {}).get("source", "")
        if src and os.path.isfile(src):
            self.state["dp"]["source"] = src

        def alive(v: Any) -> bool:
            return isinstance(v, str) and os.path.isfile(v)

        ts = data.get("国库数据校验归集") or {}
        if alive(ts.get("timeseries")):
            self.state["ts"]["timeseries"] = ts["timeseries"]
        self.state["ts"]["sources"] = [s for s in (ts.get("sources") or []) if alive(s)]
        ac = data.get("会计数据补录校验") or {}
        if alive(ac.get("timeseries")):
            self.state["ac"]["timeseries"] = ac["timeseries"]
        if alive(ac.get("fee")):
            self.state["ac"]["fee"] = ac["fee"]
        self.state["ac"]["sources"] = [s for s in (ac.get("sources") or []) if alive(s)]
        tool = data.get("工具") or {}
        if isinstance(tool.get("excel_files"), list):
            self.state["tool"]["excel_files"] = [s for s in tool["excel_files"] if alive(s)]
        if isinstance(tool.get("word_files"), list):
            self.state["tool"]["word_files"] = [s for s in tool["word_files"] if alive(s)]
        for k in ("excel_fmt", "word_fmt"):
            if tool.get(k):
                self.state["tool"][k] = tool[k]

    def dp_set_paths(self, which: str, path: str,
                     sources: Optional[List[str]] = None) -> Dict[str, Any]:
        """国库/会计模式：显式设置 时序表(timeseries)/费用余额表(fee)/源文件列表(sources)。"""
        mode = self.state.get("dp_mode")
        if mode == "国库数据校验归集":
            target = self.state["ts"]
        elif mode == "会计数据补录校验":
            target = self.state["ac"]
        else:
            raise RuntimeError("当前处理类型不支持该文件行")
        if which == "sources":
            if sources is not None:
                target["sources"] = [str(s) for s in sources if str(s).strip()]
            elif path:
                target["sources"] = [path]
        elif which in ("timeseries", "fee"):
            target[which] = str(path)
        else:
            raise RuntimeError(f"未知文件行：{which}")
        self._save_last_paths()
        self._log(f"{mode}：已设置 {which} = {path or target.get(which)}")
        return self.state

    def dp_clear_sources(self) -> Dict[str, Any]:
        """清空当前模式的源文件列表。"""
        mode = self.state.get("dp_mode")
        if mode == "国库数据校验归集":
            self.state["ts"]["sources"] = []
        elif mode == "会计数据补录校验":
            self.state["ac"]["sources"] = []
        self._save_last_paths()
        return self.state

    def dp_pick_file(self, kind: str) -> Dict[str, Any]:
        """国库/会计模式：选择 时序表(timeseries)/费用余额表(fee)/源文件(sources，多选)。"""
        mode = self.state.get("dp_mode")
        if mode not in ("国库数据校验归集", "会计数据补录校验"):
            raise RuntimeError("请先在数据处理页把「处理类型」切到国库或会计")
        multi = kind == "sources"
        flt = ("数据文件 (*.xls;*.xlsx;*.xlsm)",)
        paths: List[str] = []
        if self._window is not None:
            try:
                paths = [str(p) for p in (_sta_open_dialog(flt, allow_multiple=multi) or [])
                         if str(p).strip()]
            except Exception as e:
                self._log(f"原生对话框不可用：{e}")
                raise RuntimeError(
                    f"原生文件对话框不可用（{e}）。"
                    "可在设置页把「文件对话框」切换为「浏览器内置」") from e
        else:
            paths = _tk_open_data(multi=multi)
        if not paths:
            return self.state
        target = self.state["ts" if mode == "国库数据校验归集" else "ac"]
        if kind == "sources":
            for p in paths:
                if p not in target["sources"]:
                    target["sources"].append(p)
        elif kind == "timeseries":
            target["timeseries"] = paths[0]
        elif kind == "fee":
            target["fee"] = paths[0]
        self._save_last_paths()
        self._log(f"{mode}：已选择 {kind}：{paths}")
        return self.state

    def ts_run(self) -> Dict[str, Any]:
        """国库数据校验归集：源指标表归集进时序表 + 总分/不应有数校验。"""
        from .timeseries import run_treasury
        mode = self.state.get("dp_mode")
        if mode != "国库数据校验归集":
            raise RuntimeError("当前不是「国库数据校验归集」模式")
        ts = self.state.get("ts") or {}
        tp = ts.get("timeseries", "")
        if not tp or not Path(tp).is_file():
            raise RuntimeError("请先选择时序表")
        sources = [s for s in (ts.get("sources") or []) if Path(s).is_file()]
        if not sources:
            raise RuntimeError("请先选择源文件（可多选）")
        cfg = self._config_xlsx
        if not cfg.is_file():
            raise RuntimeError(f"未找到规则配置：{cfg}")
        res = run_treasury(Path(tp), sources, cfg)
        self._save_last_paths()
        self._log(f"国库归集校验完成：{res['files']} 个源文件，写入 {res['cells']} 格，"
                  f"总分异常 {res['total_bad']}，不应有数 {res['no_data_bad']}")
        return {"ok": True, **res}

    def _ac_inputs(self) -> tuple:
        """会计模式输入校验，返回 (时序表, 费用余额表, 源文件列表)。"""
        if self.state.get("dp_mode") != "会计数据补录校验":
            raise RuntimeError("当前不是「会计数据补录校验」模式")
        ac = self.state.get("ac") or {}
        tp = ac.get("timeseries", "")
        if not tp or not Path(tp).is_file():
            raise RuntimeError("请先选择时序表")
        fee = ac.get("fee", "")
        if not fee or not Path(fee).is_file():
            raise RuntimeError("请先选择费用余额表")
        sources = [s for s in (ac.get("sources") or []) if Path(s).is_file()]
        if not sources:
            raise RuntimeError("请先选择源文件（可多选）")
        cfg = self._config_xlsx
        if not cfg.is_file():
            raise RuntimeError(f"未找到规则配置：{cfg}")
        return Path(tp), Path(fee), sources, cfg

    def ac_run(self) -> Dict[str, Any]:
        """会计链路第一步：检测惠州市费用指标是否需要补录。
        需要补录 → 返回预览待前端确认（不写任何数据）；
        不需要 → 直接归集校验并返回统计。"""
        from .timeseries import (run_accounting, parse_file_name,
                                 read_source_table, load_config, _needs_adjustment)
        tp, fee, sources, cfg = self._ac_inputs()
        _cfg = load_config(cfg)
        regions, base, vcol = _cfg["regions"], _cfg["base_region"], _cfg["value_col"]
        pending = []
        for s in sources:
            info = parse_file_name(Path(s).name, regions)
            if info["region"] == base and _needs_adjustment(read_source_table(Path(s), vcol)):
                pending.append(s)
        if pending:
            self._ac_pending = pending
            preview = {"sources": pending, "fee": fee, "note":
                       "检测到惠州市会计表 5 个费用指标（11367~11370、11631）全 0，"
                       "可从费用余额表补录并联动父项；源文件将另存 *_已补录* 副本。"}
            self._log(f"会计补录：{len(pending)} 个惠州市文件可补录，待确认")
            return {"ok": True, "needs_confirm": True, "preview": preview}
        self._ac_pending = []
        return self._ac_finish()

    def ac_confirm(self) -> Dict[str, Any]:
        """会计链路第二步：用户确认后执行补录 + 归集校验。"""
        from .timeseries import run_accounting
        tp, fee, sources, cfg = self._ac_inputs()
        pending = getattr(self, "_ac_pending", None)
        if not pending:
            raise RuntimeError("没有待确认的补录（请先执行一次校验）")
        res = run_accounting(tp, sources, cfg, fee, do_backfill=True)
        self._ac_pending = []
        self._save_last_paths()
        self._log(f"会计补录校验完成：补录 {res['backfilled']} 个文件，总分异常 "
                  f"{res['total_bad']}，不应有数 {res['no_data_bad']}")
        return {"ok": True, **res}

    # ---- 工具页：批量格式转换 ----------------------------------------
    TOOL_EXCEL_FMTS = ("xlsx", "xlsm", "xls", "csv")
    TOOL_WORD_FMTS = ("docx", "doc")

    def tool_set_files(self, kind: str, paths: List[str]) -> Dict[str, Any]:
        """设置工具页文件列表（kind=excel/word）。"""
        key = f"{kind}_files"
        if key not in self.state["tool"]:
            raise RuntimeError(f"未知工具文件类别：{kind}")
        self.state["tool"][key] = [str(p) for p in paths if str(p).strip()]
        self._save_last_paths()
        self._log(f"工具页 {kind} 文件：{len(self.state['tool'][key])} 个")
        return self.state

    def tool_set_format(self, kind: str, fmt: str) -> Dict[str, Any]:
        key = f"{kind}_fmt"
        allowed = self.TOOL_EXCEL_FMTS if kind == "excel" else self.TOOL_WORD_FMTS
        if key not in self.state["tool"] or fmt not in allowed:
            raise RuntimeError(f"未知目标格式：{fmt}")
        self.state["tool"][key] = fmt
        self._save_last_paths()
        return self.state

    def tool_pick_files(self, kind: str) -> Dict[str, Any]:
        """工具页：多选文件加入列表（桌面壳系统对话框 / 无窗口 Tk）。"""
        if kind == "excel":
            flt = ("Excel/表格文件 (*.xls;*.xlsx;*.xlsm;*.csv;*.et)",)
        elif kind == "word":
            flt = ("Word 文档 (*.doc;*.docx)",)
        else:
            raise RuntimeError(f"未知工具文件类别：{kind}")
        paths: List[str] = []
        if self._window is not None:
            try:
                paths = [str(p) for p in (_sta_open_dialog(flt, allow_multiple=True) or [])
                         if str(p).strip()]
            except Exception as e:
                self._log(f"原生对话框不可用：{e}")
                raise RuntimeError(
                    f"原生文件对话框不可用（{e}）。"
                    "可在设置页把「文件对话框」切换为「浏览器内置」") from e
        else:
            paths = _tk_open_data(f"选择要转换的{'表格' if kind == 'excel' else '文档'}",
                                  multi=True)
        if paths:
            lst = self.state["tool"][f"{kind}_files"]
            for p in paths:
                if p not in lst:
                    lst.append(p)
            self._save_last_paths()
            self._log(f"工具页 {kind} 文件 +{len(paths)}")
        return self.state

    def tool_convert(self, kind: str) -> Dict[str, Any]:
        """批量转换（在 COM 专用线程执行，自动 CoInitialize）。"""
        from .office_convert import run_excel_convert, run_word_convert
        tool = self.state["tool"]
        if kind == "excel":
            files = [f for f in tool["excel_files"] if Path(f).is_file()]
            fmt = tool["excel_fmt"]
            if not files:
                raise RuntimeError("请先选择要转换的表格文件")
            self._log(f"批量转换 Excel → {fmt}：{len(files)} 个文件")
            res = _com_call(lambda: run_excel_convert(files, fmt))
        elif kind == "word":
            files = [f for f in tool["word_files"] if Path(f).is_file()]
            fmt = tool["word_fmt"]
            if not files:
                raise RuntimeError("请先选择要转换的文档")
            self._log(f"批量转换 Word → {fmt}：{len(files)} 个文件")
            res = _com_call(lambda: run_word_convert(files, fmt))
        else:
            raise RuntimeError(f"未知转换类别：{kind}")
        self._save_last_paths()
        self._log(f"转换完成：成功 {res['ok']}，跳过 {res['skip']}（引擎 {res['engine']}）")
        res["ok"] = True
        return res

    def set_file_picker_mode(self, mode: str) -> Dict[str, Any]:
        if mode in self.PICKER_MODES:
            self.state["file_picker_mode"] = mode
            self._log("文件对话框模式：" + mode)
        return self.state

    def _picker_mode(self) -> str:
        """解析生效模式：自动 = 有 pywebview 窗口用原生，否则 Tk。"""
        mode = str(self.state.get("file_picker_mode") or "自动")
        if mode not in self.PICKER_MODES or mode == "自动":
            return "系统原生" if self._window is not None else "Tk 对话框"
        return mode

    def _dialog_open_images(self) -> List[str]:
        mode = self._picker_mode()
        if mode != "Tk 对话框" and self._window is not None:
            # 描述里不能有 / 等非 \w 字符——pywebview parse_file_type 会拒绝
            flt = ("图片或PDF (*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;"
                   "*.webp;*.heic;*.heif;*.pdf)",)
            try:
                paths = _sta_open_dialog(flt, allow_multiple=True)
                if isinstance(paths, str):
                    paths = [paths]
                return [str(p) for p in (paths or [])]
            except Exception as e:
                self._log(f"原生对话框不可用：{e}")
                raise RuntimeError(
                    f"原生文件对话框不可用（{e}）。"
                    "可在设置页把「文件对话框」切换为「浏览器内置」") from e
        return _tk_open_images()

    def upload_files(self, items: Optional[List[Dict[str, str]]] = None) -> Dict[str, Any]:
        """浏览器内置模式：前端读文件转 base64 提交，落盘 data/uploads/ 返回路径。"""
        if not items:
            return {"ok": False, "paths": []}
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        updir = self.root / "data" / "uploads" / stamp
        updir.mkdir(parents=True, exist_ok=True)
        import base64
        paths: List[str] = []
        for i, it in enumerate(items):
            name = Path(str(it.get("name") or f"文件{i}")).name or f"文件{i}"
            raw = it.get("data") or ""
            if raw.startswith("data:") and "," in raw:   # 剥掉 data URL 前缀
                raw = raw.split(",", 1)[1]
            try:
                blob = base64.b64decode(raw)
            except Exception:
                continue
            if not blob:
                continue
            p = updir / name
            p.write_bytes(blob)
            paths.append(str(p))
        self._log(f"浏览器上传 {len(paths)} 个文件 → {updir}")
        return {"ok": bool(paths), "paths": paths}

    _SAVE_KINDS = {
        "xlsx": ("Excel 工作簿 (*.xlsx)", ".xlsx", ("Excel 工作簿", "*.xlsx")),
        "docx": ("Word 文档 (*.docx)", ".docx", ("Word 文档", "*.docx")),
        "csv": ("CSV 表格 (*.csv)", ".csv", ("CSV 表格", "*.csv")),
        "txt": ("文本文件 (*.txt)", ".txt", ("文本文件", "*.txt")),
        "md": ("Markdown (*.md)", ".md", ("Markdown", "*.md")),
        "json": ("JSON 文件 (*.json)", ".json", ("JSON 文件", "*.json")),
        "pdf": ("PDF 文件 (*.pdf)", ".pdf", ("PDF 文件", "*.pdf")),
    }

    def _dialog_save_file(self, default_name: str, kind: str = "xlsx",
                          initialdir: str = "") -> str:
        _flt, _ext, tk_ft = self._SAVE_KINDS.get(kind, self._SAVE_KINDS["xlsx"])
        mode = self._picker_mode()
        if mode == "Tk 对话框" or self._window is None:
            return _tk_save_file(default_name, tk_ft, initialdir=initialdir)
        if mode == "浏览器内置":
            # 浏览器内置模式不弹窗：落到导出目录用默认名（路径会回显给用户）
            self._exports_dir.mkdir(parents=True, exist_ok=True)
            return str(self._exports_dir / default_name)
        flt, _e2, _e3 = self._SAVE_KINDS.get(kind, self._SAVE_KINDS["xlsx"])
        try:
            return _sta_save_dialog(default_name, flt, directory=str(initialdir))
        except Exception as e:
            self._log(f"原生保存对话框不可用：{e}")
            raise RuntimeError(
                f"原生保存对话框不可用（{e}）。"
                "可在设置页把「文件对话框」切换为「浏览器内置」") from e

    # ------------------------------------------------------------------ #
    def _log(self, text: str, detail: bool = False) -> None:
        entry = {"time": datetime.now().strftime("%H:%M:%S"), "text": str(text)}
        self.state["log"].append(entry)
        if len(self.state["log"]) > 100:
            del self.state["log"][:-100]

    def _refresh_status(self) -> None:
        n = len(self.state["images"])
        if self.state["busy"]:
            return
        if n == 0:
            self.state["status"] = "就绪，请先选择照片"
        else:
            done = sum(1 for i in self.state["images"] if i["status"] == "完成")
            self.state["status"] = f"共 {n} 张照片，已识别 {done} 张"


# ---------------------------------------------------------------------- #
# 原生 WinForms 对话框（专用 STA 线程；桌面壳使用）
# ---------------------------------------------------------------------- #
def _sta_open_dialog(file_types: tuple, allow_multiple: bool,
                     directory: str = "") -> Optional[tuple]:
    """OpenFileDialog 跑在专用 STA 线程上（冻结态 js 线程是 MTA，
    直调 ShowDialog 会死锁；Form.Invoke 封送在部分环境同样异常）。"""
    from System.Threading import ApartmentState, Thread, ThreadStart
    import System.Windows.Forms as WinForms
    from webview.util import parse_file_type
    box: Dict[str, Any] = {}

    def run() -> None:
        try:
            d = WinForms.OpenFileDialog()
            d.Multiselect = allow_multiple
            d.InitialDirectory = directory or os.environ.get("HOMEPATH", "")
            if file_types:
                d.Filter = "|".join("{0} ({1})|{1}".format(*parse_file_type(f))
                                    for f in file_types)
            d.RestoreDirectory = True
            if d.ShowDialog() == WinForms.DialogResult.OK:
                box["v"] = tuple(d.FileNames)
            else:
                box["v"] = None
        except Exception as e:
            box["e"] = e

    t = Thread(ThreadStart(run))
    t.SetApartmentState(ApartmentState.STA)
    t.Start()
    t.Join()
    if "e" in box:
        raise box["e"]
    return box.get("v")


def _sta_save_dialog(default_name: str, flt: str,
                     directory: str = "") -> str:
    """SaveFileDialog 跑在专用 STA 线程上（口径同 _sta_open_dialog）。"""
    from System.Threading import ApartmentState, Thread, ThreadStart
    import System.Windows.Forms as WinForms
    box: Dict[str, Any] = {}

    def run() -> None:
        try:
            d = WinForms.SaveFileDialog()
            d.InitialDirectory = directory or os.environ.get("HOMEPATH", "")
            d.FileName = default_name
            if flt:
                d.Filter = flt
            d.RestoreDirectory = True
            if d.ShowDialog() == WinForms.DialogResult.OK:
                box["v"] = str(d.FileName)
            else:
                box["v"] = ""
        except Exception as e:
            box["e"] = e

    t = Thread(ThreadStart(run))
    t.SetApartmentState(ApartmentState.STA)
    t.Start()
    t.Join()
    if "e" in box:
        raise box["e"]
    return box.get("v", "")


def _tk_open_data(title: str = "选择数据文件", multi: bool = False) -> List[str]:
    def run():
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            title=title,
            filetypes=[("数据文件", "*.xls *.xlsx *.xlsm"), ("所有文件", "*.*")])
        root.destroy()
        return [str(p) for p in paths]
    return _TK.run(run)


def _tk_open_images() -> List[str]:
    def run():
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            title="选择要识别的照片/PDF",
            filetypes=[("图片/PDF", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.heic *.heif *.pdf"),
                       ("所有文件", "*.*")])
        root.destroy()
        return list(paths)
    return _TK.run(run)


def _tk_open_xlsx(title: str = "选择文件",
                  initialdir: str = "") -> str:
    def run():
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        kw = {}
        if initialdir:
            kw["initialdir"] = str(initialdir)
        path = filedialog.askopenfilename(
            title=title, filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"),
                                    ("所有文件", "*.*")], **kw)
        root.destroy()
        return path
    return _TK.run(run)


def _tk_save_file(default_name: str, tk_ft=("Excel 工作簿", "*.xlsx"),
                  initialdir: str = "") -> str:
    def run():
        import tkinter as tk
        from tkinter import filedialog
        ext = tk_ft[1].lstrip("*")
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        kw = {}
        if initialdir:
            kw["initialdir"] = str(initialdir)
        path = filedialog.asksaveasfilename(
            title="导出", initialfile=default_name,
            defaultextension=ext, filetypes=[tk_ft, ("所有文件", "*.*")], **kw)
        root.destroy()
        return path or ""
    return _TK.run(run)


# ---------------------------------------------------------------------- #
# pywebview 入口（shell-pywebview/run.py 调用）
# ---------------------------------------------------------------------- #
def app_title() -> str:
    return "OCR工具"


def _frontend_html(project_root: Path) -> Path:
    frozen = getattr(__import__("sys"), "frozen", False)
    if frozen:
        return Path(__import__("sys")._MEIPASS) / "web" / "index.html"
    return project_root / "frontend" / "web" / "index.html"


_BRIDGE_ALIAS = (
    "<script>window.addEventListener('pywebviewready',function(){"
    "window.bridge=window.pywebview;"
    "window.dispatchEvent(new Event('bridgeready'));});</script>"
)


def launch_web(project_root) -> None:
    import webview
    html = _frontend_html(Path(project_root))
    api = WebApi(project_root)
    window = webview.create_window(
        app_title(), html.as_uri(), js_api=api,
        width=1240, height=860, min_size=(960, 660), frameless=False)
    api.attach_window(window)
    webview.start(_setup_drag_drop, (window, api), gui="edgechromium")


def _setup_drag_drop(window, api: "WebApi") -> None:
    """订阅 DOM drop 事件：拖入的文件经 WebView2 拿到完整路径直接入列。

    pywebview 会把每个拖入文件的真实路径注入事件的
    dataTransfer.files[i].pywebviewFullPath；前端仅需阻止浏览器默认
    行为，路径收集完全在 Python 侧完成。
    """
    def on_dom_event(event) -> None:
        try:
            if event.get("type") != "drop":
                return
            files = (event.get("dataTransfer") or {}).get("files") or []
            paths = [f.get("pywebviewFullPath") for f in files
                     if f.get("pywebviewFullPath")]
            if paths:
                api.add_image_paths(paths)
        except Exception:
            pass
    try:
        window.dom.document.on("drop", on_dom_event)
    except Exception:
        pass  # 拖拽不可用时仍可用文件对话框

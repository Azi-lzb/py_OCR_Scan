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
from typing import Any, Dict, List, Optional

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}


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


class WebApi:
    """前端可调用的全部方法（方法名即 Flask 端点名）。"""

    def __init__(self, project_root) -> None:
        self.root = Path(project_root)
        self.state: Dict[str, Any] = {
            "app_title": "拍照表格转Excel",
            "busy": False,
            "status": "就绪，请先选择照片",
            "images": [],          # [{path,name,status,mode,n_rows,n_cols,elapsed,error,warped,borderless,min_score,ignore_regions}]
            "current": -1,
            "table": None,         # 当前图 {mode,title,rows,merges,scores}
            "preview_mode": "photo",   # photo | overlay
            "has_result": False,
            "high_accuracy": False,    # 高精度识别档（server rec 模型）
            "templates": [],           # 月计表模板名列表
            "template_auto": True,     # 识别时自动套用模板
            "log": [],
            "version": "25.10.8.0",
        }
        self._pages: List = []     # TablePage 对象（含预览图字节，不进 state）
        self._previews: Dict[str, bytes] = {}   # path → 加图即生成的原图预览 JPEG
        self._thumbs: Dict[str, bytes] = {}     # path → 列表缩略图（更小）
        self._templates: List = []              # 已加载的月计表模板
        self._templates_dir = self.root / "data" / "templates"
        self._load_templates()
        self._cancel = threading.Event()
        self._window = None        # pywebview 窗口引用（attach_window 注入）
        self._log("程序启动")

    # ------------------------------------------------------------------ #
    # 查询
    # ------------------------------------------------------------------ #
    def get_state(self) -> Dict[str, Any]:
        return self.state

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
        self._templates = []
        if self._templates_dir.is_dir():
            for f in sorted(self._templates_dir.glob("*.json")):
                try:
                    self._templates.append(TableTemplate.load(f))
                except Exception:
                    self._log(f"模板文件损坏已跳过：{f.name}")
        self.state["templates"] = [t.name for t in self._templates]

    def save_template(self, index: int, name: str) -> Dict[str, Any]:
        """把当前页（已校对）存为月计表模板。"""
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法保存模板")
        if not (0 <= index < len(self._pages)) or self._pages[index] is None:
            raise RuntimeError("该页尚无识别结果")
        page = self._pages[index]
        if not page.rows or not page.xs or not page.ys:
            raise RuntimeError("该页缺少网格信息，无法生成模板（仅支持有框线表格页）")
        name = (name or "").strip() or f"模板{len(self._templates) + 1}"
        from .template_mode import TableTemplate, build_page_from_result
        tpl_page = build_page_from_result("第1页", page.rows, page.xs, page.ys,
                                          merges=page.merges)
        # 同名文档：追加为新页；否则新建文档（首月正常建一次，以后每月换版式才需要加页）
        doc = next((t for t in self._templates if t.name == name), None)
        if doc is not None:
            tpl_page.page_name = "第%d页" % (len(doc.pages) + 1)
            doc.pages.append(tpl_page)
        else:
            tpl_page.page_name = "第1页"
            doc = TableTemplate(name=name, pages=[tpl_page])
        safe = self._tpl_safe(name)
        doc.save(self._templates_dir / f"{safe}.json")
        # 预览图：每页一张（第N页.jpg）；第一页同时保留文档主预览
        if page.preview_jpeg:
            try:
                pv = self._templates_dir / f"{safe}_{tpl_page.page_name}.jpg"
                pv.write_bytes(page.preview_jpeg)
                if len(doc.pages) == 1:
                    (self._templates_dir / f"{safe}.jpg").write_bytes(page.preview_jpeg)
            except OSError:
                pass
        self._load_templates()
        self._log(f"已保存月计表模板：{name} · {tpl_page.page_name}"
                  f"（{tpl_page.n_rows}行 x {tpl_page.n_cols}列，"
                  f"{len(tpl_page.value_cols)}个数值列）"
                  + ("【已追加为该文档的新页】" if len(doc.pages) > 1 else ""))
        return self.state

    def delete_template(self, name: str) -> Dict[str, Any]:
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法删除模板")
        base = self._tpl_safe(name)
        for f in list(self._templates_dir.glob("*.json")) +                 list(self._templates_dir.glob("*.jpg")):
            try:
                if f.stem == base or f.stem.startswith(base + "_第"):
                    f.unlink()
            except OSError:
                pass
        self._load_templates()
        self._log(f"已删除模板：{name}")
        return self.state

    def get_template_detail(self, name: str) -> Dict[str, Any]:
        """模板完整内容（库页查看/细调用），含基准页预览图 dataURL。"""
        tpl = next((t for t in self._templates if t.name == name), None)
        if tpl is None:
            raise RuntimeError(f"模板不存在：{name}")
        d = tpl.to_dict()
        base = self._tpl_safe(name)
        for i, pg in enumerate(d.get("pages", [])):
            pv = self._templates_dir / f"{base}_{pg['page_name']}.jpg"
            if not pv.is_file() and i == 0:
                pv = self._templates_dir / f"{base}.jpg"
            pg["preview"] = ""
            if pv.is_file():
                try:
                    pg["preview"] = ("data:image/jpeg;base64,"
                                     + base64.b64encode(pv.read_bytes()).decode("ascii"))
                except OSError:
                    pass
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

    def delete_template_page(self, name: str, page_index: int) -> Dict[str, Any]:
        """删除模板文档中的一页；最后一页删除即整文档删除。"""
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法修改模板")
        doc = next((t for t in self._templates if t.name == name), None)
        if doc is None:
            raise RuntimeError(f"模板不存在：{name}")
        pidx = int(page_index)
        if not (0 <= pidx < len(doc.pages)):
            raise RuntimeError("页码超出范围")
        if len(doc.pages) == 1:
            return self.delete_template(name)
        gone = doc.pages.pop(pidx)
        pv = self._templates_dir / f"{self._tpl_safe(name)}_{gone.page_name}.jpg"
        if pv.is_file():
            try:
                pv.unlink()
            except OSError:
                pass
        doc.save(self._templates_dir / f"{self._tpl_safe(name)}.json")
        self._load_templates()
        self._log(f"已删除模板页：{name} · {gone.page_name}")
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
        self.state["template_auto"] = bool(on)
        self._log("模板自动匹配：" + ("开启" if on else "关闭"))
        return self.state

    def set_page_template(self, index: int, name: str) -> Dict[str, Any]:
        """给某页指定模板：""=跟随自动匹配，"__none__"=本页不使用，其余=模板名。"""
        if 0 <= index < len(self.state["images"]):
            name = str(name or "")
            if name and name != "__none__" and                     not any(t.name == name for t in self._templates):
                name = ""
            self.state["images"][index]["template_name"] = name
            if name:
                self._log(f"已为「{self.state['images'][index]['name']}」指定模板："
                          + ("不使用" if name == "__none__" else name))
        return self.state

    def get_thumb(self, index: int) -> str:
        """列表缩略图（约 96px）dataURL；按路径缓存，复用原图预览解码。"""
        if not (0 <= index < len(self.state["images"])):
            return ""
        path = self.state["images"][index]["path"]
        data = self._thumbs.get(path)
        if data:
            return "data:image/jpeg;base64," + base64.b64encode(data).decode("ascii")
        base = self._previews.get(path)
        if base is None:
            try:
                from .service import imread_unicode, _encode_jpeg
                img = imread_unicode(path)
                if img is None:
                    return ""
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
    def choose_images(self) -> int:
        """系统对话框选照片（可多选），追加进列表。"""
        paths = self._dialog_open_images()
        return self.add_image_paths(paths)

    def add_image_paths(self, paths: Optional[List[str]]) -> int:
        if not paths:
            return 0
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
            if any(img["path"] == p for img in self.state["images"]):
                continue
            self.state["images"].append({
                "path": p, "name": os.path.basename(p),
                "status": "等待", "mode": "table", "n_rows": 0, "n_cols": 0,
                "elapsed": 0.0, "error": "", "warped": False, "borderless": False,
                "min_score": 1.0, "ignore_regions": [],
                "template_name": "", "template": "", "warning": "",
            })
            added += 1
            if p not in self._previews:
                try:
                    from .service import imread_unicode, _encode_jpeg
                    img = imread_unicode(p)
                    if img is not None:
                        self._previews[p] = _encode_jpeg(img)
                except Exception:
                    pass   # 预览生成失败不影响识别；识别后仍有结果预览
        if added:
            self._log(f"已添加 {added} 张照片")
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

    def remove_image(self, index: int) -> Dict[str, Any]:
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法删除照片")
        if 0 <= index < len(self.state["images"]):
            self.state["images"].pop(index)
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
        self._log("已清空照片列表")
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

                # 模板解析：显式指定 > 全局自动开关；"__none__" = 本页不用模板
                tname = str(info.get("template_name") or "")
                if tname == "__none__":
                    tpl_list, tname = None, ""
                elif tname:
                    tpl_list = list(self._templates)
                else:
                    tpl_list = (list(self._templates)
                                if self.state.get("template_auto") else None)
                try:
                    page = service.process_image(
                        info["path"], on_step=on_step,
                        cancel_event=self._cancel,
                        force_text=force_text,
                        server_rec=bool(self.state.get("high_accuracy")),
                        ignore_regions=info.get("ignore_regions") or [],
                        templates=tpl_list,
                        template_name=tname)
                except Exception:  # process_image 只在"用户取消"时向外抛
                    info["status"] = "等待"
                    self._log("已取消识别")
                    break
                self._store_page(idx, page)
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
                self.state["status"] = f"识别结束：成功 {done} 张，失败 {fail} 张"
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
    def export_excel(self, path: str = "") -> Dict[str, Any]:
        """所有识别结果保存进一张工作簿（每张照片一个 Sheet）。

        每次导出生成全新文件：选择已有文件会整体覆盖，绝不追加。
        """
        if self.state["busy"]:
            raise RuntimeError("识别进行中，无法导出")
        pages = [pg for pg in self._pages if pg is not None and pg.rows]
        if not pages:
            raise RuntimeError("还没有可导出的识别结果")
        from .excel_writer import write_workbook

        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        out = path or self._dialog_save_file(f"识别结果_{stamp}.xlsx", kind="xlsx")
        if not out:
            return {"ok": False, "paths": [], "canceled": True}
        if not out.lower().endswith(".xlsx"):
            out += ".xlsx"
        write_workbook(out, [{"name": pg.name, "title": pg.title,
                              "plain": pg.mode == "text",
                              "rows": pg.rows, "merges": pg.merges}
                             for pg in pages])
        self._log(f"已导出 Excel：{out}")
        return {"ok": True, "paths": [out], "canceled": False}

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

    def _dialog_open_images(self) -> List[str]:
        if self._window is not None:
            try:
                import webview
                flt = ("图片/PDF (*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp;*.pdf)",)
                paths = self._window.create_file_dialog(
                    webview.OPEN_DIALOG, allow_multiple=True, file_types=flt)
                if isinstance(paths, str):
                    paths = [paths]
                return [str(p) for p in (paths or [])]
            except Exception:
                pass  # 回退 Tk
        return _tk_open_images()

    _SAVE_KINDS = {
        "xlsx": ("Excel 工作簿 (*.xlsx)", ".xlsx", ("Excel 工作簿", "*.xlsx")),
        "docx": ("Word 文档 (*.docx)", ".docx", ("Word 文档", "*.docx")),
        "csv": ("CSV 表格 (*.csv)", ".csv", ("CSV 表格", "*.csv")),
        "txt": ("文本文件 (*.txt)", ".txt", ("文本文件", "*.txt")),
        "md": ("Markdown (*.md)", ".md", ("Markdown", "*.md")),
        "json": ("JSON 文件 (*.json)", ".json", ("JSON 文件", "*.json")),
    }

    def _dialog_save_file(self, default_name: str, kind: str = "xlsx") -> str:
        flt, _ext, tk_ft = self._SAVE_KINDS.get(kind, self._SAVE_KINDS["xlsx"])
        if self._window is not None:
            try:
                import webview
                paths = self._window.create_file_dialog(
                    webview.SAVE_DIALOG, save_filename=default_name, file_types=(flt,))
                # pywebview 部分平台 SAVE_DIALOG 返回单元素 list/tuple 而非字符串
                if isinstance(paths, (list, tuple)):
                    paths = paths[0] if paths else ""
                return str(paths) if paths else ""
            except Exception:
                pass
        return _tk_save_file(default_name, tk_ft)

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
# Tk 对话框（Flask 外壳 / 无 pywebview 窗口时使用）
# ---------------------------------------------------------------------- #
def _tk_open_images() -> List[str]:
    def run():
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            title="选择要识别的照片/PDF",
            filetypes=[("图片/PDF", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.pdf"),
                       ("所有文件", "*.*")])
        root.destroy()
        return list(paths)
    return _TK.run(run)


def _tk_save_file(default_name: str, tk_ft=("Excel 工作簿", "*.xlsx")) -> str:
    def run():
        import tkinter as tk
        from tkinter import filedialog
        ext = tk_ft[1].lstrip("*")
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.asksaveasfilename(
            title="导出", initialfile=default_name,
            defaultextension=ext, filetypes=[tk_ft, ("所有文件", "*.*")])
        root.destroy()
        return path or ""
    return _TK.run(run)


# ---------------------------------------------------------------------- #
# pywebview 入口（shell-pywebview/run.py 调用）
# ---------------------------------------------------------------------- #
def app_title() -> str:
    return "拍照表格转Excel"


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

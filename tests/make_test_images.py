# -*- coding: utf-8 -*-
"""合成测试用"拍照表格"图片。

生成 3 张图到 tests/data/：
  1. inventory_flat.png   平拍标准库存表（6列x11行，全框线）
  2. production_merge.png 带合并表头的产量表（2行表头，首列+跨2列合并）
  3. inventory_tilt.png   库存表的透视倾斜+噪声+亮度渐变版本（模拟手机拍照）
并同时输出 tests/data/expected.json 逐格真值，供 test_pipeline.py 比对。
"""
from __future__ import annotations

import json
import math
import os
import random
import sys

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "data")

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]


def load_font(size: int):
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


INVENTORY_HEADER = ["序号", "物料编码", "物料名称", "规格型号", "单位", "数量"]
INVENTORY_ROWS = [
    ["1", "WL-1001", "深沟球轴承", "6204-2RS", "个", "120"],
    ["2", "WL-1002", "内六角螺栓", "M8x30", "盒", "45"],
    ["3", "WL-1003", "圆锥滚子轴承", "30205", "个", "86"],
    ["4", "WL-1004", "油封", "TC-35x52x7", "件", "230"],
    ["5", "WL-1005", "三角皮带", "B-1250", "条", "64"],
    ["6", "WL-1006", "O型密封圈", "ID-25x3.5", "个", "500"],
    ["7", "WL-1007", "直线导轨滑块", "HGH20CA", "套", "12"],
    ["8", "WL-1008", "滚珠丝杆", "SFU1605", "根", "8"],
    ["9", "WL-1009", "联轴器", "LX-L2", "件", "36"],
    ["10", "WL-1010", "伺服电机", "750W", "台", "5"],
]


def draw_table(header, rows, col_widths, row_height=52, header_height=56,
               merged_header=None, out_name=None, table_title="表"):
    """画一张全框线表格，返回 (PIL图, 真值rows, merges)。"""
    n_cols = len(header)
    head_rows = 2 if merged_header else 1
    total_w = sum(col_widths)
    title_h = 64
    # 画布高度必须容纳：标题 + 表头（合并表头占 2 行）+ 全部数据行
    total_h = title_h + header_height * head_rows + len(rows) * row_height
    img = Image.new("RGB", (total_w + 112, total_h + 112), "white")
    d = ImageDraw.Draw(img)
    ox, oy = 56, 56  # 表格原点（留足纸边距，贴近真实纸张）

    # 标题
    font_title = load_font(30)
    title = table_title
    tw = d.textlength(title, font=font_title)
    d.text((ox + (total_w - tw) / 2, oy + 12), title, fill="black", font=font_title)

    # 标题占两行高度时统一基线：正文区从 oy+title_h 开始
    ty = oy + title_h

    # ---- 网格线（模拟打印：先画线再填字）----
    xs = [ox]
    for wdt in col_widths:
        xs.append(xs[-1] + wdt)
    line = 3
    # 若有合并表头则表头占 2 行
    head_rows = 2 if merged_header else 1
    ys = [ty]
    for _ in range(head_rows):
        ys.append(ys[-1] + header_height)
    for _ in range(len(rows)):
        ys.append(ys[-1] + row_height)

    for x in xs:
        d.line([(x, ty), (x, ys[-1])], fill="black", width=line)
    for y in ys:
        d.line([(ox, y), (ox + total_w, y)], fill="black", width=line)

    # 合并单元格区域：擦掉其内部的线段（真实合并表格中间无线）
    if merged_header:
        erase_w = line + 4
        for (hr, hc, hspan, _t) in merged_header[0]:
            if hspan > 1:      # 跨列：擦中间竖线段
                for ci in range(hc + 1, hc + hspan):
                    d.rectangle([xs[ci] - erase_w // 2, ys[hr] + line,
                                 xs[ci] + erase_w // 2, ys[hr + 1] - line], fill="white")
            else:              # 跨行（首列车间）：擦中间横线段
                d.rectangle([xs[hc] + line, ys[hr + 1] - erase_w // 2,
                             xs[hc + 1] - line, ys[hr + 1] + erase_w // 2], fill="white")

    font_head = load_font(24)
    font_body = load_font(22)

    def cell_text(r, c, text, font):
        x0, x1 = xs[c], xs[c + 1]
        y0, y1 = ys[r], ys[r + 1]
        wt = d.textlength(text, font=font)
        bbox = font.getbbox(text)
        th = bbox[3] - bbox[1]
        d.text((x0 + (x1 - x0 - wt) / 2, y0 + (y1 - y0 - th) / 2 - bbox[1]),
               text, fill="black", font=font)

    true_rows = []
    merges = []

    if merged_header:
        # merged_header: [(row, col_start, col_span, text), ...] + 第2行列头
        placed = {}
        first, second = merged_header
        # 第 0 行：车间(跨2行) + 一月(跨2列) + 二月(跨2列)
        for (hr, hc, hspan, text) in first:
            cell_text(hr, hc, text, font_head)
            if hspan > 1:
                # 跨列：补画上方行线？不画（合并即无线）
                merges.append([hr, hc, 1, hspan])
            else:
                merges.append([hr, hc, 2, 1])  # 首列跨2行
        for c, text in enumerate(second):
            cell_text(1, c, text, font_head)
        true_rows.append([""] * n_cols)
        true_rows.append([""] * n_cols)
        for (hr, hc, hspan, text) in first:
            true_rows[hr][hc] = text
        for c, text in enumerate(second):
            true_rows[1][c] = text
    else:
        for c, text in enumerate(header):
            cell_text(0, c, text, font_head)
        true_rows.append(list(header))

    for r, row in enumerate(rows):
        true_rows.append(list(row))
        for c, text in enumerate(row):
            cell_text(len(true_rows) - 1, c, text, font_body)

    return img, true_rows, merges


def add_photograph_effects(pil_img: Image.Image, seed: int = 7) -> Image.Image:
    """模拟拍照：透视倾斜 + 高斯噪声 + 亮度渐变，垫浅色背景。"""
    rng = random.Random(seed)
    img = cv2.cvtColor(np.asarray(pil_img), cv2.COLOR_RGB2BGR)
    h, w = img.shape[:2]
    # 放到浅灰桌面背景上，留边
    bg = np.full((int(h * 1.25), int(w * 1.3), 3), (168, 172, 176), dtype=np.uint8)
    bx, by = (bg.shape[1] - w) // 2, (bg.shape[0] - h) // 2
    bg[by:by + h, bx:bx + w] = img

    # 透视倾斜：四角随机偏移 ±3.5%
    bh, bw = bg.shape[:2]
    src = np.float32([[bx, by], [bx + w, by], [bx + w, by + h], [bx, by + h]])
    kx, ky = bw * 0.035, bh * 0.035
    dst = np.float32([
        [bx - rng.uniform(0.2, 1) * kx, by - rng.uniform(0.2, 1) * ky],
        [bx + w + rng.uniform(0.2, 1) * kx, by - rng.uniform(0.2, 1) * ky],
        [bx + w + rng.uniform(0.2, 1) * kx, by + h + rng.uniform(0.2, 1) * ky],
        [bx - rng.uniform(0.2, 1) * kx, by + h + rng.uniform(0.2, 1) * ky],
    ])
    m = cv2.getPerspectiveTransform(src, dst)
    out = cv2.warpPerspective(bg, m, (bw, bh), borderValue=(150, 150, 150))

    # 高斯噪声
    noise = np.random.default_rng(seed).normal(0, 6, out.shape).astype(np.float32)
    out = np.clip(out.astype(np.float32) + noise, 0, 255).astype(np.uint8)
    # 轻微亮度渐变（左上亮右下暗）
    grad = np.linspace(1.06, 0.9, out.shape[1], dtype=np.float32)
    out = np.clip(out.astype(np.float32) * grad[None, :, None], 0, 255).astype(np.uint8)
    return Image.fromarray(cv2.cvtColor(out, cv2.COLOR_BGR2RGB))


def add_scan_effects(pil_img: Image.Image, seed: int = 23, angle_deg: float = 1.6,
                     lowres: bool = False) -> Image.Image:
    """模拟平板/ADF 扫描仪输出：小幅歪斜、边缘深色条、低对比度、
    扫描噪声、光源不均渐变、JPEG 压缩伪影。

    lowres=True 再叠加"低分辨率扫描"效果：整图降到 ~1100px 长边 +
    轻微光学模糊（模拟 200DPI 扫描或 IM 传输压缩，文字明显变小变虚）。
    """
    rng = random.Random(seed)
    img = cv2.cvtColor(np.asarray(pil_img), cv2.COLOR_RGB2BGR)
    h, w = img.shape[:2]

    # 1) ADF 进纸歪斜：整图小角度旋转，背景为扫描仪盖板深灰
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle_deg, 1.0)
    img = cv2.warpAffine(img, m, (w, h), borderValue=(95, 95, 95))

    # 2) 扫描边缘深色条（盖板未盖严漏光/纸张边缘），加在一侧
    side = rng.choice(["left", "right"])
    bar_w = int(w * rng.uniform(0.012, 0.02))
    if side == "left":
        img[:, :bar_w] = (rng.randint(35, 70),) * 3
    else:
        img[:, -bar_w:] = (rng.randint(35, 70),) * 3

    # 3) 整体发灰（对比度下降）+ 轻微偏色
    img = cv2.convertScaleAbs(img, alpha=0.86, beta=22)

    # 4) 扫描噪声（高斯 + 少量椒盐）
    noise = np.random.default_rng(seed).normal(0, 5, img.shape).astype(np.float32)
    img = np.clip(img.astype(np.float32) + noise, 0, 255).astype(np.uint8)
    n_salt = int(img.size * 0.00012)
    ys = np.random.default_rng(seed + 1).integers(0, h, n_salt)
    xs = np.random.default_rng(seed + 2).integers(0, w, n_salt)
    img[ys, xs] = 90  # 深色噪点

    # 5) 光源不均：对角方向的暗角渐变
    gx = np.linspace(1.0, 0.88, w, dtype=np.float32).reshape(1, w, 1)
    gy = np.linspace(1.0, 0.93, h, dtype=np.float32).reshape(h, 1, 1)
    img = np.clip(img.astype(np.float32) * (gx * gy), 0, 255)
    img = img.astype(np.uint8)

    # 6) JPEG 有损压缩（由调用方保存为 .jpg 时生效；此处先做一次预压缩模拟）
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, 87])
    if ok:
        img = cv2.imdecode(buf, cv2.IMREAD_COLOR)

    # 7) 低分辨率扫描：降采样到 ~1100px 长边 + 轻微光学模糊。
    #    文字与线条同步变小变虚（等效 200DPI 扫描 / IM 压缩传输），
    #    是对切格与 OCR 最苛刻的真实场景。
    if lowres:
        side = max(img.shape[:2])
        scale = 1100.0 / side
        img = cv2.resize(img, (int(img.shape[1] * scale), int(img.shape[0] * scale)),
                         interpolation=cv2.INTER_AREA)
        img = cv2.GaussianBlur(img, (0, 0), 0.9)
        ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, 78])
        if ok:
            img = cv2.imdecode(buf, cv2.IMREAD_COLOR)
    return Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))


NOTES_TITLE = "备件入库明细（2026年10月）"
NOTES_LINES = [
    "1. 深沟球轴承 6204-2RS，入库 120 个，单价 8.50 元",
    "2. 内六角螺栓 M8x30，入库 45 盒，单价 22.00 元",
    "3. 圆锥滚子轴承 30205，入库 86 个，单价 15.80 元",
    "4. 油封 TC-35x52x7，入库 230 件，单价 6.20 元",
    "5. 三角皮带 B-1250，入库 64 条，单价 12.00 元",
    "6. O型密封圈 ID-25x3.5，入库 500 个，单价 0.90 元",
    "供应商：华东机电设备有限公司",
    "入库日期：2026-10-06",
    "经手人：张敏  复核人：李强",
]


def draw_document(title, lines, width=980, out_title=True):
    """画一页纯文字文档（无表格线），返回 (PIL图, 真值rows)。"""
    font_title = load_font(32)
    font_body = load_font(24)
    pad = 60
    line_h = 52
    total_h = pad * 2 + (56 if out_title else 0) + len(lines) * line_h + 30
    img = Image.new("RGB", (width, total_h), "white")
    d = ImageDraw.Draw(img)
    y = pad
    if out_title:
        d.text((pad, y), title, fill="black", font=font_title)
        y += 56 + 12
    for line in lines:
        d.text((pad, y), line, fill="black", font=font_body)
        y += line_h
    rows = ([[title]] if out_title else []) + [[l] for l in lines]
    return img, rows


def draw_borderless_table(rows, col_xs, row_ys, out_title=""):
    """画一张无框线表格：文字按列对齐，不画任何线。

    返回 (PIL图, 真值rows)。
    """
    font_head = load_font(26)
    font_body = load_font(24)
    pad = 56
    width = col_xs[-1] + 240
    height = row_ys[-1] + 90
    img = Image.new("RGB", (width, height), "white")
    d = ImageDraw.Draw(img)
    x0 = col_xs[0]
    if out_title:
        ft = load_font(30)
        d.text((x0 + 60, 24), out_title, fill="black", font=ft)
    for r, row in enumerate(rows):
        f = font_head if r == 0 else font_body
        for c, text in enumerate(row):
            # 数字列右对齐（真实表格习惯），文本列左对齐
            col_x = col_xs[c] if not (r > 0 and c == len(row) - 1) else col_xs[c] + 120
            d.text((col_x, row_ys[r]), text, fill="black", font=f)
    return img, [list(r) for r in rows]


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    expected = {}

    # 1. 平拍库存表
    img, rows, merges = draw_table(
        INVENTORY_HEADER, INVENTORY_ROWS,
        col_widths=[90, 160, 220, 200, 90, 110],
        table_title="仓库月末盘点表",
    )
    p = os.path.join(OUT_DIR, "inventory_flat.png")
    img.save(p)
    expected["inventory_flat.png"] = {"rows": rows, "merges": merges,
                                      "title": "仓库月末盘点表"}

    # 2. 合并表头产量表
    header = ["车间", "一月计划", "一月实际", "二月计划", "二月实际"]
    prod_rows = [
        ["冲压一班", "12500", "12380", "13000", "12960"],
        ["冲压二班", "11000", "11250", "11500", "11400"],
        ["焊接班", "8600", "8420", "9000", "9130"],
        ["装配一班", "6200", "6350", "6500", "6480"],
        ["装配二班", "5800", "5760", "6000", "6120"],
    ]
    first = [(0, 0, 1, "车间"), (0, 1, 2, "一月"), (0, 3, 2, "二月")]
    second = ["", "计划", "实际", "计划", "实际"]
    img, rows, merges = draw_table(
        header, prod_rows, col_widths=[150, 150, 150, 150, 150],
        merged_header=(first, second), table_title="车间月度产量统计",
    )
    p = os.path.join(OUT_DIR, "production_merge.png")
    img.save(p)
    expected["production_merge.png"] = {"rows": rows, "merges": merges,
                                        "title": "车间月度产量统计"}

    # 3. 倾斜拍照库存表（同表1数据）
    img, rows, merges = draw_table(
        INVENTORY_HEADER, INVENTORY_ROWS,
        col_widths=[90, 160, 220, 200, 90, 110],
        table_title="仓库月末盘点表",
    )
    photo = add_photograph_effects(img, seed=11)
    p = os.path.join(OUT_DIR, "inventory_tilt.png")
    photo.save(p)
    expected["inventory_tilt.png"] = {"rows": rows, "merges": merges,
                                      "title": "仓库月末盘点表"}

    # 4. 扫描件库存表（ADF 歪斜+黑边+低对比+噪声+JPEG）
    img, rows, merges = draw_table(
        INVENTORY_HEADER, INVENTORY_ROWS,
        col_widths=[90, 160, 220, 200, 90, 110],
        table_title="仓库月末盘点表",
    )
    scan = add_scan_effects(img, seed=23, angle_deg=1.6)
    p = os.path.join(OUT_DIR, "inventory_scan.jpg")
    scan.save(p, format="JPEG", quality=88)
    expected["inventory_scan.jpg"] = {"rows": rows, "merges": merges,
                                      "title": "仓库月末盘点表"}

    # 5. 扫描件合并表头产量表（反向歪斜 + 光源不均更明显）
    img, rows, merges = draw_table(
        header, prod_rows, col_widths=[150, 150, 150, 150, 150],
        merged_header=(first, second), table_title="车间月度产量统计",
    )
    scan = add_scan_effects(img, seed=37, angle_deg=-1.2)
    p = os.path.join(OUT_DIR, "production_scan.jpg")
    scan.save(p, format="JPEG", quality=88)
    expected["production_scan.jpg"] = {"rows": rows, "merges": merges,
                                       "title": "车间月度产量统计"}

    # 6. 低分辨率扫描（200DPI/IM压缩级：长边1100px + 模糊 + JPEG78）
    img, rows, merges = draw_table(
        INVENTORY_HEADER, INVENTORY_ROWS,
        col_widths=[90, 160, 220, 200, 90, 110],
        table_title="仓库月末盘点表",
    )
    scan = add_scan_effects(img, seed=53, angle_deg=1.1, lowres=True)
    p = os.path.join(OUT_DIR, "inventory_lowres.jpg")
    scan.save(p, format="JPEG", quality=80)
    expected["inventory_lowres.jpg"] = {"rows": rows, "merges": merges,
                                        "title": "仓库月末盘点表"}

    # 7. 低分辨率扫描（合并表头表）
    img, rows, merges = draw_table(
        header, prod_rows, col_widths=[150, 150, 150, 150, 150],
        merged_header=(first, second), table_title="车间月度产量统计",
    )
    scan = add_scan_effects(img, seed=67, angle_deg=-0.9, lowres=True)
    p = os.path.join(OUT_DIR, "production_lowres.jpg")
    scan.save(p, format="JPEG", quality=80)
    expected["production_lowres.jpg"] = {"rows": rows, "merges": merges,
                                         "title": "车间月度产量统计"}

    # 10. 无框线表格（模型结构识别路径）
    bl_rows = [
        ["物料编码", "物料名称", "库存数", "存放位置"],
        ["WL-2001", "不锈钢螺母", "3200", "A区-01架"],
        ["WL-2002", "碳钢垫片", "8600", "A区-02架"],
        ["WL-2003", "尼龙衬套", "1450", "B区-11架"],
        ["WL-2004", "铜质接线端子", "2300", "B区-12架"],
        ["WL-2005", "橡胶减震垫", "980", "C区-03架"],
    ]
    img, rows = draw_borderless_table(
        bl_rows, col_xs=[60, 260, 500, 680], row_ys=[110, 170, 230, 290, 350, 410],
        out_title="物料库存一览",
    )
    p = os.path.join(OUT_DIR, "borderless_table.png")
    img.save(p)
    expected["borderless_table.png"] = {"rows": rows, "merges": [],
                                        "title": "物料库存一览", "borderless": True}

    # 11. 90 度横放的照片（整页旋转纠正）
    img, rows, merges = draw_table(
        INVENTORY_HEADER, INVENTORY_ROWS,
        col_widths=[90, 160, 220, 200, 90, 110],
        table_title="仓库月末盘点表",
    )
    rotated = img.rotate(90, expand=True)   # 逆时针 90°
    p = os.path.join(OUT_DIR, "inventory_rot90.png")
    rotated.save(p)
    expected["inventory_rot90.png"] = {"rows": rows, "merges": merges,
                                       "title": "仓库月末盘点表", "rotated": True}

    # 8. 纯文字文档（无表格线，拍照效果）—— 走"文字模式"识别
    img, rows = draw_document(NOTES_TITLE, NOTES_LINES)
    photo = add_photograph_effects(img, seed=29)
    p = os.path.join(OUT_DIR, "notes_photo.png")
    photo.save(p)
    expected["notes_photo.png"] = {"rows": rows, "mode": "text"}

    # 9. 纯文字文档（扫描件低清）—— 文字模式 + 低分辨率
    img, rows = draw_document(NOTES_TITLE, NOTES_LINES)
    scan = add_scan_effects(img, seed=31, angle_deg=1.4, lowres=True)
    p = os.path.join(OUT_DIR, "notes_scan.jpg")
    scan.save(p, format="JPEG", quality=80)
    expected["notes_scan.jpg"] = {"rows": rows, "mode": "text"}

    with open(os.path.join(OUT_DIR, "expected.json"), "w", encoding="utf-8") as f:
        json.dump(expected, f, ensure_ascii=False, indent=1)
    print("generated:", ", ".join(sorted(os.listdir(OUT_DIR))))
    return 0


if __name__ == "__main__":
    sys.exit(main())

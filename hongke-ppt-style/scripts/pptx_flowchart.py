#!/usr/bin/env python3
"""鸿科风格流程图 → PPT 原生形状生成器。

与 flowchart.py 共用同一份 spec.json，但输出为 PowerPoint/WPS 可编辑的
原生形状（圆角矩形框、连接线箭头、文本框），不是 PNG 贴图。
框、箭头、文字均可在 PPT 里直接选中编辑、移动、改色。

用法：
    python pptx_flowchart.py spec.json 输出.pptx

    - 输出文件不存在：新建一份只有该流程图页的 16:9 演示文稿
    - 输出文件已存在：把流程图作为新一页追加到现有演示文稿末尾
      （配合 python-pptx 从零构建的主文档使用）

同时会在输出 .pptx 旁边生成 输出.预览.png（调用 flowchart.py 渲染，
用于自检文字溢出/重叠；确认无误后可删除）。

spec.json 格式见 flowchart.py 顶部注释。坐标系与 matplotlib 版一致：
原点在左下角，y 向上；映射到幻灯片时自动换算为 PPT 的左上原点并
保持宽高比、居中放置。
"""
import json
import os
import subprocess
import sys

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn

# 与 flowchart.py / brand-style.md 一致的鸿科配色
COLORS = {
    "main": "E78F13",   # 主橙
    "deep": "D87718",   # 深橙
    "gold": "B98A2E",   # 浅金（文字用深金）
    "dark": "404040",   # 深灰
    "gray": "595959",   # 中灰（说明文字/箭头）
}

SLIDE_W, SLIDE_H = 12192000, 6858000  # 16:9 EMU（brand-style.md）
MARGIN = 274320                        # 四周留白 ≈ 0.3 英寸


def set_run_font(run, size_pt, color_hex, bold, typeface="黑体"):
    """写入黑体并补 <a:latin>/<a:ea> 节点（brand-style.md 的 OOXML 规范）。"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color_hex)
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        latin = rPr.makeelement(qn("a:latin"), {})
        rPr.insert(0, latin)
    latin.set("typeface", typeface)
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = latin.makeelement(qn("a:ea"), {})
        latin.addnext(ea)
    ea.set("typeface", typeface)


def add_arrowhead(shape):
    """给线条末端加三角箭头（OOXML tailEnd）。"""
    ln = shape.line._get_or_add_ln()
    tail = ln.find(qn("a:tailEnd"))
    if tail is None:
        tail = ln.makeelement(qn("a:tailEnd"), {})
        ln.append(tail)
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")


def build(spec, prs):
    """把 spec 画成原生形状，加到 prs 的新空白页上，返回 slide。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    xlim, ylim = spec["xlim"], spec["ylim"]
    # spec 坐标 → EMU：保持宽高比、居中、四周留白
    avail_w, avail_h = SLIDE_W - 2 * MARGIN, SLIDE_H - 2 * MARGIN
    emu_per_unit = min(avail_w / xlim, avail_h / ylim)
    off_x = (SLIDE_W - emu_per_unit * xlim) / 2
    off_y = (SLIDE_H - emu_per_unit * ylim) / 2

    def X(x):  # spec x → EMU
        return Emu(int(off_x + x * emu_per_unit))

    def Y(y):  # spec y（左下原点）→ EMU（左上原点）
        return Emu(int(off_y + (ylim - y) * emu_per_unit))

    # matplotlib figsize 下每单位对应 72*figsize[0]/xlim 磅，
    # 换算到幻灯片后的字号缩放比
    figsize = spec.get("figsize", [13.5, 6.0])
    mpl_pt_per_unit = 72.0 * figsize[0] / xlim
    pptx_pt_per_unit = emu_per_unit / 12700.0
    fs_scale = pptx_pt_per_unit / mpl_pt_per_unit

    def fs(mpl_size):
        return max(8.0, mpl_size * fs_scale)

    # ---- 框 ----
    for b in spec["boxes"]:
        c = COLORS[b.get("role", "main")]
        w_emu = int(b["w"] * emu_per_unit)
        h_emu = int(b["h"] * emu_per_unit)
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            X(b["x"]), Y(b["y"] + b["h"]), w_emu, h_emu,
        )
        # 圆角半径 ≈ matplotlib rounding_size 1.2 单位
        try:
            box.adjustments[0] = max(0.0, min(0.5, 1.2 / min(b["w"], b["h"])))
        except (IndexError, ValueError):
            pass
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")  # 框内白底
        box.line.color.rgb = RGBColor.from_string(c)
        box.line.width = Pt(1.8)
        box.shadow.inherit = False

        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        for m in ("margin_left", "margin_right"):
            setattr(tf, m, Emu(45720))  # 0.05 英寸
        tf.margin_top = Emu(0)
        tf.margin_bottom = Emu(0)

        # 三行：标题（彩色加粗）/ 说明（中灰）/ 责任人（彩色加粗）
        # 行位置复刻 matplotlib 版：标题距顶 2.6 单位、说明居中、责任人距底 2.4 单位
        h_pt = h_emu / 12700.0
        ppu = pptx_pt_per_unit
        t_fs, s_fs, o_fs = fs(12), fs(9), fs(9)
        t_center, s_center, o_center = 2.6 * ppu, (b["h"] / 2 + 0.3) * ppu, (b["h"] - 2.4) * ppu

        tf.margin_top = Emu(int(max(0, (t_center - 0.6 * t_fs) * 12700)))
        lines = [(b["title"], t_fs, c, True)]
        if b.get("sub"):
            lines.append((b["sub"], s_fs, COLORS["gray"], False))
        if b.get("owner"):
            lines.append((f"责任人：{b['owner']}", o_fs, c, True))
        centers = [t_center] + ([s_center] if b.get("sub") else []) + ([o_center] if b.get("owner") else [])
        sizes = [t_fs] + ([s_fs] if b.get("sub") else []) + ([o_fs] if b.get("owner") else [])
        for i, (text, size, color, bold) in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.CENTER
            if i < len(lines) - 1:
                gap = centers[i + 1] - centers[i] - 0.6 * (sizes[i] + sizes[i + 1])
                para.space_after = Pt(max(0, gap))
            run = para.add_run()
            run.text = text
            set_run_font(run, size, color, bold)

    # ---- 箭头 ----
    for a in spec.get("arrows", []):
        color = COLORS.get(a.get("color", "gray"), COLORS["gray"])
        x1, y1 = X(a["from"][0]), Y(a["from"][1])
        x2, y2 = X(a["to"][0]), Y(a["to"][1])
        curve = a.get("curve", 0)
        if not curve:
            conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
            conn.line.color.rgb = RGBColor.from_string(color)
            conn.line.width = Pt(2.2)
            add_arrowhead(conn)
        else:
            # arc3 弧线：采样成多段折线 freeform，末端加箭头，近似 matplotlib 弧度
            pts = _arc_points(a["from"], a["to"], curve, ylim, off_x, off_y, emu_per_unit)
            fb = slide.shapes.build_freeform(pts[0][0], pts[0][1], scale=1)
            fb.add_line_segments(pts[1:], close=False)
            arc = fb.convert_to_shape()
            arc.fill.background()
            arc.line.color.rgb = RGBColor.from_string(color)
            arc.line.width = Pt(2.2)
            arc.shadow.inherit = False
            add_arrowhead(arc)

    # ---- 备注（平衡公式等，框外纯文字）----
    if spec.get("note"):
        n = spec["note"]
        tb = slide.shapes.add_textbox(X(n["x"]), Y(n["y"]), Emu(1), Emu(1))
        tf = tb.text_frame
        tf.word_wrap = False
        for i, line in enumerate(n["text"].split("\n")):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.CENTER
            run = para.add_run()
            run.text = line
            set_run_font(run, fs(10), COLORS["dark"], True)
        tb.left = Emu(int(X(n["x"]) - tb.width / 2))
        tb.top = Emu(int(Y(n["y"]) - tb.height / 2))

    return slide


def _arc_points(frm, to, rad, ylim, off_x, off_y, emu_per_unit, n=24):
    """matplotlib FancyArrowPatch arc3 的采样点（rad 为弧度偏移）。"""
    import math

    x1, y1 = frm
    x2, y2 = to
    dx, dy = x2 - x1, y2 - y1
    dist = math.hypot(dx, dy)
    # arc3 rad：控制点在垂直方向偏移 rad * dist / 2
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2
    nx, ny = -dy / dist, dx / dist
    cx, cy = mx + nx * rad * dist / 2, my + ny * rad * dist / 2

    def to_emu(px, py):
        return (Emu(int(off_x + px * emu_per_unit)),
                Emu(int(off_y + (ylim - py) * emu_per_unit)))

    # 二次贝塞尔采样（控制点即 arc3 的弧度偏移点）
    pts = []
    for i in range(n + 1):
        t = i / n
        qx = (1 - t) ** 2 * x1 + 2 * (1 - t) * t * cx + t ** 2 * x2
        qy = (1 - t) ** 2 * y1 + 2 * (1 - t) * t * cy + t ** 2 * y2
        pts.append(to_emu(qx, qy))
    return pts


def main():
    spec_path, out_path = sys.argv[1], sys.argv[2]
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)

    if os.path.exists(out_path):
        prs = Presentation(out_path)
        if prs.slide_width != SLIDE_W:
            raise SystemExit(f"{out_path} 不是 16:9（{prs.slide_width}x{prs.slide_height}），"
                             "为避免版式错乱请输出为新文件再合并")
        append = True
    else:
        prs = Presentation()
        prs.slide_width = Emu(SLIDE_W)
        prs.slide_height = Emu(SLIDE_H)
        append = False

    build(spec, prs)
    prs.save(out_path)
    print(f"saved: {out_path}（{'追加一页' if append else '新建单页演示文稿'}）")

    # 预览 PNG：调用 flowchart.py，保证与 matplotlib 版视觉一致。
    # flowchart.py 在 hongke-doc-style 技能里（两技能随同一仓库安装，互为兄弟目录）；
    # 找不到时跳过预览并提示。
    preview = os.path.splitext(out_path)[0] + ".预览.png"
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [os.path.join(here, "flowchart.py"),
                  os.path.join(here, "..", "..", "hongke-doc-style", "scripts", "flowchart.py")]
    flowchart = next((c for c in candidates if os.path.exists(c)), None)
    if flowchart:
        subprocess.run([sys.executable, flowchart, spec_path, preview], check=True)
    else:
        print("warning: 未找到 flowchart.py，跳过预览 PNG 生成", file=sys.stderr)


if __name__ == "__main__":
    main()

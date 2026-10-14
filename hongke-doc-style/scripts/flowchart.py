#!/usr/bin/env python3
"""鸿科风格流程图生成器：白底、框内无填充、彩色边框、每框标注责任人。

用法：
    python3 flowchart.py spec.json 输出.png

spec.json 结构（坐标系自定，xlim/ylim 与画布一致）：
{
  "xlim": 135, "ylim": 60, "figsize": [13.5, 6.0],
  "boxes": [
    {"x": 2, "y": 42, "w": 20, "h": 13,
     "title": "① U8出库", "sub": "仓库按订单发料",
     "owner": "仓库仓管员", "role": "main"}
  ],
  "arrows": [
    {"from": [22.6, 48.5], "to": [27.6, 48.5]},
    {"from": [37, 16], "to": [12, 41.6], "color": "deep", "curve": 0.35}
  ],
  "note": {"x": 125, "y": 10.5, "text": "平衡公式：\\n出库(不含料废)\\n＝入库＋MES报废"}
}

role 可选值：main（主橙，主流程）/ dark（深灰，终审关单）/ gold（浅金，质量判定）/ deep（深橙，报废回退分支）。
箭头 color 缺省为中灰，可选 "deep"；curve 为弧度（arc3 rad），缺省 0（直线）。
"""
import json
import sys

import matplotlib

# 独立运行时 matplotlib 默认字体不含中文，自动探测已安装的中文字体
from matplotlib import font_manager

_CJK = ["Noto Sans CJK SC", "Noto Sans SC", "WenQuanYi Zen Hei", "WenQuanYi Micro Hei", "SimHei"]
_installed = {f.name for f in font_manager.fontManager.ttflist}
for _name in _CJK:
    if _name in _installed:
        matplotlib.rcParams["font.family"] = _name
        break

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

COLORS = {
    "main": "#E78F13",   # 主橙
    "deep": "#D87718",   # 深橙
    "gold": "#B98A2E",   # 浅金（文字用深金）
    "dark": "#404040",   # 深灰
    "gray": "#595959",   # 中灰（说明文字/箭头）
}


def main():
    spec_path, out_path = sys.argv[1], sys.argv[2]
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)

    figsize = spec.get("figsize", [13.5, 6.0])
    fig, ax = plt.subplots(figsize=figsize, dpi=450)
    ax.set_xlim(0, spec["xlim"])
    ax.set_ylim(0, spec["ylim"])
    ax.axis("off")

    for b in spec["boxes"]:
        c = COLORS[b.get("role", "main")]
        ax.add_patch(FancyBboxPatch(
            (b["x"], b["y"]), b["w"], b["h"],
            boxstyle="round,pad=0.4,rounding_size=1.2",
            fc="white", ec=c, lw=1.8, mutation_aspect=1,
        ))
        cx = b["x"] + b["w"] / 2
        ax.text(cx, b["y"] + b["h"] - 2.6, b["title"], ha="center", va="center",
                fontsize=12, color=c, fontweight="bold")
        if b.get("sub"):
            ax.text(cx, b["y"] + b["h"] / 2 - 0.3, b["sub"], ha="center", va="center",
                    fontsize=9, color=COLORS["gray"])
        if b.get("owner"):
            ax.text(cx, b["y"] + 2.4, f"责任人：{b['owner']}", ha="center", va="center",
                    fontsize=9, color=c, fontweight="bold")

    for a in spec.get("arrows", []):
        ax.add_patch(FancyArrowPatch(
            tuple(a["from"]), tuple(a["to"]), arrowstyle="-|>", mutation_scale=20,
            color=COLORS.get(a.get("color", "gray"), a.get("color", "gray")),
            lw=2.2, connectionstyle=f"arc3,rad={a.get('curve', 0)}",
        ))

    if spec.get("note"):
        n = spec["note"]
        ax.text(n["x"], n["y"], n["text"], ha="center", va="center",
                fontsize=10, color=COLORS["dark"], fontweight="bold")

    fig.savefig(out_path, bbox_inches="tight", facecolor="white")
    print(f"saved: {out_path}")


if __name__ == "__main__":
    main()

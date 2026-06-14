# -*- coding: utf-8 -*-
"""gen_json_figs.py — matplotlib 纯黑白「结构化输出示例」配图(JSON / 日志面板截图风)。

**这是一类可选配图,不是每篇专利都需要。** 仅当你的发明确有"结构化输出"(JSON 报文 / 日志 /
配置 / 数据记录)值得作为附图展示判别逻辑时才用;否则跳过。图的数量与类型由技术方案决定。

依赖: pip install matplotlib Pillow ;中文字体 Noto Sans SC(FONT 变量)。

用法: 把要展示的 JSON 文本传给 render(...);示例见 __main__(占位内容,套用前替换为你的真实报文)。
    python gen_json_figs.py
图片输出到 OUT 目录,每张底部烤入"图N …"图注。

关键约定(避免本技能记录过的坑):
- 纯黑白: 朴素加粗标题行 + 细黑边框 + 黑色文本,无彩色栏。
- 不溢出: 长行按 WRAP(显示宽度,CJK 计2)换行,且 WRAP < 框内可用宽度对应的显示宽度;
  判据是 WRAP < (FIG_W - 文本左右内边距)对应的显示宽度。本默认值已验证不溢出。
- 字号: 图片按≈页宽(FIG_W 英寸)生成 → 放进 docx(宽~15.3cm)后近1:1,页面字号≈所设 FS。
- 缺字: 别用 Unicode 下标/希腊字母(₁ Δ),用 ASCII;含中文用 CJK 字体(勿 monospace,否则中文豆腐)。
"""
import os, json, textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

OUT = os.environ.get("FIG_OUT", "patent_figures")
FONT_PATH = os.environ.get("FIG_FONT_PATH", os.path.expanduser("~/.fonts/NotoSansSC.ttf"))
FP = font_manager.FontProperties(fname=FONT_PATH)
os.makedirs(OUT, exist_ok=True)

WRAP = 70      # 换行显示宽度(CJK 计2),≈5.3in < 框内可用~6.08in,确保不溢出
FIG_W = 6.4    # 英寸:≈页宽,放置后近 1:1
FS = 11        # 正文字号(与正文/其它图齐平)
ROW = 0.215    # 英寸/行


def disp_width(s):
    return sum(2 if ord(c) > 0x2E7F else 1 for c in s)


def wrap_line(line, width=WRAP):
    """按显示宽度换行,续行带悬挂缩进。阈值用 width 本身(不要乘2)。"""
    indent = len(line) - len(line.lstrip(" "))
    hang = " " * (indent + 4)
    out, cur, w = [], "", 0
    for ch in line:
        cw = 2 if ord(ch) > 0x2E7F else 1
        if w + cw > width and cur.strip():
            out.append(cur)
            cur, w = hang + ch, disp_width(hang) + cw
        else:
            cur += ch; w += cw
    if cur:
        out.append(cur)
    return out


def render(json_text, title, out_name, caption):
    raw = json_text.splitlines()
    lines = []
    for ln in raw:
        lines.extend(wrap_line(ln, WRAP))
    n = len(lines)
    fig_h = 0.5 + n * ROW + 0.55
    fig, ax = plt.subplots(figsize=(FIG_W, fig_h), dpi=200)
    ax.axis("off"); ax.set_xlim(0, FIG_W); ax.set_ylim(0, fig_h)  # 英寸坐标
    ax.text(0.12, fig_h - 0.27, title, fontproperties=FP, fontsize=FS + 1,
            color="black", weight="bold", va="center", ha="left")
    ax.plot([0.08, FIG_W - 0.08], [fig_h - 0.48, fig_h - 0.48], color="black", lw=0.9)
    # 末行恒落在 y≈0.505(因 fig_h 随行数等比放大),box_bot 取 0.42 给末行留 ~0.08in 底边距,
    # 同时高于图注(y=0.22),避免末行压在框线上(原 0.5 偏紧,末行"}"会贴底框)。
    box_top = fig_h - 0.56; box_bot = 0.42
    ax.add_patch(plt.Rectangle((0.08, box_bot), FIG_W - 0.16, box_top - box_bot,
                               fill=False, ec="black", lw=1.0))
    y = box_top - 0.2
    for ln in lines:
        ax.text(0.24, y, ln, fontproperties=FP, fontsize=FS, color="black", va="center", ha="left")
        y -= ROW
    ax.text(FIG_W / 2, 0.22, caption, fontproperties=FP, fontsize=FS + 1,
            color="black", va="center", ha="center")
    fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01)
    fig.savefig(os.path.join(OUT, out_name), bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)
    print(out_name, "ok, 行数=", n)


if __name__ == "__main__":
    # 占位示例:字段全是通用占位符,仅演示排版。套用时换成你这篇专利真实的输出报文。
    sample = json.dumps({"field_a": "……", "field_b": "……", "level": "……",
                         "detail": {"key1": "……", "key2": "……", "note": "……占位说明……"}},
                        ensure_ascii=False, indent=2)
    render(sample, "〔占位:结构化输出标题〕", "example_json.png",
           "图N  系统输出的结构化结果示例（占位，套用前替换为真实报文）")
    print("DONE ->", OUT)

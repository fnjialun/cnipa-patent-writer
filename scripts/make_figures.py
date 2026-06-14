# -*- coding: utf-8 -*-
"""make_figures.py — 专利配图(白底·纯黑·粗线·大字)。**优先用本文件的高层函数,别手搓固定尺寸的盒子**。

依赖: pip install matplotlib  (并需一个中文字体)
设计目标(逐条对应常见返工):
  ① 纯黑:线 lw≥1.8、色 'black'、白底实底保存——别用细灰线(细线缩放后发灰像"没黑透")。
  ② 不溢出:框**高随换行后文字自适应**、文字按显示宽度换行到框内(留 10% 余量),文字绝不超出框。
  ③ 竖排单列:流程/模块一律**单列自上而下**——满版宽框、字大;线只在框间直走,**不穿框、不靠白底盖线**。
  ④ 不压线:箭头/分组/回流标签一律带白底或置于留白,绝不压在结构线上;**不用竖排单字标签**(难读)。
  ⑤ 图文匹配:画布宽贴近页宽(6.4in),嵌入(~15.3cm)后近 1:1,字号≈所设 fs(~13pt),不会"图大字小"。

用法(高层,推荐):
    import make_figures as mf
    mf.vflow("/tmp/figN.png",
             ["接入与采集", ("特征处理", "对采集数据做……，输出特征"), "判定与输出"],
             down_labels={1: "满足条件"}, title="图N 整体流程示意图")
    mf.vmodules("/tmp/figN.png",
                [("A模块", "职责……"), ("B模块", "职责……"), ("C模块", "职责……")],
                groups=[("某子系统", 0, 1)], feedback=(2, 0, "结果回流"),
                title="图N 系统模块框图")
图号一律用占位「图N」——实际图号/图序由你这篇专利的附图编排决定,勿固定为某图=某类型。
**画完务必单独 Read 这张 PNG,对照 references/figures.md 的自查清单逐条核对(纯黑/不溢出/不压线/字够大),再嵌入。**

低层图元 box/diamond/arrow/ftitle 仍在(供画曲线图/特殊图时拼),但拼图时同样要遵守上面 5 条。
"""
import os
import glob
import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyArrowPatch

K = "black"
FONT = None

_CANDIDATES = [
    os.path.expanduser("~/.fonts/NotoSansSC.ttf"),
    os.path.expanduser("~/.fonts/simsun.ttc"),
    os.path.expanduser("~/.fonts/simhei.ttf"),
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
]


def setup_cjk(font_path=None):
    """注册中文字体并设 rcParams,返回字体名。找不到会抛错(中文必豆腐,必须解决)。"""
    global FONT
    path = font_path
    if path is None:
        for c in _CANDIDATES:
            if os.path.exists(c):
                path = c
                break
    if path is None:
        for c in glob.glob("/usr/share/fonts/**/*CJK*.ttc", recursive=True):
            path = c
            break
    if path is None or not os.path.exists(path):
        raise RuntimeError("未找到中文字体。请装 fonts-noto-cjk 或把宋体/黑体放到 ~/.fonts 后 fc-cache -f，"
                           "或显式传 font_path。")
    fm.fontManager.addfont(path)
    FONT = fm.FontProperties(fname=path).get_name()
    plt.rcParams["font.sans-serif"] = [FONT]
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.facecolor"] = "white"
    return FONT


# ---------- 文本换行(按显示宽度,CJK 计 2) ----------
def disp_w(s):
    return sum(2 if ord(c) > 0x2E7F else 1 for c in str(s))


def wrap_cjk(text, width):
    """按显示宽度换行(显式 \n 也断行),返回行列表。width 为每行显示宽度上限。"""
    out, cur, w = [], "", 0
    for ch in str(text):
        if ch == "\n":
            out.append(cur); cur, w = "", 0; continue
        cw = 2 if ord(ch) > 0x2E7F else 1
        if w + cw > width and cur:
            out.append(cur); cur, w = ch, cw
        else:
            cur += ch; w += cw
    if cur or not out:
        out.append(cur)
    return out


# ---------- 低层图元(拼曲线图/特殊图用;盒子尺寸由调用方保证容得下文字) ----------
def box(ax, x, y, w, h, text, fs=13, lw=1.8):
    ax.add_patch(Rectangle((x, y), w, h, facecolor="white", edgecolor=K, lw=lw, zorder=3))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=K, zorder=4)
    return (x + w / 2, y + h / 2)


def diamond(ax, cx, cy, w, h, text, fs=12):
    ax.add_patch(plt.Polygon([(cx, cy + h / 2), (cx + w / 2, cy), (cx, cy - h / 2), (cx - w / 2, cy)],
                             facecolor="white", edgecolor=K, lw=1.8, zorder=3))
    ax.text(cx, cy, text, ha="center", va="center", fontsize=fs, color=K, zorder=4)


def arrow(ax, p1, p2, text="", fs=11):
    """直箭头;若带文字,文字置中点并加白底(避免压线)。"""
    ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle="-|>", mutation_scale=18, lw=1.8,
                                 color=K, shrinkA=1, shrinkB=1, zorder=2))
    if text:
        ax.text((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2, text, ha="center", va="center",
                fontsize=fs, color=K,
                bbox=dict(boxstyle="square,pad=0.12", fc="white", ec="none"), zorder=5)


def ftitle(ax, x, y, t, fs=14):
    ax.text(x, y, t, ha="center", fontsize=fs, color=K)


# ---------- 高层:单列竖向自适应框图(流程 / 模块,推荐直接用) ----------
def _vstack(out_path, items, title="", fs=13, groups=None, feedback=None, down_labels=None):
    """单列竖向、框高随文字自适应、满版宽、线只在框间直走(不穿框)的框图核心。
    items: 每项为 str 或 (标题, 详述);groups: [(标签, 起idx, 止idx)] 套虚线分组框;
    feedback: (从idx, 到idx, 标签) 左侧留白处画干净回流箭头(标签横排带白底);
    down_labels: {i: '是'} 标在第 i→i+1 箭头旁(带白底)。"""
    if FONT is None:
        setup_cjk()
    groups = groups or []
    down_labels = down_labels or {}
    W = 6.4
    has_fb = feedback is not None
    left = 1.5 if has_fb else 0.2          # 有回流则左侧让出走线 gutter
    right = 0.2
    box_x, box_w = left, W - left - right
    pad = 0.15
    line_h = fs * 1.5 / 72.0               # 英寸/行
    gap = 0.52                             # 框间竖直间距(走箭头)
    inner_w = box_w - 2 * pad
    unit_in = fs / 72.0 / 2.0              # 每显示单位的宽(CJK=2 单位=fs/72 in)
    wrap = max(8, int(inner_w / unit_in * 0.80))   # 留 20% 余量:文字两侧留白、绝不贴边/溢出

    laid = []
    for it in items:
        head, detail = (it if isinstance(it, (tuple, list)) else (it, None))
        hl = wrap_cjk(head, wrap)
        dl = wrap_cjk(detail, wrap) if detail else []   # 详述用同一 wrap,fs-1 更小→留白更多
        h = (len(hl) + len(dl)) * line_h + 2 * pad
        laid.append((hl, dl, h))

    title_h = 0.46 if title else 0.12
    total_h = sum(b[2] for b in laid) + gap * (len(laid) - 1) + title_h + 0.18
    fig, ax = plt.subplots(figsize=(W, total_h), dpi=200)
    ax.set_xlim(0, W); ax.set_ylim(0, total_h); ax.axis("off")

    y = total_h - 0.09
    spans = []
    for (hl, dl, h) in laid:
        top, bot = y, y - h
        ax.add_patch(Rectangle((box_x, bot), box_w, h, facecolor="white", edgecolor=K, lw=1.8, zorder=3))
        ty = top - pad - line_h * 0.5
        for ln in hl:
            ax.text(box_x + box_w / 2, ty, ln, ha="center", va="center", fontsize=fs, color=K, zorder=4)
            ty -= line_h
        for ln in dl:
            ax.text(box_x + box_w / 2, ty, ln, ha="center", va="center", fontsize=fs - 1, color=K, zorder=4)
            ty -= line_h
        spans.append((top, bot))
        y = bot - gap

    cx = box_x + box_w / 2
    for i in range(len(spans) - 1):
        ax.add_patch(FancyArrowPatch((cx, spans[i][1]), (cx, spans[i + 1][0]),
                                     arrowstyle="-|>", mutation_scale=18, lw=1.8, color=K, zorder=2))
        if i in down_labels:
            ax.text(cx + 0.12, (spans[i][1] + spans[i + 1][0]) / 2, down_labels[i],
                    ha="left", va="center", fontsize=fs - 1, color=K,
                    bbox=dict(boxstyle="square,pad=0.1", fc="white", ec="none"), zorder=5)

    for (label, s, e) in groups:
        gt, gb = spans[s][0] + 0.30, spans[e][1] - 0.12
        gx0, gx1 = box_x - 0.14, box_x + box_w + 0.14
        ax.add_patch(Rectangle((gx0, gb), gx1 - gx0, gt - gb, fill=False, ec=K, lw=1.4,
                               ls=(0, (5, 4)), zorder=1))
        ax.text(gx0 + 0.1, gt - 0.03, label, ha="left", va="top", fontsize=fs - 1, color=K,
                bbox=dict(boxstyle="square,pad=0.12", fc="white", ec="none"), zorder=5)

    if has_fb:
        f, t, lab = feedback
        gx = left * 0.40
        yf = (spans[f][0] + spans[f][1]) / 2
        yt = (spans[t][0] + spans[t][1]) / 2
        ax.add_patch(FancyArrowPatch((box_x, yf), (gx, yf), arrowstyle="-", lw=1.7, color=K, zorder=2))
        ax.add_patch(FancyArrowPatch((gx, yf), (gx, yt), arrowstyle="-", lw=1.7, color=K, zorder=2))
        ax.add_patch(FancyArrowPatch((gx, yt), (box_x, yt), arrowstyle="-|>", mutation_scale=18,
                                     lw=1.7, color=K, zorder=2))
        for j, ln in enumerate(wrap_cjk(lab, max(4, int(left / unit_in * 0.8)))):
            ax.text(gx, (yf + yt) / 2 - (j - 0.5) * line_h, ln, ha="center", va="center",
                    fontsize=fs - 1, color=K,
                    bbox=dict(boxstyle="square,pad=0.1", fc="white", ec="none"), zorder=5)

    if title:
        ax.text(W / 2, 0.21, title, ha="center", va="center", fontsize=fs + 1, color=K, zorder=4)
    fig.savefig(out_path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path


def vflow(out_path, steps, title="", down_labels=None, fs=13):
    """竖向流程图(单列、框高自适应文字、满版宽、线不穿框)。steps 每项 str 或 (标题,详述)。
    判定分支别再画到右侧菱形/旁注(那会出竖排小字压线);若要标注分支,用 down_labels={步骤idx:'是'}。"""
    return _vstack(out_path, steps, title=title, fs=fs, down_labels=down_labels)


def vmodules(out_path, mods, title="", groups=None, feedback=None, fs=13):
    """竖向模块框图(单列自适应)。mods 每项 str 或 (模块名,职责);groups=[(标签,起,止)] 套虚线分组;
    feedback=(从idx,到idx,标签) 在左侧留白画干净回流箭头(标签横排白底,不压线、不竖排单字)。"""
    return _vstack(out_path, mods, title=title, fs=fs, groups=groups, feedback=feedback)


def json_examples(out_path, pos_title, pos_text, neg_title, neg_text, title=""):
    """上下堆叠两个 JSON 文本块(黑框自适应包裹文字);含中文必须用 CJK 字体(不要 monospace)。
    更稳的换行不溢出版本见 gen_json_figs.py。"""
    if FONT is None:
        setup_cjk()
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(6.8, 9.6))
    for ax, ttl, txt in [(ax1, pos_title, pos_text), (ax2, neg_title, neg_text)]:
        ax.axis("off")
        ax.set_title(ttl, fontsize=13, color=K, pad=8)
        ax.text(0.5, 0.5, txt, transform=ax.transAxes, fontsize=12, ha="center", va="center",
                ma="left", color=K, fontfamily=FONT, linespacing=1.6,
                bbox=dict(boxstyle="square,pad=0.8", facecolor="white", edgecolor=K, linewidth=1.4))
    if title:
        fig.suptitle(title, fontsize=13.5, color=K, y=0.015)
    fig.subplots_adjust(hspace=0.32, top=0.95, bottom=0.05)
    fig.savefig(out_path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path

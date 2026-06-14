# -*- coding: utf-8 -*-
"""gen_figures_graphviz.py — graphviz 纯黑白专利配图(流程图 / 模块架构框图等)。

matplotlib(make_figures.py)与 graphviz 二选一:**节点多、连线/分支复杂、要自动布局**的图,
graphviz 更省事更整齐;少量盒子的简单竖图用 matplotlib 也行。两者都是白底黑线黑字。

依赖: pip install graphviz ; 并装系统 dot:sudo apt-get install -y graphviz
中文字体: 需安装中文字体(FONT 变量,默认 Noto Sans SC),缺字会渲染成豆腐块□。

理念:**示例仅演示画法,标签全是占位符**。请把节点/边文字换成你这篇专利的真实步骤/模块,
图的数量与类型由技术方案本身决定——不要照搬任何范文的图。

用法:
    import gen_figures_graphviz as gg
    g = gg.new_graph("fig1")
    g.attr(label="图N  整体流程示意图", labelloc="b", fontsize="13")   # 图号由你的附图编排决定,勿固定
    g.node("a", "第一步骤"); g.node("b", "第二步骤"); g.edge("a", "b")
    g.render("patent_figures/figN", cleanup=True)   # -> patent_figures/figN.png
"""
import os
import graphviz

OUT = os.environ.get("FIG_OUT", "patent_figures")
FONT = os.environ.get("FIG_FONT", "Noto Sans SC")
os.makedirs(OUT, exist_ok=True)


def new_graph(name, rankdir="TB"):
    """返回一个统一纯黑白样式的 Digraph(矩形黑边白底、黑箭头、中文字体)。"""
    g = graphviz.Digraph(name, format="png")
    g.attr(rankdir=rankdir, fontname=FONT, fontsize="12", bgcolor="white",
           nodesep="0.3", ranksep="0.45")
    g.attr("node", shape="box", style="", color="black", fontcolor="black",
           fontname=FONT, fontsize="11", penwidth="1.1", margin="0.14,0.07")
    g.attr("edge", color="black", fontcolor="black", fontname=FONT, fontsize="10",
           penwidth="1.0", arrowsize="0.8")
    return g


def cluster_attrs(c, label):
    """给子图(cluster)套统一黑白样式 + 标题。"""
    c.attr(label=label, fontname=FONT, color="black", fontsize="11", style="")
    c.attr("node", shape="box", style="", color="black", fontcolor="black",
           fontname=FONT, fontsize="11", penwidth="1.1", margin="0.14,0.07")


# ============ 通用示例(占位标签,套用前请改成真实内容)============
def example_flow():
    """竖向流程图骨架:线性步骤 + 一个菱形判定 + 分支。仅演示画法。"""
    g = new_graph("example_flow")
    g.attr(label="图N  整体流程示意图", labelloc="b", fontsize="13")
    g.node("s1", "步骤一\n〔占位:输入/获取〕")
    g.node("s2", "步骤二\n〔占位:关键处理〕")
    g.node("d", "〔占位:是否满足条件〕", shape="diamond")
    g.node("s3", "步骤三\n〔占位:进一步处理〕")
    g.node("out", "〔占位:输出/决策〕")
    g.edge("s1", "s2"); g.edge("s2", "d")
    g.edge("d", "s3", label="是"); g.edge("d", "s2", label="否")
    g.edge("s3", "out")
    g.render(os.path.join(OUT, "example_flow"), cleanup=True); print("example_flow ok")


def example_arch():
    """系统架构框图骨架:若干模块 + 一个含子模块的 cluster + 反向回流虚线。仅演示画法。"""
    g = new_graph("example_arch")
    g.attr(label="图N  系统逻辑架构框图", labelloc="b", fontsize="13")
    g.node("m1", "〔占位:第一模块〕")
    g.node("m2", "〔占位:第二模块〕")
    with g.subgraph(name="cluster_sub") as c:
        cluster_attrs(c, "〔占位:某复合模块〕")
        c.node("sa", "〔占位:子模块A〕")
        c.node("sb", "〔占位:子模块B〕")
    g.node("m3", "〔占位:输出模块〕")
    g.edge("m1", "m2")
    g.edge("m2", "sa"); g.edge("m2", "sb")
    g.edge("sa", "m2", style="dashed", constraint="false")  # 回流示意
    g.edge("m2", "m3")
    g.render(os.path.join(OUT, "example_arch"), cleanup=True); print("example_arch ok")


if __name__ == "__main__":
    # 仅演示;实际请按你的专利内容自行构图,并据需要决定画几张、画哪些。
    example_flow(); example_arch()
    print("DONE ->", OUT, "（以上为占位示例,套用前务必替换为真实内容）")

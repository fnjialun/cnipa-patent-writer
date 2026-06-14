# -*- coding: utf-8 -*-
"""render_check.py — 把 docx 渲染成逐页 PNG,供"可视化校验"(结构正确≠视觉正确)。

依赖(一般需 sudo apt):
    libreoffice-writer  (提供 soffice,docx->pdf)
    poppler-utils       (提供 pdftoppm,pdf->png)
    中文字体            (fonts-noto-cjk,或把宋体/黑体放 ~/.fonts 后 fc-cache -f;否则渲染中文豆腐)

用法:
    python render_check.py 成稿.docx [输出目录] [dpi]
然后用 Read 逐页查看输出目录里的 pg-*.png,核对:
    - 四个运行页眉(说明书摘要/权利要求书/说明书/说明书附图)是否分别正确;权要/说明书页码各自从 1 重起;
    - **发明名称(黑体居中标题)逐字无豆腐块**——套模板时模板内嵌的字体子集只含原文字形,本专利的新字会变 □;
      这是最易漏的一处(别只盯配图和正文宋体),务必看清说明书首页那行黑体发明名称的每个字;
    - 正文宋体/两端对齐/首行缩进/行距;章节标题加粗;发明名称居中且仅一次;
    - 配图清晰、居中、不越右边距、中文无豆腐、图号正确;
    - 分页无标题孤行、图被截断。
"""
import os
import sys
import glob
import shutil
import subprocess


def render(docx_path, out_dir=None, dpi=110):
    docx_path = os.path.abspath(docx_path)
    if out_dir is None:
        out_dir = os.path.join(os.path.dirname(docx_path), "_render")
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    if not shutil.which("soffice"):
        raise RuntimeError("缺 soffice,请装 libreoffice-writer。")
    if not shutil.which("pdftoppm"):
        raise RuntimeError("缺 pdftoppm,请装 poppler-utils。")
    # 刷新字体缓存(若用户刚拷入字体)
    if shutil.which("fc-cache"):
        subprocess.run(["fc-cache", "-f"], capture_output=True)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                   check=True, capture_output=True, timeout=180)
    pdfs = glob.glob(os.path.join(out_dir, "*.pdf"))
    if not pdfs:
        raise RuntimeError("soffice 未生成 PDF。")
    pdf = pdfs[0]
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, os.path.join(out_dir, "pg")],
                   check=True, capture_output=True, timeout=180)
    pages = sorted(glob.glob(os.path.join(out_dir, "pg-*.png")))
    print("PDF:", pdf)
    print("页数:", len(pages))
    for p in pages:
        print(" ", p)
    print("\n→ 现在请用 Read 逐页查看上述 PNG,按 docstring 清单核对版式/页眉/配图/分页。")
    return pages


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    render(sys.argv[1],
           sys.argv[2] if len(sys.argv) > 2 else None,
           int(sys.argv[3]) if len(sys.argv) > 3 else 110)

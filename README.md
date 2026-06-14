# cnipa-patent-writer

> A [Claude Code](https://claude.com/claude-code) **Agent Skill** for drafting Chinese invention patents (CNIPA).
> It borrows **only the formatting** of a reference patent and generates all text & figures from **your own**
> materials — never copying the template's technical content.

一个用于撰写**中国发明专利**（CNIPA）申请文档的 Claude Code 技能（Agent Skill）。它**只借一份"模板专利"的版式与
行文风格**（页眉页脚、分节、字体字号、缩进行距、分页、附图编排），而**说明书/权利要求/配图的全部技术内容，一律
据你自己的技术方案重新生成**——绝不照抄模板里的具体技术。产出可直接提交/继续打磨的 `.docx`（含纯黑白配图）。

## 为什么是它：一条红线

套模板写专利最大的坑，是把模板那篇专利的**具体技术内容**（架构、层级、模块、流程、配图）也一起抄了进来——
换个用户就是另一套技术，照抄既写错又涉抄袭。本技能把"**只借格式、绝不抄技术内容**"作为**硬性红线**贯穿始终：
一切文本与配图都据你提供的材料/代码生成，模板只提供"长什么样"的通用版式参照。

## 能力

- 覆盖中国发明专利全部四部分：**说明书摘要 / 权利要求书 / 说明书**（技术领域·背景技术·发明内容·有益效果·
  附图说明·具体实施方式/实施例）**/ 说明书附图**，复刻标准的"四节四运行页眉 + 各部分独立页码"结构。
- **两种装配器**：
  - 有模板 → 克隆模板真实段落作原型，**版式逐项继承**（含逐节复刻模板真实的页脚/页码），并自动剥离模板内嵌的
    字体子集（避免跨主题套模板时新文字渲染成豆腐块）。
  - 无模板 → 按 CNIPA 标准版式从零生成。
- **配图据内容生成**（不是每篇都固定 5 张/JSON 图）：纯黑白、单列竖排、框随文字自适应、字够大；matplotlib 与
  graphviz 两套画法。
- **可视化校验**：`docx → pdf → 逐页 png`，按清单逐页核对版式/页眉/页码/配图（结构对 ≠ 视觉对）。
- 行文范式、权利要求"功能/模式抽象求宽保护"手法、以及一份"真实踩过的坑"反面清单。

## 仓库结构

```
cnipa-patent-writer/
├── SKILL.md                       技能入口：何时用、工作流、核心原则
├── references/
│   ├── writing-style.md           中国专利行文手法、各部分范式、功能抽象、反面清单
│   ├── docx-format.md             克隆复刻法、四节四页眉、页码/内嵌字体处理
│   ├── cnipa-format-spec.md       CNIPA 精确版式参数（字体/字号/页边距/段长…）
│   └── figures.md                 配图六条硬规则 + 强制逐图自查
└── scripts/
    ├── inspect_template.py        扒模板 docx 的真实版式（只读版式，不读内容）
    ├── build_patent.py            有模板装配器（克隆原型 + 四节页眉 + 页码 + 嵌图）
    ├── build_patent_cnipa.py      无模板兜底装配器（区块 JSON → CNIPA 四节）
    ├── make_figures.py            matplotlib 配图（vflow/vmodules 自适应单列）
    ├── gen_figures_graphviz.py    graphviz 纯黑白流程/架构图
    ├── gen_json_figs.py           结构化输出面板图（长行自动换行不溢出）
    └── render_check.py            docx → 逐页 png 渲染校验
```

## 安装

作为 Claude Code 技能，克隆到你的技能目录即可（**目录名须与技能名一致**）：

```bash
git clone https://github.com/fnjialun/cnipa-patent-writer.git \
  ~/.claude/skills/cnipa-patent-writer
```

之后在 Claude Code 里让它写专利时会自动触发；脚本也可单独运行。

## 依赖

**Python**（`pip install -r requirements.txt`）：`python-docx`、`matplotlib`、`Pillow`，以及可选的 `graphviz`。

**系统依赖**（非 pip，按你的系统装）：

```bash
# Debian/Ubuntu
sudo apt-get install -y graphviz libreoffice-writer poppler-utils
# macOS (Homebrew)
brew install graphviz poppler && brew install --cask libreoffice
```

- `graphviz` 的 `dot` 二进制：画 graphviz 图用（光 `pip install graphviz` 不够）。
- `libreoffice`（`soffice`）：`docx → pdf` 渲染校验。
- `poppler-utils`：`pdftoppm`/`pdfinfo`/`pdftotext`。

**中文字体（重要）**：配图与 `docx→pdf` 还原都需要中文字体，否则中文渲染成豆腐块 □。

- **推荐用开源的 [Noto Sans CJK / Noto Sans SC](https://github.com/notofonts/noto-cjk)**（可自由分发），放到 `~/.fonts/`
  后 `fc-cache -f`；脚本会自动发现，也可用环境变量 `FIG_FONT_PATH` / `FIG_FONT` 指定。
- 若想让 `docx→pdf` 与 Word 中的宋体/黑体完全一致，可**自备** `simsun.ttc` / `simhei.ttf` 放进 `~/.fonts/`。
  ⚠️ **本仓库不包含、也请勿提交 SimSun/SimHei**——它们是微软专有字体，分发会侵权。

## 用法（要点）

1. **自备一份"模板专利" `.docx`** 作格式参照（本仓库**不附带**模板）。
2. 把你要写成专利的技术方案（代码库、设计文档或交底材料）交给技能。
3. 技能按工作流执行：扒模板版式 → 据你的材料写正文 → 据内容生成配图 → 装配 docx → 渲染逐页校验。
4. 产物：一篇排版规范的 `.docx`（含配图）。

## 重要提示

- **自备模板**：技能的设计就是套用你提供的模板；不要期待仓库里有现成模板，也不要把含真实技术内容/他人 IP 的
  模板提交到公开仓库。
- **不提交专有字体**（见上）。
- 本仓库的 reference/scripts **不含任何特定专利或项目的技术内容**——只有通用的格式机制与行文方法。

## License

[MIT](./LICENSE) © 2026 fnjialun

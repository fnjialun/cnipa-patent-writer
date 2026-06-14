#!/usr/bin/env node
/*
 * cnipa-patent-writer installer.
 * Copies the skill (SKILL.md + references/ + scripts/ + requirements.txt) into an
 * AI agent's skills folder. Default target = Claude Code user skills (~/.claude/skills).
 *
 * Usage:
 *   npx cnipa-patent-writer                 # install to ~/.claude/skills/cnipa-patent-writer
 *   npx cnipa-patent-writer --project       # install to ./.claude/skills/... (current project)
 *   npx cnipa-patent-writer --dir <path>    # install to a custom skills dir (other agents)
 *   npx cnipa-patent-writer --force         # overwrite an existing install
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");

const SKILL_NAME = "cnipa-patent-writer";
const ITEMS = ["SKILL.md", "references", "scripts", "requirements.txt"]; // what a skill needs
const SRC = path.join(__dirname, "..");

function parseArgs(argv) {
  const a = { force: false, help: false, project: false, dir: null };
  for (let i = 0; i < argv.length; i++) {
    const v = argv[i];
    if (v === "--force" || v === "-f") a.force = true;
    else if (v === "--help" || v === "-h") a.help = true;
    else if (v === "--project" || v === "-p") a.project = true;
    else if (v === "--dir" || v === "-d") a.dir = argv[++i];
    else if (v.startsWith("--dir=")) a.dir = v.slice(6);
  }
  return a;
}

function helpText() {
  return `cnipa-patent-writer — 安装中国发明专利撰写技能到你的 AI agent 技能目录

用法:
  npx cnipa-patent-writer [选项]

选项:
  (默认)            安装到 Claude Code 用户技能目录 ~/.claude/skills/${SKILL_NAME}
  -p, --project    安装到当前项目 ./.claude/skills/${SKILL_NAME}
  -d, --dir <路径> 安装到自定义技能目录（其他 agent 工具用其技能目录路径）
  -f, --force      覆盖已存在的安装
  -h, --help       显示本帮助

示例:
  npx cnipa-patent-writer                      # 用户级（Claude Code）
  npx cnipa-patent-writer --project            # 仅当前项目
  npx cnipa-patent-writer --dir ~/.agent/skills  # 其他工具
`;
}

function resolveBase(a) {
  if (a.dir) return path.resolve(a.dir);
  if (a.project) return path.resolve(process.cwd(), ".claude", "skills");
  return path.join(os.homedir(), ".claude", "skills");
}

function main() {
  const a = parseArgs(process.argv.slice(2));
  if (a.help) {
    process.stdout.write(helpText());
    return;
  }
  const base = resolveBase(a);
  const dest = path.join(base, SKILL_NAME);

  if (fs.existsSync(dest)) {
    if (!a.force) {
      console.error(`✗ 目标已存在：${dest}\n  如需覆盖请加 --force。`);
      process.exit(1);
    }
    fs.rmSync(dest, { recursive: true, force: true });
  }

  fs.mkdirSync(dest, { recursive: true });
  for (const item of ITEMS) {
    const src = path.join(SRC, item);
    if (!fs.existsSync(src)) continue; // requirements.txt 可选
    fs.cpSync(src, path.join(dest, item), { recursive: true });
  }

  console.log(`✓ 已安装技能：${dest}\n`);
  console.log("下一步:");
  console.log("  1) 装 Python 依赖：  pip install python-docx matplotlib Pillow   (复杂图再加 graphviz)");
  console.log("  2) 装系统依赖(渲染校验/字体)：");
  console.log("       Debian/Ubuntu: sudo apt-get install -y graphviz libreoffice-writer poppler-utils");
  console.log("       macOS:         brew install graphviz poppler && brew install --cask libreoffice");
  console.log("     中文字体推荐开源 Noto Sans CJK / Noto Sans SC，放到 ~/.fonts 后 fc-cache -f。");
  console.log("  3) 在你的 agent 里让它写专利即可触发（需自备一份'模板专利'docx 作格式参照）。\n");
  console.log("文档与源码: https://github.com/fnjialun/cnipa-patent-writer");
}

try {
  main();
} catch (e) {
  console.error("✗ 安装失败:", e && e.message ? e.message : e);
  process.exit(1);
}

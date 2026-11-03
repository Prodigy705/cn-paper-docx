---
name: cn-paper-docx-portable
description: 中文论文/报告的 Word（.docx）通用排版，便携版。当用户要新建论文或报告、给已有 .docx 套统一格式、调标题层级字体字号、正文行距缩进、页边距、三线表、图题表注、参考文献格式时使用。适用于 ChatGPT、Grok、Gemini、Claude.ai 以及任何能运行 Python 的 Agent。
---

# 中文论文 Word 通用版式（便携版）

本 skill 的执行入口是单文件脚本 `cn_paper_docx.py`，完整操作流程见
`AI-INSTRUCTIONS.md`，版式数值见 `版式规范.md`，可直接填空的示例稿见
`稿件模板.txt`。

## 快速开始

```bash
# 从标记稿生成完整文档（A4，单栏）
python cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4

# 双栏（其余格式完全不变）
python cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4 --columns 2

# 给已有 docx 套格式（原文件不覆盖）
python cn_paper_docx.py apply 原稿.docx -o 排好版.docx --paper a4

# 核对格式
python cn_paper_docx.py verify 论文.docx --paper a4

# 改参数：导出 JSON → 编辑 → --config 读回来（不用改脚本）
python cn_paper_docx.py dump-config --paper a4 -o 我的版式.json
python cn_paper_docx.py new 稿件.txt -o 论文.docx --config 我的版式.json
```

## 必须遵守

- **只有三级标题**，三级之后直接是正文，不设四级。
- **标题一律不加粗**，靠字体 / 字号 / 缩进区分层级。
- 单栏/双栏用 `--columns 1|2`；双栏只改栏数，其余格式不变。
- 正文宋体小四 12pt、首行缩进 2 字符；A4 版正文固定 28 磅，Letter 版 1.5 倍。
- 中文必须写 `w:eastAsia`；三线表顶线/底线 1.5pt、表头线 0.5pt、无竖线。
- 图题在图下方，表题在表格上方；参考文献不编造。
- 目录是域，需在 Word 里按 F9 更新页码。
- `apply` 不改文字、不删图、不动批注；原文件永不覆盖。

## 环境

依赖 `python-docx`（必需）、`Pillow`（可选）。没有 Word 也能生成 `.docx`，
只是目录页码要用户按 F9。完整说明见 `AI-INSTRUCTIONS.md` 与 `README.md`。

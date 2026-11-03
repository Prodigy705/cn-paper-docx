# cn-paper-docx

中文论文 / 报告 Word（`.docx`）通用排版 skill。

按「标题层级 + 段落角色」组织，不绑定任何具体章节安排——论文有几章、章节叫什么
名字都不影响套用。层级只有三级，三级之后直接是正文。

## 特性

- **两套默认预设**：`letter`（Letter 纸，1.5 倍行距）和 `a4`（A4 纸，固定 28 磅）。
- **单栏 / 双栏**：`--columns 1|2`，只改栏数，其余字体、字号、行距、页边距不变。
- **JSON 参数配置**：`dump-config` 导出当前预设，改完用 `--config` 读回来，
  不用改代码。
- **完整排版能力**：三级标题、正文首行缩进、三线表、图题表注、参考文献、
  封面、目录域、页脚页码。
- **可自检**：`verify` 在没有 Word 的环境里核对字体字号、行距、页边距、
  三线表、栏数，适合 AI 代码解释器。
- **给其他 AI 用**：`portable/` 里是单文件工具 + 系统提示 + 规范，可上传到
  ChatGPT、Grok、Gemini、Claude.ai 等能运行 Python 的 AI。

## 快速开始

### Claude Code skill（主 skill）

```bash
# 从零生成完整文档
python scripts/new_paper.py 初稿.txt -o 论文.docx --paper a4

# 双栏（其余格式不变）
python scripts/new_paper.py 初稿.txt -o 论文.docx --paper a4 --columns 2

# 给已有 docx 套格式（原文件不覆盖）
python scripts/apply_format.py 原稿.docx -o 排好版.docx --paper a4

# 生成空白模板 / 样张
python scripts/make_template.py --paper a4 --out 模板.docx

# 核对
python portable/cn_paper_docx.py verify 论文.docx --paper a4
```

### 便携版（ChatGPT / Grok / Gemini / Claude.ai）

```bash
python portable/cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4
python portable/cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4 --columns 2
python portable/cn_paper_docx.py apply 原稿.docx -o 排好版.docx --paper a4
python portable/cn_paper_docx.py verify 论文.docx --paper a4
```

把 `portable/AI-INSTRUCTIONS.md` 粘进 AI 的 Instructions，上传
`portable/cn_paper_docx.py` 和 `portable/版式规范.md` 即可。也可以直接上传
`portable/cn-paper-docx-portable.zip`。各平台步骤见
[`portable/README.md`](portable/README.md)。

### 自定义参数

```bash
python portable/cn_paper_docx.py dump-config --paper a4 -o 我的版式.json
# 编辑 我的版式.json：只写要改的项，其余继承 "base": "a4"
python portable/cn_paper_docx.py new 稿件.txt -o 论文.docx --config 我的版式.json
python portable/cn_paper_docx.py verify 论文.docx --config 我的版式.json
```

可改页面（纸张、页边距、页眉页脚距离）、三级标题字体字号、正文/题注/表注/
参考文献/步骤行/引文/表格单元格/封面的字体字号行距缩进、三线表线宽、
`columns`（1/2）和 `column_gap_cm`。完整字段见
[`references/format-spec.md`](references/format-spec.md) 第 10 节。

## 目录结构

```
.
├─ SKILL.md                  # Claude Code skill 入口
├─ README.md                 # 本文件
├─ references/
│  └─ format-spec.md         # 完整版式规范
├─ scripts/
│  ├─ paper_style.py         # 版式核心：字体/行距/页面/样式/三线表
│  ├─ new_paper.py           # 从标记稿生成完整文档
│  ├─ apply_format.py        # 给已有 docx 套格式
│  ├─ make_template.py       # 生成空白模板与样张
│  └─ config_tool.py         # 导出/查看 JSON 版式配置
├─ assets/                   # 模板与样张 docx
└─ portable/                 # ChatGPT / Grok 等 AI 的便携版
   ├─ AI-INSTRUCTIONS.md
   ├─ cn_paper_docx.py
   ├─ 版式规范.md
   ├─ 稿件模板.txt
   ├─ 示例版式.json
   ├─ 示例版式-双栏.json
   └─ cn-paper-docx-portable.zip
```

## 依赖

- Python 3.9+
- `python-docx`（必需）
- `Pillow`（可选，用于读图片像素尺寸；没有也能插图，只是不自动缩放）

```bash
pip install python-docx pillow
```

## 已知边界

- 只有三级标题，四级编号会降为正文。
- 目录是 Word 域，需要在 Word 中按 F9 更新页码；Windows + Word 环境下
  `new` 会自动刷新一次。
- 支持单栏/双栏；三栏以上、不等宽栏不支持。
- `apply` 不改文字、不删图、不动批注；引文段落需要手动套 `Quote-CN` 样式。
- 字体名写进 docx，用户机器需要安装宋体 / 黑体 / 楷体。
- 不支持旧版 `.doc`、公文格式、英文论文的字体体系。

## 共同创作者

- 项目作者：Prodigy705
- 共同创作者：DeepSeek

开发协助：Claude Code（Anthropic）。

## 许可证

本项目暂未指定开源许可证。使用、分发或修改前请先与作者确认。

---
name: cn-paper-docx
description: 中文论文/报告的 Word（.docx）通用排版规范。当用户要新建论文或报告、要给已有 .docx 套统一格式、要调标题层级字体字号、正文行距缩进、页边距、三线表、图题表注、参考文献格式时使用。Triggers 包括：写论文、写报告、Word 排版、docx 格式、按论文格式、套用格式、排版规范、一级标题、二级标题、三级标题、三线表、图注表注、参考文献格式、字体字号行距、页边距。
---

# 中文论文 Word 通用版式

一套可直接套用的中文论文排版规则，**按「标题层级 + 段落角色」组织，不绑定
任何具体章节安排** —— 论文有几章、章节叫什么名字都不影响套用。

**层级只有三级，三级之后直接是正文，不设四级标题。**

## 两套版式预设（按纸张分，各对应一份参考稿）

| | `letter`（默认） | `a4` |
|---|---|---|
| 参考稿 | `总.docx` | `第四章　实验流程.docx` |
| 一级标题 | 黑体三号 16pt 不加粗 | 黑体三号 16pt 不加粗 |
| 二级标题 | 黑体四号 14pt 不加粗 | **楷体三号 16pt 不加粗 + 首行缩进2字符** |
| 三级标题 | 黑体小四 12pt 不加粗 | 黑体小四 12pt 不加粗 |
| 正文 | 宋体小四 12pt，首行缩进2字符 | 同左 |
| 行距 | 1.5 倍 | 固定 28 磅 |
| 纸张 | Letter，上下3.5cm 左右3.0cm | A4，上下3.5cm 左右2.8cm |
| 图题/表题 | 黑体小四 12pt 居中 | 同左 |

用 `--paper letter` / `--paper a4` 选择。

**动手前先读 `references/format-spec.md`** —— 段前段后、缩进、三线表线宽等
全部数值都在里面。不要凭记忆写格式参数。

## 三条使用路径

### A. 从零写完一篇完整文档（推荐）

用简化标记写稿，脚本负责排版，产出**封面 + 目录 + 正文 + 参考文献**的完整文档。
完整语法见 `scripts/new_paper.py` 文件头。

```bash
SKILL=~/.claude/skills/cn-paper-docx
python "$SKILL/scripts/new_paper.py" 初稿.txt -o 论文.docx --paper a4
python "$SKILL/scripts/new_paper.py" --demo -o 示例.docx      # 先看演示稿
```

```
@cover                     # 封面页（下面跟 `键: 值` 行）
    标题: 论文题目
    副标题: ——副标题
    作者: 张三
    单位: ×××大学 ×××学院
    日期: 2026年9月

@toc                       # 自动目录（Word 里按 F9 更新即生成页码）
@pagenum                   # 从此处另起一节并从 1 编页码
                           #   封面/目录不编页码，正文从 1 开始
                           #   @pagenum continue = 接着上一节编

# 一、研究背景              # 一级标题
## （一）子标题             # 二级标题
### 1.三级标题              # 三级标题
#! 参考文献                # 同上但居中（# 后加 ! 即居中，不按章节名猜）
- Step 1：xxx              # 步骤行
[图3-1 说明](图.png)        # 插图 + 图题（居中，超宽自动等比缩小）
[图3-1 说明]                # 只放图题，不插图
[表3-1 说明](表.png)        # 表题 + 表格图片（表题在上）
| 模型 | 准确率 |            # 三线表，首行是表头
> 注：xxx                   # 表注
[@] 作者. 题名[J]. 刊, 年.   # 参考文献条目
```

图片路径可相对稿件文件所在目录，也可绝对路径；找不到会跳过并在 stderr 提示，
不会中断生成。

**用了 `####` 会直接报错** —— 本规范只有三级标题，四级请写正文。

空行分段，段内可以硬换行（按 Markdown 软换行合并成一段）；
正文里 `**加粗**` 生效。

> **目录只刷新一次，不会每次打开都更新。** 脚本生成时已自动调用 Word
> 刷新一次目录页码并存回文件，所以打开即是带页码的正确目录。
>
> 实现上关键的一点：目录域**不能**带 `w:dirty` 属性 —— 带了的话 Word 每次
> 打开文档都会强制刷新，用户手动更新再保存也白搭，下次打开又变回去。
> 想跳过自动刷新用 `--no-toc-refresh`；那就需要手动按一次 `F9`。

### B. 给已有文档套格式

```bash
python "$SKILL/scripts/apply_format.py" 原稿.docx -o 排好版.docx --paper a4
python "$SKILL/scripts/apply_format.py" 原稿.docx --dry-run   # 先看会改什么
python "$SKILL/scripts/apply_format.py" 原稿.docx -o 出.docx --center-titles "参考文献,致谢"
```

会做：页面设置、重定义样式、按编号特征识别标题层级并套样式、识别图题表注
步骤行参考文献、补 `w:eastAsia` 字体、表格转三线表、图片居中。
不会做：改文字、删图片、动批注。

**先扫全文判断用哪套编号体系，再逐段分类**（段内的 `1.` 在两套里含义不同）：

| 体系 | 编号 | `1.` 是 |
|---|---|---|
| 中文式 | `一、` `（一）` `1.` | 三级 |
| 数字式 | `1.` `1.1` `1.1.1` | 一级 |

判定方法：全文只要出现 `1.1` 这种两段式编号就按数字式处理。
另外 `第一章` `第3章` `壹、` `图3-1` `图 3 xxx` `表2-1` `Figure 3` `Step 1：`
`第一步：` `1：` 都能识别。「（1）」这类四级编号会降成正文并报告出来。

**原文件永不覆盖**，输出到新文件。标题默认两端对齐，要居中哪些标题用
`--center-titles` 明确指定。

### C. 只要空模板

```bash
python "$SKILL/scripts/make_template.py"                    # 两套版式都生成
python "$SKILL/scripts/make_template.py" --paper a4         # 只生成 A4 版
```

产出 `assets/template-{letter,a4}.docx`（空白模板，Word 打开直接写）和
`assets/specimen-{letter,a4}.docx`（样式样张，每种段落配一行说明，用来肉眼核对）。

### C2. 单栏 / 双栏

三个脚本都支持 `--columns`，其余格式完全沿用所选预设：

```bash
python "$SKILL/scripts/new_paper.py" 初稿.txt -o 论文.docx --paper a4 --columns 2
python "$SKILL/scripts/apply_format.py" 原稿.docx -o 排好版.docx --paper a4 --columns 2
python "$SKILL/scripts/make_template.py" --paper a4 --columns 2 --out 模板.docx
```

- `--columns 1`（默认）单栏；`--columns 2` 双栏。
- `--column-gap 0.75` 设置双栏栏间距（cm，默认 0.75）。
- 双栏只改栏数：字体、字号、行距、页边距、三线表等都不变。
- `new` 遇到 `@pagenum` 时，封面/目录所在节保持单栏，正文新节双栏；没有
  `@pagenum` 时全文双栏。
- 双栏时插图按单栏宽度等比缩放，表格收进单栏宽度。
- `apply --columns 2` 会把原文所有节设为双栏（适合没有封面/目录分节的稿子）。
- `verify 论文.docx --paper a4 --columns 2` 核对最后一节的栏数。

### D. 给其他 AI 用（ChatGPT / Grok / Gemini / Claude.ai）

`portable/` 下是同一套规范的**便携版**：一段系统提示 + 一个单文件脚本 + 完整
规范，任何能跑 Python 的 AI 都能用：

| 文件 | 用途 |
|---|---|
| `portable/SKILL.md` | 标准 Agent Skills 入口，支持该格式的工具可直接用 |
| `portable/AI-INSTRUCTIONS.md` | 粘进 Project / Gem / 自定义指令的完整系统提示 |
| `portable/AI-INSTRUCTIONS-compact.md` | 字符受限时用的精简版 |
| `portable/cn_paper_docx.py` | 单文件工具，子命令 `new` / `apply` / `template` / `verify` / `spec` |
| `portable/版式规范.md` | 知识库文件，完整数值 |
| `portable/稿件模板.txt` | 可直接填空的标记稿 |
| `portable/示例版式.json` | 自定义参数示例，改完直接 `--config` |
| `portable/cn-paper-docx-portable.zip` | 一键上传包 |

给 AI 的典型命令：

```bash
python cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4
python cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4 --columns 2
python cn_paper_docx.py apply 原稿.docx -o 排好版.docx --paper a4
python cn_paper_docx.py verify 论文.docx --paper a4 --columns 2
```

`verify` 会在没有 Word 的环境里核对纸张页边距、样式字体字号行距、三线表、
四级标题等，给 AI 一个可自检的闭环。

便携版的脚本由主 skill 的 `scripts/` 自动合并生成。**改完版式后**在仓库根目录
运行 `python portable/build_portable.py`，重新生成脚本、规范和 zip —— 不要直接
改 `portable/cn_paper_docx.py`。各平台的部署步骤见 `portable/README.md`。

## 自定义版式参数（JSON 配置）

两套内置预设是默认值；用户要改字体、字号、行距、页边距、三线表线宽等，**不必改
代码**，导出一份 JSON 改完再用 `--config` 读回来即可。

```bash
# 1. 导出当前预设（a4 或 letter）
python "$SKILL/scripts/config_tool.py" dump --paper a4 -o 我的版式.json

# 2. 编辑 我的版式.json（只写要改的项，其余继承 "base"）

# 3. 用配置生成 / 重排 / 做模板
python "$SKILL/scripts/new_paper.py" 稿件.txt -o 论文.docx --config 我的版式.json
python "$SKILL/scripts/apply_format.py" 原稿.docx -o 排好版.docx --config 我的版式.json
python "$SKILL/scripts/make_template.py" --config 我的版式.json --out 模板.docx

# 4. 查看合并后的完整数值
python "$SKILL/scripts/config_tool.py" show 我的版式.json
```

最小配置示例（只改正文和页边距，其余继承 a4）：

```json
{
  "name": "my-school",
  "base": "a4",
  "label": "我的学校版式",
  "page": { "left_cm": 3.0, "right_cm": 3.0 },
  "body_line": "1.5",
  "roles": {
    "body": { "font": "仿宋", "size_pt": 14, "line": "1.5" },
    "caption": { "font": "楷体", "size_pt": 14 }
  },
  "table_lines": { "top": 2.0, "header": 0.75, "bottom": 2.0 }
}
```

可配字段：

| 段 | 字段 |
|---|---|
| `page` | `paper`（a4/letter）、`width_cm`、`height_cm`、`top_cm`、`bottom_cm`、`left_cm`、`right_cm`、`header_cm`、`footer_cm` |
| `headings` | `1`/`2`/`3`：`font`、`size_pt`、`bold`、`before_pt`、`after_pt`、`line` |
| 顶层 | `body_line`、`h2_indent_chars`、`columns`（1/2）、`column_gap_cm` |
| `roles` | `body`、`title`、`caption`、`note`、`reference`、`step`、`quote`、`table_cell`、`cover_title`、`cover_subtitle`、`cover_text` |
| `table_lines` | `top`、`header`、`bottom`（磅） |

角色字段：`font`、`ascii_font`、`size_pt`、`line`、`indent_chars`、`align`
（left/center/right/justify）、`before_pt`、`after_pt`；`step` 另有 `num_font`，
`quote` 另有 `left_cm` / `right_cm`，`table_cell` 另有 `bold_header`。

便携版对应命令：`cn_paper_docx.py dump-config`、`--config 我的版式.json`、
`cn_paper_docx.py spec --config 我的版式.json`。

## 工作目录约定（重要）

写文档时**不要把中间产物散落在工作目录里**。约定：

```
<工作目录>/
├─ ××××.docx          ← 成品（中文名）
├─ ××××稿.txt          ← 稿件源文件（中文名，和成品同名相近）
├─ 写作过程/            ← 所有中间产物都进这里
│   ├─ 网页抓取/        ← 抓来的 js / html / json
│   └─ 版式核对/        ← 转 PDF 后的核对面图（pdf / png）
└─ pkg/ 等源材料        ← 赛题包、参考资料，保持原样
```

- **不要在用户的文件夹里留脚本**。写一次性脚本请放到系统临时目录
  （`$TMPDIR` / `%TEMP%`）执行完就删，或者直接用 `python -c "..."` 内联。
- **中间产物一律用中文名**，放进 `写作过程/` 子目录，不要摊在根目录。
- **版式核对**（Word 转 PDF、再转 PNG 目视检查）产生的文件用完放进
  `写作过程/版式核对/`，核对完可以删。
- 成品和稿件源文件留在工作目录根下，用中文名。

## 硬约束（别踩）

- **中文字体必须写 `w:eastAsia`**。python-docx 的 `run.font.name` 只管西文，
  不改这个的话 Word 里中文会回退成「等线」。统一用 `paper_style.set_run_font()`。
- **首行缩进要同时写 `firstLineChars` 和 `firstLine`**，用
  `paper_style.set_first_line_indent_chars()`，别用 `paragraph_format.first_line_indent`。
- **只有三级标题，没有四级**。用户明确要求过。
- **标题一律不加粗**，靠换字体/降字号区分层级 —— 黑体再加粗会发糊。
- **多倍行距要关 `snapToGrid`**，否则 1.5 倍行距被文档网格拉歪。
- **表格用 `paper_style.make_three_line_table()`**，别手工画边框。
- **日常改参数用 JSON 配置 + `--config`**（见上文），不要改代码；只有要新增
  一套内置预设时才改 `paper_style.PROFILES` 那一张表，三个脚本共用。
- 标题行距与正文行距**互相独立**，可用 `--line-heading` / `--line-body`
  分别覆盖（`1.5` = 1.5 倍，`28pt` = 固定值 28 磅）。

## 改完格式要验证

数值对不对不能靠猜。生成 docx 后用 Word 转 PDF 再看一眼：

```bash
powershell -NoProfile -Command "
\$w = New-Object -ComObject Word.Application; \$w.Visible=\$false; \$w.DisplayAlerts=0
\$d = \$w.Documents.Open('C:\path\out.docx',\$false,\$true)
\$d.SaveAs([ref]'C:\path\out.pdf',[ref]17); \$d.Close(\$false); \$w.Quit()"
python -c "import pymupdf; d=pymupdf.open('C:/path/out.pdf'); d[0].get_pixmap(dpi=110).save('C:/path/p1.png')"
```

看 `p1.png` 核对：标题是否黑体/楷体且不加粗、正文是否首行缩进 2 字符、
行距是否符合所选预设、三线表是否只有三条线、图表题注是否黑体小四居中、
**且没有出现第四级标题**。

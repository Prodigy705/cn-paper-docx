# cn-paper-docx 便携版：给 ChatGPT / Grok / Gemini / Claude.ai 用

把原来的 Claude Code skill（中文论文 Word 通用排版）扩展成**任何有 Python 的
AI 都能用**的形式：一段系统提示 + 一个单文件脚本 + 一份规范。

```
portable/
├─ README.md                     ← 本文件
├─ SKILL.md                      ← 标准 Agent Skills 入口（支持该格式的工具用）
├─ AI-INSTRUCTIONS.md            ← 完整系统提示（粘进 Project / Gem / 自定义指令）
├─ AI-INSTRUCTIONS-compact.md    ← 精简版（字符受限时用）
├─ 版式规范.md                   ← 知识库文件，完整数值
├─ cn_paper_docx.py              ← 单文件工具（唯一需要执行的脚本）
├─ 稿件模板.txt                  ← 可直接填空的标记稿
├─ 示例版式.json                 ← 自定义参数示例，可改完直接 --config
├─ 示例版式-双栏.json            ← 双栏配置示例（其余同 A4）
├─ cli_append.py                 ← 构建用，不发给 AI
├─ build_portable.py             ← 构建用，不发给 AI
└─ cn-paper-docx-portable.zip    ← 一键上传包（README + SKILL + 提示 + 规范 + 脚本 + 模板 + 示例配置）
```

脚本依赖：`python-docx` 必需，`Pillow` 可选。没有 Word 也能生成 `.docx`；
只是目录页码要用户在 Word 里按一次 F9 更新。

---

## 一、ChatGPT

### 方式 1：Project（推荐）

1. 新建一个 Project（或自定义 GPT）。
2. 把 `AI-INSTRUCTIONS.md` 粘进 Project instructions / GPT Instructions。
   如果提示超长，改用 `AI-INSTRUCTIONS-compact.md`。
3. 把 `cn_paper_docx.py`、`版式规范.md`、`稿件模板.txt` 上传到 Project files
   （或 GPT Knowledge）。
4. 在对话里确认代码解释器可用（Advanced Data Analysis）。
5. 如果代码解释器里看不到知识库文件，**直接把 `cn_paper_docx.py` 拖进对话**；
   它会出现在 `/mnt/data/`。也可以拖整个 `cn-paper-docx-portable.zip`，让 AI
   先解压。

### 方式 2：普通对话

把 `AI-INSTRUCTIONS-compact.md` 作为第一条消息发过去，再上传
`cn_paper_docx.py` 和你的稿件 / 图片，然后说：

> 按这套规范把内容排成 A4 论文，输出 .docx。

### 第一句话示例

> 帮我写一篇关于「基于 EEG 的运动想象分类」的课程论文，A4，需要封面、目录、
> 页码。参考文献我稍后补，先留占位。

ChatGPT 会写标记稿、运行脚本、生成 `/mnt/data/论文.docx`，并跑 `verify`。

---

## 二、Grok

1. 把 `AI-INSTRUCTIONS-compact.md` 粘进 Grok 的自定义指令；或在对话第一条
   消息里整段发过去。
2. 上传 `cn_paper_docx.py`（必要时连同 `版式规范.md`）。
3. 在支持代码执行 / Python 沙箱的模式下使用；如果当前模式不能执行代码，
   Grok 会按提示进入「无代码模式」，给你 `稿件.txt` + 本地运行命令。
4. 提示词示例：

> 用我上传的 cn_paper_docx.py，把下面内容排成 A4 论文并给我 .docx：
> ……

---

## 三、Gemini / Claude.ai / 其他

- **Gemini Gem**：把 `AI-INSTRUCTIONS.md` 放进 Gem instructions，上传脚本与
  规范作为知识文件，开启代码执行。
- **Claude.ai Project**：同上，在 Project instructions 里粘贴提示，上传脚本，
  使用分析工具。
- **本地 Claude Code / Codex CLI**：直接把整个仓库当项目，或把便携版目录
  放在当前工作目录。
- 任何能读写文件、能跑 Python 的 Agent 都适用：让它读 `AI-INSTRUCTIONS.md`，
  执行 `cn_paper_docx.py` 即可。

---

## 四、本地使用（没有 AI 也能用）

```bash
pip install python-docx pillow

python cn_paper_docx.py new 初稿.txt -o 论文.docx --paper a4
python cn_paper_docx.py new 初稿.txt -o 论文.docx --paper a4 --columns 2
python cn_paper_docx.py new --demo -o 示例.docx --paper a4
python cn_paper_docx.py apply 原稿.docx -o 排好版.docx --paper a4
python cn_paper_docx.py apply 原稿.docx --dry-run
python cn_paper_docx.py template -o 模板.docx --paper a4
python cn_paper_docx.py verify 论文.docx --paper a4
python cn_paper_docx.py spec
```

Windows + 安装了 Word 时，`new` 会自动刷新一次目录页码；其他环境会跳过，
在 Word 里按 F9 即可。

---

## 五、版式速查

| | `letter`（默认） | `a4` |
|---|---|---|
| 纸张 | Letter，上下 3.5cm 左右 3.0cm | A4，上下 3.5cm 左右 2.8cm |
| 一级标题 | 黑体三号 16pt 不加粗 | 黑体三号 16pt 不加粗 |
| 二级标题 | 黑体四号 14pt 不加粗 | 楷体三号 16pt 不加粗 + 缩进 2 字符 |
| 三级标题 | 黑体小四 12pt 不加粗 | 黑体小四 12pt 不加粗 |
| 正文 | 宋体小四 12pt，首行缩进 2 字符 | 同左 |
| 行距 | 1.5 倍 | 固定 28 磅 |
| 图题 / 表题 | 黑体小四 12pt 居中 | 同左 |

**单栏 / 双栏**：`--columns 1`（默认）/ `--columns 2`，`--column-gap 0.75`
设栏间距（cm）。双栏只改栏数，其余格式完全沿用所选预设；有 `@pagenum` 时
封面/目录保持单栏、正文从新节开始双栏。`verify` 时带上相同的 `--columns`。

完整数值（段前段后、三线表线宽、封面、页码、参考文献等）见 `版式规范.md`。

---

## 六、自定义版式参数（推荐）

不用改脚本：导出一份 JSON，按需求改，再用 `--config` 读回来。

```bash
# 导出 A4 预设（或 letter）
python cn_paper_docx.py dump-config --paper a4 -o 我的版式.json

# 编辑 我的版式.json；只写要改的项，其余继承 "base": "a4"

# 用配置生成 / 重排 / 核对
python cn_paper_docx.py new 稿件.txt -o 论文.docx --config 我的版式.json
python cn_paper_docx.py apply 原稿.docx -o 排好版.docx --config 我的版式.json
python cn_paper_docx.py verify 论文.docx --config 我的版式.json
python cn_paper_docx.py spec --config 我的版式.json
```

最小配置示例：

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

分栏也能写进配置：`"columns": 2, "column_gap_cm": 0.75`。开箱示例见
`示例版式-双栏.json`。

可改：纸张与页边距、单栏/双栏、三级标题的字体/字号/段前段后/行距、正文/题注/
表注/参考文献/步骤行/引文/表格单元格/封面的字体字号行距缩进、三线表线宽。
完整字段表见 `版式规范.md` 第 10 节或 `AI-INSTRUCTIONS.md`。

---

## 七、已知边界

- **只有三级标题**，四级编号会降为正文（这是本规范的硬约束）。
- **目录是域**，不是死文本；没有 Word 的环境不会自动算页码，按 F9 更新。
- **字体名写进 docx**，用户机器需要装宋体 / 黑体 / 楷体；macOS 没装时 Word 会
  替换字体，版式可能略有差异。
- **`apply` 不改文字、不删图、不动批注**；引文段落不会自动识别，需要手动套
  `Quote-CN` 样式。
- 支持单栏/双栏（`--columns 1|2`）；三栏以上、不等宽栏不支持。不支持旧版
  `.doc`、公文格式、英文论文的字体体系；需要四级以上标题的场合不适用。
- 图片找不到会跳过并在 stderr 提示，不会中断生成。

---

## 八、重新生成便携版

`cn_paper_docx.py` 由主 skill 的 `scripts/` 自动合并生成，`版式规范.md` 由
`references/format-spec.md` 生成。改完主 skill 后，在仓库根目录运行：

```bash
python portable/build_portable.py
```

它会重新生成 `cn_paper_docx.py`、`版式规范.md` 和 `cn-paper-docx-portable.zip`。
**不要直接改 `cn_paper_docx.py`**，下次构建会被覆盖。

## 九、许可证

MIT License。见仓库根目录的 [`LICENSE`](../LICENSE)，zip 包内也带了一份。

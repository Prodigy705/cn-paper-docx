# -*- coding: utf-8 -*-
"""生成便携版文件（供 ChatGPT / Grok / Gemini / Claude.ai 等使用）。

用法（在仓库根目录或 portable/ 下都行）:
    python portable/build_portable.py

做的事:
    1. 把 scripts/paper_style.py、new_paper.py、apply_format.py、make_template.py
       合并成一个自包含的 portable/cn_paper_docx.py（去掉各自的 main 与 import）。
    2. 把 references/format-spec.md 复制成 portable/版式规范.md，并把脚本名
       换成便携版的子命令。
    3. 打一个 portable/cn-paper-docx-portable.zip，方便一次上传给 AI。

改版式请改 scripts/paper_style.py 的 PROFILES，然后重新运行本脚本；
不要直接改生成出来的 cn_paper_docx.py。
"""
import ast
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "scripts")


def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def write(path, text):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def strip_module(src):
    """去掉模块 docstring、顶层 import、sys.path.insert、main() 与 __main__ 守卫，
    并把 `ps.` 前缀去掉（合并后各函数就在同一命名空间）。"""
    tree = ast.parse(src)
    lines = src.splitlines()
    drop = set()

    # 模块 docstring
    if (tree.body and isinstance(tree.body[0], ast.Expr)
            and isinstance(tree.body[0].value, ast.Constant)
            and isinstance(tree.body[0].value.value, str)):
        drop.update(range(tree.body[0].lineno, tree.body[0].end_lineno + 1))

    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == "main":
            drop.update(range(node.lineno, node.end_lineno + 1))
        elif isinstance(node, ast.If):
            seg = ast.get_source_segment(src, node) or ""
            if "__main__" in seg:
                drop.update(range(node.lineno, node.end_lineno + 1))

    kept = []
    for i, line in enumerate(lines, 1):
        if i in drop:
            continue
        if re.match(r"^(from\s|import\s)", line):
            continue
        if re.match(r"^sys\.path\.insert", line):
            continue
        # make_template.py 里只有原 main() 用到的路径常量，合并后不需要
        if re.match(r"^(HERE|ASSETS)\s*=", line):
            continue
        # 各模块自己的 coding 声明，合并后只保留文件头那一行
        if re.match(r"^# -\*- coding: utf-8 -\*-", line):
            continue
        kept.append(line)

    text = "\n".join(kept)
    text = re.sub(r"\bps\.", "", text)          # paper_style 命名空间前缀
    return text.strip("\n")


HEADER = '''# -*- coding: utf-8 -*-
"""cn-paper-docx 便携版 —— 单个文件，任何有 Python 的 AI 沙箱都能跑。

中文论文 / 报告 Word 通用排版。按「标题层级 + 段落角色」组织，不绑定章节安排。
只有三级标题，三级之后直接是正文。

依赖:
    python-docx   必需
    Pillow        可选（读图片像素尺寸；没有也能插图，只是不自动缩放）

用法:
    python cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4
    python cn_paper_docx.py new 稿件.txt -o 论文.docx --paper a4 --columns 2
    python cn_paper_docx.py new 稿件.txt -o 论文.docx --paper letter
    python cn_paper_docx.py new --demo -o 示例.docx
    python cn_paper_docx.py apply 原稿.docx -o 排好版.docx --paper a4
    python cn_paper_docx.py apply 原稿.docx --dry-run
    python cn_paper_docx.py template -o 模板.docx --paper a4
    python cn_paper_docx.py verify 论文.docx --paper a4
    python cn_paper_docx.py spec

自定义参数（不用改脚本）:
    python cn_paper_docx.py dump-config --paper a4 -o 我的版式.json
    # 编辑 JSON：只写要改的项，其余继承 base
    python cn_paper_docx.py new 稿件.txt -o 论文.docx --config 我的版式.json
    python cn_paper_docx.py verify 论文.docx --config 我的版式.json
    python cn_paper_docx.py spec --config 我的版式.json

稿件标记语法（UTF-8 文本）:
    @cover                    → 封面页，下面跟 `键: 值` 行
        标题: 论文题目
        副标题: ——副标题
        作者: 张三
        单位: ×××大学 ×××学院
        日期: 2026年9月
    @toc                      → 自动目录域（Word 里按 F9 更新页码）
    @pagenum                  → 另起一节，页码从 1 开始
    @pagenum continue         → 另起一节，页码接着编
    @pagebreak                → 分页符
    @title 论文题目            → 居中大标题（黑体小二）
    @center 一行居中文字        → 居中正文
    # 一、一级标题             → 一级标题
    ## （一）二级标题          → 二级标题
    ### 1.三级标题             → 三级标题
    #! 参考文献                → 标题居中（# 后加 !）
    - Step 1：xxx             → 步骤行
    > 注：xxx                 → 表注
    [图3-1 说明](图.png)       → 插图 + 图题（图题在图下）
    [图3-1 说明]               → 只有图题
    [表3-1 说明](表.png)       → 表题 + 表格图片（表题在表上）
    [表3-1 说明]               → 只有表题
    | 模型 | 准确率 |          → 三线表（首行表头）
    [@] 作者. 题名[J]. 刊, 年.  → 参考文献条目
    其他非空行                 → 正文（宋体小四、首行缩进 2 字符）

    ★ 只有三级标题。用了 #### 会直接报错。

本文件由 portable/build_portable.py 从 scripts/ 自动生成，不要直接改这里。
"""
import argparse
import io
import json
import os
import re
import shutil
import sys

from docx import Document
from docx.shared import Pt, Cm, Inches, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

'''


def build_script():
    parts = [
        ("paper_style.py", "版式核心：字体、行距、页面、样式、三线表"),
        ("new_paper.py", "从零生成：解析标记语法，生成完整文档"),
        ("apply_format.py", "给已有 docx 套格式：识别段落角色并重排"),
        ("make_template.py", "生成空白模板与样式样张"),
    ]
    out = [HEADER]
    for name, title in parts:
        src = read(os.path.join(SCRIPTS, name))
        out.append("\n# " + "=" * 70 + f"\n# {title}（来自 scripts/{name}）\n"
                   + "# " + "=" * 70 + "\n\n")
        out.append(strip_module(src))
        out.append("\n\n")
    out.append(read(os.path.join(HERE, "cli_append.py")))
    return "".join(out)


def build_spec():
    text = read(os.path.join(ROOT, "references", "format-spec.md"))
    text = text.replace("`apply_format.py`", "`cn_paper_docx.py apply`")
    text = text.replace("`new_paper.py`", "`cn_paper_docx.py new`")
    text = text.replace("`make_template.py`", "`cn_paper_docx.py template`")
    text = text.replace("`paper_style.py`", "`cn_paper_docx.py` 内置")
    header = (
        "> **便携版规范。** 数值与主 skill 的 `references/format-spec.md` 一致，"
        "供 ChatGPT / Grok / Gemini / Claude.ai 等 AI 读取。\n"
        "> 命令示例里的 `cn_paper_docx.py` 指本目录下的单文件工具。\n\n"
    )
    return header + text


ZIP_FILES = [
    "README.md",
    "SKILL.md",
    "AI-INSTRUCTIONS.md",
    "AI-INSTRUCTIONS-compact.md",
    "版式规范.md",
    "cn_paper_docx.py",
    "稿件模板.txt",
    "示例版式.json",
    "示例版式-双栏.json",
]


def build_zip():
    out = os.path.join(HERE, "cn-paper-docx-portable.zip")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in ZIP_FILES:
            path = os.path.join(HERE, name)
            if os.path.exists(path):
                # arcname 用正斜杠，Windows 上解压也正常
                z.write(path, name)
    return out


def main():
    script = build_script()
    script_path = os.path.join(HERE, "cn_paper_docx.py")
    write(script_path, script)
    print("已生成:", script_path, f"({len(script.splitlines())} 行)")

    spec_path = os.path.join(HERE, "版式规范.md")
    write(spec_path, build_spec())
    print("已生成:", spec_path)

    missing = [n for n in ZIP_FILES if not os.path.exists(os.path.join(HERE, n))]
    if missing:
        print("! 以下文件不存在，zip 里会缺少它们:", ", ".join(missing))
    zip_path = build_zip()
    print("已打包:", zip_path)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""从零写一篇完整的、已排好版的 Word 论文。

用法:
    python new_paper.py 稿件.txt -o 论文.docx
    python new_paper.py 稿件.txt -o 论文.docx --paper a4
    python new_paper.py 稿件.txt -o 论文.docx --line-body 1.5 --line-heading 28pt
    python new_paper.py --demo -o 示例.docx      # 生成一份演示稿

版式预设（--paper）:
    letter  Letter 纸，二级标题黑体四号，正文 1.5 倍行距
            （对应「总.docx」）
    a4      A4 纸，二级标题楷体三号+缩进2字符，正文固定值 28 磅
            （对应「第四章　实验流程.docx」）

标题与正文的行距可分别用 --line-heading / --line-body 覆盖；
写 "1.5" 是多倍行距，写 "28pt" 是固定值 28 磅。

完整结构（封面 → 目录 → 正文 → 参考文献）:
    @cover                         → 封面页，下面跟若干 `键: 值` 行
        标题: 论文题目
        副标题: xxx
        作者: 张三
        单位: 南京航空航天大学
        日期: 2026年9月
    @toc                           → 自动目录（Word 里按 F9 更新即生成页码）
    @pagenum                       → 给页脚加居中页码
    @pagebreak                     → 分页符

标记语法（写在 .txt / .md 里，UTF-8 编码）:
    @title 论文题目                 → 居中大标题（黑体小二）
    @center 一行需要居中的文字       → 居中正文
    # 一、研究背景                  → 一级标题（黑体三号）
    ## （一）子标题                 → 二级标题（letter 版黑体四号 / a4 版楷体三号）
    ### 1.三级标题                  → 三级标题（黑体小四）
    #! 参考文献                    → 同上，但该标题居中（# 后加 ! 即居中）

    ★ 只有三级标题，四级请写正文。用了 #### 会直接报错。

    - Step 1：xxx                  → 步骤行（左对齐、缩进2字符）
    > 注：xxx                      → 表注（宋体五号）
    [图3-1 实验场地布置照片](a.png)  → 插图 + 图题（居中，超宽自动缩小）
    [图3-1 实验场地布置照片]         → 只有图题，不插图
    [表3-1 各模型准确率对比](t.png)  → 表题 + 表格图片（表题在上）
    | 模型 | 准确率 |              → 三线表（连续多行；首行是表头）
    [@] 作者. 题目[J]. 期刊, 年.     → 参考文献条目
    其他非空行                      → 正文（宋体小四、首行缩进2字符）

图片路径可以是相对路径（相对稿件文件所在目录）或绝对路径。
空行分段，段内可以硬换行（按 Markdown 软换行合并成一段）。
正文里可以用 **粗体** 标记局部加粗（如表中最高分）。
"""
import argparse
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                                    # noqa: E402
from docx.shared import Pt                                   # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK      # noqa: E402

import paper_style as ps                                     # noqa: E402

# 标题默认左对齐。需要居中的标题（如参考文献、致谢）在 # 后加 ! 显式标记，
# 例如「#! 参考文献」——不按章节名字猜，任何标题都能用。


def add_rich_text(paragraph, text, ea_font=ps.SONG, size=ps.SZ_BODY,
                  base_bold=False):
    """把 `**粗体**` 拆成多个 run 写进段落。"""
    for i, seg in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not seg:
            continue
        r = paragraph.add_run(seg)
        ps.set_run_font(r, ps.TNR, ea_font, size,
                        bold=base_bold or (i % 2 == 1))
    return paragraph


def join_wrapped(parts):
    """把软换行的多行拼成一段：中文之间直接接，英文单词之间补空格。"""
    out = ""
    for seg in parts:
        if not out:
            out = seg
            continue
        if out[-1].isascii() and seg[:1].isascii() and \
                (out[-1].isalnum() or out[-1] in ",.;:)]}") and seg[:1].isalnum():
            out += " " + seg
        else:
            out += seg
    return out


# `[图3-1 说明](图片路径)` —— 方括号里是题注，圆括号里是图片文件
RE_CAPTION_IMG = re.compile(
    r"^[\[【]([图表]\s*\d+(?:[-–.]\d+)?[^\]】]*)[\]】]\s*[（(]([^）)]+)[）)]$")


def _caption(doc, text, role=None):
    """一行居中的图题/表题。"""
    role = role or ps.DEFAULT_ROLES["caption"]
    p = doc.add_paragraph(style="Caption")
    ps.clear_indent(p)
    add_rich_text(p, text, role.get("ea", ps.HEI),
                  role.get("size", ps.SZ_CAPTION),
                  base_bold=bool(role.get("bold", False)))
    return p


def _picture(doc, path, base_dir=None):
    """插入一张居中图片，超过版面宽度自动等比缩小。"""
    full = path
    if base_dir and not os.path.isabs(path):
        full = os.path.join(base_dir, path)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ps.clear_indent(p)
    ps.set_line_spacing(p, ps.LINE_SINGLE)
    got = ps.add_picture_fitted(p, full, ps.content_width(doc))
    if got is None:
        sys.stderr.write(f"  ! 找不到图片，已跳过：{path}\n")
        return None
    return p


def render_cover(doc, lines, i, paper, line_heading, roles=None):
    """解析 @cover 之后的 `键: 值` 行，生成一页封面。返回下一个待处理行号。"""
    fields = []
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            break
        m = re.match(r"^([^：:]{1,8})\s*[：:]\s*(.*)$", s)
        if not m:
            break
        fields.append((m.group(1).strip(), m.group(2).strip()))
        i += 1

    roles = roles or ps.get_roles(paper)
    cover_title = roles["cover_title"]
    cover_subtitle = roles["cover_subtitle"]
    cover_text = roles["cover_text"]
    title_font = {   # 键 → (角色, 段前, 段后)
        "标题": (cover_title, 0, 12),
        "题目": (cover_title, 0, 12),
        "副标题": (cover_subtitle, 0, 8),
        "作者": (cover_text, 6, 4),
        "姓名": (cover_text, 6, 4),
        "单位": (cover_text, 4, 4),
        "学院": (cover_text, 4, 4),
        "日期": (cover_text, 4, 4),
    }
    ps.add_centered_line(doc, "", line=line_heading)          # 顶部留白
    ps.add_centered_line(doc, "", line=line_heading)
    for k, v in fields:
        role, sb, sa = title_font.get(k, (cover_text, 4, 4))
        ps.add_centered_line(doc, v,
                             ea_font=role.get("ea", ps.SONG),
                             size=role.get("size", Pt(14)),
                             bold=bool(role.get("bold", False)),
                             space_before=role.get("before", sb),
                             space_after=role.get("after", sa),
                             line=line_heading)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    return i


def render(doc, text, paper=ps.DEFAULT_PROFILE, line_heading=None,
           base_dir=None, columns=None, column_gap_cm=None):
    lines = text.splitlines()
    i = 0
    buf = []          # 待合并的正文行
    _, prof = ps.get_profile(paper)
    body_cols = prof.get("columns", 1) if columns is None else int(columns)
    body_gap = (prof.get("column_gap_cm", 0.75)
                if column_gap_cm is None else float(column_gap_cm))
    roles = ps.get_roles(prof)
    body = roles["body"]
    title_role = roles["title"]
    caption_role = roles["caption"]
    note_role = roles["note"]
    ref_role = roles["reference"]
    step_role = roles["step"]
    cell_role = roles["table_cell"]

    def flush():
        if buf:
            p = doc.add_paragraph(style="Normal")
            ps.set_first_line_indent_chars(
                p, body.get("indent_chars", 2),
                body.get("size", ps.SZ_BODY).pt)
            add_rich_text(p, join_wrapped(buf),
                          body.get("ea", ps.SONG),
                          body.get("size", ps.SZ_BODY),
                          base_bold=bool(body.get("bold", False)))
            buf.clear()

    while i < len(lines):
        raw = lines[i]
        line = raw.strip()
        i += 1
        # 空行 = 段落分隔；连续的非空行按软换行合并为同一段
        if not line:
            flush()
            continue

        # 遇到非正文结构前，先把攒着的正文段落吐出去
        if (line.startswith(("@", "#", "-", "*", ">", "|", "[@]"))
                or re.match(r"^[\[【][图表]", line)):
            flush()

        # ---- 封面：@cover 后面跟若干 `键: 值` 行
        if line.startswith("@cover"):
            i = render_cover(doc, lines, i, paper, line_heading, roles=roles)
            continue

        # ---- 目录（自动目录域，Word 里按 F9 更新即生成带页码的目录）
        if line.startswith("@toc"):
            title = line[4:].strip() or "目录"
            ea, size, bold, _sb, _sa, _l = ps.heading_spec(paper, "Heading 1",
                                                           line_heading)
            p = doc.add_paragraph(style="Heading 1")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ps.clear_indent(p)
            ps.set_run_font(p.add_run(title), ps.TNR, ea, size, bold=bold)
            tp = doc.add_paragraph()
            ps.clear_indent(tp)
            ps.add_toc_field(tp)
            continue

        # ---- 页码：从此处另起一节，并从这里开始编号
        #      @pagenum          → 另起一页，页码从 1 开始（封面/目录不编码）
        #      @pagenum continue → 另起一页，页码接着往下编
        if line.startswith("@pagenum"):
            arg = line[len("@pagenum"):].strip().lower()
            restart = None if arg in ("continue", "续", "接续") else (
                int(arg) if arg.isdigit() else 1)
            sec = ps.start_numbered_section(doc, restart=restart)
            ps.set_columns(sec, body_cols, body_gap)
            ps.add_page_number_footer(doc, sections=[sec],
                                      ea_font=note_role.get("ea", ps.SONG),
                                      size=note_role.get("size", ps.SZ_NOTE))
            continue

        # ---- 指令
        if line.startswith("@title"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ps.clear_indent(p)
            p.paragraph_format.space_before = Pt(title_role.get("before", 0))
            p.paragraph_format.space_after = Pt(title_role.get("after", 12))
            ps.set_line_spacing(p, line_heading or title_role.get("line")
                                or ps.LINE_15)
            add_rich_text(p, line[len("@title"):].strip(),
                          title_role.get("ea", ps.HEI),
                          title_role.get("size", ps.SZ_TITLE),
                          base_bold=bool(title_role.get("bold", False)))
            continue
        if line.startswith("@center"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ps.clear_indent(p)
            add_rich_text(p, line[len("@center"):].strip(),
                          body.get("ea", ps.SONG),
                          body.get("size", ps.SZ_BODY))
            continue
        if line.startswith("@pagebreak"):
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            continue

        # ---- 标题：# 一、标题 ；#! 标题 表示该标题居中
        m = re.match(r"^(#{1,4})(!?)\s*(.+)$", line)
        if m:
            lvl, centered, txt = len(m.group(1)), m.group(2) == "!", m.group(3).strip()
            if lvl == 4:
                sys.exit(f"第 {i} 行用了「####」四级标题，但本规范只有三级标题。\n"
                         f"请把「{txt}」改成 ### 三级标题，或直接写成正文。")
            style = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}[lvl]
            ea, size, bold, _sb, _sa, _line = ps.heading_spec(
                paper, style, line_heading)
            p = doc.add_paragraph(style=style)
            ps.clear_indent(p)
            if lvl == 2:
                prof = ps.get_profile(paper)[1]
                chars = prof["h2_indent_chars"]
                if chars:
                    ps.set_first_line_indent_chars(p, chars, size.pt)
            add_rich_text(p, txt, ea, size, base_bold=bold)
            if centered:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

        # ---- 步骤行
        if re.match(r"^[-*]\s+", line):
            p = doc.add_paragraph(style="Step")
            ps.set_first_line_indent_chars(
                p, step_role.get("indent_chars", 2),
                step_role.get("size", ps.SZ_BODY).pt)
            txt = re.sub(r"^[-*]\s+", "", line)
            m2 = re.match(r"^((?:Step|步骤)\s*\d+|\d+)\s*[：:]\s*(.*)$", txt)
            if m2:
                r = p.add_run(m2.group(1) + "：")
                ps.set_run_font(r, ps.TNR,
                                step_role.get("num_ea", ps.HEI),
                                step_role.get("size", ps.SZ_BODY), bold=False)
                add_rich_text(p, m2.group(2),
                              step_role.get("ea", ps.SONG),
                              step_role.get("size", ps.SZ_BODY))
            else:
                add_rich_text(p, txt, step_role.get("ea", ps.SONG),
                              step_role.get("size", ps.SZ_BODY))
            continue

        # ---- 表注
        if line.startswith(">"):
            p = doc.add_paragraph(style="TableNote")
            ps.set_first_line_indent_chars(
                p, note_role.get("indent_chars", 2),
                note_role.get("size", ps.SZ_NOTE).pt)
            add_rich_text(p, line.lstrip("> ").strip(),
                          note_role.get("ea", ps.SONG),
                          note_role.get("size", ps.SZ_NOTE))
            continue

        # ---- 图题 / 表题（黑体小四居中），可带图片路径
        mi = RE_CAPTION_IMG.match(line)
        m = mi or re.match(r"^[\[【]([图表]\s*\d+(?:[-–.]\d+)?.*?)[\]】]$", line)
        if m:
            cap = m.group(1).strip()
            is_table = cap.startswith("表")
            if mi:
                # 表题在表格上方，图题在图下方
                if is_table:
                    _caption(doc, cap, caption_role)
                _picture(doc, mi.group(2).strip(), base_dir)
                if not is_table:
                    _caption(doc, cap, caption_role)
            else:
                _caption(doc, cap, caption_role)
            continue

        # ---- 参考文献条目
        if line.startswith("[@]"):
            p = doc.add_paragraph(style="Reference")
            ps.clear_indent(p)
            add_rich_text(p, line[3:].strip(),
                          ref_role.get("ea", ps.SONG),
                          ref_role.get("size", ps.SZ_REF))
            continue

        # ---- 表格
        if line.startswith("|"):
            rows = []
            while i <= len(lines):
                cur = lines[i - 1].strip() if i - 1 < len(lines) else ""
                if not cur.startswith("|"):
                    break
                if not re.match(r"^[-:\s|]+$", cur.strip("|")):
                    rows.append([c.strip() for c in cur.strip("|").split("|")])
                i += 1
            i -= 1
            if rows:
                width = max(len(r) for r in rows)
                tbl = doc.add_table(rows=len(rows), cols=width)
                for ri, row in enumerate(rows):
                    for ci in range(width):
                        cell = tbl.cell(ri, ci)
                        p = cell.paragraphs[0]
                        for r in list(p.runs):      # 清掉占位 run
                            r._element.getparent().remove(r._element)
                        ps.clear_indent(p)
                        add_rich_text(p, row[ci] if ci < len(row) else "",
                                      cell_role.get("ea", ps.SONG),
                                      cell_role.get("size", ps.SZ_NOTE))
                ps.fit_table_width(tbl, ps.content_width(doc))
                ps.make_three_line_table(
                    tbl, header_rows=1, lines=prof.get("table_lines"),
                    cell_role=cell_role)
            continue

        # ---- 正文：攒进缓冲区，等空行或下一个结构再落段
        buf.append(line)
    flush()
    return doc


DEMO = """@cover
标题: 论文题目
副标题: ——副标题或研究说明
作者: 张三
单位: ×××大学 ×××学院
日期: 2026年9月

@toc
@pagenum

# 一、一级标题
这是一段正文，中文用宋体小四，首行缩进 2 字符，两端对齐。
源码里可以硬换行，同一段的多行会被自动拼接成一段。
**这一段里有加粗的部分**，用来标注关键结论。

## （一）二级标题
正文照常写。

### 1.三级标题
三级标题之后就只有正文了，本规范不设四级标题。

# 二、流程类内容
## （一）操作步骤
- Step 1：需要分步说明的内容用这种步骤行，比正文更易扫读。
- Step 2：序号用黑体，冒号后的内容用宋体。
- Step 3：实验流程、操作手册、算法步骤都适用。

[图2-1 图题写在这里，会自动居中（想插图就在后面加圆括号写图片路径）]

[表2-1 表题放在表格上方]
| 方法 | 指标 A（%） | 指标 B（%） |
| --- | --- | --- |
| 基线方法 | 68.2 ± 9.1 | 71.5 ± 8.4 |
| 对比方法 | 74.6 ± 7.8 | 78.3 ± 6.9 |
| **本文方法** | **81.3 ± 5.2** | **84.7 ± 4.6** |
> 注：粗体表示该列最优结果。表格会自动套三线表。

#! 参考文献
[@] PFURTSCHELLER G, LOPES DA SILVA F H. Motor imagery and direct
    brain-computer communication[J]. Proceedings of the IEEE, 2001, 89(7): 1123-1134.
"""


def main():
    ap = argparse.ArgumentParser(description="按论文版式生成 Word 文档")
    ap.add_argument("src", nargs="?", help="标记稿件（.txt/.md，UTF-8）")
    ap.add_argument("-o", "--out", required=True, help="输出的 .docx 路径")
    ap.add_argument("--paper", default=ps.DEFAULT_PROFILE,
                    choices=list(ps.PROFILES), help="版式预设（默认 letter）")
    ap.add_argument("--config", default=None,
                    help="JSON 版式配置（可只写要改的项，其余继承 --paper / base）")
    ap.add_argument("--line-body", default=None,
                    help="正文行距：1.5 表示 1.5 倍，28pt 表示固定值 28 磅")
    ap.add_argument("--line-heading", default=None,
                    help="标题行距：同上。与正文可分别设置")
    ap.add_argument("--columns", type=ps.parse_columns, default=None,
                    help="栏数：1 单栏（默认）或 2 双栏")
    ap.add_argument("--column-gap", type=float, default=None,
                    help="双栏栏间距，单位 cm（默认 0.75）")
    ap.add_argument("--demo", action="store_true", help="忽略 src，输出演示稿")
    ap.add_argument("--no-toc-refresh", action="store_true",
                    help="不调用 Word 刷新目录页码（默认会刷新一次，"
                         "这样打开文档目录就是好的，之后也不会再提示更新）")
    args = ap.parse_args()

    if args.config:
        try:
            args.paper = ps.apply_config(args.config, base=args.paper)
        except Exception as e:
            sys.exit(f"读取版式配置失败：{e}")

    if args.demo:
        text = DEMO
    else:
        if not args.src:
            ap.error("请给出稿件路径，或用 --demo")
        with io.open(args.src, encoding="utf-8") as f:
            text = f.read()

    _, prof = ps.get_profile(args.paper)
    body_cols = prof.get("columns", 1) if args.columns is None else args.columns
    body_gap = (prof.get("column_gap_cm", 0.75)
                if args.column_gap is None else args.column_gap)
    # 有 @pagenum 时封面/目录留在单栏节，正文从新节开始分栏
    initial_cols = 1 if (body_cols == 2 and "@pagenum" in text) else body_cols

    line_head = ps.parse_line(args.line_heading)
    doc = Document()
    ps.apply_page_setup(doc, args.paper, columns=initial_cols,
                        column_gap_cm=body_gap)
    ps.apply_styles(doc, args.paper, args.line_body, args.line_heading)
    for p in list(doc.paragraphs):          # 清掉默认空段落
        p._element.getparent().remove(p._element)
    base_dir = os.path.dirname(os.path.abspath(args.src)) if args.src else os.getcwd()
    render(doc, text, args.paper, line_head, base_dir,
           columns=body_cols, column_gap_cm=body_gap)
    doc.save(args.out)
    print("已生成:", args.out)

    # 有目录域时，用 Word 刷一次，让打开就能看到带页码的目录
    if "@toc" in text and not args.no_toc_refresh:
        ok, msg = ps.refresh_fields_with_word(args.out)
        if ok:
            print("目录已刷新（打开即为带页码的目录，之后不会再提示更新）")
        else:
            print(f"跳过目录刷新（{msg}）——在 Word 里按一次 F9 即可")


if __name__ == "__main__":
    main()

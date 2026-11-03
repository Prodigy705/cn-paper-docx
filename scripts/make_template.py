# -*- coding: utf-8 -*-
"""生成模板与样式样张。

用法:
    python make_template.py                 # 两套版式各生成一份模板和样张
    python make_template.py --paper a4      # 只生成 A4 版
    python make_template.py --out 我的模板.docx --paper letter

assets/template-<profile>.docx  空白模板：样式已定义好，用 Word 打开直接写。
assets/specimen-<profile>.docx  样式样张：展示每种段落长什么样，用来核对格式。
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                                    # noqa: E402
from docx.shared import Pt                                   # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH                # noqa: E402

import paper_style as ps                                     # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")


def build_template(paper, line_body=None, line_heading=None,
                   columns=None, column_gap_cm=None):
    doc = Document()
    ps.apply_page_setup(doc, paper, columns=columns,
                        column_gap_cm=column_gap_cm)
    ps.apply_styles(doc, paper, line_body, line_heading)
    for p in list(doc.paragraphs):          # 清掉 python-docx 默认模板的空段落
        p._element.getparent().remove(p._element)
    p = doc.add_paragraph(style="Normal")
    ps.set_first_line_indent_chars(p, 2, 12)
    return doc


def build_specimen(paper, line_body=None, line_heading=None,
                   columns=None, column_gap_cm=None):
    """样张里的每行说明都从版式预设里读出来，改预设这里自动跟着变。"""
    _, prof = ps.get_profile(paper)
    roles = ps.get_roles(prof)
    body = roles["body"]
    title = roles["title"]
    step = roles["step"]
    cap = roles["caption"]
    note_r = roles["note"]
    quote_r = roles["quote"]
    ref = roles["reference"]
    cell = roles["table_cell"]
    tbl_lines = prof.get("table_lines", {"top": 1.5, "header": 0.5, "bottom": 1.5})
    doc = Document()
    ps.apply_page_setup(doc, paper, columns=columns,
                        column_gap_cm=column_gap_cm)
    ps.apply_styles(doc, paper, line_body, line_heading)

    def note(p, text):
        """在段落尾部加一行楷体小字的说明，标注该段用的是哪个样式。"""
        r = p.add_run("　　← " + text)
        ps.set_run_font(r, ps.TNR, ps.KAI, Pt(9), bold=False)

    # 大标题
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ps.clear_indent(t)
    ps.set_line_spacing(t, line_heading or title.get("line") or ps.LINE_15)
    ps.set_run_font(t.add_run("论文版式样张"), ps.TNR,
                    title.get("ea", ps.HEI), title.get("size", ps.SZ_TITLE),
                    bold=bool(title.get("bold", False)))
    note(t, prof["label"])

    # 三级标题，逐级展示
    labels = {
        "Heading 1": "一级标题",
        "Heading 2": "二级标题",
        "Heading 3": "三级标题",
    }
    for i, (style, label) in enumerate(labels.items(), 1):
        ea, size, bold, _sb, _sa, _line = ps.heading_spec(paper, style, line_heading)
        p = doc.add_paragraph(style=style)
        ps.clear_indent(p)
        if style == "Heading 2" and prof["h2_indent_chars"]:
            ps.set_first_line_indent_chars(p, prof["h2_indent_chars"], size.pt)
        extra = "，首行缩进2字符" if (style == "Heading 2"
                                     and prof["h2_indent_chars"]) else ""
        prefix = {1: "一、", 2: "（一）", 3: "1."}[i]
        txt = prefix + label
        ps.set_run_font(p.add_run(txt), ps.TNR, ea, size, bold=bold)
        note(p, f"Heading {i}｜{ea} {size.pt}pt{'加粗' if bold else '不加粗'}{extra}")

    # 正文
    p = doc.add_paragraph(style="Normal")
    ps.set_first_line_indent_chars(p, body.get("indent_chars", 2),
                                   body.get("size", ps.SZ_BODY).pt)
    ps.set_run_font(p.add_run(
        f"正文：{body.get('ea', ps.SONG)} {body.get('size', ps.SZ_BODY).pt:g}pt，"
        f"西文 Times New Roman，首行缩进 {body.get('indent_chars', 2)} 字符，两端对齐。"
        "这是正文段落的样子，用来核对行距与缩进是否与论文一致。"),
        body.get("ascii", ps.TNR), body.get("ea", ps.SONG),
        body.get("size", ps.SZ_BODY))
    note(p, "Normal")

    p = doc.add_paragraph(style="Normal")
    ps.set_first_line_indent_chars(p, body.get("indent_chars", 2),
                                   body.get("size", ps.SZ_BODY).pt)
    ps.set_run_font(p.add_run("三级标题之后就只有正文了，本规范不设四级标题。"),
                    body.get("ascii", ps.TNR), body.get("ea", ps.SONG),
                    body.get("size", ps.SZ_BODY))
    note(p, "Normal")

    # 步骤行
    for txt in (f"Step 1：步骤行——左对齐、首行缩进 {step.get('indent_chars', 2)} 字符，"
                f"序号用{step.get('num_ea', ps.HEI)}，内容用{step.get('ea', ps.SONG)}。",
                "Step 2：实验流程、操作手册、算法步骤这类“1、2、3…”都用此样式。"):
        p = doc.add_paragraph(style="Step")
        ps.set_first_line_indent_chars(p, step.get("indent_chars", 2),
                                       step.get("size", ps.SZ_BODY).pt)
        num, _, rest = txt.partition("：")
        ps.set_run_font(p.add_run(num + "："), ps.TNR,
                        step.get("num_ea", ps.HEI),
                        step.get("size", ps.SZ_BODY), bold=False)
        ps.set_run_font(p.add_run(rest), step.get("ascii", ps.TNR),
                        step.get("ea", ps.SONG), step.get("size", ps.SZ_BODY))
        note(p, "Step")

    # 图题 / 表题
    for txt in (f"图1-1　图题：{cap.get('ea', ps.HEI)} "
                f"{cap.get('size', ps.SZ_CAPTION).pt:g}pt、居中、不缩进",
                "表1-1　表题：同样用 Caption 样式，置于表格上方"):
        p = doc.add_paragraph(style="Caption")
        ps.clear_indent(p)
        ps.set_run_font(p.add_run(txt), cap.get("ascii", ps.TNR),
                        cap.get("ea", ps.HEI), cap.get("size", ps.SZ_CAPTION))
        note(p, "Caption")

    # 三线表
    tbl = doc.add_table(rows=4, cols=3)
    data = [["模型", "数据集 A（%）", "数据集 B（%）"],
            ["基线方法", "68.2 ± 9.1", "71.5 ± 8.4"],
            ["对比方法", "74.6 ± 7.8", "78.3 ± 6.9"],
            ["本文方法", "81.3 ± 5.2", "84.7 ± 4.6"]]
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            tbl.cell(i, j).text = v
    ps.fit_table_width(tbl, ps.content_width(doc))
    ps.make_three_line_table(tbl, header_rows=1, lines=tbl_lines,
                             cell_role=cell)
    note(doc.add_paragraph(style="TableNote"),
         f"三线表：顶线{tbl_lines.get('top', 1.5):g}pt／"
         f"表头下{tbl_lines.get('header', 0.5):g}pt／"
         f"底线{tbl_lines.get('bottom', 1.5):g}pt")

    n = doc.add_paragraph(style="TableNote")
    ps.set_first_line_indent_chars(n, note_r.get("indent_chars", 2),
                                   note_r.get("size", ps.SZ_NOTE).pt)
    ps.set_run_font(n.add_run(
        f"注：粗体表示该列最高观察结果。表注用{note_r.get('ea', ps.SONG)} "
        f"{note_r.get('size', ps.SZ_NOTE).pt:g}pt。"),
        note_r.get("ascii", ps.TNR), note_r.get("ea", ps.SONG),
        note_r.get("size", ps.SZ_NOTE))
    note(n, "TableNote")

    # 引文
    p = doc.add_paragraph(style="Quote-CN")
    ps.set_run_font(p.add_run(
        f"引文段落用 Quote-CN 样式：{quote_r.get('ea', ps.KAI)} "
        f"{quote_r.get('size', ps.SZ_BODY).pt:g}pt、左右各缩进 "
        f"{quote_r.get('left_cm', 0.74):g}cm。"),
        quote_r.get("ascii", ps.TNR), quote_r.get("ea", ps.KAI),
        quote_r.get("size", ps.SZ_BODY))
    note(p, "Quote-CN")

    # 居中标题 + 参考文献
    ea, size, bold, _sb, _sa, _line = ps.heading_spec(paper, "Heading 1", line_heading)
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ps.clear_indent(p)
    ps.set_run_font(p.add_run("参考文献"), ps.TNR, ea, size, bold=bold)
    note(p, "居中标题：用 new 命令时写「#! 参考文献」")

    for txt in ["PFURTSCHELLER G, LOPES DA SILVA F H. Motor imagery and direct "
                "brain-computer communication[J]. Proceedings of the IEEE, 2001, 89(7): 1123-1134.",
                "LAWHERN V J, SOLON A J, WAYTOWICH N R, et al. EEGNet: a compact "
                "convolutional neural network for EEG-based brain-computer interfaces[J]. "
                "Journal of Neural Engineering, 2018, 15(5): 056013."]:
        p = doc.add_paragraph(style="Reference")
        ps.clear_indent(p)
        ps.set_run_font(p.add_run(txt), ref.get("ascii", ps.TNR),
                        ref.get("ea", ps.SONG),
                        ref.get("size", ps.SZ_REF), bold=False)
    return doc


def main():
    ap = argparse.ArgumentParser(description="生成论文版式模板与样张")
    ap.add_argument("--paper", default=None, choices=list(ps.PROFILES),
                    help="版式预设；不指定则两套都生成")
    ap.add_argument("--config", default=None,
                    help="JSON 版式配置（可只写要改的项，其余继承 base）")
    ap.add_argument("--line-body", default=None, help="正文行距，如 1.5 或 28pt")
    ap.add_argument("--line-heading", default=None, help="标题行距，如 1.5 或 28pt")
    ap.add_argument("--columns", type=ps.parse_columns, default=None,
                    help="栏数：1 单栏（默认）或 2 双栏")
    ap.add_argument("--column-gap", type=float, default=None,
                    help="双栏栏间距，单位 cm（默认 0.75）")
    ap.add_argument("--out", default=None, help="仅生成模板并指定输出路径")
    args = ap.parse_args()

    if args.config:
        try:
            args.paper = ps.apply_config(
                args.config, base=args.paper or ps.DEFAULT_PROFILE)
        except Exception as e:
            sys.exit(f"读取版式配置失败：{e}")

    if args.out:
        paper = args.paper or ps.DEFAULT_PROFILE
        build_template(paper, args.line_body, args.line_heading,
                       args.columns, args.column_gap).save(args.out)
        print("已生成模板:", args.out)
        return

    os.makedirs(ASSETS, exist_ok=True)
    for paper in ([args.paper] if args.paper else list(ps.PROFILES)):
        tp = os.path.join(ASSETS, f"template-{paper}.docx")
        sp = os.path.join(ASSETS, f"specimen-{paper}.docx")
        build_template(paper, args.line_body, args.line_heading,
                       args.columns, args.column_gap).save(tp)
        build_specimen(paper, args.line_body, args.line_heading,
                       args.columns, args.column_gap).save(sp)
        print(f"[{paper}] 模板 → {tp}")
        print(f"[{paper}] 样张 → {sp}")


if __name__ == "__main__":
    main()

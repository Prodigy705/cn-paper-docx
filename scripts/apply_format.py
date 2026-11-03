# -*- coding: utf-8 -*-
"""把已有的 .docx 按论文版式重新排版（原文件不动，另存为新文件）。

用法:
    python apply_format.py 原稿.docx -o 排好版.docx
    python apply_format.py 原稿.docx -o 排好版.docx --paper a4
    python apply_format.py 原稿.docx -o 排好版.docx --center-titles "参考文献,致谢"
    python apply_format.py 原稿.docx --dry-run          # 只看会改什么，不写文件

版式预设（--paper）:
    letter  Letter 纸，二级标题黑体四号，正文 1.5 倍行距（对应「总.docx」）
    a4      A4 标准页边距，二级标题楷体三号+缩进2字符，正文固定值 28 磅
            （对应「第四章　实验流程.docx」）

做的事:
    1. 页面：纸张与页边距
    2. 样式：重定义 Normal / Heading 1-3 / Caption / TableNote / Step / Reference
    3. 段落：按编号特征识别各级标题、图题表题、表注、步骤行、参考文献，套用对应
       样式，并补齐中文（eastAsia）字体，消除“等线/宋体混排”
    4. 表格：改成学术三线表，表内文字宋体五号居中
    5. 图片段落居中、取消首行缩进

不会做的事:
    - 不会改文字内容、不会删图片、不会动批注
    - 不会自动判断引文段落（Quote-CN），需要的话在 Word 里手动套用该样式
    - 标题默认两端对齐；要居中哪些标题用 --center-titles 明确指定，不按章节名猜
    - 本规范只有三级标题，「（1）」这类四级编号会降为正文
"""
import argparse
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                                    # noqa: E402
from docx.shared import Pt                                   # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH                # noqa: E402
from docx.oxml.ns import qn                                  # noqa: E402

import paper_style as ps                                     # noqa: E402

# ------------------------------------------------------------ 段落识别规则
# 中文论文常见两套编号习惯，段内的 「1.」 在两边含义不同，所以要先判断全篇
# 用的是哪一套（见 detect_scheme）：
#
#   中文式 "cn"   一、  / （一） / 1.          —— 1. 是【三级】
#   数字式 "num"  1.    / 1.1   / 1.1.1        —— 1. 是【一级】
#
# 另外「第一章」「壹、」这类章标题两套都认。
_CN_NUM = "一二三四五六七八九十壹贰叁肆伍陆柒捌玖拾"
RE_H1_CN = re.compile(rf"^[{_CN_NUM}]+\s*[、．.]\s*\S")
RE_H1_CH = re.compile(rf"^第\s*[{_CN_NUM}\d]+\s*[章篇节部分]")
RE_H2_CN = re.compile(rf"^[（(]\s*[{_CN_NUM}]+\s*[）)]\s*\S")
RE_H2_NUM = re.compile(r"^\d+\.\d+(?!\.)\s*\S")           # 4.1
RE_H3_NUM = re.compile(r"^\d+\.\d+\.\d+\s*\S")            # 4.1.1
RE_H3_CN = re.compile(r"^\d+\s*[、．.]\s*\S")              # 1.
RE_H4_ANY = re.compile(r"^[（(]\s*\d+\s*[）)]\s*\S")       # （1）→ 不设四级，降为正文
RE_STEP = re.compile(r"^(?:Step|步骤)\s*\d+\s*[：:．.]|^\d+\s*[：:]|^第\s*[一二三四五六七八九十\d]+\s*步")
RE_CAPTION = re.compile(r"^[图表]\s*\d+\s*[-–—.－]\s*\d+")
# 图/表号不带章节序号的写法（「图 3 系统架构」），只在短句且不以句号结尾时算题注
RE_CAPTION_LOOSE = re.compile(r"^[图表]\s*\d+\s*[.、：:]?\s*\S")
RE_CAPTION_EN = re.compile(r"^(?:Figure|Fig\.?|Table)\s*\d+", re.I)
RE_NOTE = re.compile(r"^注\s*[：:]")
RE_REF_ENTRY = re.compile(r"\[(J|C|D|M|N|R|S|P|EB/OL|Z|A)\]|^\[\d+\]")


def detect_scheme(doc):
    """扫全文判断编号体系，返回 "num" 或 "cn"。

    只要出现「1.1」这种两段式编号，就认定是数字式（1. / 1.1 / 1.1.1），
    此时「1.」是一级；否则按中文式（一、/（一）/1.）处理，「1.」是三级。
    """
    for p in doc.paragraphs:
        t = p.text.strip()
        if RE_H3_NUM.match(t) or RE_H2_NUM.match(t):
            return "num"
    return "cn"

# 这些样式本身就是“结论”，认不出新类型时不要覆盖。
# 注意不含 Heading 4 —— 本规范只有三级标题，原稿里的四级会被降为正文。
_PRESERVE_STYLES = {"Heading 1", "Heading 2", "Heading 3",
                    "Caption", "Reference", "Step", "TableNote", "Quote-CN"}


def looks_like_heading(text, limit=45):
    """短、且不以句号结尾，才可能是小标题（避免把正文句子误判成标题）。"""
    return len(text) <= limit and not text.endswith(("。", "；", "，", "."))


def classify(text, scheme="cn"):
    """返回该段落应使用的样式名；无法判断时返回 None。

    scheme: "cn" = 一、/（一）/1.（1. 为三级）
            "num" = 1. / 1.1 / 1.1.1（1. 为一级）
    """
    t = text.strip()
    if not t:
        return None
    if RE_CAPTION.match(t) or RE_CAPTION_EN.match(t):
        return "Caption"
    if RE_CAPTION_LOOSE.match(t) and looks_like_heading(t, 40):
        return "Caption"                    # 「图 3 系统架构」这种不带章节号的
    if RE_NOTE.match(t):
        return "TableNote"
    if (RE_H1_CN.match(t) or RE_H1_CH.match(t)) and len(t) <= 40:
        return "Heading 1"
    if RE_H3_NUM.match(t) and looks_like_heading(t):        # 4.1.1 先于 4.1 判
        return "Heading 3"
    if RE_H2_NUM.match(t) and looks_like_heading(t):        # 4.1
        return "Heading 2"
    if RE_H2_CN.match(t) and looks_like_heading(t, 50):
        return "Heading 2"
    if RE_STEP.match(t) and not RE_H3_CN.match(t):
        return "Step"
    if RE_H3_CN.match(t) and looks_like_heading(t):
        # 「1.」是三级还是（数字式下的）一级，取决于全篇用哪套编号
        return "Heading 1" if scheme == "num" else "Heading 3"
    if RE_H4_ANY.match(t):
        return "Normal"          # 不设四级标题，降为正文
    if RE_REF_ENTRY.search(t) and len(t) > 30:
        return "Reference"
    return None


def has_drawing(paragraph):
    return paragraph._p.findall(".//" + qn("w:drawing")) or \
        paragraph._p.findall(".//" + qn("w:pict"))


# ------------------------------------------------------------ 主流程
def apply_to_document(doc, paper=ps.DEFAULT_PROFILE, log=None, center_titles=(),
                      line_body=None, line_heading=None,
                      columns=None, column_gap_cm=None):
    """center_titles: 需要居中的标题文字集合（如 {"参考文献", "致谢"}）。
    默认为空 —— 标题一律两端对齐，不按章节名做猜测。"""
    stats = {}
    center_titles = {t.strip().rstrip("：:") for t in center_titles if t.strip()}
    ps.apply_page_setup(doc, paper, columns=columns,
                        column_gap_cm=column_gap_cm)
    ps.apply_styles(doc, paper, line_body, line_heading)

    _, prof = ps.get_profile(paper)
    body_cols = prof.get("columns", 1) if columns is None else int(columns)
    roles = ps.get_roles(prof)
    role_spec = _role_spec(prof, line_body)
    scheme = detect_scheme(doc)

    for p in doc.paragraphs:
        text = p.text.strip()
        cur = p.style.name

        if not text:
            if has_drawing(p):          # 纯图片段落：居中、不缩进
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ps.clear_indent(p)
                stats["图片居中"] = stats.get("图片居中", 0) + 1
            continue

        target = classify(text, scheme)
        if target is None:
            # 认不出类型时：已经是专用样式的（标题/图题/参考文献/步骤/表注）
            # 原样保留，只有普通正文类才归一为 Normal。
            target = cur if cur in _PRESERVE_STYLES else "Normal"

        if target != cur:
            try:
                p.style = doc.styles[target]
            except KeyError:
                continue
            if cur != target and target == "Normal" and RE_H4_ANY.match(text):
                key = "四级编号 → 正文（本规范不设四级）"
            else:
                key = f"{cur} → {target}"
            stats[key] = stats.get(key, 0) + 1

        _normalize_paragraph(p, target, paper, doc, line_body, line_heading,
                             role_spec=role_spec)
        # 只有用户点名要居中的标题才居中（不按章节名猜）
        if center_titles and text.rstrip("：:") in center_titles:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            stats["标题居中"] = stats.get("标题居中", 0) + 1
        if log is not None:
            log.append((target, text[:60]))

    for t in doc.tables:
        if body_cols > 1:
            ps.fit_table_width(t, ps.content_width(doc))
        ps.make_three_line_table(t, header_rows=1,
                                 lines=prof.get("table_lines"),
                                 cell_role=roles["table_cell"])
        stats["表格转三线表"] = stats.get("表格转三线表", 0) + 1

    return stats


def _role_spec(prof, line_body=None):
    """按版式预设的角色表生成各段落样式的 (字体, 西文字体, 字号, 行距,
    缩进字符, 对齐)。body/step/quote 的正文行距可被 --line-body 覆盖。"""
    roles = ps.get_roles(prof)
    override = ps.parse_line(line_body)

    def spec(name, use_body_override=False):
        r = roles[name]
        line = override if (use_body_override and override) else r.get("line")
        return (r.get("ea", ps.SONG), r.get("ascii", ps.TNR),
                r.get("size", ps.SZ_BODY), line,
                r.get("indent_chars", 0),
                ps.align_of(r.get("align")))

    return {
        "Normal": spec("body", True),
        "Step": spec("step", True),
        "Caption": spec("caption"),
        "TableNote": spec("note"),
        "Reference": spec("reference"),
        "Quote-CN": spec("quote", True),
    }


def _normalize_paragraph(p, style_name, paper, doc, line_body=None,
                         line_heading=None, role_spec=None):
    """按目标样式补齐段落/字符格式：字体、缩进、行距、对齐。"""
    key, prof = ps.get_profile(paper)
    if style_name.startswith("Heading"):
        name = style_name if style_name in prof["headings"] else "Heading 3"
        ea, size, _bold, _sb, _sa, line = ps.heading_spec(paper, name, line_heading)
        ascii_font = ps.TNR
        indent = prof["h2_indent_chars"] if name == "Heading 2" else 0
        align = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        if role_spec is None:
            role_spec = _role_spec(prof, line_body)
        ea, ascii_font, size, line, indent, align = role_spec.get(
            style_name, role_spec["Normal"])
        if line is None:
            line = ps.parse_line(line_body) or prof["body_line"]

    if align is not None:
        p.alignment = align
    if indent:
        ps.set_first_line_indent_chars(p, indent, size.pt)
    else:
        ps.clear_indent(p)
    ps.set_line_spacing(p, line)

    for r in p.runs:
        # 只补字体，不覆盖用户自己设的加粗/颜色；没设字号的按样式给。
        # 例外：标题一律不加粗（本规范硬约束），原稿里的加粗要清掉。
        bold = False if style_name.startswith("Heading") else r.font.bold
        ps.set_run_font(r, ascii_font, ea, r.font.size or size, bold=bold)


def main():
    ap = argparse.ArgumentParser(description="按论文版式重排已有的 docx")
    ap.add_argument("src", help="原 .docx（不会被修改）")
    ap.add_argument("-o", "--out", help="输出的 .docx")
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
    ap.add_argument("--center-titles", default="",
                    help="要居中的标题文字，逗号分隔，如 \"参考文献,致谢\"；"
                         "默认不居中任何标题")
    ap.add_argument("--dry-run", action="store_true", help="只报告，不写文件")
    ap.add_argument("--no-backup", action="store_true",
                    help="覆盖输出文件时不先备份")
    args = ap.parse_args()

    if args.config:
        try:
            args.paper = ps.apply_config(args.config, base=args.paper)
        except Exception as e:
            sys.exit(f"读取版式配置失败：{e}")

    src = os.path.abspath(args.src)
    if not os.path.exists(src):
        sys.exit("找不到文件: " + src)

    doc = Document(src)
    log = []
    centers = [s for s in args.center_titles.replace("，", ",").split(",") if s.strip()]
    scheme = detect_scheme(doc)
    stats = apply_to_document(doc, args.paper, log, center_titles=centers,
                              line_body=args.line_body,
                              line_heading=args.line_heading,
                              columns=args.columns,
                              column_gap_cm=args.column_gap)

    print(f"编号体系: {'数字式 1./1.1/1.1.1（1. 为一级）' if scheme == 'num' else '中文式 一、/（一）/1.（1. 为三级）'}")
    print("=== 段落归类统计 ===")
    for k, v in sorted(stats.items(), key=lambda x: -x[1]):
        print(f"  {k}: {v}")
    if not stats:
        print("  （没有需要改动的段落，可能已经是目标格式）")

    if args.dry_run:
        print("\n--dry-run：未写出文件。前 40 条归类结果：")
        for style, txt in log[:40]:
            print(f"  [{style}] {txt}")
        return

    if not args.out:
        sys.exit("请用 -o 指定输出文件（不会覆盖原文件）")
    out = os.path.abspath(args.out)
    if out == src:
        sys.exit("输出文件不能与原文件相同——为安全起见请另存新文件")
    if os.path.exists(out) and not args.no_backup:
        shutil.copy2(out, out + ".bak")
        print("已备份原输出文件 →", out + ".bak")
    doc.save(out)
    print("\n已生成:", out)


if __name__ == "__main__":
    main()

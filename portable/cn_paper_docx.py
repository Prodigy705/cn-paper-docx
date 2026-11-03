# -*- coding: utf-8 -*-
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


# ======================================================================
# 版式核心：字体、行距、页面、样式、三线表（来自 scripts/paper_style.py）
# ======================================================================

# ---------------------------------------------------------------- 字体常量
SONG = "宋体"            # 正文中文
HEI = "黑体"             # 标题中文
KAI = "楷体"             # a4 版的二级标题
TNR = "Times New Roman"  # 西文与数字

# ---------------------------------------------------------------- 字号
SZ_TITLE = Pt(18)        # 小二（论文大标题）
SZ_H1 = Pt(16)           # 三号
SZ_H2 = Pt(14)           # 四号
SZ_H3 = Pt(12)           # 小四
SZ_BODY = Pt(12)         # 小四
SZ_CAPTION = Pt(12)      # 小四（图题表题用黑体小四）
SZ_NOTE = Pt(10.5)       # 五号（表注）
SZ_REF = Pt(10.5)        # 五号

# ---------------------------------------------------------------- 行距
# 统一表示成 ("auto", twips) 多倍行距 或 ("exact", pt) 固定值。
# 多倍行距：240=单倍，300=1.25倍，360=1.5倍
LINE_SINGLE = ("auto", 240)
LINE_125 = ("auto", 300)
LINE_15 = ("auto", 360)
LINE_EXACT_28 = ("exact", 28)     # 固定值 28 磅


# ================================================================ 版式预设
# ★ 全篇格式只在这里定义 —— 要改字体字号行距，改这两张表即可。
#
# headings 每项：(中文字体, 字号, 是否加粗, 段前pt, 段后pt, 行距)
# 标题一律两端对齐、不加粗（黑体本身够重，再加粗会发糊）。
PROFILES = {
    # ------------------------------------------------ Letter 版（总.docx）
    "letter": {
        "label": "Letter 版（对应「总.docx」）",
        "page": dict(w=Inches(8.5), h=Inches(11),
                     top=Cm(3.5), bottom=Cm(3.5), left=Cm(3.0), right=Cm(3.0)),
        "headings": {
            "Heading 1": (HEI, SZ_H1, False, 12, 6, LINE_15),
            "Heading 2": (HEI, SZ_H2, False, 10, 4, LINE_15),
            "Heading 3": (HEI, SZ_H3, False, 8, 4, LINE_15),
        },
        "h2_indent_chars": 0,      # 二级标题不缩进
        "body_line": LINE_15,      # 正文 1.5 倍行距
        "columns": 1,              # 单栏；2 = 双栏
        "column_gap_cm": 0.75,     # 双栏栏间距
    },
    # ------------------------------------------------ A4 版（第四章　实验流程）
    "a4": {
        "label": "A4 版（对应「第四章　实验流程.docx」）",
        "page": dict(w=Cm(21.0), h=Cm(29.7),
                     top=Cm(3.5), bottom=Cm(3.5),
                     left=Cm(2.8), right=Cm(2.8)),     # 与参考稿一致
        "headings": {
            # 一级与二级同为三号，靠换字体（黑体 → 楷体）+ 缩进区分层级
            "Heading 1": (HEI, SZ_H1, False, 12, 6, LINE_EXACT_28),
            "Heading 2": (KAI, SZ_H1, False, 6, 6, LINE_EXACT_28),
            "Heading 3": (HEI, SZ_H3, False, 6, 3, LINE_EXACT_28),
        },
        "h2_indent_chars": 2,      # 二级标题带首行缩进 2 字符
        "body_line": LINE_EXACT_28,  # 正文固定值 28 磅
        "columns": 1,              # 单栏；2 = 双栏
        "column_gap_cm": 0.75,     # 双栏栏间距
    },
}
DEFAULT_PROFILE = "letter"


def get_profile(name=None):
    """取版式预设；name 为 None 时用默认（letter）。"""
    key = (name or DEFAULT_PROFILE).lower()
    if key not in PROFILES:
        raise KeyError(f"未知版式 {name!r}，可选：{', '.join(PROFILES)}")
    return key, PROFILES[key]


def parse_columns(value):
    """命令行栏数：1/single/单栏，2/double/two/双栏。"""
    s = str(value).strip().lower()
    if s in ("1", "single", "一栏", "单栏"):
        return 1
    if s in ("2", "two", "double", "两栏", "双栏"):
        return 2
    raise ValueError("栏数只能是 1（单栏）或 2（双栏）")


def heading_spec(profile_name, style_name, line_override=None):
    """取某级标题的格式。line_override 可覆盖行距（标题行距独立可调）。"""
    _, prof = get_profile(profile_name)
    ea, size, bold, sb, sa, line = prof["headings"][style_name]
    return ea, size, bold, sb, sa, (line_override or line)


# ================================================================ 行距工具
def parse_line(spec):
    """把命令行传来的行距值解析成 ("auto", twips) / ("exact", pt)。
    接受 "1.5" / "2" （多倍行距）或 "28pt" / "28磅" （固定值）。"""
    if spec is None:
        return None
    if isinstance(spec, tuple):
        return spec
    s = str(spec).strip().lower().replace("磅", "pt").replace("倍", "").replace("行", "")
    if s.endswith("pt"):
        return ("exact", float(s[:-2]))
    return ("auto", float(s) * 240)


def _write_line_spacing(pPr, line):
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    kind, val = line
    if kind == "exact":
        sp.set(qn("w:line"), str(int(round(val * 20))))
        sp.set(qn("w:lineRule"), "exact")
    else:
        sp.set(qn("w:line"), str(int(round(val))))
        sp.set(qn("w:lineRule"), "auto")


def set_line_spacing(paragraph, line):
    """设置段落行距。line 可以是 ('auto',twips)/('exact',pt)，或现成的倍数/磅字符串。"""
    if line is None:
        return
    line = parse_line(line)
    _write_line_spacing(paragraph._p.get_or_add_pPr(), line)


# ================================================================ 角色默认值
# 段落角色（正文、题注、表注、参考文献、步骤行、引文、表格单元格、封面）的
# 默认字体 / 字号 / 行距 / 缩进。版式预设可以用 "roles" 覆盖其中任意一项。
DEFAULT_ROLES = {
    "body": dict(ea=SONG, ascii=TNR, size=SZ_BODY, line=None,
                 indent_chars=2, align="justify", before=2.5, after=0.5),
    "title": dict(ea=HEI, ascii=TNR, size=SZ_TITLE, bold=False,
                  align="center", line=None, before=0, after=12),
    "caption": dict(ea=HEI, ascii=TNR, size=SZ_CAPTION, bold=False,
                    line=LINE_15, indent_chars=0, align="center",
                    before=3, after=6),
    "note": dict(ea=SONG, ascii=TNR, size=SZ_NOTE, line=LINE_SINGLE,
                 indent_chars=2, align="justify", before=3, after=3),
    "reference": dict(ea=SONG, ascii=TNR, size=SZ_REF, line=LINE_SINGLE,
                      indent_chars=0, align="left", before=2.5, after=0),
    "step": dict(ea=SONG, ascii=TNR, size=SZ_BODY, line=None,
                 indent_chars=2, align="left", before=3, after=3,
                 num_ea=HEI),
    "quote": dict(ea=KAI, ascii=TNR, size=SZ_BODY, line=None,
                  indent_chars=0, align="justify", before=0, after=0,
                  left_cm=0.74, right_cm=0.74),
    "table_cell": dict(ea=SONG, ascii=TNR, size=SZ_NOTE, line=LINE_SINGLE,
                       indent_chars=0, align="center", before=1, after=1,
                       bold_header=True),
    "cover_title": dict(ea=HEI, ascii=TNR, size=Pt(22), bold=False,
                        align="center"),
    "cover_subtitle": dict(ea=HEI, ascii=TNR, size=Pt(16), bold=False,
                           align="center"),
    "cover_text": dict(ea=SONG, ascii=TNR, size=Pt(14), bold=False,
                       align="center"),
}

_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def align_of(name):
    """把配置里的对齐方式字符串转成 python-docx 枚举。"""
    if name is None:
        return WD_ALIGN_PARAGRAPH.JUSTIFY
    if isinstance(name, int):
        return name
    return _ALIGN.get(str(name).lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)


def _copy_value(value):
    """深拷贝 dict/list，但保留 Length 等不可变对象原样。

    不能直接 copy.deepcopy(Pt(12))：Length 是 int 子类，deepcopy 会把它当成
    数值重新构造，12pt 会变成 152400pt。
    """
    if isinstance(value, dict):
        return {k: _copy_value(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_copy_value(v) for v in value]
    if isinstance(value, tuple):
        return tuple(_copy_value(v) for v in value)
    return value


def get_roles(profile=None):
    """取某个版式预设的完整角色表（默认值 + 预设里的覆盖）。

    profile 可以是预设名，也可以直接是 PROFILES 里的那个 dict。
    body / step / quote 没显式设行距时，跟随该预设的正文行距。
    """
    if profile is None or isinstance(profile, str):
        _, prof = get_profile(profile)
    else:
        prof = profile
    roles = _copy_value(DEFAULT_ROLES)
    for name, over in (prof.get("roles") or {}).items():
        roles[name] = {**roles.get(name, {}), **_copy_value(over)}
    for name in ("body", "step", "quote"):
        if roles[name].get("line") is None:
            roles[name]["line"] = prof["body_line"]
    return roles


def line_to_text(line):
    """把 ("auto", 360) / ("exact", 28) 转成配置里好写的 "1.5" / "28pt"。"""
    if line is None:
        return None
    kind, val = line
    if kind == "exact":
        return f"{val:g}pt"
    return f"{val / 240:g}"


# ================================================================ 配置文件
# 用户可以导出当前预设为 JSON，改完再用 --config 读回来；也可以直接手写一个
# 精简 JSON，只写要改的项，其余继承 base 预设。
_PAPER_SIZES = {
    "a4": (Cm(21.0), Cm(29.7)),
    "letter": (Inches(8.5), Inches(11)),
}

_ROLE_KEYMAP = {
    "font": "ea",
    "ascii_font": "ascii",
    "size_pt": "size",
    "size": "size",
    "line": "line",
    "before_pt": "before",
    "after_pt": "after",
    "indent_chars": "indent_chars",
    "align": "align",
    "bold": "bold",
    "num_font": "num_ea",
    "left_cm": "left_cm",
    "right_cm": "right_cm",
    "bold_header": "bold_header",
}


def _as_pt(value):
    return value if hasattr(value, "pt") else Pt(float(value))


def _as_float(value):
    return float(value.pt) if hasattr(value, "pt") else float(value)


def _merge_page(cur, cfg):
    page = dict(cur)
    paper = str(cfg.get("paper", "")).lower()
    if paper in _PAPER_SIZES:
        page["w"], page["h"] = _PAPER_SIZES[paper]
    if "width_cm" in cfg:
        page["w"] = Cm(float(cfg["width_cm"]))
    if "height_cm" in cfg:
        page["h"] = Cm(float(cfg["height_cm"]))
    for key in ("top", "bottom", "left", "right"):
        if f"{key}_cm" in cfg:
            page[key] = Cm(float(cfg[f"{key}_cm"]))
    if "header_cm" in cfg:
        page["header"] = Cm(float(cfg["header_cm"]))
    if "footer_cm" in cfg:
        page["footer"] = Cm(float(cfg["footer_cm"]))
    return page


def _merge_role(dst, values):
    for k, v in values.items():
        key = _ROLE_KEYMAP.get(k, k)
        if key == "size":
            v = _as_pt(v)
        elif key == "line":
            v = parse_line(v)
        elif key in ("before", "after", "left_cm", "right_cm"):
            v = _as_float(v)
        elif key == "indent_chars":
            v = int(v)
        elif key == "bold":
            v = bool(v)
        dst[key] = v
    return dst


def apply_config(path, base=None, name=None):
    """读取 JSON 配置，注册成一个新的版式预设，返回预设名。

    典型用法:
        apply_config("我的版式.json")
        Document(); apply_styles(doc, "我的版式") ...
    配置里没写的项继承 base（默认 letter，可写 "base": "a4"）。
    """
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    base_key = cfg.get("base") or base or DEFAULT_PROFILE
    key, prof = get_profile(base_key)
    prof = _copy_value(prof)
    prof["base"] = key
    new_key = name or cfg.get("name") or "custom"
    prof["label"] = cfg.get("label") or f"自定义版式（基于 {key}）"

    if "page" in cfg:
        prof["page"] = _merge_page(prof["page"], cfg["page"])

    if "headings" in cfg:
        for lvl, h in cfg["headings"].items():
            style = f"Heading {lvl}" if str(lvl).isdigit() else str(lvl)
            if style not in prof["headings"]:
                continue
            ea, size, bold, sb, sa, line = prof["headings"][style]
            prof["headings"][style] = (
                h.get("font", ea),
                _as_pt(h["size_pt"]) if "size_pt" in h else size,
                bool(h.get("bold", bold)),
                _as_float(h.get("before_pt", sb)),
                _as_float(h.get("after_pt", sa)),
                parse_line(h["line"]) if "line" in h else line,
            )

    if "body_line" in cfg:
        prof["body_line"] = parse_line(cfg["body_line"])
    if "h2_indent_chars" in cfg:
        prof["h2_indent_chars"] = int(cfg["h2_indent_chars"])
    if "columns" in cfg:
        prof["columns"] = int(cfg["columns"])
    if "column_gap_cm" in cfg:
        prof["column_gap_cm"] = float(cfg["column_gap_cm"])

    if "roles" in cfg:
        roles = prof.setdefault("roles", {})
        for role, values in cfg["roles"].items():
            _merge_role(roles.setdefault(role, {}), values)

    if "table_lines" in cfg:
        prof["table_lines"] = {k: _as_float(v)
                               for k, v in cfg["table_lines"].items()}

    PROFILES[new_key] = prof
    return new_key


def dump_config(profile=DEFAULT_PROFILE, path=None):
    """把某个预设导出成可编辑的 JSON 配置。返回 dict；给了 path 就同时写文件。"""
    key, prof = get_profile(profile)
    roles = get_roles(prof)
    page = prof["page"]
    base = prof.get("base", key)
    name = key if key not in ("letter", "a4") else f"{key}-custom"
    cfg = {
        "name": name,
        "base": base,
        "label": prof["label"],
        "page": {
            "width_cm": round(page["w"].cm, 2),
            "height_cm": round(page["h"].cm, 2),
            "top_cm": round(page["top"].cm, 2),
            "bottom_cm": round(page["bottom"].cm, 2),
            "left_cm": round(page["left"].cm, 2),
            "right_cm": round(page["right"].cm, 2),
            "header_cm": round(page.get("header", Cm(1.27)).cm, 2),
            "footer_cm": round(page.get("footer", Cm(1.27)).cm, 2),
        },
        "body_line": line_to_text(prof["body_line"]),
        "h2_indent_chars": prof["h2_indent_chars"],
        "columns": prof.get("columns", 1),
        "column_gap_cm": prof.get("column_gap_cm", 0.75),
        "headings": {},
        "roles": {},
        "table_lines": dict(prof.get("table_lines",
                                     {"top": 1.5, "header": 0.5, "bottom": 1.5})),
    }
    for style, (ea, size, bold, sb, sa, line) in prof["headings"].items():
        cfg["headings"][style[-1]] = {
            "font": ea,
            "size_pt": round(size.pt, 2),
            "bold": bool(bold),
            "before_pt": sb,
            "after_pt": sa,
            "line": line_to_text(line),
        }
    for rname, role in roles.items():
        item = {}
        for k, v in role.items():
            if v is None:
                continue
            if k == "ea":
                item["font"] = v
            elif k == "ascii":
                item["ascii_font"] = v
            elif k == "size":
                item["size_pt"] = round(v.pt, 2)
            elif k == "line":
                item["line"] = line_to_text(v)
            elif k == "before":
                item["before_pt"] = v
            elif k == "after":
                item["after_pt"] = v
            elif k == "num_ea":
                item["num_font"] = v
            else:
                item[k] = v
        cfg["roles"][rname] = item
    if path:
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
            f.write("\n")
    return cfg


# ================================================================ 字体工具
def set_run_font(run, ascii_font=TNR, ea_font=SONG, size=None,
                 bold=None, italic=None, color=None):
    """设置 run 的中西文字体。python-docx 的 font.name 只管西文，中文必须写
    w:eastAsia，否则 Word 里会回退成默认字体（常见“等线”）。"""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), ea_font)
    rFonts.set(qn("w:cs"), ascii_font)
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def set_style_font(style, ascii_font=TNR, ea_font=SONG, size=None, bold=None):
    """把字体写进样式定义（而不是每个 run），这样用户在 Word 里继续打字也保持格式。"""
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), ea_font)
    rFonts.set(qn("w:cs"), ascii_font)
    if size is not None:
        style.font.size = size
    if bold is not None:
        style.font.bold = bold


# ================================================================ 缩进 / 对齐
def set_first_line_indent_chars(paragraph, chars=2, size_pt=12):
    """首行缩进 N 字符。必须同时写 firstLineChars（字符数×100）和 firstLine
    （等价 twips），否则在关闭“使用字符单位”的 Word 里会失效。"""
    _indent_on(paragraph._p.get_or_add_pPr(), chars, size_pt)


def clear_indent(paragraph):
    set_first_line_indent_chars(paragraph, chars=0)


def _indent_on(pPr, chars, size_pt):
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    if chars:
        ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
        ind.set(qn("w:firstLine"), str(int(chars * size_pt * 20)))
    else:
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")


def set_snap_to_grid(paragraph, off=True):
    """关掉“对齐到网格”，否则多倍行距会被文档网格拉成整行高，间距忽大忽小。"""
    pPr = paragraph._p.get_or_add_pPr()
    el = pPr.find(qn("w:snapToGrid"))
    if el is None:
        el = OxmlElement("w:snapToGrid")
        pPr.append(el)
    el.set(qn("w:val"), "0" if off else "1")


def set_keep_with_next(paragraph, keep=True):
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("w:keepNext", "w:keepLines"):
        el = pPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            pPr.append(el)
        el.set(qn("w:val"), "1" if keep else "0")


# ================================================================ 页面设置
def set_columns(section, num=1, gap_cm=0.75):
    """设置节的栏数：1=单栏，2=双栏。gap_cm 是双栏时的栏间距（厘米）。"""
    num = max(1, int(num))
    sectPr = section._sectPr
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sectPr.append(cols)
    cols.set(qn("w:num"), str(num))
    if num > 1:
        cols.set(qn("w:space"), str(int(round(Cm(float(gap_cm)).twips))))
        cols.set(qn("w:equalWidth"), "1")
    else:
        cols.set(qn("w:space"), "0")
        if cols.get(qn("w:equalWidth")) is not None:
            del cols.attrib[qn("w:equalWidth")]
    return section


def apply_page_setup(doc, paper=DEFAULT_PROFILE, columns=None, column_gap_cm=None):
    """页面与栏设置。columns 传 1/2 覆盖预设；None 用预设里的值。"""
    _, prof = get_profile(paper)
    cfg = prof["page"]
    num = prof.get("columns", 1) if columns is None else int(columns)
    gap = (prof.get("column_gap_cm", 0.75)
           if column_gap_cm is None else float(column_gap_cm))
    for s in doc.sections:
        s.page_width, s.page_height = cfg["w"], cfg["h"]
        s.top_margin, s.bottom_margin = cfg["top"], cfg["bottom"]
        s.left_margin, s.right_margin = cfg["left"], cfg["right"]
        s.header_distance = cfg.get("header", Cm(1.27))
        s.footer_distance = cfg.get("footer", Cm(1.27))
        grid = s._sectPr.find(qn("w:docGrid"))
        if grid is None:
            grid = OxmlElement("w:docGrid")
            s._sectPr.append(grid)
        grid.set(qn("w:linePitch"), "360")
        set_columns(s, num, gap)
    return doc


# ================================================================ 样式定义
def _set_doc_defaults(doc):
    """docDefaults 里指定中文默认字体，避免正文里没写 rFonts 的 run 回退成等线。"""
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is None:
        return
    rPrDefault = dd.find(qn("w:rPrDefault"))
    if rPrDefault is None:
        return
    r = rPrDefault.find(qn("w:rPr"))
    if r is None:
        r = OxmlElement("w:rPr")
        rPrDefault.append(r)
    rFonts = r.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        r.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), TNR)
    rFonts.set(qn("w:hAnsi"), TNR)
    rFonts.set(qn("w:eastAsia"), SONG)


def _get_or_create_style(doc, name):
    try:
        return doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
        return st


def apply_styles(doc, paper=DEFAULT_PROFILE, line_body=None, line_heading=None):
    """在文档里定义全套论文样式。用户打开生成的 docx 继续打字也会沿用。

    line_body / line_heading：可分别覆盖正文行距、标题行距
    （如 "1.5" 表示 1.5 倍，"28pt" 表示固定值 28 磅）。
    """
    key, prof = get_profile(paper)
    roles = get_roles(prof)
    body = roles["body"]
    body_line = parse_line(line_body) or body.get("line") or prof["body_line"]
    head_line = parse_line(line_heading)          # None 表示各标题用预设值
    _set_doc_defaults(doc)

    # ---- 正文 Normal
    normal = doc.styles["Normal"]
    body_size = body.get("size", SZ_BODY)
    set_style_font(normal, body.get("ascii", TNR), body.get("ea", SONG),
                   body_size, bold=False)
    pf = normal.paragraph_format
    pf.alignment = align_of(body.get("align"))
    pf.space_before = Pt(body.get("before", 2.5))
    pf.space_after = Pt(body.get("after", 0.5))
    pf.first_line_indent = Pt(0)
    pPr = normal.element.get_or_add_pPr()
    _indent_on(pPr, body.get("indent_chars", 2), body_size.pt)
    _write_line_spacing(pPr, body_line)
    snap = pPr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        pPr.append(snap)
    snap.set(qn("w:val"), "0")                  # 关网格吸附，保证行距稳定

    # ---- 各级标题（只有三级）
    for name, (ea, size, bold, sb, sa, line) in prof["headings"].items():
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        set_style_font(st, TNR, ea, size, bold=bold)
        st.font.color.rgb = RGBColor(0x00, 0x00, 0x00)   # 清掉 Word 默认的蓝色
        st.font.italic = False
        p = st.paragraph_format
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.space_before = Pt(sb)
        p.space_after = Pt(sa)
        p.first_line_indent = Pt(0)
        pPr = st.element.get_or_add_pPr()
        # 二级标题在 a4 版里带首行缩进
        chars = prof["h2_indent_chars"] if name == "Heading 2" else 0
        _indent_on(pPr, chars, size.pt)
        _write_line_spacing(pPr, head_line or line)
        keep = pPr.find(qn("w:keepNext"))
        if keep is None:
            keep = OxmlElement("w:keepNext")
            pPr.insert(0, keep)

    # 本规范只有三级标题，Heading 4 不定义 —— apply_format.py 会把原稿里的
    # 四级标题降为正文，new_paper.py 遇到 #### 直接报错。

    # ---- 图题 / 表题 Caption
    cap_role = roles["caption"]
    caption = doc.styles["Caption"]
    cap_size = cap_role.get("size", SZ_CAPTION)
    set_style_font(caption, cap_role.get("ascii", TNR),
                   cap_role.get("ea", HEI), cap_size,
                   bold=bool(cap_role.get("bold", False)))
    cp = caption.paragraph_format
    cp.alignment = align_of(cap_role.get("align", "center"))
    cp.space_before = Pt(cap_role.get("before", 3))
    cp.space_after = Pt(cap_role.get("after", 6))
    cp.first_line_indent = Pt(0)
    _indent_on(caption.element.get_or_add_pPr(),
               cap_role.get("indent_chars", 0), cap_size.pt)
    _write_line_spacing(caption.element.get_or_add_pPr(),
                        cap_role.get("line") or LINE_15)

    # ---- 表注 TableNote
    note_role = roles["note"]
    note = _get_or_create_style(doc, "TableNote")
    note_size = note_role.get("size", SZ_NOTE)
    set_style_font(note, note_role.get("ascii", TNR),
                   note_role.get("ea", SONG), note_size, bold=False)
    npf = note.paragraph_format
    npf.alignment = align_of(note_role.get("align"))
    npf.space_before = Pt(note_role.get("before", 3))
    npf.space_after = Pt(note_role.get("after", 3))
    npf.first_line_indent = Pt(0)
    _indent_on(note.element.get_or_add_pPr(),
               note_role.get("indent_chars", 2), note_size.pt)
    _write_line_spacing(note.element.get_or_add_pPr(),
                        note_role.get("line") or LINE_SINGLE)

    # ---- 参考文献 Reference
    ref_role = roles["reference"]
    ref = _get_or_create_style(doc, "Reference")
    ref_size = ref_role.get("size", SZ_REF)
    set_style_font(ref, ref_role.get("ascii", TNR),
                   ref_role.get("ea", SONG), ref_size, bold=False)
    rpf = ref.paragraph_format
    rpf.alignment = align_of(ref_role.get("align"))
    rpf.space_before = Pt(ref_role.get("before", 2.5))
    rpf.space_after = Pt(ref_role.get("after", 0))
    rpf.first_line_indent = Pt(0)
    _indent_on(ref.element.get_or_add_pPr(),
               ref_role.get("indent_chars", 0), ref_size.pt)
    _write_line_spacing(ref.element.get_or_add_pPr(),
                        ref_role.get("line") or LINE_SINGLE)

    # ---- 步骤行 Step
    step_role = roles["step"]
    step = _get_or_create_style(doc, "Step")
    step_size = step_role.get("size", SZ_BODY)
    set_style_font(step, step_role.get("ascii", TNR),
                   step_role.get("ea", SONG), step_size, bold=False)
    spf = step.paragraph_format
    spf.alignment = align_of(step_role.get("align"))
    spf.space_before = Pt(step_role.get("before", 3))
    spf.space_after = Pt(step_role.get("after", 3))
    spf.first_line_indent = Pt(0)
    _indent_on(step.element.get_or_add_pPr(),
               step_role.get("indent_chars", 2), step_size.pt)
    step_line = body_line if line_body else (step_role.get("line") or body_line)
    _write_line_spacing(step.element.get_or_add_pPr(), step_line)

    # ---- 引文 Quote-CN
    quote_role = roles["quote"]
    quote = _get_or_create_style(doc, "Quote-CN")
    quote_size = quote_role.get("size", SZ_BODY)
    set_style_font(quote, quote_role.get("ascii", TNR),
                   quote_role.get("ea", KAI), quote_size, bold=False)
    qpf = quote.paragraph_format
    qpf.alignment = align_of(quote_role.get("align"))
    qpf.left_indent = Cm(float(quote_role.get("left_cm", 0.74)))
    qpf.right_indent = Cm(float(quote_role.get("right_cm", 0.74)))
    qpf.first_line_indent = Pt(0)
    _indent_on(quote.element.get_or_add_pPr(),
               quote_role.get("indent_chars", 0), quote_size.pt)
    quote_line = body_line if line_body else (quote_role.get("line") or body_line)
    _write_line_spacing(quote.element.get_or_add_pPr(), quote_line)

    return doc


# ================================================================ 整篇文档
# 下面几个是“从零写完一篇完整文档”需要的零件：封面、目录、页码、插图。

def content_width(doc):
    """当前节正文可用宽度（EMU）。双栏时返回单栏宽度，插图就不会跨栏。"""
    s = doc.sections[-1]
    width = int(s.page_width - s.left_margin - s.right_margin)
    cols = s._sectPr.find(qn("w:cols"))
    num = int(cols.get(qn("w:num")) or 1) if cols is not None else 1
    if num <= 1:
        return width
    space = int(cols.get(qn("w:space")) or 0)
    return max(1, (width - space * 635 * (num - 1)) // num)


def fit_table_width(table, width_emu):
    """把表格总宽设成给定宽度，并按列数均分。

    双栏排版时表格默认可能比一栏宽，调用这个把表格收进当前栏。
    """
    width_emu = int(width_emu)
    grid = table._tbl.tblGrid
    cols = grid.findall(qn("w:gridCol")) if grid is not None else []
    n = max(1, len(cols))
    each = width_emu // n
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(round(width_emu / 635))))
    table.autofit = False
    for c in cols:
        c.set(qn("w:w"), str(int(round(each / 635))))
    for row in table.rows:
        for cell in row.cells:
            cell.width = Emu(each)
    return table


def add_picture_fitted(paragraph, path, max_width=None, dpi=96.0):
    """按页面可用宽度插入图片：比版面宽就等比缩到版面宽，否则用原始尺寸。

    返回 (宽EMU, 高EMU)；找不到文件或非图片时返回 None。
    """
    if not os.path.exists(path):
        return None
    try:
        from PIL import Image
        with Image.open(path) as im:
            px_w, px_h = im.size
        emu = 914400 / dpi
        nat_w = px_w * emu
        if max_width and nat_w > max_width:
            ratio = max_width / nat_w
            paragraph.add_run().add_picture(path, width=Emu(int(nat_w * ratio)))
            return int(nat_w * ratio), int(px_h * emu * ratio)
        paragraph.add_run().add_picture(path)
        return int(nat_w), int(px_h * emu)
    except Exception:
        # 没有 Pillow 或读不出尺寸：交给 Word 按原始尺寸放
        try:
            paragraph.add_run().add_picture(path)
            return 0, 0
        except Exception:
            return None


def add_toc_field(paragraph, levels="1-3", placeholder="（按 F9 更新目录）"):
    """插入自动目录域。

    注意：**不要**写 w:dirty —— 那会让 Word 每次打开文档都强制刷新目录，
    用户按“更新域”存了一次也白搭，下次打开又刷新。这里不写 dirty，
    目录只在用户主动按 F9（或运行时用 refresh_fields_with_word 刷新）时才变。
    """
    r = paragraph.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    r._r.append(fld)
    r2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' TOC \\o "{levels}" \\h \\z \\u '
    r2._r.append(instr)
    r3 = paragraph.add_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r3._r.append(sep)
    r4 = paragraph.add_run(placeholder)
    set_run_font(r4, TNR, SONG, SZ_NOTE, bold=False)
    r5 = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r5._r.append(end)
    return paragraph


def refresh_fields_with_word(path, timeout=180):
    """用 Word 打开文件、刷新一次所有域（目录页码），再存回去。

    生成的文档一打开就是带正确页码的目录，之后不会再提示更新 —— 因为
    add_toc_field 不写 w:dirty，Word 不会每次打开都强制刷新。

    返回 (是否成功, 说明)。没装 Word 或刷新失败返回 (False, 原因)，
    不影响已生成的文档（用户手动按一次 F9 效果相同）。
    """
    import subprocess
    if os.name != "nt":
        # ChatGPT / Grok / Linux 沙箱里没有 Word，直接降级，避免无谓等待
        return False, "当前环境不是 Windows，无法调用 Word；请在 Word 里按 F9 更新目录"
    full = os.path.abspath(path).replace("\\", "/")
    ps = (
        "$ErrorActionPreference='Stop';"
        "$w = New-Object -ComObject Word.Application;"
        "$w.Visible=$false; $w.DisplayAlerts=0;"
        "try {"
        f"$d = $w.Documents.Open('{full}', $false, $false);"
        "$d.Fields.Update() | Out-Null;"
        "foreach ($t in $d.TablesOfContents) { $t.Update() };"
        "$d.Save(); $d.Close(0); Write-Output 'OK'"
        "} finally { $w.Quit() }"
    )
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True, timeout=timeout)
        if "OK" in (r.stdout or ""):
            return True, "目录已刷新"
        return False, (r.stderr or r.stdout or "未知错误").strip()[:200]
    except FileNotFoundError:
        return False, "没找到 PowerShell"
    except subprocess.TimeoutExpired:
        return False, "刷新超时（Word 可能弹了对话框）"
    except Exception as e:
        return False, str(e)[:200]


def start_numbered_section(doc, restart=1):
    """另起一新页作为新节，并从 restart 开始重新编页码，返回该节。

    用于“封面和目录不编页码、正文从 1 开始”——新节会继承上一节的纸张与页边距。
    restart 传 None 表示接着上一节继续编号。
    """
    from docx.enum.section import WD_SECTION
    s = doc.add_section(WD_SECTION.NEW_PAGE)
    if restart is not None:
        sectPr = s._sectPr
        for old in sectPr.findall(qn("w:pgNumType")):
            sectPr.remove(old)
        el = OxmlElement("w:pgNumType")
        el.set(qn("w:start"), str(int(restart)))
        sectPr.append(el)
    return s


def add_page_number_footer(doc, sections=None, ea_font=SONG, size=SZ_NOTE,
                           prefix="", suffix=""):
    """给指定节的页脚加上居中页码（PAGE 域）。sections 为空则处理全部节。"""
    for s in (sections if sections is not None else doc.sections):
        ftr = s.footer
        ftr.is_linked_to_previous = False
        p = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clear_indent(p)
        _write_line_spacing(p._p.get_or_add_pPr(), LINE_SINGLE)
        if prefix:
            set_run_font(p.add_run(prefix), TNR, ea_font, size, bold=False)
        r = p.add_run()
        b = OxmlElement("w:fldChar")
        b.set(qn("w:fldCharType"), "begin")
        r._r.append(b)
        r2 = p.add_run()
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = " PAGE "
        r2._r.append(it)
        r3 = p.add_run()
        sp = OxmlElement("w:fldChar")
        sp.set(qn("w:fldCharType"), "separate")
        r3._r.append(sp)
        set_run_font(p.add_run("1"), TNR, ea_font, size, bold=False)
        r5 = p.add_run()
        e = OxmlElement("w:fldChar")
        e.set(qn("w:fldCharType"), "end")
        r5._r.append(e)
        for rr in p.runs:
            if rr.font.size is None:
                set_run_font(rr, TNR, ea_font, size, bold=False)
        if suffix:
            set_run_font(p.add_run(suffix), TNR, ea_font, size, bold=False)
    return doc


def add_centered_line(doc, text, ea_font=SONG, size=SZ_BODY, bold=False,
                      space_before=None, space_after=None, line=None):
    """加一行居中文字（封面/扉页用）。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clear_indent(p)
    if line:
        set_line_spacing(p, line)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    set_run_font(p.add_run(text), TNR, ea_font, size, bold=bold)
    return p


# ================================================================ 三线表
def _border_sz(pt):
    """把磅值换成 OOXML 的 1/8 磅单位（1.5pt → 12）。"""
    return str(int(round(float(pt) * 8)))


def make_three_line_table(table, header_rows=1, lines=None, cell_role=None):
    """把表格改成学术三线表。

    lines: 线宽（磅），可选 {"top": 1.5, "header": 0.5, "bottom": 1.5}
    cell_role: 单元格角色表（默认 DEFAULT_ROLES["table_cell"]）
    """
    lines = lines or {}
    top_sz = _border_sz(lines.get("top", 1.5))
    header_sz = _border_sz(lines.get("header", 0.5))
    bottom_sz = _border_sz(lines.get("bottom", 1.5))
    cell_role = cell_role or DEFAULT_ROLES["table_cell"]

    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for tag, sz, val in (("top", top_sz, "single"),
                         ("bottom", bottom_sz, "single"),
                         ("left", "0", "none"), ("right", "0", "none"),
                         ("insideH", "0", "none"), ("insideV", "0", "none")):
        el = OxmlElement("w:" + tag)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    n = len(table.rows)
    for ri, row in enumerate(table.rows):
        if ri < header_rows:                     # 表头跨页重复
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:tblHeader")) is None:
                th = OxmlElement("w:tblHeader")
                th.set(qn("w:val"), "true")
                trPr.append(th)
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(old)
            tcB = OxmlElement("w:tcBorders")
            for tag, on, sz in (("top", ri == 0, top_sz),
                                ("bottom", ri == header_rows - 1 or ri == n - 1,
                                 bottom_sz if ri == n - 1 else header_sz),
                                ("left", False, "0"), ("right", False, "0")):
                el = OxmlElement("w:" + tag)
                el.set(qn("w:val"), "single" if on else "nil")
                el.set(qn("w:sz"), sz)
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), "000000")
                tcB.append(el)
            tcPr.append(tcB)
            for old in tcPr.findall(qn("w:tcMar")):
                tcPr.remove(old)
            mar = OxmlElement("w:tcMar")
            for tag, w in (("top", "70"), ("left", "90"),
                           ("bottom", "70"), ("right", "90")):
                el = OxmlElement("w:" + tag)
                el.set(qn("w:w"), w)
                el.set(qn("w:type"), "dxa")
                mar.append(el)
            tcPr.append(mar)
            bold = ri < header_rows and cell_role.get("bold_header", True)
            _format_table_cell(cell, bold=bold, role=cell_role)
    return table


def _format_table_cell(cell, bold=False, role=None):
    role = role or DEFAULT_ROLES["table_cell"]
    for p in cell.paragraphs:
        p.paragraph_format.alignment = align_of(role.get("align", "center"))
        p.paragraph_format.space_before = Pt(role.get("before", 1))
        p.paragraph_format.space_after = Pt(role.get("after", 1))
        _indent_on(p._p.get_or_add_pPr(), role.get("indent_chars", 0),
                   role.get("size", SZ_NOTE).pt)
        set_line_spacing(p, role.get("line") or LINE_SINGLE)
        for r in p.runs:
            set_run_font(r, role.get("ascii", TNR), role.get("ea", SONG),
                         r.font.size or role.get("size", SZ_NOTE), bold=bold)


# ======================================================================
# 从零生成：解析标记语法，生成完整文档（来自 scripts/new_paper.py）
# ======================================================================

# 标题默认左对齐。需要居中的标题（如参考文献、致谢）在 # 后加 ! 显式标记，
# 例如「#! 参考文献」——不按章节名字猜，任何标题都能用。


def add_rich_text(paragraph, text, ea_font=SONG, size=SZ_BODY,
                  base_bold=False):
    """把 `**粗体**` 拆成多个 run 写进段落。"""
    for i, seg in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not seg:
            continue
        r = paragraph.add_run(seg)
        set_run_font(r, TNR, ea_font, size,
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
    role = role or DEFAULT_ROLES["caption"]
    p = doc.add_paragraph(style="Caption")
    clear_indent(p)
    add_rich_text(p, text, role.get("ea", HEI),
                  role.get("size", SZ_CAPTION),
                  base_bold=bool(role.get("bold", False)))
    return p


def _picture(doc, path, base_dir=None):
    """插入一张居中图片，超过版面宽度自动等比缩小。"""
    full = path
    if base_dir and not os.path.isabs(path):
        full = os.path.join(base_dir, path)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clear_indent(p)
    set_line_spacing(p, LINE_SINGLE)
    got = add_picture_fitted(p, full, content_width(doc))
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

    roles = roles or get_roles(paper)
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
    add_centered_line(doc, "", line=line_heading)          # 顶部留白
    add_centered_line(doc, "", line=line_heading)
    for k, v in fields:
        role, sb, sa = title_font.get(k, (cover_text, 4, 4))
        add_centered_line(doc, v,
                             ea_font=role.get("ea", SONG),
                             size=role.get("size", Pt(14)),
                             bold=bool(role.get("bold", False)),
                             space_before=role.get("before", sb),
                             space_after=role.get("after", sa),
                             line=line_heading)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    return i


def render(doc, text, paper=DEFAULT_PROFILE, line_heading=None,
           base_dir=None, columns=None, column_gap_cm=None):
    lines = text.splitlines()
    i = 0
    buf = []          # 待合并的正文行
    _, prof = get_profile(paper)
    body_cols = prof.get("columns", 1) if columns is None else int(columns)
    body_gap = (prof.get("column_gap_cm", 0.75)
                if column_gap_cm is None else float(column_gap_cm))
    roles = get_roles(prof)
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
            set_first_line_indent_chars(
                p, body.get("indent_chars", 2),
                body.get("size", SZ_BODY).pt)
            add_rich_text(p, join_wrapped(buf),
                          body.get("ea", SONG),
                          body.get("size", SZ_BODY),
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
            ea, size, bold, _sb, _sa, _l = heading_spec(paper, "Heading 1",
                                                           line_heading)
            p = doc.add_paragraph(style="Heading 1")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clear_indent(p)
            set_run_font(p.add_run(title), TNR, ea, size, bold=bold)
            tp = doc.add_paragraph()
            clear_indent(tp)
            add_toc_field(tp)
            continue

        # ---- 页码：从此处另起一节，并从这里开始编号
        #      @pagenum          → 另起一页，页码从 1 开始（封面/目录不编码）
        #      @pagenum continue → 另起一页，页码接着往下编
        if line.startswith("@pagenum"):
            arg = line[len("@pagenum"):].strip().lower()
            restart = None if arg in ("continue", "续", "接续") else (
                int(arg) if arg.isdigit() else 1)
            sec = start_numbered_section(doc, restart=restart)
            set_columns(sec, body_cols, body_gap)
            add_page_number_footer(doc, sections=[sec],
                                      ea_font=note_role.get("ea", SONG),
                                      size=note_role.get("size", SZ_NOTE))
            continue

        # ---- 指令
        if line.startswith("@title"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clear_indent(p)
            p.paragraph_format.space_before = Pt(title_role.get("before", 0))
            p.paragraph_format.space_after = Pt(title_role.get("after", 12))
            set_line_spacing(p, line_heading or title_role.get("line")
                                or LINE_15)
            add_rich_text(p, line[len("@title"):].strip(),
                          title_role.get("ea", HEI),
                          title_role.get("size", SZ_TITLE),
                          base_bold=bool(title_role.get("bold", False)))
            continue
        if line.startswith("@center"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clear_indent(p)
            add_rich_text(p, line[len("@center"):].strip(),
                          body.get("ea", SONG),
                          body.get("size", SZ_BODY))
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
            ea, size, bold, _sb, _sa, _line = heading_spec(
                paper, style, line_heading)
            p = doc.add_paragraph(style=style)
            clear_indent(p)
            if lvl == 2:
                prof = get_profile(paper)[1]
                chars = prof["h2_indent_chars"]
                if chars:
                    set_first_line_indent_chars(p, chars, size.pt)
            add_rich_text(p, txt, ea, size, base_bold=bold)
            if centered:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

        # ---- 步骤行
        if re.match(r"^[-*]\s+", line):
            p = doc.add_paragraph(style="Step")
            set_first_line_indent_chars(
                p, step_role.get("indent_chars", 2),
                step_role.get("size", SZ_BODY).pt)
            txt = re.sub(r"^[-*]\s+", "", line)
            m2 = re.match(r"^((?:Step|步骤)\s*\d+|\d+)\s*[：:]\s*(.*)$", txt)
            if m2:
                r = p.add_run(m2.group(1) + "：")
                set_run_font(r, TNR,
                                step_role.get("num_ea", HEI),
                                step_role.get("size", SZ_BODY), bold=False)
                add_rich_text(p, m2.group(2),
                              step_role.get("ea", SONG),
                              step_role.get("size", SZ_BODY))
            else:
                add_rich_text(p, txt, step_role.get("ea", SONG),
                              step_role.get("size", SZ_BODY))
            continue

        # ---- 表注
        if line.startswith(">"):
            p = doc.add_paragraph(style="TableNote")
            set_first_line_indent_chars(
                p, note_role.get("indent_chars", 2),
                note_role.get("size", SZ_NOTE).pt)
            add_rich_text(p, line.lstrip("> ").strip(),
                          note_role.get("ea", SONG),
                          note_role.get("size", SZ_NOTE))
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
            clear_indent(p)
            add_rich_text(p, line[3:].strip(),
                          ref_role.get("ea", SONG),
                          ref_role.get("size", SZ_REF))
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
                        clear_indent(p)
                        add_rich_text(p, row[ci] if ci < len(row) else "",
                                      cell_role.get("ea", SONG),
                                      cell_role.get("size", SZ_NOTE))
                fit_table_width(tbl, content_width(doc))
                make_three_line_table(
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


# ======================================================================
# 给已有 docx 套格式：识别段落角色并重排（来自 scripts/apply_format.py）
# ======================================================================

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
def apply_to_document(doc, paper=DEFAULT_PROFILE, log=None, center_titles=(),
                      line_body=None, line_heading=None,
                      columns=None, column_gap_cm=None):
    """center_titles: 需要居中的标题文字集合（如 {"参考文献", "致谢"}）。
    默认为空 —— 标题一律两端对齐，不按章节名做猜测。"""
    stats = {}
    center_titles = {t.strip().rstrip("：:") for t in center_titles if t.strip()}
    apply_page_setup(doc, paper, columns=columns,
                        column_gap_cm=column_gap_cm)
    apply_styles(doc, paper, line_body, line_heading)

    _, prof = get_profile(paper)
    body_cols = prof.get("columns", 1) if columns is None else int(columns)
    roles = get_roles(prof)
    role_spec = _role_spec(prof, line_body)
    scheme = detect_scheme(doc)

    for p in doc.paragraphs:
        text = p.text.strip()
        cur = p.style.name

        if not text:
            if has_drawing(p):          # 纯图片段落：居中、不缩进
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                clear_indent(p)
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
            fit_table_width(t, content_width(doc))
        make_three_line_table(t, header_rows=1,
                                 lines=prof.get("table_lines"),
                                 cell_role=roles["table_cell"])
        stats["表格转三线表"] = stats.get("表格转三线表", 0) + 1

    return stats


def _role_spec(prof, line_body=None):
    """按版式预设的角色表生成各段落样式的 (字体, 西文字体, 字号, 行距,
    缩进字符, 对齐)。body/step/quote 的正文行距可被 --line-body 覆盖。"""
    roles = get_roles(prof)
    override = parse_line(line_body)

    def spec(name, use_body_override=False):
        r = roles[name]
        line = override if (use_body_override and override) else r.get("line")
        return (r.get("ea", SONG), r.get("ascii", TNR),
                r.get("size", SZ_BODY), line,
                r.get("indent_chars", 0),
                align_of(r.get("align")))

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
    key, prof = get_profile(paper)
    if style_name.startswith("Heading"):
        name = style_name if style_name in prof["headings"] else "Heading 3"
        ea, size, _bold, _sb, _sa, line = heading_spec(paper, name, line_heading)
        ascii_font = TNR
        indent = prof["h2_indent_chars"] if name == "Heading 2" else 0
        align = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        if role_spec is None:
            role_spec = _role_spec(prof, line_body)
        ea, ascii_font, size, line, indent, align = role_spec.get(
            style_name, role_spec["Normal"])
        if line is None:
            line = parse_line(line_body) or prof["body_line"]

    if align is not None:
        p.alignment = align
    if indent:
        set_first_line_indent_chars(p, indent, size.pt)
    else:
        clear_indent(p)
    set_line_spacing(p, line)

    for r in p.runs:
        # 只补字体，不覆盖用户自己设的加粗/颜色；没设字号的按样式给。
        # 例外：标题一律不加粗（本规范硬约束），原稿里的加粗要清掉。
        bold = False if style_name.startswith("Heading") else r.font.bold
        set_run_font(r, ascii_font, ea, r.font.size or size, bold=bold)


# ======================================================================
# 生成空白模板与样式样张（来自 scripts/make_template.py）
# ======================================================================

def build_template(paper, line_body=None, line_heading=None,
                   columns=None, column_gap_cm=None):
    doc = Document()
    apply_page_setup(doc, paper, columns=columns,
                        column_gap_cm=column_gap_cm)
    apply_styles(doc, paper, line_body, line_heading)
    for p in list(doc.paragraphs):          # 清掉 python-docx 默认模板的空段落
        p._element.getparent().remove(p._element)
    p = doc.add_paragraph(style="Normal")
    set_first_line_indent_chars(p, 2, 12)
    return doc


def build_specimen(paper, line_body=None, line_heading=None,
                   columns=None, column_gap_cm=None):
    """样张里的每行说明都从版式预设里读出来，改预设这里自动跟着变。"""
    _, prof = get_profile(paper)
    roles = get_roles(prof)
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
    apply_page_setup(doc, paper, columns=columns,
                        column_gap_cm=column_gap_cm)
    apply_styles(doc, paper, line_body, line_heading)

    def note(p, text):
        """在段落尾部加一行楷体小字的说明，标注该段用的是哪个样式。"""
        r = p.add_run("　　← " + text)
        set_run_font(r, TNR, KAI, Pt(9), bold=False)

    # 大标题
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clear_indent(t)
    set_line_spacing(t, line_heading or title.get("line") or LINE_15)
    set_run_font(t.add_run("论文版式样张"), TNR,
                    title.get("ea", HEI), title.get("size", SZ_TITLE),
                    bold=bool(title.get("bold", False)))
    note(t, prof["label"])

    # 三级标题，逐级展示
    labels = {
        "Heading 1": "一级标题",
        "Heading 2": "二级标题",
        "Heading 3": "三级标题",
    }
    for i, (style, label) in enumerate(labels.items(), 1):
        ea, size, bold, _sb, _sa, _line = heading_spec(paper, style, line_heading)
        p = doc.add_paragraph(style=style)
        clear_indent(p)
        if style == "Heading 2" and prof["h2_indent_chars"]:
            set_first_line_indent_chars(p, prof["h2_indent_chars"], size.pt)
        extra = "，首行缩进2字符" if (style == "Heading 2"
                                     and prof["h2_indent_chars"]) else ""
        prefix = {1: "一、", 2: "（一）", 3: "1."}[i]
        txt = prefix + label
        set_run_font(p.add_run(txt), TNR, ea, size, bold=bold)
        note(p, f"Heading {i}｜{ea} {size.pt}pt{'加粗' if bold else '不加粗'}{extra}")

    # 正文
    p = doc.add_paragraph(style="Normal")
    set_first_line_indent_chars(p, body.get("indent_chars", 2),
                                   body.get("size", SZ_BODY).pt)
    set_run_font(p.add_run(
        f"正文：{body.get('ea', SONG)} {body.get('size', SZ_BODY).pt:g}pt，"
        f"西文 Times New Roman，首行缩进 {body.get('indent_chars', 2)} 字符，两端对齐。"
        "这是正文段落的样子，用来核对行距与缩进是否与论文一致。"),
        body.get("ascii", TNR), body.get("ea", SONG),
        body.get("size", SZ_BODY))
    note(p, "Normal")

    p = doc.add_paragraph(style="Normal")
    set_first_line_indent_chars(p, body.get("indent_chars", 2),
                                   body.get("size", SZ_BODY).pt)
    set_run_font(p.add_run("三级标题之后就只有正文了，本规范不设四级标题。"),
                    body.get("ascii", TNR), body.get("ea", SONG),
                    body.get("size", SZ_BODY))
    note(p, "Normal")

    # 步骤行
    for txt in (f"Step 1：步骤行——左对齐、首行缩进 {step.get('indent_chars', 2)} 字符，"
                f"序号用{step.get('num_ea', HEI)}，内容用{step.get('ea', SONG)}。",
                "Step 2：实验流程、操作手册、算法步骤这类“1、2、3…”都用此样式。"):
        p = doc.add_paragraph(style="Step")
        set_first_line_indent_chars(p, step.get("indent_chars", 2),
                                       step.get("size", SZ_BODY).pt)
        num, _, rest = txt.partition("：")
        set_run_font(p.add_run(num + "："), TNR,
                        step.get("num_ea", HEI),
                        step.get("size", SZ_BODY), bold=False)
        set_run_font(p.add_run(rest), step.get("ascii", TNR),
                        step.get("ea", SONG), step.get("size", SZ_BODY))
        note(p, "Step")

    # 图题 / 表题
    for txt in (f"图1-1　图题：{cap.get('ea', HEI)} "
                f"{cap.get('size', SZ_CAPTION).pt:g}pt、居中、不缩进",
                "表1-1　表题：同样用 Caption 样式，置于表格上方"):
        p = doc.add_paragraph(style="Caption")
        clear_indent(p)
        set_run_font(p.add_run(txt), cap.get("ascii", TNR),
                        cap.get("ea", HEI), cap.get("size", SZ_CAPTION))
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
    fit_table_width(tbl, content_width(doc))
    make_three_line_table(tbl, header_rows=1, lines=tbl_lines,
                             cell_role=cell)
    note(doc.add_paragraph(style="TableNote"),
         f"三线表：顶线{tbl_lines.get('top', 1.5):g}pt／"
         f"表头下{tbl_lines.get('header', 0.5):g}pt／"
         f"底线{tbl_lines.get('bottom', 1.5):g}pt")

    n = doc.add_paragraph(style="TableNote")
    set_first_line_indent_chars(n, note_r.get("indent_chars", 2),
                                   note_r.get("size", SZ_NOTE).pt)
    set_run_font(n.add_run(
        f"注：粗体表示该列最高观察结果。表注用{note_r.get('ea', SONG)} "
        f"{note_r.get('size', SZ_NOTE).pt:g}pt。"),
        note_r.get("ascii", TNR), note_r.get("ea", SONG),
        note_r.get("size", SZ_NOTE))
    note(n, "TableNote")

    # 引文
    p = doc.add_paragraph(style="Quote-CN")
    set_run_font(p.add_run(
        f"引文段落用 Quote-CN 样式：{quote_r.get('ea', KAI)} "
        f"{quote_r.get('size', SZ_BODY).pt:g}pt、左右各缩进 "
        f"{quote_r.get('left_cm', 0.74):g}cm。"),
        quote_r.get("ascii", TNR), quote_r.get("ea", KAI),
        quote_r.get("size", SZ_BODY))
    note(p, "Quote-CN")

    # 居中标题 + 参考文献
    ea, size, bold, _sb, _sa, _line = heading_spec(paper, "Heading 1", line_heading)
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clear_indent(p)
    set_run_font(p.add_run("参考文献"), TNR, ea, size, bold=bold)
    note(p, "居中标题：用 new 命令时写「#! 参考文献」")

    for txt in ["PFURTSCHELLER G, LOPES DA SILVA F H. Motor imagery and direct "
                "brain-computer communication[J]. Proceedings of the IEEE, 2001, 89(7): 1123-1134.",
                "LAWHERN V J, SOLON A J, WAYTOWICH N R, et al. EEGNet: a compact "
                "convolutional neural network for EEG-based brain-computer interfaces[J]. "
                "Journal of Neural Engineering, 2018, 15(5): 056013."]:
        p = doc.add_paragraph(style="Reference")
        clear_indent(p)
        set_run_font(p.add_run(txt), ref.get("ascii", TNR),
                        ref.get("ea", SONG),
                        ref.get("size", SZ_REF), bold=False)
    return doc


# ================================================================ 便携版 CLI
# 这一文件由 build_portable.py 追加到 cn_paper_docx.py 末尾。
# 上面合并进来的函数都来自 scripts/，这里只做参数解析、核对与打印规范。


def _ensure_parent(path):
    d = os.path.dirname(os.path.abspath(path))
    if d:
        os.makedirs(d, exist_ok=True)


def _load_config_arg(args):
    """有 --config 就加载并注册成预设，把 args.paper 指向它。"""
    path = getattr(args, "config", None)
    if not path:
        return
    try:
        args.paper = apply_config(path, base=getattr(args, "paper", None))
    except Exception as e:
        sys.exit(f"读取版式配置失败：{e}")


def cmd_new(args):
    _load_config_arg(args)
    if getattr(args, "demo", False):
        text = DEMO
    else:
        if not args.src:
            sys.exit("请给出稿件路径，或用 --demo（demo 子命令也可以）")
        with io.open(args.src, encoding="utf-8") as f:
            text = f.read()

    _, prof = get_profile(args.paper)
    body_cols = (prof.get("columns", 1)
                 if getattr(args, "columns", None) is None else args.columns)
    body_gap = (prof.get("column_gap_cm", 0.75)
                if getattr(args, "column_gap", None) is None else args.column_gap)
    # 有 @pagenum 时封面/目录留在单栏节，正文从新节开始分栏
    initial_cols = 1 if (body_cols == 2 and "@pagenum" in text) else body_cols

    line_head = parse_line(args.line_heading)
    doc = Document()
    apply_page_setup(doc, args.paper, columns=initial_cols,
                     column_gap_cm=body_gap)
    apply_styles(doc, args.paper, args.line_body, args.line_heading)
    for p in list(doc.paragraphs):          # 清掉 python-docx 默认模板的空段落
        p._element.getparent().remove(p._element)

    base_dir = os.path.dirname(os.path.abspath(args.src)) if args.src else os.getcwd()
    render(doc, text, args.paper, line_head, base_dir,
           columns=body_cols, column_gap_cm=body_gap)
    _ensure_parent(args.out)
    doc.save(args.out)
    print("已生成:", args.out)

    if "@toc" in text and not getattr(args, "no_toc_refresh", False):
        ok, msg = refresh_fields_with_word(args.out)
        if ok:
            print("目录已刷新（打开即为带页码的目录，之后不会再提示更新）")
        else:
            print(f"跳过目录刷新（{msg}）——在 Word 里按一次 F9 即可")


def cmd_apply(args):
    _load_config_arg(args)
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
                              columns=getattr(args, "columns", None),
                              column_gap_cm=getattr(args, "column_gap", None))

    print("编号体系: " + ("数字式 1./1.1/1.1.1（1. 为一级）"
                         if scheme == "num" else
                         "中文式 一、/（一）/1.（1. 为三级）"))
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
    _ensure_parent(out)
    doc.save(out)
    print("\n已生成:", out)


def cmd_template(args):
    _load_config_arg(args)
    _ensure_parent(args.out)
    cols = getattr(args, "columns", None)
    gap = getattr(args, "column_gap", None)
    build_template(args.paper, args.line_body, args.line_heading,
                   cols, gap).save(args.out)
    print("已生成模板:", args.out)
    if args.specimen:
        _ensure_parent(args.specimen)
        build_specimen(args.paper, args.line_body, args.line_heading,
                       cols, gap).save(args.specimen)
        print("已生成样张:", args.specimen)


# ---------------------------------------------------------------- 核对
def _style(doc, name):
    try:
        return doc.styles[name]
    except KeyError:
        return None


def _ea_font(style):
    rPr = style.element.find(qn("w:rPr"))
    if rPr is None:
        return None
    rFonts = rPr.find(qn("w:rFonts"))
    return rFonts.get(qn("w:eastAsia")) if rFonts is not None else None


def _font_size(style):
    if style.font.size is None:
        return None
    return float(style.font.size.pt)


def _style_line(style):
    pPr = style.element.find(qn("w:pPr"))
    if pPr is None:
        return None, None
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        return None, None
    return sp.get(qn("w:line")), sp.get(qn("w:lineRule"))


def _line_matches(style, expected):
    line, rule = _style_line(style)
    kind, val = expected
    if kind == "exact":
        exp_line, exp_rule = str(int(round(val * 20))), "exact"
    else:
        exp_line, exp_rule = str(int(round(val))), "auto"
    return line == exp_line and rule == exp_rule


def _style_indent_chars(style):
    pPr = style.element.find(qn("w:pPr"))
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    return ind.get(qn("w:firstLineChars")) if ind is not None else None


def _style_snap(style):
    pPr = style.element.find(qn("w:pPr"))
    if pPr is None:
        return None
    snap = pPr.find(qn("w:snapToGrid"))
    return snap.get(qn("w:val")) if snap is not None else None


def _paragraph_indent_chars(p):
    pPr = p._p.pPr
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    return ind.get(qn("w:firstLineChars")) if ind is not None else None


def _report(label, issues, fails, warns=None, warn=False):
    if issues:
        kind = "注意" if warn else "失败"
        print(f"  [{kind}] {label}")
        for i in issues:
            print(f"        - {i}")
        bucket = warns if warn else fails
        if bucket is not None:
            bucket.extend(f"{label}：{i}" for i in issues)
    else:
        print(f"  [通过] {label}")


def _line_text(line):
    kind, val = line
    if kind == "exact":
        return f"固定值 {val:g} 磅"
    return f"{val / 240:g} 倍行距"


def cmd_verify(args):
    _load_config_arg(args)
    src = os.path.abspath(args.src)
    if not os.path.exists(src):
        sys.exit("找不到文件: " + src)
    doc = Document(src)
    key, prof = get_profile(args.paper)
    roles = get_roles(prof)
    body_role = roles["body"]
    cap_role = roles["caption"]
    body_line = (parse_line(getattr(args, "line_body", None))
                 or body_role.get("line") or prof["body_line"])
    head_line = parse_line(getattr(args, "line_heading", None))
    want_cols = (prof.get("columns", 1)
                 if getattr(args, "columns", None) is None else args.columns)
    want_gap = (prof.get("column_gap_cm", 0.75)
                if getattr(args, "column_gap", None) is None
                else args.column_gap)
    fails, warns = [], []

    print(f"核对文件：{src}")
    print(f"版式预设：{key}（{prof['label']}）\n")

    # ---- 页面
    s = doc.sections[0]
    pg = prof["page"]
    page_issues = []
    for label, actual, expected in (
            ("纸张宽", s.page_width, pg["w"]),
            ("纸张高", s.page_height, pg["h"]),
            ("上边距", s.top_margin, pg["top"]),
            ("下边距", s.bottom_margin, pg["bottom"]),
            ("左边距", s.left_margin, pg["left"]),
            ("右边距", s.right_margin, pg["right"]),
            ("页眉距离", s.header_distance, Cm(1.27)),
            ("页脚距离", s.footer_distance, Cm(1.27))):
        if abs(int(actual) - int(expected)) > 3000:
            page_issues.append(
                f"{label}：实际 {int(actual) / 360000:.2f} cm ≠ "
                f"期望 {int(expected) / 360000:.2f} cm")
    _report("页面设置（纸张 / 页边距 / 页眉页脚距离）", page_issues, fails)

    # ---- 分栏
    last = doc.sections[-1]
    cols_el = last._sectPr.find(qn("w:cols"))
    got_cols = int(cols_el.get(qn("w:num")) or 1) if cols_el is not None else 1
    col_issues = []
    if got_cols != want_cols:
        col_issues.append(f"栏数应为 {want_cols}，实际 {got_cols}")
    elif want_cols > 1:
        got_gap = int(cols_el.get(qn("w:space")) or 0)
        want_gap_tw = int(round(Cm(want_gap).twips))
        if abs(got_gap - want_gap_tw) > 2:
            col_issues.append(f"栏间距应为 {want_gap:g}cm，"
                              f"实际 {got_gap / 567:.2f}cm")
    _report(f"分栏（{want_cols} 栏）", col_issues, fails)

    # ---- 正文 Normal
    normal = _style(doc, "Normal")
    if normal is None:
        fails.append("缺少 Normal 样式")
        print("  [失败] 缺少 Normal 样式")
    else:
        issues = []
        body_ea = body_role.get("ea", SONG)
        body_size = float(body_role.get("size", SZ_BODY).pt)
        body_chars = str(int(body_role.get("indent_chars", 2) * 100))
        if _ea_font(normal) != body_ea:
            issues.append(f"中文字体应为{body_ea}，实际 {_ea_font(normal)}")
        if _font_size(normal) != body_size:
            issues.append(f"字号应为 {body_size:g}pt，实际 {_font_size(normal)}")
        if not _line_matches(normal, body_line):
            line, rule = _style_line(normal)
            issues.append(f"行距应为 {_line_text(body_line)}，"
                          f"实际 line={line} rule={rule}")
        if _style_indent_chars(normal) != body_chars:
            issues.append(f"首行缩进应为 {int(body_chars) / 100:g} 字符"
                          f"（firstLineChars={body_chars}）")
        if _style_snap(normal) != "0":
            issues.append("未关闭「对齐到网格」（snapToGrid=0），多倍行距会被拉歪")
        if normal.font.bold:
            issues.append("正文被设为加粗")
        _report(f"正文 Normal（{body_ea} {body_size:g}pt / "
                f"{int(body_chars) / 100:g} 字符缩进 / 行距 / 关网格）",
                issues, fails)

    # ---- 各级标题
    for name, (ea, size, bold, sb, sa, line) in prof["headings"].items():
        st = _style(doc, name)
        lvl = name[-1]
        line = head_line or line
        if st is None:
            fails.append(f"缺少样式 {name}")
            print(f"  [失败] {lvl} 级标题：缺少样式 {name}")
            continue
        issues = []
        if _ea_font(st) != ea:
            issues.append(f"中文字体应为{ea}，实际 {_ea_font(st)}")
        if _font_size(st) != float(size.pt):
            issues.append(f"字号应为 {size.pt:g}pt，实际 {_font_size(st)}")
        if st.font.bold:
            issues.append("标题被加粗（本规范标题一律不加粗）")
        try:
            rgb = st.font.color.rgb
            if rgb is not None and str(rgb).upper() != "000000":
                issues.append(f"字色不是黑色，实际 {rgb}")
        except Exception:
            pass
        if not _line_matches(st, line):
            actual, rule = _style_line(st)
            issues.append(f"行距应为 {_line_text(line)}，"
                          f"实际 line={actual} rule={rule}")
        want_chars = str(prof["h2_indent_chars"] * 100) if name == "Heading 2" else "0"
        got_chars = _style_indent_chars(st) or "0"
        if got_chars != want_chars:
            issues.append(f"首行缩进应为 {int(want_chars) / 100:g} 字符，"
                          f"实际 firstLineChars={got_chars}")
        if st.paragraph_format.space_before != Pt(sb):
            issues.append(f"段前应为 {sb}pt")
        if st.paragraph_format.space_after != Pt(sa):
            issues.append(f"段后应为 {sa}pt")
        pPr = st.element.find(qn("w:pPr"))
        if pPr is None or pPr.find(qn("w:keepNext")) is None:
            issues.append("缺少「与下段同页」（keepNext）")
        _report(f"{lvl} 级标题 {name}", issues, fails)

    # ---- 图题 / 表题
    cap = _style(doc, "Caption")
    if cap is None:
        fails.append("缺少 Caption 样式")
        print("  [失败] 缺少 Caption 样式")
    else:
        issues = []
        cap_ea = cap_role.get("ea", HEI)
        cap_size = float(cap_role.get("size", SZ_CAPTION).pt)
        cap_chars = str(int(cap_role.get("indent_chars", 0) * 100))
        if _ea_font(cap) != cap_ea:
            issues.append(f"中文字体应为{cap_ea}，实际 {_ea_font(cap)}")
        if _font_size(cap) != cap_size:
            issues.append(f"字号应为 {cap_size:g}pt，实际 {_font_size(cap)}")
        if bool(cap.font.bold) != bool(cap_role.get("bold", False)):
            issues.append("题注加粗设置与配置不符")
        if cap.paragraph_format.alignment != align_of(
                cap_role.get("align", "center")):
            issues.append("题注对齐方式与配置不符")
        if _style_indent_chars(cap) not in (None, cap_chars):
            issues.append(f"题注首行缩进应为 {int(cap_chars) / 100:g} 字符")
        if not _line_matches(cap, cap_role.get("line") or LINE_15):
            actual, rule = _style_line(cap)
            issues.append(f"题注行距应为 {_line_text(cap_role.get('line') or LINE_15)}，"
                          f"实际 line={actual} rule={rule}")
        _report(f"图题 / 表题 Caption（{cap_ea} {cap_size:g}pt / 居中）",
                issues, fails)

    # ---- 其他角色样式是否齐全
    missing = []
    for name in ("TableNote", "Reference", "Step", "Quote-CN"):
        if _style(doc, name) is None:
            missing.append(name)
    if missing:
        fails.append("缺少样式：" + "、".join(missing))
        print("  [失败] 缺少角色样式：" + "、".join(missing))
    else:
        print("  [通过] 角色样式齐全（TableNote / Reference / Step / Quote-CN）")

    # ---- 其余角色样式的字体 / 字号 / 行距 / 缩进 / 对齐
    body_override = parse_line(getattr(args, "line_body", None))
    for style_name, role_name, label in (
            ("TableNote", "note", "表注 TableNote"),
            ("Reference", "reference", "参考文献 Reference"),
            ("Step", "step", "步骤行 Step"),
            ("Quote-CN", "quote", "引文 Quote-CN")):
        st = _style(doc, style_name)
        if st is None:
            continue
        r = roles[role_name]
        issues = []
        if _ea_font(st) != r.get("ea"):
            issues.append(f"中文字体应为{r.get('ea')}，实际 {_ea_font(st)}")
        if _font_size(st) != float(r.get("size").pt):
            issues.append(f"字号应为 {r.get('size').pt:g}pt，实际 {_font_size(st)}")
        expected_line = r.get("line")
        if role_name in ("step", "quote") and body_override:
            expected_line = body_override
        if not _line_matches(st, expected_line):
            actual, rule = _style_line(st)
            issues.append(f"行距应为 {_line_text(expected_line)}，"
                          f"实际 line={actual} rule={rule}")
        want_chars = str(int(r.get("indent_chars", 0) * 100))
        if _style_indent_chars(st) not in (None, want_chars):
            issues.append(f"首行缩进应为 {int(want_chars) / 100:g} 字符")
        if st.paragraph_format.alignment != align_of(r.get("align")):
            issues.append("对齐方式与配置不符")
        if role_name == "quote":
            want_cm = float(r.get("left_cm", 0.74))
            if st.paragraph_format.left_indent is None or \
                    abs(int(st.paragraph_format.left_indent) - int(Cm(want_cm))) > 3000:
                issues.append(f"左缩进应为 {want_cm:g}cm")
        _report(label, issues, fails)

    # ---- 三线表
    if not doc.tables:
        print("  [通过] 未使用表格（跳过三线表检查）")
    else:
        issues = []
        table_lines = prof.get("table_lines",
                               {"top": 1.5, "header": 0.5, "bottom": 1.5})
        top_sz = _border_sz(table_lines.get("top", 1.5))
        header_sz = _border_sz(table_lines.get("header", 0.5))
        bottom_sz = _border_sz(table_lines.get("bottom", 1.5))
        for i, t in enumerate(doc.tables, 1):
            tblPr = t._tbl.tblPr
            borders = tblPr.find(qn("w:tblBorders"))
            if borders is None:
                issues.append(f"表{i}：没有 tblBorders")
            else:
                top = borders.find(qn("w:top"))
                bottom = borders.find(qn("w:bottom"))
                insideH = borders.find(qn("w:insideH"))
                insideV = borders.find(qn("w:insideV"))
                if not (top is not None and top.get(qn("w:val")) == "single"
                        and top.get(qn("w:sz")) == top_sz):
                    issues.append(f"表{i}：顶线不是 {table_lines.get('top', 1.5):g}pt 单实线")
                if not (bottom is not None and bottom.get(qn("w:val")) == "single"
                        and bottom.get(qn("w:sz")) == bottom_sz):
                    issues.append(f"表{i}：底线不是 {table_lines.get('bottom', 1.5):g}pt 单实线")
                for tag, el in (("insideH", insideH), ("insideV", insideV)):
                    if el is not None and el.get(qn("w:val")) not in ("none", "0", None):
                        issues.append(f"表{i}：不应有 {tag} 线")
            if t.alignment != WD_TABLE_ALIGNMENT.CENTER:
                issues.append(f"表{i}：表格未居中")
            if t.rows:
                trPr = t.rows[0]._tr.find(qn("w:trPr"))
                if trPr is None or trPr.find(qn("w:tblHeader")) is None:
                    issues.append(f"表{i}：表头行未设置跨页重复")
                tcPr = t.rows[0].cells[0]._tc.tcPr
                tcB = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
                top = tcB.find(qn("w:top")) if tcB is not None else None
                if top is None or top.get(qn("w:sz")) != top_sz:
                    issues.append(f"表{i}：表头格顶线不是 {table_lines.get('top', 1.5):g}pt")
        _report(f"三线表（{table_lines.get('top', 1.5):g}pt 顶线 / "
                f"{table_lines.get('header', 0.5):g}pt 表头线 / "
                f"{table_lines.get('bottom', 1.5):g}pt 底线 / 表头重复）",
                issues, fails)

    # ---- 四级标题
    h4 = [p for p in doc.paragraphs
          if p.style is not None and p.style.name == "Heading 4"]
    if h4:
        fails.append(f"出现 {len(h4)} 个 Heading 4（本规范只有三级标题）")
        print(f"  [失败] 出现 {len(h4)} 个四级标题（Heading 4）")
    else:
        print("  [通过] 没有四级标题")

    # ---- 段落级细节（只提示，不判失败）
    para_issues = []
    want_body_chars = str(int(body_role.get("indent_chars", 2) * 100))
    want_cap_chars = str(int(cap_role.get("indent_chars", 0) * 100))
    for p in doc.paragraphs:
        st = p.style.name if p.style is not None else ""
        text = p.text.strip()
        # 目录域所在段落是工具故意清零缩进的，不参与正文缩进检查
        if p._p.findall(".//" + qn("w:instrText")):
            continue
        if st == "Normal" and text and p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            chars = _paragraph_indent_chars(p)
            if chars == "0" and want_body_chars != "0":
                para_issues.append(f"正文被清零缩进：{text[:30]}")
            elif chars not in (None, want_body_chars):
                para_issues.append(f"正文缩进异常（{chars}）：{text[:30]}")
        if st in ("Heading 1", "Heading 2", "Heading 3"):
            if any(r.bold for r in p.runs):
                para_issues.append(f"标题段落有加粗：{text[:30]}")
        if st == "Caption":
            chars = _paragraph_indent_chars(p)
            if chars not in (None, want_cap_chars):
                para_issues.append(f"题注首行缩进应为 {int(want_cap_chars) / 100:g} 字符：{text[:30]}")
        if has_drawing(p) and p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            para_issues.append(f"图片段落未居中：{text[:30] or '（无文字）'}")
    _report("段落级细节（正文缩进 / 题注缩进 / 图片居中）",
            para_issues, fails, warns, warn=True)

    # ---- 疑似四级编号（只提示）
    h4_like = [p.text.strip() for p in doc.paragraphs
               if RE_H4_ANY.match(p.text.strip())]
    if h4_like:
        warns.append(f"{len(h4_like)} 段以「（1）」类编号开头，"
                     f"本规范只有三级，已按正文处理：" + h4_like[0][:30])

    # ---- 目录域提示
    toc = any("TOC" in (it.text or "")
              for it in doc.element.iter(qn("w:instrText")))
    if toc:
        print("  [提示] 检测到目录域：在 Word 中按 F9（或右键→更新域）生成页码。")

    # ---- 总结
    print("\n=== 核对结果 ===")
    if not fails and not warns:
        print("全部通过。")
    else:
        if fails:
            print(f"失败 {len(fails)} 项：")
            for f in fails:
                print("  - " + f)
        if warns:
            print(f"注意 {len(warns)} 项：")
            for w in warns:
                print("  - " + w)
    sys.exit(1 if fails else 0)


def cmd_spec(args):
    config_path = getattr(args, "config", None)
    _load_config_arg(args)
    print("中文论文 / 报告 Word 通用版式（便携版）")
    print("层级只有三级：一级 / 二级 / 三级，三级之后直接是正文；不设四级标题。")
    items = ([(args.paper, PROFILES[args.paper])] if config_path
             else list(PROFILES.items()))
    for key, prof in items:
        p = prof["page"]
        print(f"\n=== {key}：{prof['label']} ===")
        print(f"纸张：{int(p['w']) / 360000:.2f} × {int(p['h']) / 360000:.2f} cm")
        print(f"页边距：上 {int(p['top']) / 360000:.2f} / 下 {int(p['bottom']) / 360000:.2f}"
              f" / 左 {int(p['left']) / 360000:.2f} / 右 {int(p['right']) / 360000:.2f} cm")
        ncols = prof.get("columns", 1)
        gap = prof.get("column_gap_cm", 0.75)
        print(f"栏数：{ncols}" + (f"，栏间距 {gap:g}cm" if ncols > 1 else ""))
        for name, (ea, size, bold, sb, sa, line) in prof["headings"].items():
            lvl = name[-1]
            extra = ""
            if name == "Heading 2" and prof["h2_indent_chars"]:
                extra = f"，首行缩进 {prof['h2_indent_chars']} 字符"
            print(f"{lvl} 级标题：{ea} {size.pt:g}pt，"
                  f"{'加粗' if bold else '不加粗'}{extra}，"
                  f"段前 {sb}pt / 段后 {sa}pt，{_line_text(line)}")
        roles = get_roles(prof)
        body = roles["body"]
        cap = roles["caption"]
        note = roles["note"]
        ref = roles["reference"]
        step = roles["step"]
        quote = roles["quote"]
        print(f"正文：{body.get('ea')} {body.get('size').pt:g}pt，"
              f"西文 {body.get('ascii', TNR)}，"
              f"首行缩进 {body.get('indent_chars', 2)} 字符，"
              f"{_line_text(body.get('line') or prof['body_line'])}，"
              f"{body.get('align', 'justify')}，"
              f"段前 {body.get('before', 2.5):g}pt / 段后 {body.get('after', 0.5):g}pt")
        print(f"图题 / 表题：{cap.get('ea')} {cap.get('size').pt:g}pt、"
              f"{cap.get('align', 'center')}、首行缩进 "
              f"{cap.get('indent_chars', 0)} 字符；图题在图下方，表题在表格上方。")
        print(f"表注：{note.get('ea')} {note.get('size').pt:g}pt、"
              f"缩进 {note.get('indent_chars', 2)} 字符、"
              f"{_line_text(note.get('line') or LINE_SINGLE)}。")
        tl = prof.get("table_lines", {"top": 1.5, "header": 0.5, "bottom": 1.5})
        print(f"三线表：顶线 {tl.get('top', 1.5):g}pt、表头下 "
              f"{tl.get('header', 0.5):g}pt、底线 {tl.get('bottom', 1.5):g}pt，"
              f"无竖线；表头行跨页重复。")
        print(f"参考文献：{ref.get('ea')} {ref.get('size').pt:g}pt、"
              f"{_line_text(ref.get('line') or LINE_SINGLE)}、"
              f"缩进 {ref.get('indent_chars', 0)} 字符、"
              f"{ref.get('align', 'left')}。")
        print(f"步骤行：{step.get('align', 'left')}、首行缩进 "
              f"{step.get('indent_chars', 2)} 字符，序号"
              f"{step.get('num_ea', HEI)}、内容{step.get('ea')}。")
        print(f"引文：{quote.get('ea')} {quote.get('size').pt:g}pt，"
              f"左右缩进 {quote.get('left_cm', 0.74):g}cm。")
        print(f"封面：@cover + `键: 值`；标题 "
              f"{roles['cover_title'].get('ea')} "
              f"{roles['cover_title'].get('size').pt:g}pt。")
    print("目录：@toc 插入 TOC 域；Windows + Word 环境会自动刷新一次，否则按 F9。")
    print("页码：@pagenum 另起一节并从 1 开始；封面 / 目录不编页码。")
    print("\n硬约束：中文必须写 w:eastAsia；标题一律不加粗；"
          "多倍行距关闭 snapToGrid；首行缩进同时写 firstLineChars 与 firstLine。")


def cmd_dump_config(args):
    _load_config_arg(args)
    cfg = dump_config(args.paper, args.out)
    if args.out:
        print("已生成配置:", args.out)
    else:
        print(json.dumps(cfg, ensure_ascii=False, indent=2))


# ---------------------------------------------------------------- 参数解析
def build_parser():
    ap = argparse.ArgumentParser(
        prog="cn_paper_docx.py",
        description="中文论文 / 报告 Word 通用排版（便携版，单文件）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add_common(p):
        p.add_argument("--paper", default=DEFAULT_PROFILE,
                       choices=list(PROFILES), help="版式预设（默认 letter）")
        p.add_argument("--config", default=None,
                       help="JSON 版式配置（只写要改的项，其余继承 --paper / base）")
        p.add_argument("--line-body", default=None,
                       help="正文行距：1.5 表示 1.5 倍，28pt 表示固定值 28 磅")
        p.add_argument("--line-heading", default=None,
                       help="标题行距：同上，与正文可分别设置")
        p.add_argument("--columns", type=parse_columns, default=None,
                       help="栏数：1 单栏（默认）或 2 双栏")
        p.add_argument("--column-gap", type=float, default=None,
                       help="双栏栏间距，单位 cm（默认 0.75）")

    p = sub.add_parser("new", help="从标记稿生成完整文档")
    p.add_argument("src", nargs="?", help="标记稿（.txt/.md，UTF-8）")
    p.add_argument("-o", "--out", required=True, help="输出 .docx")
    add_common(p)
    p.add_argument("--demo", action="store_true", help="忽略 src，生成演示稿")
    p.add_argument("--no-toc-refresh", action="store_true",
                   help="不尝试用 Word 刷新目录")
    p.set_defaults(func=cmd_new)

    p = sub.add_parser("demo", help="生成演示稿")
    p.add_argument("-o", "--out", required=True, help="输出 .docx")
    add_common(p)
    p.add_argument("--no-toc-refresh", action="store_true")
    p.set_defaults(func=cmd_new, src=None, demo=True)

    p = sub.add_parser("apply", help="给已有 .docx 套格式")
    p.add_argument("src", help="原 .docx（不会被修改）")
    p.add_argument("-o", "--out", help="输出 .docx")
    add_common(p)
    p.add_argument("--center-titles", default="",
                   help='要居中的标题，逗号分隔，如 "参考文献,致谢"')
    p.add_argument("--dry-run", action="store_true", help="只报告，不写文件")
    p.add_argument("--no-backup", action="store_true",
                   help="覆盖输出文件时不先备份")
    p.set_defaults(func=cmd_apply)

    p = sub.add_parser("template", help="生成空白模板 / 样式样张")
    p.add_argument("-o", "--out", required=True, help="模板输出路径")
    p.add_argument("--specimen", default=None, help="同时生成样式样张的路径")
    add_common(p)
    p.set_defaults(func=cmd_template)

    p = sub.add_parser("verify", help="核对生成的 docx 是否符合规范")
    p.add_argument("src", help="要核对的 .docx")
    p.add_argument("--paper", default=DEFAULT_PROFILE, choices=list(PROFILES))
    p.add_argument("--config", default=None,
                   help="生成时用的 JSON 版式配置")
    p.add_argument("--line-body", default=None,
                   help="与生成时一致的正文行距覆盖（1.5 / 28pt）")
    p.add_argument("--line-heading", default=None,
                   help="与生成时一致的标题行距覆盖（1.5 / 28pt）")
    p.add_argument("--columns", type=parse_columns, default=None,
                   help="与生成时一致的栏数（1 / 2）")
    p.add_argument("--column-gap", type=float, default=None,
                   help="与生成时一致的栏间距（cm）")
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser("spec", help="打印当前版式数值")
    p.add_argument("--config", default=None,
                   help="打印配置合并后的版式数值")
    p.set_defaults(func=cmd_spec)

    p = sub.add_parser("dump-config", help="导出可编辑的 JSON 版式配置")
    p.add_argument("--paper", default=DEFAULT_PROFILE, choices=list(PROFILES),
                   help="要导出的版式预设（默认 letter）")
    p.add_argument("--config", default=None,
                   help="基于已有配置导出（可与 --paper 的 base 叠加）")
    p.add_argument("-o", "--out", default=None,
                   help="输出 JSON 路径；不给就打印到屏幕")
    p.set_defaults(func=cmd_dump_config)
    return ap


def main():
    args = build_parser().parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

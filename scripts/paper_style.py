# -*- coding: utf-8 -*-
"""中文论文/报告 Word 版式核心模块。

两套版式预设（按纸张分），各对应一份参考稿：
    letter —— 对应「总.docx」        Letter 纸 / 二级黑体四号 / 1.5 倍行距
    a4     —— 对应「第四章　实验流程.docx」  A4 纸 / 二级楷体三号+缩进 / 固定 28 磅

层级只有三级，三级之后直接是正文（不设四级标题）。

单位说明：docx 内部长度用 twips(1/20 pt) 或 EMU，本模块统一用 python-docx 的
Pt/Cm/Inches 封装，只有“首行缩进按字符”这种必须写 XML 的地方才直接操作 lxml。
"""
import json
import os

from docx.shared import Pt, Cm, Inches, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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

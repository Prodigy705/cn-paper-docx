
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

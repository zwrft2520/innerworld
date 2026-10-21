# -*- coding: utf-8 -*-
"""Build a single DOCX book from the depression-treatment Markdown library.

The DOCX is intentionally the source of truth; LibreOffice converts it to PDF.
"""
import os
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.opc.constants import RELATIONSHIP_TYPE as RT

ROOT = Path(r"D:\zwr\idea\project\111\innerworld\精神疾病确诊")
OUT_DIR = ROOT / ".build"
OUT_DIR.mkdir(parents=True, exist_ok=True)
DOCX_PATH = OUT_DIR / os.environ.get("BOOK_DOCX", "精神疾病确诊.docx")
BOOK_LIMIT = int(os.environ.get("BOOK_LIMIT", "0"))
BOOK_MODULE = os.environ.get("BOOK_MODULE", "")

BODY_FONT_EA = "宋体"
BODY_FONT_ASCII = "SimSun"
HEAD_FONT_EA = "黑体"
HEAD_FONT_ASCII = "Arial"
COVER_FONT_EA = "微软雅黑"
COVER_FONT_ASCII = "Microsoft YaHei"
MONO_FONT = "Consolas"

DARK = "17365D"
BLUE = "1F4E79"
MID_BLUE = "2E74B5"
GREY = "595959"
LIGHT_GREY = "F2F2F2"
QUOTE_FILL = "F2F6FB"
TABLE_ALT = "F7F9FC"
BORDER = "BFBFBF"

PAGE_W = Cm(16.8)


def set_run_font(run, ascii_name=BODY_FONT_ASCII, ea_name=BODY_FONT_EA,
                 size=None, bold=None, color=None, italic=None, underline=None):
    run.font.name = ascii_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), ascii_name)
    rFonts.set(qn("w:hAnsi"), ascii_name)
    rFonts.set(qn("w:eastAsia"), ea_name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    if italic is not None:
        run.font.italic = italic
    if underline is not None:
        run.font.underline = underline


def shade_paragraph(p, fill):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def paragraph_border(p, edge="bottom", color=MID_BLUE, sz="6", space="4"):
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), space)
    el.set(qn("w:color"), color)
    pBdr.append(el)


def add_page_field(paragraph):
    run = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t"); txt.text = "1"
    fld3 = OxmlElement("w:fldChar"); fld3.set(qn("w:fldCharType"), "end")
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
    run._r.append(txt); run._r.append(fld3)
    set_run_font(run, size=9, color=GREY)


def add_hyperlink(paragraph, url, label, size=10.5, color=MID_BLUE, bold=False):
    if not re.match(r"^(https?://|mailto:)", url):
        run = paragraph.add_run(label)
        set_run_font(run, size=size, color=color, bold=bold, underline=True)
        return
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), BODY_FONT_ASCII)
    rFonts.set(qn("w:hAnsi"), BODY_FONT_ASCII)
    rFonts.set(qn("w:eastAsia"), BODY_FONT_EA)
    rPr.append(rFonts)
    c = OxmlElement("w:color"); c.set(qn("w:val"), color); rPr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    if bold:
        b = OxmlElement("w:b"); rPr.append(b)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2))); rPr.append(sz)
    new_run.append(rPr)
    t = OxmlElement("w:t"); t.text = label; t.set(qn("xml:space"), "preserve")
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


TOKEN_RE = re.compile(r"(\*\*.+?\*\*|\[[^\]]+\]\([^)]+\)|`[^`]+`|\*[^*\n]+\*)")
URL_RE = re.compile(r"(https?://[^\s，。；、）\)]+)")


def add_text_with_urls(p, text, size=10.5, color=None, bold=False):
    for bit in URL_RE.split(text):
        if not bit:
            continue
        if URL_RE.fullmatch(bit):
            add_hyperlink(p, bit, bit, size=size, color=color or MID_BLUE, bold=bold)
        else:
            run = p.add_run(bit)
            set_run_font(run, size=size, color=color, bold=bold)


def add_runs(p, text, size=10.5, color=None, bold_all=False, header=False):
    """Parse a small subset of inline markdown into runs."""
    pos = 0
    for m in TOKEN_RE.finditer(text):
        if m.start() > pos:
            add_text_with_urls(p, text[pos:m.start()], size=size, color=color, bold=bold_all)
        token = m.group(0)
        if token.startswith("**") and token.endswith("**") and len(token) > 4:
            run = p.add_run(token[2:-2])
            set_run_font(run, size=size, color=color, bold=True)
        elif token.startswith("[") and "](" in token and token.endswith(")"):
            mm = re.match(r"\[([^\]]+)\]\(([^)]+)\)", token)
            label, url = mm.group(1), mm.group(2)
            add_hyperlink(p, url, label, size=size,
                          color=("FFFFFF" if header else (color or MID_BLUE)),
                          bold=bold_all)
        elif token.startswith("`") and token.endswith("`"):
            run = p.add_run(token[1:-1])
            set_run_font(run, ascii_name=MONO_FONT, ea_name=MONO_FONT, size=size - 0.5,
                         color=color, bold=bold_all)
        elif token.startswith("*") and token.endswith("*") and len(token) > 2:
            run = p.add_run(token[1:-1])
            set_run_font(run, size=size, color=color, bold=bold_all, italic=True)
        else:
            run = p.add_run(token)
            set_run_font(run, size=size, color=color, bold=bold_all)
        pos = m.end()
    if pos < len(text):
        add_text_with_urls(p, text[pos:], size=size, color=color, bold=bold_all)


def add_heading(doc, text, level, page_break=False, color=None, size=None):
    p = doc.add_paragraph(style=f"Heading {level}")
    pf = p.paragraph_format
    pf.page_break_before = page_break
    pf.keep_with_next = True
    pf.space_before = Pt({1: 0, 2: 4, 3: 14, 4: 10}.get(level, 8))
    pf.space_after = Pt({1: 8, 2: 6, 3: 5, 4: 3}.get(level, 4))
    pf.line_spacing = 1.2
    run = p.add_run(text)
    if level == 1:
        set_run_font(run, HEAD_FONT_ASCII, HEAD_FONT_EA, size or 23, True, color or DARK)
    elif level == 2:
        set_run_font(run, HEAD_FONT_ASCII, HEAD_FONT_EA, size or 18, True, color or BLUE)
    elif level == 3:
        set_run_font(run, HEAD_FONT_ASCII, HEAD_FONT_EA, size or 13.5, True, color or BLUE)
    else:
        set_run_font(run, HEAD_FONT_ASCII, HEAD_FONT_EA, size or 11.5, True, color or "333333")
    return p


def add_hr(doc, color=BORDER, space_before=2, space_after=8):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    paragraph_border(p, "bottom", color=color, sz="6", space="1")
    return p


def add_body(doc, text, indent=True, size=10.5, color=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             space_after=3, line=1.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align
    pf.line_spacing = line
    pf.space_before = Pt(0)
    pf.space_after = Pt(space_after)
    if indent:
        pf.first_line_indent = Pt(size * 2)
    add_runs(p, text, size=size, color=color)
    return p


def add_list_item(doc, marker, text, level=0, quote=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.75 + 0.45 * level)
    pf.first_line_indent = Cm(-0.45)
    pf.line_spacing = 1.35
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    if quote:
        shade_paragraph(p, QUOTE_FILL)
        paragraph_border(p, "left", color=MID_BLUE, sz="12", space="4")
    run = p.add_run(marker + " ")
    set_run_font(run, size=10.5, color=MID_BLUE if not quote else GREY, bold=True)
    add_runs(p, text, size=10.5, color=None if not quote else "404040")
    return p


def add_quote_paragraph(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.55)
    pf.right_indent = Cm(0.25)
    pf.line_spacing = 1.35
    pf.space_before = Pt(2)
    pf.space_after = Pt(3)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    shade_paragraph(p, QUOTE_FILL)
    paragraph_border(p, "left", color=MID_BLUE, sz="14", space="5")
    add_runs(p, text, size=9.5, color="404040")
    return p


def add_code_block(doc, lines):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.35)
    pf.right_indent = Cm(0.2)
    pf.line_spacing = 1.0
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)
    shade_paragraph(p, "F4F4F4")
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "3"); el.set(qn("w:color"), "D9D9D9")
        pBdr.append(el)
    pPr.append(pBdr)
    for i, line in enumerate(lines):
        txt = line.replace("\t", "    ").replace(" ", "\u00A0")
        run = p.add_run(txt)
        set_run_font(run, MONO_FONT, MONO_FONT, 8.2, color="333333")
        if i < len(lines) - 1:
            run.add_break()
    return p


def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def split_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip().replace("\\|", "|") for c in s.split("|")]


def is_separator(line):
    s = line.strip()
    if not s or "-" not in s:
        return False
    s = s.strip("|").strip()
    parts = [p.strip() for p in s.split("|")]
    if not parts:
        return False
    for p in parts:
        if not re.fullmatch(r":?-{2,}:?", p):
            return False
    return True


def add_table(doc, lines):
    header = split_row(lines[0])
    sep = split_row(lines[1])
    n = max(len(header), len(sep))
    header = (header + [""] * n)[:n]
    aligns = []
    for c in (sep + [""] * n)[:n]:
        if c.startswith(":") and c.endswith(":"):
            aligns.append(WD_ALIGN_PARAGRAPH.CENTER)
        elif c.endswith(":"):
            aligns.append(WD_ALIGN_PARAGRAPH.RIGHT)
        else:
            aligns.append(WD_ALIGN_PARAGRAPH.LEFT)

    body = []
    for line in lines[2:]:
        if not line.strip() or not line.lstrip().startswith("|"):
            break
        cells = split_row(line)
        cells = (cells + [""] * n)[:n]
        body.append(cells)

    table = doc.add_table(rows=1, cols=n)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    if n == 1:
        widths = [Cm(16.6)]
    elif n == 2:
        widths = [Cm(4.3), Cm(12.3)]
    elif n == 3:
        widths = [Cm(3.9), Cm(6.3), Cm(6.4)]
    else:
        widths = [Cm(3.2), Cm(4.5), Cm(4.4), Cm(4.5)]

    hdr = table.rows[0]
    set_repeat_header(hdr)
    for i, text in enumerate(header):
        cell = hdr.cells[i]
        cell.width = widths[i]
        set_cell_shading(cell, BLUE)
        p = cell.paragraphs[0]
        p.alignment = aligns[i]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.1
        add_runs(p, text, size=9, color="FFFFFF", bold_all=True, header=True)

    for r_idx, cells in enumerate(body):
        row = table.add_row()
        for i, text in enumerate(cells):
            cell = row.cells[i]
            cell.width = widths[i]
            if r_idx % 2 == 1:
                set_cell_shading(cell, TABLE_ALT)
            p = cell.paragraphs[0]
            p.alignment = aligns[i]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.15
            add_runs(p, text, size=8.8, color="333333")

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(4)
    spacer.paragraph_format.line_spacing = 1.0
    return table


def join_soft(lines):
    out = ""
    for s in lines:
        s = s.strip()
        if not s:
            continue
        if not out:
            out = s
            continue
        if out[-1].isascii() and out[-1].isalnum() and s[0].isascii() and s[0].isalnum():
            out += " " + s
        else:
            out += s
    return out


def collect_table(lines, i):
    start = i
    i += 1
    while i < len(lines) and lines[i].lstrip().startswith("|"):
        i += 1
    return lines[start:i], i


def collect_quote(lines, i):
    start = i
    while i < len(lines) and lines[i].lstrip().startswith(">"):
        i += 1
    return lines[start:i], i


def collect_list(lines, i):
    line = lines[i]
    m = re.match(r"^(\s*)([-*])\s+(\[[ xX]\]\s+)?(.*)$", line)
    if m:
        indent = len(m.group(1))
        marker = "□" if m.group(3) else "•"
        if m.group(3) and m.group(3).lower().startswith("[x"):
            marker = "☑"
        text = m.group(4).strip()
        j = i + 1
        while j < len(lines) and lines[j].startswith(("  ", "\t")) and lines[j].strip():
            nxt = lines[j].strip()
            if nxt.startswith(("#", "|", ">", "```")):
                break
            if text and text[-1].isascii() and text[-1].isalnum() and nxt[0].isascii() and nxt[0].isalnum():
                text += " " + nxt
            else:
                text += nxt
            j += 1
        return marker, text, indent, j
    m = re.match(r"^(\s*)(\d+)[.、)]\s+(.*)$", line)
    if m:
        indent = len(m.group(1))
        marker = m.group(2) + "."
        text = m.group(3).strip()
        j = i + 1
        while j < len(lines) and lines[j].startswith(("  ", "\t")) and lines[j].strip():
            nxt = lines[j].strip()
            if nxt.startswith(("#", "|", ">", "```")):
                break
            if text and text[-1].isascii() and text[-1].isalnum() and nxt[0].isascii() and nxt[0].isalnum():
                text += " " + nxt
            else:
                text += nxt
            j += 1
        return marker, text, indent, j
    return None


def parse_markdown(doc, path, skip_title=True):
    lines = path.read_text(encoding="utf-8").splitlines()
    i = 0
    para = []

    def flush():
        nonlocal para
        if para:
            add_body(doc, join_soft(para))
            para = []

    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()

        if stripped.startswith("```"):
            flush()
            i += 1
            code = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i].rstrip("\n"))
                i += 1
            if i < len(lines):
                i += 1
            add_code_block(doc, code)
            continue

        if stripped.startswith("|") and i + 1 < len(lines) and is_separator(lines[i + 1]):
            flush()
            table_lines, i = collect_table(lines, i)
            add_table(doc, table_lines)
            continue

        if stripped.startswith("# "):
            flush()
            i += 1
            continue

        if stripped.startswith("## "):
            flush()
            add_heading(doc, stripped[3:].strip(), 3)
            i += 1
            continue

        if stripped.startswith("### "):
            flush()
            add_heading(doc, stripped[4:].strip(), 4)
            i += 1
            continue

        if stripped.startswith(">"):
            flush()
            qlines, i = collect_quote(lines, i)
            for q in qlines:
                qt = q.lstrip()
                qt = qt[1:] if qt.startswith(">") else qt
                qt = qt.lstrip()
                if not qt:
                    continue
                lm = re.match(r"^([-*])\s+(\[[ xX]\]\s+)?(.*)$", qt)
                nm = re.match(r"^(\d+)[.、)]\s+(.*)$", qt)
                if lm:
                    marker = "□" if lm.group(2) else "•"
                    add_list_item(doc, marker, lm.group(3).strip(), quote=True)
                elif nm:
                    add_list_item(doc, nm.group(1) + ".", nm.group(2).strip(), quote=True)
                else:
                    add_quote_paragraph(doc, qt)
            continue

        lm = collect_list(lines, i)
        if lm:
            flush()
            marker, text, indent, i = lm
            add_list_item(doc, marker, text, level=min(indent // 2, 2))
            continue

        if stripped in ("---", "***", "___"):
            flush()
            add_hr(doc, color="D9D9D9", space_before=4, space_after=6)
            i += 1
            continue

        if not stripped:
            flush()
            i += 1
            continue

        para.append(raw)
        i += 1

    flush()


def build():
    toc_text = (ROOT / "目录.md").read_text(encoding="utf-8")
    link_re = re.compile(r"\[([^\]]+)\]\(<([^>]+\.md)>\)")
    entries = []
    for label, target in link_re.findall(toc_text):
        if target.startswith(("模块", "附录")):
            entries.append((label, target))
    # De-duplicate while preserving order
    seen = set()
    unique = []
    for label, target in entries:
        if target not in seen:
            seen.add(target)
            unique.append((label, target))
    entries = unique
    if BOOK_MODULE:
        entries = [e for e in entries if e[1].startswith(BOOK_MODULE)]
    if BOOK_LIMIT:
        entries = entries[:BOOK_LIMIT]
    if not BOOK_LIMIT and not BOOK_MODULE and len(entries) != 64:
        print("WARNING: expected 64 article links, got", len(entries))
        for e in entries:
            print("  ", e)

    intro_re = re.compile(r"^## (模块\d+：[^\n]+)\n\n> ([^\n]+)", re.M)
    intros = {m.group(1): m.group(2).strip() for m in intro_re.finditer(toc_text)}

    modules = []
    for label, target in entries:
        folder = target.split("/")[0]
        modules.append((folder, label, target))
    module_order = []
    for folder, _, _ in modules:
        if folder not in module_order:
            module_order.append(folder)

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.1)
    section.bottom_margin = Cm(1.9)
    section.left_margin = Cm(2.1)
    section.right_margin = Cm(2.1)
    section.header_distance = Cm(0.9)
    section.footer_distance = Cm(0.9)
    section.different_first_page_header_footer = True

    # Normal style
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT_ASCII
    normal.font.size = Pt(10.5)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT_EA)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(3)

    # Header / footer
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("精神疾病确诊  ·  患者与家属的循证诊断科普")
    set_run_font(hr, size=8.5, color="808080")
    paragraph_border(hp, "bottom", color="D9D9D9", sz="4", space="2")

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = fp.add_run("第 ")
    set_run_font(r1, size=9, color=GREY)
    add_page_field(fp)
    r2 = fp.add_run(" 页")
    set_run_font(r2, size=9, color=GREY)

    # Cover
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(95)
    r = p.add_run("患者与家属的循证治疗科普")
    set_run_font(r, COVER_FONT_ASCII, COVER_FONT_EA, 12, False, MID_BLUE)
    p.paragraph_format.space_after = Pt(14)

    p = doc.add_paragraph()
    r = p.add_run("精神疾病确诊")
    set_run_font(r, COVER_FONT_ASCII, COVER_FONT_EA, 40, True, DARK)
    p.paragraph_format.space_after = Pt(8)

    p = doc.add_paragraph()
    r = p.add_run("从诊断标准到鉴别、误诊与就诊沟通")
    set_run_font(r, COVER_FONT_ASCII, COVER_FONT_EA, 14, False, "404040")
    p.paragraph_format.space_after = Pt(22)

    p = doc.add_paragraph()
    paragraph_border(p, "bottom", color=MID_BLUE, sz="18", space="1")
    p.paragraph_format.space_after = Pt(12)

    p = doc.add_paragraph()
    r = p.add_run("13 个模块   ·   64 篇   ·   V1.0 / 2026-10")
    set_run_font(r, COVER_FONT_ASCII, COVER_FONT_EA, 11, False, GREY)
    p.paragraph_format.space_after = Pt(160)

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.right_indent = Cm(0.3)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    shade_paragraph(p, "F2F6FB")
    r = p.add_run("本资料库仅用于科普学习，不能替代精神科医生面诊、诊断与处方。\n"
                  "出现自杀念头、意识障碍、严重过敏、胸痛晕厥等情况，请立即拨打 120 或前往急诊；也可拨打 12356 心理援助热线。")
    set_run_font(r, COVER_FONT_ASCII, COVER_FONT_EA, 9, False, "595959")

    doc.add_page_break()

    # Copyright / safety page
    add_heading(doc, "使用边界与安全声明", 1, page_break=False, size=20)
    add_body(doc, "这是一本写给患者和家属的诊断科普书。它把精神疾病确诊拆成 13 个模块，覆盖从主诉、病史、精神检查、诊断标准，到鉴别、共病、误诊和就诊沟通的完整路径。", indent=False, size=11)
    add_body(doc, "它不能替代面诊、诊断和处方。所有具体用药、剂量、减量、住院与特殊人群决策，都必须由具备资质的精神科医生结合个体情况作出。", indent=False, size=11, color="333333")
    add_heading(doc, "请先记住三件事", 3, page_break=False)
    for t in [
        "诊断不是对症状的简单计数；病程、功能损害、鉴别与排除，缺一不可。",
        "不要自行突然停药、减量、换药或叠加药物；停药需要在医生指导下缓慢进行。",
        "出现自杀念头、自伤行为、意识障碍、高热、严重过敏、胸痛、晕厥、抽搐，立即拨打 120 或去急诊。",
    ]:
        add_list_item(doc, "•", t)
    add_heading(doc, "版本与引用", 3, page_break=False)
    add_body(doc, "版本 V1.0，最后更新 2026-10。指南、药物适应证、医保政策与热线号码会变化；与面诊医生或最新指南不一致时，以医生和最新指南为准。引用请注明《精神疾病确诊资料库》V1.0，并附具体篇目。", indent=False, size=10, color="595959")
    add_body(doc, "书中案例均为合并、改写后的示意，不对应任何真实个人。", indent=False, size=10, color="595959")
    doc.add_page_break()

    # Manual TOC
    add_heading(doc, "总目录", 1, page_break=False, size=26)
    add_body(doc, "按模块顺序排列。每篇正文都可独立阅读；书末附 60 个术语与主要指南来源。", indent=False, size=10, color=GREY)
    add_hr(doc, color="D9D9D9", space_before=2, space_after=8)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("使用说明")
    set_run_font(r, HEAD_FONT_ASCII, HEAD_FONT_EA, 12.5, True, DARK)

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.7)
    p.paragraph_format.space_after = Pt(1.5)
    r = p.add_run("00  使用说明与免责声明")
    set_run_font(r, size=10, color="333333")

    for folder in module_order:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(11)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(folder)
        set_run_font(r, HEAD_FONT_ASCII, HEAD_FONT_EA, 12.5, True, DARK)
        for f, label, target in modules:
            if f != folder:
                continue
            # Use H1 title from file for consistency
            md = (ROOT / target).read_text(encoding="utf-8").splitlines()
            title = label
            for ln in md:
                if ln.startswith("# "):
                    title = ln[2:].strip()
                    break
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.7)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.25
            r = p.add_run(title)
            set_run_font(r, size=10, color="333333")

    doc.add_page_break()

    # Preface / 00
    add_heading(doc, "使用说明", 1, page_break=False, size=23)
    zero = ROOT / "00 使用说明与免责声明.md"
    zero_title = "00 使用说明与免责声明"
    for ln in zero.read_text(encoding="utf-8").splitlines():
        if ln.startswith("# "):
            zero_title = ln[2:].strip()
            break
    add_heading(doc, zero_title, 2, page_break=False)
    parse_markdown(doc, zero, skip_title=True)

    # Modules
    for mi, folder in enumerate(module_order):
        add_heading(doc, folder, 1, page_break=True, size=23)
        intro = intros.get(folder)
        if not intro:
            intro = "本模块的篇目可以按顺序读，也可以按问题单独查。"
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.5
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(intro)
        set_run_font(r, size=9.8, color="595959")
        add_hr(doc, color="D9D9D9", space_before=0, space_after=8)

        first = True
        for f, label, target in modules:
            if f != folder:
                continue
            path = ROOT / target
            md_lines = path.read_text(encoding="utf-8").splitlines()
            title = label
            for ln in md_lines:
                if ln.startswith("# "):
                    title = ln[2:].strip()
                    break
            add_heading(doc, title, 2, page_break=not first)
            first = False
            parse_markdown(doc, path, skip_title=True)

    # Metadata
    cp = doc.core_properties
    cp.title = "精神疾病确诊"
    cp.subject = "患者与家属的循证诊断科普"
    cp.author = "精神疾病确诊资料库"
    cp.keywords = "精神疾病, 诊断, 鉴别诊断, DSM-5-TR, ICD-11, 误诊"
    cp.comments = "V1.0 / 2026-10"
    cp.category = "精神健康科普"

    doc.save(DOCX_PATH)
    print("saved", DOCX_PATH)
    print("modules", len(module_order), "articles", len(entries), "intros", len(intros))


if __name__ == "__main__":
    build()

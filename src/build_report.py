"""Build the case study Word report (A4, print-ready) from reports/results.json.

Usage:  python src/build_report.py [--toc-pages reports/toc_pages.json]

The document contains a real Word TOC field (Word rebuilds it on open) whose
cached entries can be pre-filled with page numbers from a JSON file produced by
rendering a first pass to PDF.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "reports" / "results.json"
FIG = ROOT / "reports" / "figures"
OUT = ROOT / "reports" / "Case_Study_Data_Mining.docx"

NAVY = RGBColor(0x1F, 0x38, 0x64)
STEEL = RGBColor(0x2E, 0x54, 0x96)
GRAY = RGBColor(0x59, 0x59, 0x59)
BODY_FONT = "Cambria"
SANS = "Calibri"
MONO = "Consolas"
TEXT_W = 15.5  # usable text width in cm on portrait A4 (21 - 3.0 - 2.5)

HEADINGS: list[tuple[int, str]] = []   # (level, title) in document order

# schema-order successors so raw OOXML inserts stay valid for Word
_P_AFTER_PBDR = ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                 "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
                 "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
                 "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                 "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
                 "w:jc", "w:textDirection", "w:textAlignment",
                 "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
                 "w:rPr", "w:sectPr", "w:pPrChange")
_P_AFTER_SHD = _P_AFTER_PBDR[1:]
_TC_AFTER_BORDER = ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                    "w:tcFitText", "w:vAlign", "w:hideMark", "w:tcPrChange")
_TC_AFTER_SHD = _TC_AFTER_BORDER[1:]


# ---------------------------------------------------------------- helpers
def _pborder(par, edge: str, sz: str, color: str, space: str = "4"):
    pPr = par._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.insert_element_before(pBdr, *_P_AFTER_PBDR)
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), space)
    el.set(qn("w:color"), color)
    pBdr.append(el)


def _pshade(par, fill: str):
    pPr = par._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.insert_element_before(shd, *_P_AFTER_SHD)


def add_field(par, instr: str, placeholder: str = "1"):
    r = par.add_run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "begin")
    r._r.append(fc)
    r = par.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    r._r.append(it)
    r = par.add_run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "separate")
    r._r.append(fc)
    par.add_run(placeholder)
    r = par.add_run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "end")
    r._r.append(fc)


def rich(par, text: str):
    """Markup: **bold**, `mono`, *italic*."""
    buf, i = "", 0
    while i < len(text):
        if text.startswith("**", i):
            j = text.find("**", i + 2)
            if j > 0:
                if buf:
                    par.add_run(buf)
                    buf = ""
                par.add_run(text[i + 2:j]).bold = True
                i = j + 2
                continue
        if text.startswith("*", i) and not text.startswith("**", i):
            j = text.find("*", i + 1)
            if j > 0:
                if buf:
                    par.add_run(buf)
                    buf = ""
                par.add_run(text[i + 1:j]).italic = True
                i = j + 1
                continue
        if text.startswith("`", i):
            j = text.find("`", i + 1)
            if j > 0:
                if buf:
                    par.add_run(buf)
                    buf = ""
                run = par.add_run(text[i + 1:j])
                run.font.name = MONO
                run.font.size = Pt(10)
                i = j + 1
                continue
        buf += text[i]
        i += 1
    if buf:
        par.add_run(buf)


def _set_runs(par, name, size, color=None, bold=None):
    for run in par.runs:
        run.font.name = name
        run.font.size = size
        if color is not None:
            run.font.color.rgb = color
        if bold is not None:
            run.bold = bold


def body(doc, text: str):
    par = doc.add_paragraph(style="Normal")
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rich(par, text)
    return par


def bullet(doc, text: str):
    par = doc.add_paragraph(style="List Bullet")
    par.paragraph_format.space_after = Pt(3)
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    rich(par, text)
    return par


def h1(doc, title: str):
    HEADINGS.append((1, title))
    par = doc.add_paragraph(title, style="Heading 1")
    _pborder(par, "bottom", "8", "B4C6E7", space="3")
    return par


def h2(doc, title: str):
    HEADINGS.append((2, title))
    return doc.add_paragraph(title, style="Heading 2")


def quote_box(doc, text: str):
    par = doc.add_paragraph()
    par.paragraph_format.left_indent = Cm(0.6)
    par.paragraph_format.right_indent = Cm(0.6)
    par.paragraph_format.space_before = Pt(8)
    par.paragraph_format.space_after = Pt(8)
    _pshade(par, "F2F6FC")
    pPr = par._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    for edge in ("top", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "6")
        el.set(qn("w:color"), "B4C6E7")
        pBdr.append(el)
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), "1F3864")
    pBdr.append(left)
    pPr.insert_element_before(pBdr, *_P_AFTER_PBDR)
    run = par.add_run(text)
    run.italic = True
    run.font.size = Pt(10.5)
    return par


def mono_block(doc, lines: list[str]):
    par = doc.add_paragraph()
    par.paragraph_format.left_indent = Cm(0.6)
    par.paragraph_format.space_before = Pt(6)
    par.paragraph_format.space_after = Pt(6)
    par.paragraph_format.line_spacing = 1.0
    _pshade(par, "F4F4F4")
    for k, line in enumerate(lines):
        if k:
            par.add_run().add_break()
        run = par.add_run(line)
        run.font.name = MONO
        run.font.size = Pt(9.5)
    return par


def caption_above(doc, label: str, text: str):
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(10)
    par.paragraph_format.space_after = Pt(3)
    par.paragraph_format.keep_with_next = True
    r = par.add_run(f"{label}. ")
    r.bold = True
    r.font.name = SANS
    r.font.size = Pt(9.5)
    r.font.color.rgb = NAVY
    r = par.add_run(text)
    r.font.name = SANS
    r.font.size = Pt(9.5)
    return par


def caption_below(doc, text: str):
    par = doc.add_paragraph(style="Caption")
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(3)
    par.paragraph_format.space_after = Pt(10)
    rich(par, text)
    return par


def figure(doc, path: Path, width_cm: float, text: str):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(6)
    par.paragraph_format.keep_with_next = True
    par.add_run().add_picture(str(path), width=Cm(width_cm))
    caption_below(doc, text)


def _cell_border_bottom(cell, color="808080", sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    el = OxmlElement("w:bottom")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color)
    borders.append(el)
    tcPr.insert_element_before(borders, *_TC_AFTER_BORDER)


def _shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.insert_element_before(shd, *_TC_AFTER_SHD)


def table(doc, label: str, caption: str, headers, rows, widths,
          aligns=None, font=9.5):
    caption_above(doc, label, caption)
    t = doc.add_table(rows=1, cols=len(headers))
    t.autofit = False
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge, val, sz in (("top", "single", "12"), ("left", "none", "0"),
                          ("bottom", "single", "12"), ("right", "none", "0"),
                          ("insideH", "none", "0"), ("insideV", "none", "0")):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000" if val == "single" else "auto")
        borders.append(el)
    tblPr.insert_element_before(
        borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription", "w:tblPrChange")

    aligns = aligns or ["left"] * len(headers)
    amap = {"left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT}

    hdr = t.rows[0]
    trPr = hdr._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))   # repeat header if table splits
    for i, text in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = Cm(widths[i])
        _cell_border_bottom(cell)
        _shade(cell, "EDF1F8")
        par = cell.paragraphs[0]
        par.alignment = amap[aligns[i]]
        par.paragraph_format.space_after = Pt(1)
        par.paragraph_format.space_before = Pt(1)
        rich(par, str(text))
        _set_runs(par, SANS, Pt(font), NAVY, bold=True)

    for row in rows:
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cells[i].width = Cm(widths[i])
            par = cells[i].paragraphs[0]
            par.alignment = amap[aligns[i]]
            par.paragraph_format.space_after = Pt(1)
            par.paragraph_format.space_before = Pt(1)
            rich(par, str(value))
            _set_runs(par, SANS, Pt(font))
    return t


def new_section(doc, landscape: bool):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.different_first_page_header_footer = False
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
        sec.left_margin = sec.right_margin = Cm(2.0)
        sec.top_margin = sec.bottom_margin = Cm(2.2)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(2.5)
        sec.top_margin = sec.bottom_margin = Cm(2.5)
    return sec


# ---------------------------------------------------------------- document
def setup(doc: Document):
    st = doc.styles["Normal"]
    st.font.name = BODY_FONT
    st.font.size = Pt(11)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    pf = st.paragraph_format
    pf.line_spacing = 1.15
    pf.space_after = Pt(6)

    for name, size, color, before, after in (
        ("Heading 1", 15, NAVY, 16, 6),
        ("Heading 2", 12.5, STEEL, 12, 4),
        ("Heading 3", 11.5, STEEL, 10, 3),
    ):
        s = doc.styles[name]
        s.font.name = SANS
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = color
        s.font.italic = False
        s.paragraph_format.space_before = Pt(before)
        s.paragraph_format.space_after = Pt(after)
        s.paragraph_format.keep_with_next = True

    cap = doc.styles["Caption"]
    cap.font.name = SANS
    cap.font.size = Pt(9.5)
    cap.font.italic = True
    cap.font.color.rgb = GRAY

    # hyphenate justified text; ask Word to refresh fields (TOC) on open
    settings = doc.settings.element
    if settings.find(qn("w:defaultTabStop")) is not None:
        hyph = OxmlElement("w:autoHyphenation")
        hyph.set(qn("w:val"), "true")
        settings.find(qn("w:defaultTabStop")).addnext(hyph)
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    for tag in ("w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
                "w:docVars", "w:rsids", "w:themeFontLang", "w:shapeDefaults"):
        found = settings.find(qn(tag))
        if found is not None:
            found.addprevious(update)
            break

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.5)
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.header_distance = Cm(1.3)
    sec.footer_distance = Cm(1.4)
    sec.different_first_page_header_footer = True

    hdr = sec.header
    hdr.is_linked_to_previous = False
    hp = hdr.paragraphs[0]
    hp.text = ""
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hp.add_run("Data Mining  \u00b7  Case Study: Hospital Patient Records")
    r.font.name = SANS
    r.font.size = Pt(8.5)
    r.font.color.rgb = GRAY
    _pborder(hp, "bottom", "4", "BFBFBF", space="4")

    ftr = sec.footer
    ftr.is_linked_to_previous = False
    fp = ftr.paragraphs[0]
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(fp, " PAGE ")
    _set_runs(fp, SANS, Pt(9), GRAY)

    # cover: explicitly blank first-page header/footer
    sec.first_page_header.is_linked_to_previous = False
    sec.first_page_header.paragraphs[0].text = ""
    sec.first_page_footer.is_linked_to_previous = False
    sec.first_page_footer.paragraphs[0].text = ""


def cover(doc):
    for _ in range(2):
        doc.add_paragraph()
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run("CHANDIGARH UNIVERSITY")
    run.font.name = SANS
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = NAVY
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run("Subject: Data Mining")
    run.font.name = SANS
    run.font.size = Pt(12)
    run.font.color.rgb = GRAY
    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(6)
    rule.paragraph_format.space_after = Pt(24)
    _pborder(rule, "bottom", "12", "1F3864")

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run("Case Study")
    run.font.name = SANS
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = STEEL

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(6)
    run = par.add_run("Identifying Common Patterns in\nHospital Patient Records")
    run.font.name = SANS
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = NAVY

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(10)
    run = par.add_run(
        "How clustering and classification turn raw clinical records\n"
        "into patient groups and predictions")
    run.font.name = BODY_FONT
    run.font.size = Pt(13)
    run.font.italic = True
    run.font.color.rgb = GRAY

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(14)
    run = par.add_run("Case study topic 3")
    run.font.name = SANS
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = STEEL

    for _ in range(3):
        doc.add_paragraph()

    last = None
    for label, value in [
        ("Submitted by", "[Your Name]"),
        ("Roll number", "[Your Roll Number]"),
        ("Section / Session", "[Your Section]  /  2026-27"),
        ("Submitted to", "[Teacher's Name]"),
        ("Date of submission", "10 November 2026"),
    ]:
        last = doc.add_paragraph()
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        last.paragraph_format.space_after = Pt(4)
        r1 = last.add_run(f"{label}:  ")
        r1.font.name = SANS
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = NAVY
        r2 = last.add_run(value)
        r2.font.name = SANS
        r2.font.size = Pt(11)
        r2.font.italic = True
    return last   # TOC gets moved in right after this paragraph


def place_contents(doc, anchor_par, toc_pages: dict):
    """Create the TOC at the end, then move it behind the cover."""
    created = []

    par = doc.add_paragraph()          # spacer; page break ends the cover
    par.paragraph_format.page_break_before = True
    created.append(par)
    par = doc.add_paragraph()
    run = par.add_run("Contents")
    run.font.name = SANS
    run.font.size = Pt(17)
    run.font.bold = True
    run.font.color.rgb = NAVY
    par.paragraph_format.space_after = Pt(10)
    created.append(par)

    for level, title in HEADINGS:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0 if level == 1 else 0.7)
        p.paragraph_format.tab_stops.add_tab_stop(
            Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        r = p.add_run(title)
        r.font.name = SANS
        r.font.size = Pt(10.5)
        r.font.bold = level == 1
        r.font.color.rgb = NAVY if level == 1 else GRAY
        p.add_run("\t")
        num = str(toc_pages.get(title, ""))
        r = p.add_run(num)
        r.font.name = SANS
        r.font.size = Pt(10.5)
        r.font.bold = level == 1
        r.font.color.rgb = NAVY if level == 1 else GRAY
        created.append(p)

    # wrap entries in a real Word TOC field: begin/instr/separate ... end
    first, last = created[2], created[-1]
    for idx, frag in ((1, "begin"), (2, "instr"), (3, "separate")):
        r = OxmlElement("w:r")
        if frag == "instr":
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = ' TOC \\o "1-2" \\h \\z \\u '
            r.append(it)
        else:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), frag)
            r.append(fc)
        first._p.insert(idx, r)
    r = OxmlElement("w:r")
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "end")
    r.append(fc)
    last._p.append(r)

    # move from document end to directly after the cover
    prev = anchor_par._p
    for p in created:
        prev.addnext(p._p)
        prev = p._p


# ------------------------------------------------------------------ content
def build(toc_pages: dict) -> Document:
    r = json.loads(RESULTS.read_text())
    ds, cl, cn = r["dataset"], r["clustering"], r["classification"]
    met = cn["metrics"]
    prof, rate = cl["profiles"], cl["disease_rate_pct"]
    dist, ex = cl["worked_example"], cn["tree_example"]
    sm = cl["scaler_mean"]
    ss = cl["scaler_scale"]
    HEADINGS.clear()

    doc = Document()
    setup(doc)
    doc.core_properties.title = "Case Study: Hospital Patient Records"
    doc.core_properties.subject = "Data Mining"
    doc.core_properties.author = "Data Mining case study"

    anchor = cover(doc)

    # ---------------------------------------------------- 1. Introduction
    intro = h1(doc, "1. Introduction")
    intro.paragraph_format.page_break_before = True   # body starts on a fresh page
    body(doc,
         "Hospitals sit on records that nobody has time to read one by one. "
         "A cardiology department keeps, for every admission, the patient's age, "
         "resting blood pressure, cholesterol, how fast the heart recovered "
         "during exercise, and eventually a diagnosis. Piled up over years these "
         "rows become hard to interpret: too many to eyeball, and a diagnosis "
         "column on its own says nothing about *who* the patients are as "
         "groups.")
    body(doc, "This case study answers a simple assignment question:")
    quote_box(doc,
              "\u201cA hospital wants to analyze patient records to identify "
              "common patterns. Explain how classification or clustering can "
              "help.\u201d   (Case study topic 3)")
    body(doc,
         "We work with real clinical data rather than an invented example, and "
         "we run both families of methods the question mentions. Clustering "
         "groups the patients without ever looking at their diagnoses, which "
         "suits the phrase *common patterns*: it asks what structure exists in "
         "the data, not who fell ill. Classification then learns from past "
         "diagnoses and judges new patients one at a time. Every number "
         "reported below came out of the accompanying code in this repository, "
         "so any table can be checked or re-run.")
    body(doc, "By the end of the report the reader should be able to:")
    bullet(doc, "describe the dataset, how it was cleaned, and why each "
                "preprocessing step was needed;")
    bullet(doc, "explain in plain words how K-means and a decision tree work, "
                "and why these two were picked;")
    bullet(doc, "read the two patient groups the algorithm found and say what "
                "they mean clinically;")
    bullet(doc, "evaluate the classifier honestly, including the cases it gets "
                "wrong, and")
    bullet(doc, "translate both results into concrete hospital use, which is "
                "the actual assignment question.")

    # -------------------------------------------------------- 2. Dataset
    h1(doc, "2. About the dataset")
    h2(doc, "2.1 Source and variables")
    body(doc,
         "The data is the processed Cleveland set from the UCI Machine Learning "
         "Repository, collected at the Cleveland Clinic Foundation and first "
         "described by Detrano and colleagues in 1988 [2, 3]. It contains "
         f"**{ds['records']} patient records**, {ds['raw_attributes']} clinical "
         "variables and one diagnosis column. The diagnosis is stored as 0 to "
         "4; following the usual convention we treat any value above 0 as heart "
         "disease present, which gives a clean yes/no target: "
         f"**{r['cleaning']['class_counts']['no_disease']} patients without "
         f"disease, {r['cleaning']['class_counts']['disease']} with it**. The "
         "set is small enough to reason about line by line and openly "
         "available, which is exactly what a course project needs.")
    table(doc, "Table 1", "The fourteen variables recorded for each patient.",
          ["Attribute", "Meaning", "Range / values in this data"],
          [
              ["age", "Age of the patient", "29 to 77 years"],
              ["sex", "Sex", "1 = male, 0 = female (206 male, 97 female)"],
              ["cp", "Chest pain type",
               "1 typical, 2 atypical, 3 non-anginal, 4 asymptomatic"],
              ["trestbps", "Resting blood pressure", "94 to 200 mm Hg"],
              ["chol", "Serum cholesterol", "126 to 564 mg/dl"],
              ["fbs", "Fasting blood sugar > 120 mg/dl", "1 = true, 0 = false"],
              ["restecg", "Resting electrocardiogram", "0, 1, 2"],
              ["thalach", "Maximum heart rate achieved", "71 to 202 bpm"],
              ["exang", "Exercise-induced angina", "1 = yes, 0 = no"],
              ["oldpeak", "ST depression from exercise", "0.0 to 6.2"],
              ["slope", "Slope of peak exercise ST segment", "1, 2, 3"],
              ["ca", "Major vessels seen by fluoroscopy",
               "0 to 3 (4 missing values)"],
              ["thal", "Thalassemia test result",
               "3 = normal, 6 = fixed, 7 = reversible"],
              ["num", "Diagnosis of heart disease", "0 = no, 1-4 = yes (target)"],
          ],
          [3.0, 6.0, 6.5], font=9)

    h2(doc, "2.2 A first look at the records")
    body(doc,
         "The first ten records, already cleaned, look like this. One detail "
         "worth noticing early: the last column is the answer we want to "
         "predict, and in a real workflow it would stay hidden while the model "
         "is applied.")
    sample = r["sample_rows"][:10]
    table(doc, "Table 2", "The first ten patient records after cleaning.",
          ["Age", "Sex", "Chest pain", "BP", "Chol", "Max HR", "Disease"],
          [[s["age"], s["sex"], s["cp"], s["trestbps"], s["chol"],
            s["thalach"], s["disease"]] for s in sample],
          [1.6, 1.4, 4.0, 2.0, 2.0, 2.0, 2.5],
          aligns=["center", "center", "left", "center", "center",
                  "center", "center"], font=9)

    h2(doc, "2.3 What the numbers look like")
    body(doc,
         "Five variables are continuous and carry most of the physiological "
         "signal, so we summarise them first. Cholesterol has the widest spread "
         "and one obvious outlier at 564 mg/dl; maximum heart rate varies by "
         "more than 130 beats per minute across patients. Such differences in "
         "scale matter later, because K-means measures distance.")
    pretty = {"age": "age", "trestbps": "resting BP", "chol": "cholesterol",
              "thalach": "max heart rate", "oldpeak": "ST depression"}
    table(doc, "Table 3",
          "Descriptive statistics of the five continuous variables used for "
          "clustering.",
          ["Variable", "Mean", "Std. dev.", "Min", "Max"],
          [[pretty.get(d["attribute"], d["attribute"]), f"{d['mean']:.1f}",
            f"{d['std']:.1f}", f"{d['min']:.0f}", f"{d['max']:.0f}"]
           for d in r["descriptives"]],
          [3.5, 3.0, 3.0, 3.0, 3.0],
          aligns=["left", "right", "right", "right", "right"], font=9.5)

    # --------------------------------------------------- 3. Preprocessing
    h1(doc, "3. Preparing the data")
    body(doc,
         "Raw records rarely arrive ready for mining, and the steps taken here "
         "are deliberately ordinary: nothing clever, only the changes a careful "
         "analyst would make anyway.")
    bullet(doc, "**Missing values.** Six cells are empty: `ca` (vessels seen by "
                "fluoroscopy) is missing in 4 records and `thal` in 2. That is "
                "0.15% of all cells. We fill them with the median of each "
                f"column (ca = {r['cleaning']['median_imputed']['ca']:.0f}, "
                f"thal = {r['cleaning']['median_imputed']['thal']:.0f}). The "
                "median is robust to the cholesterol outlier, and dropping six "
                "rows out of 303 would throw away real patients for no gain.")
    bullet(doc, "**Duplicates.** Exactly "
                f"{r['cleaning']['duplicate_rows']} duplicate rows were found, "
                "so nothing was removed.")
    bullet(doc, "**Encoding categories.** Variables such as chest pain type or "
                "thalassemia result are categories, not numbers that can be "
                "averaged. One-hot encoding expands the "
                f"{ds['raw_attributes']} variables into "
                f"{met['n_features_after_encoding']} binary columns, so a "
                "split like *chest pain is asymptomatic* becomes a yes/no test.")
    bullet(doc, "**Standardisation.** For clustering each variable is rescaled "
                "to z-scores, (x \u2212 mean) / std. Without this, cholesterol "
                "(mean 247) would dominate age (mean 54) purely because of its "
                "units, and distance would mean the wrong thing.")
    bullet(doc, "**Train/test split.** The labelled data is split 70/30 into "
                f"{met['train_size']} training and {met['test_size']} test "
                "records, stratified so both halves keep the same disease "
                "proportion, with `random_state=42` so the numbers in this "
                "report reproduce exactly. The test set is touched only at the "
                "end, to report honest performance.")

    # ---------------------------------------------------- 4. Methodology
    h1(doc, "4. How we approached the problem")
    body(doc,
         "The work follows the usual CRISP-DM sequence [5]: understand the "
         "question, understand the data, prepare it (sections 2 and 3), model, "
         "then evaluate. The modelling phase uses two algorithms, chosen "
         "because they answer two different halves of the assignment question.")
    h2(doc, "4.1 Why clustering, and why K-means")
    body(doc,
         "Clustering is unsupervised learning: the algorithm never sees a "
         "diagnosis. It partitions the patients into k groups so that points "
         "inside a group sit close to their group's centre, repeating two steps "
         "until nothing moves: assign each patient to the nearest centroid, "
         "then recompute each centroid as the mean of its members [1, 6]. "
         "K-means was picked over hierarchical clustering because the "
         "hospital's question is \u201cwhat groups exist\u201d rather than a "
         "full family tree of every pair of patients, and over DBSCAN because "
         "these physiological measures form broad clouds rather than the "
         "tight, unevenly-dense blobs DBSCAN expects. The only real decision "
         "is k, and it is made from data: we sweep k = 1 to 8, watch how "
         "within-cluster inertia drops (the elbow), and confirm with the "
         "silhouette score, which averages how similar each patient is to its "
         "own group versus the neighbouring group [4]. Values near 1 mean "
         "crisp groups; near 0, overlapping ones.")
    h2(doc, "4.2 Why classification, and why a decision tree")
    body(doc,
         "Classification is supervised: the tree reads past patients whose "
         "diagnoses are known and learns a chain of yes/no questions that "
         "separates the two classes, much like a doctor's rule of thumb. Each "
         "split picks the question that makes the resulting boxes as pure as "
         "possible (Gini impurity), and recursion stops at depth 4 or once a "
         "box holds fewer than 8 patients, which keeps the tree from memorising "
         "noise in only 212 training rows [1, 7]. We chose a tree over naive "
         "Bayes because clinical variables are strongly correlated (age, blood "
         "pressure, maximum heart rate) and naive Bayes assumes they are not, "
         "and over logistic regression mainly because a tree hands the hospital "
         "readable rules instead of coefficients to interpret. Its "
         "straightforward weakness is also worth stating up front: a single "
         "tree is unstable around its training data, so we report 10-fold "
         "cross-validation alongside the single test score.")
    h2(doc, "4.3 Reproducing these results")
    body(doc,
         "Everything lives in this repository: `data/raw/` holds the UCI "
         "download, `src/analyze.py` performs the whole pipeline and writes "
         "`reports/results.json` plus four figures, and `src/build_report.py` "
         "turns those results into this document. Running the two scripts in "
         "that order regenerates every table and number here; seeds are fixed, "
         "so the output is identical on any machine.")

    # ----------------------------------------------- 5. Clustering results
    h1(doc, "5. Finding the patient groups")
    h2(doc, "5.1 Choosing the number of groups")
    body(doc,
         "Sweeping k from 1 to 8, inertia drops sharply when going from one "
         f"group to two (from {cl['k_inertia']['1']:.0f} to "
         f"{cl['k_inertia']['2']:.0f}) and then flattens into a slow decline, "
         "the classic elbow shape. The silhouette score agrees: it peaks at "
         f"k = 2 with {cl['silhouette']:.2f}, marginally above k = 3 "
         f"({cl['k_silhouette']['3']:.2f}) and clearly above k = 8 "
         f"({cl['k_silhouette']['8']:.2f}). So two groups it is. The score of "
         "0.24 is honest to report as modest: these are tendencies in the "
         "data, not two sharply separated species of patient, and section 5.3 "
         "shows why the split is still useful.")
    figure(doc, FIG / "fig1_elbow_silhouette.png", 15.0,
           "**Figure 1.** Elbow curve (left) and silhouette score (right) for "
           "k = 1 to 8. Both point at k = 2.")
    h2(doc, "5.2 The two groups side by side")
    body(doc,
         "With two groups fixed, the profiles below are just the mean of each "
         "group. The contrast is easy to state in one sentence: group 1 is the "
         "older, more stressed heart; group 2 the younger, quieter one.")
    table(doc, "Table 4",
          "The two groups K-means found. The diagnosis column was not used to "
          "build the groups; it is shown afterwards to interpret them.",
          ["Group", "n", "Age", "Resting BP", "Chol", "Max HR", "ST dep.",
           "Diagnosed"],
          [
              ["Group 1", prof[0]["size"], f"{prof[0]['age']:.1f}",
               f"{prof[0]['trestbps']:.1f}", f"{prof[0]['chol']:.1f}",
               f"{prof[0]['thalach']:.1f}", f"{prof[0]['oldpeak']:.2f}",
               f"{rate['0']:.1f}%"],
              ["Group 2", prof[1]["size"], f"{prof[1]['age']:.1f}",
               f"{prof[1]['trestbps']:.1f}", f"{prof[1]['chol']:.1f}",
               f"{prof[1]['thalach']:.1f}", f"{prof[1]['oldpeak']:.2f}",
               f"{rate['1']:.1f}%"],
          ],
          [2.2, 1.1, 1.7, 2.2, 2.0, 2.0, 1.7, 2.6],
          aligns=["left", "center", "right", "right", "right", "right",
                  "right", "right"], font=9)
    h2(doc, "5.3 The groups line up with the diagnoses")
    body(doc,
         "Here is the part that answers *identify common patterns*. Group 1 "
         f"({prof[0]['size']} patients) averages "
         f"{prof[0]['age'] - prof[1]['age']:.0f} years older, runs about "
         f"{prof[0]['trestbps'] - prof[1]['trestbps']:.0f} mm Hg higher at "
         f"rest, carries more cholesterol, reaches "
         f"{prof[1]['thalach'] - prof[0]['thalach']:.0f} bpm less during "
         f"exercise, and shows roughly three times the ST depression "
         f"({prof[0]['oldpeak']:.2f} against {prof[1]['oldpeak']:.2f}). After "
         "the fact, and without the algorithm ever seeing a diagnosis, "
         f"{rate['0']:.1f}% of group 1 turns out to have heart disease against "
         f"only {rate['1']:.1f}% of group 2. That gap is the pattern the "
         "hospital was looking for: a risk profile described purely from "
         "measurements.")
    figure(doc, FIG / "fig2_clusters.png", 13.0,
           f"**Figure 2.** The same 303 patients projected onto the first two "
           f"principal components ({cl['pca_explained_variance_pct'][0]}% + "
           f"{cl['pca_explained_variance_pct'][1]}% of the variance), coloured "
           "by group. The clouds overlap, which is why the silhouette score "
           "is modest.")

    # -------------------------------------------- 6. Classification results
    h1(doc, "6. Predicting heart disease")
    h2(doc, "6.1 How well the tree performs")
    body(doc,
         f"The tree was trained on {met['train_size']} records and evaluated on "
         f"the {met['test_size']} it had never seen. It answers correctly "
         f"**{met['accuracy']*100:.1f}%** of the time, against "
         f"{met['baseline_majority']*100:.1f}% for a lazy model that always "
         "guesses the majority class, so the tree earns its keep. Precision and "
         "recall split that headline number in the two directions a hospital "
         "cares about: of the patients the tree flags, "
         f"{met['precision']*100:.1f}% really have the disease, and of the "
         f"patients who really have it, {met['recall']*100:.1f}% are caught. "
         f"Ten-fold cross-validation on the full set gives "
         f"{met['cv10_mean']*100:.1f}% \u00b1 {met['cv10_std']*100:.1f}%, a "
         "spread that honestly reflects how much the score moves with a "
         "different split.")
    tn, fp = met["confusion"][0]
    fn, tp = met["confusion"][1]
    table(doc, "Table 5",
          "Decision tree performance on the 91 held-out test patients.",
          ["Metric", "Value", "What it means"],
          [
              ["Accuracy", f"{met['accuracy']*100:.1f}%",
               f"{tn + tp} of {met['test_size']} patients judged correctly "
               f"(baseline {met['baseline_majority']*100:.1f}%)"],
              ["Precision", f"{met['precision']*100:.1f}%",
               f"Of {tp + fp} flagged patients, {tp} truly have the disease"],
              ["Recall", f"{met['recall']*100:.1f}%",
               f"Of {tp + fn} actual cases, {tp} were caught; {fn} were missed"],
              ["F1 score", f"{met['f1']*100:.1f}%",
               "Balance of precision and recall"],
              ["10-fold CV", f"{met['cv10_mean']*100:.1f}%",
               f"Average over 10 re-splits: \u00b1{met['cv10_std']*100:.1f}%"],
          ],
          [3.2, 2.4, 9.9], aligns=["left", "right", "left"], font=9.5)
    figure(doc, FIG / "fig3_confusion_matrix.png", 9.5,
           f"**Figure 3.** Confusion matrix on the test set: {tn} correct "
           f"\u201cno disease\u201d, {tp} correct \u201cdisease\u201d, {fp} "
           f"false alarms, {fn} missed cases.")
    body(doc,
         f"The {fn} missed cases and {fp} false alarms deserve a sentence "
         "each. A missed case means a sick patient sent home; a false alarm "
         "means an extra test. A hospital that prefers the first error over "
         "the second would later tune the decision threshold towards recall, "
         "but at this stage the honest reading is: useful signal, not a "
         "verdict.")
    h2(doc, "6.2 What the tree actually learned")
    body(doc,
         "Feature importances show where the tree's decisions came from. The "
         "thalassemia result alone accounts for almost half of them, followed "
         "by the chest pain type and the number of visible vessels:")
    plain = {"thal_3.0": "Whether the thalassemia scan was normal",
             "cp_4.0": "Whether chest pain was asymptomatic",
             "ca_0.0": "Whether fluoroscopy saw any major vessel",
             "age": "Patient age in years",
             "oldpeak": "Exercise-induced ST depression",
             "thalach": "Maximum heart rate reached in exercise"}
    table(doc, "Table 6",
          "Features the tree relied on most (importances sum to 1 across all "
          "features).",
          ["Feature", "Importance", "In plain words"],
          [[f["feature"], f"{f['importance']*100:.0f}%",
            plain.get(f["feature"], "\u2014")]
           for f in cn["top_features"]],
          [3.4, 2.4, 9.7], aligns=["left", "right", "left"], font=9.5)
    body(doc,
         "Read as rules rather than percentages, the tree says things a doctor "
         "would recognise:")
    mono_block(doc, [
        "IF thalassemia abnormal AND chest pain asymptomatic  -> disease",
        "IF thalassemia normal AND pain NOT asymptomatic      -> no disease",
        "IF thal normal AND asymptomatic AND a vessel visible -> disease",
        "IF thal normal AND asymptomatic AND no vessel AND age <= 59 -> no",
        "                               ...but disease if older than 59",
    ])
    body(doc,
         "That is the practical form of the assignment answer: the classifier "
         "converts 212 past patients into a page of rules that can be applied "
         "to the next admission in milliseconds.")

    # landscape page for the wide tree figure
    new_section(doc, landscape=True)
    figure(doc, FIG / "fig4_decision_tree.png", 24.5,
           "**Figure 4.** The full learned tree (depth 4, 13 leaves). Orange "
           "nodes lean towards \u201cno disease\u201d, blue towards "
           "\u201cdisease\u201d; each node shows Gini impurity, patient count "
           "and class.")
    new_section(doc, landscape=False)

    # ---------------------------------------------- 7. Worked examples
    h1(doc, "7. Two worked examples, step by step")
    h2(doc, "7.1 Assigning a patient to a group (clustering)")
    p0 = dist["patient"]
    body(doc,
         "Take the very first patient in the file: a "
         f"{p0['age']:.0f}-year-old man with resting BP {p0['trestbps']:.0f}, "
         f"cholesterol {p0['chol']:.0f}, maximum heart rate "
         f"{p0['thalach']:.0f} and ST depression {p0['oldpeak']:.1f}. Each "
         "value is standardised with the mean and standard deviation of the "
         "data set, then the Euclidean distance to each group centre is taken:")
    feats = ["age", "trestbps", "chol", "thalach", "oldpeak"]
    labels = ["age", "resting BP", "cholesterol", "max HR", "oldpeak"]
    mono_block(doc, [
        f"z({lab:<12}) = ({p0[f]:g} - {sm[f]}) / {ss[f]} = "
        f"{(p0[f] - sm[f]) / ss[f]:+.2f}"
        for lab, f in zip(labels, feats)
    ])
    c1, c2 = cl["centroids_scaled"]
    z = dist["scaled"]
    table(doc, "Table 7",
          "Distance of the standardised patient to each group centre "
          "(squared terms; z rounded to 2 dp).",
          ["Term", "Patient z", "Centre z (G1)", "(z\u2212c1)\u00b2",
           "Centre z (G2)", "(z\u2212c2)\u00b2"],
          [[lab, f"{z[k]:.2f}", f"{c1[k]:.3f}", f"{(z[k] - c1[k])**2:.3f}",
            f"{c2[k]:.3f}", f"{(z[k] - c2[k])**2:.3f}"]
           for k, lab in enumerate(labels)]
          + [["distance", "", "", f"**{dist['distances'][0]:.3f}**", "",
              f"**{dist['distances'][1]:.3f}**"]],
          [2.8, 2.3, 2.8, 2.4, 2.8, 2.4],
          aligns=["left", "right", "right", "right", "right", "right"], font=9)
    body(doc,
         f"(The distance is the square root of the sum of the squared "
         f"columns.) The patient is {dist['distances'][0]:.3f} from the centre "
         f"of group 1 and {dist['distances'][1]:.3f} from group 2, so he lands "
         "in group 1, the older and higher-risk one. Note what this does *not* "
         f"mean: group 1 is {rate['0']:.1f}% diagnosed, not 100%, and this "
         "particular patient was in fact healthy. A group is a tendency, which "
         "is precisely why the hospital also wants the classifier for "
         "individual judgement.")
    h2(doc, "7.2 Classifying a test patient (decision tree)")
    body(doc,
         "For the second example we deliberately pick a patient the tree never "
         f"trained on (record {ex['record_index']} of the test set): a "
         f"{ex['raw']['age']:.0f}-year-old man with asymptomatic chest pain, "
         f"resting BP {ex['raw']['trestbps']:.0f}, cholesterol "
         f"{ex['raw']['chol']:.0f}, maximum heart rate "
         f"{ex['raw']['thalach']:.0f}, ST depression "
         f"{ex['raw']['oldpeak']:.1f}, {ex['raw']['ca']:.0f} visible vessel "
         "and a reversible thalassemia defect. The tree walks four tests and "
         "stops at a leaf:")
    steps = []
    for k, s in enumerate(ex["path"], 1):
        op = "\u2264" if s["go"] == "left" else ">"
        steps.append([f"Step {k}", s["feature"].replace("_", " "),
                      f"{s['value']:g} {op} {s['threshold']:g}",
                      f"go {'left' if s['go'] == 'left' else 'right'}"])
    steps.append(["Leaf", "prediction", "\u2014",
                  "**Disease** (actual: Disease) \u2014 correct"])
    table(doc, "Table 8",
          "The four tests the tree applies to the test patient.",
          ["", "Test", "Comparison", "Result"],
          steps, [2.0, 4.0, 4.4, 5.1], font=9.5)
    body(doc,
         "The reasoning is clinical rather than statistical: the thalassemia "
         "scan was not normal, the pain was asymptomatic, the ST depression of "
         "2.8 far exceeds the 0.45 threshold, and cholesterol above 202 pushes "
         "the final branch to disease. Four questions, one answer, well under "
         "a millisecond \u2014 that is the form in which classification helps.")

    # ---------------------------------------- 8. Interpretation (the ask)
    h1(doc, "8. How classification and clustering help the hospital")
    body(doc,
         "The assignment asks one direct question, so here is the direct "
         "answer, grounded in what the previous sections actually produced.")
    body(doc,
         "**Clustering helps by describing the population without needing "
         "labels.** The two groups it found are a usable segmentation: a "
         f"higher-risk cohort of {prof[0]['size']} patients averaging "
         f"{prof[0]['age']:.0f} years with a reduced exercise heart rate, "
         f"where {rate['0']:.1f}% already show heart disease, and a "
         f"lower-risk cohort of {prof[1]['size']}. A hospital can aim a "
         "screening camp at the profile of group 1 before any diagnosis "
         "exists, plan cardiology capacity around the expected split, or "
         "compare next year's admissions against these baselines. Because "
         "clustering needs no diagnosis column it also works in departments "
         "where labels are expensive, delayed, or missing altogether.")
    body(doc,
         "**Classification helps by judging the individual patient.** Once "
         "diagnoses exist, the tree converts them into rules that score a new "
         f"admission at {met['accuracy']*100:.1f}% accuracy "
         f"({met['cv10_mean']*100:.1f}% under cross-validation), far above the "
         f"{met['baseline_majority']*100:.1f}% of guessing. Used as a triage "
         "aid it can flag high-risk arrivals for an earlier test, rank "
         "referrals for a specialist's limited time, or prompt a second look "
         "when the record says low risk but the tree says otherwise. Its "
         f"recall of {met['recall']*100:.1f}% means roughly three in four "
         "actual cases are caught at admission.")
    body(doc,
         "**The two answer different halves of one question.** Clustering "
         "explains what the patient population looks like (the common "
         "patterns); classification says what to expect from the next "
         "individual. Used together, the group structure can even be fed back "
         "as an extra feature to the classifier, which is the natural next "
         "step if this project were extended. Both remain decision *support*: "
         f"with roughly one prediction in four wrong, no result here overrides "
         "a cardiologist, and any real deployment would be validated on the "
         "hospital's own current patients first.")

    # ----------------------------------------------------- 9. Limitations
    h1(doc, "9. Limitations")
    bullet(doc, "**Data age and size.** 303 patients from a single clinic in "
                "1988 is a small, dated sample; modern populations, equipment "
                "and treatment may shift every relationship in this report.")
    bullet(doc, "**Soft groups.** A silhouette of "
                f"{cl['silhouette']:.2f} means the two groups overlap "
                "noticeably. They are a summary, not a diagnosis, and some "
                "patients sit between them.")
    bullet(doc, "**Honest error rate.** At "
                f"{met['accuracy']*100:.1f}% accuracy the tree is wrong about "
                f"1 patient in 4, including {fn} missed disease cases in a "
                "test set of 91. That bounds its role to supporting, never "
                "replacing, a clinician.")
    bullet(doc, "**Simplified preparation.** Median imputation for the six "
                "missing values and one random 70/30 split are defensible but "
                "crude; the \u00b1"
                f"{met['cv10_std']*100:.1f}% cross-validation spread shows "
                "how much a single split can move the numbers.")
    bullet(doc, "**No causation, no history.** The variables are snapshots "
                "without medication, lifestyle or family history, so the tree "
                "correlates rather than explains, and no claim here is causal.")

    # ---------------------------------------------------- 10. Conclusion
    h1(doc, "10. Conclusion")
    body(doc,
         "The hospital's question has two answers and we ran both on the same "
         "303 real records. K-means split the patients into an older, "
         "higher-stress group and a younger one; the diagnosis rates of "
         f"{rate['0']:.1f}% against {rate['1']:.1f}% appeared afterwards, "
         "confirming that the pattern the algorithm found is the clinically "
         "meaningful one. The decision tree turned the labelled history into "
         f"{cn['tree_depth']} levels of yes/no rules that classify unseen "
         f"patients with {met['accuracy']*100:.1f}% accuracy against a "
         f"{met['baseline_majority']*100:.1f}% baseline. Together they cover "
         "the assignment question end to end: clustering tells the hospital "
         "what its patient population looks like, classification tells it what "
         "to expect about the next patient, and the limitations section marks "
         "where both stop being decision support and start needing a doctor's "
         "judgement.")

    # --------------------------------------------------- 11. References
    h1(doc, "11. References")
    refs = [
        "Han, J., Kamber, M., Pei, J. *Data Mining: Concepts and Techniques*, "
        "3rd ed. Morgan Kaufmann, 2012.",
        "Detrano, R., Janosi, A., Steinbrunn, W., et al. \u201cInternational "
        "application of a new probability algorithm for the diagnosis of "
        "coronary artery disease.\u201d *American Journal of Cardiology*, "
        "62(8), 1988, pp. 504-510. (Source of the UCI Heart Disease data)",
        "UCI Machine Learning Repository: Heart Disease. "
        "https://archive.ics.uci.edu/dataset/45/heart+disease",
        "Rousseeuw, P. J. \u201cSilhouettes: A graphical aid to the "
        "interpretation and validation of cluster analysis.\u201d *Journal of "
        "Computational and Applied Mathematics*, vol. 20, 1987, pp. 53-65.",
        "Wirth, R., Hipp, J. \u201cCRISP-DM: Towards a process model for data "
        "mining.\u201d *Proceedings of the 4th International Conference on "
        "Knowledge Discovery and Data Mining*, 2000.",
        "MacQueen, J. \u201cSome methods for classification and analysis of "
        "multivariate observations.\u201d *Proceedings of the 5th Berkeley "
        "Symposium on Mathematical Statistics and Probability*, vol. 1, 1967, "
        "pp. 281-297.",
        "Breiman, L., Friedman, J. H., Olshen, R. A., Stone, C. J. "
        "*Classification and Regression Trees*. Wadsworth, 1984.",
        "Pedregosa, F., et al. \u201cScikit-learn: Machine Learning in "
        "Python.\u201d *Journal of Machine Learning Research*, vol. 12, 2011, "
        "pp. 2825-2830.",
    ]
    for ref in refs:
        par = doc.add_paragraph()
        par.paragraph_format.left_indent = Cm(0.9)
        par.paragraph_format.first_line_indent = Cm(-0.9)
        par.paragraph_format.space_after = Pt(6)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        rich(par, ref)
        _set_runs(par, BODY_FONT, Pt(10))

    place_contents(doc, anchor, toc_pages)
    return doc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--toc-pages", type=Path, default=None,
                    help="JSON mapping heading title -> printed page number")
    ap.add_argument("--out", type=Path, default=OUT)
    args = ap.parse_args()
    toc_pages = {}
    if args.toc_pages and args.toc_pages.exists():
        toc_pages = json.loads(args.toc_pages.read_text())
    doc = build(toc_pages)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.out)
    print(f"wrote {args.out} ({args.out.stat().st_size} bytes)")


if __name__ == "__main__":
    main()

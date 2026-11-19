# -*- coding: utf-8 -*-
"""Допоміжні функції: PDF (reportlab), DOCX (python-docx), XLSX (openpyxl), MD.

Усі згенеровані PDF/DOCX несуть позначку «ДЕМО · вигадані дані».
"""
import os
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                Table, TableStyle)

DEMO_NOTE = "ДЕМО-ДОКУМЕНТ · усі назви, особи, коди, суми й номери вигадані"
FONT = "DV"
FONT_B = "DV-B"


def init_fonts(font_dir):
    d = Path(font_dir)
    pdfmetrics.registerFont(TTFont(FONT, str(d / "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont(FONT_B, str(d / "DejaVuSans-Bold.ttf")))
    pdfmetrics.registerFontFamily(FONT, normal=FONT, bold=FONT_B, italic=FONT, boldItalic=FONT_B)


def _styles():
    base = dict(fontName=FONT, fontSize=9.5, leading=13.5)
    return {
        "p": ParagraphStyle("p", alignment=TA_JUSTIFY, spaceAfter=4, **base),
        "pl": ParagraphStyle("pl", spaceAfter=3, **base),
        "h1": ParagraphStyle("h1", fontName=FONT_B, fontSize=14, leading=18, alignment=TA_CENTER, spaceAfter=8),
        "h2": ParagraphStyle("h2", fontName=FONT_B, fontSize=10.5, leading=14, spaceBefore=6, spaceAfter=3),
        "c": ParagraphStyle("c", alignment=TA_CENTER, spaceAfter=3, **base),
        "cell": ParagraphStyle("cell", fontName=FONT, fontSize=8.5, leading=11),
        "cellb": ParagraphStyle("cellb", fontName=FONT_B, fontSize=8.5, leading=11),
    }


def _page(canvas, doc):
    canvas.saveState()
    w, h = A4
    canvas.setFont(FONT_B, 54)
    canvas.setFillColor(colors.Color(0.5, 0.5, 0.5, alpha=0.09))
    canvas.translate(w / 2, h / 2)
    canvas.rotate(35)
    canvas.drawCentredString(0, 0, "ДЕМО")
    canvas.restoreState()
    canvas.saveState()
    canvas.setFont(FONT, 7)
    canvas.setFillColor(colors.grey)
    canvas.drawCentredString(w / 2, 9 * mm, f"{DEMO_NOTE} · стор. {doc.page}")
    canvas.restoreState()


def redact(text, on):
    """Закрити суму чорним прямокутником (як «замазано»)."""
    if not on:
        return text
    return '<span backColor="#000000" color="#000000">' + ("█" * max(6, len(text))) + "</span>"


def pdf(path, blocks, title=""):
    """blocks: ('h1',t) ('h2',t) ('p',t) ('pl',t) ('c',t) ('sp',) ('pb',) ('table',header,rows,widths_mm)"""
    st = _styles()
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=20 * mm, rightMargin=18 * mm,
                            topMargin=18 * mm, bottomMargin=18 * mm, title=title or Path(path).stem,
                            author="Демо-генератор", subject=DEMO_NOTE)
    flow = []
    for b in blocks:
        k = b[0]
        if k in ("h1", "h2", "p", "pl", "c"):
            flow.append(Paragraph(b[1], st[k]))
        elif k == "sp":
            flow.append(Spacer(1, 6))
        elif k == "pb":
            flow.append(PageBreak())
        elif k == "table":
            _, header, rows, widths = b
            data = []
            if header:
                data.append([Paragraph(str(c), st["cellb"]) for c in header])
            for r in rows:
                data.append([Paragraph(str(c), st["cell"]) for c in r])
            t = Table(data, colWidths=[w * mm for w in widths] if widths else None, repeatRows=1 if header else 0)
            style = [("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                     ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
            if header:
                style.append(("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.9, 0.92, 0.95)))
            t.setStyle(TableStyle(style))
            flow.append(t)
            flow.append(Spacer(1, 6))
    doc.build(flow, onFirstPage=_page, onLaterPages=_page)


# ------------------------------------------------------------------ DOCX
def docx_file(path, blocks, landscape=False, note=True):
    """blocks: ('title',t) ('h1',t) ('h2',t) ('p',t) ('bullets',[..]) ('table',header,rows) ('note',t)"""
    import docx
    from docx.enum.section import WD_ORIENT
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    d = docx.Document()
    sec = d.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for m in ("left_margin", "right_margin"):
        setattr(sec, m, Cm(2))
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    stl = d.styles["Normal"]
    stl.font.name = "Calibri"
    stl.font.size = Pt(10.5)
    if note:
        fp = sec.footer.paragraphs[0]
        fp.text = DEMO_NOTE
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in fp.runs:
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    def shade(cell, color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    for b in blocks:
        k = b[0]
        if k == "title":
            p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(b[1]); r.bold = True; r.font.size = Pt(16)
        elif k == "h1":
            d.add_heading(b[1], level=1)
        elif k == "h2":
            d.add_heading(b[1], level=2)
        elif k == "p":
            d.add_paragraph(b[1])
        elif k == "note":
            p = d.add_paragraph(); r = p.add_run(b[1]); r.italic = True; r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
        elif k == "bullets":
            for it in b[1]:
                d.add_paragraph(it, style="List Bullet")
        elif k == "table":
            _, header, rows = b[:3]
            ncols = len(header) if header else len(rows[0])
            t = d.add_table(rows=0, cols=ncols)
            t.style = "Table Grid"
            if header:
                cells = t.add_row().cells
                for i, h in enumerate(header):
                    cells[i].text = ""
                    run = cells[i].paragraphs[0].add_run(str(h)); run.bold = True; run.font.size = Pt(9.5)
                    shade(cells[i], "E6EBF2")
            for r_ in rows:
                cells = t.add_row().cells
                for i, v in enumerate(r_):
                    cells[i].text = ""
                    run = cells[i].paragraphs[0].add_run(str(v)); run.font.size = Pt(9.5)
            d.add_paragraph()
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    d.core_properties.author = "Демо-генератор"
    d.core_properties.comments = DEMO_NOTE
    d.save(str(path))


# ------------------------------------------------------------------ MD
def md(path, text):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(text.strip("\n") + "\n", encoding="utf-8", newline="\n")


def keep(path, note=""):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(note, encoding="utf-8", newline="\n")


def money(v, cur="грн"):
    s = f"{v:,.2f}".replace(",", " ").replace(".", ",")
    return f"{s} {cur}"

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка отчёта по практической работе № 1 (РОСА Фреш на VMware) в форматах DOCX и PDF.

Скриншоты берутся из ../screenshots, результат складывается в ../report.
Требуется: python-docx, reportlab, Pillow (см. README).
"""

import os
import sys

from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(BASE)
SHOTS = os.path.join(LAB, 'screenshots')
REPORT_DIR = os.path.join(LAB, 'report')
FONTS = os.path.join(BASE, 'fonts')

sys.path.insert(0, BASE)
import content  # noqa: E402

IMG_WIDTH_CM = 16.0
DOCX_NAME = 'Отчёт_ЛР1_Установка_РОСА_на_VMware.docx'
PDF_NAME = 'Отчёт_ЛР1_Установка_РОСА_на_VMware.pdf'


def img_size(path, width_cm):
    """Возвращает (ширина, высота) картинки в сантиметрах с сохранением пропорций."""
    with Image.open(path) as im:
        w, h = im.size
    return width_cm, width_cm * h / w


def toc_entries():
    """Список (уровень, текст) для содержания в порядке следования."""
    out = []
    for block in content.SECTIONS:
        if block[0] == 'h1':
            out.append((0, block[1]))
        elif block[0] == 'h2':
            out.append((1, block[1]))
        elif block[0] == 'h1c':
            out.append((0, block[1].capitalize() if False else block[1]))
    return out


# ----------------------------------------------------------------------------- DOCX


def build_docx(page_map, path):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    TNR = 'Times New Roman'
    doc = Document()

    # --- страница / поля
    sec = doc.sections[0]
    sec.left_margin, sec.right_margin = Cm(3), Cm(1.5)
    sec.top_margin, sec.bottom_margin = Cm(2), Cm(2)
    sec.different_first_page_header_footer = True

    # --- базовый стиль
    normal = doc.styles['Normal']
    normal.font.name = TNR
    normal.font.size = Pt(14)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), TNR)
    normal.element.rPr.rFonts.set(qn('w:cs'), TNR)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(1.25)

    # --- заголовки
    for name in ('Heading 1', 'Heading 2'):
        st = doc.styles[name]
        st.font.name = TNR
        st.font.size = Pt(14)
        st.font.bold = True
        st.font.italic = False
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), TNR)
        st.element.rPr.rFonts.set(qn('w:cs'), TNR)
        p = st.paragraph_format
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.first_line_indent = Cm(1.25)
        p.space_before = Pt(12)
        p.space_after = Pt(6)
        p.line_spacing = 1.5
        p.keep_with_next = True

    def make_style(name, **kw):
        st = doc.styles.add_style(name, 1)  # WD_STYLE_TYPE.PARAGRAPH
        st.base_style = doc.styles['Normal']
        f = st.font
        f.name = TNR
        f.size = Pt(kw.get('size', 14))
        f.bold = kw.get('bold', False)
        f.italic = kw.get('italic', False)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), TNR)
        st.element.rPr.rFonts.set(qn('w:cs'), TNR)
        p = st.paragraph_format
        p.alignment = kw.get('align', WD_ALIGN_PARAGRAPH.JUSTIFY)
        p.first_line_indent = Cm(kw.get('indent', 1.25))
        p.left_indent = Cm(kw.get('left', 0))
        p.line_spacing = 1.5
        p.space_before = Pt(kw.get('before', 0))
        p.space_after = Pt(kw.get('after', 0))
        p.keep_with_next = kw.get('keep', False)
        return st

    cap_style = make_style('FigureCaption', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
                           before=6, after=12, keep=True)
    tab_cap_style = make_style('TableCaption', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
                               before=6, after=6, keep=True)
    list_style = make_style('ListGOST', align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=-0.5, left=1.75)
    center_style = make_style('CenterNoIndent', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0)
    left_style = make_style('LeftNoIndent', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0)

    # --- нумерация страниц (кроме титульного листа)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.first_line_indent = Cm(0)
    run = footer.add_run()
    run.font.name = TNR
    run.font.size = Pt(14)
    fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld); run._r.append(instr); run._r.append(fld2)

    # --- титульный лист
    for line in content.TITLE:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(line.get('gap', 0) or 6)
        r = p.add_run(line['t'])
        r.font.name = TNR
        r.font.size = Pt(line.get('sz', 14))
        r.bold = bool(line.get('b'))
        r.italic = bool(line.get('i'))

    doc.add_page_break()

    # --- содержание
    p = doc.add_paragraph(style='CenterNoIndent')
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run('СОДЕРЖАНИЕ')
    r.bold = True
    r.font.size = Pt(14)

    for level, text in toc_entries():
        par = doc.add_paragraph(style='LeftNoIndent')
        par.paragraph_format.space_after = Pt(0)
        par.paragraph_format.left_indent = Cm(0.75 * level)
        par.paragraph_format.tab_stops.add_tab_stop(Cm(16.5 - 0.75 * level), 2, 1)  # RIGHT, DOTS
        par.add_run('%s\t%s' % (text, page_map.get(text, '')))

    doc.add_page_break()

    fig_no = [0]

    def add_figure(fname, caption):
        fig_no[0] += 1
        path = os.path.join(SHOTS, fname)
        w_cm, h_cm = img_size(path, IMG_WIDTH_CM)
        par = doc.add_paragraph(style='CenterNoIndent')
        par.paragraph_format.space_before = Pt(6)
        par.add_run().add_picture(path, width=Cm(w_cm), height=Cm(h_cm))
        cap = doc.add_paragraph(style='FigureCaption')
        cap.add_run('Рисунок %d — %s' % (fig_no[0], caption))

    def add_table(spec):
        cap = doc.add_paragraph(style='TableCaption')
        cap.add_run(spec['caption'])
        rows = [spec['head']] + spec['rows']
        table = doc.add_table(rows=len(rows), cols=len(spec['head']))
        table.style = 'Table Grid'
        table.autofit = True
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = table.cell(i, j)
                cp = cell.paragraphs[0]
                cp.paragraph_format.first_line_indent = Cm(0)
                cp.paragraph_format.line_spacing = 1.0
                cp.paragraph_format.space_after = Pt(2)
                cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = cp.add_run(val)
                run.font.name = TNR
                run.font.size = Pt(12)
                run.bold = (i == 0)
        table.columns[0].width = Cm(6.0)
        table.columns[1].width = Cm(10.5)
        doc.add_paragraph(style='LeftNoIndent').paragraph_format.space_after = Pt(6)

    for block in content.SECTIONS:
        kind = block[0]
        if kind == 'h1':
            doc.add_paragraph(block[1], style='Heading 1')
        elif kind == 'h2':
            doc.add_paragraph(block[1], style='Heading 2')
        elif kind == 'h1c':
            p = doc.add_paragraph(style='CenterNoIndent')
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.keep_with_next = True
            p.add_run(block[1]).bold = True
        elif kind == 'p':
            doc.add_paragraph(block[1], style='Normal')
        elif kind == 'list':
            for item in block[1]:
                doc.add_paragraph('– ' + item, style='ListGOST')
        elif kind == 'fig':
            add_figure(block[1], block[2])
        elif kind == 'table':
            add_table(block[1])

    doc.core_properties.title = 'Отчёт по практической работе № 1. Установка ОС РОСА на VMware'
    doc.core_properties.author = '____________'
    doc.save(path)
    return fig_no[0]


# ------------------------------------------------------------------------------ PDF


def build_pdf(path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont as RLTTFont
    from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
                                    Image as RLImage, Table, TableStyle, PageBreak,
                                    NextPageTemplate)
    from reportlab.platypus.tableofcontents import TableOfContents

    for name, fname in (('Tinos', 'Tinos-Regular.ttf'), ('Tinos-Bold', 'Tinos-Bold.ttf'),
                        ('Tinos-Italic', 'Tinos-Italic.ttf'), ('Tinos-BoldItalic', 'Tinos-BoldItalic.ttf')):
        pdfmetrics.registerFont(RLTTFont(name, os.path.join(FONTS, fname)))
    pdfmetrics.registerFontFamily('Tinos', normal='Tinos', bold='Tinos-Bold',
                                  italic='Tinos-Italic', boldItalic='Tinos-BoldItalic')

    body = ParagraphStyle('body', fontName='Tinos', fontSize=14, leading=21, alignment=TA_JUSTIFY,
                          firstLineIndent=1.25 * cm)
    h1 = ParagraphStyle('h1', parent=body, fontName='Tinos-Bold', alignment=TA_LEFT,
                        spaceBefore=12, spaceAfter=6, keepWithNext=1)
    h2 = ParagraphStyle('h2', parent=h1, spaceBefore=12, spaceAfter=6, keepWithNext=1)
    h1c = ParagraphStyle('h1c', parent=body, fontName='Tinos-Bold', alignment=TA_CENTER,
                         firstLineIndent=0, spaceBefore=12, spaceAfter=6, keepWithNext=1)
    cap = ParagraphStyle('cap', parent=body, alignment=TA_CENTER, firstLineIndent=0,
                         spaceBefore=6, spaceAfter=12)
    tabcap = ParagraphStyle('tabcap', parent=body, alignment=TA_CENTER, firstLineIndent=0,
                            spaceBefore=6, spaceAfter=6, keepWithNext=1)
    li = ParagraphStyle('li', parent=body, firstLineIndent=0, leftIndent=1.75 * cm, bulletIndent=1.25 * cm)
    cover = ParagraphStyle('cover', parent=body, alignment=TA_CENTER, firstLineIndent=0, leading=18)

    class ReportDoc(BaseDocTemplate):
        def __init__(self, filename, **kw):
            BaseDocTemplate.__init__(self, filename, **kw)
            self.toc_pages = {}

        def afterFlowable(self, flowable):
            if isinstance(flowable, Paragraph):
                name = flowable.style.name
                if name in ('h1', 'h2', 'h1c'):
                    text = flowable.getPlainText()
                    level = 1 if name == 'h2' else 0
                    self.notify('TOCEntry', (level, text, self.page))
                    self.toc_pages[text] = self.page

    left, right, top, bottom = 3 * cm, 1.5 * cm, 2 * cm, 2 * cm
    frame = Frame(left, bottom, A4[0] - left - right, A4[1] - top - bottom, id='main',
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

    def draw_page_number(canvas, doc_):
        canvas.saveState()
        canvas.setFont('Tinos', 14)
        canvas.drawCentredString(A4[0] / 2.0, 1.2 * cm, str(doc_.page))
        canvas.restoreState()

    doc = ReportDoc(path, pagesize=A4, leftMargin=left, rightMargin=right, topMargin=top,
                    bottomMargin=bottom, title='Отчёт по практической работе № 1. Установка ОС РОСА на VMware',
                    author='____________')
    doc.addPageTemplates([
        PageTemplate(id='cover', frames=[frame], onPage=lambda c, d: None),
        PageTemplate(id='main', frames=[frame], onPage=draw_page_number),
    ])

    story = []
    story.append(NextPageTemplate('main'))
    for line in content.TITLE:
        story.append(Paragraph(line['t'] or '&nbsp;',
                               ParagraphStyle('c', parent=cover, fontName='Tinos-Bold' if line.get('b') else 'Tinos',
                                              fontSize=line.get('sz', 14), spaceAfter=(line.get('gap') or 6))))
    story.append(PageBreak())

    story.append(Paragraph('СОДЕРЖАНИЕ', ParagraphStyle('toc', parent=h1c, spaceAfter=12)))
    toc = TableOfContents()
    toc.dotsMinLevel = 0
    toc.levelStyles = [
        ParagraphStyle('TOC1', fontName='Tinos', fontSize=12, leading=18),
        ParagraphStyle('TOC2', fontName='Tinos', fontSize=12, leading=18, leftIndent=0.75 * cm),
    ]
    story.append(toc)
    story.append(PageBreak())

    fig_no = [0]

    def add_figure(fname, caption):
        fig_no[0] += 1
        p = os.path.join(SHOTS, fname)
        w_cm, h_cm = img_size(p, IMG_WIDTH_CM)
        story.append(Spacer(1, 6))
        story.append(RLImage(p, width=w_cm * cm, height=h_cm * cm))
        story.append(Paragraph('Рисунок %d — %s' % (fig_no[0], caption), cap))

    for block in content.SECTIONS:
        kind = block[0]
        if kind == 'h1':
            story.append(Paragraph(block[1], h1))
        elif kind == 'h2':
            story.append(Paragraph(block[1], h2))
        elif kind == 'h1c':
            story.append(Paragraph(block[1], h1c))
        elif kind == 'p':
            story.append(Paragraph(block[1], body))
        elif kind == 'list':
            for item in block[1]:
                story.append(Paragraph(item, li, bulletText='–'))
        elif kind == 'fig':
            add_figure(block[1], block[2])
        elif kind == 'table':
            spec = block[1]
            story.append(Paragraph(spec['caption'], tabcap))
            data = [spec['head']] + spec['rows']
            cell = ParagraphStyle('cell', fontName='Tinos', fontSize=12, leading=15)
            cellb = ParagraphStyle('cellb', parent=cell, fontName='Tinos-Bold')
            data = [[Paragraph(c, cellb if i == 0 else cell) for c in row] for i, row in enumerate(data)]
            t = Table(data, colWidths=[6.0 * cm, 10.5 * cm], repeatRows=1)
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, (0, 0, 0)),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            story.append(t)
            story.append(Spacer(1, 12))

    doc.multiBuild(story)
    return doc.toc_pages, fig_no[0]


def main():
    os.makedirs(REPORT_DIR, exist_ok=True)
    pdf_path = os.path.join(REPORT_DIR, PDF_NAME)
    docx_path = os.path.join(REPORT_DIR, DOCX_NAME)

    pages, n_figs = build_pdf(pdf_path)
    print('PDF: %s (%d рисунков)' % (pdf_path, n_figs))
    n_figs_docx = build_docx(pages, docx_path)
    print('DOCX: %s (%d рисунков)' % (docx_path, n_figs_docx))


if __name__ == '__main__':
    main()

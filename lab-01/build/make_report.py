#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка отчёта по практической работе № 1 (РОСА Фреш на VMware) в форматах DOCX, PDF и ODT.

Оформление — 1:1 по образцам lab-01/examples/*.odt и «Краткой выписке из
ГОСТ 7.32-2017»: Times New Roman 14 пт, полуторный межстрочный интервал,
выравнивание по ширине, абзацный отступ 1,25 см, поля 30/15/20/20 мм,
нумерация страниц внизу по центру тем же шрифтом (титульный лист входит
в нумерацию, но номер не проставляется). Титульный лист — по образцу
«Лаба 2»/«Gentoo»: «Колледж Научно-Технологического Университета Сириус»,
линия, «ПРАКТИЧЕСКАЯ РАБОТА № 1.», блок «Выполнил/Принял» справа, внизу
«IT-Колледж «Сириус», год. Заголовки — полужирные, после каждого заголовка
и подзаголовка — пустая строка; структурные элементы (РЕФЕРАТ, СОДЕРЖАНИЕ,
ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ) — прописными по
центру с нового листа; вокруг рисунков и подписей — пустые строки.

Скриншоты берутся из ../screenshots, готовые файлы складываются в ../report.
Требуется: python-docx, reportlab, Pillow, pypdfium2 (см. requirements.txt).
"""

import os
import re
import sys

from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(BASE)
SHOTS = os.path.join(LAB, 'screenshots')
REPORT_DIR = os.path.join(LAB, 'report')
FONTS = os.path.join(BASE, 'fonts')

sys.path.insert(0, BASE)
import content  # noqa: E402

# ----------------------------------------------------------------- параметры страницы
PAGE = {
    'margin_left': 3.0,     # см
    'margin_right': 1.5,    # см
    'margin_top': 2.0,      # см
    'margin_bottom': 2.0,   # см
    'indent': 1.25,         # см, абзацный отступ
    'font_size': 14,        # пт, основной текст
    'leading': 21.0,        # пт, полуторный интервал для 14 пт («пустая строка»)
    'img_width': float(os.environ.get('IMG_WIDTH_CM', '15')),  # см, ширина рисунков
}

DOCX_NAME = 'Отчёт_ЛР1_Установка_РОСА_на_VMware.docx'
PDF_NAME = 'Отчёт_ЛР1_Установка_РОСА_на_VMware.pdf'
ODT_NAME = 'Отчёт_ЛР1_Установка_РОСА_на_VMware.odt'


def img_size(path, width_cm):
    """Возвращает (ширина, высота) картинки в сантиметрах с сохранением пропорций."""
    with Image.open(path) as im:
        w, h = im.size
    return width_cm, width_cm * h / w


def in_toc(block):
    """Нужно ли включать структурный элемент (h1c) в содержание."""
    return len(block) < 3 or bool(block[2])


def skip_in_toc():
    """Тексты структурных элементов, которые не включаются в содержание."""
    return {b[1] for b in content.SECTIONS if b[0] == 'h1c' and not in_toc(b)}


def toc_entries(page_map):
    """Список (уровень, текст, номер страницы) для содержания."""
    out = []
    for block in content.SECTIONS:
        if block[0] == 'h1':
            out.append((0, block[1], page_map.get(block[1], '')))
        elif block[0] == 'h2':
            out.append((1, block[1], page_map.get(block[1], '')))
        elif block[0] == 'h1c' and in_toc(block):
            out.append((0, block[1], page_map.get(block[1], '')))
    return out


def counts_for(content_counts):
    """Числа для реферата: страницы, рисунки, таблицы, источники."""
    figs = sum(1 for b in content.SECTIONS if b[0] == 'fig')
    tables = sum(1 for b in content.SECTIONS if b[0] == 'table')
    return {
        'pages': content_counts.get('pages', ''),
        'figs': figs,
        'tables': tables,
        'sources': len(content.REFS),
    }


def format_text(text, counts):
    """Подставляет числа в текст (для реферата) и убирает экранирование."""
    try:
        return text.format(**counts)
    except (KeyError, IndexError):
        return text


# --------------------------------------------------------- перекрёстные ссылки на рисунки
# Ссылки в тексте вида «рисунок N» становятся кликабельными: щелчок (в Word —
# Ctrl+щелчок) переносит к подписи соответствующего рисунка. Работает во всех
# трёх форматах: DOCX (гиперссылка на закладку), PDF (внутренняя ссылка),
# ODT (text:a на text:bookmark).

FIG_REF_RE = re.compile(r'([Рр]исунок)\s+(\d+)')


def iter_fig_refs(text):
    """Разбивает текст на части: ('t', строка) и ('ref', фрагмент, номер_рисунка)."""
    pos = 0
    for m in FIG_REF_RE.finditer(text):
        if m.start() > pos:
            yield ('t', text[pos:m.start()])
        yield ('ref', m.group(0), int(m.group(2)))
        pos = m.end()
    if pos < len(text):
        yield ('t', text[pos:])


# ----------------------------------------------------------------------------- DOCX

def build_docx(page_map, path, counts):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    TNR = 'Times New Roman'
    SIZE = PAGE['font_size']
    IND = PAGE['indent']
    doc = Document()

    # --- страница, поля
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(PAGE['margin_left']), Cm(PAGE['margin_right'])
    sec.top_margin, sec.bottom_margin = Cm(PAGE['margin_top']), Cm(PAGE['margin_bottom'])
    sec.different_first_page_header_footer = True   # титульный лист без номера

    # --- базовый стиль (как Standard (WW) в образцах: TNR 14, 150%, по ширине, отступ 1,25)
    normal = doc.styles['Normal']
    normal.font.name = TNR
    normal.font.size = Pt(SIZE)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), TNR)
    normal.element.rPr.rFonts.set(qn('w:cs'), TNR)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(IND)

    def set_spacing(style, pt):
        """Разрежение текста заголовка (межбуквенный интервал), pt в пунктах."""
        rpr = style.element.get_or_add_rPr()
        el = rpr.find(qn('w:spacing'))
        if el is None:
            el = OxmlElement('w:spacing')
            rpr.append(el)
        el.set(qn('w:val'), str(int(pt * 20)))

    bookmark_id = [0]

    def add_bookmark(paragraph, name):
        """Закладка на абзаце (цель перекрёстной ссылки)."""
        bookmark_id[0] += 1
        start = OxmlElement('w:bookmarkStart')
        start.set(qn('w:id'), str(bookmark_id[0]))
        start.set(qn('w:name'), name)
        end = OxmlElement('w:bookmarkEnd')
        end.set(qn('w:id'), str(bookmark_id[0]))
        paragraph._p.insert(0, start)
        paragraph._p.append(end)

    def add_internal_link(paragraph, anchor, text):
        """Внутренняя гиперссылка на закладку (вид — как у обычного текста)."""
        hl = OxmlElement('w:hyperlink')
        hl.set(qn('w:anchor'), anchor)
        hl.set(qn('w:history'), '1')
        r = OxmlElement('w:r')
        rpr = OxmlElement('w:rPr')
        rfonts = OxmlElement('w:rFonts')
        for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
            rfonts.set(qn(a), TNR)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), str(PAGE['font_size'] * 2))
        rpr.append(rfonts)
        rpr.append(sz)
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text
        r.append(rpr)
        r.append(t)
        hl.append(r)
        paragraph._p.append(hl)

    # --- заголовки разделов (уровень 1) и подразделов (уровень 2, разреженный на 3 пт)
    for name, spacing in (('Heading 1', 0), ('Heading 2', 3)):
        st = doc.styles[name]
        st.font.name = TNR
        st.font.size = Pt(SIZE)
        st.font.bold = True
        st.font.italic = False
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), TNR)
        st.element.rPr.rFonts.set(qn('w:cs'), TNR)
        p = st.paragraph_format
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.first_line_indent = Cm(IND)
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.5
        p.keep_with_next = True
        if spacing:
            set_spacing(st, spacing)

    def make_style(name, **kw):
        st = doc.styles.add_style(name, 1)  # WD_STYLE_TYPE.PARAGRAPH
        st.base_style = doc.styles['Normal']
        f = st.font
        f.name = TNR
        f.size = Pt(kw.get('size', SIZE))
        f.bold = kw.get('bold', False)
        f.italic = kw.get('italic', False)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), TNR)
        st.element.rPr.rFonts.set(qn('w:cs'), TNR)
        p = st.paragraph_format
        p.alignment = kw.get('align', WD_ALIGN_PARAGRAPH.JUSTIFY)
        p.first_line_indent = Cm(kw.get('indent', IND))
        p.left_indent = Cm(kw.get('left', 0))
        p.line_spacing = 1.5
        p.space_before = Pt(kw.get('before', 0))
        p.space_after = Pt(kw.get('after', 0))
        p.keep_with_next = kw.get('keep', False)
        return st

    cap_style = make_style('FigureCaption', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0)
    tab_cap_style = make_style('TableCaption', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0,
                               keep=True)
    list_style = make_style('ListGOST')
    center_style = make_style('CenterNoIndent', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0)
    struct_style = make_style('StructHeading', align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
                              bold=True, keep=True)
    toc_style = make_style('TocGOST', align=WD_ALIGN_PARAGRAPH.LEFT, indent=0)

    # --- номер страницы внизу по центру тем же шрифтом (кроме титульного листа)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.first_line_indent = Cm(0)
    footer.paragraph_format.line_spacing = 1.0
    run = footer.add_run()
    run.font.name = TNR
    run.font.size = Pt(SIZE)
    fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld); run._r.append(instr); run._r.append(fld2)

    def add_empty(keep=False):
        """Пустая строка (как в образцах — после заголовков, вокруг рисунков)."""
        p = doc.add_paragraph(style='CenterNoIndent')
        if keep:
            p.paragraph_format.keep_with_next = True
        return p

    # --- титульный лист (1:1 по образцам «Лаба 2» / «Gentoo»)
    for line in content.TITLE:
        if line.get('gap') == 'BOTTOM':
            for _ in range(content.TITLE_BOTTOM_LINES):
                add_empty()
        p = doc.add_paragraph()
        p.alignment = (WD_ALIGN_PARAGRAPH.RIGHT if line.get('r')
                       else WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(line.get('gap') if isinstance(line.get('gap'), int) else 12)
        r = p.add_run(line['t'])
        r.font.name = TNR
        r.font.size = Pt(line.get('sz', SIZE))
        r.bold = bool(line.get('b'))

    fig_no = [0]

    def add_figure(fname, caption):
        fig_no[0] += 1
        path_img = os.path.join(SHOTS, fname)
        w_cm, h_cm = img_size(path_img, PAGE['img_width'])
        add_empty()                                   # пустая строка перед рисунком
        par = doc.add_paragraph(style='CenterNoIndent')
        par.paragraph_format.keep_with_next = True    # рисунок не отрывается от подписи
        par.add_run().add_picture(path_img, width=Cm(w_cm), height=Cm(h_cm))
        add_empty(keep=True)                          # пустая строка под рисунком
        cap = doc.add_paragraph(style='FigureCaption')
        cap.add_run('Рисунок %d – %s' % (fig_no[0], caption))
        add_bookmark(cap, 'fig%d' % fig_no[0])        # цель перекрёстных ссылок
        add_empty()                                   # пустая строка после подписи

    def add_table(spec):
        add_empty()
        cap = doc.add_paragraph(style='TableCaption')
        cap.add_run(spec['caption'])
        rows = [spec['head']] + spec['rows']
        table = doc.add_table(rows=len(rows), cols=len(spec['head']))
        table.style = 'Table Grid'
        table.autofit = False
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = table.cell(i, j)
                cp = cell.paragraphs[0]
                cp.paragraph_format.first_line_indent = Cm(0)
                cp.paragraph_format.line_spacing = 1.0
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after = Pt(2)
                cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = cp.add_run(val)
                run.font.name = TNR
                run.font.size = Pt(12)
                run.bold = (i == 0)
        table.columns[0].width = Cm(6.0)
        table.columns[1].width = Cm(10.5)
        for row in table.rows:      # ширины ячеек фиксируются по столбцам
            row.cells[0].width = Cm(6.0)
            row.cells[1].width = Cm(10.5)
        add_empty()

    for block in content.SECTIONS:
        kind = block[0]
        if kind == 'h1':
            doc.add_paragraph(block[1], style='Heading 1')
            add_empty(keep=True)
        elif kind == 'h2':
            doc.add_paragraph(block[1], style='Heading 2')
            add_empty(keep=True)
        elif kind == 'h1c':
            p = doc.add_paragraph(style='StructHeading')
            p.paragraph_format.page_break_before = True
            p.add_run(block[1])
            add_empty(keep=True)
        elif kind == 'p':
            par = doc.add_paragraph(style='Normal')
            for part in iter_fig_refs(format_text(block[1], counts)):
                if part[0] == 't':
                    par.add_run(part[1])
                else:
                    add_internal_link(par, 'fig%d' % part[2], part[1])
        elif kind == 'list':
            for item in block[1]:
                doc.add_paragraph('– ' + item, style='ListGOST')
        elif kind == 'refs':
            for i, item in enumerate(block[1], 1):
                doc.add_paragraph('%d %s' % (i, item), style='ListGOST')
        elif kind == 'toc':
            for level, text, page in toc_entries(page_map):
                par = doc.add_paragraph(style='TocGOST')
                par.paragraph_format.left_indent = Cm(0.5 * level)
                par.paragraph_format.first_line_indent = Cm(0)
                par.paragraph_format.tab_stops.add_tab_stop(
                    Cm(21.0 - PAGE['margin_left'] - PAGE['margin_right'] - 0.5 * level), 2, 1)
                par.add_run('%s\t%s' % (text, page))
        elif kind == 'fig':
            add_figure(block[1], block[2])
        elif kind == 'table':
            add_table(block[1])

    doc.core_properties.title = ('Отчёт о выполнении практической работы № 1. '
                                 'Установка ОС РОСА на VMware Workstation')
    doc.core_properties.author = '____________'
    doc.save(path)
    return fig_no[0]



# ------------------------------------------------------------------------------ ODT

def build_odt(page_map, path, counts):
    """Собирает .odt с тем же оформлением, что и DOCX/PDF (образцы — examples/*.odt):
    внутренние ссылки «рисунок N» ведут к закладкам на подписях рисунков."""
    from odf.opendocument import OpenDocumentText
    from odf.style import (Style, PageLayout, PageLayoutProperties, MasterPage, Footer,
                           ParagraphProperties, TextProperties, TabStops, TabStop,
                           TableColumnProperties, TableCellProperties, GraphicProperties,
                           FontFace)
    from odf.text import P, A, Tab, PageNumber, Bookmark
    from odf.draw import Frame, Image as DrawImage
    from odf.table import Table, TableColumn, TableRow, TableCell

    TNR = 'Times New Roman'
    odt = OpenDocumentText()

    odt.fontfacedecls.addElement(FontFace(name=TNR, fontfamily="'%s'" % TNR,
                                          fontfamilygeneric='roman', fontpitch='variable'))

    # --- страница и колонтитулы: Standard — с номером, FirstPage — без (титульный)
    pagenum_cm = lambda v: ('%.3fcm' % v)
    pgl = PageLayout(name='Mpm1')
    pgl.addElement(PageLayoutProperties(pagewidth='21.001cm', pageheight='29.7cm',
                                        printorientation='portrait',
                                        margintop=pagenum_cm(PAGE['margin_top']),
                                        marginbottom=pagenum_cm(PAGE['margin_bottom']),
                                        marginleft=pagenum_cm(PAGE['margin_left']),
                                        marginright=pagenum_cm(PAGE['margin_right']),
                                        writingmode='lr-tb'))
    odt.automaticstyles.addElement(pgl)

    def make_master(name, with_footer):
        mp = MasterPage(name=name, pagelayoutname=pgl)
        if with_footer:
            ftr = Footer()
            fp = P(stylename='FooterP')
            fp.addElement(PageNumber(selectpage='current'))
            ftr.addElement(fp)
            mp.addElement(ftr)
        odt.masterstyles.addElement(mp)

    make_master('Standard', with_footer=True)
    make_master('FirstPage', with_footer=False)

    # --- стили абзацев (по образцам: TNR 14, 150 %, по ширине, отступ 1,25 см)
    def pstyle(name, parent=None, master=None, pp=None, tp=None):
        st = Style(name=name, family='paragraph')
        if parent:
            st.setAttribute('parentstylename', parent)
        if master:
            st.setAttribute('masterpagename', master)
        if pp:
            pp = dict(pp)
            tabs = pp.pop('tabstops', None)
            pr = ParagraphProperties(**pp)
            if tabs is not None:
                pr.addElement(tabs)
            st.addElement(pr)
        if tp:
            st.addElement(TextProperties(**tp))
        odt.styles.addElement(st)
        return st

    base_pp = dict(margintop='0cm', marginbottom='0cm', lineheight='150%',
                   textalign='justify', textindent='1.251cm')
    base_tp = dict(fontname=TNR, fontfamily=TNR, fontsize='14pt', fontweight='normal',
                   color='#000000')
    pstyle('ReportBase', pp=base_pp, tp=base_tp)
    pstyle('ReportH1', parent='ReportBase', pp=dict(keepwithnext='always'),
           tp=dict(fontweight='bold'))
    pstyle('ReportH2', parent='ReportBase', pp=dict(keepwithnext='always'),
           tp=dict(fontweight='bold', letterspacing='0.106cm'))
    pstyle('ReportH1c', parent='ReportBase', master='Standard',
           pp=dict(textalign='center', textindent='0cm', breakbefore='page',
                   keepwithnext='always'),
           tp=dict(fontweight='bold'))
    pstyle('ReportCaption', parent='ReportBase',
           pp=dict(textalign='center', textindent='0cm'))
    pstyle('ReportTabCap', parent='ReportBase',
           pp=dict(textalign='left', textindent='0cm', keepwithnext='always'))
    pstyle('ReportImg', parent='ReportBase',
           pp=dict(textalign='center', textindent='0cm', keepwithnext='always'))
    pstyle('ReportEmpty', parent='ReportBase')
    pstyle('FooterP', parent='ReportBase',
           pp=dict(textalign='center', textindent='0cm', lineheight='normal'))

    def toc_style(name, left):
        tabs = TabStops()
        # отточие и номер страницы по правому краю — как в образцах (16,484 см)
        tabs.addElement(TabStop(position=pagenum_cm(21.0 - PAGE['margin_left']
                                                - PAGE['margin_right'] - 0.01),
                                type='right', leaderstyle='dotted', leadertext='.'))
        pstyle(name, parent='ReportBase',
               pp=dict(textalign='left', textindent='0cm', marginleft=pagenum_cm(left),
                       numberlines='false', tabstops=tabs))
    toc_style('ReportToc1', 0.0)
    toc_style('ReportToc2', 0.5)

    # титульный лист: центры/право, жирные варианты, первая строка — мастер без номера
    pstyle('TitleC', parent='ReportBase',
           pp=dict(textalign='center', textindent='0cm', marginbottom='0.42cm'))
    pstyle('TitleCB', parent='ReportBase',
           pp=dict(textalign='center', textindent='0cm', marginbottom='0.42cm'),
           tp=dict(fontweight='bold'))
    pstyle('TitleR', parent='ReportBase',
           pp=dict(textalign='end', textindent='0cm', marginbottom='0.42cm'))
    pstyle('TitleCFirst', parent='ReportBase', master='FirstPage',
           pp=dict(textalign='center', textindent='0cm', marginbottom='0.42cm'),
           tp=dict(fontweight='bold'))

    frame_style = Style(name='ReportFrame', family='graphic')
    frame_style.addElement(GraphicProperties(wrap='run-through'))
    odt.styles.addElement(frame_style)

    cellb = '0.5pt solid #000000'
    col1 = Style(name='TableCol1', family='table-column')
    col1.addElement(TableColumnProperties(columnwidth='6.0cm'))
    col2 = Style(name='TableCol2', family='table-column')
    col2.addElement(TableColumnProperties(columnwidth='10.49cm'))
    tcell = Style(name='TCell', family='table-cell')
    tcell.addElement(TableCellProperties(border=cellb, paddingtop='0.06cm',
                                         paddingbottom='0.06cm', paddingleft='0.12cm',
                                         paddingright='0.12cm'))
    for st in (col1, col2, tcell):
        odt.styles.addElement(st)
    pstyle('CellP', parent='ReportBase',
           pp=dict(textalign='left', textindent='0cm', lineheight='115%'),
           tp=dict(fontsize='12pt'))
    pstyle('CellPH', parent='ReportBase',
           pp=dict(textalign='left', textindent='0cm', lineheight='115%'),
           tp=dict(fontsize='12pt', fontweight='bold'))

    def add_p(stylename, content=None):
        p = P(stylename=stylename)
        if content:
            p.addText(content)
        odt.text.addElement(p)
        return p

    def add_empty():
        add_p('ReportEmpty')

    # --- титульный лист (1:1 по образцам «Лаба 2» / «Gentoo»)
    first = [True]
    for line in content.TITLE:
        if line.get('gap') == 'BOTTOM':
            add_empty()
        if line.get('r'):
            st = 'TitleR'
        elif first[0]:
            st = 'TitleCFirst'
            first[0] = False
        else:
            st = 'TitleCB' if line.get('b') else 'TitleC'
        add_p(st, line['t'])

    fig_no = [0]
    tbl_no = [0]

    def add_figure(fname, caption):
        fig_no[0] += 1
        p_img = os.path.join(SHOTS, fname)
        w_cm, h_cm = img_size(p_img, PAGE['img_width'])
        pic = odt.addPicture(p_img)
        add_empty()
        ip = P(stylename='ReportImg')
        frame = Frame(stylename=frame_style, width=pagenum_cm(w_cm),
                      height=pagenum_cm(h_cm), anchortype='as-char')
        frame.addElement(DrawImage(href=pic))
        ip.addElement(frame)
        odt.text.addElement(ip)
        add_empty()
        cp = P(stylename='ReportCaption')
        cp.addElement(Bookmark(name='fig%d' % fig_no[0]))   # цель перекрёстных ссылок
        cp.addText('Рисунок %d – %s' % (fig_no[0], caption))
        odt.text.addElement(cp)
        add_empty()

    def add_table(spec):
        tbl_no[0] += 1
        add_empty()
        add_p('ReportTabCap', spec['caption'])
        t = Table(name='Table%d' % tbl_no[0])
        t.addElement(TableColumn(stylename='TableCol1'))
        t.addElement(TableColumn(stylename='TableCol2'))
        for i, row in enumerate([spec['head']] + spec['rows']):
            tr = TableRow()
            for val in row:
                c = TableCell(stylename='TCell', valuetype='string')
                cp = P(stylename='CellPH' if i == 0 else 'CellP')
                cp.addText(val)
                c.addElement(cp)
                tr.addElement(c)
            t.addElement(tr)
        odt.text.addElement(t)
        add_empty()

    for block in content.SECTIONS:
        kind = block[0]
        if kind == 'h1':
            add_p('ReportH1', block[1])
            add_empty()
        elif kind == 'h2':
            add_p('ReportH2', block[1])
            add_empty()
        elif kind == 'h1c':
            add_p('ReportH1c', block[1])
            add_empty()
        elif kind == 'p':
            p = P(stylename='ReportBase')
            for part in iter_fig_refs(format_text(block[1], counts)):
                if part[0] == 't':
                    p.addText(part[1])
                else:
                    a = A(href='#fig%d' % part[2], type='simple')
                    a.addText(part[1])
                    p.addElement(a)
            odt.text.addElement(p)
        elif kind == 'list':
            for item in block[1]:
                add_p('ReportBase', '– ' + item)
        elif kind == 'refs':
            for i, item in enumerate(block[1], 1):
                add_p('ReportBase', '%d %s' % (i, item))
        elif kind == 'toc':
            for level, text, page in toc_entries(page_map):
                p = P(stylename='ReportToc2' if level else 'ReportToc1')
                p.addText(text)
                p.addElement(Tab())
                p.addText(str(page))
                odt.text.addElement(p)
        elif kind == 'fig':
            add_figure(block[1], block[2])
        elif kind == 'table':
            add_table(block[1])

    odt.meta.addElement(__import__('odf.dc', fromlist=['Title']).Title(
        text='Отчёт о выполнении практической работы № 1. '
             'Установка ОС РОСА на VMware Workstation'))
    odt.save(path)
    return fig_no[0]



# ------------------------------------------------------------------------------ PDF

def build_pdf(path, counts):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont as RLTTFont
    from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
                                    Image as RLImage, Table, TableStyle, PageBreak,
                                    NextPageTemplate, KeepTogether)
    from reportlab.platypus.tableofcontents import TableOfContents

    for name, fname in (('Tinos', 'Tinos-Regular.ttf'), ('Tinos-Bold', 'Tinos-Bold.ttf'),
                        ('Tinos-Italic', 'Tinos-Italic.ttf'), ('Tinos-BoldItalic', 'Tinos-BoldItalic.ttf')):
        pdfmetrics.registerFont(RLTTFont(name, os.path.join(FONTS, fname)))
    pdfmetrics.registerFontFamily('Tinos', normal='Tinos', bold='Tinos-Bold',
                                  italic='Tinos-Italic', boldItalic='Tinos-BoldItalic')

    SIZE = PAGE['font_size']
    LEAD = PAGE['leading']
    TOC_NUM_RESERVE = 1.2 * cm   # место под отточие и номер страницы в содержании
    IND = PAGE['indent'] * cm
    EMPTY = LEAD                 # «пустая строка», как в образцах

    body = ParagraphStyle('body', fontName='Tinos', fontSize=SIZE, leading=LEAD,
                          alignment=TA_JUSTIFY, firstLineIndent=IND)
    h1 = ParagraphStyle('h1', parent=body, fontName='Tinos-Bold', spaceBefore=0,
                        spaceAfter=0, keepWithNext=1)
    h2 = ParagraphStyle('h2', parent=h1, charSpace=3, keepWithNext=1)
    h1c = ParagraphStyle('h1c', parent=body, fontName='Tinos-Bold', alignment=TA_CENTER,
                         firstLineIndent=0, spaceBefore=0, spaceAfter=0, keepWithNext=1)
    cap = ParagraphStyle('cap', parent=body, alignment=TA_CENTER, firstLineIndent=0,
                         spaceBefore=0, spaceAfter=0)
    tabcap = ParagraphStyle('tabcap', parent=body, alignment=TA_LEFT, firstLineIndent=0,
                            spaceBefore=0, spaceAfter=0, keepWithNext=1)
    li = ParagraphStyle('li', parent=body, firstLineIndent=IND)
    cover = ParagraphStyle('cover', parent=body, alignment=TA_CENTER, firstLineIndent=0,
                           spaceAfter=12)
    cover_r = ParagraphStyle('cover_r', parent=cover, alignment=TA_RIGHT)

    skip = skip_in_toc()

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
                    self.toc_pages[text] = self.page
                    if text not in skip:
                        self.notify('TOCEntry', (level, text, self.page))

    left = PAGE['margin_left'] * cm
    right = PAGE['margin_right'] * cm
    top = PAGE['margin_top'] * cm
    bottom = PAGE['margin_bottom'] * cm
    frame = Frame(left, bottom, A4[0] - left - right, A4[1] - top - bottom, id='main',
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

    def draw_page_number(canvas, doc_):
        canvas.saveState()
        canvas.setFont('Tinos', SIZE)
        canvas.drawCentredString(A4[0] / 2.0, 1.2 * cm, str(doc_.page))
        canvas.restoreState()

    doc = ReportDoc(path, pagesize=A4, leftMargin=left, rightMargin=right, topMargin=top,
                    bottomMargin=bottom,
                    title='Отчёт о выполнении практической работы № 1. Установка ОС РОСА на VMware',
                    author='____________')
    doc.addPageTemplates([
        PageTemplate(id='cover', frames=[frame], onPage=lambda c, d: None),
        PageTemplate(id='main', frames=[frame], onPage=draw_page_number),
    ])

    # --- титульный лист (1:1 по образцам «Лаба 2» / «Gentoo»)
    story = [NextPageTemplate('main')]
    for line in content.TITLE:
        if line.get('gap') == 'BOTTOM':
            story.append(Spacer(1, content.TITLE_BOTTOM_GAP))
        st = cover_r if line.get('r') else cover
        story.append(Paragraph(line['t'] or '&nbsp;',
                               ParagraphStyle('c', parent=st,
                                              fontName='Tinos-Bold' if line.get('b') else 'Tinos',
                                              fontSize=line.get('sz', SIZE))))

    fig_no = [0]

    def add_figure(fname, caption):
        fig_no[0] += 1
        p = os.path.join(SHOTS, fname)
        w_cm, h_cm = img_size(p, PAGE['img_width'])
        story.append(Spacer(1, EMPTY))
        story.append(KeepTogether([
            RLImage(p, width=w_cm * cm, height=h_cm * cm),
            Spacer(1, EMPTY),
            # <a name> — якорь-цель перекрёстных ссылок на этот рисунок
            Paragraph('<a name="fig%d"/>Рисунок %d – %s' % (fig_no[0], fig_no[0], caption), cap),
        ]))
        story.append(Spacer(1, EMPTY))

    for block in content.SECTIONS:
        kind = block[0]
        if kind == 'h1':
            story.append(Paragraph(block[1], h1))
            story.append(Spacer(1, EMPTY))
        elif kind == 'h2':
            story.append(Paragraph(block[1], h2))
            story.append(Spacer(1, EMPTY))
        elif kind == 'h1c':
            story.append(PageBreak())
            story.append(Paragraph(block[1], h1c))
            story.append(Spacer(1, EMPTY))
        elif kind == 'p':
            txt = FIG_REF_RE.sub(
                lambda m: '<a href="#fig%s">%s</a>' % (m.group(2), m.group(0)),
                format_text(block[1], counts))
            story.append(Paragraph(txt, body))
        elif kind == 'list':
            for item in block[1]:
                story.append(Paragraph('– ' + item, li))
        elif kind == 'refs':
            for i, item in enumerate(block[1], 1):
                story.append(Paragraph('%d %s' % (i, item), li))
        elif kind == 'toc':
            toc = TableOfContents()
            toc.dotsMinLevel = 0
            # rightIndent резервирует место под отточие и номер страницы: иначе у длинной
            # строки номера, занимающей всю полосу набора, номер уезжает за поле.
            toc.levelStyles = [
                ParagraphStyle('TOC1', fontName='Tinos', fontSize=SIZE, leading=LEAD,
                               firstLineIndent=0, rightIndent=TOC_NUM_RESERVE),
                ParagraphStyle('TOC2', fontName='Tinos', fontSize=SIZE, leading=LEAD,
                               leftIndent=0.5 * cm, firstLineIndent=0,
                               rightIndent=TOC_NUM_RESERVE),
            ]
            story.append(toc)
        elif kind == 'fig':
            add_figure(block[1], block[2])
        elif kind == 'table':
            spec = block[1]
            story.append(Spacer(1, EMPTY))
            story.append(Paragraph(spec['caption'], tabcap))
            data = [spec['head']] + spec['rows']
            cell = ParagraphStyle('cell', fontName='Tinos', fontSize=12, leading=15,
                                  firstLineIndent=0)
            cellb = ParagraphStyle('cellb', parent=cell, fontName='Tinos-Bold')
            data = [[Paragraph(c, cellb if i == 0 else cell) for c in row]
                    for i, row in enumerate(data)]
            t = Table(data, colWidths=[6.0 * cm, 10.5 * cm], repeatRows=1)
            t.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, (0, 0, 0)),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            story.append(KeepTogether([t]))
            story.append(Spacer(1, EMPTY))

    doc.multiBuild(story)
    return doc.toc_pages, doc.page


def main():
    os.makedirs(REPORT_DIR, exist_ok=True)
    pdf_path = os.path.join(REPORT_DIR, PDF_NAME)
    docx_path = os.path.join(REPORT_DIR, DOCX_NAME)

    counts = counts_for({'pages': 0})     # первая прикидка — для числа страниц
    pages_map, n_pages = {}, 0
    for _ in range(4):                    # число страниц уточняется итеративно
        pages_map, n_pages = build_pdf(pdf_path, counts)
        if counts['pages'] == n_pages:
            break
        counts = counts_for({'pages': n_pages})
    print('PDF: %s — %d с., %d рис., %d табл., %d ист.'
          % (pdf_path, n_pages, counts['figs'], counts['tables'], counts['sources']))

    n_figs_docx = build_docx(pages_map, docx_path, counts)
    print('DOCX: %s — %d рис.' % (docx_path, n_figs_docx))

    odt_path = os.path.join(REPORT_DIR, ODT_NAME)
    n_figs_odt = build_odt(pages_map, odt_path, counts)
    print('ODT: %s — %d рис.' % (odt_path, n_figs_odt))


if __name__ == '__main__':
    main()

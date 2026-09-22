#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка отчёта по лабораторной работе № 3 «Настройка загрузчика GRUB2» (ODT).

Оформление — по «Краткой выписке из ГОСТ 7.32-2017» (файл
lab-01/examples/ОформлениеОтчета_Краткая_выписка_из_ГОСТ_с_Примерами.pdf):

  * Times New Roman 14 пт, чёрный, выравнивание по ширине;
  * межстрочный интервал — полуторный, абзацный отступ 1,25 см,
    интервалы перед и после абзаца — 0;
  * поля: левое 3,0 см, правое 1,5 см, верхнее и нижнее 2,0 см;
  * номера страниц — внизу по центру, тем же шрифтом и размером; на титульном
    листе номер не проставляется;
  * заголовки разделов — полужирные, с абзацного отступа, после заголовка одна
    пустая строка; заголовки подразделов дополнительно разрежены на 3 пт;
  * названия структурных элементов (РЕФЕРАТ, СОДЕРЖАНИЕ, СПИСОК
    ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ) — с нового листа, по центру, прописными;
  * подпись рисунка — по центру под рисунком («Рисунок N – Название»),
    название таблицы — слева над таблицей («Таблица N – Название»);
  * выводы команд — моноширинным шрифтом Courier New 10 пт;
  * номера страниц в СОДЕРЖАНИИ — поля-перекрёстные ссылки на закладки: в
    LibreOffice они пересчитываются автоматически; в файл записываются
    значения, рассчитанные модулем paginate.py.

Скриншоты берутся из ../screenshots, выводы команд — из ../outputs,
готовый файл складывается в ../report.
"""

import os
import sys

from PIL import Image

from odf.namespaces import (FONS, STYLENS, TEXTNS, SVGNS, DRAWNS, XLINKNS,
                            TABLENS, OFFICENS)

BASE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(BASE)
SHOTS = os.path.join(LAB, 'screenshots')
REPORT_DIR = os.path.join(LAB, 'report')
FONTS = os.path.normpath(os.path.join(BASE, '..', '..', 'lab-01', 'build', 'fonts'))

sys.path.insert(0, BASE)
import content  # noqa: E402
import paginate  # noqa: E402

# ------------------------------------------------------------------ параметры
CM = 72.0 / 2.54

PAGE = {
    'width': 21.0, 'height': 29.7,
    'margin_left': 3.0, 'margin_right': 1.5,
    'margin_top': 2.0, 'margin_bottom': 2.0,
    'text_width': 21.0 - 3.0 - 1.5,      # 16,5 см
    'body_size': 14, 'body_indent': 1.25,
    'blank_line': 21.0,                  # «одна пустая строка» после заголовка, пт
    'code_size': 10,
    'table_size': 12,
    'img_width': 16.5,                   # см
    'img_max_height': 21.0,              # см
}

ODT_NAME = 'Отчёт_ЛР3_Настройка_загрузчика_GRUB2.odt'


# -------------------------------------------------------------------- утилиты
def img_size(path, width_cm, max_height_cm):
    """Размер рисунка в сантиметрах с сохранением пропорций."""
    with Image.open(path) as im:
        w, h = im.size
    width = float(width_cm)
    height = width * h / float(w)
    if height > max_height_cm:
        height = float(max_height_cm)
        width = height * w / float(h)
    return round(width, 3), round(height, 3)


def cm(value):
    return '%.3fcm' % value


def pt(value):
    return '%.2fpt' % value


def in_toc(block):
    return len(block) < 3 or bool(block[2])


# ----------------------------------------------------------- работа с odfpy
NSMAP = {
    'fo': FONS, 'style': STYLENS, 'text': TEXTNS, 'svg': SVGNS,
    'draw': DRAWNS, 'xlink': XLINKNS, 'table': TABLENS, 'office': OFFICENS,
}


def with_attrs(element, mapping):
    """Присваивает элементу атрибуты с учётом пространств имён ODF.

    odfpy не принимает имена вида «fo:font-size» напрямую, поэтому имена
    разбираются на префикс и локальную часть и устанавливаются через setAttrNS.
    """
    for key, value in mapping.items():
        if isinstance(key, tuple):
            namespace, local = key
        else:
            prefix, local = key.split(':', 1)
            namespace = NSMAP[prefix]
        element.setAttrNS(namespace, local, value)
    return element


def new(cls, mapping=None):
    element = cls()
    return with_attrs(element, mapping) if mapping else element


# --------------------------------------------------------------------- стили
def make_styles(doc):
    """Определения стилей абзацев, ячеек, разметки страницы и колонтитулов."""
    from odf.style import (Style, TextProperties, ParagraphProperties, PageLayout,
                           PageLayoutProperties, MasterPage, Footer, TabStop,
                           TabStops, TableCellProperties)
    from odf.text import P, PageNumber

    TNR = 'Times New Roman'
    MONO = 'Courier New'

    def text_props(size=14, bold=False, mono=False, spacing=None):
        props = {
            'fo:font-size': pt(size),
            'fo:font-family': "'%s'" % (MONO if mono else TNR),
            'style:font-name': MONO if mono else TNR,
            'style:font-family-generic': 'modern' if mono else 'roman',
            'style:font-pitch': 'fixed' if mono else 'variable',
            'fo:color': '#000000',
            'fo:font-weight': 'bold' if bold else 'normal',
            'fo:hyphenate': 'false',
        }
        if spacing:
            props['fo:letter-spacing'] = cm(spacing)
        return with_attrs(TextProperties(), props)

    def para_props(align='justify', indent=0.0, left=None, before=0.0, after=0.0,
                   line_height=150, keep_next=False, break_before=False,
                   master=None, tab=False):
        props = {
            'fo:text-align': align,
            'fo:text-indent': cm(indent),
            'fo:margin-top': pt(before),
            'fo:margin-bottom': pt(after),
            'fo:line-height': '%d%%' % line_height,
            'fo:orphans': '2',
            'fo:widows': '2',
            'style:justify-single-word': 'false',
            'style:contextual-spacing': 'false',
            'style:auto-text-indent': 'false',
        }
        if left is not None:
            props['fo:margin-left'] = cm(left)
        if keep_next:
            props['fo:keep-with-next'] = 'always'
        if break_before:
            props['fo:break-before'] = 'page'
        if master:
            props['style:master-page-name'] = master
        element = with_attrs(ParagraphProperties(), props)
        if tab:                     # отточие и номер страницы справа — для СОДЕРЖАНИЯ
            stops = TabStops()
            stops.addElement(with_attrs(TabStop(position=cm(PAGE['text_width'])), {
                'style:type': 'right',
                'style:leader-char': '.',
            }))
            element.addElement(stops)
        return element

    def add(name, text=None, para=None):
        style = Style(name=name, family='paragraph')
        if text is not None:
            style.addElement(text)
        if para is not None:
            style.addElement(para)
        doc.styles.addElement(style)
        return style

    indent = PAGE['body_indent']
    blank = PAGE['blank_line']

    # --- основной текст, заголовки, структурные элементы
    add('Body', text_props(PAGE['body_size']), para_props(indent=indent))
    add('BodyBold', text_props(PAGE['body_size'], bold=True), para_props(indent=indent))
    add('H1', text_props(PAGE['body_size'], bold=True),
        para_props(align='start', indent=indent, after=blank, keep_next=True))
    add('H1Break', text_props(PAGE['body_size'], bold=True),
        para_props(align='start', indent=indent, after=blank, keep_next=True,
                   break_before=True))
    add('H2', text_props(PAGE['body_size'], bold=True, spacing=0.106),
        para_props(align='start', indent=indent, after=blank, keep_next=True))
    add('StructH', text_props(PAGE['body_size'], bold=True),
        para_props(align='center', indent=0, after=blank, keep_next=True,
                   break_before=True, master='Standard'))
    # --- рисунки и таблицы
    add('FigPara', text_props(PAGE['body_size']),
        para_props(align='center', indent=0, before=6, line_height=100,
                   keep_next=True))
    add('FigCaption', text_props(PAGE['body_size']),
        para_props(align='center', indent=0, before=6, after=blank))
    add('TabCaption', text_props(PAGE['body_size']),
        para_props(align='start', indent=0, before=12, after=6, keep_next=True))
    # подпись таблицы, которую нельзя разрывать между страницами: переносим её
    # вместе с таблицей на следующий лист
    add('TabCaptionBreak', text_props(PAGE['body_size']),
        para_props(align='start', indent=0, before=12, after=6, keep_next=True,
                   break_before=True))
    # --- выводы команд
    add('Code', text_props(PAGE['code_size'], mono=True),
        para_props(align='start', indent=0, line_height=100))
    # --- содержание
    add('TOC1', text_props(PAGE['body_size']), para_props(align='start', indent=0, tab=True))
    add('TOC2', text_props(PAGE['body_size']),
        para_props(align='start', indent=0, left=0.5, tab=True))
    # --- титульный лист
    add('TitleFirst', text_props(PAGE['body_size'], bold=True),
        para_props(align='center', indent=0, after=6, line_height=100,
                   master='FirstPage'))
    add('TitleLine', text_props(PAGE['body_size']),
        para_props(align='center', indent=0, after=6, line_height=100))
    add('TitleBold', text_props(PAGE['body_size'], bold=True),
        para_props(align='center', indent=0, after=6, line_height=100))
    add('TitleSmall', text_props(12),
        para_props(align='center', indent=0, after=6, line_height=100))
    add('TitleBig', text_props(18, bold=True),
        para_props(align='center', indent=0, after=6, line_height=100))
    add('FootP', text_props(PAGE['body_size']),
        para_props(align='center', indent=0, line_height=100))

    # --- ячейки таблиц
    cell_text = text_props(PAGE['table_size'])
    for name, bold in (('TblCell', False), ('TblHead', True)):
        style = Style(name=name, family='table-cell')
        style.addElement(with_attrs(TableCellProperties(), {
            'fo:border': '0.5pt solid #000000',
            'fo:padding': '0.08cm',
            'fo:vertical-align': 'middle',
        }))
        style.addElement(text_props(PAGE['table_size'], bold=bold))
        doc.styles.addElement(style)
    for name, bold in (('TblP', False), ('TblPHead', True)):
        style = Style(name=name, family='paragraph')
        style.addElement(cell_text)
        style.addElement(with_attrs(ParagraphProperties(), {
            'fo:text-align': 'start',
            'fo:text-indent': '0cm',
            'fo:margin-top': '0cm',
            'fo:margin-bottom': '2pt',
            'fo:line-height': '100%',
            'fo:orphans': '2',
            'fo:widows': '2',
            'fo:font-weight': 'bold' if bold else 'normal',
        }))
        doc.styles.addElement(style)

    # --- параметры страницы
    layout = PageLayout(name='pm1')
    layout.addElement(with_attrs(PageLayoutProperties(), {
        'fo:page-width': cm(PAGE['width']),
        'fo:page-height': cm(PAGE['height']),
        'fo:margin-top': cm(PAGE['margin_top']),
        'fo:margin-bottom': cm(PAGE['margin_bottom']),
        'fo:margin-left': cm(PAGE['margin_left']),
        'fo:margin-right': cm(PAGE['margin_right']),
        'style:print-orientation': 'portrait',
        'style:writing-mode': 'lr-tb',
    }))
    doc.automaticstyles.addElement(layout)

    # --- колонтитулы: номер страницы внизу по центру (титульный лист — без номера)
    number = PageNumber(selectpage='current')
    number.addText('1')
    footer_p = P(stylename='FootP')
    footer_p.addElement(number)
    footer = Footer()
    footer.addElement(footer_p)

    standard = MasterPage(name='Standard', pagelayoutname='pm1')
    standard.addElement(footer)
    doc.masterstyles.addElement(standard)

    doc.masterstyles.addElement(MasterPage(name='FirstPage', pagelayoutname='pm1',
                                           nextstylename='Standard'))


# ------------------------------------------------------------------- документ
def build_document(page_map=None, page_count=None, path=None, table_breaks=None):
    """Собирает ODT. page_map — {(уровень, заголовок): страница}."""
    from odf.opendocument import OpenDocumentText
    from odf.text import (P, Tab, BookmarkStart, BookmarkEnd, BookmarkRef,
                          PageCount)
    from odf.dc import Title as MetaTitle, Creator
    from odf.draw import Frame, Image as DrawImage
    from odf.table import Table, TableColumn, TableRow, TableCell
    from odf.style import Style, TableColumnProperties

    page_map = page_map or {}
    table_breaks = set(table_breaks or ())
    doc = OpenDocumentText()
    make_styles(doc)

    body = doc.text
    fig_no = [0]
    tab_no = [0]

    # --- предварительный проход: список заголовков для СОДЕРЖАНИЯ
    headings = []
    all_names = []
    index = 0
    for block in content.SECTIONS:
        kind = block[0]
        if kind == 'h1c' or kind in ('h1', 'h2'):
            index += 1
            name = '_Toc%d' % (10000 + index)
            all_names.append(name)
            if kind == 'h1c':
                if in_toc(block):
                    headings.append((1, block[1], name))
            else:
                headings.append((2 if kind == 'h2' else 1, block[1], name))
    heading_iter = iter(all_names)

    def add_par(text, style='Body'):
        p = P(stylename=style)
        p.addText(text)
        body.addElement(p)
        return p

    # ------------------------------------------------------------ титульный лист
    for i, line in enumerate(content.TITLE):
        text = line['t']
        if i == 0:
            style = 'TitleFirst'
        elif line.get('sz') == 18:
            style = 'TitleBig'
        elif line.get('sz') == 12:
            style = 'TitleSmall'
        elif line.get('b'):
            style = 'TitleBold'
        else:
            style = 'TitleLine'
        add_par(text, style=style)
        for _ in range(int(round(line.get('gap', 0) / 18.0))):   # пустые строки-отбивки
            add_par('', style='TitleLine')

    # ------------------------------------------------------------------- текст
    for block in content.SECTIONS:
        kind = block[0]
        if kind == 'h1c':
            name = next(heading_iter)
            p = P(stylename='StructH')
            p.addElement(BookmarkStart(name=name))
            p.addText(block[1])
            p.addElement(BookmarkEnd(name=name))
            body.addElement(p)
        elif kind in ('h1', 'h2'):
            name = next(heading_iter)
            if kind == 'h1':
                style = 'H1Break' if len(block) > 2 and block[2] == 'break' else 'H1'
            else:
                style = 'H2'
            p = P(stylename=style)
            p.addElement(BookmarkStart(name=name))
            p.addText(block[1])
            p.addElement(BookmarkEnd(name=name))
            body.addElement(p)
        elif kind == 'p':
            add_par(block[1])
        elif kind == 'pb':
            add_par(block[1], style='BodyBold')
        elif kind == 'list':
            for item in block[1]:
                add_par('– ' + item)
        elif kind == 'refs':
            for i, item in enumerate(block[1], 1):
                add_par('%d %s' % (i, item))
        elif kind == 'code':
            for line in block[1].split('\n'):
                add_par(line, style='Code')
            add_par('')                       # пустая строка после вывода команды
        elif kind == 'fig':
            fig_no[0] += 1
            path_img = os.path.join(SHOTS, block[1])
            w, h = img_size(path_img, PAGE['img_width'], PAGE['img_max_height'])
            href = doc.addPictureFromFile(path_img)
            p = P(stylename='FigPara')
            frame = with_attrs(Frame(), {
                'draw:name': 'Picture%d' % fig_no[0],
                'text:anchor-type': 'as-char',
                'svg:width': cm(w),
                'svg:height': cm(h),
                'draw:z-index': '0',
            })
            frame.addElement(DrawImage(href=href, type='simple', show='embed',
                                       actuate='onLoad'))
            p.addElement(frame)
            body.addElement(p)
            add_par('Рисунок %d – %s' % (fig_no[0], block[2]), style='FigCaption')
        elif kind == 'table':
            spec = block[1]
            tab_no[0] += 1
            add_par(spec['caption'],
                    style='TabCaptionBreak' if tab_no[0] - 1 in table_breaks
                    else 'TabCaption')
            table = Table(name='Таблица %d' % tab_no[0])
            for i, width in enumerate(spec['widths']):
                col_name = 'Col%d_%d' % (tab_no[0], i)
                col_style = Style(name=col_name, family='table-column')
                col_style.addElement(with_attrs(TableColumnProperties(),
                                            {'style:column-width': cm(width)}))
                doc.automaticstyles.addElement(col_style)
                table.addElement(TableColumn(stylename=col_name))
            for r, row in enumerate([spec['head']] + spec['rows']):
                tr = TableRow()
                for value in row:
                    cell = TableCell(stylename='TblHead' if r == 0 else 'TblCell',
                                     valuetype='string')
                    cp = P(stylename='TblPHead' if r == 0 else 'TblP')
                    cp.addText(u'%s' % value)
                    cell.addElement(cp)
                    tr.addElement(cell)
                table.addElement(tr)
            body.addElement(table)
            add_par('')
        elif kind == 'counts':
            head, tail = block[1].split('{pages}', 1)
            p = P(stylename='Body')
            p.addText(head)
            count = PageCount()
            count.addText(str(page_count if page_count else 1))
            p.addElement(count)
            p.addText(tail.format(figs=sum(1 for b in content.SECTIONS if b[0] == 'fig'),
                                  tables=sum(1 for b in content.SECTIONS if b[0] == 'table'),
                                  sources=len(content.REFS)))
            body.addElement(p)
        elif kind == 'toc':
            for level, text, name in headings:
                p = P(stylename='TOC1' if level == 1 else 'TOC2')
                p.addText(text)
                p.addElement(Tab())
                ref = BookmarkRef(refname=name, referenceformat='page')
                ref.addText(str(page_map.get((level, text), '')))
                p.addElement(ref)
                body.addElement(p)

    doc.meta.addElement(MetaTitle(text='Отчёт о лабораторной работе № 3. '
                                       'Настройка загрузчика GRUB2'))
    doc.meta.addElement(Creator(text='Лернер В. И.'))
    if path:
        doc.save(path)
    return doc, headings


def main():
    os.makedirs(REPORT_DIR, exist_ok=True)
    path = os.path.join(REPORT_DIR, ODT_NAME)

    # --- итерации: подбираем таблицы, которые нужно перенести целиком,
    #     и рассчитываем номера страниц для СОДЕРЖАНИЯ
    table_breaks = set()
    for step in range(6):
        build_document(page_map={}, page_count=1, path=path,
                       table_breaks=table_breaks)
        pages, outlines, split = paginate.analyse(path, fonts_dir=FONTS)
        page_map = {(1 if level == 1 else 2, text): number
                    for level, text, number, anchor in outlines}
        total = len(pages)
        if split <= table_breaks:
            break
        table_breaks |= split
        print('шаг %d: переносим таблицы %s целиком'
              % (step + 1, [n + 1 for n in sorted(table_breaks)]))

    # --- итоговая сборка: номера страниц в СОДЕРЖАНИИ и общее число страниц
    build_document(page_map=page_map, page_count=total, path=path,
                   table_breaks=table_breaks)
    pages, outlines, split = paginate.analyse(path, fonts_dir=FONTS)
    print('Страниц в документе: %d' % len(pages))
    for level, text, number, anchor in outlines:
        print('   %s%d  стр. %d  %s' % ('   ' * (level - 1), level, number, text[:70]))
    if split:
        print('ВНИМАНИЕ: таблицы %s всё ещё разрываются между страницами'
              % [n + 1 for n in sorted(split)])
    print('Готово: %s' % path)


if __name__ == '__main__':
    main()

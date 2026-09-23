#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Предпросмотр отчёта в PDF по той же модели разметки, что и paginate.py.

Модуль не заменяет LibreOffice: он рисует страницы так, как их рассчитала
модель пагинации (шрифт Tinos метрически совпадает с Times New Roman,
Courier — с Courier New). Файл нужен для быстрой проверки вёрстки и
содержания; рабочим форматом отчёта остаётся ODT.

Запуск:  python3 preview.py <отчёт.odt> <предпросмотр.pdf>
"""

import os
import sys
import zipfile

from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import paginate as pg  # noqa: E402

FONTS = os.path.normpath(os.path.join(BASE, '..', '..', 'lab-01', 'build', 'fonts'))
CM = 72.0 / 2.54

ASCENT = {'roman': 0.892, 'mono': 0.8325}


def build_pdf(odt_path, pdf_path):
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    for name, fname in (('Tinos', 'Tinos-Regular.ttf'), ('Tinos-Bold', 'Tinos-Bold.ttf'),
                        ('Tinos-Italic', 'Tinos-Italic.ttf'),
                        ('Tinos-BoldItalic', 'Tinos-BoldItalic.ttf')):
        pdfmetrics.registerFont(TTFont(name, os.path.join(FONTS, fname)))
    # моноширинный шрифт с кириллицей — только для предпросмотра
    # (в самом отчёте указан Courier New, как и требует методичка)
    dejavu = '/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf'
    if os.path.exists(dejavu):
        pdfmetrics.registerFont(TTFont('CourierMono', dejavu))
        pdfmetrics.registerFontFamily('CourierMono', normal='CourierMono', bold='CourierMono')
    pdfmetrics.registerFontFamily('Tinos', normal='Tinos', bold='Tinos-Bold',
                                  italic='Tinos-Italic', boldItalic='Tinos-BoldItalic')

    zf = zipfile.ZipFile(odt_path)
    styles = pg.load_styles(zf)
    root_items, _ = pg.read_odt(odt_path, pg.Ctx(fonts_dir=FONTS), styles)

    # геометрия
    import xml.etree.ElementTree as ET
    root = ET.fromstring(zf.read('content.xml'))
    margins, size = (3.0 * CM, 1.5 * CM, 2.0 * CM, 2.0 * CM), (21.0 * CM, 29.7 * CM)
    for pl in ET.fromstring(zf.read('styles.xml')).iter(pg.qn('style', 'page-layout')):
        props = pl.find(pg.qn('style', 'page-layout-properties'))
        if props is None:
            continue
        margins = (pg.parse_len(props.get(pg.qn('fo', 'margin-left')), margins[0]),
                   pg.parse_len(props.get(pg.qn('fo', 'margin-right')), margins[1]),
                   pg.parse_len(props.get(pg.qn('fo', 'margin-top')), margins[2]),
                   pg.parse_len(props.get(pg.qn('fo', 'margin-bottom')), margins[3]))
        size = (pg.parse_len(props.get(pg.qn('fo', 'page-width')), size[0]),
                pg.parse_len(props.get(pg.qn('fo', 'page-height')), size[1]))
        break
    ctx = pg.Ctx(size[0], size[1], margins, fonts_dir=FONTS)
    pages, outlines, _ = pg.layout(root_items, ctx)

    canvas = rl_canvas.Canvas(pdf_path, pagesize=(size[0], size[1]))
    canvas.setTitle('Отчёт ЛР3 (предпросмотр)')

    def font_of(style):
        family = style.get('font_family') or 'Times New Roman'
        if ctx.is_mono(family):
            return 'CourierMono' if os.path.exists(
                '/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf') else 'Courier'
        if style.get('bold') and style.get('italic'):
            return 'Tinos-BoldItalic'
        if style.get('bold'):
            return 'Tinos-Bold'
        if style.get('italic'):
            return 'Tinos-Italic'
        return 'Tinos'

    def ascent_of(style):
        if ctx.is_mono(style.get('font_family')):
            return 0.76 * style.get('font_size')      # DejaVu Sans Mono: 1556+483 / 2048
        return ASCENT['roman'] * style.get('font_size')

    def draw_text_run(x, y_baseline, text, style, width=None, align='left'):
        size_pt = style.get('font_size')
        font = font_of(style)
        canvas.setFont(font, size_pt)
        canvas.setFillColorRGB(0, 0, 0)
        spacing = style.get('letter_spacing', 0.0)
        if spacing:                      # разрежение подзаголовков (3 пт)
            canvas._code.append('%.2f Tc' % spacing)
        if align == 'center':
            x = ctx.ml + (ctx.text_w - style.get('left', 0.0)) / 2.0 \
                - ctx.text_width(text, style) / 2.0
        elif align == 'right' and width is not None:
            x = x + width - ctx.text_width(text, style)
        canvas.drawString(x, y_baseline, text)
        if spacing:
            canvas._code.append('0 Tc')

    for index, page in enumerate(pages, 1):
        # номер страницы (кроме титульного листа)
        if index > 1:
            canvas.setFont('Tinos', 14)
            canvas.drawCentredString(size[0] / 2.0, margins[3] * 0.55, str(index))
        for placed in page.placed:
            item = placed.item
            style = item.style
            if item.kind == 'img':
                import io as _io
                from reportlab.lib.utils import ImageReader
                data = zf.read(item.image)
                image = Image.open(_io.BytesIO(data))
                top = size[1] - placed.y
                canvas.drawImage(ImageReader(image), placed.x, top - item.h,
                                 width=item.w, height=item.h,
                                 preserveAspectRatio=False)
                if item.caption:        # подпись внутри кадра, под рисунком
                    cap_style = item.caption_style or style
                    cap_lh = ctx.line_height(cap_style)
                    lines = ctx.wrap(item.caption, cap_style, width=item.w)
                    for n, line in enumerate(lines):
                        baseline = top - item.h - n * cap_lh - ascent_of(cap_style)
                        draw_text_run(ctx.ml, baseline, line, cap_style, align='center')
                continue
            if item.kind == 'table':
                row = item.table['rows'][placed.meta.get('row_index', 0)]
                widths = item.table['cols']
                total = sum(widths) or ctx.text_w
                widths = [w * ctx.text_w / total for w in widths]
                x = ctx.ml
                canvas.setLineWidth(0.5)
                for cell, width in zip(row, widths):
                    canvas.rect(x, size[1] - (placed.y + placed.height), width,
                                placed.height, stroke=1, fill=0)
                    text, cell_style = cell
                    if text:
                        lines = ctx.wrap(text, cell_style, width=width - 0.16 * CM)
                        font_size = cell_style.get('font_size')
                        lh = ctx.line_height(cell_style)
                        for n, line in enumerate(lines):
                            baseline = size[1] - (placed.y + 0.08 * CM + n * lh +
                                                  cell_style.get('font_size') * 0.892)
                            draw_text_run(x + 0.08 * CM, baseline, line, cell_style)
                    x += width
                continue
            # --- обычный абзац
            lh = ctx.line_height(style)
            align = style.get('align')
            for n, line in enumerate(placed.lines):
                top = placed.y + n * lh
                baseline = size[1] - top - ascent_of(style)
                x = placed.x + (style.get('indent') if n == 0 else 0.0)
                if line == '':
                    continue
                if '\t' in line:            # отточие и номер страницы в СОДЕРЖАНИИ
                    left, right = line.split('\t', 1)
                    draw_text_run(x, baseline, left, style)
                    width = ctx.text_w - style.get('left', 0.0)
                    draw_text_run(x, baseline, right.strip(), style,
                                  width=width, align='right')
                    continue
                if align == 'center':
                    draw_text_run(x, baseline, line, style, align='center')
                elif align == 'end':
                    draw_text_run(x, baseline, line, style,
                                  width=ctx.text_w - style.get('left', 0.0),
                                  align='right')
                else:
                    draw_text_run(x, baseline, line, style)
        canvas.showPage()
    canvas.save()
    return len(pages)


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 1
    total = build_pdf(argv[1], argv[2])
    print('Предпросмотр: %s (%d с.)' % (argv[2], total))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))

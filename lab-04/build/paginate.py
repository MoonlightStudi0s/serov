# -*- coding: utf-8 -*-
"""Разметка (пагинация) ODT-документа: определяет, на какой странице окажется
каждый заголовок. Нужна для того, чтобы номера страниц в СОДЕРЖАНИИ отчёта
были вычислены без запуска LibreOffice.

Модуль читает готовый ODT-файл (content.xml + styles.xml), восстанавливает
параметры абзацев с учётом наследования стилей и «раскладывает» поток по
страницам формата A4 с заданными полями.

Модель линейки:

    высота строки = (ascent + descent) / em * размер шрифта * (line-height / 100)

Так же считает и LibreOffice: межстрочный интервал «150 %» берётся не от
кегля шрифта, а от его собственной высоты строки (для Times New Roman / Tinos
это 1,109 em).

Модуль можно запускать из командной строки для проверки готового документа:

    python3 paginate.py ../report/Отчёт_ЛР3_….odt
"""

import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

from PIL import ImageFont

# --------------------------------------------------------------------- единицы
CM = 72.0 / 2.54          # пунктов в сантиметре
UNITS = {
    'cm': CM, 'mm': CM / 10.0, 'in': 72.0, 'pt': 1.0, 'pc': 12.0, 'px': 0.75,
    '': 1.0,
}

NS = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    'xlink': 'http://www.w3.org/1999/xlink',
}


def qn(prefix, tag):
    return '{%s}%s' % (NS[prefix], tag)


def parse_len(value, base=0.0):
    """'1.25cm' -> пункты; None -> base."""
    if value is None:
        return base
    value = value.strip()
    m = re.match(r'^(-?[\d.]+)(cm|mm|in|pt|pc|px)?$', value)
    if not m:
        return base
    num, unit = m.group(1), m.group(2) or ''
    return float(num) * UNITS.get(unit, 1.0)


def parse_percent(value, base=100.0):
    if value is None:
        return base
    m = re.match(r'^(-?[\d.]+)%$', value.strip())
    return float(m.group(1)) if m else base


# ---------------------------------------------------------------------- стили
class Style(object):
    """Разрешённые (с учётом наследования) свойства абзацного стиля."""

    DEFAULTS = {
        'font_family': 'Times New Roman',
        'font_size': 12.0,
        'bold': False,
        'italic': False,
        'align': 'start',
        'indent': 0.0,          # первая строка, пт
        'left': 0.0,            # отступ слева, пт
        'right': 0.0,
        'line_height': 100.0,   # проценты от высоты строки шрифта
        'sb': 0.0,              # интервал перед абзацем, пт
        'sa': 0.0,              # интервал после абзаца, пт
        'keep_next': False,
        'break_before': False,
        'letter_spacing': 0.0,  # межбуквенный интервал, пт
        'orphans': 2,
        'widows': 2,
    }

    def __init__(self, name, parent=None):
        self.name = name
        self.parent = parent
        self.props = dict(self.DEFAULTS)

    def get(self, key, default=None):
        return self.props.get(key, default)

    def __repr__(self):
        return '<Style %s %spt %s>' % (self.name, self.props['font_size'],
                                       self.props['align'])


def _text_props(el):
    out = {}
    if el is None:
        return out
    get = el.get
    if get(qn('fo', 'font-size')):
        out['font_size'] = parse_len(get(qn('fo', 'font-size')), 12.0)
    if get(qn('fo', 'font-family')):
        out['font_family'] = get(qn('fo', 'font-family')).strip('\'"')
    if get(qn('fo', 'font-weight')) == 'bold':
        out['bold'] = True
    if get(qn('fo', 'font-style')) == 'italic':
        out['italic'] = True
    if get(qn('fo', 'letter-spacing')):
        out['letter_spacing'] = parse_len(get(qn('fo', 'letter-spacing')), 0.0)
    return out


def _para_props(el):
    out = {}
    if el is None:
        return out
    get = el.get
    if get(qn('fo', 'text-align')):
        out['align'] = get(qn('fo', 'text-align'))
    if get(qn('fo', 'text-indent')):
        out['indent'] = parse_len(get(qn('fo', 'text-indent')))
    if get(qn('fo', 'margin-left')):
        out['left'] = parse_len(get(qn('fo', 'margin-left')))
    if get(qn('fo', 'margin-right')):
        out['right'] = parse_len(get(qn('fo', 'margin-right')))
    if get(qn('fo', 'margin-top')):
        out['sb'] = parse_len(get(qn('fo', 'margin-top')))
    if get(qn('fo', 'margin-bottom')):
        out['sa'] = parse_len(get(qn('fo', 'margin-bottom')))
    if get(qn('fo', 'line-height')):
        lh = get(qn('fo', 'line-height'))
        if lh.endswith('%'):
            out['line_height'] = parse_percent(lh)
        else:                                  # фиксированное значение
            out['line_height'] = ('fixed', parse_len(lh))
    if get(qn('fo', 'keep-with-next')) == 'always':
        out['keep_next'] = True
    if get(qn('fo', 'break-before')) == 'page':
        out['break_before'] = True
    for key, attr in (('orphans', 'orphans'), ('widows', 'widows')):
        if get(qn('fo', attr)):
            try:
                out[key] = int(get(qn('fo', attr)))
            except ValueError:
                pass
    return out


def load_styles(zf):
    """Читает styles.xml и автоматические стили content.xml."""
    styles = {}
    defaults = {}

    def collect(root, skip_defaults=False):
        for st in root.iter(qn('style', 'style')):
            if st.get(qn('style', 'family')) not in ('paragraph', 'table-column', None):
                continue
            name = st.get(qn('style', 'name'))
            if name is None:
                continue
            obj = Style(name, st.get(qn('style', 'parent-style-name')))
            obj.props.update(_para_props(st.find(qn('style', 'paragraph-properties'))))
            obj.props.update(_text_props(st.find(qn('style', 'text-properties'))))
            col = st.find(qn('style', 'table-column-properties'))
            if col is not None and col.get(qn('style', 'column-width')):
                obj.props['column_width'] = parse_len(col.get(qn('style', 'column-width')), 0.0)
            styles[name] = obj

    for part in ('styles.xml', 'content.xml'):
        if part in zf.namelist():
            root = ET.fromstring(zf.read(part))
            collect(root)
            for st in root.iter(qn('style', 'default-style')):
                if st.get(qn('style', 'family')) in ('paragraph', None):
                    defaults.update(_para_props(st.find(qn('style', 'paragraph-properties'))))
                    defaults.update(_text_props(st.find(qn('style', 'text-properties'))))

    # наследование свойств от родительских стилей
    def resolved(name, seen=None):
        seen = seen or set()
        st = styles.get(name)
        if st is None or name in seen:
            return dict(Style.DEFAULTS)
        seen.add(name)
        props = resolved(st.parent, seen) if st.parent else dict(Style.DEFAULTS)
        props.update(defaults)
        props.update({k: v for k, v in st.props.items()
                      if v != Style.DEFAULTS.get(k, None) or k not in Style.DEFAULTS})
        return props

    for name, st in styles.items():
        st.props = resolved(name)
    return styles


# ------------------------------------------------------------------- геометрия
class Ctx(object):
    """Геометрия страницы, метрики шрифтов и перенос строк."""

    def __init__(self, page_w=None, page_h=None, margins=None, fonts_dir=None,
                 default_styles=None):
        self.page_w = page_w or 21.0 * CM
        self.page_h = page_h or 29.7 * CM
        m = margins or (3.0 * CM, 1.5 * CM, 2.0 * CM, 2.0 * CM)
        self.ml, self.mr, self.mt, self.mb = m
        self.text_w = self.page_w - self.ml - self.mr
        self.text_h = self.page_h - self.mt - self.mb
        self.fonts_dir = fonts_dir
        self.default_styles = default_styles or {}
        self._cache = {}

    # --- метрики -----------------------------------------------------------
    def _pil(self, family, bold, italic):
        key = (family, bold, italic)
        if key in self._cache:
            return self._cache[key]
        font = None
        if self.fonts_dir and 'times' in (family or '').lower():
            fname = 'Tinos-%s.ttf' % ('BoldItalic' if bold and italic else
                                      'Bold' if bold else
                                      'Italic' if italic else 'Regular')
            path = os.path.join(self.fonts_dir, fname)
            if os.path.exists(path):
                font = ImageFont.truetype(path, 1000)
        self._cache[key] = font
        return font

    def is_mono(self, family):
        return 'courier' in (family or '').lower() or 'mono' in (family or '').lower() \
            or 'consol' in (family or '').lower()

    def text_width(self, text, style):
        size = style.get('font_size')
        family = style.get('font_family')
        if self.is_mono(family):
            w = 0.6 * size * len(text)          # Courier New — моноширинный 600/1000 em
        else:
            font = self._pil(family, style.get('bold'), style.get('italic'))
            if font is None:                    # неизвестный шрифт — грубая оценка
                w = 0.5 * size * len(text)
            else:
                w = font.getlength(text) / 1000.0 * size
        return w + style.get('letter_spacing', 0.0) * max(len(text) - 1, 0)

    def line_height(self, style):
        size = style.get('font_size')
        lh = style.get('line_height')
        if isinstance(lh, tuple):
            return lh[1]
        family = style.get('font_family')
        if self.is_mono(family):
            em = 1.1328                          # Courier New: 1705 + 615 / 2048
        else:
            em = 1.109                           # Times New Roman / Tinos
        return em * size * lh / 100.0

    # --- перенос строк -----------------------------------------------------
    def wrap(self, text, style, width=None, indent=None):
        """Список строк после переноса по словам."""
        width = self.text_w if width is None else width
        first = style.get('indent') if indent is None else indent
        width -= style.get('left', 0.0)
        words = re.split(r'(\s+)', text.strip())
        lines, cur, avail = [], '', width - max(first, 0.0)
        for chunk in words:
            if not chunk:
                continue
            candidate = cur + chunk
            if self.text_width(candidate.rstrip(), style) <= avail or not cur.strip():
                if self.text_width(candidate.rstrip(), style) > avail and not cur.strip():
                    # слово длиннее строки — режем по символам
                    piece = ''
                    for ch in chunk:
                        if self.text_width(piece + ch, style) > avail and piece:
                            lines.append(piece)
                            piece = ch
                            avail = width
                        else:
                            piece += ch
                    cur = piece
                    continue
                cur = candidate
            else:
                lines.append(cur.rstrip())
                avail = width
                cur = chunk.lstrip()
        if cur.strip():
            lines.append(cur.rstrip())
        return lines or ['']


# ------------------------------------------------------------------- элементы
class Item(object):
    """Элемент потока: абзац, рисунок или таблица."""

    def __init__(self, kind, style, **kw):
        self.kind = kind
        self.style = style
        self.text = kw.get('text', '')
        self.image = kw.get('image')
        self.w = kw.get('w', 0.0)
        self.h = kw.get('h', 0.0)
        self.table = kw.get('table')
        self.anchor = kw.get('anchor')
        self.outline = kw.get('outline', 0)
        self.table_index = kw.get('table_index', -1)
        self.keep_next = kw.get('keep_next', style.get('keep_next'))
        self.break_before = kw.get('break_before', style.get('break_before'))


def _para_text(p):
    """Текст абзаца с учётом <text:s/>, <text:tab/> и <text:line-break/>."""
    out = []
    if p.text:
        out.append(p.text)
    for child in p:
        tag = child.tag
        if tag == qn('text', 's'):
            out.append(' ' * int(child.get(qn('text', 'c')) or 1))
        elif tag == qn('text', 'tab'):
            out.append('\t')
        elif tag == qn('text', 'line-break'):
            out.append('\n')
        elif tag in (qn('draw', 'frame'), qn('draw', 'custom-shape')):
            continue
        else:
            out.append(_para_text(child) if len(child) else (child.text or ''))
        if child.tail:
            out.append(child.tail)
    return ''.join(out)


def _frame_size(frame):
    w = frame.get(qn('svg', 'width'))
    h = frame.get(qn('svg', 'height'))
    href = None
    image = frame.find(qn('draw', 'image'))
    if image is not None:
        href = image.get(qn('xlink', 'href'))
    return (parse_len(w, 0.0), parse_len(h, 0.0), href)


def read_odt(path, ctx, styles):
    """Преобразует тело ODT в список элементов потока."""
    zf = zipfile.ZipFile(path)
    root = ET.fromstring(zf.read('content.xml'))
    body = root.find(qn('office', 'body'))
    text = body.find(qn('office', 'text'))
    items = []
    outline = 0
    for el in text:
        tag = el.tag
        if tag == qn('text', 'h'):
            level = int(el.get(qn('text', 'outline-level')) or 1)
            outline += 1
            name = el.get(qn('text', 'style-name'), 'Standard')
            items.append(Item('p', styles.get(name, Style(name)),
                              text=_para_text(el), outline=level,
                              anchor=_bookmark_of(el)))
        elif tag == qn('text', 'p'):
            name = el.get(qn('text', 'style-name'), 'Standard')
            style = styles.get(name, Style(name))
            frames = el.findall(qn('draw', 'frame'))
            if frames:
                for fr in frames:
                    w, h, href = _frame_size(fr)
                    items.append(Item('img', style, image=href, w=w, h=h))
                continue
            # заголовки распознаются по именам стилей отчёта: H1/H1Break/H2/StructH
            level = 2 if name.startswith('H2') else (
                1 if (name.startswith('H1') or name == 'StructH') else 0)
            items.append(Item('p', style, text=_para_text(el), outline=level,
                              anchor=_bookmark_of(el)))
        elif tag == qn('table', 'table'):
            n_tables = sum(1 for it in items if it.kind == 'table')
            items.append(Item('table', styles.get('TableCell', Style('TableCell')),
                              table=_read_table(el, ctx, styles),
                              table_index=n_tables))
        elif tag == qn('text', 'table-of-content'):
            items.append(Item('toc', styles.get('Standard', Style('Standard')),
                              text='(СОДЕРЖАНИЕ)'))
        elif tag == qn('text', 'section'):
            pass
    return items, zf


def _bookmark_of(el):
    bm = el.find(qn('text', 'bookmark-start'))
    return bm.get(qn('text', 'name')) if bm is not None else None


def _read_table(el, ctx, styles):
    cols = []
    for col in el.iter(qn('table', 'table-column')):
        st = styles.get(col.get(qn('table', 'style-name')))
        cols.append(float(st.props.get('column_width') or 0.0) if st else 0.0)
    rows = []
    for tr in el.findall(qn('table', 'table-row')):
        cells = []
        for tc in tr.findall(qn('table', 'table-cell')):
            txt = '\n'.join(_para_text(p) for p in tc.findall(qn('text', 'p')))
            st_name = tc.get(qn('table', 'style-name'))
            st = styles.get(st_name, None) or Style(st_name or 'TableCell')
            cells.append((txt, st))
        rows.append(cells)
    return {'cols': cols, 'rows': rows}


# -------------------------------------------------------------------- разметка
class Placed(object):
    def __init__(self, item, x, y, lines=None, height=0.0, meta=None):
        self.item = item
        self.x = x
        self.y = y
        self.lines = lines or []
        self.height = height
        self.meta = meta or {}


class Page(object):
    def __init__(self, number):
        self.number = number
        self.placed = []
        self.left = 0.0
    @property
    def used(self):
        return sum(p.height for p in self.placed)


def layout(items, ctx, start_page=1):
    """Раскладка потока по страницам. Возвращает (страницы, карта заголовков)."""
    pages = [Page(start_page)]
    outlines = []
    split_tables = set()
    i = 0
    items = list(items)

    def new_page():
        pages.append(Page(pages[-1].number + 1))
        return pages[-1]

    def free(page):
        return ctx.text_h - page.used

    def place(page, item, x, y, lines, height, meta=None):
        page.placed.append(Placed(item, x, y, lines, height, meta))
        if item.outline:
            outlines.append((item.outline, item.text, page.number, item.anchor))

    while i < len(items):
        item = items[i]
        style = item.style
        lh = ctx.line_height(style)

        # явный разрыв страницы (fo:break-before="page" или style:page-number="auto")
        if item.break_before and pages[-1].placed:
            new_page()

        if item.kind == 'img':
            x = (ctx.text_w - item.w) / 2.0
            block = item.h + style.get('sb') + 0.0
            nxt = items[i + 1] if i + 1 < len(items) and items[i + 1].kind == 'p' else None
            if item.keep_next and nxt is not None:
                cap_lh = ctx.line_height(nxt.style)
                cap_lines = ctx.wrap(nxt.text, nxt.style)
                block += len(cap_lines) * cap_lh + nxt.style.get('sa')
            page = pages[-1]
            if block > free(page) and free(page) < ctx.text_h - 1e-6:
                page = new_page()
            y = ctx.mt + page.used + style.get('sb')
            place(page, item, x, y, [], item.h + style.get('sb'))
            i += 1
            continue

        if item.kind == 'table':
            rows = item.table['rows']
            page = pages[-1]
            y = ctx.mt + page.used + style.get('sb')
            used = style.get('sb')
            caption_page = None
            if page.placed and page.placed[-1].item.kind == 'p':
                prev = page.placed[-1].item
                if prev.style.name.startswith('TabCaption'):
                    caption_page = page.number
            table_pages = set()
            for row_index, row in enumerate(rows):
                rh = 0.0
                for idx, (txt, st) in enumerate(row):
                    cw = item.table['cols'][idx] if idx < len(item.table['cols']) else ctx.text_w / max(len(row), 1)
                    cw = cw or ctx.text_w / max(len(row), 1)
                    n = len(ctx.wrap(txt, st, width=cw - 2 * 0.097 * 0 + -0.0)) if txt else 1
                    rh = max(rh, n * ctx.line_height(st) + 0.12 * CM)
                if used + rh > free(page) and page.placed:
                    page = new_page()
                    used = 0.0
                    y = ctx.mt
                place(page, item, ctx.ml, y, [], rh, {'row_index': row_index})
                table_pages.add(page.number)
                used += rh
                y += rh
            if style.get('sa'):
                place(page, item, ctx.ml, y, [], style.get('sa'))
            if len(table_pages) > 1 or (caption_page is not None and
                                        min(table_pages) != caption_page):
                split_tables.add(item.table_index)
            i += 1
            continue

        if item.kind == 'toc':
            # содержание: N строк по числу записей (оценка)
            n = item.table['n'] if item.table else 20
            h = n * lh
            page = pages[-1]
            if h > free(page) and page.placed:
                page = new_page()
            place(page, item, ctx.ml, ctx.mt + page.used, [], h)
            i += 1
            continue

        # --- обычный абзац
        lines = []
        for part in item.text.split('\n'):
            lines.extend(ctx.wrap(part, style))
        if not item.text:
            lines = ['']
        lh_total = len(lines) * lh
        sb = style.get('sb')
        page = pages[-1]

        # keep-with-next: заголовок не отрывается от следующего блока
        if item.keep_next and i + 1 < len(items):
            nxt = items[i + 1]
            need = sb + lh_total + style.get('sa')
            if nxt.kind == 'img':
                need += nxt.h + (ctx.line_height(nxt.style) if nxt.keep_next else 0)
            elif nxt.kind == 'table':
                need += sum(max(ctx.line_height(st) for _, st in row)
                            for row in nxt.table['rows'][:2])
            else:
                first_lines = ctx.wrap(nxt.text, nxt.style)
                need += sum(ctx.line_height(nxt.style) for _ in first_lines[:2])
            if need > free(page) and page.placed:
                page = new_page()

        if sb + lh_total + style.get('sa') > free(page) and page.placed:
            # переносим целиком либо разрываем по правилам orphans/widows
            orphans = style.get('orphans')
            widows = style.get('widows')
            fits = int(max(free(page) - sb, 0) // lh)
            if len(lines) > orphans + widows and fits >= orphans + 1:
                fits = min(fits, len(lines) - widows)
                y = ctx.mt + page.used + sb
                place(page, item, ctx.ml, y, lines[:fits], fits * lh + sb)
                page = new_page()
                rest = lines[fits:]
                y = ctx.mt
                place(page, item, ctx.ml, y, rest, len(rest) * lh + style.get('sa'))
                i += 1
                continue
            page = new_page()

        y = ctx.mt + page.used + sb
        place(page, item, ctx.ml, y, lines, lh_total + sb + style.get('sa'))
        i += 1

    return pages, outlines, split_tables


def analyse(path, fonts_dir=None):
    """Разметка ODT-файла. Возвращает (число страниц, список заголовков)."""
    zf = zipfile.ZipFile(path)
    styles = load_styles(zf)
    root = ET.fromstring(zf.read('content.xml'))
    # поля страницы из автоматических стилей документа
    auto = root.find(qn('office', 'automatic-styles'))
    margins, size = (3.0 * CM, 1.5 * CM, 2.0 * CM, 2.0 * CM), (21.0 * CM, 29.7 * CM)
    if auto is not None:
        for pl in auto.iter(qn('style', 'page-layout')):
            props = pl.find(qn('style', 'page-layout-properties'))
            if props is None:
                continue
            margins = (parse_len(props.get(qn('fo', 'margin-left')), margins[0]),
                       parse_len(props.get(qn('fo', 'margin-right')), margins[1]),
                       parse_len(props.get(qn('fo', 'margin-top')), margins[2]),
                       parse_len(props.get(qn('fo', 'margin-bottom')), margins[3]))
            size = (parse_len(props.get(qn('fo', 'page-width')), size[0]),
                    parse_len(props.get(qn('fo', 'page-height')), size[1]))
            break
    for pl in zf.namelist():
        pass
    styles_root = ET.fromstring(zf.read('styles.xml'))
    for pl in styles_root.iter(qn('style', 'page-layout')):
        props = pl.find(qn('style', 'page-layout-properties'))
        if props is None:
            continue
        margins = (parse_len(props.get(qn('fo', 'margin-left')), margins[0]),
                   parse_len(props.get(qn('fo', 'margin-right')), margins[1]),
                   parse_len(props.get(qn('fo', 'margin-top')), margins[2]),
                   parse_len(props.get(qn('fo', 'margin-bottom')), margins[3]))
        size = (parse_len(props.get(qn('fo', 'page-width')), size[0]),
                parse_len(props.get(qn('fo', 'page-height')), size[1]))
        break
    ctx = Ctx(size[0], size[1], margins, fonts_dir=fonts_dir)
    items, _ = read_odt(path, ctx, styles)
    return layout(items, ctx)


def main(argv):
    fonts = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..',
                         'lab-01', 'build', 'fonts')
    for path in argv[1:]:
        pages, outlines, split = analyse(path, fonts_dir=fonts)
        print('=== %s' % path)
        if split:
            print('таблицы, разорванные между страницами: %s' % [n + 1 for n in sorted(split)])
        print('страниц: %d' % len(pages))
        for level, text, num, anchor in outlines:
            print('  %s%d  стр. %d  %s' % ('    ' * (level - 1), level, num,
                                           text[:70]))


if __name__ == '__main__':
    main(sys.argv)

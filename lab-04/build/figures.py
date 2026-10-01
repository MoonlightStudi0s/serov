# -*- coding: utf-8 -*-
"""Рисунки отчёта: подпись «Рисунок N – Название» вписывается в изображение.

По ГОСТ 7.32-2017 подпись располагается под рисунком; здесь она не выносится
отдельным абзацем, а входит в состав самого изображения — в нижнюю белую полосу,
пристроенную к скриншоту. Такой рисунок нельзя «разорвать» с подписью при
переносе на другую страницу, и подпись всегда остаётся вместе с картинкой.

Подпись набирается шрифтом Tinos (метрически совпадает с Times New Roman)
кеглем 14 пт *в масштабе страницы*: изображение печатается шириной width_sm,
плотность пикселей dpi = ширина_в_пикселах / (width_sm / 2,54), поэтому высота
шрифта в пикселах подбирается так, чтобы на странице подпись имела ровно 14 пт.
Полоса с подписью рисуется с четырёхкратным увеличением и уменьшается обратно —
текст получается чётким.

Исходные скриншоты не изменяются: готовые изображения складываются в каталог
build/figures (в репозиторий не попадает) и оттуда вкладываются в ODT.

Модуль можно запускать из командной строки для проверки:

    python3 figures.py            # собрать все подписи в build/figures
"""

import hashlib
import os

from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(BASE)
FONTS = os.path.normpath(os.path.join(BASE, '..', '..', 'lab-01', 'build', 'fonts'))
OUT_DIR = os.path.join(BASE, 'figures')

# --------------------------------------------------------------- параметры подписи
CAPTION = {
    'size': 14.0,          # кегль подписи на странице, пт
    'leading': 1.15,       # межстрочный интервал подписи, долей кегля
    'pad': 6.0,            # отступ от края полосы до текста, пт
    'side': 0.20,          # боковые поля подписи, см
    'color': (0, 0, 0),
    'fill': (255, 255, 255),
    'supersample': 4,      # во сколько раз крупнее рисуется полоса
}


def _font(size_px, bold=False):
    name = 'Tinos-%s.ttf' % ('Bold' if bold else 'Regular')
    return ImageFont.truetype(os.path.join(FONTS, name), max(int(round(size_px)), 1))


def wrap(text, font, max_width):
    """Перенос строки по словам под заданную ширину (в пикселах)."""
    lines, current = [], ''
    for word in text.split():
        candidate = (current + ' ' + word).strip()
        if font.getlength(candidate) <= max_width or not current:
            current = candidate
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or ['']


def caption_lines(number, caption):
    return 'Рисунок %d – %s' % (number, caption)


def captioned(path, number, caption, out_dir=OUT_DIR, width_cm=16.5,
              max_height_cm=21.0, min_dpi=98.0, params=None):
    """Изображение скриншота с подписью, вписанной в нижнюю белую полосу.

    Возвращает (путь к изображению, ширина в см, высота в см) — размер на
    странице с сохранением пропорций: мелкие диалоговые окна не растягиваются
    на всю ширину текста (плотность не ниже min_dpi).

    Готовые файлы кэшируются: имя содержит отпечаток параметров, поэтому
    повторная сборка отчёта изображения не перерисовывает.
    """
    par = dict(CAPTION)
    par.update(params or {})
    with Image.open(path) as im:
        source = im.convert('RGB')
        px_w, px_h = source.size

        # --- размер на странице и плотность пикселей
        width = min(float(width_cm), px_w * 2.54 / float(min_dpi))
        for _ in range(3):
            dpi = px_w / (width / 2.54)
            font_px = par['size'] * dpi / 72.0
            side_px = par['side'] / 2.54 * dpi
            measure = _font(font_px)
            lines = wrap(caption_lines(number, caption), measure,
                         px_w - 2 * side_px)
            band_pt = len(lines) * par['size'] * par['leading'] + 2 * par['pad']
            band_cm = band_pt / 72.0 * 2.54
            total = width * px_h / float(px_w) + band_cm
            if total <= max_height_cm or width <= 1.0:
                break
            # изображение вместе с подписью выше текстовой полосы — уменьшаем
            width *= max_height_cm / total
        band_px = max(int(round(band_cm / 2.54 * dpi)), 1)

        # --- полоса с подписью (рисуем крупно, затем уменьшаем)
        scale = int(par['supersample'])
        band = Image.new('RGB', (px_w * scale, band_px * scale), par['fill'])
        draw = ImageDraw.Draw(band)
        draw_font = _font(font_px * scale)
        line_px = par['size'] * par['leading'] * dpi / 72.0 * scale
        pad_px = par['pad'] * dpi / 72.0 * scale
        for i, line in enumerate(lines):
            draw.text((px_w * scale / 2.0, pad_px + i * line_px), line,
                      font=draw_font, fill=par['color'], anchor='ma')
        band = band.resize((px_w, band_px), Image.LANCZOS)

        # --- скриншот + полоса
        result = Image.new('RGB', (px_w, px_h + band_px), par['fill'])
        result.paste(source, (0, 0))
        result.paste(band, (0, px_h))

    key = hashlib.sha1()
    key.update(('%s|%s|%d|%s|%.3f|%.3f|%.3f|%.1f'
                % (os.path.basename(path), os.path.getmtime(path), number, caption,
                   width_cm, max_height_cm, min_dpi,
                   par['size'])).encode('utf-8'))
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    out_path = os.path.join(out_dir, 'fig%02d_%s.png' % (number, key.hexdigest()[:12]))
    if not os.path.exists(out_path):
        for old in os.listdir(out_dir):            # убираем устаревшие варианты
            if old.startswith('fig%02d_' % number):
                os.remove(os.path.join(out_dir, old))
        result.save(out_path, 'PNG', optimize=True)
    return out_path, round(width, 3), round(width * (px_h + band_px) / float(px_w), 3)


def main():
    import content
    number = 0
    for block in content.SECTIONS:
        if block[0] != 'fig':
            continue
        number += 1
        path, width, height = captioned(os.path.join(LAB, 'screenshots', block[1]),
                                        number, block[2])
        print('%2d  %5.2f × %5.2f см  %s' % (number, width, height,
                                            os.path.basename(path)))


if __name__ == '__main__':
    main()

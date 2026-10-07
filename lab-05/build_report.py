#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генератор отчёта ЛР №5 — Администрирование Calculate Linux Desktop."""
import tempfile, shutil
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from odf.opendocument import OpenDocumentText
from odf.text import P, Span, BookmarkStart, BookmarkEnd, PageNumber, A as TextA
from odf.style import Style, TextProperties, ParagraphProperties, PageLayout, PageLayoutProperties, MasterPage, Footer, FontFace
from odf.draw import Frame, Image as ODFImage
from odf import dc

BASE = Path(__file__).resolve().parent
SHOTS = BASE / 'screenshots'
REPORT = BASE / 'report'

def _tp(**kw): return TextProperties(**kw)
def _pp(**kw): return ParagraphProperties(**kw)

# ── Изображения с подписями ──
def make_img_caption(img_path, caption, tmp_dir):
    img = Image.open(img_path).convert('RGB')
    w, h = img.size
    fsz = max(14, int(h * 0.032))
    for fp in ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
               "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"]:
        try: font = ImageFont.truetype(fp, fsz); break
        except: font = ImageFont.load_default()
    draw = ImageDraw.Draw(img)
    maxw = int(w * 0.95)
    words = caption.split()
    lines, ln = [], ''
    for wd in words:
        t = f'{ln} {wd}'.strip()
        if draw.textbbox((0,0), t, font=font)[2] > maxw and ln:
            lines.append(ln); ln = wd
        else: ln = t
    if ln: lines.append(ln)
    th = sum(draw.textbbox((0,0), l, font=font)[3] - draw.textbbox((0,0), l, font=font)[1] for l in lines) + len(lines)*4
    bh = th + 24
    c = Image.new('RGB', (w, h+bh), (255,255,255))
    c.paste(img, (0,0))
    d = ImageDraw.Draw(c)
    d.rectangle([0,h,w,h+bh], fill=(235,235,235))
    y = h + 12
    for l in lines:
        bb = d.textbbox((0,0), l, font=font)
        d.text(((w-bb[2]+bb[0])//2, y), l, fill=(40,40,40), font=font)
        y += bb[3]-bb[1]+4
    out = tmp_dir / img_path.name
    c.save(str(out), quality=90)
    return out

def add_fig(doc, body, img_path, fnum, caption, tmp_dir, wcm=15):
    out = make_img_caption(img_path, caption, tmp_dir)
    img = Image.open(out)
    w, h = img.size
    hcm = wcm * h / w
    p = P(stylename='CenterNoIndent')
    p.addElement(BookmarkStart(name=f'fig{fnum}'))
    p.addElement(BookmarkEnd(name=f'fig{fnum}'))
    fr = Frame(width=f'{wcm}cm', height=f'{hcm}cm', anchortype='paragraph')
    href = doc.addPicture(str(out))
    fr.addElement(ODFImage(href=href))
    p.addElement(fr)
    body.addElement(p)

def add_p(body, text, style='Normal'):
    p = P(stylename=style)
    if text: p.addText(text)
    body.addElement(p)

def add_p_bm(body, text, style, bm=None):
    p = P(stylename=style)
    if bm:
        p.addElement(BookmarkStart(name=bm))
        p.addElement(BookmarkEnd(name=bm))
    if text: p.addText(text)
    body.addElement(p)

def add_code(body, text):
    for line in text.split('\n'):
        add_p(body, line, 'CodeBlock')

def add_link(body, text, href, style='Normal', link_text=None):
    """Параграф с гиперссылкой."""
    p = P(stylename=style)
    if text: p.addText(text)
    a = TextA(href=href)
    sp = Span(stylename='RefGray')
    sp.addText(link_text or '')
    a.addElement(sp)
    p.addElement(a)
    body.addElement(p)

# ── Стили ──
def setup_styles(doc):
    doc.fontfacedecls.addElement(FontFace(name='Times New Roman', fontfamily='Times New Roman', fontfamilygeneric='roman', fontpitch='variable'))

    def add_s(name, parent, **kw):
        s = Style(name=name, family='paragraph', parentstylename=parent)
        if 'pp' in kw: s.addElement(_pp(**kw['pp']))
        if 'tp' in kw: s.addElement(_tp(**kw['tp']))
        doc.styles.addElement(s)

    add_s('Normal', 'Standard',
          pp=dict(lineheight='150%', textalign='justify', textindent='1.25cm', marginbottom='0cm', margintop='0cm'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt'))
    add_s('Heading1', 'Heading 1',
          pp=dict(lineheight='150%', textalign='left', textindent='1.25cm', marginbottom='0.35cm', keepwithnext='true'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt', fontweight='bold'))
    add_s('Heading2', 'Heading 2',
          pp=dict(lineheight='150%', textalign='left', textindent='1.25cm', marginbottom='0.35cm', keepwithnext='true'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt', fontweight='bold'))
    add_s('StructHeading', 'Standard',
          pp=dict(lineheight='150%', textalign='center', textindent='0cm', breakbefore='page', marginbottom='0.35cm', keepwithnext='true'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt', fontweight='bold', texttransform='uppercase'))
    add_s('CenterNoIndent', 'Standard',
          pp=dict(lineheight='150%', textalign='center', textindent='0cm', marginbottom='0cm', margintop='0cm'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt'))
    add_s('TitlePage', 'Standard',
          pp=dict(lineheight='115%', textalign='center', textindent='0cm', marginbottom='0.2cm', margintop='0cm'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt'))
    add_s('TitleBold', 'TitlePage', tp=dict(fontweight='bold'))
    add_s('TitleBold16', 'TitlePage', tp=dict(fontsize='16pt', fontsizeasian='16pt', fontweight='bold'))
    add_s('TitleUnderline', 'TitlePage', pp=dict(lineheight='130%'))
    add_s('CodeBlock', 'Standard',
          pp=dict(lineheight='120%', textalign='left', textindent='0cm', backgroundcolor='#f0f0f0',
                  marginleft='0.5cm', marginright='0.5cm', margintop='0.2cm', marginbottom='0.2cm'),
          tp=dict(fontfamily='Courier New', fontsize='11pt', fontsizeasian='11pt'))
    add_s('TOC1', 'Standard',
          pp=dict(lineheight='150%', textalign='left', textindent='0cm', marginleft='0cm', marginbottom='0.05cm', margintop='0.05cm'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt', fontweight='bold'))
    add_s('TOC2', 'Standard',
          pp=dict(lineheight='150%', textalign='left', textindent='0cm', marginleft='0.75cm', marginbottom='0.05cm', margintop='0.05cm'),
          tp=dict(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt'))

    rg = Style(name='RefGray', family='text')
    rg.addElement(_tp(color='#808080'))
    doc.styles.addElement(rg)

    std = Style(name='Standard', family='text')
    std.addElement(_pp(lineheight='150%', textalign='justify', textindent='1.25cm', marginbottom='0cm', margintop='0cm'))
    std.addElement(_tp(fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt', fontsizecomplex='14pt'))
    doc.styles.addElement(std)

def setup_layout(doc):
    pl = PageLayout(name='MPl')
    pl.addElement(PageLayoutProperties(pagewidth='21cm', pageheight='29.7cm',
        marginleft='3cm', marginright='1.5cm', margintop='2cm', marginbottom='2cm', printorientation='portrait'))
    doc.automaticstyles.addElement(pl)
    mp = MasterPage(name='Standard', pagelayoutname=pl)
    doc.masterstyles.addElement(mp)
    ft = Footer()
    ftp = P(stylename='CenterNoIndent')
    ftp.addElement(PageNumber())
    ft.addElement(ftp)
    mp.addElement(ft)

# ── FIGURES: все 23 скриншота ──
FIGURES = [
    # Task 1: 1..6.jpg (1 цифра = задание 1)
    ('1.jpg',   'Синхронизация репозиториев командой cl-update --sync-only'),       # 0
    ('2.jpg',   'Полное обновление системы командой cl-update — установка пакетов'), # 1
    ('3.jpg',   'Завершение обновления — удаление устаревших пакетов'),              # 2
    ('4.jpg',   'Синхронизация репозиториев через emerge --sync'),                   # 3
    ('5.jpg',   'Проверка обновлений через emerge --pretend'),                       # 4
    ('6.jpg',   'Установка обновлений через emerge --update --deep --newuse @world'),# 5
    # Task 2: 11..15.jpg (2 цифры = задание 2)
    ('11.jpg',  'Поиск пакета htop командой eix'),                                   # 6
    ('12.jpg',  'Установка пакета htop через emerge -a'),                            # 7
    ('13.jpg',  'Утилита htop — интерактивный монитор процессов'),                   # 8
    ('14.jpg',  'Удаление пакета htop командой emerge --unmerge'),                   # 9
    ('15.jpg',  'Очистка от осиротевших зависимостей emerge --depclean'),            # 10
    # Task 3: 111..115.jpg (3 цифры = задание 3)
    ('111.jpg', 'Просмотр списка запущенных служб (rc-status)'),                     # 11
    ('112.jpg', 'Проверка статуса службы sshd'),                                     # 12
    ('113.jpg', 'Остановка и запуск службы sshd'),                                   # 13
    ('114.jpg', 'Управление автозапуском службы sshd (rc-update)'),                  # 14
    ('115.jpg', 'Проверка изменений автозапуска'),                                    # 15
    # Task 4: 1111..1117.jpg (4 цифры = задание 4)
    ('1111.jpg','Создание пользователя adminuser'),                                   # 16
    ('1112.jpg','Создание пользователя simpleuser'),                                  # 17
    ('1113.jpg','Назначение паролей пользователям'),                                  # 18
    ('1114.jpg','Проверка групп пользователей (id)'),                                 # 19
    ('1115.jpg','simpleuser: попытка sudo — доступ запрещён'),                        # 20
    ('1116.jpg','adminuser: выполнение sudo — доступ разрешён'),                      # 21
    ('1117.jpg','Завершение настройки учётных записей'),                              # 22
]

TOC = [
    (0, 'ВВЕДЕНИЕ', 'intro'),
    (0, '1 Обновление операционной системы', 's1'),
    (1, '1.1 Синхронизация репозиториев', 's1_1'),
    (1, '1.2 Полное обновление системы', 's1_2'),
    (1, '1.3 Проверка обновлений через emerge', 's1_3'),
    (0, '2 Управление программным обеспечением', 's2'),
    (1, '2.1 Поиск и установка пакета', 's2_1'),
    (1, '2.2 Проверка работоспособности', 's2_2'),
    (1, '2.3 Удаление пакета и очистка зависимостей', 's2_3'),
    (0, '3 Управление службами', 's3'),
    (1, '3.1 Просмотр состояния служб', 's3_1'),
    (1, '3.2 Остановка и запуск службы', 's3_2'),
    (1, '3.3 Управление автозапуском', 's3_3'),
    (0, '4 Управление учётными записями', 's4'),
    (1, '4.1 Создание пользователей', 's4_1'),
    (1, '4.2 Назначение паролей', 's4_2'),
    (1, '4.3 Проверка разграничения прав', 's4_3'),
    (0, 'КОНТРОЛЬНЫЕ ВОПРОСЫ', 'questions'),
    (0, 'ЗАКЛЮЧЕНИЕ', 'conclusion'),
]

TITLE = [
    ('Колледж Научно-Технологического Университета «Сириус»', 'TitlePage'),
    ('                                                          ', 'TitleUnderline'),
    ('', 'TitlePage'), ('', 'TitlePage'), ('', 'TitlePage'),
    ('ЛАБОРАТОРНАЯ РАБОТА № 5', 'TitleBold16'),
    ('по дисциплине «Операционные системы и среды»', 'TitlePage'),
    ('на тему «Администрирование Calculate Linux Desktop»', 'TitleBold'),
    ('', 'TitlePage'), ('', 'TitlePage'),
    ('Выполнил: студент группы', 'TitlePage'),
    ('К0709-25/2 Лернер Владислав', 'TitlePage'),
    ('Принял:', 'TitlePage'),
    ('Серов Валерий Александрович', 'TitlePage'),
    ('_____________________', 'TitleUnderline'),
    ('', 'TitlePage'), ('', 'TitlePage'),
    ('IT-Колледж «Сириус»', 'TitlePage'),
    ('2026', 'TitlePage'),
]

# ── Содержание отчёта ──
# (тип, данные...)
# h1c=структурный, h1=заголовок1, h2=заголовок2, p=абзац, code=код, toc=содержание, fig=рисунок
CONTENT = [
    # РЕФЕРАТ
    ('h1c', 'РЕФЕРАТ', 'ref_abstract'),
    ('p', 'Пояснительная записка содержит сведения об администрировании операционной системы Calculate Linux Desktop: обновление системы, управление пакетами, работа со службами и управление учётными записями пользователей.'),
    ('p', 'Ключевые слова: CALCULATE LINUX, PORTAGE, CL-UPDATE, EMERGE, OPENRC, ОБНОВЛЕНИЕ СИСТЕМЫ, УПРАВЛЕНИЕ ПАКЕТАМИ, УПРАВЛЕНИЕ СЛУЖБАМИ, УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ.'),
    ('p', 'Объектом исследования является ОС Calculate Linux Desktop. Цель — приобретение практических навыков по управлению обновлениями, установке и удалению ПО, работе со службами и учётными записями.'),

    # СОДЕРЖАНИЕ
    ('h1c', 'СОДЕРЖАНИЕ', 'toc'),
    ('toc',),

    # ВВЕДЕНИЕ
    ('h1c', 'ВВЕДЕНИЕ', 'intro'),
    ('p', 'Calculate Linux Desktop — российский дистрибутив на базе Gentoo, использующий систему Portage для управления пакетами. Дистрибутив применяет модель rolling-release, что позволяет обновлять систему без переустановки.'),
    ('p', 'Система инициализации — OpenRC. Для администрирования используется набор утилит (cl-update, cl-useradd и др.), а также поддерживается установка бинарных пакетов из репозиториев Calculate и Gentoo.'),
    ('p', 'Цель — приобретение практических навыков по управлению обновлениями ОС, установке и удалению ПО, работе со службами и учётными записями в Calculate Linux Desktop.'),
    ('p', 'Задачи: 1) обновить ОС и зафиксировать обновлённые компоненты; 2) установить и удалить htop; 3) управлять службами через OpenRC; 4) создать пользователей с разными правами.'),

    # ЗАДАНИЕ 1
    ('h1', '1 Обновление операционной системы', 's1'),
    ('p', 'Обновление системы выполняется с помощью утилиты cl-update, которая автоматически синхронизирует репозитории, рассчитывает зависимости, устанавливает обновления и удаляет устаревшие зависимости.'),

    ('h2', '1.1 Синхронизация репозиториев', 's1_1'),
    ('p', 'Первым шагом выполнена синхронизация репозиториев:'),
    ('code', 'sudo cl-update --sync-only'),
    ('p', 'Команда обновила дерево Portage и оверлеи. Сервер обновлений — http://mirror.mephi.ru/calculate.'),
    ('fig', 0),

    ('h2', '1.2 Полное обновление системы', 's1_2'),
    ('p', 'Полное обновление системы:'),
    ('code', 'sudo cl-update'),
    ('p', 'В ходе обновления установлено 206 пакетов (676 895 KiB). Обновлены: glibc, GCC, OpenSSL, libxml2, Qt6, calculate-utils, portage, LibreOffice, GIMP и другие.'),
    ('fig', 1),
    ('fig', 2),

    ('h2', '1.3 Проверка обновлений через emerge', 's1_3'),
    ('p', 'Дополнительная проверка средствами Portage:'),
    ('code', 'sudo emerge --sync'),
    ('fig', 3),
    ('code', 'sudo emerge --pretend --update --deep --newuse @world'),
    ('p', 'Утилита сообщила об отсутствии устаревших пакетов.'),
    ('fig', 4),
    ('code', 'sudo emerge --update --deep --newuse @world'),
    ('fig', 5),

    # ЗАДАНИЕ 2
    ('h1', '2 Управление программным обеспечением', 's2'),
    ('p', 'Управление пакетами выполняется через Portage с использованием emerge.'),

    ('h2', '2.1 Поиск и установка пакета', 's2_1'),
    ('p', 'Установлена утилита htop — монитор процессов:'),
    ('code', 'eix htop'),
    ('fig', 6),
    ('code', 'sudo emerge -a sys-process/htop'),
    ('p', 'Пакет загружен из бинарного репозитория и установлен.'),
    ('fig', 7),

    ('h2', '2.2 Проверка работоспособности', 's2_2'),
    ('p', 'Работоспособность htop подтверждена запуском в терминале.'),
    ('fig', 8),

    ('h2', '2.3 Удаление пакета и очистка зависимостей', 's2_3'),
    ('code', 'sudo emerge --unmerge sys-process/htop'),
    ('fig', 9),
    ('code', 'sudo emerge --depclean'),
    ('p', 'Утилита depclean не обнаружила пакетов для удаления.'),
    ('fig', 10),

    # ЗАДАНИЕ 3
    ('h1', '3 Управление службами', 's3'),
    ('p', 'В Calculate Linux Desktop используется OpenRC. Выбрана служба sshd.'),

    ('h2', '3.1 Просмотр состояния служб', 's3_1'),
    ('code', 'sudo rc-status\nsudo rc-service sshd status'),
    ('p', 'Среди запущенных: NetworkManager, alsasound, bluetooth, cupsd, sshd и др.'),
    ('fig', 11),
    ('fig', 12),

    ('h2', '3.2 Остановка и запуск службы', 's3_2'),
    ('code', 'sudo rc-service sshd stop\nsudo rc-service sshd start'),
    ('p', 'Обе команды выполнены успешно.'),
    ('fig', 13),

    ('h2', '3.3 Управление автозапуском', 's3_3'),
    ('code', 'sudo rc-update add sshd default\nsudo rc-update del sshd default'),
    ('p', 'Служба sshd уже была в runlevel default. Удаление выполнено.'),
    ('fig', 14),
    ('fig', 15),

    # ЗАДАНИЕ 4
    ('h1', '4 Управление учётными записями пользователей', 's4'),
    ('p', 'Права администратора — членство в группе wheel.'),

    ('h2', '4.1 Создание пользователей', 's4_1'),
    ('p', 'Создан adminuser (с правами):'),
    ('code', 'useradd --create-home --groups users,wheel,audio,cdrom,video,usb,plugdev,games,scanner,lp,lpadmin adminuser'),
    ('p', 'Создан simpleuser (без прав):'),
    ('code', 'useradd --create-home --groups users,audio,cdrom,video,usb,plugdev,games,scanner,lp,lpadmin simpleuser'),
    ('p', 'Ключевое различие: adminuser в группе wheel, simpleuser — нет.'),
    ('fig', 16),
    ('fig', 17),

    ('h2', '4.2 Назначение паролей', 's4_2'),
    ('code', 'passwd adminuser\npasswd simpleuser'),
    ('p', 'Пароли установлены.'),
    ('fig', 18),

    ('h2', '4.3 Проверка разграничения прав', 's4_3'),
    ('code', 'id adminuser\nid simpleuser'),
    ('p', 'adminuser в группе wheel, simpleuser — нет.'),
    ('fig', 19),
    ('p', 'simpleuser получил отказ при sudo. adminuser успешно выполнил sudo ls /root.'),
    ('fig', 20),
    ('fig', 21),
    ('fig', 22),

    # КОНТРОЛЬНЫЕ ВОПРОСЫ
    ('h1c', 'КОНТРОЛЬНЫЕ ВОПРОСЫ', 'questions'),
    ('h2', '1. Чем отличается cl-update --sync-only от cl-update?', None),
    ('p', 'cl-update --sync-only — только синхронизация. cl-update — полный цикл: синхронизация, расчёт зависимостей, установка, очистка.'),
    ('h2', '2. Как найти пакет без точного имени?', None),
    ('p', 'eix <слово> или emerge --search <слово>. eix быстрее благодаря индексной базе.'),
    ('h2', '3. Какие группы для звука и USB?', None),
    ('p', 'audio — звук, usb и plugdev — USB, video — видео.'),
    ('h2', '4. Как добавить службу в автозапуск OpenRC?', None),
    ('p', 'sudo rc-update add <служба> default. Для удаления — sudo rc-update del.'),
    ('h2', '5. Что с домашней директорией при userdel без флагов?', None),
    ('p', 'Не удаляется. Для удаления — userdel -r.'),
    ('h2', '6. Какая группа даёт права администратора?', None),
    ('p', 'Группа wheel. cl-useradd добавляет автоматически, useradd — вручную.'),
    ('h2', '7. Разница emerge --unmerge и --depclean?', None),
    ('p', '--unmerge удаляет указанный пакет. --depclean удаляет осиротевшие зависимости.'),
    ('h2', '8. cl-update vs emerge -u -D -N @world?', None),
    ('p', 'cl-update — автоматизированный цикл. emerge — только установка обновлений.'),
    ('h2', '9. Что делает флаг --pretend (-p)?', None),
    ('p', 'Показывает план без выполнения. Позволяет оценить изменения.'),

    # ЗАКЛЮЧЕНИЕ
    ('h1c', 'ЗАКЛЮЧЕНИЕ', 'conclusion'),
    ('p', 'В ходе работы получены навыки администрирования Calculate Linux Desktop.'),
    ('p', 'Выполнено обновление системы (206 пакетов), установлена и удалена утилита htop, изучено управление службами через OpenRC.'),
    ('p', 'Созданы два пользователя: adminuser (wheel) и simpleuser (без wheel). Разграничение прав подтверждено.'),
    ('p', 'Цель достигнута, все задачи выполнены.'),
]

def main():
    tmp_dir = Path(tempfile.mkdtemp(prefix='odt_'))
    try:
        doc = OpenDocumentText()
        doc.meta.addElement(dc.Title(text='Отчёт ЛР №5. Администрирование Calculate Linux'))
        doc.meta.addElement(dc.Creator(text='Лернер Владислав'))
        setup_styles(doc)
        setup_layout(doc)
        body = doc.text

        # Титульный лист
        for text, style in TITLE:
            p = P(stylename=style)
            if text: p.addText(text)
            body.addElement(p)

        # Основное содержание
        fnum = 0
        for block in CONTENT:
            kind = block[0]
            if kind == 'h1c':
                add_p_bm(body, block[1], 'StructHeading', block[2] if len(block) > 2 else None)
            elif kind == 'h1':
                add_p_bm(body, block[1], 'Heading1', block[2] if len(block) > 2 else None)
            elif kind == 'h2':
                add_p_bm(body, block[1], 'Heading2', block[2] if len(block) > 2 else None)
            elif kind == 'p':
                add_p(body, block[1])
            elif kind == 'code':
                add_code(body, block[1])
            elif kind == 'toc':
                for level, text, bm in TOC:
                    p = P(stylename=f'TOC{level+1}')
                    p.addElement(BookmarkStart(name=f'toc_{bm}'))
                    p.addElement(BookmarkEnd(name=f'toc_{bm}'))
                    a = TextA(href=f'#{bm}')
                    a.addText(('    ' * level) + text)
                    p.addElement(a)
                    body.addElement(p)
            elif kind == 'fig':
                idx = block[1]
                fname, caption = FIGURES[idx]
                fnum += 1
                img_path = SHOTS / fname
                if img_path.exists():
                    add_fig(doc, body, img_path, fnum, caption, tmp_dir)
                else:
                    print(f'[WARN] {fname} не найден')
                # Подпись с перекрёстной ссылкой
                p_cap = P(stylename='CenterNoIndent')
                a_fig = TextA(href=f'#fig{fnum}')
                sp_fig = Span(stylename='RefGray')
                sp_fig.addText(f'Рисунок {fnum}')
                a_fig.addElement(sp_fig)
                p_cap.addElement(a_fig)
                p_cap.addText(f' — {caption}')
                body.addElement(p_cap)

        REPORT.mkdir(parents=True, exist_ok=True)
        out = REPORT / 'Отчёт_ЛР5_Администрирование_Calculate_Linux.odt'
        doc.save(str(out))
        print(f'Сохранено: {out}')
        print(f'Рисунков: {fnum}')
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

if __name__ == '__main__':
    main()

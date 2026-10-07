#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генератор отчёта по ЛР №5 — Администрирование Calculate Linux Desktop.

Формат: .odt, оформление по ГОСТ 7.32-2017.
Шрифт: Times New Roman 14 pt, полуторный интервал, выравнивание по ширине,
       абзацный отступ 1.25 см, поля 30/15/20/20 мм.
"""

import os, sys, tempfile, shutil
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

BASE   = Path(__file__).resolve().parent
SHOTS  = BASE / 'screenshots'
REPORT = BASE / 'report'

from odf.opendocument import OpenDocumentText
from odf.text import P, H, Span, BookmarkStart, BookmarkEnd, PageNumber
from odf.style import (
    Style, TextProperties, ParagraphProperties, GraphicProperties,
    PageLayout, PageLayoutProperties, MasterPage, Footer,
)
from odf.draw import Frame, Image as ODFImage
from odf import dc
import odf.table as T

# Атрибуты в odfpy задаются БЕЗ дефисов в нижнем регистре.
# fo:font-family  → 'fontfamily'
# fo:font-size    → 'fontsize'
# style:font-size-asian → 'fontsizeasian'

def _tp(**kw):
    """TextProperties с удобным именованием."""
    return TextProperties(**kw)

def _pp(**kw):
    """ParagraphProperties с удобным именованием."""
    return ParagraphProperties(**kw)

def _gp(**kw):
    """GraphicProperties."""
    return GraphicProperties(**kw)


# ── НАСТРОЙКА СТИЛЕЙ ──────────────────────────────────────────────────────────

def setup_styles(doc):
    # ── Базовый стиль Standard ──
    std = Style(name='Standard', family='text')
    std.addElement(_pp(
        lineheight='150%', textalign='justify',
        textindent='1.25cm', marginbottom='0cm', margintop='0cm',
    ))
    std.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt',
        fontsizeasian='14pt', fontsizecomplex='14pt',
    ))

    # ── Normal (paragraph) ──
    normal = Style(name='Normal', family='paragraph', parentstylename='Standard')
    normal.addElement(_pp(
        lineheight='150%', textalign='justify',
        textindent='1.25cm', marginbottom='0cm', margintop='0cm',
    ))
    normal.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
    ))
    doc.styles.addElement(normal)

    # ── Heading 1 ──
    h1 = Style(name='Heading1', family='paragraph', parentstylename='Heading 1')
    h1.addElement(_pp(
        lineheight='150%', textalign='left', textindent='1.25cm',
        marginbottom='0.35cm', keepwithnext='true',
    ))
    h1.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
        fontweight='bold',
    ))
    doc.styles.addElement(h1)

    # ── Heading 2 ──
    h2 = Style(name='Heading2', family='paragraph', parentstylename='Heading 2')
    h2.addElement(_pp(
        lineheight='150%', textalign='left', textindent='1.25cm',
        marginbottom='0.35cm', keepwithnext='true',
    ))
    h2.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
        fontweight='bold',
    ))
    doc.styles.addElement(h2)

    # ── StructHeading (центрированный, с разрывом страницы) ──
    sh = Style(name='StructHeading', family='paragraph', parentstylename='Standard')
    sh.addElement(_pp(
        lineheight='150%', textalign='center', textindent='0cm',
        breakbefore='page', marginbottom='0.35cm', keepwithnext='true',
    ))
    sh.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
        fontweight='bold', texttransform='uppercase',
    ))
    doc.styles.addElement(sh)

    # ── CenterNoIndent ──
    cn = Style(name='CenterNoIndent', family='paragraph', parentstylename='Standard')
    cn.addElement(_pp(
        lineheight='150%', textalign='center', textindent='0cm',
        marginbottom='0cm', margintop='0cm',
    ))
    cn.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
    ))
    doc.styles.addElement(cn)

    # ── TitlePage ──
    tp = Style(name='TitlePage', family='paragraph', parentstylename='Standard')
    tp.addElement(_pp(
        lineheight='115%', textalign='center', textindent='0cm',
        marginbottom='0.2cm', margintop='0cm',
    ))
    tp.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
    ))
    doc.styles.addElement(tp)

    # ── TitleBold ──
    tb = Style(name='TitleBold', family='paragraph', parentstylename='TitlePage')
    tb.addElement(_tp(fontweight='bold'))
    doc.styles.addElement(tb)

    # ── TitlePage16 (для заголовка "ОТЧЁТ") ──
    tp16 = Style(name='TitlePage16', family='paragraph', parentstylename='TitlePage')
    tp16.addElement(_tp(fontsize='16pt', fontsizeasian='16pt'))
    doc.styles.addElement(tp16)

    # ── TitleBold16 ──
    tb16 = Style(name='TitleBold16', family='paragraph', parentstylename='TitlePage')
    tb16.addElement(_tp(fontsize='16pt', fontsizeasian='16pt', fontweight='bold'))
    doc.styles.addElement(tb16)

    # ── CodeBlock ──
    cb = Style(name='CodeBlock', family='paragraph', parentstylename='Standard')
    cb.addElement(_pp(
        lineheight='120%', textalign='left', textindent='0cm',
        backgroundcolor='#f0f0f0',
        marginleft='0.5cm', marginright='0.5cm',
        margintop='0.2cm', marginbottom='0.2cm',
    ))
    cb.addElement(_tp(
        fontfamily='Courier New', fontsize='11pt', fontsizeasian='11pt',
    ))
    doc.styles.addElement(cb)

    # ── TOC1 ──
    toc1 = Style(name='TOC1', family='paragraph', parentstylename='Standard')
    toc1.addElement(_pp(
        lineheight='150%', textalign='left', textindent='0cm',
        marginleft='0cm',
    ))
    toc1.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
        fontweight='bold',
    ))
    doc.styles.addElement(toc1)

    # ── TOC2 ──
    toc2 = Style(name='TOC2', family='paragraph', parentstylename='Standard')
    toc2.addElement(_pp(
        lineheight='150%', textalign='left', textindent='0cm',
        marginleft='0.75cm',
    ))
    toc2.addElement(_tp(
        fontfamily='Times New Roman', fontsize='14pt', fontsizeasian='14pt',
    ))
    doc.styles.addElement(toc2)

    # ── Bold / Itonic ──
    bold_s = Style(name='Bold', family='text')
    bold_s.addElement(_tp(fontweight='bold'))
    doc.styles.addElement(bold_s)

    italic_s = Style(name='Italic', family='text')
    italic_s.addElement(_tp(fontstyle='italic'))
    doc.styles.addElement(italic_s)

    # ── TableText ──
    tt = Style(name='TableText', family='paragraph', parentstylename='Standard')
    tt.addElement(_pp(
        lineheight='115%', textalign='left', textindent='0cm',
    ))
    tt.addElement(_tp(fontsize='12pt', fontsizeasian='12pt'))
    doc.styles.addElement(tt)


def setup_page_layout(doc):
    # Основная разметка страницы
    pl = PageLayout(name='MPl')
    pl.addElement(PageLayoutProperties(
        pagewidth='21cm', pageheight='29.7cm',
        marginleft='3cm', marginright='1.5cm',
        margintop='2cm', marginbottom='2cm',
        printorientation='portrait',
    ))
    doc.automaticstyles.addElement(pl)

    mp = MasterPage(name='Standard', pagelayoutname=pl)
    doc.masterstyles.addElement(mp)

    # Футер с номером страницы
    ft = Footer()
    ftp = P(stylename='CenterNoIndent')
    ftp.addElement(PageNumber())
    ft.addElement(ftp)
    mp.addElement(ft)


# ── РАБОТА С ИЗОБРАЖЕНИЯМИ ────────────────────────────────────────────────────

def make_image_with_caption(img_path: Path, caption: str, tmp_dir: Path) -> Path:
    """Накладывает подпись НА изображение (снизу, на полосе)."""
    img = Image.open(img_path).convert('RGB')
    w, h = img.size
    font_size = max(14, int(h * 0.032))
    try:
        font = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", font_size)
    except Exception:
        try:
            font = ImageFont.truetype(
                "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", font_size)
        except Exception:
            font = ImageFont.load_default()

    draw_tmp = ImageDraw.Draw(img)
    max_text_w = int(w * 0.95)
    words = caption.split()
    lines, line = [], ''
    for word in words:
        test = f'{line} {word}'.strip()
        bbox = draw_tmp.textbbox((0, 0), test, font=font)
        if bbox[2] - bbox[0] > max_text_w and line:
            lines.append(line)
            line = word
        else:
            line = test
    if line:
        lines.append(line)

    text_h = sum(
        draw_tmp.textbbox((0, 0), ln, font=font)[3] -
        draw_tmp.textbbox((0, 0), ln, font=font)[1] for ln in lines
    ) + len(lines) * 4
    bar_h = text_h + 24

    canvas = Image.new('RGB', (w, h + bar_h), (255, 255, 255))
    canvas.paste(img, (0, 0))
    draw = ImageDraw.Draw(canvas)
    # Серый фон полосы
    draw.rectangle([0, h, w, h + bar_h], fill=(235, 235, 235))
    y = h + 12
    for ln in lines:
        bbox = draw.textbbox((0, 0), ln, font=font)
        tw = bbox[2] - bbox[0]
        x = (w - tw) // 2
        draw.text((x, y), ln, fill=(40, 40, 40), font=font)
        y += (bbox[3] - bbox[1]) + 4

    out = tmp_dir / img_path.name
    canvas.save(str(out), quality=90)
    return out


# ── ПОСТРОЕНИЕ ДОКУМЕНТА ──────────────────────────────────────────────────────

TITLE_LINES = [
    ('Колледж Научно-Технологического Университета «Сириус»',  'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('ОТЧЁТ',                                                   'TitleBold16'),
    ('о выполнении лабораторной работы № 5',                    'TitlePage'),
    ('по дисциплине',                                           'TitlePage'),
    ('«Операционные системы и среды»',                         'TitlePage'),
    ('',                                                        'TitlePage'),
    ('Тема: «Администрирование операционной системы',           'TitleBold'),
    ('Calculate Linux Desktop»',                                'TitleBold'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('Выполнил: студент группы К0709-25/2',                     'TitlePage'),
    ('Лернер Владислав',                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('Принял:',                                                 'TitlePage'),
    ('Серов Валерий Александрович',                             'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('',                                                        'TitlePage'),
    ('IT-Колледж «Сириус»',                                    'TitlePage'),
    ('2026',                                                    'TitlePage'),
]


def add_p(body, text, stylename='Normal'):
    p = P(stylename=stylename)
    if text:
        p.addText(text)
    body.addElement(p)
    return p


def add_code(body, text):
    for line in text.split('\n'):
        add_p(body, line, 'CodeBlock')


def add_figure(doc, body, img_path, fig_num, caption, tmp_dir, width_cm=15):
    """Вставка рисунка с подписью ВНУТРИ изображения + закладка для ссылок."""
    out = make_image_with_caption(img_path, caption, tmp_dir)
    img = Image.open(out)
    w, h = img.size
    h_cm = width_cm * h / w

    bm_name = f'fig{fig_num}'
    bm_start = BookmarkStart(name=bm_name)
    bm_end = BookmarkEnd(name=bm_name)

    p = P(stylename='CenterNoIndent')
    p.addElement(bm_start)
    p.addElement(bm_end)

    frame = Frame(
        width=f'{width_cm}cm',
        height=f'{h_cm}cm',
        anchortype='paragraph',
    )
    href = doc.addPicture(str(out))
    frame.addElement(ODFImage(href=href))
    p.addElement(frame)
    body.addElement(p)
    return bm_name


CONTENT = [
    # ── РЕФЕРАТ ──
    ('h1c', 'РЕФЕРАТ'),
    ('p',   'Пояснительная записка содержит сведения об администрировании '
            'операционной системы Calculate Linux Desktop: обновление системы, '
            'управление пакетами, работа со службами и управление учётными '
            'записями пользователей.'),
    ('p',   'Ключевые слова: CALCULATE LINUX, PORTAGE, CL-UPDATE, EMERGE, OPENRC, '
            'ОБНОВЛЕНИЕ СИСТЕМЫ, УПРАВЛЕНИЕ ПАКЕТАМИ, УПРАВЛЕНИЕ СЛУЖБАМИ, '
            'УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ.'),
    ('p',   'Объектом исследования является операционная система Calculate Linux Desktop. '
            'Целью работы является приобретение практических навыков по управлению '
            'обновлениями операционной системы, установке и удалению программного '
            'обеспечения, работе со службами и учётными записями пользователей.'),

    # ── СОДЕРЖАНИЕ ──
    ('h1c', 'СОДЕРЖАНИЕ'),
    ('toc', None),

    # ── ВВЕДЕНИЕ ──
    ('h1c', 'ВВЕДЕНИЕ'),
    ('p',   'Calculate Linux Desktop — российский дистрибутив на базе Gentoo, '
            'использующий систему Portage для управления пакетами. Дистрибутив '
            'применяет модель «滚动-release» (rolling-release), что позволяет '
            'обновлять систему без необходимости переустановки.'),
    ('p',   'Система инициализации — OpenRC. Для администрирования используется '
            'собственный набор утилит (cl-update, cl-useradd и др.), а также '
            'поддерживается установка бинарных пакетов из репозиториев Calculate '
            'и Gentoo.'),
    ('p',   'Целью данной лабораторной работы является приобретение практических '
            'навыков по управлению обновлениями операционной системы, установке '
            'и удалению программного обеспечения, работе со службами и учётными '
            'записями пользователей в ОС Calculate Linux Desktop.'),
    ('p',   'Для достижения поставленной цели необходимо решить следующие задачи:'),
    ('p',   '    1) выполнить обновление операционной системы и зафиксировать '
            'перечень обновлённых компонентов;'),
    ('p',   '    2) установить и удалить консольную утилиту htop, очистить '
            'систему от осиротевших зависимостей;'),
    ('p',   '    3) выполнить управление службами: остановку, запуск, добавление '
            'и удаление из автозагрузки;'),
    ('p',   '    4) создать двух пользователей с разными правами и проверить '
            'разграничение доступа.'),

    # ── ЗАДАНИЕ 1 ──
    ('h1', '1 Обновление операционной системы'),
    ('p',  'Обновление системы в Calculate Linux Desktop выполняется с помощью '
           'специализированной утилиты cl-update, которая автоматически '
           'синхронизирует репозитории, рассчитывает зависимости, устанавливает '
           'обновления и удаляет устаревшие зависимости.'),
    ('h2', '1.1 Синхронизация репозиториев'),
    ('p',  'Первым шагом выполнена синхронизация репозиториев без установки '
           'обновлений с помощью команды:'),
    ('code', 'sudo cl-update --sync-only'),
    ('p',  'Команда обновила дерево Portage и оверлеи, не изменяя установленные '
           'пакеты. Сервер обновлений — http://mirror.mephi.ru/calculate.'),
    ('fig', '1.jpg', 'Синхронизация репозиториев командой cl-update --sync-only'),

    ('h2', '1.2 Полное обновление системы'),
    ('p',  'Далее выполнено полное обновление системы командой:'),
    ('code', 'sudo cl-update'),
    ('p',  'Утилита cl-update автоматически: синхронизировала Portage и оверлеи, '
           'рассчитала зависимости, установила обновлённые пакеты и удалила '
           'устаревшие зависимости.'),
    ('p',  'В ходе обновления было установлено 206 пакетов (676 895 KiB загружено). '
           'Среди обновлённых компонентов: ядро glibc, компилятор GCC, библиотеки '
           'OpenSSL, libxml2, Qt6, утилиты calculate-utils, portage, а также '
           'пакеты LibreOffice, GIMP, ImageMagick и другие.'),
    ('p',  'После установки пакетов система предложила удалить 4 неиспользуемых '
           'пакета (включая устаревший GCC 15.3.0), после чего обновление было '
           'завершено.'),
    ('fig', '2.jpg', 'Обновление системы командой cl-update — установка пакетов'),

    ('h2', '1.3 Проверка обновлений через emerge'),
    ('p',  'Дополнительно выполнена проверка обновлений с помощью стандартных '
           'средств Portage:'),
    ('code', 'sudo emerge --sync'),
    ('p',  'Синхронизация выполнена успешно для трёх репозиториев: gentoo, '
           'calculate, distros.'),
    ('p',  'Проверка наличия обновлений без их установки:'),
    ('code', 'sudo emerge --pretend --update --deep --newuse @world'),
    ('p',  'Утилита сообщила об отсутствии устаревших пакетов на момент проверки.'),
    ('fig', '3.jpg', 'Проверка обновлений через emerge'),

    # ── ЗАДАНИЕ 2 ──
    ('h1', '2 Управление программным обеспечением'),
    ('p',  'В Calculate Linux Desktop управление пакетами выполняется через '
           'систему Portage с использованием утилиты emerge.'),
    ('h2', '2.1 Поиск и установка пакета'),
    ('p',  'В качестве примера установлена консольная утилита htop — интерактивный '
           'монитор процессов. Поиск пакета:'),
    ('code', 'eix htop'),
    ('p',  'Утилита eix нашла пакет sys-process/htop версии 3.5.3. '
           'Установка пакета:'),
    ('code', 'sudo emerge -a sys-process/htop'),
    ('p',  'Пакет загружен из бинарного репозитория и успешно установлен.'),
    ('fig', '4.jpg', 'Поиск и установка пакета htop'),

    ('h2', '2.2 Проверка работоспособности'),
    ('p',  'Работоспособность утилиты htop подтверждена её запуском в терминале.'),
    ('fig', '5.jpg', 'Утилита htop — интерактивный монитор процессов'),

    ('h2', '2.3 Удаление пакета и очистка зависимостей'),
    ('p',  'Удаление пакета:'),
    ('code', 'sudo emerge --unmerge sys-process/htop'),
    ('p',  'Удаление осиротевших зависимостей:'),
    ('code', 'sudo emerge --depclean'),
    ('p',  'Утилита depclean не обнаружила пакетов для удаления.'),
    ('fig', '6.jpg', 'Удаление пакета htop и очистка от осиротевших зависимостей'),

    # ── ЗАДАНИЕ 3 ──
    ('h1', '3 Управление службами'),
    ('p',  'В Calculate Linux Desktop используется система инициализации OpenRC. '
           'В качестве примера выбрана служба sshd.'),
    ('h2', '3.1 Просмотр состояния служб'),
    ('code', 'sudo rc-status'),
    ('p',  'Среди запущенных служб: NetworkManager, alsasound, bluetooth, chronyd, '
           'cupsd, sshd и другие.'),
    ('code', 'sudo rc-service sshd status'),
    ('p',  'Служба sshd была в состоянии started.'),
    ('fig', '11.jpg', 'Просмотр списка служб и проверка статуса sshd'),

    ('h2', '3.2 Остановка и запуск службы'),
    ('code', 'sudo rc-service sshd stop\nsudo rc-service sshd start'),
    ('p',  'Обе команды выполнены успешно.'),
    ('fig', '12.jpg', 'Остановка и запуск службы sshd'),

    ('h2', '3.3 Управление автозапуском'),
    ('code', 'sudo rc-update add sshd default'),
    ('p',  'Служба sshd уже была в runlevel default.'),
    ('code', 'sudo rc-update del sshd default'),
    ('p',  'Удаление из автозагрузки выполнено успешно.'),
    ('fig', '13.jpg', 'Управление автозапуском службы sshd'),

    # ── ЗАДАНИЕ 4 ──
    ('h1', '4 Управление учётными записями пользователей'),
    ('p',  'Права администратора в Calculate Linux Desktop предоставляются '
           'членством в группе wheel.'),
    ('h2', '4.1 Создание пользователей'),
    ('p',  'Создан пользователь с правами администратора:'),
    ('code', 'useradd --create-home --groups users,wheel,audio,cdrom,video,'
             'usb,plugdev,games,scanner,lp,lpadmin adminuser'),
    ('p',  'Создан пользователь без прав администратора:'),
    ('code', 'useradd --create-home --groups users,audio,cdrom,video,usb,'
             'plugdev,games,scanner,lp,lpadmin simpleuser'),
    ('p',  'Ключевое различие: adminuser включён в группу wheel, simpleuser — нет.'),
    ('fig', '111.jpg', 'Создание пользователей adminuser и simpleuser'),

    ('h2', '4.2 Назначение паролей'),
    ('code', 'passwd adminuser\npasswd simpleuser'),
    ('p',  'Пароли успешно установлены обоим пользователям.'),
    ('fig', '112.jpg', 'Назначение паролей пользователям'),

    ('h2', '4.3 Проверка разграничения прав'),
    ('code', 'id adminuser\nid simpleuser'),
    ('p',  'adminuser состоит в группе wheel, simpleuser — нет.'),
    ('fig', '113.jpg', 'Проверка принадлежности пользователей к группам'),

    ('p',  'Пользователь simpleuser получил отказ при попытке использовать sudo: '
           '«simpleuser отсутствует в файле sudoers».'),
    ('p',  'Пользователь adminuser успешно выполнил sudo ls /root, подтвердив '
           'наличие прав администратора.'),
    ('fig', '114.jpg', 'Попытка simpleuser выполнить sudo — доступ запрещён'),
    ('fig', '115.jpg', 'Выполнение sudo пользователем adminuser — доступ разрешён'),

    # ── КОНТРОЛЬНЫЕ ВОПРОСЫ ──
    ('h1c', 'КОНТРОЛЬНЫЕ ВОПРОСЫ'),

    ('h2', '1. Чем отличается cl-update --sync-only от cl-update?'),
    ('p',  'cl-update --sync-only выполняет только синхронизацию дерева Portage '
           'и оверлеев, не устанавливая обновлений. cl-update выполняет полный '
           'цикл: синхронизацию, расчёт зависимостей, установку обновлений и '
           'удаление устаревших зависимостей.'),

    ('h2', '2. Как найти пакет, если не знаете его точного имени?'),
    ('p',  'Используется утилита eix (eix <слово>) или emerge --search <слово>. '
           'Утилита eix быстрее благодаря предварительно построенной индексной базе.'),

    ('h2', '3. Какие группы нужны для доступа к звуковой карте и USB-накопителям?'),
    ('p',  'Для звуковой карты — группа audio. Для USB-накопителей — usb и plugdev. '
           'Для видеоустройств — video.'),

    ('h2', '4. Как в OpenRC добавить службу в автозапуск?'),
    ('p',  'Команда: sudo rc-update add <имя_службы> default. Для удаления: '
           'sudo rc-update del <имя_службы> default.'),

    ('h2', '5. Что произойдёт с домашней директорией при userdel без флагов?'),
    ('p',  'Домашняя директория НЕ удаляется. Для удаления нужен флаг -r: '
           'userdel -r <имя_пользователя>.'),

    ('h2', '6. Какая группа предоставляет права администратора в Calculate Linux?'),
    ('p',  'Права администратора предоставляет группа wheel. При создании '
           'пользователя через cl-useradd она добавляется автоматически; '
           'при useradd — вручную через --groups.'),

    ('h2', '7. В чём разница между emerge --unmerge и emerge --depclean?'),
    ('p',  'emerge --unmerge удаляет указанный пакет, не затрагивая зависимости. '
           'emerge --depclean удаляет пакеты, не требующиеся ни одному из '
           'установленных пакетов (осиротевшие зависимости).'),

    ('h2', '8. Чем отличается обновление через cl-update от emerge -u -D -N @world?'),
    ('p',  'cl-update — автоматизированный цикл с выбором зеркала, синхронизацией, '
           'обновлением и очисткой. emerge -u -D -N @world — только установка '
           'обновлений, требует ручного вызова синхронизации и очистки.'),

    ('h2', '9. Что делает флаг --pretend (-p) в emerge?'),
    ('p',  'Выводит план действий без фактического выполнения. Показывает, какие '
           'пакеты будут установлены/обновлены/удалены, что позволяет оценить '
           'изменения до их применения.'),

    # ── ЗАКЛЮЧЕНИЕ ──
    ('h1c', 'ЗАКЛЮЧЕНИЕ'),
    ('p',  'В ходе выполнения лабораторной работы были получены практические навыки '
           'администрирования операционной системы Calculate Linux Desktop.'),
    ('p',  'Было выполнено полное обновление системы с помощью утилиты cl-update: '
           'синхронизированы репозитории, обновлены 206 пакетов, удалены '
           'устаревшие компоненты.'),
    ('p',  'Установлена и проверена работоспособность консольной утилиты htop. '
           'Пакет успешно удалён, система очищена от осиротевших зависимостей.'),
    ('p',  'Изучено управление службами на примере sshd: просмотр состояния, '
           'остановка, запуск, добавление и удаление из автозапуска.'),
    ('p',  'Созданы два пользователя с разными уровнями привилегий: adminuser '
           '(с правами администратора) и simpleuser (без таких прав). '
           'Разграничение доступа подтверждено.'),
    ('p',  'Поставленная цель работы достигнута, все задачи выполнены.'),
]


def main():
    tmp_dir = Path(tempfile.mkdtemp(prefix='odt_build_'))
    try:
        doc = OpenDocumentText()
        doc.meta.addElement(dc.Title(text='Отчёт ЛР №5. Администрирование Calculate Linux Desktop'))
        doc.meta.addElement(dc.Creator(text='Лернер Владислав'))

        setup_styles(doc)
        setup_page_layout(doc)
        body = doc.text

        # ── ТИТУЛЬНЫЙ ЛИСТ ──
        for text, style_name in TITLE_LINES:
            p = P(stylename=style_name)
            if text:
                p.addText(text)
            body.addElement(p)

        # ── СОДЕРЖАНИЕ ──
        fig_num = 0
        toc = [
            (0, 'ВВЕДЕНИЕ'),
            (0, '1 Обновление операционной системы'),
            (1, '1.1 Синхронизация репозиториев'),
            (1, '1.2 Полное обновление системы'),
            (1, '1.3 Проверка обновлений через emerge'),
            (0, '2 Управление программным обеспечением'),
            (1, '2.1 Поиск и установка пакета'),
            (1, '2.2 Проверка работоспособности'),
            (1, '2.3 Удаление пакета и очистка зависимостей'),
            (0, '3 Управление службами'),
            (1, '3.1 Просмотр состояния служб'),
            (1, '3.2 Остановка и запуск службы'),
            (1, '3.3 Управление автозапуском'),
            (0, '4 Управление учётными записями пользователей'),
            (1, '4.1 Создание пользователей'),
            (1, '4.2 Назначение паролей'),
            (1, '4.3 Проверка разграничения прав'),
            (0, 'КОНТРОЛЬНЫЕ ВОПРОСЫ'),
            (0, 'ЗАКЛЮЧЕНИЕ'),
        ]

        for block in CONTENT:
            kind = block[0]
            if kind == 'h1c':
                add_p(body, block[1], 'StructHeading')
            elif kind == 'h1':
                add_p(body, block[1], 'Heading1')
            elif kind == 'h2':
                add_p(body, block[1], 'Heading2')
            elif kind == 'p':
                add_p(body, block[1], 'Normal')
            elif kind == 'code':
                add_code(body, block[1])
            elif kind == 'toc':
                for level, text in toc:
                    add_p(body, ('    ' * level) + text, f'TOC{level+1}')
            elif kind == 'fig':
                fname, caption = block[1], block[2]
                fig_num += 1
                img_path = SHOTS / fname
                if img_path.exists():
                    add_figure(doc, body, img_path, fig_num, caption, tmp_dir)
                else:
                    print(f'[WARN] {fname} не найден')

        REPORT.mkdir(parents=True, exist_ok=True)
        out_path = REPORT / 'Отчёт_ЛР5_Администрирование_Calculate_Linux.odt'
        doc.save(str(out_path))
        print(f'Отчёт сохранён: {out_path}')
        print(f'Рисунков: {fig_num}')
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == '__main__':
    main()
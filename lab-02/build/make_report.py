#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка отчёта по лабораторной работе № 2 в формате ODT.

Документ собран в стилистике файла «Лаба 2.odt» из lab-01/examples:
A4, поля 3/1,5/2/1,27 см, Times New Roman 14 pt, полуторный интервал,
абзацный отступ 1,25 см, нумерация страниц внизу по центру и титульный
лист без номера. Все 63 исходных скриншота из lab-02/screenshots включаются
в отчёт и нумеруются сквозным образом.
"""
from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
EXAMPLE = ROOT / "lab-01" / "examples" / "Лаба 2.odt"
SHOTS = ROOT / "lab-02" / "screenshots"
OUT = ROOT / "lab-02" / "report" / "Отчёт_ЛР2_Основы_работы_в_Linux_GUI_РОСА.odt"

# Ширина области набора шаблона: 21 - 3 - 1,499 = 16,501 см.
IMAGE_W = 16.50
IMAGE_H = 9.28

NS = (
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
    'xmlns:xlink="http://www.w3.org/1999/xlink" '
    'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
    'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:ooo="http://openoffice.org/2004/office" '
    'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" '
    'office:version="1.3"'
)


def esc(s: str) -> str:
    return html.escape(s, quote=False)


def p(text: str = "", style: str = "Pbody", *, outline: int | None = None) -> str:
    attr = f' text:outline-level="{outline}"' if outline else ""
    return f'<text:p text:style-name="{style}"{attr}>{esc(text)}</text:p>'


def styled(text: str, style: str) -> str:
    return f'<text:span text:style-name="{style}">{esc(text)}</text:span>'


def title_p(text: str, style: str = "Pcover", span: str = "Tcover") -> str:
    return f'<text:p text:style-name="{style}">{styled(text, span)}</text:p>'


def page_break() -> str:
    # Пустой абзац с разрывом страницы: так Writer сохраняет ручной разрыв
    # и не добавляет к нему видимый текст.
    return '<text:p text:style-name="Ppage"/>'


def paragraph_with_spans(parts: list[tuple[str, str]], style: str = "Pbody") -> str:
    body = "".join(styled(text, span) if span else esc(text) for text, span in parts)
    return f'<text:p text:style-name="{style}">{body}</text:p>'


def figure(n: int, description: str, image_no: int) -> str:
    fname = f"{image_no}.jpg"
    return (
        page_break()
        + p("Результат выполнения соответствующего этапа задания приведён на скриншоте.", "Pfigintro")
        + '<text:p text:style-name="Pimage">'
        f'<draw:frame draw:style-name="fr3" draw:name="Рисунок{n}" '
        f'text:anchor-type="as-char" svg:width="{IMAGE_W:.2f}cm" '
        f'svg:height="{IMAGE_H:.2f}cm" draw:z-index="0">'
        f'<draw:image xlink:href="Pictures/fig-{image_no:03d}.jpg" '
        'xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" '
        'draw:mime-type="image/jpeg"/>'
        '</draw:frame></text:p>'
        + p(f"Рисунок {n} – {description}", "Pcaption")
    )


def toc_line(text: str, page: int, level: int = 0) -> str:
    style = "Ptoc2" if level else "Ptoc"
    return f'<text:p text:style-name="{style}">{esc(text)}<text:tab/>{page}</text:p>'


def automatic_styles() -> str:
    # В styles.xml шаблона уже находятся базовые стили Standard (WW),
    # «1 уровень», «2 уровень», Footer и графический стиль fr3. Здесь
    # добавлены только небольшие автоматические стили документа.
    return '''<office:automatic-styles>
 <style:style style:name="Pcover" style:family="paragraph" style:parent-style-name="Normal_20__28_Web_29_" style:master-page-name="First_20_Page">
  <style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.423cm" fo:text-align="center" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false"/>
 </style:style>
 <style:style style:name="PcoverGap" style:family="paragraph" style:parent-style-name="Normal_20__28_Web_29_" style:master-page-name="First_20_Page">
  <style:paragraph-properties fo:margin-top="0.85cm" fo:margin-bottom="0.85cm" fo:text-align="center" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false"/>
 </style:style>
 <style:style style:name="Ppage" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:break-before="page" fo:text-indent="0cm" style:auto-text-indent="false"/>
 </style:style>
 <style:style style:name="Pstruct" style:family="paragraph" style:parent-style-name="_31__20_уровень">
  <style:paragraph-properties fo:break-before="page" fo:text-align="center" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false" fo:keep-with-next="always"/>
  <style:text-properties fo:font-name="Times New Roman" fo:font-family="&apos;Times New Roman&apos;" fo:font-size="14pt" fo:font-weight="bold" style:font-size-asian="14pt" style:font-weight-asian="bold"/>
 </style:style>
 <style:style style:name="Psection" style:family="paragraph" style:parent-style-name="_32__20_уровень">
  <style:paragraph-properties fo:text-align="justify" style:justify-single-word="false" fo:keep-with-next="always"/>
  <style:text-properties fo:font-name="Times New Roman" fo:font-family="&apos;Times New Roman&apos;" fo:font-size="14pt" fo:font-weight="bold" style:font-size-asian="14pt" style:font-weight-asian="bold"/>
 </style:style>
 <style:style style:name="Psub" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:text-align="justify" style:justify-single-word="false" fo:keep-with-next="always"/>
  <style:text-properties fo:font-name="Times New Roman" fo:font-family="&apos;Times New Roman&apos;" fo:font-size="14pt" fo:font-weight="bold" style:font-size-asian="14pt" style:font-weight-asian="bold"/>
 </style:style>
 <style:style style:name="Pbody" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:line-height="150%" fo:text-align="justify" style:justify-single-word="false"/>
 </style:style>
 <style:style style:name="Pbody0" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:line-height="150%" fo:text-align="justify" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false"/>
 </style:style>
 <style:style style:name="Plist" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:line-height="150%" fo:text-align="justify" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false" fo:margin-left="0.8cm"/>
 </style:style>
 <style:style style:name="Pfigintro" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:line-height="150%" fo:text-align="justify" style:justify-single-word="false" fo:keep-with-next="always"/>
 </style:style>
 <style:style style:name="Pimage" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:line-height="100%" fo:text-align="center" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false" fo:keep-with-next="always"/>
 </style:style>
 <style:style style:name="Pcaption" style:family="paragraph" style:parent-style-name="Standard_20__28_WW_29_">
  <style:paragraph-properties fo:line-height="150%" fo:text-align="center" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false" fo:margin-top="0.21cm"/>
  <style:text-properties fo:font-size="12pt" fo:font-style="italic" style:font-size-asian="12pt" style:font-style-asian="italic"/>
 </style:style>
 <style:style style:name="Ptoc" style:family="paragraph" style:parent-style-name="Contents_20_1">
  <style:paragraph-properties fo:text-align="start" style:justify-single-word="false" fo:text-indent="0cm"><style:tab-stops><style:tab-stop style:position="16.48cm" style:type="right" style:leader-style="dotted" style:leader-text="."/></style:tab-stops></style:paragraph-properties>
 </style:style>
 <style:style style:name="Ptoc2" style:family="paragraph" style:parent-style-name="Contents_20_2">
  <style:paragraph-properties fo:text-align="start" style:justify-single-word="false" fo:text-indent="0cm"><style:tab-stops><style:tab-stop style:position="16.48cm" style:type="right" style:leader-style="dotted" style:leader-text="."/></style:tab-stops></style:paragraph-properties>
 </style:style>
 <style:style style:name="Tcover" style:family="text">
  <style:text-properties style:font-name="Times New Roman" fo:font-family="&apos;Times New Roman&apos;" fo:font-size="14pt" style:font-size-asian="14pt"/>
 </style:style>
 <style:style style:name="TcoverBold" style:family="text">
  <style:text-properties style:font-name="Times New Roman" fo:font-family="&apos;Times New Roman&apos;" fo:font-size="14pt" fo:font-weight="bold" style:font-size-asian="14pt" style:font-weight-asian="bold"/>
 </style:style>
 <style:style style:name="TcoverBig" style:family="text">
  <style:text-properties style:font-name="Times New Roman" fo:font-family="&apos;Times New Roman&apos;" fo:font-size="16pt" fo:font-weight="bold" style:font-size-asian="16pt" style:font-weight-asian="bold"/>
 </style:style>
 <style:style style:name="TcoverSmall" style:family="text">
  <style:text-properties style:font-name="Times New Roman" fo:font-family="&apos;Times New Roman&apos;" fo:font-size="12pt" fo:font-style="italic" style:font-size-asian="12pt" style:font-style-asian="italic"/>
 </style:style>
 <style:style style:name="Tbold" style:family="text">
  <style:text-properties fo:font-weight="bold" style:font-weight-asian="bold"/>
 </style:style>
 <style:style style:name="Titalic" style:family="text">
  <style:text-properties fo:font-style="italic" style:font-style-asian="italic"/>
 </style:style>
 <style:style style:name="fr3" style:family="graphic" style:parent-style-name="Graphics">
  <style:graphic-properties style:mirror="none" fo:clip="rect(0cm, 0cm, 0cm, 0cm)" draw:luminance="0%" draw:contrast="0%" draw:red="0%" draw:green="0%" draw:blue="0%" draw:gamma="100%" draw:color-inversion="false" draw:image-opacity="100%" draw:color-mode="standard"/>
 </style:style>
</office:automatic-styles>'''


def cover() -> str:
    out = []
    out.append(title_p("Колледж Научно-Технологического Университета Сириус", "Pcover", "TcoverBold"))
    out.append(title_p("_____________________________________________________________", "Pcover"))
    out.append(title_p("ЛАБОРАТОРНАЯ РАБОТА № 2.", "PcoverGap", "TcoverBig"))
    out.append(title_p("по дисциплине «Операционные системы и среды»", "Pcover", "Tcover"))
    out.append(title_p("на тему «Основы работы в Linux с GUI.", "Pcover", "TcoverBold"))
    out.append(title_p("Работа с интерфейсом. Пользовательские настройки ОС РОСА»", "Pcover", "TcoverBold"))
    out.append(title_p("", "PcoverGap"))
    out.append(title_p("Выполнил:", "Pcover", "Tcover"))
    out.append(title_p("Студент группы ______________________________", "Pcover", "Tcover"))
    out.append(title_p("______________________________________________", "Pcover", "Tcover"))
    out.append(title_p("", "PcoverGap"))
    out.append(title_p("Принял:", "Pcover", "Tcover"))
    out.append(title_p("Серов Валерий Александрович", "Pcover", "Tcover"))
    out.append(title_p("_____________________", "Pcover", "Tcover"))
    out.append(title_p("", "PcoverGap"))
    out.append(title_p("IT-Колледж «Сириус»", "Pcover", "Tcover"))
    out.append(title_p("2026", "Pcover", "Tcover"))
    return "".join(out)


ABSTRACT = [
    "Пояснительная записка 76 с., 63 рис., 3 источника.",
    "ОС РОСА, LINUX, KDE PLASMA, ГРАФИЧЕСКИЙ ИНТЕРФЕЙС, СЕАНС ПОЛЬЗОВАТЕЛЯ, ФАЙЛОВЫЙ МЕНЕДЖЕР, НАСТРОЙКИ, ШРИФТЫ.",
    "Объектом работы является графическая среда рабочего стола операционной системы РОСА, запущенной на виртуальной машине VMware Workstation. Предмет работы – основные приёмы управления сеансами, приложениями, файлами, съёмными носителями и персональными настройками KDE.",
    "Целью работы является освоение базовых навыков работы в Linux с графическим интерфейсом: вход и завершение сеанса, запуск и закрытие приложений, управление элементами рабочего стола, выполнение файловых операций и настройка среды пользователя.",
    "Для достижения цели были последовательно выполнены задания по управлению питанием и сеансами, работе с приложениями Chromium, KWrite, Gwenview, KCalc, KPatience и Dolphin, созданию каталогов, файлов, ярлыков и ZIP-архива, подключению USB-накопителя, а также настройке панели, раскладки клавиатуры, сети, звука, шрифтов и обоев.",
    "Результатом работы стали закреплённые навыки использования графического интерфейса ОС РОСА и оформленный отчёт, в котором все выполненные действия подтверждены скриншотами экрана.",
]


def abstract_and_toc() -> str:
    out = [p("РЕФЕРАТ", "Pstruct")]
    out.extend(p(x, "Pbody") for x in ABSTRACT)
    out.append(p("СОДЕРЖАНИЕ", "Pstruct"))
    toc = [
        ("ВВЕДЕНИЕ", 4, 0),
        ("1 Вход, завершение работы и управление сеансами", 5, 0),
        ("1.1 Вход в систему и выключение виртуальной машины", 5, 1),
        ("1.2 Блокировка и завершение сеанса пользователя", 5, 1),
        ("2 Запуск приложений и управление окнами", 16, 0),
        ("2.1 Запуск прикладных программ", 16, 1),
        ("2.2 Элементы рабочего стола и панель задач", 16, 1),
        ("3 Домашний каталог и файловые операции", 28, 0),
        ("3.1 Создание каталогов и текстовых файлов", 28, 1),
        ("3.2 Ярлыки, копирование и архивирование", 28, 1),
        ("4 Работа со съёмным носителем", 37, 0),
        ("5 Область уведомлений и панель рабочего стола", 45, 0),
        ("6 Персональные настройки рабочего стола", 54, 0),
        ("6.1 Параметры панели, меню и обоев", 54, 1),
        ("7 Шрифты и текст", 65, 0),
        ("ЗАКЛЮЧЕНИЕ", 75, 0),
        ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 76, 0),
    ]
    out.extend(toc_line(t, pg, lev) for t, pg, lev in toc)
    return "".join(out)


SECTIONS = [
    {
        "title": "1 Вход, завершение работы и управление сеансами",
        "subs": ["1.1 Вход в систему и выключение виртуальной машины", "1.2 Блокировка и завершение сеанса пользователя"],
        "intro": [
            "Работа выполнялась в ОС РОСА, установленной на виртуальную машину VMware Workstation. После запуска виртуальной машины отображается экран загрузки, затем появляется рабочий стол KDE. Вход в систему выполняется выбором учётной записи и вводом пароля.",
            "Команды выключения, перезапуска, блокировки и выхода из сеанса находятся в меню пользователя и в меню завершения работы. Выключение корректно завершает процессы гостевой ОС и отключает виртуальную машину; блокировка сохраняет запущенный сеанс и закрывает доступ к рабочему столу до повторной аутентификации.",
            "Повторный вход после блокировки восстанавливает открытые приложения. Завершение сеанса, в отличие от блокировки, закрывает приложения и возвращает к экрану выбора пользователя. Таким образом, в работе были проверены оба варианта окончания работы с сеансом.",
        ],
        "figs": [
            "Загрузка операционной системы РОСА на виртуальной машине VMware Workstation.",
            "Рабочий стол KDE после входа в систему.",
            "Рабочий сеанс пользователя и элементы панели KDE.",
            "Экран блокировки сеанса пользователя.",
            "Возврат к рабочему столу после разблокировки сеанса.",
            "Меню управления питанием и завершения сеанса.",
            "Повторная загрузка виртуальной машины после выключения.",
            "Рабочий стол после повторного входа в ОС РОСА.",
            "Экран выбора пользователя для входа в систему.",
            "Повторный вход пользователя и восстановление рабочего сеанса.",
        ],
    },
    {
        "title": "2 Запуск приложений и управление окнами",
        "subs": ["2.1 Запуск прикладных программ", "2.2 Элементы рабочего стола и панель задач"],
        "intro": [
            "Приложения в KDE запускаются через меню «Пуск», значки панели задач или ярлыки рабочего стола. В ходе работы последовательно открыты Chromium, KWrite, Gwenview, KCalc и KPatience. Запущенные окна отображаются кнопками на панели задач и могут быть свернуты, развернуты или закрыты.",
            "Dolphin запускался двумя способами: через меню «Пуск» и через панель задач. Повторное нажатие на уже активную пиктограмму переключает фокус на существующее окно, а не создаёт вторую копию. Закрытие приложений выполнено кнопками окна и командами контекстного меню панели задач.",
        ],
        "figs": [
            "Запуск приложений через меню «Пуск».",
            "Открытие файлового менеджера Dolphin из меню приложений.",
            "Окно Dolphin, запущенное с панели задач.",
            "Открытые окна приложений на рабочем столе.",
            "Управление окном приложения кнопками заголовка.",
            "Контекстное меню кнопки запущенного приложения на панели задач.",
            "Игра KPatience, открытая в графической среде KDE.",
            "Рабочее окно KPatience и элементы управления игрой.",
            "Калькулятор KCalc среди открытых приложений.",
            "Закрытие запущенных приложений через элементы управления окнами.",
            "Проверка рабочего стола после закрытия приложений.",
        ],
    },
    {
        "title": "3 Домашний каталог и файловые операции",
        "subs": ["3.1 Создание каталогов и текстовых файлов", "3.2 Ярлыки, копирование и архивирование"],
        "intro": [
            "Домашний каталог пользователя содержит стандартные папки: «Документы», «Загрузки», «Изображения», «Музыка», «Видео» и «Общедоступные». Для навигации и выполнения операций использовался Dolphin.",
            "В папке «Документы» создан каталог ПР1 и текстовый файл ПР1.txt. Файл открыт в редакторе KWrite, дополнен несколькими строками произвольного текста и сохранён. В каталоге «Общедоступные» создана папка ПР2 с файлом ПР2.txt; для файла создан ярлык на рабочем столе.",
            "Папка ПР2 скопирована в ПР1. После проверки содержимого каталог ПР1 упакован в ZIP-архив. Ярлык ПР2.txt проверен открытием исходного файла из рабочего стола.",
        ],
        "figs": [
            "Домашний каталог пользователя в файловом менеджере Dolphin.",
            "Содержимое папки «Документы».",
            "Создание папки ПР1 в каталоге «Документы».",
            "Создание файла ПР1.txt в папке ПР1.",
            "Открытие ПР1.txt в текстовом редакторе KWrite.",
            "Сохранённый текстовый файл ПР1.txt.",
            "Папка ПР2 и файл ПР2.txt в каталоге «Общедоступные».",
            "Копирование ПР2 в ПР1 и подготовка ZIP-архива.",
        ],
    },
    {
        "title": "4 Работа со съёмным носителем",
        "subs": [],
        "intro": [
            "USB-накопитель подключён к виртуальной машине и открыт в Dolphin. Среда РОСА автоматически смонтировала носитель и показала его в боковой панели файлового менеджера. На внешний носитель скопированы ZIP-архив ПР1 и папка ПР2.",
            "Перед физическим отключением использована команда безопасного извлечения. После повторного подключения накопитель снова открыт в Dolphin, наличие архива и папки проверено, затем носитель повторно извлечён безопасным способом. Такой порядок предотвращает потерю данных из-за незавершённой записи.",
        ],
        "figs": [
            "Рабочий стол после подключения USB-накопителя.",
            "Уведомление о подключённом съёмном носителе.",
            "Содержимое USB-накопителя в файловом менеджере Dolphin.",
            "Выбор ZIP-архива и папки ПР2 для копирования.",
            "Копирование подготовленных данных на внешний носитель.",
            "Команда безопасного извлечения USB-накопителя.",
            "Повторное подключение носителя и проверка сохранённых данных.",
        ],
    },
    {
        "title": "5 Область уведомлений и панель рабочего стола",
        "subs": [],
        "intro": [
            "Кнопка «Взглянуть на рабочий стол» сворачивает все окна и повторным нажатием возвращает их на экран. Она полезна для быстрого доступа к ярлыкам и файлам рабочего стола.",
            "В области уведомлений расположены цифровые часы, индикатор раскладки клавиатуры, сетевое соединение и регулятор громкости. Через свойства цифровых часов открыта вкладка «Внешний вид», где задаётся формат даты и времени. Панель также позволяет быстро просмотреть состояние сети, изменить язык раскладки и уровень громкости.",
        ],
        "figs": [
            "Свернутые окна после нажатия кнопки «Взглянуть на рабочий стол».",
            "Развернутые окна после повторного нажатия кнопки.",
            "Вызов цифровых часов из области уведомлений.",
            "Окно настройки цифровых часов, вкладка «Внешний вид».",
            "Меню выбора языка раскладки клавиатуры.",
            "Просмотр сведений о текущем сетевом соединении.",
            "Изменение уровня громкости в области уведомлений.",
            "Панель KDE с настроенными элементами области уведомлений.",
        ],
    },
    {
        "title": "6 Персональные настройки рабочего стола",
        "subs": ["6.1 Параметры панели, меню и обоев"],
        "intro": [
            "Персональные параметры KDE изменяются через «Параметры системы» и контекстные меню рабочего стола и панели. В настройках шрифтов выбрана средняя степень хинтинга, улучшающая читаемость символов на экране. Для панели задач включено автоскрытие, благодаря чему увеличивается свободная площадь рабочего стола.",
            "Внешний вид меню «Пуск» изменён: заменён значок меню, создана категория «Часто используемое» и помещена в нижнюю часть списка. В категорию скопированы ярлыки KWrite, KolourPaint, Konsole и KCalc.",
            "Через браузер загружено несколько изображений, после чего одно из них установлено в качестве обоев рабочего стола. Настройка обоев выполняется в свойствах рабочего стола и применяется к выбранному рабочему столу или ко всем рабочим столам.",
        ],
        "figs": [
            "Открытие настроек рабочего стола KDE.",
            "Настройка степени хинтинга шрифтов: среднее значение.",
            "Включение автоскрытия панели задач.",
            "Изменение значка меню «Пуск».",
            "Создание в меню новой категории «Часто используемое».",
            "Перемещение категории «Часто используемое» в нижнюю часть списка.",
            "Ярлыки KWrite, KolourPaint, Konsole и KCalc в новой категории.",
            "Скачивание изображений для оформления рабочего стола через браузер.",
            "Выбор загруженного изображения в настройках обоев.",
            "Рабочий стол с установленными пользовательскими обоями.",
        ],
    },
    {
        "title": "7 Шрифты и текст",
        "subs": [],
        "intro": [
            "Клавиша Compose используется для ввода символов, которые не представлены отдельной клавишей клавиатуры. В параметрах клавиатуры функция Compose включена и закреплена за правой клавишей Alt.",
            "Шрифт Times New Roman скачан из сети Интернет и добавлен в систему как системный шрифт. Системная установка делает шрифт доступным не только текущему пользователю, но и приложениям, которые используют общий каталог шрифтов.",
            "После изменения настроек проверена доступность установленных параметров в интерфейсе ОС РОСА и приложениях рабочего стола.",
        ],
        "figs": [
            "Открытие параметров клавиатуры и дополнительных способов ввода.",
            "Включение клавиши Compose.",
            "Назначение клавиши Compose на правый Alt.",
            "Настройки способов ввода текста в ОС РОСА.",
            "Открытие браузера для загрузки Times New Roman.",
            "Выбор файла шрифта Times New Roman.",
            "Установка шрифта в системный каталог.",
            "Проверка Times New Roman в списке установленных шрифтов.",
            "Итоговая проверка настроек рабочего стола и шрифтов.",
        ],
    },
]


def main_body() -> str:
    out = [p("ВВЕДЕНИЕ", "Pstruct")]
    out.extend([
        p("Графический интерфейс Linux предоставляет пользователю набор средств для управления системой без обязательного ввода команд в терминале. В среде KDE Plasma действия выполняются через меню приложений, панели, контекстные меню, файловый менеджер и окна параметров. Освоение этих элементов необходимо для повседневной работы с ОС РОСА.", "Pbody"),
        p("Цель лабораторной работы – изучить основы работы в Linux с графическим интерфейсом на примере ОС РОСА, запущенной на виртуальной машине VMware Workstation, и получить навыки настройки пользовательского рабочего окружения.", "Pbody"),
        p("Для достижения цели необходимо выполнить следующие задачи:", "Pbody"),
    ])
    for item in [
        "войти в ОС, проверить выключение компьютера, блокировку и завершение сеанса пользователя;",
        "запустить и закрыть приложения, освоить управление окнами и элементами рабочего стола;",
        "создать каталоги, файлы, ярлыки и ZIP-архив, выполнить копирование данных;",
        "подключить USB-накопитель, проверить монтирование и безопасное извлечение;",
        "настроить область уведомлений, панель задач, меню «Пуск», обои и шрифты;",
        "включить Compose и назначить правый Alt, установить системный Times New Roman.",
    ]:
        out.append(p("– " + item, "Plist"))
    out.append(p("Все действия выполнены в графической среде ОС РОСА и подтверждены скриншотами, приведёнными в основной части отчёта.", "Pbody"))

    for sec in SECTIONS:
        out.append(page_break())
        out.append(p(sec["title"], "Psection", outline=1))
        for sub in sec["subs"]:
            out.append(p(sub, "Psub", outline=2))
        out.extend(p(x, "Pbody") for x in sec["intro"])
        for idx, desc in enumerate(sec["figs"]):
            # Изображения в репозитории идут в том же порядке, что и задания.
            # Общая нумерация рисунков начинается с 1.
            image_no = sum(len(s["figs"]) for s in SECTIONS[:SECTIONS.index(sec)]) + idx + 1
            out.append(figure(image_no, desc, image_no))
    return "".join(out)


def conclusion_and_refs() -> str:
    out = [p("ЗАКЛЮЧЕНИЕ", "Pstruct")]
    out.append(p("В ходе лабораторной работы изучены основные элементы графического интерфейса ОС РОСА на базе KDE Plasma. Выполнены вход в систему, выключение виртуальной машины, блокировка и завершение пользовательского сеанса с последующим повторным входом.", "Pbody"))
    out.append(p("Последовательно запущены Chromium, KWrite, Gwenview, KCalc, KPatience и Dolphin; освоены способы управления окнами через кнопки, контекстные меню и панель задач. В домашнем каталоге созданы ПР1 и ПР2, текстовые файлы, ярлык, копия каталога и ZIP-архив.", "Pbody"))
    out.append(p("Проверена работа со съёмным носителем: выполнены подключение, копирование данных, безопасное извлечение, повторное подключение и контроль сохранности файлов. Настроены область уведомлений, часы, язык раскладки, сеть, громкость, панель задач и меню «Пуск».", "Pbody"))
    out.append(p("Дополнительно выполнены персональные настройки: выбран средний хинтинг шрифтов, включено автоскрытие панели, изменены значок меню и обои, создана категория «Часто используемое». Функция Compose назначена на правый Alt, системно установлен шрифт Times New Roman. Цель и задачи лабораторной работы выполнены.", "Pbody"))
    out.extend([page_break(), p("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "Pstruct")])
    refs = [
        "ОС РОСА : официальный сайт компании ООО «НТЦ ИТ РОСА». – URL: https://rosa.ru/ (дата обращения: 17.09.2026).",
        "KDE UserBase : документация по рабочему столу KDE Plasma. – URL: https://userbase.kde.org/ (дата обращения: 17.09.2026).",
        "VMware Workstation Pro Documentation : руководство пользователя. – URL: https://techdocs.broadcom.com/ (дата обращения: 17.09.2026).",
    ]
    for i, ref in enumerate(refs, 1):
        out.append(p(f"{i} {ref}", "Pbody0"))
    return "".join(out)


def content_xml() -> str:
    body = cover() + abstract_and_toc() + main_body() + conclusion_and_refs()
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<office:document-content {NS}>'
        '<office:scripts/>'
        f'{automatic_styles()}'
        '<office:body><office:text text:use-soft-page-breaks="true">'
        '<text:sequence-decls>'
        '<text:sequence-decl text:display-outline-level="0" text:name="Illustration"/>'
        '<text:sequence-decl text:display-outline-level="0" text:name="Table"/>'
        '<text:sequence-decl text:display-outline-level="0" text:name="Text"/>'
        '<text:sequence-decl text:display-outline-level="0" text:name="Drawing"/>'
        '<text:sequence-decl text:display-outline-level="0" text:name="Figure"/>'
        '</text:sequence-decls>'
        + body
        + '</office:text></office:body></office:document-content>'
    )


def meta_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3"><office:meta><meta:generator>LibreOffice Writer-compatible ODT generator</meta:generator><dc:title>Отчёт по лабораторной работе № 2. Основы работы в Linux с GUI</dc:title><dc:subject>ОС РОСА, KDE, графический интерфейс</dc:subject><dc:description>Отчёт по работе в ОС РОСА на виртуальной машине VMware Workstation.</dc:description><dc:language>ru-RU</dc:language><meta:creation-date>2026-09-17T00:00:00</meta:creation-date></office:meta></office:document-meta>'''


def manifest_xml() -> str:
    entries = [
        '<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"/>',
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>',
        '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>',
        '<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>',
        '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>',
    ]
    entries += [f'<manifest:file-entry manifest:full-path="Pictures/fig-{i:03d}.jpg" manifest:media-type="image/jpeg"/>' for i in range(1, 64)]
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
        + ''.join(entries) + '</manifest:manifest>'
    )


def build() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(EXAMPLE) as src:
        styles = src.read("styles.xml")
        settings = src.read("settings.xml")
        thumbnail = src.read("Thumbnails/thumbnail.png") if "Thumbnails/thumbnail.png" in src.namelist() else None
    # ODT requires mimetype to be the first, uncompressed entry.
    with zipfile.ZipFile(OUT, "w") as dst:
        dst.writestr("mimetype", "application/vnd.oasis.opendocument.text", compress_type=zipfile.ZIP_STORED)
        dst.writestr("content.xml", content_xml().encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
        dst.writestr("styles.xml", styles, compress_type=zipfile.ZIP_DEFLATED)
        dst.writestr("settings.xml", settings, compress_type=zipfile.ZIP_DEFLATED)
        dst.writestr("meta.xml", meta_xml().encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
        dst.writestr("META-INF/manifest.xml", manifest_xml().encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
        if thumbnail:
            dst.writestr("Thumbnails/thumbnail.png", thumbnail, compress_type=zipfile.ZIP_DEFLATED)
        for i in range(1, 64):
            image = SHOTS / f"{i}.jpg"
            if not image.exists():
                raise FileNotFoundError(image)
            dst.write(image, f"Pictures/fig-{i:03d}.jpg", compress_type=zipfile.ZIP_DEFLATED)
    print(f"Создан файл: {OUT}")
    print(f"Размер: {OUT.stat().st_size / 1024 / 1024:.2f} MiB")
    print("Рисунков: 63")


if __name__ == "__main__":
    build()

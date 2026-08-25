#!/usr/bin/env python3
"""
Генератор эталонного DOCX-шаблона.
Демонстрирует все стили из конфига config_from_Миасс.yaml
с описаниями параметров форматирования.
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


OUTPUT = os.path.join(os.path.dirname(__file__), 'dist', 'templates', 'reference_format.docx')


def create_styles(doc):
    """Создаёт все нужные стили в документе."""
    from docx.enum.style import WD_STYLE_TYPE

    styles = [
        # (имя, шрифт, размер, жирный, выравнивание, first_line_cm, space_before, space_after, line_spacing)
        ('021216Раздел', 'Times New Roman', 14, True, None, 0, 0, 0, 1.15),
        ('021216Глава', 'Times New Roman', 12, True, None, 0, 6, 0, 1.15),
        ('021216Подглава', 'Times New Roman', 12, False, None, 0, 6, 0, 1.15),
        ('Пункт', 'Times New Roman', 12, True, None, 0, 6, 0, None),
        ('Текст доклада', 'Times New Roman', 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 1.25, 0, 0, 1.15),
        ('List Paragraph', 'Times New Roman', 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 0.75, 0, 0, 1.15),
        ('Абзац', None, None, False, None, 1.25, 6, 3, 1.15),
        ('Caption', 'Times New Roman', 12, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 12, 6, 1.0),
        ('Table Note', 'Times New Roman', 10, False, None, 0, 0, 0, None),
        ('Note', 'Times New Roman', 10, False, None, 0, 0, 0, None),
    ]

    for name, font_name, size, bold, align, fl, sb, sa, ls in styles:
        try:
            doc.styles[name]
        except KeyError:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            if font_name:
                style.font.name = font_name
                try:
                    style._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                except Exception:
                    pass
            if size:
                style.font.size = Pt(size)
            style.font.bold = bold
            if align is not None:
                style.paragraph_format.alignment = align
            if fl:
                style.paragraph_format.first_line_indent = Cm(fl)
            if sb:
                style.paragraph_format.space_before = Pt(sb)
            if sa is not None and sa > 0:
                style.paragraph_format.space_after = Pt(sa)
            if ls:
                style.paragraph_format.line_spacing = ls


def setup_page(doc):
    """Настройка полей страницы A4."""
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)


def setup_header_footer(doc):
    """Колонтитулы."""
    section = doc.sections[0]
    # Верхний колонтитул — номер страницы
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = hp.add_run()
    fld_char1 = run._element.makeelement(qn('w:fldChar'), {qn('w:fldCharType'): 'begin'})
    run._element.append(fld_char1)
    run2 = hp.add_run(' PAGE ')
    fld_code = run2._element.makeelement(qn('w:instrText'), {qn('xml:space'): 'preserve'})
    fld_code.text = ' PAGE '
    run2._element.append(fld_code)
    run3 = hp.add_run()
    fld_char2 = run3._element.makeelement(qn('w:fldChar'), {qn('w:fldCharType'): 'end'})
    run3._element.append(fld_char2)

    # Нижний колонтитул — SECTION/CHAPTER
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run('РАЗДЕЛ 1. ГЛАВА 1.1')
    run.font.bold = True
    run.font.size = Pt(10)


def add_title(doc):
    """Титул справочника."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(72)
    p.paragraph_format.space_after = Pt(24)
    run = p.add_run('СПРАВОЧНИК ФОРМАТИРОВАНИЯ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(16)
    run.font.bold = True

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(12)
    run2 = p2.add_run('Конфиг: config_from_Миасс.yaml')
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(12)
    run2.font.italic = True
    run2.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(36)
    run3 = p3.add_run('Каждый стиль показан с примером и описанием параметров из конфига.')
    run3.font.name = 'Times New Roman'
    run3.font.size = Pt(10)
    run3.font.color.rgb = RGBColor(0x99, 0x99, 0x99)


def add_style_header(doc, text, level=1):
    """Заголовок секции стиля."""
    style_name = {1: '021216Раздел', 2: '021216Глава', 3: '021216Подглава'}[level]
    p = doc.add_paragraph(style=style_name)
    run = p.add_run(text)
    return p


def add_description(doc, text):
    """Описание параметров стиля (серый текст)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.name = 'Consolas'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x33, 0x99, 0x33)
    return p


def add_example(doc, style_name, text):
    """Пример использования стиля."""
    p = doc.add_paragraph(style=style_name)
    run = p.add_run(text)
    return p


def build():
    doc = Document()
    create_styles(doc)
    setup_page(doc)
    setup_header_footer(doc)
    add_title(doc)

    # ============================================================
    # РАЗДЕЛ 1: ЗАГОЛОВКИ
    # ============================================================
    add_style_header(doc, 'РАЗДЕЛ 1. ЗАГОЛОВКИ', 1)
    add_description(doc, '021216Раздел: TNR 14pt bold, line_spacing: 1.15')

    add_style_header(doc, 'ГЛАВА 1.1 Пример главы', 2)
    add_description(doc, '021216Глава: TNR 12pt bold, space_before: 6pt, line_spacing: 1.15')

    add_style_header(doc, '1.1.1 Пример подзаголовка', 3)
    add_description(doc, '021216Подглава: TNR 12pt, space_before: 6pt, space_after: 0pt, line_spacing: 1.15')

    p = doc.add_paragraph(style='021216Подглава')
    p.add_run('1.1.2 Ещё один подзаголовок')
    add_description(doc, '021216Подглава: TNR 12pt, space_before: 6pt, space_after: 0pt, line_spacing: 1.15')

    p = doc.add_paragraph(style='Пункт')
    p.add_run('1.1.1.1 Пример пункта')
    add_description(doc, 'Пункт: TNR 12pt bold, space_before: 6pt, space_after: 0pt')

    # ============================================================
    # РАЗДЕЛ 2: ОСНОВНОЙ ТЕКСТ
    # ============================================================
    add_style_header(doc, 'РАЗДЕЛ 2. ТЕКСТ', 1)
    add_description(doc, 'Базовый текст: Normal (TNR 12pt), Текст доклада (TNR 12pt, отступ 1.25см), Абзац (отступ 1.25см)')

    add_style_header(doc, 'ГЛАВА 2.1 Основные стили текста', 2)

    add_example(doc, 'Абзац', 'Абзац с отступом красной строки 1,25 см. Интервал перед 6пт, после 3пт, межстрочный 1.15. Используется для выделенных абзацев.')
    add_description(doc, 'Абзац: first_line_indent_cm: 1.25, space_before_pt: 6.0, space_after_pt: 3.0, line_spacing: 1.15')

    add_example(doc, 'Текст доклада', 'Текст доклада — выравнивание по ширине, отступ красной строки 1,25 см. Основной стиль для содержательной части документа.')
    add_description(doc, 'Текст доклада: TNR 12pt, first_line_indent_cm: 1.25, alignment: justify, line_spacing: 1.15')

    add_example(doc, 'Normal', 'Обычный текст без отступа. Наследует Times New Roman 12pt из базового стиля.')
    add_description(doc, 'Normal: TNR 12pt (базовый стиль документа)')

    add_example(doc, 'Normal', 'Второй абзац Normal — без отступа красной строки, выравнивание по левому краю.')

    # ============================================================
    # РАЗДЕЛ 3: СПИСКИ
    # ============================================================
    add_style_header(doc, 'РАЗДЕЛ 3. СПИСКИ', 1)

    add_style_header(doc, 'ГЛАВА 3.1 Маркированный список', 2)

    p = doc.add_paragraph(style='List Paragraph')
    p.paragraph_format.first_line_indent = Cm(0.75)
    run = p.add_run('1) Первый пункт списка с нумерацией')
    add_description(doc, 'List Paragraph: TNR 12pt, first_line_indent_cm: 0.75, alignment: justify, line_spacing: 1.15')

    p = doc.add_paragraph(style='List Paragraph')
    p.paragraph_format.first_line_indent = Cm(0.75)
    run = p.add_run('2) Второй пункт списка')

    p = doc.add_paragraph(style='List Paragraph')
    p.paragraph_format.first_line_indent = Cm(0.75)
    run = p.add_run('3) Третий пункт списка')

    # ============================================================
    # РАЗДЕЛ 4: ТАБЛИЦЫ
    # ============================================================
    add_style_header(doc, 'РАЗДЕЛ 4. ТАБЛИЦЫ', 1)

    add_style_header(doc, 'ГЛАВА 4.1 Пример таблицы', 2)

    # Заголовок таблицы (Caption)
    p = doc.add_paragraph(style='Caption')
    run = p.add_run('Таблица 4.1 — Пример таблицы с данными')
    add_description(doc, 'Caption: TNR 12pt bold, centered, перед таблицей')

    # Таблица
    table = doc.add_table(rows=4, cols=3, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Заголовки
    headers = ['Наименование', 'Значение', 'Ед. изм.']
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(10)

    # Данные
    data = [
        ('Показатель 1', '1 234,5', 'чел.'),
        ('Показатель 2', '98,7', '%'),
        ('Показатель 3', '456,0', 'га'),
    ]
    for row_idx, row_data in enumerate(data, 1):
        for col_idx, val in enumerate(row_data):
            cell = table.rows[row_idx].cells[col_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.size = Pt(10)

    add_description(doc, 'Table Grid: стиль таблицы по умолчанию')

    # Примечание к таблице
    p = doc.add_paragraph(style='Table Note')
    run = p.add_run('Примечание — данные приведены по состоянию на 01.01.2025 г.')
    run.font.size = Pt(10)
    add_description(doc, 'Table Note: TNR 10pt (шрифт примечания к таблице)')

    # ============================================================
    # РАЗДЕЛ 5: ПРИМЕЧАНИЯ
    # ============================================================
    add_style_header(doc, 'РАЗДЕЛ 5. ПРИМЕЧАНИЯ', 1)

    add_style_header(doc, 'ГЛАВА 5.1 Типы примечаний', 2)

    p = doc.add_paragraph(style='Table Note')
    run = p.add_run('Примечание к таблице — шрифт 10pt, без отступа красной строки')
    add_description(doc, 'Table Note: TNR 10pt, first_line_indent: 0, space_before: 0')

    p = doc.add_paragraph(style='Note')
    run = p.add_run('Общее примечание — шрифт 10pt')
    add_description(doc, 'Note: TNR 10pt')

    # ============================================================
    # РАЗДЕЛ 6: НАСТРОЙКИ СТРАНИЦЫ
    # ============================================================
    add_style_header(doc, 'РАЗДЕЛ 6. НАСТРОЙКИ СТРАНИЦЫ', 1)

    add_style_header(doc, 'ГЛАВА 6.1 Поля и колонтитулы', 2)

    p = doc.add_paragraph(style='Текст доклада')
    run = p.add_run('Поля страницы: левое 2,5 см, правое 1,0 см, верхнее 2,0 см, нижнее 2,0 см')
    add_description(doc, 'page_setup: left_margin_cm: 2.5, right_margin_cm: 1.0, top_margin_cm: 2.0, bottom_margin_cm: 2.0')

    p = doc.add_paragraph(style='Текст доклада')
    run = p.add_run('Верхний колонтитул: номер страницы (по правому краю, отступ 1,0 см)')
    add_description(doc, 'header: enabled, position_from_edge_cm: 1.0')

    p = doc.add_paragraph(style='Текст доклада')
    run = p.add_run('Нижний колонтитул: РАЗДЕЛ X. ГЛАВА X.X (по центру, жирный, отступ 1,0 см)')
    add_description(doc, 'footer: enabled, position_from_edge_cm: 1.0, alignment: center, bold: true')

    # ============================================================
    # СВОДНАЯ ТАБЛИЦА СТИЛЕЙ
    # ============================================================
    doc.add_page_break()
    add_style_header(doc, 'РАЗДЕЛ 7. СВОДНАЯ ТАБЛИЦА СТИЛЕЙ', 1)
    add_description(doc, 'Все стили из конфига config_from_Миасс.yaml')

    summary_table = doc.add_table(rows=9, cols=5, style='Table Grid')
    summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Заголовки
    sh = ['Стиль', 'Шрифт', 'Размер', 'Жирный', 'Отступ красной строки']
    for i, h in enumerate(sh):
        cell = summary_table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)

    # Данные
    styles_data = [
        ('021216Раздел', 'TNR', '14', 'Да', 'Нет'),
        ('021216Глава', 'TNR', '12', 'Да', 'Нет'),
        ('021216Подглава', 'TNR', '12', 'Нет', 'Нет'),
        ('Пункт', 'TNR', '12', 'Да', 'Нет'),
        ('Текст доклада', 'TNR', '12', 'Нет', '1,25 см'),
        ('List Paragraph', 'TNR', '12', 'Нет', '0,75 см'),
        ('Абзац', '—', '—', 'Нет', '1,25 см'),
        ('Normal', 'TNR', '12', 'Нет', 'Нет'),
    ]
    for row_idx, row_data in enumerate(styles_data, 1):
        for col_idx, val in enumerate(row_data):
            cell = summary_table.rows[row_idx].cells[col_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            run = p.add_run(val)
            run.font.size = Pt(9)
            if col_idx == 0:
                run.font.bold = True

    # Сохранение
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    doc.save(OUTPUT)
    print(f'Справочник создан: {OUTPUT}')


if __name__ == '__main__':
    build()

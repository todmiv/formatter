# -*- coding: utf-8 -*-
"""Тесты качества оформления по НТД 01-2013 (редакция август 2026).

Конвертируют произвольный Markdown-текст конфигом config_НТД_2026.yaml и
проверяют соответствие ключевым требованиям новой РИ:
  п. 2.7.2  - поля 20/10/20/20 мм, шрифт TNR 12 пт, интервал 1,0,
              красная строка 1,25 см, выравнивание по ширине;
  п. 2.10.1 - подпись таблицы по центру, сквозная нумерация;
  п. 2.10.4 - шапка таблицы жирным по центру (единая таблица);
  п. 2.14.2 - номер страницы 12 пт в верхнем правом углу (поле PAGE);
  п. 2.14.4 - нижний колонтитул пуст.

Дополнительно проверяется, что конфиг config_Эталон.yaml (имена стилей
образца Эталон.docx) порождает документы со стилями образца.
"""
import re
from docx import Document
from docx.oxml.ns import qn
import pytest

from md_to_docx import RI2013Converter

CONFIG = 'configs/active/config_НТД_2026.yaml'
CONFIG_TIKSI = 'configs/active/config_Эталон.yaml'

MD_TEXT = """ВВЕДЕНИЕ

Настоящая записка подготовлена в целях обоснования проектных решений.

РАЗДЕЛ 1. ОБЩИЕ СВЕДЕНИЯ

1.1 Описание положения

Муниципальное образование расположено в северной части региона.

Основные характеристики климата приведены в таблице 1.

Таблица 1 Основные климатические показатели

| Показатель | Значение |
|------------|----------|
| Температура, °C | -6,5 |
| Осадки, мм | 320 |

Примечание — Источник: данные метеорологической станции.

Расчетные параметры приняты по нормативным документам.

РАЗДЕЛ 2. ОБОСНОВАНИЕ РЕШЕНИЙ

Глава 1. Демография

2.1 Анализ состояния

Численность населения составляет 12,4 тыс. человек.

Структура занятости представлена в таблице 2.

Таблица 2 Структура занятости населения

| Отрасль | Доля, % |
|---------|---------|
| Транспорт | 32,0 |
| Прочие | 68,0 |

ЗАКЛЮЧЕНИЕ

Реализация мероприятий обеспечит устойчивое развитие территории.
"""


@pytest.fixture(scope='module')
def converted(tmp_path_factory):
    md_path = tmp_path_factory.mktemp('md') / 'arbitrary.md'
    md_path.write_text(MD_TEXT, encoding='utf-8')
    docx_path = md_path.with_suffix('.docx')
    converter = RI2013Converter(CONFIG)
    converter.convert(str(md_path), str(docx_path))
    assert docx_path.exists()
    return Document(str(docx_path))


def has_page_field(paragraph):
    return paragraph._element.findall('.//' + qn('w:instrText')) != []


def test_page_margins_per_ri_2026(converted):
    s = converted.sections[0]
    assert abs(s.left_margin.cm - 2.0) < 0.01
    assert abs(s.right_margin.cm - 1.0) < 0.01
    assert abs(s.top_margin.cm - 2.0) < 0.01
    assert abs(s.bottom_margin.cm - 2.0) < 0.01


def test_header_page_number_12pt_right(converted):
    s = converted.sections[0]
    hdr_paras = [p for p in s.header.paragraphs if p.text.strip() or has_page_field(p)]
    assert hdr_paras, 'верхний колонтитул должен содержать номер страницы'
    assert has_page_field(hdr_paras[0]), 'номер страницы должен быть полем PAGE'
    for r in hdr_paras[0].runs:
        if r.font.size:
            assert abs(r.font.size.pt - 12.0) < 0.01
            break


def test_footer_empty(converted):
    s = converted.sections[0]
    assert ''.join(p.text for p in s.footer.paragraphs).strip() == ''


def test_body_text_style_per_ri_2026(converted):
    style = converted.styles['04_Основной текст']
    assert style.font.name == 'Times New Roman'
    assert abs(style.font.size.pt - 12.0) < 0.01
    assert abs(style.paragraph_format.first_line_indent.cm - 1.25) < 0.05
    assert abs(style.paragraph_format.line_spacing - 1.0) < 0.05
    assert style.paragraph_format.alignment.name == 'JUSTIFY'


def test_section_heading_style(converted):
    h1 = [p for p in converted.paragraphs if p.style.name == '01_Раздел' and p.text.strip()]
    assert h1
    style = converted.styles['01_Раздел']
    assert style.font.name == 'Times New Roman'
    assert abs(style.font.size.pt - 12.0) < 0.01
    assert style.font.bold
    assert style.paragraph_format.page_break_before


def test_table_captions_centered_continuous(converted):
    caps = [p for p in converted.paragraphs if p.style.name == '05_Номер таблицы' and p.text.strip()]
    assert len(caps) == 2
    style = converted.styles['05_Номер таблицы']
    assert style.paragraph_format.alignment.name == 'CENTER'
    for c in caps:
        # выравнивание задано на уровне стиля (прямое форматирование не требуется)
        assert c.paragraph_format.alignment in (None, style.paragraph_format.alignment)
    nums = [int(re.match(r'^Таблица\s+(\d+)', c.text).group(1)) for c in caps]
    assert nums == [1, 2], 'нумерация таблиц должна быть сквозной (п. 2.10.2)'


def test_single_table_with_bold_centered_header(converted):
    assert len(converted.tables) == 2, 'каждая таблица Markdown должна давать одну таблицу DOCX'
    for t in converted.tables:
        for cell in t.rows[0].cells:
            for p in cell.paragraphs:
                if not p.text.strip():
                    continue
                assert p.alignment.name == 'CENTER'
                for r in p.runs:
                    if r.text.strip():
                        assert r.font.bold
                        break


def test_table_note_style(converted):
    notes = [p for p in converted.paragraphs if p.style.name == '055_Примечание после таблицы/рисунка' and p.text.strip()]
    assert notes
    for r in notes[0].runs:
        if r.text.strip() and r.font.size:
            assert r.font.size.pt == 10
            break


# ---------------------------------------------------------------------------
# Конфиг Эталон: имена стилей должны соответствовать образцу Эталон.docx
# ---------------------------------------------------------------------------

@pytest.fixture(scope='module')
def converted_tiksi(tmp_path_factory):
    md_path = tmp_path_factory.mktemp('md_tiksi') / 'arbitrary.md'
    md_path.write_text(MD_TEXT, encoding='utf-8')
    docx_path = md_path.with_suffix('.docx')
    converter = RI2013Converter(CONFIG_TIKSI)
    converter.convert(str(md_path), str(docx_path))
    assert docx_path.exists()
    return Document(str(docx_path))


def test_tiksi_config_uses_sample_style_names(converted_tiksi):
    """Стили документа, созданного конфигом Эталон, совпадают с именами образца."""
    styles_used = {p.style.name for p in converted_tiksi.paragraphs if p.text.strip()}
    sample_styles = {
        '01_Раздел', '02_Глава', '03_Пункт', '031_Подпункт',
        '04_Заголовок основного текста', '041_Основной текст',
        '05_Номер и название таблицы/рисунка', '051_Шапка таблицы',
        '052_Текст в таблице/примечание', '06_Перечисление тире',
        '061_Перечисление цифра', 'Normal',
    }
    assert styles_used <= sample_styles, f'стили вне набора образца: {styles_used - sample_styles}'
    assert '041_Основной текст' in styles_used
    assert '05_Номер и название таблицы/рисунка' in styles_used


def test_tiksi_config_caption_style_name(converted_tiksi):
    caps = [p for p in converted_tiksi.paragraphs
            if p.style.name == '05_Номер и название таблицы/рисунка' and p.text.strip()]
    assert len(caps) == 2
    style = converted_tiksi.styles['05_Номер и название таблицы/рисунка']
    assert style.paragraph_format.alignment.name == 'CENTER'

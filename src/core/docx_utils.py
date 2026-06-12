"""
Утилиты для работы с python-docx, общие для audit_engine и apply_engine.
"""

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from typing import Optional, Dict, Any
import docx
from docx.document import Document as DocxDocument
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.shared import RGBColor
import logging

logger = logging.getLogger(__name__)

# Константы преобразования единиц
MM_TO_TWIPS = 56.7   # 1 мм = 56.7 twips
PT_TO_TWIPS = 20     # 1 pt = 20 twips


# --- Работа с XML ---
def get_xml_attr_twips(element: Optional[OxmlElement], attr_name: str) -> Optional[int]:
    """
    Безопасно извлекает значение атрибута в twips из XML элемента.
    Если атрибут отсутствует, возвращает None.
    """
    if element is None:
        return None
    val_str = element.get(qn(attr_name))
    if val_str:
        try:
            return int(val_str)
        except ValueError:
            return None
    return None


def get_paragraph_indent_twips(paragraph: Paragraph) -> Dict[str, int]:
    """
    Возвращает отступы параграфа в twips.
    """
    pPr = paragraph._element.pPr
    first_line = 0
    left = 0
    if pPr is not None:
        indent_elem = pPr.find(qn('w:ind'))
        if indent_elem is not None:
            first_line = get_xml_attr_twips(indent_elem, 'w:firstLine') or 0
            left = get_xml_attr_twips(indent_elem, 'w:left') or 0
    # Если нет явного отступа, берём из стиля
    if first_line == 0 and paragraph.style.paragraph_format.first_line_indent:
        first_line = paragraph.style.paragraph_format.first_line_indent.twips
    if left == 0 and paragraph.style.paragraph_format.left_indent:
        left = paragraph.style.paragraph_format.left_indent.twips
    return {'first_line': first_line, 'left': left}


def get_paragraph_spacing_twips(paragraph: Paragraph) -> Dict[str, int]:
    """
    Возвращает интервалы перед и после параграфа в twips.
    """
    pPr = paragraph._element.pPr
    before = 0
    after = 0
    if pPr is not None:
        spacing_elem = pPr.find(qn('w:spacing'))
        if spacing_elem is not None:
            before = get_xml_attr_twips(spacing_elem, 'w:before') or 0
            after = get_xml_attr_twips(spacing_elem, 'w:after') or 0
    if before == 0 and paragraph.style.paragraph_format.space_before:
        before = paragraph.style.paragraph_format.space_before.twips
    if after == 0 and paragraph.style.paragraph_format.space_after:
        after = paragraph.style.paragraph_format.space_after.twips
    return {'before': before, 'after': after}


def get_paragraph_alignment(paragraph: Paragraph) -> str:
    """
    Возвращает выравнивание параграфа в виде строки ('left', 'center', 'right', 'justify').
    """
    pPr = paragraph._element.pPr
    if pPr is not None:
        jc_elem = pPr.find(qn('w:jc'))
        if jc_elem is not None:
            val = jc_elem.get(qn('w:val'))
            if val:
                map_align = {'both': 'justify', 'left': 'left', 'center': 'center', 'right': 'right'}
                return map_align.get(val, 'left')
    # По умолчанию
    return 'left'


def normalize_color(color_spec: Any) -> Optional[RGBColor]:
    """
    Преобразует спецификацию цвета (строка, RGBColor, None) в RGBColor.
    Поддерживаются:
      - RGBColor объект (возвращается как есть)
      - Строка в формате hex с # или без (например, '#000000', 'FF0000', 'fff')
      - Строка с именованным цветом (black, white, red, green, blue, gray, grey, yellow, cyan, magenta)
      - Строка в формате rgb(r,g,b) (например, 'rgb(255,0,0)')
      - None -> None
    """
    if color_spec is None:
        return None
    if isinstance(color_spec, RGBColor):
        return color_spec
    if isinstance(color_spec, str):
        color_spec = color_spec.strip().lower()
        # Именованные цвета
        named_colors = {
            'black': RGBColor(0, 0, 0),
            'white': RGBColor(255, 255, 255),
            'red': RGBColor(255, 0, 0),
            'green': RGBColor(0, 255, 0),
            'blue': RGBColor(0, 0, 255),
            'gray': RGBColor(128, 128, 128),
            'grey': RGBColor(128, 128, 128),
            'yellow': RGBColor(255, 255, 0),
            'cyan': RGBColor(0, 255, 255),
            'magenta': RGBColor(255, 0, 255),
        }
        if color_spec in named_colors:
            return named_colors[color_spec]
        
        # Формат rgb(r,g,b)
        import re
        rgb_match = re.match(r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', color_spec)
        if rgb_match:
            try:
                r = int(rgb_match.group(1))
                g = int(rgb_match.group(2))
                b = int(rgb_match.group(3))
                if 0 <= r <= 255 and 0 <= g <= 255 and 0 <= b <= 255:
                    return RGBColor(r, g, b)
            except ValueError:
                pass
        
        # Удалить # если есть
        if color_spec.startswith('#'):
            color_spec = color_spec[1:]
        
        # Hex строка длиной 3 (удваиваем каждый символ)
        if len(color_spec) == 3:
            try:
                r = int(color_spec[0] * 2, 16)
                g = int(color_spec[1] * 2, 16)
                b = int(color_spec[2] * 2, 16)
                return RGBColor(r, g, b)
            except ValueError:
                pass
        
        # Hex строка длиной 6
        if len(color_spec) == 6:
            try:
                r = int(color_spec[0:2], 16)
                g = int(color_spec[2:4], 16)
                b = int(color_spec[4:6], 16)
                return RGBColor(r, g, b)
            except ValueError:
                pass
    
    # Если не удалось распознать, возвращаем None
    logger.warning(f"Не удалось распознать цвет: {color_spec}")
    return None


def get_font_properties(run) -> Dict[str, Any]:
    """
    Возвращает свойства шрифта Run.
    """
    from docx.oxml.ns import qn
    rPr = run._element.find(qn('w:rPr'))
    font_name = None
    if rPr is not None:
        rFonts_elem = rPr.find(qn('w:rFonts'))
        if rFonts_elem is not None:
            font_name = rFonts_elem.get(qn('w:ascii')) or rFonts_elem.get(qn('w:hAnsi'))
    size_twips = run.font.size.twips if run.font.size else None
    color = run.font.color.rgb if run.font.color else None
    return {
        'name': font_name,
        'size_twips': size_twips,
        'bold': run.font.bold,
        'italic': run.font.italic,
        'underline': run.font.underline,
        'color': color,
    }


# --- Открытие документа ---
def open_document(doc_path: str) -> DocxDocument:
    """
    Открывает документ с обработкой ошибок.
    """
    try:
        return docx.Document(doc_path)
    except Exception as e:
        logger.error(f"Не удалось открыть документ {doc_path}: {e}")
        raise RuntimeError(f"Не удалось открыть документ: {e}")


# --- Конвертация единиц ---
def cm_to_twips(cm: float) -> int:
    """Конвертирует сантиметры в twips (1 см = 567 twips)."""
    return int(cm * 567)


def pt_to_twips(pt: float) -> int:
    """Конвертирует пункты в twips (1 pt = 20 twips)."""
    return int(pt * 20)


def twips_to_cm(twips: int) -> float:
    """Конвертирует twips в сантиметры."""
    return twips / 567


def twips_to_pt(twips: int) -> float:
    """Конвертирует twips в пункты."""
    return twips / 20
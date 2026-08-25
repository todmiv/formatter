"""
DocumentModel - абстракция над python-docx для улучшения тестируемости и управления документом.
Инкапсулирует работу с документом, предоставляет единый интерфейс для чтения и модификации.
"""

import docx
from docx.document import Document as DocxDocument
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Twips, Pt, Cm
from docx.oxml.ns import qn
from typing import List, Dict, Any, Optional, Iterator
import logging

logger = logging.getLogger(__name__)


class DocumentModel:
    """Модель документа, оборачивающая python-docx Document."""

    def __init__(self, doc_path: Optional[str] = None, doc: Optional[DocxDocument] = None):
        """
        Инициализация модели.
        :param doc_path: путь к файлу .docx (если None, создаётся новый документ)
        :param doc: существующий объект docx.Document (опционально)
        """
        if doc is not None:
            self._doc = doc
            self._path = None
        elif doc_path is not None:
            self._doc = docx.Document(doc_path)
            self._path = doc_path
        else:
            self._doc = docx.Document()
            self._path = None

    @property
    def path(self) -> Optional[str]:
        """Возвращает путь к файлу документа, если он был загружен из файла."""
        return self._path

    @property
    def docx_document(self) -> DocxDocument:
        """Возвращает внутренний объект docx.Document (для совместимости)."""
        return self._doc

    def save(self, save_path: Optional[str] = None):
        """Сохраняет документ в указанный путь или перезаписывает исходный."""
        target_path = save_path or self._path
        if target_path is None:
            raise ValueError("Не указан путь для сохранения документа.")
        self._doc.save(target_path)
        if save_path:
            self._path = save_path
        logger.debug(f"Документ сохранён: {target_path}")

    # --- Чтение структуры ---
    def paragraphs(self) -> List[Paragraph]:
        """Возвращает список всех параграфов документа."""
        return self._doc.paragraphs

    def paragraph_at(self, index: int) -> Optional[Paragraph]:
        """Возвращает параграф по индексу или None, если индекс вне диапазона."""
        if 0 <= index < len(self._doc.paragraphs):
            return self._doc.paragraphs[index]
        return None

    def tables(self) -> List[Table]:
        """Возвращает список всех таблиц документа."""
        return self._doc.tables

    def sections(self):
        """Возвращает список секций документа."""
        return self._doc.sections

    def styles(self):
        """Возвращает объект стилей документа."""
        return self._doc.styles

    # --- Получение свойств ---
    def get_paragraph_properties(self, para: Paragraph) -> Dict[str, Any]:
        """Извлекает свойства параграфа: стиль, выравнивание, отступы, интервалы."""
        p_format = para.paragraph_format
        props = {
            'style': para.style.name,
            'alignment': self._get_alignment_name(p_format.alignment),
            'first_line_indent_twips': p_format.first_line_indent.twips if p_format.first_line_indent else 0,
            'left_indent_twips': p_format.left_indent.twips if p_format.left_indent else 0,
            'space_before_twips': p_format.space_before.twips if p_format.space_before else 0,
            'space_after_twips': p_format.space_after.twips if p_format.space_after else 0,
            'line_spacing': p_format.line_spacing,
        }
        return props

    def get_run_font_properties(self, run) -> Dict[str, Any]:
        """Извлекает свойства шрифта Run."""
        font = run.font
        return {
            'name': font.name,
            'size_twips': font.size.twips if font.size else None,
            'bold': font.bold,
            'italic': font.italic,
            'underline': font.underline,
            'color': font.color.rgb if font.color else None,
        }

    def get_style_config(self, style_name: str) -> Optional[Dict]:
        """Возвращает конфигурацию стиля из документа (если стиль существует)."""
        try:
            style = self._doc.styles[style_name]
            # Здесь можно преобразовать в словарь, но для простоты вернём объект стиля
            # В будущем можно расширить.
            return {'name': style.name, 'type': style.type}
        except KeyError:
            return None

    # --- Модификация ---
    def apply_style(self, paragraph: Paragraph, style_name: str):
        """Применяет стиль к параграфу."""
        paragraph.style = style_name

    def set_paragraph_format(self, paragraph: Paragraph, **kwargs):
        """Устанавливает форматирование параграфа."""
        p_format = paragraph.paragraph_format
        if 'alignment' in kwargs:
            align_map = {
                'left': 0,  # WD_ALIGN_PARAGRAPH.LEFT
                'center': 1,
                'right': 2,
                'justify': 3,
            }
            val = kwargs['alignment']
            if isinstance(val, str):
                p_format.alignment = align_map.get(val, 0)
            else:
                p_format.alignment = val
        if 'first_line_indent_cm' in kwargs:
            p_format.first_line_indent = Cm(kwargs['first_line_indent_cm'])
        if 'left_indent_cm' in kwargs:
            p_format.left_indent = Cm(kwargs['left_indent_cm'])
        if 'space_before_pt' in kwargs:
            p_format.space_before = Pt(kwargs['space_before_pt'])
        if 'space_after_pt' in kwargs:
            p_format.space_after = Pt(kwargs['space_after_pt'])
        if 'line_spacing' in kwargs:
            p_format.line_spacing = kwargs['line_spacing']

    def set_run_font(self, run, **kwargs):
        """Устанавливает свойства шрифта Run."""
        if 'name' in kwargs:
            run.font.name = kwargs['name']
        if 'size_pt' in kwargs:
            run.font.size = Pt(kwargs['size_pt'])
        if 'bold' in kwargs:
            run.font.bold = kwargs['bold']
        if 'italic' in kwargs:
            run.font.italic = kwargs['italic']
        if 'underline' in kwargs:
            run.font.underline = kwargs['underline']

    def clear_direct_formatting(self, paragraph: Paragraph):
        """Очищает прямое форматирование параграфа (сбрасывает runs)."""
        # Простой способ: объединить все runs в один с текстом и очистить форматирование
        text = paragraph.text
        paragraph.clear()
        run = paragraph.add_run(text)
        # Установить шрифт по умолчанию (стиль параграфа)
        run.font.name = None
        run.font.size = None
        run.font.bold = None
        run.font.italic = None
        run.font.underline = None

    # --- Вспомогательные методы ---
    @staticmethod
    def _get_alignment_name(alignment) -> str:
        """Преобразует значение выравнивания в строку."""
        if alignment is None:
            return 'left'
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        mapping = {
            WD_ALIGN_PARAGRAPH.LEFT: 'left',
            WD_ALIGN_PARAGRAPH.CENTER: 'center',
            WD_ALIGN_PARAGRAPH.RIGHT: 'right',
            WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify',
        }
        return mapping.get(alignment, 'left')

    def estimate_page_number(self, paragraph_index: int, lines_per_page: int = 25) -> int:
        """Оценивает номер страницы для параграфа по его индексу."""
        # Простая эвристика: считаем, что каждый параграф занимает примерно 1 строку
        return (paragraph_index // lines_per_page) + 1

    # --- Контекстный менеджер ---
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # При необходимости можно реализовать автосохранение
        pass
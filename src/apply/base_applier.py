"""
Базовый класс Applier и общие утилиты для применения исправлений.
"""

import logging
from typing import Dict, Any, Optional, Tuple
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

from src.core.audit_engine import AuditIssue
from src.core.config_loader import ConfigLoader
from src.core.style_manager import StyleManager
from src.core.page_manager import PageManager
from src.core.docx_utils import open_document

logger = logging.getLogger(__name__)


class BaseApplier:
    """
    Базовый класс для всех апплеров.
    Содержит общие методы для поиска элементов в документе и очистки форматирования.
    """

    def __init__(self, config_loader: ConfigLoader):
        self.config = config_loader
        self.style_manager = StyleManager(config_loader)
        self.page_manager = PageManager(config_loader)

    def _get_paragraph(self, doc: Document, location: Dict[str, Any]) -> Tuple[Paragraph, str]:
        """
        Находит параграф в документе по location.

        Поддерживаемые форматы location:
        - {'index': int} - параграф в основном теле документа
        - {'paragraph_index': int} - альтернативный ключ для основного тела
        - {'section': int, 'paragraph': int} - параграф в колонтитуле (без указания container)
        - {'section': int, 'paragraph': int, 'container': 'header'|'footer'} - явное указание колонтитула

        Возвращает кортеж (paragraph, container_type), где container_type:
        'body', 'header', 'footer'.

        Выбрасывает FormatterError, если параграф не найден.
        """
        from src.core.formatter_errors import (
            raise_paragraph_index_error, raise_header_footer_error,
            raise_location_invalid, FormatterError, ErrorKind,
        )

        # Основное тело документа
        if 'index' in location:
            para_index = location['index']
            if para_index is None or para_index >= len(doc.paragraphs):
                raise_paragraph_index_error(para_index, len(doc.paragraphs))
            logger.debug(f"Найден параграф в основном теле по индексу {para_index}")
            return doc.paragraphs[para_index], 'body'

        if 'paragraph_index' in location:
            para_index = location['paragraph_index']
            if para_index is None or para_index >= len(doc.paragraphs):
                raise_paragraph_index_error(para_index, len(doc.paragraphs))
            logger.debug(f"Найден параграф в основном теле по paragraph_index {para_index}")
            return doc.paragraphs[para_index], 'body'

        # Колонтитулы
        if 'section' in location and 'paragraph' in location:
            sect_idx = location['section']
            para_idx = location['paragraph']
            container = location.get('container')

            if container is None:
                container = 'header'
                logger.warning(f"Ключ 'container' отсутствует в location {location}, предполагаем 'header'")

            if sect_idx >= len(doc.sections):
                raise_header_footer_error(sect_idx, f"секция {sect_idx} вне диапазона (всего {len(doc.sections)})")

            section = doc.sections[sect_idx]
            if container == 'header':
                hf = section.header
                if hf is None:
                    raise_header_footer_error(sect_idx, "верхний колонтитул отсутствует")
                if para_idx >= len(hf.paragraphs):
                    raise_header_footer_error(sect_idx, f"параграф {para_idx} вне диапазона (всего {len(hf.paragraphs)})")
                logger.debug(f"Найден параграф в верхнем колонтитуле секции {sect_idx}, параграф {para_idx}")
                return hf.paragraphs[para_idx], 'header'
            elif container == 'footer':
                hf = section.footer
                if hf is None:
                    raise_header_footer_error(sect_idx, "нижний колонтитул отсутствует")
                if para_idx >= len(hf.paragraphs):
                    raise_header_footer_error(sect_idx, f"параграф {para_idx} вне диапазона (всего {len(hf.paragraphs)})")
                logger.debug(f"Найден параграф в нижнем колонтитуле секции {sect_idx}, параграф {para_idx}")
                return hf.paragraphs[para_idx], 'footer'
            else:
                raise FormatterError(ErrorKind.LOCATION_INVALID, detail=f"неизвестный контейнер '{container}'", location=location)

        raise_location_invalid(location)

    def _get_table(self, doc: Document, location: Dict[str, Any]) -> Table:
        """Находит таблицу в документе по location."""
        from src.core.formatter_errors import raise_table_index_error, FormatterError, ErrorKind

        if 'table_index' not in location:
            raise FormatterError(ErrorKind.LOCATION_INVALID, detail="отсутствует ключ table_index", location=location)
        table_idx = location['table_index']
        if table_idx is None or table_idx >= len(doc.tables):
            raise_table_index_error(table_idx, len(doc.tables))
        return doc.tables[table_idx]

    def _get_element(self, doc: Document, location: Dict[str, Any]) -> Tuple[Any, str]:
        """
        Находит элемент в документе по location с поддержкой всех типов.
        Возвращает кортеж (элемент, тип_элемента).
        """
        from src.core.formatter_errors import FormatterError, ErrorKind

        element_type = location.get('type', 'paragraph')

        if element_type == 'paragraph':
            para, container = self._get_paragraph(doc, location)
            return para, 'paragraph'
        elif element_type == 'table':
            table = self._get_table(doc, location)
            return table, 'table'
        elif element_type == 'figure':
            raise FormatterError(ErrorKind.LOCATION_INVALID, detail=f"тип '{element_type}' пока не поддерживается")
        else:
            raise FormatterError(ErrorKind.LOCATION_INVALID, detail=f"неизвестный тип '{element_type}'")

    def _clear_paragraph_direct_formatting(self, paragraph: Paragraph):
        """
        Очищает прямое форматирование в одном параграфе (шрифт, начертание, цвет и т.д.).
        Сохраняет стиль параграфа.
        """
        for run in paragraph.runs:
            run.font.bold = None
            run.font.italic = None
            run.font.underline = None
            run.font.strike = None
            run.font.color.rgb = None
            run.font.highlight_color = None
            # Шрифт и размер лучше не сбрасывать в None, если стиль их не задает явно,
            # но в нашей логике стиль уже задан.
            # run.font.name = None
            # run.font.size = None

    def clear_direct_formatting(self, doc_path: str, output_path: str):
        """
        Вспомогательный метод: удаляет ВСЕ прямое форматирование в документе, оставляя только стили.
        Полезно как радикальная мера перед применением стилей.
        """
        doc = open_document(doc_path)

        for para in doc.paragraphs:
            self._clear_paragraph_direct_formatting(para)

        doc.save(output_path)
        logger.info(f"Прямое форматирование очищено, файл сохранен: {output_path}")
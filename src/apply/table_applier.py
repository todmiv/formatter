"""
TableApplier - применение исправлений таблиц.
"""

import logging
from typing import Dict, Any
import docx
from docx.document import Document
from docx.table import Table
from docx.enum.style import WD_STYLE_TYPE

from src.core.audit_engine import AuditIssue
from src.core.config_loader import ConfigLoader
from .base_applier import BaseApplier

logger = logging.getLogger(__name__)


class TableApplier(BaseApplier):
    """
    Апплер для применения исправлений таблиц.
    """

    def apply(self, doc: Document, issue: AuditIssue) -> bool:
        """
        Применяет исправление таблицы к указанной таблице.
        Возвращает True в случае успеха, False при ошибке.
        """
        try:
            element, element_type = self._get_element(doc, issue.location)
            if element_type != 'table':
                logger.warning(f"TableApplier ожидает таблицу, получен {element_type}. Пропускаем.")
                return False

            table = element
            expected_style = issue.expected_value
            if expected_style and hasattr(table, 'style'):
                # Убеждаемся, что стиль существует в документе
                self._ensure_table_style(doc, expected_style)
                table.style = expected_style
                logger.debug(f"Стиль таблицы {expected_style} применен к таблице {issue.location}")
            else:
                logger.warning(f"Не удалось применить стиль таблицы {expected_style} к таблице {issue.location}")

            logger.debug(f"Исправление таблицы {issue.id} применено к {issue.location}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при применении исправления таблицы {issue.id}: {e}")
            return False

    def _ensure_table_style(self, doc: Document, style_name: str):
        """Создаёт стиль таблицы, если он отсутствует в документе."""
        try:
            doc.styles[style_name]
        except KeyError:
            # Стиль не найден — создаём
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement

            style = doc.styles.add_style(style_name, WD_STYLE_TYPE.TABLE)
            # Базовый стиль для Table Grid — это Table Normal
            try:
                base = doc.styles['Table Normal']
                style.base_style = base
            except KeyError:
                pass

            # Устанавливаем сетку (все границы)
            tblPr = style.element.tblPr
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                style.element.insert(0, tblPr)

            borders = OxmlElement('w:tblBorders')
            for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'single')
                border.set(qn('w:sz'), '4')
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), 'auto')
                borders.append(border)
            tblPr.append(borders)

            logger.debug(f"Стиль таблицы '{style_name}' создан в документе")

    def apply_style(self, table: Table, style_name: str) -> bool:
        """
        Применяет стиль к таблице по имени.
        """
        try:
            table.style = style_name
            return True
        except Exception as e:
            logger.error(f"Не удалось применить стиль таблицы {style_name}: {e}")
            return False
"""
HeaderFooterApplier - применение исправлений колонтитулов.
"""

import logging
from typing import Dict, Any, List
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph

from src.core.audit_engine import AuditIssue
from src.core.config_loader import ConfigLoader
from src.core.page_manager import PageManager
from .base_applier import BaseApplier

logger = logging.getLogger(__name__)


class HeaderFooterApplier(BaseApplier):
    """
    Апплер для применения исправлений колонтитулов.
    """

    def __init__(self, config_loader: ConfigLoader):
        super().__init__(config_loader)
        self.page_manager = PageManager(config_loader)

    def apply(self, doc: Document, issue: AuditIssue) -> bool:
        """
        Применяет исправление колонтитула к указанному элементу.
        Возвращает True в случае успеха, False при ошибке.
        """
        try:
            element, element_type = self._get_element(doc, issue.location)
            if element_type != 'paragraph':
                logger.warning(f"HeaderFooterApplier ожидает параграф, получен {element_type}. Пропускаем.")
                return False

            paragraph, container = self._get_paragraph(doc, issue.location)
            target_style_name = issue.element_type
            style_config = self.config.get_style_config(target_style_name)
            if not style_config:
                logger.warning(f"Конфигурация для стиля '{target_style_name}' не найдена.")
                return False

            # Применяем стиль к параграфу колонтитула
            try:
                paragraph.style = target_style_name
                logger.debug(f"Стиль {target_style_name} применен к {container} {issue.location}")
            except Exception as e:
                logger.warning(f"Не удалось применить стиль {target_style_name}: {e}")

            # Применяем прямые переопределения
            self.style_manager.apply_direct_overrides(paragraph, style_config)

            logger.debug(f"Исправление колонтитула {issue.id} применено к {container} {issue.location}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при применении исправления колонтитула {issue.id}: {e}")
            return False

    def apply_page_settings(self, doc: Document):
        """
        Применяет настройки страницы (поля, колонтитулы) через PageManager.
        """
        try:
            self.page_manager.apply_all(doc)
            logger.debug("Настройки страницы применены")
            return True
        except Exception as e:
            logger.error(f"Ошибка при применении настроек страницы: {e}")
            return False
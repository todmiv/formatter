"""
FontApplier - применение исправлений шрифтов.
"""

import logging
from typing import Dict, Any, List
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph

from src.core.audit_engine import AuditIssue
from src.core.config_loader import ConfigLoader
from .base_applier import BaseApplier

logger = logging.getLogger(__name__)


class FontApplier(BaseApplier):
    """
    Апплер для применения исправлений шрифтов.
    """

    def apply(self, doc: Document, issue: AuditIssue) -> bool:
        """
        Применяет исправление шрифта к указанному элементу.
        Возвращает True в случае успеха, False при ошибке.
        """
        try:
            element, element_type = self._get_element(doc, issue.location)
            if element_type != 'paragraph':
                logger.warning(f"FontApplier ожидает параграф, получен {element_type}. Пропускаем.")
                return False

            paragraph = element
            target_style_name = issue.element_type
            style_config = self.config.get_style_config(target_style_name)
            if not style_config:
                logger.warning(f"Конфигурация для стиля '{target_style_name}' не найдена.")
                return False

            # Применяем настройки шрифта из конфига
            if 'font' in style_config:
                font_cfg = style_config['font']
                self._apply_font_config(paragraph, font_cfg)

            logger.debug(f"Исправление шрифта {issue.id} применено к {issue.location}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при применении исправления шрифта {issue.id}: {e}")
            return False

    def _apply_font_config(self, paragraph: Paragraph, font_cfg: Dict[str, Any]):
        """
        Применяет конфигурацию шрифта ко всем runs параграфа.
        """
        for run in paragraph.runs:
            if 'name' in font_cfg:
                run.font.name = font_cfg['name']
            if 'size_pt' in font_cfg:
                run.font.size = docx.shared.Pt(font_cfg['size_pt'])
            if 'bold' in font_cfg:
                run.font.bold = font_cfg['bold']
            if 'italic' in font_cfg:
                run.font.italic = font_cfg['italic']
            if 'underline' in font_cfg:
                run.font.underline = font_cfg['underline']
            if 'color' in font_cfg:
                # Ожидается шестнадцатеричное значение, например '#000000'
                from docx.shared import RGBColor
                hex_color = font_cfg['color'].lstrip('#')
                if len(hex_color) == 6:
                    r = int(hex_color[0:2], 16)
                    g = int(hex_color[2:4], 16)
                    b = int(hex_color[4:6], 16)
                    run.font.color.rgb = RGBColor(r, g, b)
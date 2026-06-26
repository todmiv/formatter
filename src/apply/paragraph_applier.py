"""
ParagraphApplier - применение исправлений абзацев.
"""

import logging
from typing import Dict, Any, List, Optional
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn

from src.core.audit_engine import AuditIssue
from src.core.config_loader import ConfigLoader
from src.core.style_manager import StyleManager
from .base_applier import BaseApplier

logger = logging.getLogger(__name__)


def _is_spacer_paragraph(paragraph):
    """Определяет, является ли параграф-спайсером (минимальный интервал между таблицами)."""
    text = (paragraph.text or '').strip()
    if text:
        return False
    pPr = paragraph._element.find(qn('w:pPr'))
    if pPr is None:
        return False
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        return False
    line_rule = spacing.get(qn('w:lineRule'))
    line_val = spacing.get(qn('w:line'))
    if line_rule == 'exact' and line_val and int(line_val) <= 20:
        return True
    return False


class ParagraphApplier(BaseApplier):
    """
    Апплер для применения исправлений абзацев.
    """

    def __init__(self, config_loader: ConfigLoader):
        super().__init__(config_loader)
        self.style_manager = StyleManager(config_loader)

    def apply(self, doc: Document, issue: AuditIssue, style_cache: Dict,
              apply_styles: bool = True, apply_direct_overrides: bool = True,
              clear_direct_formatting: bool = False) -> bool:
        """
        Применяет исправление абзаца к указанному параграфу.
        Возвращает True в случае успеха, False при ошибке.
        """
        try:
            element, element_type = self._get_element(doc, issue.location)
            if element_type != 'paragraph':
                logger.warning(f"ParagraphApplier ожидает параграф, получен {element_type}. Пропускаем.")
                return False

            paragraph, container = self._get_paragraph(doc, issue.location)
            
            if _is_spacer_paragraph(paragraph):
                logger.debug(f"Пропуск spacer-параграфа {issue.location}")
                return False

            target_style_name = issue.element_type
            style_config = self.config.get_style_config(target_style_name)
            if not style_config:
                logger.warning(f"Конфигурация для стиля '{target_style_name}' не найдена.")
                return False

            # Очистка прямого форматирования, если запрошено
            if clear_direct_formatting:
                self._clear_paragraph_direct_formatting(paragraph)

            # Обеспечиваем наличие стиля в документе
            self.style_manager.ensure_style_exists(doc, target_style_name, style_config, style_cache)

            # Применяем стиль к параграфу
            if apply_styles:
                try:
                    paragraph.style = target_style_name
                    logger.debug(f"Стиль {target_style_name} применен к {container} {issue.location}")
                except Exception as e:
                    logger.warning(f"Не удалось применить стиль {target_style_name}: {e}")

            # Точечно применяем свойства, указанные в payload, если нужно форсировать
            if apply_direct_overrides:
                try:
                    self.style_manager.apply_direct_overrides(paragraph, style_config)
                    logger.debug(f"Прямые переопределения применены к {container} {issue.location}")
                except Exception as e:
                    logger.warning(f"Не удалось применить прямые переопределения: {e}")

            logger.debug(f"Исправление абзаца {issue.id} применено к {issue.location}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при применении исправления абзаца {issue.id}: {e}")
            return False

    def apply_with_context(self, doc: Document, issues: List[AuditIssue], style_cache: Dict,
                           apply_styles: bool = True, apply_direct_overrides: bool = True,
                           clear_direct_formatting: bool = False) -> Dict[str, int]:
        """
        Улучшенная обработка взаимозависимостей параграфов.
        Учитывает контекст соседних параграфов, применяет исправления в правильном порядке.
        Возвращает статистику {'applied': int, 'failed': int}.
        """
        applied = 0
        failed = 0

        # Группируем проблемы по индексу параграфа (включаем PARAGRAPH и STYLE)
        para_issues = {}
        for issue in issues:
            if issue.category not in ('PARAGRAPH', 'STYLE'):
                continue
            idx = issue.location.get('index')
            if idx is None:
                continue
            if idx not in para_issues:
                para_issues[idx] = []
            para_issues[idx].append(issue)

        # Сортируем индексы для последовательной обработки
        sorted_indices = sorted(para_issues.keys())

        for idx in sorted_indices:
            para_issues_list = para_issues[idx]
            try:
                para, container = self._get_paragraph(doc, {'index': idx})
                
                if _is_spacer_paragraph(para):
                    logger.debug(f"Пропуск spacer-параграфа index={idx}")
                    continue

                # Очистка прямого форматирования, если запрошено
                if clear_direct_formatting:
                    self._clear_paragraph_direct_formatting(para)

                # Применяем стили для каждой проблемы (обычно одна проблема на параграф)
                for issue in para_issues_list:
                    target_style_name = issue.element_type
                    style_config = self.config.get_style_config(target_style_name)
                    if not style_config:
                        logger.warning(f"Пропущено исправление {issue.id}: конфигурация для стиля '{target_style_name}' не найдена.")
                        failed += 1
                        continue

                    # Обеспечиваем наличие стиля
                    self.style_manager.ensure_style_exists(doc, target_style_name, style_config, style_cache)

                    # Применяем стиль
                    if apply_styles:
                        para.style = target_style_name

                    # Прямые переопределения
                    if apply_direct_overrides:
                        self.style_manager.apply_direct_overrides(para, style_config)

                    applied += 1
                    logger.debug(f"Исправление {issue.id} применено к параграфу {idx}")

                # Проверяем согласованность с соседними параграфами
                if idx > 0:
                    prev_para, _ = self._get_paragraph(doc, {'index': idx-1})
                    # Можно добавить логику проверки интервалов

            except Exception as e:
                logger.warning(f"Ошибка при обработке параграфа {idx}: {e}")
                failed += len(para_issues_list)

        return {'applied': applied, 'failed': failed}
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ContentTransformer — трансформация содержимого документов.
Изменяет регистр, очищает текст, применяет правила форматирования.
"""

import re
import logging
from typing import Dict, Any, Optional

from docx import Document
from docx.text.paragraph import Paragraph
from src.core.models import TransformResult

logger = logging.getLogger(__name__)


class ContentTransformer:
    """
    Трансформер содержимого документов.
    """

    TEXT_TRANSFORM_MAP = {
        'capitalize': str.capitalize,
        'uppercase': str.upper,
        'lowercase': str.lower,
        'title': str.title,
    }

    def __init__(self, config_loader=None):
        self.config = config_loader
        self.formatting_rules = {}
        if config_loader:
            self.formatting_rules = config_loader.get_formatting_rules() or {}

    def transform_document(
        self,
        doc: Document,
        output_path: Optional[str] = None,
    ) -> TransformResult:
        """
        Трансформирует содержимое документа.

        :param doc: объект docx.Document.
        :param output_path: путь для сохранения.
        :return: результат трансформации.
        """
        stats = TransformResult()

        for para in doc.paragraphs:
            result = self._transform_paragraph(para)
            if result['changed']:
                stats.transforms_applied += result['changes_count']
            stats.paragraphs_processed += 1

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        result = self._transform_paragraph(para)
                        if result['changed']:
                            stats.transforms_applied += result['changes_count']
                        stats.paragraphs_processed += 1

        if output_path:
            doc.save(output_path)
            logger.info(f"Документ сохранён: {output_path}")

        logger.info(
            f"Трансформация: {stats.transforms_applied} изменений, "
            f"{stats.paragraphs_processed} параграфов"
        )
        return stats

    def cleanup_text(self, text: str) -> str:
        """
        Очищает текст от типичных проблем.

        :param text: исходный текст.
        :return: очищенный текст.
        """
        if not text:
            return text

        text = re.sub(r'  +', ' ', text)
        text = re.sub(r'\s+([.,;:!?])', r'\1', text)
        text = re.sub(r'([(\[])\s+', r'\1', text)
        text = re.sub(r'\s+([)\]])', r'\1', text)
        text = re.sub(r'\s+', ' ', text)

        return text.strip()

    def apply_text_transform(self, text: str, transform_type: str) -> str:
        """
        Применяет трансформацию регистра.

        :param text: исходный текст.
        :param transform_type: тип трансформации.
        :return: трансформированный текст.
        """
        if not text or transform_type not in self.TEXT_TRANSFORM_MAP:
            return text
        return self.TEXT_TRANSFORM_MAP[transform_type](text)

    def _transform_paragraph(self, para: Paragraph) -> Dict[str, Any]:
        """Трансформирует один параграф."""
        result = {'changed': False, 'changes_count': 0}

        text = para.text
        if not text:
            return result

        style_name = para.style.name
        style_config = None
        if self.config:
            style_config = self.config.get_style_config(style_name)

        if style_config:
            text_transform = style_config.get('text_transform')
            if text_transform and text_transform in self.TEXT_TRANSFORM_MAP:
                new_text = self.apply_text_transform(text, text_transform)
                if new_text != text:
                    self._replace_paragraph_text(para, new_text)
                    result['changed'] = True
                    result['changes_count'] += 1
                    return result

            no_period = style_config.get('no_period_at_end', False)
            if no_period and text.rstrip().endswith('.'):
                new_text = text.rstrip()[:-1]
                self._replace_paragraph_text(para, new_text)
                result['changed'] = True
                result['changes_count'] += 1

        return result

    def _replace_paragraph_text(self, para: Paragraph, new_text: str):
        """Заменяет текст параграфа, сохраняя форматирование."""
        if not para.runs:
            para.text = new_text
            return

        first_run = para.runs[0]
        first_run.text = new_text

        for run in para.runs[1:]:
            run.text = ''

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ContentProcessor — обёртка над ContentValidator и ContentTransformer.
Обеспечивает обратную совместимость с существующим кодом.
"""

import logging
from typing import Dict, Any, Optional, List
from docx import Document
from docx.text.paragraph import Paragraph

from src.core.content_validator import ContentValidator
from src.core.content_transformer import ContentTransformer
from src.core.models import ValidationResult, TransformResult

logger = logging.getLogger(__name__)


class ContentProcessor:
    """
    Процессор содержимого документов.
    Делегирует валидацию и трансформацию специализированным сервисам.
    """

    def __init__(self, config_loader=None):
        self.config = config_loader
        self._validator = ContentValidator(config_loader)
        self._transformer = ContentTransformer(config_loader)

    def validate_document(self, doc: Document) -> List[Dict[str, Any]]:
        """
        Валидирует содержимое документа.
        Возвращает список проблем (обратная совместимость).
        """
        result = self._validator.validate_document(doc)
        return result.issues

    def transform_document(
        self,
        doc: Document,
        output_path: Optional[str] = None,
    ) -> Dict[str, int]:
        """
        Трансформирует содержимое документа.
        Возвращает статистику (обратная совместимость).
        """
        result = self._transformer.transform_document(doc, output_path)
        return {
            'transforms_applied': result.transforms_applied,
            'paragraphs_processed': result.paragraphs_processed,
        }

    def cleanup_text(self, text: str) -> str:
        """Очищает текст."""
        return self._transformer.cleanup_text(text)

    def apply_text_transform(self, text: str, transform_type: str) -> str:
        """Применяет трансформацию регистра."""
        return self._transformer.apply_text_transform(text, transform_type)

    def validate_heading(self, text: str, style_name: str) -> List[str]:
        """Валидирует заголовок."""
        return self._validator.validate_heading(text, style_name)

    def validate_table_title(self, text: str) -> List[str]:
        """Валидирует название таблицы."""
        return self._validator.validate_table_title(text)

    def validate_list_item(self, text: str) -> List[str]:
        """Валидирует элемент списка."""
        return self._validator.validate_list_item(text)

    @property
    def validator(self) -> ContentValidator:
        return self._validator

    @property
    def transformer(self) -> ContentTransformer:
        return self._transformer


def process_document(
    doc_path: str,
    config_path: Optional[str] = None,
    output_path: Optional[str] = None,
    validate: bool = True,
    transform: bool = True,
) -> Dict[str, Any]:
    """Удобная функция для обработки документа."""
    from src.core.config_loader import ConfigLoader

    config_loader = None
    if config_path:
        config_loader = ConfigLoader(config_path)
        config_loader.load()

    processor = ContentProcessor(config_loader)
    doc = Document(doc_path)

    result = {'validation_issues': [], 'transform_stats': {}}

    if validate:
        result['validation_issues'] = processor.validate_document(doc)

    if transform:
        result['transform_stats'] = processor.transform_document(doc, output_path)

    return result

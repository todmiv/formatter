#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ContentValidator — валидация содержимого документов.
Проверяет текст по правилам форматирования.
"""

import re
import logging
from typing import Dict, Any, List

from docx import Document
from src.core.models import ValidationResult

logger = logging.getLogger(__name__)


class ContentValidator:
    """
    Валидатор содержимого документов.
    Проверяет заголовки, подписи, списки, нумерацию.
    """

    def __init__(self, config_loader=None):
        self.config = config_loader
        self.formatting_rules = {}
        if config_loader:
            self.formatting_rules = config_loader.get_formatting_rules() or {}

    def validate_document(self, doc: Document) -> ValidationResult:
        """
        Валидирует содержимое документа.

        :param doc: объект docx.Document.
        :return: результат валидации.
        """
        issues = []
        issues.extend(self._validate_paragraphs(doc))
        issues.extend(self._validate_tables(doc))
        issues.extend(self._validate_numbering(doc))

        logger.info(f"Валидация завершена: {len(issues)} проблем")
        return ValidationResult(issues=issues)

    def validate_heading(self, text: str, style_name: str) -> List[str]:
        """Валидирует заголовок."""
        errors = []
        rules = self.formatting_rules.get('headings', {})

        if rules.get('no_period_at_end', False) and text.rstrip().endswith('.'):
            errors.append("В конце заголовка не должна ставиться точка")

        return errors

    def validate_table_title(self, text: str) -> List[str]:
        """Валидирует название таблицы."""
        errors = []
        tables_rules = self.formatting_rules.get('tables', {})
        title_rules = tables_rules.get('title', {})

        if title_rules.get('no_period_at_end', False) and text.rstrip().endswith('.'):
            errors.append("В конце названия таблицы не ставится точка")

        if title_rules.get('no_abbreviations', False) and self._contains_abbreviation(text):
            errors.append("В названии таблицы не допускаются сокращения")

        return errors

    def validate_list_item(self, text: str) -> List[str]:
        """Валидирует элемент списка."""
        errors = []
        text = text.strip()
        if text and text.endswith('.'):
            errors.append("Элемент списка не должен заканчиваться точкой")
        return errors

    def _validate_paragraphs(self, doc: Document) -> List[Dict[str, Any]]:
        """Валидирует параграфы документа."""
        issues = []
        math_symbols = self.formatting_rules.get('math_symbols_require_numbers', [])

        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue

            style_name = para.style.name

            if style_name.startswith('Heading'):
                for error in self.validate_heading(text, style_name):
                    issues.append({
                        'type': 'HEADING_FORMAT',
                        'location': {'index': idx},
                        'message': error,
                        'severity': 'WARNING',
                    })

            if self._is_caption(text):
                for error in self.validate_table_title(text):
                    issues.append({
                        'type': 'CAPTION_FORMAT',
                        'location': {'index': idx},
                        'message': error,
                        'severity': 'WARNING',
                    })

            if self._is_list_item(text):
                for error in self.validate_list_item(text):
                    issues.append({
                        'type': 'LIST_FORMAT',
                        'location': {'index': idx},
                        'message': error,
                        'severity': 'INFO',
                    })

            if math_symbols:
                for issue in self._validate_math_symbols(text, math_symbols, idx):
                    issues.append(issue)

        return issues

    def _validate_math_symbols(self, text: str, symbols: List[str], idx: int) -> List[Dict[str, Any]]:
        """Валидирует использование математических знаков."""
        issues = []
        for symbol in symbols:
            pattern = rf'(?<!\d){re.escape(symbol)}(?!\d)'
            for match in re.finditer(pattern, text):
                issues.append({
                    'type': 'MATH_SYMBOL',
                    'location': {'index': idx, 'column': match.start()},
                    'message': f"Мат. знак '{symbol}' должен использоваться с числами",
                    'severity': 'WARNING',
                })
        return issues

    def _validate_tables(self, doc: Document) -> List[Dict[str, Any]]:
        """Валидирует содержимое таблиц."""
        return []

    def _validate_numbering(self, doc: Document) -> List[Dict[str, Any]]:
        """Валидирует нумерацию таблиц и рисунков."""
        issues = []
        tables_rules = self.formatting_rules.get('tables', {})
        numbering_rules = tables_rules.get('numbering', {})

        if not numbering_rules.get('enabled', True):
            return issues

        table_pattern = re.compile(r'^Таблица\s+(\d+\.?\d*)')
        found_tables = []

        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            table_match = table_pattern.match(text)
            if table_match:
                found_tables.append({'index': idx, 'number': table_match.group(1)})

        if len(found_tables) > 1:
            numbers = [t['number'] for t in found_tables]
            if not self._is_sequential_numbering(numbers):
                issues.append({
                    'type': 'TABLE_NUMBERING',
                    'location': {'index': found_tables[-1]['index']},
                    'message': f"Нарушена нумерация таблиц: {', '.join(numbers)}",
                    'severity': 'WARNING',
                })

        return issues

    def _is_sequential_numbering(self, numbers: List[str]) -> bool:
        """Проверяет последовательность нумерации."""
        try:
            parsed = [[int(p) for p in num.split('.')] for num in numbers]
            return all(
                parsed[i] == parsed[i-1][:-1] + [parsed[i-1][-1] + 1]
                for i in range(1, len(parsed))
            )
        except (ValueError, IndexError):
            return True

    def _is_caption(self, text: str) -> bool:
        """Определяет, является ли текст подписью."""
        patterns = [r'^Таблица\s+\d+', r'^Рисунок\s+\d+', r'^Примечание']
        return any(re.match(p, text.strip()) for p in patterns)

    def _is_list_item(self, text: str) -> bool:
        """Определяет, является ли текст элементом списка."""
        patterns = [r'^\s*[–−-]\s+', r'^\s*\d+[\.\)]\s+', r'^\s*[а-яa-z][\.\)]\s+']
        return any(re.match(p, text) for p in patterns)

    def _contains_abbreviation(self, text: str) -> bool:
        """Проверяет наличие сокращений."""
        abbreviations = [r'\bт\.\s*д\.', r'\bт\.\s*е\.', r'\bт\.\s*к\.', r'\bт\.\s*н\.']
        return any(re.search(p, text, re.IGNORECASE) for p in abbreviations)

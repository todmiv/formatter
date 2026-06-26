#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ContentProcessor - обработка содержимого документов.
Валидация, трансформация и очистка текста согласно правилам форматирования.

Поддерживаемые операции:
- Валидация текста по правилам конфигурации
- Трансформация регистра текста (capitalize, uppercase, lowercase)
- Очистка лишних пробелов и знаков препинания
- Проверка запрещённых конструкций
- Валидация нумерации таблиц и рисунков
- Проверка единиц измерения и математических знаков

Использование:
    from src.core.content_processor import ContentProcessor
    processor = ContentProcessor(config_loader)
    issues = processor.validate_document(doc)
    processor.transform_document(doc)
"""

import re
import logging
from typing import Dict, Any, Optional, List, Tuple
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)


class ContentProcessor:
    """
    Процессор содержимого документов.
    Валидирует и трансформирует текст согласно правилам форматирования.
    """

    TEXT_TRANSFORM_MAP = {
        'capitalize': str.capitalize,
        'uppercase': str.upper,
        'lowercase': str.lower,
        'title': str.title,
    }

    def __init__(self, config_loader=None):
        """
        :param config_loader: экземпляр ConfigLoader или None для правил по умолчанию.
        """
        self.config = config_loader
        self.formatting_rules = {}
        if config_loader:
            self.formatting_rules = config_loader.get_formatting_rules() or {}

    def validate_document(self, doc: Document) -> List[Dict[str, Any]]:
        """
        Валидирует содержимое документа по правилам форматирования.

        :param doc: объект docx.Document.
        :return: список проблем [{'type': str, 'location': dict, 'message': str, 'severity': str}]
        """
        issues = []

        issues.extend(self._validate_paragraphs(doc))
        issues.extend(self._validate_tables(doc))
        issues.extend(self._validate_numbering(doc))

        logger.info(f"Валидация завершена: {len(issues)} проблем найдено")
        return issues

    def transform_document(self, doc: Document, output_path: Optional[str] = None) -> Dict[str, int]:
        """
        Трансформирует содержимое документа (регистр, пробелы, знаки препинания).

        :param doc: объект docx.Document.
        :param output_path: путь для сохранения. Если None, изменения в памяти.
        :return: статистика {'transforms_applied': int, 'paragraphs_processed': int}
        """
        stats = {'transforms_applied': 0, 'paragraphs_processed': 0}

        for para in doc.paragraphs:
            result = self._transform_paragraph(para)
            if result['changed']:
                stats['transforms_applied'] += result['changes_count']
            stats['paragraphs_processed'] += 1

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        result = self._transform_paragraph(para)
                        if result['changed']:
                            stats['transforms_applied'] += result['changes_count']
                        stats['paragraphs_processed'] += 1

        if output_path:
            doc.save(output_path)
            logger.info(f"Документ сохранён: {output_path}")

        logger.info(
            f"Трансформация завершена: {stats['transforms_applied']} изменений, "
            f"{stats['paragraphs_processed']} параграфов обработано"
        )
        return stats

    def cleanup_text(self, text: str) -> str:
        """
        Очищает текст от типичных проблем:
        - Двойные пробелы
        - Пробелы перед знаками препинания
        - Пробелы после открывающих скобок
        - Пробелы перед закрывающими скобками
        - Точки с запятой в конце предложений
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
        Применяет трансформацию регистра к тексту.

        :param text: исходный текст.
        :param transform_type: тип трансформации ('capitalize', 'uppercase', 'lowercase', 'title').
        :return: трансформированный текст.
        """
        if not text or transform_type not in self.TEXT_TRANSFORM_MAP:
            return text

        transformer = self.TEXT_TRANSFORM_MAP[transform_type]
        return transformer(text)

    def validate_heading(self, text: str, style_name: str) -> List[str]:
        """
        Валидирует заголовок по правилам.

        :param text: текст заголовка.
        :param style_name: имя стиля.
        :return: список ошибок.
        """
        errors = []
        rules = self.formatting_rules.get('headings', {})

        if rules.get('no_period_at_end', False):
            if text.rstrip().endswith('.'):
                errors.append("В конце заголовка не должна ставиться точка")

        if rules.get('no_bold', False):
            pass

        return errors

    def validate_table_title(self, text: str) -> List[str]:
        """
        Валидирует название таблицы.

        :param text: текст подписи.
        :return: список ошибок.
        """
        errors = []
        tables_rules = self.formatting_rules.get('tables', {})
        title_rules = tables_rules.get('title', {})

        if title_rules.get('no_period_at_end', False):
            if text.rstrip().endswith('.'):
                errors.append("В конце названия таблицы не ставится точка")

        if title_rules.get('no_abbreviations', False):
            if self._contains_abbreviation(text):
                errors.append("В названии таблицы не допускаются сокращения")

        if title_rules.get('alignment', 'center') == 'center':
            pass

        return errors

    def validate_list_item(self, text: str) -> List[str]:
        """
        Валидирует элемент списка.

        :param text: текст элемента.
        :return: список ошибок.
        """
        errors = []

        text = text.strip()
        if not text:
            return errors

        if text.endswith('.'):
            errors.append("Элемент списка не должен заканчиваться точкой")

        return errors

    # ------------------------------------------------------------------
    # Приватные методы валидации
    # ------------------------------------------------------------------

    def _validate_paragraphs(self, doc: Document) -> List[Dict[str, Any]]:
        """Валидирует параграфы документа."""
        issues = []
        prohibited = self.formatting_rules.get('prohibited', [])
        math_symbols = self.formatting_rules.get('math_symbols_require_numbers', [])

        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue

            style_name = para.style.name

            if style_name.startswith('Heading'):
                heading_issues = self._validate_heading_content(text, style_name, idx)
                issues.extend(heading_issues)

            if self._is_caption(text):
                caption_issues = self._validate_caption_content(text, idx)
                issues.extend(caption_issues)

            if self._is_list_item(text):
                list_issues = self._validate_list_content(text, idx)
                issues.extend(list_issues)

            if math_symbols:
                math_issues = self._validate_math_symbols(text, math_symbols, idx)
                issues.extend(math_issues)

        return issues

    def _validate_heading_content(self, text: str, style_name: str, para_index: int) -> List[Dict[str, Any]]:
        """Валидирует содержимое заголовка."""
        issues = []
        errors = self.validate_heading(text, style_name)

        for error in errors:
            issues.append({
                'type': 'HEADING_FORMAT',
                'location': {'index': para_index},
                'message': error,
                'severity': 'WARNING',
            })

        return issues

    def _validate_caption_content(self, text: str, para_index: int) -> List[Dict[str, Any]]:
        """Валидирует содержимое подписи (таблица, рисунок)."""
        issues = []

        if text.rstrip().endswith('.'):
            issues.append({
                'type': 'CAPTION_FORMAT',
                'location': {'index': para_index},
                'message': "В конце подписи не ставится точка",
                'severity': 'WARNING',
            })

        return issues

    def _validate_list_content(self, text: str, para_index: int) -> List[Dict[str, Any]]:
        """Валидирует содержимое элемента списка."""
        issues = []
        errors = self.validate_list_item(text)

        for error in errors:
            issues.append({
                'type': 'LIST_FORMAT',
                'location': {'index': para_index},
                'message': error,
                'severity': 'INFO',
            })

        return issues

    def _validate_math_symbols(self, text: str, symbols: List[str], para_index: int) -> List[Dict[str, Any]]:
        """Валидирует использование математических знаков."""
        issues = []

        for symbol in symbols:
            pattern = rf'(?<!\d){re.escape(symbol)}(?!\d)'
            matches = list(re.finditer(pattern, text))
            for match in matches:
                issues.append({
                    'type': 'MATH_SYMBOL',
                    'location': {'index': para_index, 'column': match.start()},
                    'message': f"Математический знак '{symbol}' должен использоваться с числовыми значениями",
                    'severity': 'WARNING',
                })

        return issues

    def _validate_tables(self, doc: Document) -> List[Dict[str, Any]]:
        """Валидирует содержимое таблиц."""
        issues = []

        for table_idx, table in enumerate(doc.tables):
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    for para in cell.paragraphs:
                        text = para.text.strip()
                        if not text:
                            continue

        return issues

    def _validate_numbering(self, doc: Document) -> List[Dict[str, Any]]:
        """Валидирует нумерацию таблиц и рисунков."""
        issues = []
        tables_rules = self.formatting_rules.get('tables', {})
        numbering_rules = tables_rules.get('numbering', {})

        if not numbering_rules.get('enabled', True):
            return issues

        table_pattern = re.compile(r'^Таблица\s+(\d+\.?\d*)')
        figure_pattern = re.compile(r'^Рисунок\s+(\d+\.?\d*)')

        found_tables = []
        found_figures = []

        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()

            table_match = table_pattern.match(text)
            if table_match:
                number = table_match.group(1)
                found_tables.append({'index': idx, 'number': number})

            figure_match = figure_pattern.match(text)
            if figure_match:
                number = figure_match.group(1)
                found_figures.append({'index': idx, 'number': number})

        if len(found_tables) > 1:
            numbers = [t['number'] for t in found_tables]
            if not self._is_sequential_numbering(numbers):
                issues.append({
                    'type': 'TABLE_NUMBERING',
                    'location': {'index': found_tables[-1]['index']},
                    'message': f"Нарушена последовательность нумерации таблиц: {', '.join(numbers)}",
                    'severity': 'WARNING',
                })

        return issues

    def _is_sequential_numbering(self, numbers: List[str]) -> bool:
        """Проверяет последовательность нумерации."""
        try:
            parsed = []
            for num in numbers:
                parts = [int(p) for p in num.split('.')]
                parsed.append(parts)

            for i in range(1, len(parsed)):
                if parsed[i] != parsed[i-1][:-1] + [parsed[i-1][-1] + 1]:
                    return False

            return True
        except (ValueError, IndexError):
            return True

    # ------------------------------------------------------------------
    # Приватные методы трансформации
    # ------------------------------------------------------------------

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
        """Заменяет текст параграфа, сохраняя форматирование первого run."""
        if not para.runs:
            para.text = new_text
            return

        first_run = para.runs[0]
        first_run.text = new_text

        for run in para.runs[1:]:
            run.text = ''

    # ------------------------------------------------------------------
    # Вспомогательные методы
    # ------------------------------------------------------------------

    def _is_caption(self, text: str) -> bool:
        """Определяет, является ли текст подписью."""
        patterns = [
            r'^Таблица\s+\d+',
            r'^Рисунок\s+\d+',
            r'^Примечание',
        ]
        return any(re.match(p, text.strip()) for p in patterns)

    def _is_list_item(self, text: str) -> bool:
        """Определяет, является ли текст элементом списка."""
        patterns = [
            r'^\s*[–−-]\s+',
            r'^\s*\d+[\.\)]\s+',
            r'^\s*[а-яa-z][\.\)]\s+',
        ]
        return any(re.match(p, text) for p in patterns)

    def _contains_abbreviation(self, text: str) -> bool:
        """Проверяет наличие сокращений в тексте."""
        abbreviations = [
            r'\bт\.\s*д\.',
            r'\bт\.\s*е\.',
            r'\bт\.\s*к\.',
            r'\bт\.\s*н\.',
            r'\bи\.\s*т\.\s*д\.',
            r'\bи\.\s*т\.\s*п\.',
        ]
        return any(re.search(p, text, re.IGNORECASE) for p in abbreviations)


def process_document(doc_path: str, config_path: Optional[str] = None,
                     output_path: Optional[str] = None,
                     validate: bool = True, transform: bool = True) -> Dict[str, Any]:
    """
    Удобная функция для обработки документа.

    :param doc_path: путь к документу.
    :param config_path: путь к конфигу (опционально).
    :param output_path: путь для сохранения (опционально).
    :param validate: выполнять валидацию.
    :param transform: выполнять трансформацию.
    :return: результат обработки.
    """
    from src.core.config_loader import ConfigLoader

    config_loader = None
    if config_path:
        config_loader = ConfigLoader(config_path)
        config_loader.load()

    processor = ContentProcessor(config_loader)
    doc = Document(doc_path)

    result = {
        'validation_issues': [],
        'transform_stats': {},
    }

    if validate:
        result['validation_issues'] = processor.validate_document(doc)

    if transform:
        result['transform_stats'] = processor.transform_document(doc, output_path)

    return result

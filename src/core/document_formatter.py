#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DocumentFormatter — высокоуровневый API для форматирования документов.
Делегирует работу сервисам: FormatterService, BatchService, DiffService.

Использование:
    from src.core.document_formatter import DocumentFormatter

    formatter = DocumentFormatter(config_path='configs/active/config.yaml')
    result = formatter.format_document('input.docx', 'output.docx')
"""

import os
import logging
import argparse
from typing import Optional

from src.core.formatter_service import FormatterService
from src.core.batch_service import BatchService
from src.core.diff_service import DiffService
from src.core.models import FormatResult, AuditFormatResult, BatchResult, DiffResult

logger = logging.getLogger(__name__)


class DocumentFormatter:
    """
    Высокоуровневый форматировщик документов.
    Делегирует работу специализированным сервисам.
    """

    def __init__(
        self,
        config_path: Optional[str] = None,
        template_path: Optional[str] = None,
        verbose: bool = False,
    ):
        self.config_path = config_path
        self.template_path = template_path
        self.verbose = verbose

        self._formatter = FormatterService(config_path, template_path)
        self._batch = BatchService(config_path, template_path)
        self._diff = DiffService()

    @property
    def mode(self) -> str:
        return self._formatter.mode

    def format_document(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        options: Optional[dict] = None,
    ) -> FormatResult:
        """Форматирует документ."""
        options = options or {}
        return self._formatter.format_document(
            input_path,
            output_path,
            clear_formatting=options.get('clear_direct_formatting', False),
            process_content=options.get('process_content', True),
        )

    def format_with_audit(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        max_iterations: int = 3,
    ) -> AuditFormatResult:
        """Полный цикл: аудит + исправление."""
        if not self.config_loader:
            from src.core.formatter_errors import FormatterError, ErrorKind
            raise FormatterError(
                ErrorKind.CONFIG_NOT_FOUND,
                detail="format_with_audit требует config_path"
            )

        from src.core.audit_engine import AuditEngine
        from src.apply.apply_orchestrator import ApplyOrchestrator

        t_start = __import__('time').perf_counter()
        output = output_path or self._formatter._default_output_path(input_path)

        result = AuditFormatResult(
            input=input_path,
            output=output,
            mode=self.mode,
            success=False,
        )

        try:
            current_path = input_path
            audit_engine = AuditEngine(self.config_loader)

            for iteration in range(max_iterations):
                t_iter = __import__('time').perf_counter()

                issues = audit_engine.scan_document(current_path)
                fixable = [i for i in issues if i.auto_fixable]

                if not fixable:
                    break

                orchestrator = ApplyOrchestrator(self.config_loader)
                stats = orchestrator.apply_fixes(current_path, fixable, output)

                result.iterations.append({
                    'iteration': iteration + 1,
                    'issues_found': len(issues),
                    'issues_fixed': stats['applied'],
                    'issues_failed': stats['failed'],
                    'duration': __import__('time').perf_counter() - t_iter,
                })

                current_path = output
                if stats['failed'] == 0:
                    break

            final_audit = audit_engine.scan_document(output)
            result.final_stats = {'remaining_issues': len(final_audit)}
            result.success = True

        except Exception as e:
            logger.error(f"Ошибка: {e}", exc_info=True)
            result.error = str(e)

        result.duration = __import__('time').perf_counter() - t_start
        return result

    def batch_format(
        self,
        input_dir: str,
        output_dir: str,
        options: Optional[dict] = None,
    ) -> BatchResult:
        """Пакетная обработка."""
        options = options or {}
        return self._batch.batch_format(
            input_dir,
            output_dir,
            clear_formatting=options.get('clear_direct_formatting', False),
            max_workers=options.get('max_workers', 1),
        )

    def diff(self, input_path: str, template_path: Optional[str] = None) -> DiffResult:
        """Сравнение с шаблоном."""
        template = template_path or self.template_path
        if not template:
            from src.core.formatter_errors import FormatterError, ErrorKind
            raise FormatterError(
                ErrorKind.TEMPLATE_NOT_FOUND,
                detail="для сравнения необходим template_path"
            )
        return self._diff.diff(input_path, template)

    def extract_config(self, template_path: str, output_path: str) -> dict:
        """Извлечение конфига из шаблона."""
        return self._diff.extract_config(template_path, output_path)

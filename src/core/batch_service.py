#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BatchService — пакетная обработка документов.
Поддерживает последовательную и параллельную обработку.
"""

import os
import glob
import logging
from typing import Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

from src.core.formatter_service import FormatterService
from src.core.models import FormatResult, BatchResult

logger = logging.getLogger(__name__)


class BatchService:
    """
    Сервис пакетной обработки документов.
    """

    def __init__(
        self,
        config_path: Optional[str] = None,
        template_path: Optional[str] = None,
    ):
        self.config_path = config_path
        self.template_path = template_path

    def batch_format(
        self,
        input_dir: str,
        output_dir: str,
        pattern: str = '*.docx',
        clear_formatting: bool = False,
        max_workers: int = 1,
    ) -> BatchResult:
        """
        Пакетная обработка документов.

        :param input_dir: входная директория.
        :param output_dir: выходная директория.
        :param pattern: паттерн файлов.
        :param clear_formatting: очистить прямое форматирование.
        :param max_workers: количество потоков (1=последовательно, >1=параллельно).
        :return: результат пакетной обработки.
        """
        os.makedirs(output_dir, exist_ok=True)

        search = os.path.join(input_dir, pattern)
        files = glob.glob(search)

        result = BatchResult()

        if max_workers <= 1:
            result = self._process_sequential(files, output_dir, clear_formatting)
        else:
            result = self._process_parallel(
                files, output_dir, clear_formatting, max_workers
            )

        logger.info(
            f"Пакетная обработка завершена: "
            f"{result.success_count}/{result.total} успешно"
        )

        return result

    def _process_sequential(
        self,
        files: list,
        output_dir: str,
        clear_formatting: bool,
    ) -> BatchResult:
        """Последовательная обработка."""
        result = BatchResult()

        for file_path in files:
            file_name = os.path.basename(file_path)
            output_path = os.path.join(output_dir, file_name)

            logger.info(f"Обработка: {file_name}")
            format_result = self._process_file(
                file_path, output_path, clear_formatting
            )
            result.results.append(format_result)

        return result

    def _process_parallel(
        self,
        files: list,
        output_dir: str,
        clear_formatting: bool,
        max_workers: int,
    ) -> BatchResult:
        """Параллельная обработка."""
        result = BatchResult()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for file_path in files:
                file_name = os.path.basename(file_path)
                output_path = os.path.join(output_dir, file_name)
                future = executor.submit(
                    self._process_file, file_path, output_path, clear_formatting
                )
                futures[future] = file_name

            for future in as_completed(futures):
                file_name = futures[future]
                try:
                    format_result = future.result()
                    result.results.append(format_result)
                    logger.info(f"Завершено: {file_name}")
                except Exception as e:
                    logger.error(f"Ошибка {file_name}: {e}")
                    result.results.append(FormatResult(
                        input=file_name,
                        output='',
                        mode='unknown',
                        success=False,
                        error=str(e),
                    ))

        return result

    def _process_file(
        self,
        file_path: str,
        output_path: str,
        clear_formatting: bool,
    ) -> FormatResult:
        """Обрабатывает один файл."""
        try:
            formatter = FormatterService(
                config_path=self.config_path,
                template_path=self.template_path,
            )
            return formatter.format_document(
                file_path,
                output_path,
                clear_formatting=clear_formatting,
            )
        except Exception as e:
            logger.error(f"Ошибка обработки {file_path}: {e}")
            return FormatResult(
                input=file_path,
                output=output_path,
                mode='unknown',
                success=False,
                error=str(e),
            )

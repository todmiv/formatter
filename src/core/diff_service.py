#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DiffService — сравнение документов и извлечение конфигов.
"""

import os
import logging
from typing import Dict, Any

from src.core.models import DiffResult

logger = logging.getLogger(__name__)


class DiffService:
    """
    Сервис сравнения документов и извлечения конфигов.
    """

    def diff(self, input_path: str, template_path: str) -> DiffResult:
        """
        Сравнивает документ с шаблоном.

        :param input_path: путь к документу.
        :param template_path: путь к шаблону.
        :return: результат сравнения.
        """
        from src.apply.style_applier import StyleApplier

        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Файл не найден: {input_path}")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Шаблон не найден: {template_path}")

        applier = StyleApplier(template_path)
        diff_data = applier.diff(input_path)

        return DiffResult(
            missing_styles=diff_data.get('missing_styles', []),
            different_styles=diff_data.get('different_styles', []),
            page_setup_diff=diff_data.get('page_setup_diff', False),
        )

    def extract_config(self, template_path: str, output_path: str) -> Dict[str, Any]:
        """
        Извлекает YAML-конфиг из DOCX-шаблона.

        :param template_path: путь к DOCX-шаблону.
        :param output_path: путь для сохранения YAML.
        :return: извлечённая конфигурация.
        """
        from src.core.template_extractor import TemplateExtractor

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Шаблон не найден: {template_path}")

        extractor = TemplateExtractor(template_path)
        config = extractor.extract()
        extractor.save_yaml(output_path, config)

        logger.info(f"Конфиг извлечён: {output_path}")
        return config

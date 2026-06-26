#!/usr/bin/env python3
"""
Тестирование функции _get_paragraph из apply_engine.py.
Проверяет обработку всех типов location, edge cases и корректность работы с колонтитулами.
"""

import sys
import os
import docx
import pytest
from src.apply.apply_orchestrator import ApplyOrchestrator
from src.core.config_loader import ConfigLoader


def find_config_file():
    """Находит конфигурационный файл."""
    candidates = ["configs/active/config_v4.2.yaml", "configs/active/config.yaml"]
    for c in candidates:
        if os.path.exists(c):
            return c
    raise FileNotFoundError("Конфиг не найден")


def test_get_paragraph():
    """Основная функция тестирования."""
    from src.apply.base_applier import BaseApplier
    from src.core.formatter_errors import FormatterError

    config_path = find_config_file()
    config = ConfigLoader(config_path)
    config.load()
    applier = BaseApplier(config)

    test_doc_path = "tests/test.docx"
    if not os.path.exists(test_doc_path):
        test_doc_path = "test_header_container.docx"
    if not os.path.exists(test_doc_path):
        test_doc_path = "Том II_Волжск.docx"

    if not os.path.exists(test_doc_path):
        pytest.skip("Тестовый документ не найден (test_header_container.docx или Том II_Волжск.docx)")

    doc = docx.Document(test_doc_path)

    test_cases = [
        {"name": "index valid", "location": {"index": 0}, "should_pass": True},
        {"name": "index out of range", "location": {"index": 9999}, "should_pass": False},
        {"name": "index None", "location": {"index": None}, "should_pass": False},
        {"name": "paragraph_index valid", "location": {"paragraph_index": 0}, "should_pass": True},
        {"name": "paragraph_index out of range", "location": {"paragraph_index": 9999}, "should_pass": False},
        {"name": "paragraph_index None", "location": {"paragraph_index": None}, "should_pass": False},
        {"name": "header with container", "location": {"section": 0, "paragraph": 0, "container": "header"}, "should_pass": True},
        {"name": "footer with container", "location": {"section": 0, "paragraph": 0, "container": "footer"}, "should_pass": True},
        {"name": "header without container", "location": {"section": 0, "paragraph": 0}, "should_pass": True},
        {"name": "empty location", "location": {}, "should_pass": False},
        {"name": "unknown keys", "location": {"unknown": 1}, "should_pass": False},
        {"name": "both index and paragraph_index", "location": {"index": 0, "paragraph_index": 1}, "should_pass": True},
        {"name": "table_index", "location": {"table_index": 0}, "should_pass": False},
        {"name": "figure_index", "location": {"figure_index": 0}, "should_pass": False},
    ]

    for test in test_cases:
        name = test["name"]
        location = test["location"]
        should_pass = test["should_pass"]

        try:
            paragraph, container = applier._get_paragraph(doc, location)
            if not should_pass:
                pytest.fail(f"{name}: НЕОЖИДАННЫЙ УСПЕХ (должен был провалиться)")
        except FormatterError:
            if should_pass:
                pytest.fail(f"{name}: НЕОЖИДАННАЯ ОШИБКА")
        except Exception as e:
            pytest.fail(f"{name}: ИСКЛЮЧЕНИЕ: {e}")

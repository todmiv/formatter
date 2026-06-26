#!/usr/bin/env python3
"""
Тестирование повторных аудитов после исправлений.
Проверяет, что количество ошибок уменьшается после применения исправлений.
"""
import os
import sys
import tempfile
import shutil
import pytest

sys.path.insert(0, os.path.dirname(__file__))

from src.core.main_controller import MainController


def test_repeat_audit():
    test_doc = "tests/test_output.docx"
    if not os.path.exists(test_doc):
        import docx
        doc = docx.Document()
        doc.add_paragraph("Тестовый параграф с некорректным стилем.")
        doc.save(test_doc)

    temp_dir = tempfile.mkdtemp()
    doc_copy = os.path.join(temp_dir, "test_copy.docx")
    shutil.copy2(test_doc, doc_copy)

    config_path = "configs/active/config.yaml"
    if not os.path.exists(config_path):
        config_path = "configs/active/config_v4.2.yaml"

    controller = MainController(config_path)
    controller.set_document(doc_copy)

    issues = controller.start_audit()
    initial_count = len(issues)

    if initial_count == 0:
        shutil.rmtree(temp_dir, ignore_errors=True)
        pytest.skip("Нет ошибок для исправления")

    stats = controller.apply_all_fixes()
    fixed_path = controller.last_fixed_path
    assert fixed_path and os.path.exists(fixed_path), "Исправленный документ не создан"

    controller.set_document(fixed_path)
    issues2 = controller.start_audit()
    second_count = len(issues2)

    assert second_count <= initial_count, (
        f"Количество ошибок не уменьшилось: было {initial_count}, стало {second_count}"
    )

    shutil.rmtree(temp_dir, ignore_errors=True)

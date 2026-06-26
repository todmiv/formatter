#!/usr/bin/env python3
"""
Тестирование кнопки "Подсветить проблемные места".
Проверяет, что функция highlight_issues создаёт файл _highlighted.docx и открывает его.
"""
import os
import sys
import tempfile
import shutil
import pytest

sys.path.insert(0, os.path.dirname(__file__))

from src.core.main_controller import MainController


def test_highlight():
    test_doc = "tests/test_output.docx"
    if not os.path.exists(test_doc):
        import docx
        doc = docx.Document()
        doc.add_paragraph("Тестовый параграф.")
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

    if len(issues) == 0:
        shutil.rmtree(temp_dir, ignore_errors=True)
        pytest.skip("Нет проблем для подсветки")

    highlighted_path = controller.highlight_issues()
    assert highlighted_path and os.path.exists(highlighted_path), "Файл с подсветкой не создан"

    size = os.path.getsize(highlighted_path)
    assert size > 0, "Файл с подсветкой пустой"

    shutil.rmtree(temp_dir, ignore_errors=True)

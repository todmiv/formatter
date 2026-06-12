#!/usr/bin/env python3
"""
Тестирование функции _get_paragraph из apply_engine.py.
Проверяет обработку всех типов location, edge cases и корректность работы с колонтитулами.
"""

import sys
import os
import docx
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
    config_path = find_config_file()
    config = ConfigLoader(config_path)
    config.load()
    engine = ApplyOrchestrator(config)
    
    # Используем тестовый документ с колонтитулами
    test_doc_path = "tests/test.docx"
    if not os.path.exists(test_doc_path):
        test_doc_path = "test_header_container.docx"
    if not os.path.exists(test_doc_path):
        test_doc_path = "Том II_Волжск.docx"
    
    if not os.path.exists(test_doc_path):
        import pytest
        pytest.skip("Тестовый документ не найден (test_header_container.docx или Том II_Волжск.docx)")
    
    doc = docx.Document(test_doc_path)
    print(f"Тестируем документ: {test_doc_path}")
    print(f"Параграфов в документе: {len(doc.paragraphs)}")
    print(f"Секций: {len(doc.sections)}")
    
    # Подготовим тестовые location
    test_cases = [
        # Основное тело документа
        {"name": "index valid", "location": {"index": 0}, "should_pass": True},
        {"name": "index out of range", "location": {"index": 9999}, "should_pass": False},
        {"name": "index None", "location": {"index": None}, "should_pass": False},
        {"name": "paragraph_index valid", "location": {"paragraph_index": 0}, "should_pass": True},
        {"name": "paragraph_index out of range", "location": {"paragraph_index": 9999}, "should_pass": False},
        {"name": "paragraph_index None", "location": {"paragraph_index": None}, "should_pass": False},
        # Колонтитулы с явным container
        {"name": "header with container", "location": {"section": 0, "paragraph": 0, "container": "header"}, "should_pass": True},
        {"name": "footer with container", "location": {"section": 0, "paragraph": 0, "container": "footer"}, "should_pass": True},
        # Колонтитулы без container (должен предположить header)
        {"name": "header without container", "location": {"section": 0, "paragraph": 0}, "should_pass": True},
        # Неподдерживаемые форматы
        {"name": "empty location", "location": {}, "should_pass": False},
        {"name": "unknown keys", "location": {"unknown": 1}, "should_pass": False},
        # Смешанные ключи
        {"name": "both index and paragraph_index", "location": {"index": 0, "paragraph_index": 1}, "should_pass": True},  # должен использовать index
    ]
    
    # Дополнительные тесты для таблиц и рисунков (не должны поддерживаться)
    test_cases.append({"name": "table_index", "location": {"table_index": 0}, "should_pass": False})
    test_cases.append({"name": "figure_index", "location": {"figure_index": 0}, "should_pass": False})
    
    passed = 0
    failed = 0
    
    for test in test_cases:
        name = test["name"]
        location = test["location"]
        should_pass = test["should_pass"]
        
        try:
            paragraph, container = engine._get_paragraph(doc, location)
            if should_pass:
                print(f"[PASS] {name}: УСПЕХ, найден параграф, контейнер: {container}")
                passed += 1
            else:
                print(f"[FAIL] {name}: НЕОЖИДАННЫЙ УСПЕХ (должен был провалиться)")
                failed += 1
        except ValueError as e:
            if not should_pass:
                print(f"[PASS] {name}: ОЖИДАЕМАЯ ОШИБКА: {e}")
                passed += 1
            else:
                print(f"[FAIL] {name}: НЕОЖИДАННАЯ ОШИБКА: {e}")
                failed += 1
        except Exception as e:
            print(f"[FAIL] {name}: ИСКЛЮЧЕНИЕ: {e}")
            failed += 1
    
    print(f"\nИтог: {passed} пройдено, {failed} провалено")
    
    # Дополнительная проверка: тестирование location из реальных проблем
    print("\n--- Тестирование location из реальных проблем ---")
    real_locations = [
        {"index": 3, "page": 1},
        {"paragraph_index": 10, "page": 0},
        {"table_index": 0, "page": 0},
        {"section": 0, "paragraph": 0, "container": "header"},
        {"section": 0, "paragraph": 0, "container": "footer"},
    ]
    
    for loc in real_locations:
        try:
            paragraph, container = engine._get_paragraph(doc, loc)
            print(f"[OK] Реальный location {loc}: УСПЕХ, контейнер: {container}")
        except ValueError as e:
            print(f"[ERR] Реальный location {loc}: ОШИБКА: {e}")
        except Exception as e:
            print(f"[ERR] Реальный location {loc}: ИСКЛЮЧЕНИЕ: {e}")
    
    # Проверка edge cases с None значениями
    print("\n--- Проверка edge cases ---")
    edge_cases = [
        {"location": {"index": -1}, "desc": "Отрицательный индекс"},
        {"location": {"index": len(doc.paragraphs) - 1}, "desc": "Последний параграф"},
        {"location": {"section": 999, "paragraph": 0, "container": "header"}, "desc": "Несуществующая секция"},
        {"location": {"section": 0, "paragraph": 999, "container": "header"}, "desc": "Несуществующий параграф в колонтитуле"},
        {"location": {"section": 0, "paragraph": 0, "container": "unknown"}, "desc": "Неизвестный контейнер"},
    ]
    
    for case in edge_cases:
        try:
            paragraph, container = engine._get_paragraph(doc, case["location"])
            print(f"[OK] {case['desc']}: УСПЕХ")
        except ValueError as e:
            print(f"[OK] {case['desc']}: ОЖИДАЕМАЯ ОШИБКА: {e}")
        except Exception as e:
            print(f"[ERR] {case['desc']}: ИСКЛЮЧЕНИЕ: {e}")
    
    return failed == 0

if __name__ == "__main__":
    success = test_get_paragraph()
    sys.exit(0 if success else 1)
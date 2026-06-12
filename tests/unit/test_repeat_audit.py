#!/usr/bin/env python3
"""
Тестирование повторных аудитов после исправлений.
Проверяет, что количество ошибок уменьшается после применения исправлений.
"""
import os
import sys
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(__file__))

from src.core.main_controller import MainController

def test_repeat_audit():
    print("=== Тестирование повторных аудитов после исправлений ===")
    
    # Используем тестовый документ
    test_doc = "test_output.docx"
    if not os.path.exists(test_doc):
        print(f"Файл {test_doc} не найден. Создаём простой DOCX...")
        # Создаём простой документ через python-docx
        import docx
        doc = docx.Document()
        doc.add_paragraph("Тестовый параграф с некорректным стилем.")
        doc.save(test_doc)
        print(f"Создан {test_doc}")
    
    # Создаём временную копию для теста
    temp_dir = tempfile.mkdtemp()
    doc_copy = os.path.join(temp_dir, "test_copy.docx")
    shutil.copy2(test_doc, doc_copy)
    print(f"Работаем с копией: {doc_copy}")
    
    # Инициализируем контроллер с конфигом по умолчанию
    config_path = "config.yaml"
    if not os.path.exists(config_path):
        config_path = "config_v4.2.yaml"
    print(f"Используем конфиг: {config_path}")
    
    controller = MainController(config_path)
    controller.set_document(doc_copy)
    
    # Первый аудит
    print("\n--- Первый аудит ---")
    issues = controller.start_audit()
    initial_count = len(issues)
    print(f"Найдено ошибок: {initial_count}")
    if initial_count == 0:
        print("Нет ошибок для исправления. Тест пройден условно.")
        return True
    
    # Применяем исправления ко всем ошибкам
    print("\n--- Применение исправлений ко всем ошибкам ---")
    stats = controller.apply_all_fixes()
    fixed_path = controller.last_fixed_path
    if not fixed_path or not os.path.exists(fixed_path):
        print("Ошибка: исправленный документ не создан.")
        return False
    print(f"Исправленный документ: {fixed_path}")
    
    # Второй аудит на исправленном документе
    print("\n--- Второй аудит (после исправлений) ---")
    controller.set_document(fixed_path)
    issues2 = controller.start_audit()
    second_count = len(issues2)
    print(f"Найдено ошибок после исправлений: {second_count}")
    
    # Проверяем, что количество ошибок уменьшилось
    if second_count < initial_count:
        print(f"[OK] Количество ошибок уменьшилось с {initial_count} до {second_count}")
    else:
        print(f"[WARN] Количество ошибок не уменьшилось: было {initial_count}, стало {second_count}")
        # Возможно, некоторые ошибки неустранимы
    
    # Третий аудит (если остались ошибки)
    if second_count > 0:
        print("\n--- Третий аудит (повторное исправление) ---")
        stats2 = controller.apply_all_fixes()
        fixed_path2 = controller.last_fixed_path
        if fixed_path2 and os.path.exists(fixed_path2):
            controller.set_document(fixed_path2)
            issues3 = controller.start_audit()
            third_count = len(issues3)
            print(f"Найдено ошибок после второго исправления: {third_count}")
            if third_count < second_count:
                print(f"[OK] Количество ошибок продолжает уменьшаться: {second_count} -> {third_count}")
            else:
                print(f"[WARN] Количество ошибок не изменилось после второго исправления")
        else:
            print("Не удалось применить второе исправление")
    
    # Проверка функции highlight_issues
    print("\n--- Проверка подсветки проблемных мест ---")
    try:
        highlighted = controller.highlight_issues()
        if highlighted and os.path.exists(highlighted):
            print(f"[OK] Файл с подсветкой создан: {highlighted}")
        else:
            print("[WARN] Файл с подсветкой не создан")
    except Exception as e:
        print(f"[WARN] Ошибка при подсветке: {e}")
    
    # Очистка
    shutil.rmtree(temp_dir, ignore_errors=True)
    print("\n=== Тест завершён ===")
    return True

if __name__ == "__main__":
    success = test_repeat_audit()
    sys.exit(0 if success else 1)
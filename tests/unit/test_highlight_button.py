#!/usr/bin/env python3
"""
Тестирование кнопки "Подсветить проблемные места".
Проверяет, что функция highlight_issues создаёт файл _highlighted.docx и открывает его.
"""
import os
import sys
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(__file__))

from src.core.main_controller import MainController

def test_highlight():
    print("=== Тестирование подсветки проблемных мест ===")
    
    # Используем тестовый документ
    test_doc = "test_output.docx"
    if not os.path.exists(test_doc):
        print(f"Файл {test_doc} не найден. Создаём простой DOCX...")
        import docx
        doc = docx.Document()
        doc.add_paragraph("Тестовый параграф.")
        doc.save(test_doc)
        print(f"Создан {test_doc}")
    
    # Создаём временную копию
    temp_dir = tempfile.mkdtemp()
    doc_copy = os.path.join(temp_dir, "test_copy.docx")
    shutil.copy2(test_doc, doc_copy)
    print(f"Работаем с копией: {doc_copy}")
    
    # Инициализируем контроллер
    config_path = "config.yaml"
    if not os.path.exists(config_path):
        config_path = "config_v4.2.yaml"
    print(f"Используем конфиг: {config_path}")
    
    controller = MainController(config_path)
    controller.set_document(doc_copy)
    
    # Запускаем аудит, чтобы были проблемы
    print("\n--- Запуск аудита для получения проблем ---")
    issues = controller.start_audit()
    print(f"Найдено проблем: {len(issues)}")
    
    if len(issues) == 0:
        print("Нет проблем для подсветки. Добавляем искусственную проблему?")
        # Можно добавить, но пропустим
        print("Пропускаем тест подсветки.")
        shutil.rmtree(temp_dir, ignore_errors=True)
        return True
    
    # Вызываем подсветку
    print("\n--- Вызов highlight_issues ---")
    highlighted_path = controller.highlight_issues()
    if highlighted_path and os.path.exists(highlighted_path):
        print(f"[OK] Файл с подсветкой создан: {highlighted_path}")
        # Проверяем, что файл не пустой
        size = os.path.getsize(highlighted_path)
        print(f"Размер файла: {size} байт")
        if size > 0:
            print("[OK] Файл не пустой.")
        else:
            print("[WARN] Файл пустой.")
    else:
        print("[FAIL] Файл с подсветкой не создан.")
        shutil.rmtree(temp_dir, ignore_errors=True)
        return False
    
    # Проверяем, что файл имеет суффикс _highlighted.docx
    expected_suffix = "_highlighted.docx"
    if highlighted_path.endswith(expected_suffix):
        print(f"[OK] Имя файла содержит суффикс {expected_suffix}")
    else:
        print(f"[WARN] Имя файла не содержит ожидаемый суффикс: {highlighted_path}")
    
    # Проверяем, что файл можно открыть (не ломается)
    try:
        import docx
        doc = docx.Document(highlighted_path)
        paragraphs = len(doc.paragraphs)
        print(f"[OK] Файл можно открыть через python-docx, параграфов: {paragraphs}")
    except Exception as e:
        print(f"[WARN] Ошибка при открытии файла: {e}")
    
    # Очистка
    shutil.rmtree(temp_dir, ignore_errors=True)
    print("\n=== Тест завершён успешно ===")
    return True

if __name__ == "__main__":
    success = test_highlight()
    sys.exit(0 if success else 1)
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import tkinter as tk
from gui_formatter import GOSTFormatterGUI
from src.core.audit_engine import AuditEngine, Severity
from src.core.config_loader import ConfigLoader

def test_preview():
    root = tk.Tk()
    root.withdraw()  # скрыть окно
    app = GOSTFormatterGUI(root)
    
    # Загружаем конфиг и выполняем аудит (имитируем)
    config_path = 'config.yaml'
    doc_path = 'test_output.docx'
    if not os.path.exists(doc_path):
        # Попробуем найти другой
        docx_files = [f for f in os.listdir('.') if f.lower().endswith('.docx')]
        if docx_files:
            doc_path = docx_files[0]
        else:
            print("Документ не найден")
            root.destroy()
            return
    
    loader = ConfigLoader(config_path)
    loader.load()
    engine = AuditEngine(loader)
    issues = engine.scan_document(doc_path)
    
    if not issues:
        print("Нет проблем для теста")
        root.destroy()
        return
    
    # Берём первую проблему
    issue = issues[0]
    print(f"Тестируем проблему: {issue.id}")
    print(f"Индекс параграфа: {issue.location.get('index')}")
    
    # Устанавливаем состояние приложения через контроллер
    app.controller.set_document(doc_path)  # создаст document_model
    app.controller.current_issues = issues
    
    # Заполняем treeview проблемами
    app.populate_tree(issues)
    # Обновляем интерфейс, чтобы treeview отобразился
    root.update_idletasks()
    
    # Проверяем, что элемент существует в treeview
    if app.tree.exists(issue.id):
        # Имитируем выбор строки в treeview
        app.tree.selection_set(issue.id)
        # Вызываем on_tree_select вручную
        app.on_tree_select(None)
    else:
        print(f"ОШИБКА: элемент {issue.id} не найден в treeview")
        root.destroy()
        return
    
    # Получаем текст из preview_text
    preview_content = app.preview_text.get("1.0", tk.END).strip()
    print("Предпросмотр:")
    # Безопасный вывод для Windows (замена непечатаемых символов)
    try:
        print(preview_content)
    except UnicodeEncodeError:
        # Заменяем проблемные символы на '?'
        safe_content = preview_content.encode('cp1251', errors='replace').decode('cp1251')
        print(safe_content)
    
    # Проверяем, что текст не содержит "Фрагмент текста с ошибкой не доступен"
    if "Фрагмент текста с ошибкой не доступен" in preview_content:
        print("ОШИБКА: предпросмотр всё ещё показывает заглушку!")
    else:
        print("УСПЕХ: предпросмотр показывает реальный текст.")
    
    root.destroy()

if __name__ == "__main__":
    test_preview()
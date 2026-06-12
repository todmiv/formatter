# -*- coding: utf-8 -*-
# test_audit.py

import sys
import io
import os

# Принудительно устанавливаем UTF-8 для стандартного вывода, если возможно
# Это помогает в некоторых терминалах, но замена эмодзи - более надежное решение для Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        # Для старых версий Python (< 3.7) используем обертку
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine

def main():
    # Путь к конфигурации и файлу
    CONFIG_PATH = 'config.yaml'
    # Используем существующий тестовый файл
    DOC_PATH = 'test_output.docx'
    # Если файла нет, попробуем другой
    if not os.path.exists(DOC_PATH):
        # Попробуем найти любой .docx файл
        docx_files = [f for f in os.listdir('.') if f.lower().endswith('.docx')]
        if docx_files:
            DOC_PATH = docx_files[0]
        else:
            print("[ERROR] В папке нет DOCX файлов для теста.")
            return

    print("=" * 60)
    print("ЗАПУСК АУДИТА ДОКУМЕНТА")
    print("=" * 60)

    try:
        # 1. Загрузка конфига
        print(f"[INFO] Загрузка конфигурации из {CONFIG_PATH}...")
        loader = ConfigLoader(CONFIG_PATH)
        loader.load()
        print("[OK] Конфигурация загружена успешно")
        print(f"     Стандарт: {loader.config.get('name', 'Не указан')}")

        # 2. Запуск аудита
        print(f"\n[INFO] Сканирование документа: {DOC_PATH}...")
        engine = AuditEngine(loader)
        issues = engine.scan_document(DOC_PATH)
        
        print("[OK] Аудит завершен")
        
        # 3. Вывод сводки
        summary = engine.get_summary()
        total = sum(summary.values())
        
        print("\n" + "-" * 40)
        print(f"ВСЕГО НАЙДЕНО ПРОБЛЕМ: {total}")
        print("-" * 40)
        print(f"  КРИТИЧЕСКИЕ (CRITICAL): {summary['CRITICAL']}")
        print(f"  ПРЕДУПРЕЖДЕНИЯ (WARNING): {summary['WARNING']}")
        print(f"  ИНФОРМАЦИЯ (INFO):      {summary['INFO']}")
        print("-" * 40)

        if not issues:
            print("\n[RESULT] Документ полностью соответствует конфигурации!")
            return

        # 4. Детальный вывод первых 10 проблем
        print("\nСПИСОК ПЕРВЫХ 10 НАРУШЕНИЙ:")
        print("{:<5} | {:<10} | {:<20} | {:<30}".format("№", "Тип", "Стиль", "Проблема"))
        print("-" * 80)

        for i, issue in enumerate(issues[:10]):
            # Форматированный вывод без эмодзи
            status = "CRIT" if issue.severity.value == "CRITICAL" else "WARN"
            print("{:<5} | {:<10} | {:<20} | {:<30}".format(
                i + 1, 
                status, 
                issue.element_type, 
                issue.description
            ))
            print("      | Текущее: {}".format(issue.current_value))
            print("      | Ожидается: {}".format(issue.expected_value))
            print("-" * 80)

        if len(issues) > 10:
            print(f"... и еще {len(issues) - 10} проблем (см. полный лог или GUI)")

        print("\n[INFO] Для применения исправлений запустите gui_formatter.py")

    except FileNotFoundError as e:
        print(f"\n[ERROR] Файл не найден: {e}")
        print("Проверьте, что config.yaml и документ лежат в папке со скриптом.")
    except Exception as e:
        print(f"\n[ERROR] Произошла критическая ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

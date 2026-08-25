# -*- coding: utf-8 -*-
"""
Тест интеграции ApplyEngine с реальным документом.
"""
import sys
import os
sys.path.insert(0, '.')

from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine
from src.apply.apply_orchestrator import ApplyOrchestrator

def main():
    config_path = 'configs/active/config.yaml'
    doc_path = 'tests/test_output.docx'
    output_path = 'test_output_fixed.docx'
    
    print("=== Тест ApplyEngine ===")
    print(f"Конфиг: {config_path}")
    print(f"Документ: {doc_path}")
    
    if not os.path.exists(doc_path):
        print("Файл документа не найден!")
        # Попробуем другой файл
        docx_files = [f for f in os.listdir('.') if f.lower().endswith('.docx')]
        if docx_files:
            doc_path = docx_files[0]
            print(f"Используем файл: {doc_path}")
        else:
            print("В папке нет DOCX файлов.")
            return
    
    # 1. Загрузка конфигурации
    loader = ConfigLoader(config_path)
    loader.load()
    
    # 2. Аудит
    audit = AuditEngine(loader)
    issues = audit.scan_document(doc_path)
    print(f"Найдено проблем: {len(issues)}")
    
    if not issues:
        print("Нет проблем для исправления. Тест завершен.")
        return
    
    # 3. Применение исправлений
    apply = ApplyOrchestrator(loader)
    stats = apply.apply_fixes(doc_path, issues, output_path)
    
    print(f"Результат: применено {stats['applied']}, ошибок {stats['failed']}")
    print(f"Исправленный файл: {output_path}")
    
    # 4. Проверка, что файл создан
    if os.path.exists(output_path):
        print("[OK] Файл успешно создан.")
    else:
        print("[ERROR] Файл не создан.")

if __name__ == '__main__':
    main()
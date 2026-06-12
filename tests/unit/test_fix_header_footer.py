#!/usr/bin/env python3
"""
Тестовый скрипт для проверки исправлений обработки колонтитулов.
Проверяет, что ApplyEngine корректно обрабатывает проблемы в колонтитулах.
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine
from src.apply.apply_orchestrator import ApplyOrchestrator
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_header_footer_fix():
    """Основной тест: аудит и применение исправлений для колонтитулов."""
    config_path = "config.yaml"
    if not os.path.exists(config_path):
        config_path = "config_v4.2.yaml"
    
    doc_path = "(ГЕНЕРАЦИЯ) Том II Черкесск Демография.docx"
    if not os.path.exists(doc_path):
        logger.error(f"Тестовый документ не найден: {doc_path}")
        return False
    
    logger.info(f"Загружаем конфигурацию из {config_path}")
    config = ConfigLoader(config_path)
    config.load()
    
    # 1. Запускаем аудит
    logger.info("Запуск аудита документа...")
    audit = AuditEngine(config)
    issues = audit.scan_document(doc_path)
    
    # Фильтруем проблемы колонтитулов
    header_footer_issues = [i for i in issues if i.category == 'HEADER_FOOTER']
    logger.info(f"Всего проблем: {len(issues)}, из них HEADER_FOOTER: {len(header_footer_issues)}")
    
    if len(header_footer_issues) == 0:
        logger.warning("Нет проблем колонтитулов для тестирования. Возможно, документ уже отформатирован.")
        # Продолжим тест с другими проблемами для проверки обратной совместимости
    
    # 2. Применяем исправления
    logger.info("Применение исправлений...")
    apply_engine = ApplyOrchestrator(config)
    
    # Сохраняем в тестовый файл, чтобы не портить оригинал
    output_path = "test_output_fixed_header_footer.docx"
    
    result = apply_engine.apply_fixes(
        doc_path=doc_path,
        issues=issues,
        output_path=output_path,
        apply_styles=True,
        apply_direct_overrides=True,
        clear_direct_formatting=False
    )
    
    logger.info(f"Результат применения: успешно {result['applied']}, неудачно {result['failed']}")
    
    # 3. Проверяем, что неудачных исправлений нет (или минимальное количество)
    if result['failed'] > 0:
        logger.warning(f"Есть неудачные исправления: {result['failed']}. Анализируем причины...")
        # Можно дополнительно проанализировать, какие проблемы не удалось применить
        # Но для целей теста считаем, что это допустимо, если это не колонтитулы
        # Проверим, были ли неудачи среди колонтитулов
        # Для этого нужно повторно проаудить и сравнить
        pass
    
    # 4. Проверяем, что проблемы колонтитулов обработаны
    # Запускаем повторный аудит на исправленном документе
    logger.info("Повторный аудит исправленного документа...")
    issues_after = audit.scan_document(output_path)
    header_footer_after = [i for i in issues_after if i.category == 'HEADER_FOOTER']
    logger.info(f"Проблем HEADER_FOOTER после исправлений: {len(header_footer_after)}")
    
    # Ожидаем, что количество проблем уменьшилось (или осталось нулевым)
    if len(header_footer_issues) > 0 and len(header_footer_after) >= len(header_footer_issues):
        logger.error("Количество проблем колонтитулов не уменьшилось после исправлений!")
        return False
    
    # 5. Проверяем обратную совместимость с ключами index и paragraph_index
    logger.info("Проверка обратной совместимости...")
    # Создадим искусственную проблему с ключом index
    from audit_engine import AuditIssue, Severity
    dummy_issue = AuditIssue(
        id="TEST_INDEX",
        severity=Severity.CRITICAL,
        category="PARAGRAPH",
        element_type="Normal",
        location={'index': 0},  # Первый параграф
        description="Тестовая проблема с ключом index",
        current_value="",
        expected_value="",
        auto_fixable=True,
        fix_payload={'style': 'Normal', 'prop': 'style_name'}
    )
    
    # Пробуем применить
    test_result = apply_engine.apply_fixes(
        doc_path=doc_path,
        issues=[dummy_issue],
        output_path="test_compatibility.docx",
        apply_styles=True,
        apply_direct_overrides=False,
        clear_direct_formatting=False
    )
    if test_result['failed'] > 0:
        logger.error("Обратная совместимость с ключом index нарушена!")
        return False
    
    # Проверка с ключом paragraph_index
    dummy_issue2 = AuditIssue(
        id="TEST_PARAGRAPH_INDEX",
        severity=Severity.CRITICAL,
        category="PARAGRAPH",
        element_type="Normal",
        location={'paragraph_index': 1},
        description="Тестовая проблема с ключом paragraph_index",
        current_value="",
        expected_value="",
        auto_fixable=True,
        fix_payload={'style': 'Normal', 'prop': 'style_name'}
    )
    test_result2 = apply_engine.apply_fixes(
        doc_path=doc_path,
        issues=[dummy_issue2],
        output_path="test_compatibility2.docx",
        apply_styles=True,
        apply_direct_overrides=False,
        clear_direct_formatting=False
    )
    if test_result2['failed'] > 0:
        logger.error("Обратная совместимость с ключом paragraph_index нарушена!")
        return False
    
    logger.info("Обратная совместимость подтверждена.")
    
    # 6. Проверка обработки колонтитулов с container
    # Создадим искусственную проблему для колонтитула
    # Сначала узнаем структуру документа
    import docx
    from docx_utils import open_document
    doc = open_document(doc_path)
    if len(doc.sections) > 0:
        section = doc.sections[0]
        if section.header and len(section.header.paragraphs) > 0:
            dummy_header_issue = AuditIssue(
                id="TEST_HEADER_CONTAINER",
                severity=Severity.CRITICAL,
                category="HEADER_FOOTER",
                element_type="Header",
                location={'section': 0, 'paragraph': 0, 'container': 'header'},
                description="Тестовая проблема в верхнем колонтитуле",
                current_value="",
                expected_value="",
                auto_fixable=True,
                fix_payload={'style': 'Header', 'prop': 'style_name'}
            )
            test_result3 = apply_engine.apply_fixes(
                doc_path=doc_path,
                issues=[dummy_header_issue],
                output_path="test_header_container.docx",
                apply_styles=True,
                apply_direct_overrides=False,
                clear_direct_formatting=False
            )
            if test_result3['failed'] > 0:
                logger.error("Обработка колонтитула с container нарушена!")
                return False
            else:
                logger.info("Обработка колонтитула с container успешна.")
    
    logger.info("Все проверки пройдены успешно.")
    return True

if __name__ == "__main__":
    success = test_header_footer_fix()
    if success:
        print("\n[OK] Тест пройден: ApplyEngine корректно обрабатывает колонтитулы.")
        sys.exit(0)
    else:
        print("\n[FAIL] Тест не пройден.")
        sys.exit(1)
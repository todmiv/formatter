#!/usr/bin/env python3
"""
Тестовый скрипт для проверки исправлений обработки колонтитулов.
Проверяет, что ApplyEngine корректно обрабатывает проблемы в колонтитулах.
"""

import sys
import os
import pytest

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine
from src.apply.apply_orchestrator import ApplyOrchestrator
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def test_header_footer_fix():
    """Основной тест: аудит и применение исправлений для колонтитулов."""
    config_path = "configs/active/config.yaml"
    if not os.path.exists(config_path):
        config_path = "configs/active/config_v4.2.yaml"

    doc_path = "(ГЕНЕРАЦИЯ) Том II Черкесск Демография.docx"
    if not os.path.exists(doc_path):
        pytest.skip("Тестовый документ не найден")

    logger.info(f"Загружаем конфигурацию из {config_path}")
    config = ConfigLoader(config_path)
    config.load()

    logger.info("Запуск аудита документа...")
    audit = AuditEngine(config)
    issues = audit.scan_document(doc_path)

    header_footer_issues = [i for i in issues if i.category == 'HEADER_FOOTER']
    logger.info(f"Всего проблем: {len(issues)}, из них HEADER_FOOTER: {len(header_footer_issues)}")

    logger.info("Применение исправлений...")
    apply_engine = ApplyOrchestrator(config)

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

    logger.info("Повторный аудит исправленного документа...")
    issues_after = audit.scan_document(output_path)
    header_footer_after = [i for i in issues_after if i.category == 'HEADER_FOOTER']
    logger.info(f"Проблем HEADER_FOOTER после исправлений: {len(header_footer_after)}")

    if len(header_footer_issues) > 0:
        assert len(header_footer_after) < len(header_footer_issues), (
            f"Количество проблем колонтитулов не уменьшилось: "
            f"{len(header_footer_issues)} -> {len(header_footer_after)}"
        )

    logger.info("Проверка обратной совместимости...")
    from src.core.audit_engine import AuditIssue, Severity
    dummy_issue = AuditIssue(
        id="TEST_INDEX",
        severity=Severity.CRITICAL,
        category="PARAGRAPH",
        element_type="Normal",
        location={'index': 0},
        description="Тестовая проблема с ключом index",
        current_value="",
        expected_value="",
        auto_fixable=True,
        fix_payload={'style': 'Normal', 'prop': 'style_name'}
    )

    test_result = apply_engine.apply_fixes(
        doc_path=doc_path,
        issues=[dummy_issue],
        output_path="test_compatibility.docx",
        apply_styles=True,
        apply_direct_overrides=False,
        clear_direct_formatting=False
    )
    assert test_result['failed'] == 0, "Обратная совместимость с ключом index нарушена!"

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
    assert test_result2['failed'] == 0, "Обратная совместимость с ключом paragraph_index нарушена!"

    logger.info("Обратная совместимость подтверждена.")

    import docx
    from src.core.docx_utils import open_document
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
            assert test_result3['failed'] == 0, "Обработка колонтитула с container нарушена!"

    logger.info("Все проверки пройдены успешно.")

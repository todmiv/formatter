#!/usr/bin/env python3
"""Тесты для BaseAuditor и AuditIssue."""

import os
import tempfile
import shutil
import pytest
from unittest.mock import Mock, MagicMock
from docx import Document
from docx.shared import Pt, Cm

from src.audit.base_auditor import BaseAuditor, AuditIssue, Severity
from src.core.config_loader import ConfigLoader


def make_config():
    """Создаёт тестовую конфигурацию."""
    temp_dir = tempfile.mkdtemp()
    config_path = os.path.join(temp_dir, 'config.yaml')
    with open(config_path, 'w', encoding='utf-8') as f:
        f.write("""
styles:
  Normal:
    enabled: true
    font:
      name: "Times New Roman"
      size: 12
      bold: false
      italic: false
      color: "black"
    paragraph:
      first_line_indent_cm: 1.25
      line_spacing: 1.15
      alignment: "justify"
  Heading 1:
    enabled: true
    font:
      name: "Calibri Light"
      size: 16
      bold: true
detection_rules:
  Heading 1:
    - "^RAZDEL\\\\s+\\\\d+"
""")
    config = ConfigLoader(config_path)
    config.load()
    return config, temp_dir


class TestAuditIssue:
    """Тесты для AuditIssue."""

    def test_create_issue(self):
        issue = AuditIssue(
            id='TEST_001',
            severity=Severity.CRITICAL,
            category='FONT',
            element_type='Normal',
            location={'index': 0},
            description='Test issue',
            current_value='Arial',
            expected_value='Times New Roman',
        )
        assert issue.id == 'TEST_001'
        assert issue.severity == Severity.CRITICAL
        assert issue.auto_fixable is True

    def test_issue_with_recommendation(self):
        issue = AuditIssue(
            id='TEST_001',
            severity=Severity.WARNING,
            category='PARAGRAPH',
            element_type='Normal',
            location={'index': 0},
            description='Test',
            current_value='',
            expected_value='',
            recommendation='Исправьте отступ',
        )
        assert issue.recommendation == 'Исправьте отступ'


class TestBaseAuditor:
    """Тесты для BaseAuditor."""

    def setup_method(self):
        self.config, self.temp_dir = make_config()

    def teardown_method(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_init(self):
        auditor = BaseAuditor(self.config)
        assert auditor.config == self.config
        assert auditor.issues == []

    def test_init_with_custom_tolerances(self):
        auditor = BaseAuditor(
            self.config,
            font_size_tolerance_twips=5,
            indent_tolerance_twips=10,
            spacing_tolerance_twips=15,
        )
        assert auditor.FONT_SIZE_TOLERANCE_TWIPS == 5
        assert auditor.INDENT_TOLERANCE_TWIPS == 10
        assert auditor.SPACING_TOLERANCE_TWIPS == 15

    def test_create_issue(self):
        auditor = BaseAuditor(self.config)
        issue = auditor._create_issue(
            severity=Severity.WARNING,
            category='FONT',
            elem_type='Normal',
            loc={'index': 0},
            desc='Test',
            current='Arial',
            expected='Times New Roman',
        )
        assert issue.id == 'FONT_0001'
        assert issue.severity == Severity.WARNING
        assert issue.auto_fixable is True

    def test_create_issue_counter(self):
        auditor = BaseAuditor(self.config)
        issue1 = auditor._create_issue(
            severity=Severity.WARNING, category='FONT',
            elem_type='Normal', loc={'index': 0},
            desc='Test', current='', expected='',
        )
        issue2 = auditor._create_issue(
            severity=Severity.WARNING, category='FONT',
            elem_type='Normal', loc={'index': 0},
            desc='Test', current='', expected='',
        )
        assert issue1.id != issue2.id

    def test_detect_target_style_explicit(self):
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Test')
        para.style = doc.styles['Normal']

        target = auditor._detect_target_style(para)
        assert target == 'Normal'

    def test_detect_target_style_by_pattern(self):
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Test text')

        target = auditor._detect_target_style(para)
        assert target == 'Normal'

    def test_detect_target_style_fallback(self):
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Обычный текст')

        target = auditor._detect_target_style(para)
        assert target == 'Normal'

    def test_update_progress(self):
        callback = Mock()
        auditor = BaseAuditor(self.config, progress_callback=callback)
        auditor._update_progress('test', 50, 100, 'message')
        callback.assert_called_once_with('test', 50, 100, 'message')

    def test_update_progress_no_callback(self):
        auditor = BaseAuditor(self.config)
        auditor._update_progress('test', 50, 100, 'message')

    def test_check_font_correct(self):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Test')
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        issues = auditor._check_paragraph_formatting(
            para,
            {'alignment': 'justify'},
            'Normal', 0, 1
        )
        assert len(issues) == 0

    def test_check_font_wrong_name(self):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Test')
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        issues = auditor._check_paragraph_formatting(
            para,
            {'alignment': 'justify'},
            'Normal', 0, 1
        )
        assert len(issues) > 0
        assert issues[0].category == 'PARAGRAPH'

    def test_check_font_wrong_size(self):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Test')
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        issues = auditor._check_paragraph_formatting(
            para,
            {'alignment': 'justify'},
            'Normal', 0, 1
        )
        assert len(issues) > 0

    def test_check_paragraph_alignment_correct(self):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Test')
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        issues = auditor._check_paragraph_formatting(
            para,
            {'alignment': 'justify'},
            'Normal', 0, 1
        )
        assert len(issues) == 0

    def test_check_paragraph_alignment_wrong(self):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        auditor = BaseAuditor(self.config)
        doc = Document()
        para = doc.add_paragraph('Test')
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        issues = auditor._check_paragraph_formatting(
            para,
            {'alignment': 'justify'},
            'Normal', 0, 1
        )
        assert len(issues) > 0
        assert issues[0].category == 'PARAGRAPH'

#!/usr/bin/env python3
"""Тесты для ReportManager."""

import os
import tempfile
import shutil
import pytest
from src.core.report_manager import ReportManager
from src.core.audit_engine import AuditIssue, Severity


def make_issues():
    """Создаёт тестовые проблемы."""
    return [
        AuditIssue(
            id='FONT_1', severity=Severity.CRITICAL, category='FONT',
            element_type='Normal', location={'index': 0},
            description='Неверный шрифт', current_value='Arial',
            expected_value='Times New Roman', auto_fixable=True
        ),
        AuditIssue(
            id='PARA_1', severity=Severity.WARNING, category='PARAGRAPH',
            element_type='Heading 1', location={'index': 5},
            description='Неверный отступ', current_value='0',
            expected_value='1.25', auto_fixable=True
        ),
        AuditIssue(
            id='TABLE_1', severity=Severity.INFO, category='TABLE',
            element_type='Table Grid', location={'table_index': 0},
            description='Информация о таблице', current_value='',
            expected_value='', auto_fixable=False
        ),
    ]


class TestReportManager:
    """Тесты для ReportManager."""

    def setup_method(self):
        self.temp_dir = tempfile.mkdtemp()
        self.issues = make_issues()
        self.manager = ReportManager(
            self.issues,
            doc_path='test.docx',
            config_path='config.yaml'
        )

    def teardown_method(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_filter_by_severity(self):
        critical = self.manager.filter_by_severity(Severity.CRITICAL)
        assert len(critical) == 1
        assert critical[0].id == 'FONT_1'

    def test_filter_by_category(self):
        font_issues = self.manager.filter_by_category('FONT')
        assert len(font_issues) == 1

    def test_apply_filter_severity(self):
        filtered = self.manager.apply_filter(severity=Severity.WARNING)
        assert len(filtered) == 1
        assert filtered[0].id == 'PARA_1'

    def test_apply_filter_category(self):
        filtered = self.manager.apply_filter(category='TABLE')
        assert len(filtered) == 1
        assert filtered[0].id == 'TABLE_1'

    def test_apply_filter_combined(self):
        filtered = self.manager.apply_filter(
            severity=Severity.CRITICAL,
            category='FONT'
        )
        assert len(filtered) == 1

    def test_exclude_issues(self):
        self.manager.exclude_issues(['FONT_1'])
        filtered = self.manager.apply_filter()
        assert len(filtered) == 2

    def test_include_issues(self):
        self.manager.exclude_issues(['FONT_1'])
        self.manager.include_issues(['FONT_1'])
        filtered = self.manager.apply_filter()
        assert len(filtered) == 3

    def test_get_summary(self):
        summary = self.manager.get_summary()
        assert summary['CRITICAL'] == 1
        assert summary['WARNING'] == 1
        assert summary['INFO'] == 1

    def test_get_summary_excluded(self):
        self.manager.exclude_issues(['FONT_1'])
        summary = self.manager.get_summary()
        assert summary['CRITICAL'] == 0

    def test_group_by_severity(self):
        groups = self.manager.group_by_severity()
        assert Severity.CRITICAL in groups
        assert len(groups[Severity.CRITICAL]) == 1

    def test_group_by_category(self):
        groups = self.manager.group_by_category()
        assert 'FONT' in groups
        assert 'PARAGRAPH' in groups

    def test_group_by_element_type(self):
        groups = self.manager.group_by_element_type()
        assert 'Normal' in groups
        assert 'Heading 1' in groups

    def test_group_by_category_and_element(self):
        groups = self.manager.group_by_category_and_element()
        assert 'FONT: Normal' in groups

    def test_export_txt(self):
        path = os.path.join(self.temp_dir, 'report.txt')
        self.manager.export_txt(path)
        assert os.path.exists(path)

        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        assert 'ОТЧЁТ АУДИТА' in content
        assert 'FONT_1' in content

    def test_export_html(self):
        path = os.path.join(self.temp_dir, 'report.html')
        self.manager.export_html(path)
        assert os.path.exists(path)

        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        assert '<html>' in content
        assert 'Неверный шрифт' in content
        assert 'CRITICAL' in content

    def test_export_csv(self):
        path = os.path.join(self.temp_dir, 'report.csv')
        self.manager.export_csv(path)
        assert os.path.exists(path)

        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        assert 'FONT_1' in content
        assert 'CRITICAL' in content

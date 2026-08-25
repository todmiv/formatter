"""
Модуль для управления отчётами аудита: группировка, фильтрация, экспорт.
"""
import os
import datetime
from typing import List, Dict, Any, Optional
from dataclasses import dataclass
from enum import Enum

from src.core.audit_engine import AuditIssue, Severity


class ReportManager:
    """
    Управляет отчётами по проблемам аудита.
    """
    def __init__(self, issues: List[AuditIssue], doc_path: str, config_path: str):
        self.issues = issues
        self.doc_path = doc_path
        self.config_path = config_path
        self.filtered_issues = issues.copy()
        self.excluded_ids = set()  # ID проблем, исключённых из отчёта

    def filter_by_severity(self, severity: Severity) -> List[AuditIssue]:
        """
        Возвращает проблемы с заданной важностью.
        """
        return [issue for issue in self.issues if issue.severity == severity]

    def filter_by_category(self, category: str) -> List[AuditIssue]:
        """
        Возвращает проблемы по категории.
        """
        return [issue for issue in self.issues if issue.category == category]

    def apply_filter(self, severity: Optional[Severity] = None,
                     category: Optional[str] = None) -> List[AuditIssue]:
        """
        Применяет фильтры и возвращает отфильтрованный список.
        """
        filtered = self.issues
        if severity is not None:
            filtered = [i for i in filtered if i.severity == severity]
        if category is not None:
            filtered = [i for i in filtered if i.category == category]
        # Исключаем проблемы, помеченные как исключённые
        filtered = [i for i in filtered if i.id not in self.excluded_ids]
        self.filtered_issues = filtered
        return filtered

    def exclude_issues(self, issue_ids: List[str]) -> None:
        """
        Исключает проблемы из отчёта.
        """
        self.excluded_ids.update(issue_ids)

    def include_issues(self, issue_ids: List[str]) -> None:
        """
        Возвращает проблемы в отчёт.
        """
        self.excluded_ids.difference_update(issue_ids)

    def get_summary(self) -> Dict[str, int]:
        """
        Возвращает сводку по всем проблемам (включая исключённые?).
        """
        summary = {s.value: 0 for s in Severity}
        for issue in self.issues:
            if issue.id not in self.excluded_ids:
                summary[issue.severity.value] += 1
        return summary

    def group_by_severity(self) -> Dict[Severity, List[AuditIssue]]:
        """
        Группирует проблемы по важности.
        """
        groups = {s: [] for s in Severity}
        for issue in self.issues:
            if issue.id not in self.excluded_ids:
                groups[issue.severity].append(issue)
        return groups

    def group_by_category(self) -> Dict[str, List[AuditIssue]]:
        """
        Группирует проблемы по категории.
        """
        groups = {}
        for issue in self.issues:
            if issue.id not in self.excluded_ids:
                cat = issue.category
                groups.setdefault(cat, []).append(issue)
        return groups

    def group_by_element_type(self) -> Dict[str, List[AuditIssue]]:
        """
        Группирует проблемы по типу элемента (стилю).
        """
        groups = {}
        for issue in self.issues:
            if issue.id not in self.excluded_ids:
                elem = issue.element_type
                groups.setdefault(elem, []).append(issue)
        return groups

    def group_by_category_and_element(self) -> Dict[str, List[AuditIssue]]:
        """
        Группирует проблемы по комбинации категория + элемент.
        Ключ: "Категория: Элемент"
        """
        groups = {}
        for issue in self.issues:
            if issue.id not in self.excluded_ids:
                key = f"{issue.category}: {issue.element_type}"
                groups.setdefault(key, []).append(issue)
        return groups

    def export_txt(self, file_path: str) -> None:
        """
        Экспортирует отчёт в текстовый файл.
        """
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("ОТЧЁТ АУДИТА ДОКУМЕНТА\n")
            f.write("=" * 50 + "\n")
            f.write(f"Документ: {os.path.basename(self.doc_path)}\n")
            f.write(f"Конфигурация: {self.config_path}\n")
            f.write(f"Дата аудита: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("\n")

            summary = self.get_summary()
            f.write("СВОДКА:\n")
            f.write(f"  Всего проблем: {sum(summary.values())}\n")
            f.write(f"  Критических: {summary['CRITICAL']}\n")
            f.write(f"  Предупреждений: {summary['WARNING']}\n")
            f.write(f"  Информационных: {summary['INFO']}\n")
            f.write("\n")

            f.write("ДЕТАЛИ ПРОБЛЕМ:\n")
            f.write("-" * 50 + "\n")
            for idx, issue in enumerate(self.issues, 1):
                if issue.id in self.excluded_ids:
                    continue
                f.write(f"{idx}. ID: {issue.id}\n")
                f.write(f"   Важность: {issue.severity.value}\n")
                f.write(f"   Категория: {issue.category}\n")
                f.write(f"   Стиль: {issue.element_type}\n")
                f.write(f"   Локация: стр. ~{issue.location.get('page', '?')}\n")
                f.write(f"   Описание: {issue.description}\n")
                f.write(f"   Текущее значение: {issue.current_value}\n")
                f.write(f"   Ожидаемое значение: {issue.expected_value}\n")
                f.write(f"   Автоисправление: {'Да' if issue.auto_fixable else 'Нет'}\n")
                f.write(f"   Рекомендация: {issue.recommendation if issue.recommendation else '(не указана)'}\n")
                f.write("\n")

    def export_html(self, file_path: str) -> None:
        """
        Экспортирует отчёт в HTML файл.
        """
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Отчёт аудита документа</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .critical { background-color: #ffcccc; }
        .warning { background-color: #fff3cd; }
        .info { background-color: #d1ecf1; }
    </style>
</head>
<body>
    <h1>ОТЧЁТ АУДИТА ДОКУМЕНТА</h1>
    <p><strong>Документ:</strong> """)
            f.write(os.path.basename(self.doc_path))
            f.write("""</p>
    <p><strong>Конфигурация:</strong> """)
            f.write(self.config_path)
            f.write("""</p>
    <p><strong>Дата аудита:</strong> """)
            f.write(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            f.write("""</p>
    
    <h2>Сводка</h2>
    <ul>
        <li>Всего проблем: """)
            total = sum(self.get_summary().values())
            f.write(str(total))
            f.write("""</li>
        <li>Критических: """)
            f.write(str(self.get_summary()['CRITICAL']))
            f.write("""</li>
        <li>Предупреждений: """)
            f.write(str(self.get_summary()['WARNING']))
            f.write("""</li>
        <li>Информационных: """)
            f.write(str(self.get_summary()['INFO']))
            f.write("""</li>
    </ul>
    
    <h2>Детали проблем</h2>
    <table>
        <thead>
            <tr>
                <th>#</th>
                <th>Важность</th>
                <th>Категория</th>
                <th>Стиль</th>
                <th>Локация</th>
                <th>Описание</th>
                <th>Текущее значение</th>
                <th>Ожидаемое значение</th>
                <th>Автоисправление</th>
                <th>Рекомендация</th>
            </tr>
        </thead>
        <tbody>
""")
            for idx, issue in enumerate(self.issues, 1):
                if issue.id in self.excluded_ids:
                    continue
                f.write(f"            <tr class='{issue.severity.value.lower()}'>\n")
                f.write(f"                <td>{idx}</td>\n")
                f.write(f"                <td>{issue.severity.value}</td>\n")
                f.write(f"                <td>{issue.category}</td>\n")
                f.write(f"                <td>{issue.element_type}</td>\n")
                f.write(f"                <td>стр. ~{issue.location.get('page', '?')}</td>\n")
                f.write(f"                <td>{issue.description}</td>\n")
                f.write(f"                <td>{issue.current_value}</td>\n")
                f.write(f"                <td>{issue.expected_value}</td>\n")
                f.write(f"                <td>{'Да' if issue.auto_fixable else 'Нет'}</td>\n")
                f.write(f"                <td>{issue.recommendation if issue.recommendation else ''}</td>\n")
                f.write("            </tr>\n")
            f.write("""        </tbody>
    </table>
</body>
</html>""")

    def export_csv(self, file_path: str) -> None:
        """
        Экспортирует отчёт в CSV файл.
        """
        import csv
        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['ID', 'Severity', 'Category', 'Element', 'Location', 'Description',
                          'Current', 'Expected', 'AutoFixable', 'Recommendation']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for issue in self.issues:
                if issue.id in self.excluded_ids:
                    continue
                writer.writerow({
                    'ID': issue.id,
                    'Severity': issue.severity.value,
                    'Category': issue.category,
                    'Element': issue.element_type,
                    'Location': f"стр. ~{issue.location.get('page', '?')}",
                    'Description': issue.description,
                    'Current': issue.current_value,
                    'Expected': issue.expected_value,
                    'AutoFixable': 'Да' if issue.auto_fixable else 'Нет',
                    'Recommendation': issue.recommendation if issue.recommendation else ''
                })
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ExportMixin — экспорт отчётов.
Вынесена из GOSTFormatterGUI для уменьшения размера класса.
"""

import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox


class ExportMixin:
    """
    Миксин для экспорта отчётов.
    Требует атрибутов: controller, log.
    """

    def export_report(self):
        """Экспорт отчёта об аудите."""
        if not self.controller.current_issues:
            messagebox.showwarning("Внимание", "Нет данных для экспорта")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("Text Files", "*.txt"),
                ("HTML Files", "*.html"),
                ("CSV Files", "*.csv"),
            ],
            title="Сохранить отчёт"
        )

        if not file_path:
            return

        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == '.html':
                self._export_report_html(file_path)
            elif ext == '.csv':
                self._export_report_csv(file_path)
            else:
                self._export_report_txt(file_path)

            self.log(f"Отчёт экспортирован: {file_path}")
            messagebox.showinfo("Готово", f"Отчёт сохранён:\n{file_path}")
        except Exception as e:
            self.log(f"Ошибка экспорта: {e}")
            messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{e}")

    def _export_report_txt(self, file_path: str):
        """Генерация текстового отчёта."""
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("ОТЧЁТ АУДИТА ДОКУМЕНТА\n")
            f.write("=" * 50 + "\n")
            f.write(f"Документ: {os.path.basename(self.controller.doc_path)}\n")
            f.write(f"Конфигурация: {self.controller.config_path}\n")
            f.write(f"Дата: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("\n")

            summary = self.controller.get_audit_summary()
            f.write("СВОДКА:\n")
            f.write(f"  Всего проблем: {len(self.controller.current_issues)}\n")
            f.write(f"  Критических: {summary.get('CRITICAL', 0)}\n")
            f.write(f"  Предупреждений: {summary.get('WARNING', 0)}\n")
            f.write(f"  Информационных: {summary.get('INFO', 0)}\n")
            f.write("\n")
            f.write("ДЕТАЛИ:\n")
            f.write("-" * 50 + "\n")

            for issue in self.controller.current_issues:
                f.write(f"[{issue.severity.value}] {issue.category}: {issue.description}\n")
                f.write(f"  Стиль: {issue.element_type}\n")
                f.write(f"  Текущее: {issue.current_value}\n")
                f.write(f"  Ожидается: {issue.expected_value}\n")
                f.write("\n")

    def _export_report_html(self, file_path: str):
        """Генерация HTML отчёта."""
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("<!DOCTYPE html>\n<html>\n<head>\n")
            f.write("<meta charset='utf-8'>\n")
            f.write("<title>Отчёт аудита</title>\n")
            f.write("<style>\n")
            f.write("body { font-family: Arial; margin: 20px; }\n")
            f.write("table { border-collapse: collapse; width: 100%; }\n")
            f.write("th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }\n")
            f.write("th { background-color: #4CAF50; color: white; }\n")
            f.write(".critical { background-color: #ffcccc; }\n")
            f.write(".warning { background-color: #fff3cd; }\n")
            f.write(".info { background-color: #d1ecf1; }\n")
            f.write("</style>\n</head>\n<body>\n")

            f.write("<h1>Отчёт аудита документа</h1>\n")
            f.write(f"<p><b>Документ:</b> {os.path.basename(self.controller.doc_path)}</p>\n")
            f.write(f"<p><b>Дата:</b> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>\n")

            f.write("<table>\n<tr>\n")
            f.write("<th>Важность</th><th>Категория</th><th>Стиль</th><th>Описание</th>\n")
            f.write("</tr>\n")

            for issue in self.controller.current_issues:
                severity = issue.severity.value.lower()
                f.write(f"<tr class='{severity}'>\n")
                f.write(f"<td>{issue.severity.value}</td>\n")
                f.write(f"<td>{issue.category}</td>\n")
                f.write(f"<td>{issue.element_type}</td>\n")
                f.write(f"<td>{issue.description}</td>\n")
                f.write("</tr>\n")

            f.write("</table>\n</body>\n</html>")

    def _export_report_csv(self, file_path: str):
        """Генерация CSV отчёта."""
        import csv

        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Важность', 'Категория', 'Стиль', 'Описание', 'Текущее', 'Ожидается'])

            for issue in self.controller.current_issues:
                writer.writerow([
                    issue.severity.value,
                    issue.category,
                    issue.element_type,
                    issue.description,
                    issue.current_value,
                    issue.expected_value,
                ])

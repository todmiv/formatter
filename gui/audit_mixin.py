#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AuditMixin — логика аудита и исправления ошибок.
Вынесена из GOSTFormatterGUI для уменьшения размера класса.
"""

import threading
import tkinter as tk
from tkinter import messagebox
from typing import List


class AuditMixin:
    """
    Миксин для работы с аудитом и исправлениями.
    Требует атрибутов: controller, status_var, progress_var, log, root.
    """

    def start_audit_thread(self):
        """Запуск аудита в отдельном потоке."""
        if not self.controller.doc_path:
            self.log("Документ не выбран")
            return

        self.status_var.set("Запуск аудита...")
        self.progress_var.set(0)

        thread = threading.Thread(target=self._run_audit, daemon=True)
        thread.start()

    def _run_audit(self):
        """Выполнение аудита (в отдельном потоке)."""
        try:
            issues = self.controller.start_audit()
            self.root.after(0, lambda: self._on_audit_complete(issues))
        except Exception as e:
            self.root.after(0, lambda: self._on_audit_error(e))

    def _on_audit_complete(self, issues):
        """Обработка завершения аудита."""
        self.status_var.set(f"Аудит завершён: {len(issues)} проблем")
        self.progress_var.set(100)
        self.log(f"Аудит завершён: {len(issues)} проблем")

        if hasattr(self, 'btn_fix_all'):
            self.btn_fix_all.config(state='normal')
        if hasattr(self, 'btn_fix_selected'):
            self.btn_fix_selected.config(state='normal')
        if hasattr(self, 'btn_fix_styles_only'):
            self.btn_fix_styles_only.config(state='normal')
        if hasattr(self, 'btn_export_report'):
            self.btn_export_report.config(state='normal')

    def _on_audit_error(self, error):
        """Обработка ошибки аудита."""
        self.status_var.set("Ошибка аудита")
        self.log(f"Ошибка аудита: {error}")
        messagebox.showerror("Ошибка", f"Ошибка аудита:\n{error}")

    def fix_selected(self):
        """Исправление выбранных проблем."""
        selected = self.get_selected_issues()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите проблемы для исправления")
            return

        self.status_var.set("Применение исправлений...")
        thread = threading.Thread(
            target=self._apply_fixes_thread,
            args=(selected,),
            daemon=True
        )
        thread.start()

    def fix_styles_only(self):
        """Исправление только стилей."""
        selected = self.get_selected_issues()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите проблемы для исправления")
            return

        self.status_var.set("Применение стилей...")
        thread = threading.Thread(
            target=self._apply_fixes_thread,
            args=(selected, True),
            daemon=True
        )
        thread.start()

    def fix_all(self):
        """Исправление всех проблем."""
        self.status_var.set("Исправление всех проблем...")
        thread = threading.Thread(
            target=self._apply_all_fixes_thread,
            daemon=True
        )
        thread.start()

    def _apply_fixes_thread(self, issue_ids, styles_only=False):
        """Поток применения исправлений."""
        try:
            stats = self.controller.apply_fixes(issue_ids)
            self.root.after(0, lambda: self._on_fix_complete(stats))
        except Exception as e:
            self.root.after(0, lambda: self._on_fix_error(e))

    def _apply_all_fixes_thread(self):
        """Поток применения всех исправлений."""
        try:
            issue_ids = [i.id for i in self.controller.current_issues]
            stats = self.controller.apply_fixes(issue_ids)
            self.root.after(0, lambda: self._on_fix_complete(stats))
        except Exception as e:
            self.root.after(0, lambda: self._on_fix_error(e))

    def _on_fix_complete(self, stats):
        """Обработка завершения исправлений."""
        applied = stats.get('applied', 0)
        failed = stats.get('failed', 0)
        self.status_var.set(f"Исправлено: {applied}, ошибок: {failed}")
        self.log(f"Исправления применены: {applied} успешно, {failed} ошибок")

        if hasattr(self, 'btn_save'):
            self.btn_save.config(state='normal')

    def _on_fix_error(self, error):
        """Обработка ошибки исправления."""
        self.status_var.set("Ошибка исправления")
        self.log(f"Ошибка: {error}")
        messagebox.showerror("Ошибка", f"Ошибка исправления:\n{error}")

    def apply_filter(self, event=None):
        """Применение фильтра к таблице проблем."""
        if not hasattr(self, 'tree'):
            return

        filter_value = self.filter_var.get() if hasattr(self, 'filter_var') else 'ALL'

        for item in self.tree.get_children():
            self.tree.delete(item)

        for issue in self.controller.current_issues:
            if filter_value != 'ALL' and issue.severity.value != filter_value:
                continue

            severity = issue.severity.value
            tag = severity.lower() if hasattr(self, 'tree') else ''

            self.tree.insert('', 'end', values=(
                issue.id,
                severity,
                issue.category,
                issue.element_type,
                str(issue.location.get('index', '')),
                issue.description[:50],
                issue.current_value[:30],
                issue.expected_value[:30],
            ), tags=(tag,))

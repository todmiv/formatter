#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HighlightMixin — подсветка и поиск в документе.
Вынесена из GOSTFormatterGUI для уменьшения размера класса.
"""

import threading
import tkinter as tk
from tkinter import messagebox


class HighlightMixin:
    """
    Миксин для подсветки и поиска в документе.
    Требует атрибутов: controller, status_var, log, root.
    """

    def highlight_issues(self):
        """Подсветка проблемных мест в документе."""
        if not self.controller.doc_path:
            messagebox.showwarning("Внимание", "Документ не выбран")
            return

        self.status_var.set("Создание подсветки...")
        thread = threading.Thread(target=self._highlight_thread, daemon=True)
        thread.start()

    def _highlight_thread(self):
        """Поток подсветки."""
        try:
            highlighted_path = self.controller.highlight_issues()
            self.root.after(0, lambda: self._on_highlight_complete(highlighted_path))
        except Exception as e:
            self.root.after(0, lambda: self._on_highlight_error(e))

    def _on_highlight_complete(self, path):
        """Обработка завершения подсветки."""
        self.status_var.set("Подсветка создана")
        if path:
            self.log(f"Подсветка: {path}")
        else:
            self.log("Подсветка не создана (нет проблем)")

    def _on_highlight_error(self, error):
        """Обработка ошибки подсветки."""
        self.status_var.set("Ошибка подсветки")
        self.log(f"Ошибка подсветки: {error}")

    def find_in_document(self):
        """Поиск выбранной проблемы в документе."""
        selected = self.get_selected_issues()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите проблему для поиска")
            return

        issue = selected[0]
        self.status_var.set("Поиск в документе...")

        thread = threading.Thread(
            target=self._find_thread,
            args=(issue,),
            daemon=True
        )
        thread.start()

    def _find_thread(self, issue):
        """Поток поиска."""
        try:
            path = self.controller.find_issue_in_document(issue)
            self.root.after(0, lambda: self._on_find_complete(path))
        except Exception as e:
            self.root.after(0, lambda: self._on_find_error(e))

    def _on_find_complete(self, path):
        """Обработка завершения поиска."""
        self.status_var.set("Документ открыт")
        if path:
            self.log(f"Найдено: {path}")

    def _on_find_error(self, error):
        """Обработка ошибки поиска."""
        self.status_var.set("Ошибка поиска")
        self.log(f"Ошибка поиска: {error}")

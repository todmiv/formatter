"""Дерево конфига — левая панель редактора."""

import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional


class TreePanel(ttk.Frame):
    def __init__(self, parent, model, on_select: Callable):
        super().__init__(parent)
        self.model = model
        self.on_select = on_select

        ttk.Label(self, text="Конфигурация", font=("", 10, "bold")).pack(anchor=tk.W, padx=5, pady=5)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        ttk.Button(btn_frame, text="Развернуть", command=self.expand_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Свернуть", command=self.collapse_all).pack(side=tk.LEFT, padx=2)

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, show="tree")
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

    def build_tree(self):
        self.tree.delete(*self.tree.get_children())
        root_id = self.tree.insert("", tk.END, text="Конфигурация", open=True)

        styles_id = self.tree.insert(root_id, tk.END, text="Стили", open=True)
        for name in sorted(self.model.styles.keys()):
            style = self.model.styles[name]
            prefix = "✓ " if style.enabled else "✗ "
            self.tree.insert(styles_id, tk.END, text=f"{prefix}{name}",
                             values=("style", name))

        for section_key, label in [
            ("detection_rules", "Правила определения"),
            ("page_setup", "Настройки страницы"),
            ("headers_footers", "Колонтитулы"),
            ("formatting_rules", "Правила форматирования"),
            ("validation", "Валидация"),
            ("federal_regional_projects", "Федеральные/региональные проекты"),
        ]:
            if section_key in self.model.sections:
                self.tree.insert(root_id, tk.END, text=label,
                                 values=("section", section_key))

        self.tree.item(root_id, open=True)

    def expand_all(self):
        for item in self.tree.get_children(""):
            self._expand_recursive(item)

    def _expand_recursive(self, item):
        self.tree.item(item, open=True)
        for child in self.tree.get_children(item):
            self._expand_recursive(child)

    def collapse_all(self):
        for item in self.tree.get_children(""):
            self.tree.item(item, open=False)
            for child in self.tree.get_children(child):
                self.tree.item(child, open=False)

    def _on_select(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        item = selection[0]
        values = self.tree.item(item, "values")
        text = self.tree.item(item, "text")

        if values and len(values) >= 2:
            kind, key = values[0], values[1]
            if kind == "style":
                self.on_select("style", key)
            elif kind == "section":
                self.on_select("section", key)
        elif text == "Стили":
            self.on_select("styles_root", None)
        elif text == "Конфигурация":
            self.on_select("root", None)

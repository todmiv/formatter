"""Панель редактирования — правая панель с полями."""

import tkinter as tk
from tkinter import ttk
from typing import Any, Callable, Dict, List

from .field_widgets import FieldWidgetFactory
from .model import FieldDescriptor, StyleDescriptor


class EditorPanel(ttk.Frame):
    def __init__(self, parent, on_field_change: Callable = None):
        super().__init__(parent)
        self.on_field_change = on_field_change
        self._widgets: Dict[str, tk.Widget] = {}
        self._current_target = None
        self._canvas = None
        self._inner_frame = None

        self._build_scrollable()

    def _build_scrollable(self):
        self._canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self._canvas.yview)
        self._inner_frame = ttk.Frame(self._canvas)

        self._inner_frame.bind("<Configure>",
                              lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.create_window((0, 0), window=self._inner_frame, anchor=tk.NW)
        self._canvas.configure(yscrollcommand=scrollbar.set)

        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._canvas.bind_all("<MouseWheel>",
                             lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))

    def clear(self):
        for w in self._inner_frame.winfo_children():
            w.destroy()
        self._widgets.clear()
        self._current_target = None

    def show_style(self, style: StyleDescriptor):
        self.clear()
        self._current_target = ("style", style.name)

        ttk.Label(self._inner_frame, text=f"Стиль: {style.name}",
                  font=("", 11, "bold")).grid(row=0, column=0, columnspan=3,
                                               sticky=tk.W, padx=5, pady=5)

        self._build_enabled_row(style)

        row = 2
        if style.font:
            row = self._build_section("Шрифт", style.font, row, "font")
        if style.paragraph:
            row = self._build_section("Абзац", style.paragraph, row, "paragraph")
        if style.table:
            row = self._build_section("Таблица", style.table, row, "table")
        if style.position:
            row = self._build_section("Позиция", style.position, row, "position")
        if style.content:
            row = self._build_section("Содержимое", style.content, row, "content")

    def _build_enabled_row(self, style: StyleDescriptor):
        enabled_var = tk.BooleanVar(value=style.enabled)
        cb = ttk.Checkbutton(self._inner_frame, text="Включён", variable=enabled_var)
        cb.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=2)

        def _on_toggle(*_):
            style.enabled = enabled_var.get()
            if self.on_field_change:
                self.on_field_change("style_enabled", style.name, enabled_var.get())

        enabled_var.trace_add("write", _on_toggle)

    def _build_section(self, title: str, fields: Dict[str, FieldDescriptor],
                       start_row: int, section_name: str) -> int:
        frame = ttk.LabelFrame(self._inner_frame, text=title, padding=5)
        frame.grid(row=start_row, column=0, columnspan=3, sticky=tk.EW,
                   padx=5, pady=5)
        self._inner_frame.columnconfigure(0, weight=1)

        row = 0
        for key, fd in fields.items():
            ttk.Label(frame, text=fd.label).grid(row=row, column=0, sticky=tk.W, pady=2, padx=2)

            widget = FieldWidgetFactory.create(frame, fd,
                                               on_change=lambda k, v, sn=section_name: self._on_change(sn, k, v))
            widget.grid(row=row, column=1, sticky=tk.W, pady=2, padx=5)

            if fd.unit:
                ttk.Label(frame, text=fd.unit).grid(row=row, column=2, sticky=tk.W, pady=2)
            self._widgets[f"{section_name}.{key}"] = widget
            row += 1

        return start_row + 1

    def _on_change(self, section: str, key: str, value: Any):
        if self.on_field_change:
            self.on_field_change(section, key, value)

    def show_section(self, section_name: str, data: Any):
        self.clear()
        self._current_target = ("section", section_name)

        labels = {
            "detection_rules": "Правила определения",
            "page_setup": "Настройки страницы",
            "headers_footers": "Колонтитулы",
            "formatting_rules": "Правила форматирования",
            "validation": "Валидация",
            "federal_regional_projects": "Федеральные/региональные проекты",
        }
        ttk.Label(self._inner_frame, text=labels.get(section_name, section_name),
                  font=("", 11, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

        self._text = tk.Text(self._inner_frame, wrap=tk.WORD, width=60, height=25)
        self._text.grid(row=1, column=0, sticky=tk.NSEW, padx=5, pady=5)
        self._inner_frame.rowconfigure(1, weight=1)
        self._inner_frame.columnconfigure(0, weight=1)

        import yaml
        text = yaml.dump(data, default_flow_style=False, allow_unicode=True, sort_keys=False)
        self._text.insert("1.0", text)

    def get_section_data(self) -> Any:
        if self._current_target and self._current_target[0] == "section":
            import yaml
            text = self._text.get("1.0", tk.END)
            try:
                return yaml.safe_load(text)
            except Exception:
                return None
        return None

    def show_root(self):
        self.clear()
        self._current_target = ("root", None)
        ttk.Label(self._inner_frame, text="Конфигурация ГОСТ",
                  font=("", 12, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=10)
        ttk.Label(self._inner_frame, text="Выберите элемент в дереве для редактирования.",
                  foreground="gray").grid(row=1, column=0, sticky=tk.W, padx=5)

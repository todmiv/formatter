#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SettingsMixin — настройки допусков и игнорирования.
Вынесена из GOSTFormatterGUI для уменьшения размера класса.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List

from src.core.docx_utils import MM_TO_TWIPS, PT_TO_TWIPS


class SettingsMixin:
    """
    Миксин для настроек допусков и игнорирования.
    Требует атрибутов: controller, log, root.
    """

    def open_settings(self):
        """Открывает окно настроек допусков."""
        settings_win = tk.Toplevel(self.root)
        settings_win.title("Настройки допусков")
        settings_win.geometry("450x350")
        settings_win.transient(self.root)
        settings_win.grab_set()

        ttk.Label(settings_win, text="Допуски для сравнения свойств").pack(pady=10)

        settings_frame = ttk.Frame(settings_win, padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=True)

        def create_tolerance_row(parent, label, row, default_value):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=5)

            var_mm = tk.StringVar(value=f"{default_value / MM_TO_TWIPS:.2f}")
            var_pt = tk.StringVar(value=f"{default_value / PT_TO_TWIPS:.2f}")
            var_twips = tk.StringVar(value=str(default_value))

            ttk.Label(parent, text="мм:").grid(row=row, column=1, padx=5)
            entry_mm = ttk.Entry(parent, textvariable=var_mm, width=8)
            entry_mm.grid(row=row, column=2, padx=5)

            ttk.Label(parent, text="pt:").grid(row=row, column=3, padx=5)
            entry_pt = ttk.Entry(parent, textvariable=var_pt, width=8)
            entry_pt.grid(row=row, column=4, padx=5)

            ttk.Label(parent, text="twips:").grid(row=row, column=5, padx=5)
            entry_twips = ttk.Entry(parent, textvariable=var_twips, width=8)
            entry_twips.grid(row=row, column=6, padx=5)

            def update_from_mm(*args):
                try:
                    mm = float(var_mm.get())
                    var_pt.set(f"{mm * MM_TO_TWIPS / PT_TO_TWIPS:.2f}")
                    var_twips.set(str(int(round(mm * MM_TO_TWIPS))))
                except ValueError:
                    pass

            def update_from_pt(*args):
                try:
                    pt = float(var_pt.get())
                    var_mm.set(f"{pt * PT_TO_TWIPS / MM_TO_TWIPS:.2f}")
                    var_twips.set(str(int(round(pt * PT_TO_TWIPS))))
                except ValueError:
                    pass

            var_mm.trace_add('write', update_from_mm)
            var_pt.trace_add('write', update_from_pt)

            return var_twips

        font_var = create_tolerance_row(settings_frame, "Размер шрифта:", 0, self.controller.font_tolerance)
        indent_var = create_tolerance_row(settings_frame, "Отступ:", 1, self.controller.indent_tolerance)
        spacing_var = create_tolerance_row(settings_frame, "Интервал:", 2, self.controller.spacing_tolerance)

        btn_frame = ttk.Frame(settings_win)
        btn_frame.pack(pady=10)

        def save_settings():
            try:
                new_font = int(font_var.get())
                new_indent = int(indent_var.get())
                new_spacing = int(spacing_var.get())

                self.controller.set_tolerances(
                    font=new_font,
                    indent=new_indent,
                    spacing=new_spacing
                )

                self.log(f"Допуски обновлены: font={new_font}, indent={new_indent}, spacing={new_spacing}")
                messagebox.showinfo("Сохранено", "Допуски успешно сохранены.")
                settings_win.destroy()
            except ValueError as e:
                messagebox.showerror("Ошибка", f"Некорректное значение: {e}")

        ttk.Button(btn_frame, text="Сохранить", command=save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=settings_win.destroy).pack(side=tk.LEFT, padx=5)

    def open_ignore_settings(self):
        """Открывает окно настроек игнорирования типов ошибок."""
        ignore_win = tk.Toplevel(self.root)
        ignore_win.title("Настройки игнорирования")
        ignore_win.geometry("400x350")
        ignore_win.transient(self.root)
        ignore_win.grab_set()

        ttk.Label(ignore_win, text="Выберите типы ошибок для игнорирования").pack(pady=10)

        categories_frame = ttk.LabelFrame(ignore_win, text="Категории", padding=10)
        categories_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        all_categories = ['FONT', 'PARAGRAPH', 'STYLE', 'TABLE', 'HEADER_FOOTER', 'PAGE']
        cat_vars = {}

        for i, cat in enumerate(all_categories):
            var = tk.BooleanVar(value=cat in self.controller.ignored_categories)
            cat_vars[cat] = var
            ttk.Checkbutton(categories_frame, text=cat, variable=var).grid(
                row=i // 3, column=i % 3, sticky=tk.W, padx=5
            )

        btn_frame = ttk.Frame(ignore_win)
        btn_frame.pack(pady=10)

        def save_ignore_settings():
            selected = [cat for cat, var in cat_vars.items() if var.get()]
            self.controller.set_ignore_settings(categories=selected)
            self.log(f"Игнорируемые категории: {selected}")
            messagebox.showinfo("Сохранено", "Настройки игнорирования сохранены.")
            ignore_win.destroy()

        ttk.Button(btn_frame, text="Сохранить", command=save_ignore_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=ignore_win.destroy).pack(side=tk.LEFT, padx=5)

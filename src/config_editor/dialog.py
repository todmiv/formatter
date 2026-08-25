"""Главное диалоговое окно редактора конфигурации."""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .model import ConfigModel
from .tree_panel import TreePanel
from .editor_panel import EditorPanel
from .style_preview import StylePreview
from .validators import Validators
from .yaml_utils import YamlUtils


class ConfigEditorDialog:
    def __init__(self, parent, config_path: str = None, on_close=None):
        self.parent = parent
        self.config_path = config_path or "configs/active/config.yaml"
        self.is_saved = False
        self._on_close_callback = on_close

        self.window = tk.Toplevel(parent)
        self.window.title(f"Редактор конфигураций — {os.path.basename(self.config_path)}")
        self.window.geometry("1200x800")
        self.window.minsize(900, 600)
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

        self.model = ConfigModel(self.config_path)
        self.model.load()

        self._build_ui()
        self._refresh_tree()
        self.window.deiconify()
        self.window.update_idletasks()

    def _build_ui(self):
        toolbar = ttk.Frame(self.window, padding=5)
        toolbar.pack(fill=tk.X)

        ttk.Button(toolbar, text="Открыть", command=self._open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Сохранить", command=self._save).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Сохранить как...", command=self._save_as).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Проверить", command=self._validate).pack(side=tk.LEFT, padx=2)

        self.lbl_status = ttk.Label(toolbar, text="", foreground="gray")
        self.lbl_status.pack(side=tk.RIGHT, padx=5)

        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_frame = ttk.Frame(main_frame, width=280)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        left_frame.pack_propagate(False)

        self.tree_panel = TreePanel(left_frame, self.model, self._on_tree_select)
        self.tree_panel.pack(fill=tk.BOTH, expand=True)

        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        editor_frame = ttk.Frame(right_frame)
        editor_frame.pack(fill=tk.BOTH, expand=True)

        self.editor = EditorPanel(editor_frame, on_field_change=self._on_field_change)
        self.editor.pack(fill=tk.BOTH, expand=True)

        preview_frame = ttk.LabelFrame(right_frame, text="Предпросмотр стиля", padding=5)
        preview_frame.pack(fill=tk.X, pady=(5, 0))

        self.preview = StylePreview(preview_frame)
        self.preview.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(self.window, padding=5)
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="Сохранить", command=self._save).pack(side=tk.RIGHT, padx=2)
        ttk.Button(btn_frame, text="Отмена", command=self._on_close).pack(side=tk.RIGHT, padx=2)

    def _refresh_tree(self):
        self.tree_panel.build_tree()

    def _on_tree_select(self, kind: str, key: str):
        if kind == "style":
            style = self.model.get_style(key)
            if style:
                self.editor.show_style(style)
                self.preview.update_preview(style)
        elif kind == "section":
            data = self.model.sections.get(key, {})
            self.editor.show_section(key, data)
            self.preview.clear()
        elif kind == "root":
            self.editor.show_root()
            self.preview.clear()
        elif kind == "styles_root":
            self.editor.show_root()
            self.preview.clear()

    def _on_field_change(self, section: str, key: str, value):
        if section == "style_enabled":
            self.model.set_style_enabled(key, value)
            self._refresh_tree()
            return
        style_name = self._get_selected_style_name()
        if style_name and section in ("font", "paragraph", "table"):
            self.model.update_style_field(style_name, section, key, value)
            style = self.model.get_style(style_name)
            if style:
                self.preview.update_preview(style)

    def _get_selected_style_name(self) -> str:
        sel = self.tree_panel.tree.selection()
        if sel:
            values = self.tree_panel.tree.item(sel[0], "values")
            if values and len(values) >= 2 and values[0] == "style":
                return values[1]
        return None

    def _open_file(self):
        if self.model.modified:
            if not messagebox.askyesno("Внимание", "Есть несохранённые изменения. Открыть новый файл?"):
                return
        path = filedialog.askopenfilename(
            title="Открыть конфигурацию",
            filetypes=[("YAML", "*.yaml *.yml"), ("JSON", "*.json"), ("Все", "*.*")],
            initialdir=os.path.dirname(self.config_path),
        )
        if path:
            self.config_path = path
            self.model = ConfigModel(path)
            self.model.load()
            self.tree_panel.model = self.model
            self._refresh_tree()
            self.editor.clear()
            self.preview.clear()
            self.window.title(f"Редактор конфигураций — {os.path.basename(path)}")
            self._set_status("Загружено")

    def _save(self):
        section_data = self.editor.get_section_data()
        if section_data is not None and self.editor._current_target:
            section_name = self.editor._current_target[1]
            self.model.update_section(section_name, section_data)

        path = self.config_path
        if self.model.save(path):
            self.is_saved = True
            self._set_status("Сохранено")
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить файл")

    def _save_as(self):
        section_data = self.editor.get_section_data()
        if section_data is not None and self.editor._current_target:
            section_name = self.editor._current_target[1]
            self.model.update_section(section_name, section_data)

        path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".yaml",
            filetypes=[("YAML", "*.yaml"), ("Все", "*.*")],
            initialdir=os.path.dirname(self.config_path),
            initialfile=os.path.basename(self.config_path),
        )
        if path:
            self.config_path = path
            self.model.config_path = path
            if self.model.save(path):
                self.is_saved = True
                self._set_status("Сохранено")
                self.window.title(f"Редактор конфигураций — {os.path.basename(path)}")

    def _validate(self):
        errors = Validators.validate_all(self.model)
        if not errors:
            messagebox.showinfo("Валидация", "Ошибок не найдено!")
            return
        msg = "\n".join(f"[{e.severity}] {e.field_path}: {e.message}" for e in errors)
        messagebox.showwarning("Результаты валидации", msg)

    def _set_status(self, text: str):
        self.lbl_status.config(text=text)

    def _on_close(self):
        if self.model.modified:
            if messagebox.askyesno("Выход", "Есть несохранённые изменения. Сохранить?"):
                self._save()
        self.window.destroy()
        if self._on_close_callback:
            self._on_close_callback(self.is_saved)

    def show(self):
        pass

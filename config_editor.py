"""
ConfigEditorWindow — визуальный редактор YAML-конфигураций ГОСТ.
Отдельное окно с вкладками для редактирования каждого раздела конфига.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import yaml
import os
import copy
import logging

logger = logging.getLogger(__name__)


class ConfigEditorWindow:
    """Окно редактора конфигурации ГОСТ."""

    def __init__(self, root, config_path: str = None):
        self.root = root
        self.config_path = config_path or "configs/active/config.yaml"
        self.config_data = {}
        self.modified = False

        self.window = tk.Toplevel(root)
        self.window.title(f"Редактор конфигурации — {os.path.basename(self.config_path)}")
        self.window.geometry("950x700")
        self.window.minsize(800, 500)

        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self._load_config()

    def _build_ui(self):
        # --- Toolbar ---
        toolbar = ttk.Frame(self.window, padding=5)
        toolbar.pack(fill=tk.X)

        ttk.Button(toolbar, text="📂 Открыть", command=self._open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="💾 Сохранить", command=self._save_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="💾 Сохранить как…", command=self._save_file_as).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="↩ Сбросить", command=self._reload).pack(side=tk.LEFT, padx=2)

        self.lbl_status = ttk.Label(toolbar, text="", foreground="gray")
        self.lbl_status.pack(side=tk.RIGHT, padx=5)

        # --- Notebook (вкладки) ---
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self._build_general_tab()
        self._build_styles_tab()
        self._build_detection_tab()
        self._build_page_setup_tab()
        self._build_headers_footers_tab()
        self._build_formatting_tab()

    # =====================================================================
    # Общие
    # =====================================================================
    def _build_general_tab(self):
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Общие")

        fields = [
            ("version", "Версия:"),
            ("name", "Название:"),
            ("description", "Описание:"),
            ("last_modified", "Дата изменения:"),
        ]
        self.general_entries = {}
        for i, (key, label) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=4)
            entry = ttk.Entry(frame, width=60)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=5, pady=4)
            self.general_entries[key] = entry

        frame.columnconfigure(1, weight=1)

    # =====================================================================
    # Стили
    # =====================================================================
    def _build_styles_tab(self):
        frame = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(frame, text="Стили")

        # Левая панель — список стилей
        left = ttk.LabelFrame(frame, text="Стили документа", padding=5)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))

        self.style_listbox = tk.Listbox(left, width=28, height=30)
        self.style_listbox.pack(fill=tk.BOTH, expand=True)
        self.style_listbox.bind("<<ListboxSelect>>", self._on_style_select)

        btn_frame = ttk.Frame(left)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_frame, text="＋ Добавить", command=self._add_style).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="✕ Удалить", command=self._delete_style).pack(side=tk.LEFT, padx=2)

        # Правая панель — редактор свойств стиля
        right = ttk.LabelFrame(frame, text="Свойства стиля", padding=10)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.style_notebook = ttk.Notebook(right)
        self.style_notebook.pack(fill=tk.BOTH, expand=True)

        self._build_style_font_tab()
        self._build_style_paragraph_tab()
        self._build_style_table_tab()

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)

    def _build_style_font_tab(self):
        frame = ttk.Frame(self.style_notebook, padding=10)
        self.style_notebook.add(frame, text="Шрифт")

        self.style_font_entries = {}
        fields = [
            ("font.name", "Имя шрифта:", "Times New Roman"),
            ("font.size", "Размер (пт):", "12"),
            ("font.bold", "Жирный:", "false"),
            ("font.italic", "Курсив:", "false"),
            ("font.color", "Цвет:", "black"),
        ]
        for i, (key, label, default) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=4)
            entry = ttk.Entry(frame, width=30)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=5, pady=4)
            self.style_font_entries[key] = entry

        frame.columnconfigure(1, weight=1)

    def _build_style_paragraph_tab(self):
        frame = ttk.Frame(self.style_notebook, padding=10)
        self.style_notebook.add(frame, text="Абзац")

        self.style_para_entries = {}
        fields = [
            ("paragraph.first_line_indent_cm", "Отступ 1-й строки (см):", "1.25"),
            ("paragraph.line_spacing", "Межстрочный интервал:", "1.15"),
            ("paragraph.space_before_pt", "Интервал перед (пт):", "0"),
            ("paragraph.space_after_pt", "Интервал после (пт):", "0"),
            ("paragraph.alignment", "Выравнивание:", "justify"),
        ]
        for i, (key, label, default) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=4)
            entry = ttk.Entry(frame, width=30)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=5, pady=4)
            self.style_para_entries[key] = entry

        frame.columnconfigure(1, weight=1)

    def _build_style_table_tab(self):
        frame = ttk.Frame(self.style_notebook, padding=10)
        self.style_notebook.add(frame, text="Таблица")

        self.style_table_entries = {}
        fields = [
            ("table.auto_fit", "Автоподбор:", "window"),
            ("table.fixed_column_width", "Фикс. ширина:", "true"),
            ("table.row_height_cm", "Высота строки (см):", "0"),
            ("table.row_height_mode", "Режим высоты:", "minimum"),
            ("table.border_style", "Стиль границ:", "single"),
            ("table.border_width_pt", "Ширина границ (пт):", "0.5"),
        ]
        for i, (key, label, default) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=4)
            entry = ttk.Entry(frame, width=30)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=5, pady=4)
            self.style_table_entries[key] = entry

        frame.columnconfigure(1, weight=1)

    # =====================================================================
    # Правила определения
    # =====================================================================
    def _build_detection_tab(self):
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Правила определения")

        ttk.Label(frame, text="Шаблоны для автоматического определения стилей (по одному на строку):").pack(anchor=tk.W)

        self.detection_text = tk.Text(frame, wrap=tk.WORD, width=80, height=30)
        self.detection_text.pack(fill=tk.BOTH, expand=True, pady=5)

    # =====================================================================
    # Настройки страницы
    # =====================================================================
    def _build_page_setup_tab(self):
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Страница")

        self.page_entries = {}
        fields = [
            ("portrait.left_margin_cm", "Книжная — левое поле (см):", "2.5"),
            ("portrait.right_margin_cm", "Книжная — правое поле (см):", "1.0"),
            ("portrait.top_margin_cm", "Книжная — верхнее поле (см):", "2.0"),
            ("portrait.bottom_margin_cm", "Книжная — нижнее поле (см):", "2.0"),
            ("landscape.left_margin_cm", "Альбомная — левое поле (см):", "2.0"),
            ("landscape.right_margin_cm", "Альбомная — правое поле (см):", "2.0"),
            ("landscape.top_margin_cm", "Альбомная — верхнее поле (см):", "2.5"),
            ("landscape.bottom_margin_cm", "Альбомная — нижнее поле (см):", "1.0"),
            ("paper_size", "Формат бумаги:", "A4"),
            ("orientation_default", "Ориентация по умолчанию:", "portrait"),
        ]
        for i, (key, label, default) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=4)
            entry = ttk.Entry(frame, width=20)
            entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=4)
            self.page_entries[key] = entry

    # =====================================================================
    # Колонтитулы
    # =====================================================================
    def _build_headers_footers_tab(self):
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Колонтитулы")

        # Header
        hf_frame = ttk.LabelFrame(frame, text="Верхний колонтитул (Header)", padding=10)
        hf_frame.pack(fill=tk.X, pady=(0, 10))

        self.hf_entries = {}
        fields = [
            ("header.enabled", "Включён:", "true"),
            ("header.position_from_edge_cm", "Положение от края (см):", "1.0"),
            ("header.content", "Содержимое:", "page_number"),
            ("header.alignment", "Выравнивание:", "right"),
            ("header.font_size_min", "Мин. размер шрифта:", "12"),
            ("header.font_size_max", "Макс. размер шрифта:", "14"),
            ("header.arabic_numerals", "Арабские цифры:", "true"),
            ("header.no_page_word", "Без слова 'страница':", "true"),
            ("header.no_punctuation", "Без знаков препинания:", "true"),
            ("header.continuous_numbering", "Сквозная нумерация:", "true"),
        ]
        for i, (key, label, default) in enumerate(fields):
            ttk.Label(hf_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2)
            entry = ttk.Entry(hf_frame, width=20)
            entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
            self.hf_entries[key] = entry

        # Footer
        ff_frame = ttk.LabelFrame(frame, text="Нижний колонтитул (Footer)", padding=10)
        ff_frame.pack(fill=tk.X)

        fields_footer = [
            ("footer.enabled", "Включён:", "true"),
            ("footer.position_from_edge_cm", "Положение от края (см):", "1.0"),
            ("footer.content", "Содержимое:", "section_chapter_number"),
            ("footer.alignment", "Выравнивание:", "center"),
            ("footer.style", "Стиль:", "uppercase_spaced"),
            ("footer.bold", "Жирный:", "true"),
            ("footer.character_spacing", "Межбуквенный интервал:", "1.5"),
            ("footer.section_break_independent", "Независимые разделы:", "true"),
        ]
        for i, (key, label, default) in enumerate(fields_footer):
            ttk.Label(ff_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2)
            entry = ttk.Entry(ff_frame, width=20)
            entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
            self.hf_entries[key] = entry

    # =====================================================================
    # Правила форматирования
    # =====================================================================
    def _build_formatting_tab(self):
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Форматирование")

        self.formatting_text = tk.Text(frame, wrap=tk.WORD, width=80, height=30)
        self.formatting_text.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(frame, text="Секция formatting_rules редактируется как YAML-текст.").pack(anchor=tk.W)

    # =====================================================================
    # Загрузка / Сохранение
    # =====================================================================
    def _load_config(self):
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config_data = yaml.safe_load(f) or {}
            self._populate_fields()
            self.modified = False
            self._update_status("Загружено")
            logger.info(f"Конфиг загружен: {self.config_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить конфиг:\n{e}")
            logger.error(f"Ошибка загрузки конфига: {e}")

    def _populate_fields(self):
        cfg = self.config_data

        # General
        for key, entry in self.general_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, str(cfg.get(key, "")))

        # Styles list
        self.style_listbox.delete(0, tk.END)
        styles = cfg.get('styles', {})
        for name in styles:
            self.style_listbox.insert(tk.END, name)
        self._selected_style = None

        # Detection rules
        self.detection_text.delete("1.0", tk.END)
        detection = cfg.get('detection_rules', {})
        self.detection_text.insert("1.0", yaml.dump(detection, default_flow_style=False, allow_unicode=True))

        # Page setup
        page = cfg.get('page_setup', {})
        for key, entry in self.page_entries.items():
            entry.delete(0, tk.END)
            val = self._get_nested(page, key)
            entry.insert(0, str(val) if val is not None else "")

        # Headers/footers
        hf = cfg.get('headers_footers', {})
        for key, entry in self.hf_entries.items():
            entry.delete(0, tk.END)
            val = self._get_nested(hf, key)
            entry.insert(0, str(val) if val is not None else "")

        # Formatting rules
        self.formatting_text.delete("1.0", tk.END)
        fmt = cfg.get('formatting_rules', {})
        self.formatting_text.insert("1.0", yaml.dump(fmt, default_flow_style=False, allow_unicode=True))

    def _get_nested(self, d, key):
        """Получает значение по точечному ключу (например, 'portrait.left_margin_cm')."""
        keys = key.split('.')
        for k in keys:
            if isinstance(d, dict) and k in d:
                d = d[k]
            else:
                return None
        return d

    def _collect_fields(self):
        """Собирает данные из полей обратно в config_data."""
        cfg = self.config_data

        # General
        for key, entry in self.general_entries.items():
            cfg[key] = entry.get().strip()

        # Styles
        styles = cfg.get('styles', {})
        if self._selected_style and self._selected_style in styles:
            style = styles[self._selected_style]
            self._collect_style_entries(style)

        # Detection rules
        try:
            cfg['detection_rules'] = yaml.safe_load(self.detection_text.get("1.0", tk.END)) or {}
        except yaml.YAMLError as e:
            messagebox.showwarning("Внимание", f"Ошибка YAML в правилах определения:\n{e}")

        # Page setup
        page = cfg.get('page_setup', {})
        for key, entry in self.page_entries.items():
            self._set_nested(page, key, entry.get().strip())
        cfg['page_setup'] = page

        # Headers/footers
        hf = cfg.get('headers_footers', {})
        for key, entry in self.hf_entries.items():
            self._set_nested(hf, key, entry.get().strip())
        cfg['headers_footers'] = hf

        # Formatting rules
        try:
            cfg['formatting_rules'] = yaml.safe_load(self.formatting_text.get("1.0", tk.END)) or {}
        except yaml.YAMLError as e:
            messagebox.showwarning("Внимание", f"Ошибка YAML в правилах форматирования:\n{e}")

    def _set_nested(self, d, key, value):
        """Устанавливает значение по точечному ключу с авто-конвертацией типов."""
        keys = key.split('.')
        for k in keys[:-1]:
            if k not in d or not isinstance(d[k], dict):
                d[k] = {}
            d = d[k]

        final_key = keys[-1]
        d[final_key] = self._auto_convert(value)

    def _auto_convert(self, value: str):
        """Автоматическое преобразование строки в подходящий тип."""
        if value.lower() in ('true', 'false'):
            return value.lower() == 'true'
        try:
            return int(value)
        except ValueError:
            pass
        try:
            return float(value)
        except ValueError:
            pass
        return value

    def _collect_style_entries(self, style):
        """Собирает данные полей стиля в словарь."""
        # Font
        font = style.get('font', {})
        for key in self.style_font_entries:
            parts = key.split('.', 1)
            section = parts[1] if len(parts) > 1 else parts[0]
            entry = self.style_font_entries[key]
            val = entry.get().strip()
            if val:
                font[section] = self._auto_convert(val)
        if font:
            style['font'] = font

        # Paragraph
        para = style.get('paragraph', {})
        for key in self.style_para_entries:
            parts = key.split('.', 1)
            section = parts[1] if len(parts) > 1 else parts[0]
            entry = self.style_para_entries[key]
            val = entry.get().strip()
            if val:
                para[section] = self._auto_convert(val)
        if para:
            style['paragraph'] = para

        # Table
        tbl = style.get('table', {})
        for key in self.style_table_entries:
            parts = key.split('.', 1)
            section = parts[1] if len(parts) > 1 else parts[0]
            entry = self.style_table_entries[key]
            val = entry.get().strip()
            if val:
                tbl[section] = self._auto_convert(val)
        if tbl:
            style['table'] = tbl

    def _populate_style_entries(self, style):
        """Заполняет поля стиля из словаря."""
        font = style.get('font', {})
        for key in self.style_font_entries:
            parts = key.split('.', 1)
            section = parts[1] if len(parts) > 1 else parts[0]
            entry = self.style_font_entries[key]
            entry.delete(0, tk.END)
            entry.insert(0, str(font.get(section, "")))

        para = style.get('paragraph', {})
        for key in self.style_para_entries:
            parts = key.split('.', 1)
            section = parts[1] if len(parts) > 1 else parts[0]
            entry = self.style_para_entries[key]
            entry.delete(0, tk.END)
            entry.insert(0, str(para.get(section, "")))

        tbl = style.get('table', {})
        for key in self.style_table_entries:
            parts = key.split('.', 1)
            section = parts[1] if len(parts) > 1 else parts[0]
            entry = self.style_table_entries[key]
            entry.delete(0, tk.END)
            entry.insert(0, str(tbl.get(section, "")))

    # =====================================================================
    # Обработчики событий
    # =====================================================================
    def _on_style_select(self, event):
        selection = self.style_listbox.curselection()
        if not selection:
            return

        # Сохраняем текущий стиль перед переключением
        if self._selected_style:
            self._collect_fields()

        idx = selection[0]
        self._selected_style = self.style_listbox.get(idx)
        styles = self.config_data.get('styles', {})
        if self._selected_style in styles:
            self._populate_style_entries(styles[self._selected_style])

    def _add_style(self):
        name = tk.simpledialog.askstring("Новый стиль", "Введите имя стиля:")
        if not name:
            return
        if 'styles' not in self.config_data:
            self.config_data['styles'] = {}
        self.config_data['styles'][name] = {
            'enabled': True,
            'font': {'name': 'Times New Roman', 'size': 12, 'bold': False, 'italic': False},
            'paragraph': {'first_line_indent_cm': 0, 'line_spacing': 1.0, 'alignment': 'left'},
        }
        self.style_listbox.insert(tk.END, name)
        self.modified = True
        self._update_status("Изменено")

    def _delete_style(self):
        selection = self.style_listbox.curselection()
        if not selection:
            return
        name = self.style_listbox.get(selection[0])
        if messagebox.askyesno("Удаление", f"Удалить стиль '{name}'?"):
            self.style_listbox.delete(selection[0])
            self.config_data.get('styles', {}).pop(name, None)
            self._selected_style = None
            self.modified = True
            self._update_status("Изменено")

    def _open_file(self):
        if self.modified:
            if not messagebox.askyesno("Внимание", "Есть несохранённые изменения. Открыть новый файл?"):
                return
        path = filedialog.askopenfilename(
            title="Открыть конфигурацию",
            filetypes=[("YAML", "*.yaml *.yml"), ("JSON", "*.json"), ("Все", "*.*")],
            initialdir=os.path.dirname(self.config_path)
        )
        if path:
            self.config_path = path
            self.window.title(f"Редактор конфигурации — {os.path.basename(path)}")
            self._load_config()

    def _save_file(self):
        self._collect_fields()
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(self.config_data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
            self.modified = False
            self._update_status("Сохранено")
            logger.info(f"Конфиг сохранён: {self.config_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    def _save_file_as(self):
        self._collect_fields()
        path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".yaml",
            filetypes=[("YAML", "*.yaml"), ("Все", "*.*")],
            initialdir=os.path.dirname(self.config_path),
            initialfile=os.path.basename(self.config_path)
        )
        if path:
            self.config_path = path
            self._save_file()
            self.window.title(f"Редактор конфигурации — {os.path.basename(path)}")

    def _reload(self):
        if self.modified:
            if not messagebox.askyesno("Внимание", "Отменить несохранённые изменения?"):
                return
        self._load_config()

    def _on_close(self):
        if self.modified:
            if messagebox.askyesno("Выход", "Есть несохранённые изменения. Сохранить?"):
                self._save_file()
        self.window.destroy()

    def _update_status(self, text):
        self.lbl_status.config(text=text)


def open_config_editor(root, config_path=None):
    """Открывает окно редактора конфигурации."""
    return ConfigEditorWindow(root, config_path)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    open_config_editor(root)
    root.mainloop()

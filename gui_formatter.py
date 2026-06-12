# gui_formatter.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import threading
import os
import datetime
import logging
import shutil
import time
import docx

# Импорт новых модулей ядра
from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine, Severity
from src.apply.apply_orchestrator import ApplyOrchestrator
from src.core.report_manager import ReportManager
from md_to_docx import RI2013Converter
from src.core.highlight_engine import HighlightEngine
from src.core.docx_utils import open_document, MM_TO_TWIPS, PT_TO_TWIPS
from src.core.main_controller import MainController
from src.core.document_model import DocumentModel

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("formatter.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class GOSTFormatterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Форматировщик документов по ГОСТ (Audit & Apply)")
        self.root.geometry("1000x700")

        # Конфигурация по умолчанию
        self.config_path = "configs/active/config.yaml"

        # Инициализация контроллера
        self.controller = MainController(self.config_path)
        # Установка колбэков для обновления UI
        self.controller.set_callbacks(
            progress_cb=self.update_audit_progress,
            status_cb=lambda msg: self.root.after(0, lambda m=msg: self.status_var.set(m)) if hasattr(self, 'status_var') and self.root else None,
            log_cb=self.log
        )

        # Состояние сортировки таблицы (остаётся в GUI)
        self.sort_column = None
        self.sort_reverse = False

        # Флаги GUI
        self.is_scanning = False

        self._init_ui()
        # Загрузка конфига уже выполнена в контроллере, но можно обновить UI
        self._sync_ui_with_controller()

    def _init_ui(self):
        # --- Верхняя панель: Файлы и Конфиг ---
        top_frame = ttk.LabelFrame(self.root, text="Настройки", padding=10)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        # Выбор файла
        ttk.Label(top_frame, text="Документ:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.entry_file = ttk.Entry(top_frame, width=50)
        self.entry_file.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="Обзор...", command=self.browse_file).grid(row=0, column=2, pady=5)
        ttk.Button(top_frame, text="Открыть для просмотра", command=self.open_document_for_view).grid(row=0, column=3, padx=(5,0), pady=5)
        ttk.Button(top_frame, text=" Импорт из Markdown", command=self.import_markdown).grid(row=0, column=4, padx=(5,0), pady=5)

        # Выбор конфига
        ttk.Label(top_frame, text="Конфиг:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.entry_config = ttk.Entry(top_frame, width=50)
        self.entry_config.insert(0, self.config_path)
        self.entry_config.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="Обзор...", command=self.browse_config).grid(row=1, column=2, pady=5)
        ttk.Button(top_frame, text="Открыть для редактирования", command=self.open_config_editor_window).grid(row=1, column=3, padx=(5,0), pady=5)

        # --- Панель действий ---
        action_frame = ttk.Frame(self.root, padding=5)
        action_frame.pack(fill=tk.X, padx=10, pady=5)

        self.btn_audit = ttk.Button(action_frame, text="🔍 Начать аудит", command=self.start_audit_thread)
        self.btn_audit.pack(side=tk.LEFT, padx=5)

        self.btn_highlight = ttk.Button(action_frame, text="🔦 Подсветить проблемные места", command=self.highlight_issues)
        self.btn_highlight.pack(side=tk.LEFT, padx=5)

        self.btn_fix_selected = ttk.Button(action_frame, text="🛠 Исправить выбранные", command=self.fix_selected, state=tk.DISABLED)
        self.btn_fix_selected.pack(side=tk.LEFT, padx=5)

        self.btn_fix_styles_only = ttk.Button(action_frame, text="🎨 Исправить только стили", command=self.fix_styles_only, state=tk.DISABLED)
        self.btn_fix_styles_only.pack(side=tk.LEFT, padx=5)

        self.btn_fix_all = ttk.Button(action_frame, text="⚡ Исправить всё", command=self.fix_all, state=tk.DISABLED)
        self.btn_fix_all.pack(side=tk.LEFT, padx=5)

        self.btn_export_report = ttk.Button(action_frame, text="📄 Экспорт отчёта", command=self.export_report, state=tk.DISABLED)
        self.btn_export_report.pack(side=tk.LEFT, padx=5)

        self.btn_settings = ttk.Button(action_frame, text="⚙ Настройки", command=self.open_settings)
        self.btn_settings.pack(side=tk.LEFT, padx=5)

        self.btn_ignore_settings = ttk.Button(action_frame, text="👁 Игнорировать ошибки", command=self.open_ignore_settings)
        self.btn_ignore_settings.pack(side=tk.LEFT, padx=5)

        self.btn_monitor = ttk.Button(action_frame, text="📊 Мониторинг", command=self.open_monitor_window)
        self.btn_monitor.pack(side=tk.LEFT, padx=5)

        self.btn_save = ttk.Button(action_frame, text="💾 Сохранить результат", command=self.save_document, state=tk.DISABLED)
        self.btn_save.pack(side=tk.RIGHT, padx=5)

        # --- Статус бар ---
        self.status_var = tk.StringVar(value="Готов к работе. Выберите документ.")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # Веса этапов аудита (в процентах от общего прогресса)
        self.audit_stage_weights = {
            'page_setup': 5,
            'paragraphs': 70,
            'tables': 10,
            'headers_footers': 10,
            'images': 5
        }
        # Текущий прогресс по этапам (накопленная сумма весов завершённых этапов)
        self.audit_stage_completed = 0
        # Текущий этап (для отображения)
        self.current_stage = ''

        # --- Основная область: Трёхпанельный интерфейс ---
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # --- Левая панель: Дерево документа ---
        left_frame = ttk.LabelFrame(main_paned, text="Структура документа", padding=5)
        self.doc_tree = ttk.Treeview(left_frame, columns=("type", "text"), show="tree", selectmode="browse")
        self.doc_tree.heading("#0", text="Элемент")
        self.doc_tree.column("#0", width=200)
        self.doc_tree.column("type", width=80, stretch=False)
        self.doc_tree.column("text", width=200)
        doc_tree_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.doc_tree.yview)
        self.doc_tree.configure(yscrollcommand=doc_tree_scroll.set)
        self.doc_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        doc_tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        main_paned.add(left_frame, weight=1)

        # --- Центральная панель: Таблица проблем ---
        center_frame = ttk.LabelFrame(main_paned, text="Отчет аудита (Найденные несоответствия)", padding=5)
        
        # Фильтры и кнопки над таблицей
        filter_frame = ttk.Frame(center_frame)
        filter_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(filter_frame, text="Фильтр:").pack(side=tk.LEFT)
        self.filter_var = tk.StringVar(value="ALL")
        cb_filter = ttk.Combobox(filter_frame, textvariable=self.filter_var, values=["ALL", "CRITICAL", "WARNING", "INFO"], state="readonly", width=15)
        cb_filter.pack(side=tk.LEFT, padx=5)
        cb_filter.bind("<<ComboboxSelected>>", self.apply_filter)
        
        # Чекбокс "Выделить все"
        self.select_all_var = tk.BooleanVar(value=False)
        select_all_cb = ttk.Checkbutton(filter_frame, text="Выделить все", variable=self.select_all_var, command=self.toggle_select_all)
        select_all_cb.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(filter_frame, text="Снять выделение", command=self.deselect_all).pack(side=tk.RIGHT, padx=5)
        ttk.Button(filter_frame, text="Исключить из отчёта", command=self.exclude_selected).pack(side=tk.RIGHT, padx=5)

        # Контейнер для таблицы
        tree_container = ttk.Frame(center_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        # Таблица (Treeview) с чекбоксами (пока без чекбоксов, добавим позже)
        columns = ("id", "severity", "category", "element", "location", "description", "current", "expected")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="extended")
        
        # Настройка заголовков с поддержкой сортировки
        self.tree.heading("id", text="#", command=lambda: self.treeview_sort_column("id", False))
        self.tree.column("id", width=30)
        
        self.tree.heading("severity", text="Важность", command=lambda: self.treeview_sort_column("severity", False))
        self.tree.column("severity", width=80)
        
        self.tree.heading("category", text="Категория", command=lambda: self.treeview_sort_column("category", False))
        self.tree.column("category", width=80)
        
        self.tree.heading("element", text="Стиль", command=lambda: self.treeview_sort_column("element", False))
        self.tree.column("element", width=100)
        
        self.tree.heading("location", text="Локация", command=lambda: self.treeview_sort_column("location", False))
        self.tree.column("location", width=80)
        
        self.tree.heading("description", text="Проблема", command=lambda: self.treeview_sort_column("description", False))
        self.tree.column("description", width=200)
        
        self.tree.heading("current", text="Текущее", command=lambda: self.treeview_sort_column("current", False))
        self.tree.column("current", width=120)
        
        self.tree.heading("expected", text="Ожидается", command=lambda: self.treeview_sort_column("expected", False))
        self.tree.column("expected", width=120)

        # Скроллбары
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Размещение внутри tree_container с grid
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # Привязка события выбора для обновления деталей
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        # Настройка тегов для цветовой кодировки строк
        self.tree.tag_configure('critical', background='#ffcccc', foreground='black')
        self.tree.tag_configure('warning', background='#fff3cd', foreground='black')
        self.tree.tag_configure('info', background='#d1ecf1', foreground='black')

        main_paned.add(center_frame, weight=3)

        # --- Правая панель: Детали ошибки и предпросмотр ---
        right_frame = ttk.LabelFrame(main_paned, text="Детали ошибки", padding=5)
        
        # Детали
        details_label = ttk.Label(right_frame, text="Выберите ошибку для просмотра деталей", wraplength=250)
        details_label.pack(pady=5)
        
        self.details_text = scrolledtext.ScrolledText(right_frame, height=10, state='disabled')
        self.details_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # Кнопка "Найти в документе"
        self.btn_find_in_doc = ttk.Button(right_frame, text="🔍 Найти в документе", command=self.find_in_document, state=tk.DISABLED)
        self.btn_find_in_doc.pack(pady=5)
        
        # Предпросмотр
        preview_label = ttk.Label(right_frame, text="Предпросмотр фрагмента")
        preview_label.pack(pady=(10,0))
        
        self.preview_text = scrolledtext.ScrolledText(right_frame, height=8, state='disabled')
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        
        main_paned.add(right_frame, weight=2)

        # --- Лог операций (под трёхпанельным интерфейсом) ---
        log_frame = ttk.LabelFrame(self.root, text="Лог операций", padding=5)
        log_frame.pack(fill=tk.X, padx=10, pady=(0,5))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=4, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _load_config(self):
        """Предварительная загрузка конфига для проверки (оставлено для совместимости)"""
        # Конфиг уже загружен в контроллере, но можно проверить наличие файла
        if os.path.exists(self.config_path):
            self.log("Конфигурация загружена успешно.")
        else:
            self.log("Файл конфигурации не найден. Используйте стандартный или выберите другой.")
    
    def _sync_ui_with_controller(self):
        """Синхронизирует UI с состоянием контроллера (например, допуски)."""
        # Устанавливаем допуски из контроллера в локальные переменные для UI
        self.font_tolerance = self.controller.font_tolerance
        self.indent_tolerance = self.controller.indent_tolerance
        self.spacing_tolerance = self.controller.spacing_tolerance
        # Синхронизируем путь конфига
        self.config_path = self.controller.config_path
        # Обновляем поле конфига, если нужно
        if hasattr(self, 'entry_config'):
            self.entry_config.delete(0, tk.END)
            self.entry_config.insert(0, self.controller.config_path)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")])
        if filename:
            try:
                self.controller.set_document(filename)
                self.entry_file.delete(0, tk.END)
                self.entry_file.insert(0, filename)
                self.status_var.set(f"Файл выбран: {os.path.basename(filename)}")
                # Сброс состояния при выборе нового файла
                self.clear_results()
                # Заполняем дерево документа (заглушка)
                self.populate_doc_tree(filename)
            except Exception as e:
                self.log(f"Ошибка установки документа: {e}")
                messagebox.showerror("Ошибка", f"Не удалось загрузить документ:\n{e}")

    def import_markdown(self):
        """Импорт Markdown файла и конвертация в DOCX с последующей загрузкой в интерфейс."""
        # Выбор исходного .md файла
        md_path = filedialog.askopenfilename(
            filetypes=[("Markdown Files", "*.md"), ("Text Files", "*.txt"), ("All Files", "*.*")],
            title="Выберите Markdown файл для импорта"
        )
        if not md_path:
            return  # пользователь отменил
        
        # Выбор пути сохранения .docx
        default_name = os.path.splitext(os.path.basename(md_path))[0] + "_converted.docx"
        docx_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word Documents", "*.docx")],
            title="Сохранить как DOCX"
        )
        if not docx_path:
            return
        
        # Конфигурация для конвертера (используем текущий конфиг или стандартный)
        config_path = self.entry_config.get().strip()
        if not config_path or not os.path.exists(config_path):
            config_path = "configs/active/config.yaml"
        
        # Запуск конвертации в отдельном потоке, чтобы не блокировать GUI
        self.status_var.set("Конвертация Markdown в DOCX...")
        self.progress_var.set(0)
        self.progress_bar.configure(mode='determinate', maximum=100)
        self.log(f"Начата конвертация: {os.path.basename(md_path)} -> {os.path.basename(docx_path)}")
        
        # Событие для синхронизации завершения потока
        completion_event = threading.Event()
        result_container = {'success': False, 'error': None, 'docx_path': None}
        timeout_occurred = False
        # Динамический таймаут: 5 минут на большие файлы
        timeout_seconds = 300
        
        # Throttling для callback'ов прогресса (не чаще 1 раза в 200 мс)
        _last_update_time = [0.0]
        
        def conversion_progress_callback(stage, progress, total, message):
            """Callback прогресса конвертации — обновляет статус в GUI."""
            if stage == 'processing' and total > 0:
                percent = int(progress * 100 / total)
                # Throttling: обновляем GUI не чаще 1 раза в 200 мс
                now = time.perf_counter()
                if now - _last_update_time[0] < 0.2 and percent < 100:
                    return
                _last_update_time[0] = now
                self.root.after(0, lambda p=percent: (
                    self.status_var.set(f"Конвертация Markdown: {p}%"),
                    self.progress_var.set(p)
                ))
            elif stage == 'save':
                self.root.after(0, lambda: (
                    self.status_var.set("Сохранение DOCX файла..."),
                    self.progress_var.set(95)
                ))
            elif stage == 'done':
                self.root.after(0, lambda: (
                    self.status_var.set("Конвертация завершена"),
                    self.progress_var.set(100)
                ))
        
        def convert_thread():
            nonlocal timeout_occurred
            try:
                # Используем контроллер для импорта с callback прогресса
                success = self.controller.import_markdown(
                    md_path, docx_path,
                    progress_callback=conversion_progress_callback
                )
                if timeout_occurred:
                    return  # таймаут уже сработал, игнорируем
                result_container['success'] = success
                result_container['docx_path'] = docx_path
            except Exception as e:
                if timeout_occurred:
                    return
                result_container['success'] = False
                result_container['error'] = e
            finally:
                completion_event.set()
        
        def check_completion():
            nonlocal timeout_occurred
            if completion_event.is_set():
                # Поток завершился — обрабатываем результат
                if result_container['success']:
                    self.on_conversion_success(result_container['docx_path'])
                else:
                    error = result_container['error'] or RuntimeError("Контроллер вернул неудачу")
                    self.on_conversion_error(error)
            elif timeout_occurred:
                # Таймаут сработал, но поток мог уже завершиться
                if completion_event.is_set() and result_container['success']:
                    # Поток успел завершиться — игнорируем таймаут
                    self.on_conversion_success(result_container['docx_path'])
                # Иначе ничего не делаем — таймаут уже обработан
            else:
                # Проверяем снова через 200мс (реже, чтобы меньше нагружать)
                self.root.after(200, check_completion)
        
        # Запускаем поток
        thread = threading.Thread(target=convert_thread, daemon=True)
        thread.start()
        
        # Запускаем таймер таймаута
        def on_timeout():
            nonlocal timeout_occurred
            if completion_event.is_set():
                # Поток уже завершился, таймаут не нужен
                return
            timeout_occurred = True
            self._on_conversion_timeout(docx_path, timeout_seconds)
        
        self.root.after(timeout_seconds * 1000, on_timeout)
        
        # Запускаем проверку завершения
        self.root.after(200, check_completion)

    def _on_conversion_timeout(self, docx_path, timeout_seconds):
        """Обработка таймаута конвертации."""
        self.progress_bar.stop()
        self.progress_bar.configure(mode='determinate')
        self.progress_var.set(0)
        self.status_var.set("Ошибка: таймаут конвертации")
        error_msg = (
            f"Конвертация Markdown в DOCX не завершилась за {timeout_seconds} секунд.\n"
            f"Файл назначения: {docx_path}\n\n"
            f"Возможные причины:\n"
            f"• Слишком большой Markdown-файл\n"
            f"• Отсутствует конфигурационный файл\n"
            f"• Проблемы с доступом к файловой системе\n\n"
            f"Попробуйте:\n"
            f"• Уменьшить размер файла\n"
            f"• Проверить наличие configs/active/config.yaml\n"
            f"• Запустить конвертацию через CLI: python md_to_docx.py <file.md> <file.docx>"
        )
        self.log(f"Ошибка: таймаут конвертации ({timeout_seconds}с): {docx_path}")
        messagebox.showerror("Таймаут конвертации", error_msg)

    def on_conversion_success(self, docx_path):
        """Обработка успешной конвертации."""
        self.progress_bar.stop()
        self.progress_var.set(100)
        self.status_var.set(f"Конвертация завершена: {os.path.basename(docx_path)}")
        self.log(f"Конвертация успешно завершена: {docx_path}")
        
        # Автоматически загружаем полученный DOCX в интерфейс
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, docx_path)
        self.clear_results()
        
        # Запускаем populate_doc_tree в фоновом потоке, чтобы не блокировать GUI
        self.status_var.set("Загрузка структуры документа...")
        self.progress_bar.configure(mode='indeterminate')
        self.progress_bar.start()
        
        def populate_thread():
            """Заполняет дерево документа в фоновом потоке."""
            try:
                # Собираем данные для дерева в фоновом потоке
                tree_data = self._collect_doc_tree_data(docx_path)
                # Передаём данные в главный поток для вставки в Treeview
                self.root.after(0, lambda: self._insert_doc_tree_data(docx_path, tree_data))
            except Exception as e:
                self.root.after(0, lambda: self._on_populate_error(docx_path, e))
        
        thread = threading.Thread(target=populate_thread, daemon=True)
        thread.start()

    def on_conversion_error(self, error):
        """Обработка ошибки конвертации."""
        self.progress_bar.stop()
        self.progress_bar.configure(mode='determinate')
        self.progress_var.set(0)
        self.status_var.set("Ошибка конвертации")
        self.log(f"Ошибка конвертации: {error}")
        messagebox.showerror("Ошибка конвертации", f"Не удалось сконвертировать Markdown файл:\n{error}")

    def browse_config(self):
        filename = filedialog.askopenfilename(filetypes=[("Configuration Files", "*.yaml;*.yml;*.json"), ("YAML Files", "*.yaml;*.yml"), ("JSON Files", "*.json")])
        if filename:
            self.config_path = filename
            self.entry_config.delete(0, tk.END)
            self.entry_config.insert(0, filename)
            try:
                self.controller.load_config(filename)
                self.log("Новая конфигурация загружена.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить конфиг: {e}")

    def open_document_for_view(self):
        """Открыть выбранный документ для просмотра в ассоциированном приложении."""
        if not self.controller.doc_path or not os.path.exists(self.controller.doc_path):
            messagebox.showwarning("Внимание", "Документ не выбран или файл не существует.")
            return
        try:
            os.startfile(self.controller.doc_path)
            self.log(f"Открыт документ для просмотра: {self.controller.doc_path}")
        except Exception as e:
            self.log(f"Ошибка при открытии документа: {e}")
            messagebox.showerror("Ошибка", f"Не удалось открыть документ:\n{str(e)}")

    def open_config_editor_window(self):
        """Открыть визуальный редактор конфигурации."""
        config_path = self.entry_config.get().strip()
        if not config_path or not os.path.exists(config_path):
            messagebox.showwarning("Внимание", "Конфиг не выбран или файл не существует.")
            return
        try:
            open_config_editor(self.root, config_path)
            self.log(f"Открыт редактор конфигурации: {config_path}")
        except Exception as e:
            self.log(f"Ошибка при открытии редактора: {e}")
            messagebox.showerror("Ошибка", f"Не удалось открыть редактор:\n{str(e)}")

    def open_monitor_window(self):
        """Открыть панель мониторинга в реальном времени."""
        try:
            open_monitor(self.root, self.controller)
            self.log("Открыт мониторинг системы")
        except Exception as e:
            self.log(f"Ошибка при открытии мониторинга: {e}")
            messagebox.showerror("Ошибка", f"Не удалось открыть мониторинг:\n{str(e)}")

    def highlight_issues(self):
        """Создать копию документа с подсветкой проблемных элементов и открыть её."""
        doc_path = self.controller.doc_path if self.controller else None
        if not doc_path or not os.path.exists(doc_path):
            messagebox.showwarning("Внимание", "Документ не выбран или файл не существует.")
            return
        
        current_issues = self.controller.current_issues if self.controller else []
        if not current_issues:
            if messagebox.askyesno("Нет данных",
                                   "Проблемы не найдены. Выполнить аудит перед подсветкой?"):
                self.start_audit_thread()
                # После аудита пользователь должен нажать кнопку подсветки снова
                return
            else:
                return
        
        self.status_var.set("Создание документа с подсветкой...")
        self.progress_bar.start()
        self.log(f"Начата подсветка проблемных мест в документе {os.path.basename(doc_path)}")
        
        def highlight_thread():
            try:
                highlighted_path = self.controller.highlight_issues()
                if highlighted_path:
                    self.root.after(0, lambda: self.on_highlight_success(highlighted_path))
                else:
                    raise RuntimeError("Контроллер вернул None")
            except Exception as e:
                self.root.after(0, lambda: self.on_highlight_error(e))
        
        thread = threading.Thread(target=highlight_thread, daemon=True)
        thread.start()

    def on_highlight_success(self, highlighted_path):
        """Обработка успешной подсветки."""
        self.progress_bar.stop()
        self.status_var.set(f"Подсветка завершена: {os.path.basename(highlighted_path)}")
        self.log(f"Документ с подсветкой сохранён: {highlighted_path}")
        
        # Открываем файл для просмотра
        try:
            os.startfile(highlighted_path)
            self.log(f"Открыт документ с подсветкой: {highlighted_path}")
        except Exception as e:
            self.log(f"Ошибка при открытии документа с подсветкой: {e}")
            messagebox.showwarning("Внимание",
                                   f"Файл создан, но не удалось открыть:\n{highlighted_path}\n\nОшибка: {e}")
        
        messagebox.showinfo("Успех",
                            f"Документ с подсветкой создан и открыт.\n{highlighted_path}")

    def on_highlight_error(self, error):
        """Обработка ошибки подсветки."""
        self.progress_bar.stop()
        self.status_var.set("Ошибка подсветки")
        self.log(f"Ошибка подсветки: {error}")
        messagebox.showerror("Ошибка подсветки",
                             f"Не удалось создать документ с подсветкой:\n{error}")
    def find_in_document(self):
        """Создать временный документ с подсветкой только выбранной ошибки и открыть его."""
        selected_ids = self.tree.selection()
        if not selected_ids:
            messagebox.showinfo("Инфо", "Выберите ошибку в таблице для поиска в документе.")
            return
        # Берём первую выбранную ошибку
        item_id = selected_ids[0]
        issue = next((i for i in self.controller.current_issues if i.id == item_id), None)
        if issue is None:
            messagebox.showwarning("Ошибка", "Выбранная проблема не найдена в списке.")
            return
        
        if not self.controller.doc_path or not os.path.exists(self.controller.doc_path):
            messagebox.showwarning("Внимание", "Документ не выбран или файл не существует.")
            return
        
        self.status_var.set("Создание документа с выделением выбранной ошибки...")
        self.progress_bar.start()
        self.log(f"Поиск в документе: {issue.id} - {issue.description[:50]}...")
        
        def find_thread():
            try:
                # Генерируем имя файла для результата
                base_name = os.path.splitext(self.controller.doc_path)[0]
                output_path = f"{base_name}_find_{issue.id}.docx"
                
                # Используем HighlightEngine для подсветки только этой ошибки
                highlighted_path = HighlightEngine.highlight_issues(
                    self.controller.doc_path, [issue], output_path
                )
                
                self.root.after(0, lambda: self.on_find_success(highlighted_path, issue))
            except Exception as e:
                self.root.after(0, lambda: self.on_find_error(e))
        
        thread = threading.Thread(target=find_thread, daemon=True)
        thread.start()
    
    def on_find_success(self, highlighted_path, issue):
        """Обработка успешного создания документа с выделением."""
        self.progress_bar.stop()
        self.status_var.set(f"Документ с выделением создан: {os.path.basename(highlighted_path)}")
        self.log(f"Документ с выделением ошибки сохранён: {highlighted_path}")
        
        # Открываем файл для просмотра
        try:
            os.startfile(highlighted_path)
            self.log(f"Открыт документ с выделением: {highlighted_path}")
        except Exception as e:
            self.log(f"Ошибка при открытии документа с выделением: {e}")
            messagebox.showwarning("Внимание",
                                   f"Файл создан, но не удалось открыть:\n{highlighted_path}\n\nОшибка: {e}")
        
        messagebox.showinfo("Успех",
                            f"Документ с выделением ошибки создан и открыт.\n{highlighted_path}")
    
    def on_find_error(self, error):
        """Обработка ошибки поиска."""
        self.progress_bar.stop()
        self.status_var.set("Ошибка поиска")
        self.log(f"Ошибка поиска в документе: {error}")
        messagebox.showerror("Ошибка поиска",
                             f"Не удалось создать документ с выделением:\n{error}")

    def clear_results(self, clear_doc_tree=True):
        self.tree.delete(*self.tree.get_children())
        # Очищаем список проблем в контроллере (если нужно)
        if self.controller:
            self.controller.current_issues = []
            self.controller.last_fixed_path = None
        self.btn_fix_selected.config(state=tk.DISABLED)
        self.btn_fix_all.config(state=tk.DISABLED)
        self.btn_export_report.config(state=tk.DISABLED)
        self.btn_save.config(state=tk.DISABLED)
        self.btn_find_in_doc.config(state=tk.DISABLED)
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        # Очищаем дерево документа только если указано
        if clear_doc_tree:
            self.doc_tree.delete(*self.doc_tree.get_children())

    def populate_doc_tree(self, doc_path):
        """Заполняет дерево документа (синхронная версия для обратной совместимости).
        
        Для больших файлов использует асинхронное заполнение через _collect_doc_tree_data + _insert_doc_tree_data.
        Для маленьких файлов заполняет синхронно.
        """
        # Очищаем предыдущее содержимое
        self.doc_tree.delete(*self.doc_tree.get_children())
        
        if not os.path.exists(doc_path):
            self.log(f"Файл {doc_path} не существует, дерево не может быть заполнено.")
            return
        
        # Для маленьких файлов используем синхронное заполнение
        # Для больших - асинхронное
        try:
            file_size = os.path.getsize(doc_path)
        except:
            file_size = 0
        
        if file_size < 1024 * 100:  # меньше 100 КБ — синхронно
            self._populate_doc_tree_sync(doc_path)
        else:
            self.status_var.set("Загрузка структуры документа...")
            self.progress_bar.configure(mode='indeterminate')
            self.progress_bar.start()
            
            def populate_thread():
                try:
                    tree_data = self._collect_doc_tree_data(doc_path)
                    self.root.after(0, lambda: self._insert_doc_tree_data(doc_path, tree_data))
                except Exception as e:
                    self.root.after(0, lambda: self._on_populate_error(doc_path, e))
            
            thread = threading.Thread(target=populate_thread, daemon=True)
            thread.start()
    
    def _populate_doc_tree_sync(self, doc_path):
        """Синхронное заполнение дерева документа (для маленьких файлов)."""
        # Используем DocumentModel из контроллера, если он соответствует пути, иначе создаём временный
        doc_model = None
        if self.controller.document_model and self.controller.doc_path == doc_path:
            doc_model = self.controller.document_model
        else:
            try:
                doc_model = DocumentModel(doc_path)
            except Exception as e:
                self.log(f"Ошибка при создании модели документа {doc_path}: {e}")
                root = self.doc_tree.insert("", tk.END, text="Документ (ошибка)", open=True)
                self.doc_tree.insert(root, tk.END, text="Не удалось загрузить структуру", values=("Error", ""))
                return
        
        try:
            doc = doc_model.docx_document
        except Exception as e:
            self.log(f"Ошибка при открытии документа {doc_path}: {e}")
            root = self.doc_tree.insert("", tk.END, text="Документ (ошибка)", open=True)
            self.doc_tree.insert(root, tk.END, text="Не удалось загрузить структуру", values=("Error", ""))
            return
        
        root = self.doc_tree.insert("", tk.END, text=os.path.basename(doc_path), open=True)
        
        heading_count = 0
        table_count = 0
        list_count = 0
        para_count = 0
        section_count = 0
        
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            
            style_name = para.style.name
            style_lower = style_name.lower()
            if style_lower.startswith('heading') or 'заголовок' in style_lower:
                elem_type = "Heading"
                heading_count += 1
                display_text = f"Заголовок {heading_count}: {text[:50]}"
                values = (elem_type, f"стр. ~{i//20+1}")
                self.doc_tree.insert(root, tk.END, text=display_text, values=values)
            elif para._element.pPr is not None and para._element.pPr.numPr is not None:
                elem_type = "List"
                list_count += 1
                display_text = f"Список {list_count}: {text[:50]}"
                values = (elem_type, f"стр. ~{i//20+1}")
                self.doc_tree.insert(root, tk.END, text=display_text, values=values)
            else:
                elem_type = "Normal"
                para_count += 1
                display_text = f"Параграф {para_count}: {text[:50]}"
                values = (elem_type, f"стр. ~{i//20+1}")
                self.doc_tree.insert(root, tk.END, text=display_text, values=values)
        
        for i, table in enumerate(doc.tables):
            table_count += 1
            first_cell_text = ""
            try:
                if table.rows and table.columns:
                    first_cell_text = table.cell(0, 0).text.strip()[:30]
            except:
                pass
            display_text = f"Таблица {table_count}: {first_cell_text}"
            values = ("Table", f"стр. ~{table_count*2}")
            self.doc_tree.insert(root, tk.END, text=display_text, values=values)
        
        for i, section in enumerate(doc.sections):
            section_count += 1
            header = section.header
            footer = section.footer
            has_header = header is not None and any(p.text.strip() for p in header.paragraphs)
            has_footer = footer is not None and any(p.text.strip() for p in footer.paragraphs)
            if has_header or has_footer:
                display_text = f"Секция {section_count}"
                if has_header and has_footer:
                    display_text += " (колонтитулы)"
                elif has_header:
                    display_text += " (верхний колонтитул)"
                elif has_footer:
                    display_text += " (нижний колонтитул)"
                values = ("Section", f"стр. ~{section_count*10}")
                self.doc_tree.insert(root, tk.END, text=display_text, values=values)
        
        total_elements = heading_count + table_count + list_count + para_count + section_count
        if total_elements == 0:
            self.doc_tree.insert(root, tk.END, text="Документ пуст или не содержит распознаваемых элементов", values=("Info", ""))
        
        self.log(f"Дерево документа заполнено: {heading_count} заголовков, {table_count} таблиц, {list_count} списков, {para_count} параграфов, {section_count} секций с колонтитулами.")
    
    def _collect_doc_tree_data(self, doc_path):
        """Собирает данные для дерева документа в фоновом потоке.
        Возвращает словарь с данными для вставки в Treeview.
        """
        import docx
        
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"Файл {doc_path} не найден")
        
        doc = docx.Document(doc_path)
        
        data = {
            'filename': os.path.basename(doc_path),
            'items': [],
            'counts': {
                'heading': 0,
                'table': 0,
                'list': 0,
                'para': 0,
                'section': 0,
            }
        }
        
        heading_count = 0
        table_count = 0
        list_count = 0
        para_count = 0
        
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            
            style_name = para.style.name
            style_lower = style_name.lower()
            
            if style_lower.startswith('heading') or 'заголовок' in style_lower:
                heading_count += 1
                data['items'].append({
                    'type': 'Heading',
                    'display': f"Заголовок {heading_count}: {text[:50]}",
                    'values': ("Heading", f"стр. ~{i//20+1}")
                })
            elif para._element.pPr is not None and para._element.pPr.numPr is not None:
                list_count += 1
                data['items'].append({
                    'type': 'List',
                    'display': f"Список {list_count}: {text[:50]}",
                    'values': ("List", f"стр. ~{i//20+1}")
                })
            else:
                para_count += 1
                data['items'].append({
                    'type': 'Normal',
                    'display': f"Параграф {para_count}: {text[:50]}",
                    'values': ("Normal", f"стр. ~{i//20+1}")
                })
        
        data['counts']['heading'] = heading_count
        data['counts']['table'] = table_count
        data['counts']['list'] = list_count
        data['counts']['para'] = para_count
        
        for i, table in enumerate(doc.tables):
            table_count += 1
            first_cell_text = ""
            try:
                if table.rows and table.columns:
                    first_cell_text = table.cell(0, 0).text.strip()[:30]
            except:
                pass
            data['items'].append({
                'type': 'Table',
                'display': f"Таблица {table_count}: {first_cell_text}",
                'values': ("Table", f"стр. ~{table_count*2}")
            })
        
        data['counts']['table'] = table_count
        
        section_count = 0
        for i, section in enumerate(doc.sections):
            section_count += 1
            header = section.header
            footer = section.footer
            has_header = header is not None and any(p.text.strip() for p in header.paragraphs)
            has_footer = footer is not None and any(p.text.strip() for p in footer.paragraphs)
            if has_header or has_footer:
                display_text = f"Секция {section_count}"
                if has_header and has_footer:
                    display_text += " (колонтитулы)"
                elif has_header:
                    display_text += " (верхний колонтитул)"
                elif has_footer:
                    display_text += " (нижний колонтитул)"
                data['items'].append({
                    'type': 'Section',
                    'display': display_text,
                    'values': ("Section", f"стр. ~{section_count*10}")
                })
        
        data['counts']['section'] = section_count
        
        return data
    
    def _insert_doc_tree_data(self, doc_path, tree_data):
        """Вставляет собранные данные в дерево документа (в главном потоке)."""
        self.progress_bar.stop()
        self.progress_bar.configure(mode='determinate')
        
        self.doc_tree.delete(*self.doc_tree.get_children())
        
        root = self.doc_tree.insert("", tk.END, text=tree_data['filename'], open=True)
        
        for item in tree_data['items']:
            self.doc_tree.insert(root, tk.END, text=item['display'], values=item['values'])
        
        counts = tree_data['counts']
        total = sum(counts.values())
        if total == 0:
            self.doc_tree.insert(root, tk.END, text="Документ пуст или не содержит распознаваемых элементов", values=("Info", ""))
        
        self.log(f"Дерево документа заполнено: {counts['heading']} заголовков, {counts['table']} таблиц, {counts['list']} списков, {counts['para']} параграфов, {counts['section']} секций с колонтитулами.")
        self.status_var.set(f"Загружен: {os.path.basename(doc_path)}")
        
        # Показываем сообщение об успехе (уже в главном потоке)
        messagebox.showinfo("Успех", f"Markdown файл успешно сконвертирован и загружен в интерфейс.\n{doc_path}")
    
    def _on_populate_error(self, doc_path, error):
        """Обработка ошибки заполнения дерева документа."""
        self.progress_bar.stop()
        self.progress_bar.configure(mode='determinate')
        self.log(f"Ошибка при загрузке структуры документа {doc_path}: {error}")
        root = self.doc_tree.insert("", tk.END, text="Документ (ошибка)", open=True)
        self.doc_tree.insert(root, tk.END, text=f"Не удалось загрузить структуру: {error}", values=("Error", ""))
        self.status_var.set("Ошибка загрузки структуры документа")

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}"
        print(full_msg)
        # Обновление UI должно происходить в главном потоке
        def update_log():
            if hasattr(self, 'log_text') and self.log_text.winfo_exists():
                self.log_text.config(state='normal')
                self.log_text.insert(tk.END, full_msg + "\n")
                self.log_text.see(tk.END)
                self.log_text.config(state='disabled')
        # Если мы в главном потоке, вызываем напрямую, иначе через after
        if self.root and hasattr(self.root, 'after'):
            self.root.after(0, update_log)
        else:
            update_log()

    def update_audit_progress(self, stage, progress, total, message):
        """
        Колбэк для обновления прогресса аудита в реальном времени.
        
        Args:
            stage (str): идентификатор этапа ('page_setup', 'paragraphs', 'tables', 'headers_footers', 'images')
            progress (int): текущий прогресс внутри этапа (например, обработано параграфов)
            total (int): общее количество элементов в этапе (например, всего параграфов)
            message (str): текстовое описание текущего действия для лога и статуса
        """
        # Вычисляем прогресс внутри этапа (от 0 до 1)
        if total > 0:
            stage_ratio = progress / total
        else:
            stage_ratio = 1.0
        
        # Вычисляем вклад этого этапа в общий прогресс
        stage_weight = self.audit_stage_weights.get(stage, 0)
        stage_contribution = stage_weight * stage_ratio
        
        # Общий прогресс = завершённые этапы + вклад текущего этапа
        overall = self.audit_stage_completed + stage_contribution
        
        # Ограничиваем 0-100
        overall = max(0, min(100, overall))
        
        def update_ui():
            # Обновляем прогресс-бар
            self.progress_var.set(overall)
            
            # Обновляем статусную строку
            self.status_var.set(f"Аудит: {stage} ({progress}/{total}) - {message}")
            
            # Логируем сообщение
            self.log(f"Прогресс аудита: {stage} {progress}/{total} - {message}")
            
            # Если этап завершён (progress == total), обновляем завершённые этапы
            if progress == total and stage in self.audit_stage_weights:
                # Добавляем вес этапа к завершённым (но не превышаем 100)
                self.audit_stage_completed += stage_weight
                # Сбрасываем текущий этап
                self.current_stage = ''
                self.log(f"Этап '{stage}' завершён.")
            else:
                self.current_stage = stage
            
            # Принудительное обновление GUI
            self.root.update_idletasks()
        
        # Вызываем обновление UI в главном потоке
        if self.root and hasattr(self.root, 'after'):
            self.root.after(0, update_ui)
        else:
            update_ui()

    # --- Логика Аудита (в потоке) ---

    def start_audit_thread(self):
        if not self.controller.doc_path:
            messagebox.showwarning("Внимание", "Выберите документ DOCX!")
            return
        if not self.controller.config_loader:
            messagebox.showwarning("Внимание", "Конфигурация не загружена!")
            return

        self.clear_results(clear_doc_tree=False)
        self.is_scanning = True
        self.btn_audit.config(state=tk.DISABLED)
        # Сброс прогресса
        self.audit_stage_completed = 0
        self.current_stage = ''
        self.progress_var.set(0)
        self.progress_bar.start()
        self.status_var.set("Идет сканирование документа...")
        
        thread = threading.Thread(target=self.run_audit)
        thread.daemon = True
        thread.start()

    def run_audit(self):
        try:
            self.log("Запуск двигателя аудита...")
            # Используем контроллер для запуска аудита асинхронно
            def on_complete(issues):
                self.root.after(0, lambda: self.on_audit_complete(issues))
            def on_error(e):
                self.root.after(0, lambda: self.on_audit_error(e))
            self.controller.start_audit_async(on_complete, on_error)
        except Exception as e:
            logger.error(f"Ошибка аудита: {e}", exc_info=True)
            self.root.after(0, lambda: messagebox.showerror("Ошибка аудита", str(e)))
            self.root.after(0, self.reset_ui_state)
            self.root.after(0, self.finish_audit)

    def on_audit_complete(self, issues):
        """Обработка завершения аудита."""
        # issues уже сохранены в контроллере, используем их
        self.root.after(0, lambda: self.populate_tree(issues))
        self.root.after(0, self.finish_audit)

    def on_audit_error(self, error):
        """Обработка ошибки аудита."""
        logger.error(f"Ошибка аудита: {error}", exc_info=True)
        self.root.after(0, lambda: messagebox.showerror("Ошибка аудита", str(error)))
        self.root.after(0, self.reset_ui_state)
        self.root.after(0, self.finish_audit)

    def populate_tree(self, issues, group_by=None):
        self.tree.delete(*self.tree.get_children())
        
        # Просто отображаем плоский список (группировка удалена)
        for issue in issues:
            self._insert_issue_row(issue)
        
        # Обновляем сводку
        if issues and hasattr(self.controller, 'audit_engine') and self.controller.audit_engine:
            summary = self.controller.audit_engine.get_summary()
            self.log(f"Аудит завершен. Найдено: {len(issues)} проблем.")
            self.log(f"   Критических: {summary['CRITICAL']}, Предупреждений: {summary['WARNING']}")
            self.status_var.set(f"Аудит завершен. Найдено проблем: {len(issues)}")
        elif issues:
            self.status_var.set(f"Найдено проблем: {len(issues)}")

    def _insert_issue_row(self, issue):
        """Вставляет строку проблемы в дерево с цветовой кодировкой."""
        # Маппинг важности на цвета/текст
        sev_map = {
            Severity.CRITICAL: ("❌ CRITICAL", "critical"),
            Severity.WARNING: ("⚠️ WARNING", "warning"),
            Severity.INFO: ("ℹ️ INFO", "info")
        }
        sev_text, tag = sev_map.get(issue.severity, ("?", "info"))
        
        loc_str = f"стр. ~{issue.location.get('page', '?')}"
        
        self.tree.insert("", tk.END, iid=issue.id, values=(
            issue.id.split('_')[-1], # Короткий ID
            sev_text,
            issue.category,
            issue.element_type,
            loc_str,
            issue.description,
            issue.current_value,
            issue.expected_value
        ), tags=(tag,))

    def treeview_sort_column(self, col, reverse=None):
        """Сортирует дерево по колонке col (прямая или обратная).
        Если reverse не указан, определяется автоматически:
        - если колонка уже отсортирована, инвертировать направление,
        - иначе сортировать по возрастанию.
        """
        # Определяем направление сортировки
        if reverse is None:
            if self.sort_column == col:
                # Инвертируем текущее направление
                reverse = not self.sort_reverse
            else:
                reverse = False
        else:
            # Используем переданное значение (для обратной совместимости)
            pass
        
        # Получаем все элементы
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        
        # Определяем тип данных для сортировки
        # Если колонка "id" (числовой индекс) или "location" (строка с числами) или "current"/"expected" (числа)
        # Для простоты сортируем как строки, но для числовых колонок можно преобразовать
        if col in ("id", "location"):
            # Пробуем извлечь числа из строки
            def try_int(val):
                try:
                    # Ищем первое число в строке
                    import re
                    match = re.search(r'\d+', val)
                    if match:
                        return int(match.group())
                except:
                    pass
                return 0
            items.sort(key=lambda t: try_int(t[0]), reverse=reverse)
        else:
            # Сортировка как строки (регистронезависимая)
            items.sort(key=lambda t: t[0].lower(), reverse=reverse)
        
        # Переставляем элементы в дереве
        for index, (_, k) in enumerate(items):
            self.tree.move(k, '', index)
        
        # Запоминаем состояние сортировки
        self.sort_column = col
        self.sort_reverse = reverse
        
        # Обновляем заголовок для следующего клика (инвертируем направление)
        self.tree.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))

    def finish_audit(self):
        self.is_scanning = False
        self.progress_bar.stop()
        self.btn_audit.config(state=tk.NORMAL)
        self.reset_ui_state()
        
        if len(self.controller.current_issues) > 0:
            self.btn_fix_all.config(state=tk.NORMAL)
            self.btn_fix_selected.config(state=tk.NORMAL)
            self.btn_fix_styles_only.config(state=tk.NORMAL)
            self.btn_export_report.config(state=tk.NORMAL)
            self.btn_find_in_doc.config(state=tk.NORMAL)  # Включаем кнопку поиска
            self.btn_save.config(state=tk.NORMAL) # Разрешаем сохранение даже если просто открыли (хотя менять нечего)
            # Логичнее разрешать сохранение только после изменений, но пока оставим так
            self.btn_save.config(state=tk.DISABLED) # Блокируем, пока нет изменений
        else:
            self.log("Документ соответствует конфигурации!")
            self.status_var.set("Документ соответствует ГОСТ (ошибок не найдено).")

    def reset_ui_state(self):
        pass # Заглушка для будущей логики

    # --- Логика Применения (Apply) ---

    def get_selected_issues(self):
        selected_ids = self.tree.selection()
        # Возвращаем объекты Issue из списка current_issues контроллера, чьи ID есть в selection
        # Создаем словарь для быстрого поиска
        issues_map = {i.id: i for i in self.controller.current_issues}
        return [issues_map[iid] for iid in selected_ids if iid in issues_map]

    def fix_selected(self):
        selected = self.get_selected_issues()
        if not selected:
            messagebox.showinfo("Инфо", "Выберите ошибки в списке для исправления.")
            return
        if not messagebox.askyesno("Подтверждение", f"Исправить выбранные {len(selected)} ошибок?"):
            return
        self.apply_fixes(selected)

    def fix_styles_only(self):
        """Исправить только стили для выбранных проблем."""
        selected = self.get_selected_issues()
        if not selected:
            messagebox.showinfo("Инфо", "Выберите ошибки в списке для исправления.")
            return
        if not messagebox.askyesno("Подтверждение",
                                   f"Применить только стили к выбранным {len(selected)} ошибкам?\nПрямое форматирование останется неизменным."):
            return
        
        self.status_var.set("Применение только стилей...")
        self.progress_bar.start()
        self.root.update_idletasks()
        
        try:
            issue_ids = [issue.id for issue in selected]
            stats = self.controller.apply_styles_only(issue_ids)
            
            applied = stats['applied']
            failed = stats['failed']
            output_path = self.controller.last_fixed_path
            
            msg = f"Готово!\nУспешно применено стилей: {applied}\nОшибок: {failed}\nФайл сохранен: {output_path}"
            
            if failed > 0:
                messagebox.showwarning("Результат", msg)
            else:
                messagebox.showinfo("Результат", msg)
            
            # Удаляем исправленные строки из таблицы
            for issue in selected:
                item_id = issue.id
                if self.tree.exists(item_id):
                    self.tree.delete(item_id)
            
            self.log(f"Успешно применено стилей: {applied}")
            self.btn_save.config(state=tk.NORMAL)
            
            if len(self.controller.current_issues) == 0:
                self.status_var.set("Все исправления применены. Документ готов к сохранению.")
                self.btn_fix_all.config(state=tk.DISABLED)
                self.btn_fix_selected.config(state=tk.DISABLED)
                self.btn_fix_styles_only.config(state=tk.DISABLED)
                self.btn_export_report.config(state=tk.DISABLED)
                self.btn_find_in_doc.config(state=tk.DISABLED)
            else:
                self.status_var.set(f"Стили применены. Осталось проблем: {len(self.controller.current_issues)}")
            self.update_summary_stats()
            
        except Exception as e:
            logger.error(f"Ошибка при применении стилей: {e}", exc_info=True)
            messagebox.showerror("Критическая ошибка", f"Не удалось применить стили:\n{str(e)}")
            self.status_var.set("Ошибка при применении стилей")
        finally:
            self.progress_bar.stop()

    def fix_all(self):
        if not self.controller.current_issues:
            return
        if messagebox.askyesno("Подтверждение", f"Исправить все {len(self.controller.current_issues)} найденных ошибок?"):
            self.apply_fixes(self.controller.current_issues)

    def apply_fixes(self, issues_to_fix):
        if not issues_to_fix:
            return
            
        self.status_var.set("Применение исправлений...")
        self.progress_bar.start()
        self.root.update_idletasks()

        try:
            # Используем контроллер для применения исправлений
            issue_ids = [issue.id for issue in issues_to_fix]
            stats = self.controller.apply_fixes(issue_ids)
            
            applied = stats['applied']
            failed = stats['failed']
            output_path = self.controller.last_fixed_path
            
            msg = f"Готово!\nУспешно применено: {applied}\nОшибок: {failed}\nФайл сохранен: {output_path}"
            
            if failed > 0:
                messagebox.showwarning("Результат", msg)
            else:
                messagebox.showinfo("Результат", msg)
            
            # Обновляем интерфейс: удаляем исправленные строки из таблицы
            for issue in issues_to_fix:
                item_id = issue.id
                if self.tree.exists(item_id):
                    self.tree.delete(item_id)
            
            # Переключаемся на исправленный документ
            import os
            if output_path and os.path.exists(output_path):
                self.controller.set_document(output_path)
                self.entry_file.delete(0, tk.END)
                self.entry_file.insert(0, output_path)
                self.log(f"Переключено на исправленный документ: {output_path}")
                # Очищаем дерево проблем, так как документ изменился
                self.tree.delete(*self.tree.get_children())
            
            self.log(f"Успешно применено исправлений: {applied}")
            self.btn_save.config(state=tk.NORMAL) # Теперь можно сохранять
            
            if len(self.controller.current_issues) == 0:
                self.status_var.set("Все исправления применены. Документ готов к сохранению.")
                self.btn_fix_all.config(state=tk.DISABLED)
                self.btn_fix_selected.config(state=tk.DISABLED)
                self.btn_fix_styles_only.config(state=tk.DISABLED)
                self.btn_export_report.config(state=tk.DISABLED)
                self.btn_find_in_doc.config(state=tk.DISABLED)  # Отключаем кнопку поиска
            else:
                self.status_var.set(f"Исправления применены. Осталось проблем: {len(self.controller.current_issues)}")
            self.update_summary_stats()

        except Exception as e:
            logger.error(f"Ошибка при применении исправлений: {e}", exc_info=True)
            messagebox.showerror("Критическая ошибка", f"Не удалось применить исправления:\n{str(e)}")
            self.status_var.set("Ошибка при применении")
        finally:
            self.progress_bar.stop()

    def update_summary_stats(self):
        """Обновляет метки статистики после изменений."""
        if self.controller.report_manager:
            summary = self.controller.report_manager.get_summary()
            total = sum(summary.values())
            crit = summary.get('CRITICAL', 0)
            warn = summary.get('WARNING', 0)
            info = summary.get('INFO', 0)
        elif hasattr(self.controller, 'current_issues'):
            # Подсчет оставшихся проблем по важности
            crit = sum(1 for i in self.controller.current_issues if i.severity.value == 'CRITICAL')
            warn = sum(1 for i in self.controller.current_issues if i.severity.value == 'WARNING')
            info = sum(1 for i in self.controller.current_issues if i.severity.value == 'INFO')
            total = len(self.controller.current_issues)
        else:
            total = crit = warn = info = 0
        
        # Вывод в лог (можно заменить на обновление GUI-меток, если они будут добавлены)
        self.log(f"Осталось проблем: Критических={crit}, Предупреждений={warn}, Инфо={info}")
        # Обновление статусной строки
        self.status_var.set(f"Осталось проблем: {total} (Критических: {crit}, Предупреждений: {warn})")

    def deselect_all(self):
        def deselect():
            self.tree.selection_remove(self.tree.selection())
            self.select_all_var.set(False)
        self.root.after(0, deselect)

    def toggle_select_all(self):
        """Выделяет или снимает выделение со всех строк в таблице."""
        if self.select_all_var.get():
            # Выделить все
            def select():
                self.tree.selection_set(self.tree.get_children())
            self.root.after(0, select)
        else:
            # Снять выделение
            def deselect():
                self.tree.selection_remove(self.tree.selection())
            self.root.after(0, deselect)

    def apply_filter(self, event=None):
        filter_val = self.filter_var.get()
        self.tree.delete(*self.tree.get_children())
        
        if filter_val == "ALL":
            target_sev = None
        else:
            target_sev = Severity[filter_val]
        
        if self.controller.report_manager:
            filtered_issues = self.controller.report_manager.apply_filter(severity=target_sev)
        else:
            # fallback на старую логику
            if filter_val == "ALL":
                filtered_issues = self.controller.current_issues
            else:
                filtered_issues = [i for i in self.controller.current_issues if i.severity == target_sev]
            
        self.populate_tree(filtered_issues)

    # Метод apply_grouping удалён

    def on_tree_select(self, event):
        """Обновляет панель деталей и предпросмотра при выборе строки в таблице."""
        selected_ids = self.tree.selection()
        if not selected_ids:
            return
        # Берем первую выбранную строку
        item_id = selected_ids[0]
        # Находим соответствующую проблему
        issue = next((i for i in self.controller.current_issues if i.id == item_id), None)
        if issue is None:
            return
        
        # Формируем текст деталей
        details = f"""
ID: {issue.id}
Важность: {issue.severity.value}
Категория: {issue.category}
Стиль: {issue.element_type}
Локация: стр. ~{issue.location.get('page', '?')}
Описание: {issue.description}
Текущее значение: {issue.current_value}
Ожидаемое значение: {issue.expected_value}
Автоисправление: {'Да' if issue.auto_fixable else 'Нет'}
Рекомендация: {issue.recommendation if issue.recommendation else '—'}
"""
        self.details_text.config(state='normal')
        self.details_text.delete(1.0, tk.END)
        self.details_text.insert(tk.END, details.strip())
        self.details_text.config(state='disabled')
        
        # Предпросмотр (извлечение реального текста параграфа)
        para_index = issue.location.get('index')
        preview = ""
        if para_index is not None and self.controller.document_model:
            try:
                doc = self.controller.document_model.docx_document
                if 0 <= para_index < len(doc.paragraphs):
                    para = doc.paragraphs[para_index]
                    preview = para.text.strip()
                    if not preview:
                        preview = "(пустой параграф)"
                    # Добавим контекст: предыдущий и следующий параграфы
                    context = []
                    if para_index > 0:
                        prev = doc.paragraphs[para_index - 1].text.strip()
                        if prev:
                            context.append(f"← {prev[:100]}...")
                    if para_index < len(doc.paragraphs) - 1:
                        nxt = doc.paragraphs[para_index + 1].text.strip()
                        if nxt:
                            context.append(f"→ {nxt[:100]}...")
                    if context:
                        preview += "\n\nКонтекст:\n" + "\n".join(context)
                else:
                    preview = f"Параграф #{para_index} не найден в документе."
            except Exception as e:
                preview = f"Ошибка при извлечении фрагмента: {e}"
        else:
            preview = f"Фрагмент текста с ошибкой не доступен.\nПараграф #{para_index if para_index is not None else '?'}"
        self.preview_text.config(state='normal')
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(tk.END, preview)
        self.preview_text.config(state='disabled')
        
        # Активируем кнопку "Найти в документе"
        self.btn_find_in_doc.config(state=tk.NORMAL)

    def exclude_selected(self):
        """Исключает выбранные проблемы из отчёта (удаляет из таблицы и списка)."""
        selected_ids = self.tree.selection()
        if not selected_ids:
            messagebox.showinfo("Инфо", "Выберите ошибки для исключения из отчёта.")
            return
        if not messagebox.askyesno("Подтверждение", f"Исключить {len(selected_ids)} выбранных ошибок из отчёта?"):
            return
        
        # Удаляем из таблицы
        for item_id in selected_ids:
            if self.tree.exists(item_id):
                self.tree.delete(item_id)
        
        # Удаляем из списка текущих проблем в контроллере
        self.controller.current_issues = [i for i in self.controller.current_issues if i.id not in selected_ids]
        
        # Исключаем из менеджера отчётов
        if self.controller.report_manager:
            self.controller.report_manager.exclude_issues(selected_ids)
        
        self.log(f"Исключено {len(selected_ids)} ошибок из отчёта.")
        self.update_summary_stats()
        
        # Если проблем не осталось, отключаем кнопки
        if len(self.controller.current_issues) == 0:
            self.btn_fix_all.config(state=tk.DISABLED)
            self.btn_fix_selected.config(state=tk.DISABLED)
            self.btn_export_report.config(state=tk.DISABLED)
            self.btn_save.config(state=tk.DISABLED)
            self.btn_find_in_doc.config(state=tk.DISABLED)
            self.status_var.set("Все проблемы исключены. Документ готов к сохранению.")

    def save_document(self):
        if not self.controller.doc_path:
            return
        
        # Определяем источник для сохранения
        source_path = None
        if self.controller.last_fixed_path and os.path.exists(self.controller.last_fixed_path):
            source_path = self.controller.last_fixed_path
            source_desc = "исправленный документ"
        else:
            source_path = self.controller.doc_path
            source_desc = "исходный документ (исправления не применялись)"
            if self.controller.current_issues:
                # Есть неисправленные проблемы, предупреждаем
                if not messagebox.askyesno(
                    "Внимание",
                    f"Исправления ещё не применены к {len(self.controller.current_issues)} проблемам. "
                    "Сохранить исходный документ без исправлений?"
                ):
                    return  # пользователь отказался
        
        # Создаём резервную копию исходного документа перед сохранением
        backup_path = None
        if messagebox.askyesno("Резервная копия", "Создать резервную копию исходного документа перед сохранением?"):
            backup_path = self.controller.create_backup(self.controller.doc_path)
            if backup_path:
                self.log(f"Создана резервная копия: {backup_path}")
            else:
                self.log("Не удалось создать резервную копию.")
        
        default_name = os.path.splitext(self.controller.doc_path)[0] + "_fixed.docx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word Documents", "*.docx")]
        )
        
        if file_path:
            try:
                logger.info(f"Сохранение {source_desc} в {file_path}...")
                shutil.copy(source_path, file_path)
                msg = f"Документ сохранен:\n{file_path}"
                if backup_path:
                    msg += f"\nРезервная копия создана:\n{backup_path}"
                messagebox.showinfo("Успех", msg)
                self.log(f"Файл сохранен: {file_path}")
            except Exception as e:
                logger.error(f"Ошибка сохранения: {e}", exc_info=True)
                messagebox.showerror("Ошибка сохранения", str(e))

    def export_report(self):
        """Экспортирует отчёт аудита в текстовый, HTML или CSV файл."""
        if not self.controller.current_issues:
            messagebox.showinfo("Инфо", "Нет данных для экспорта. Сначала выполните аудит.")
            return

        # Диалог выбора формата
        format_choice = simpledialog.askstring("Формат экспорта",
                                                      "Введите формат (txt, html или csv):",
                                                      initialvalue="txt")
        if not format_choice:
            return
        format_choice = format_choice.strip().lower()
        if format_choice not in ("txt", "html", "csv"):
            messagebox.showerror("Ошибка", "Поддерживаются только txt, html и csv.")
            return

        # Диалог выбора файла для сохранения
        default_name = os.path.splitext(self.controller.doc_path)[0] + "_report." + format_choice
        file_path = filedialog.asksaveasfilename(
            defaultextension="." + format_choice,
            initialfile=default_name,
            filetypes=[("Text files", "*.txt"), ("HTML files", "*.html"), ("CSV files", "*.csv")]
        )
        if not file_path:
            return

        try:
            if self.controller.report_manager:
                if format_choice == "txt":
                    self.controller.report_manager.export_txt(file_path)
                elif format_choice == "html":
                    self.controller.report_manager.export_html(file_path)
                else:  # csv
                    self.controller.report_manager.export_csv(file_path)
            else:
                # fallback на старые методы (только txt и html)
                if format_choice == "txt":
                    self._export_report_txt(file_path)
                elif format_choice == "html":
                    self._export_report_html(file_path)
                else:
                    messagebox.showerror("Ошибка", "CSV экспорт недоступен без менеджера отчётов.")
                    return
            messagebox.showinfo("Успех", f"Отчёт сохранён в {file_path}")
            self.log(f"Экспорт отчёта: {file_path}")
        except Exception as e:
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
            messagebox.showerror("Ошибка экспорта", str(e))

    def open_settings(self):
        """Открывает окно настроек допусков с взаимозависимыми полями мм и pt."""
        # Коэффициенты преобразования (импортированы из docx_utils)
        # MM_TO_TWIPS и PT_TO_TWIPS уже доступны

        # Создаем дочернее окно
        settings_win = tk.Toplevel(self.root)
        settings_win.title("Настройки допусков")
        settings_win.geometry("500x400")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()

        # Преобразуем текущие значения из twips в миллиметры и пункты
        font_tol_mm = self.font_tolerance / MM_TO_TWIPS
        font_tol_pt = self.font_tolerance / PT_TO_TWIPS
        indent_tol_mm = self.indent_tolerance / MM_TO_TWIPS
        indent_tol_pt = self.indent_tolerance / PT_TO_TWIPS
        spacing_tol_mm = self.spacing_tolerance / MM_TO_TWIPS
        spacing_tol_pt = self.spacing_tolerance / PT_TO_TWIPS

        # Фрейм для полей ввода
        frame = ttk.Frame(settings_win, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Метка заголовка
        ttk.Label(frame, text="Настройка допусков (мм и pt)", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=4, pady=(0, 15))

        # Заголовки колонок
        ttk.Label(frame, text="Допуск", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        ttk.Label(frame, text="мм", font=("Arial", 10, "bold")).grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(frame, text="pt", font=("Arial", 10, "bold")).grid(row=1, column=2, sticky=tk.W, pady=5, padx=5)
        ttk.Label(frame, text="twips", font=("Arial", 10, "bold")).grid(row=1, column=3, sticky=tk.W, pady=5, padx=5)

        # Допуск размера шрифта
        ttk.Label(frame, text="Размер шрифта:").grid(row=2, column=0, sticky=tk.W, pady=5)
        font_var_mm = tk.StringVar(value=f"{font_tol_mm:.2f}")
        font_var_pt = tk.StringVar(value=f"{font_tol_pt:.2f}")
        font_var_twips = tk.StringVar(value=str(self.font_tolerance))
        font_entry_mm = ttk.Entry(frame, textvariable=font_var_mm, width=10)
        font_entry_mm.grid(row=2, column=1, sticky=tk.W, padx=5)
        font_entry_pt = ttk.Entry(frame, textvariable=font_var_pt, width=10)
        font_entry_pt.grid(row=2, column=2, sticky=tk.W, padx=5)
        ttk.Label(frame, textvariable=font_var_twips).grid(row=2, column=3, sticky=tk.W, padx=5)

        # Допуск отступа
        ttk.Label(frame, text="Отступ:").grid(row=3, column=0, sticky=tk.W, pady=5)
        indent_var_mm = tk.StringVar(value=f"{indent_tol_mm:.2f}")
        indent_var_pt = tk.StringVar(value=f"{indent_tol_pt:.2f}")
        indent_var_twips = tk.StringVar(value=str(self.indent_tolerance))
        indent_entry_mm = ttk.Entry(frame, textvariable=indent_var_mm, width=10)
        indent_entry_mm.grid(row=3, column=1, sticky=tk.W, padx=5)
        indent_entry_pt = ttk.Entry(frame, textvariable=indent_var_pt, width=10)
        indent_entry_pt.grid(row=3, column=2, sticky=tk.W, padx=5)
        ttk.Label(frame, textvariable=indent_var_twips).grid(row=3, column=3, sticky=tk.W, padx=5)

        # Допуск интервалов
        ttk.Label(frame, text="Интервалы:").grid(row=4, column=0, sticky=tk.W, pady=5)
        spacing_var_mm = tk.StringVar(value=f"{spacing_tol_mm:.2f}")
        spacing_var_pt = tk.StringVar(value=f"{spacing_tol_pt:.2f}")
        spacing_var_twips = tk.StringVar(value=str(self.spacing_tolerance))
        spacing_entry_mm = ttk.Entry(frame, textvariable=spacing_var_mm, width=10)
        spacing_entry_mm.grid(row=4, column=1, sticky=tk.W, padx=5)
        spacing_entry_pt = ttk.Entry(frame, textvariable=spacing_var_pt, width=10)
        spacing_entry_pt.grid(row=4, column=2, sticky=tk.W, padx=5)
        ttk.Label(frame, textvariable=spacing_var_twips).grid(row=4, column=3, sticky=tk.W, padx=5)

        # Пояснение
        ttk.Label(frame, text="1 мм ≈ 56.7 twips, 1 pt = 20 twips", font=("Arial", 9), foreground="gray").grid(row=5, column=0, columnspan=4, pady=10)

        # Функции синхронизации
        def update_from_mm(var_mm, var_pt, var_twips):
            """Обновляет pt и twips при изменении мм."""
            try:
                mm = float(var_mm.get().replace(',', '.'))
                twips = int(round(mm * MM_TO_TWIPS))
                var_twips.set(str(twips))
                pt = twips / PT_TO_TWIPS
                var_pt.set(f"{pt:.2f}")
            except ValueError:
                pass  # Игнорируем некорректный ввод

        def update_from_pt(var_pt, var_mm, var_twips):
            """Обновляет мм и twips при изменении pt."""
            try:
                pt = float(var_pt.get().replace(',', '.'))
                twips = int(round(pt * PT_TO_TWIPS))
                var_twips.set(str(twips))
                mm = twips / MM_TO_TWIPS
                var_mm.set(f"{mm:.2f}")
            except ValueError:
                pass

        # Привязка событий
        def make_trace_mm(var_mm, var_pt, var_twips):
            def callback(*args):
                update_from_mm(var_mm, var_pt, var_twips)
            return callback

        def make_trace_pt(var_pt, var_mm, var_twips):
            def callback(*args):
                update_from_pt(var_pt, var_mm, var_twips)
            return callback

        font_var_mm.trace_add("write", make_trace_mm(font_var_mm, font_var_pt, font_var_twips))
        font_var_pt.trace_add("write", make_trace_pt(font_var_pt, font_var_mm, font_var_twips))
        indent_var_mm.trace_add("write", make_trace_mm(indent_var_mm, indent_var_pt, indent_var_twips))
        indent_var_pt.trace_add("write", make_trace_pt(indent_var_pt, indent_var_mm, indent_var_twips))
        spacing_var_mm.trace_add("write", make_trace_mm(spacing_var_mm, spacing_var_pt, spacing_var_twips))
        spacing_var_pt.trace_add("write", make_trace_pt(spacing_var_pt, spacing_var_mm, spacing_var_twips))

        # Кнопки сохранения/отмены
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=6, column=0, columnspan=4, pady=20)

        def save_settings():
            try:
                # Берем значения из twips (они всегда актуальны)
                new_font = int(font_var_twips.get())
                new_indent = int(indent_var_twips.get())
                new_spacing = int(spacing_var_twips.get())
                if new_font < 0 or new_indent < 0 or new_spacing < 0:
                    raise ValueError("Допуски не могут быть отрицательными")
                # Обновляем атрибуты GUI
                self.font_tolerance = new_font
                self.indent_tolerance = new_indent
                self.spacing_tolerance = new_spacing
                # Обновляем значения в контроллере
                self.controller.font_tolerance = new_font
                self.controller.indent_tolerance = new_indent
                self.controller.spacing_tolerance = new_spacing
                # Обновляем значения в аудит-движке (если он существует)
                if self.controller.audit_engine:
                    self.controller.audit_engine.FONT_SIZE_TOLERANCE_TWIPS = new_font
                    self.controller.audit_engine.INDENT_TOLERANCE_TWIPS = new_indent
                    self.controller.audit_engine.SPACING_TOLERANCE_TWIPS = new_spacing
                # Также можно сохранить в конфиг (дополнительно)
                self.log(f"Допуски обновлены: шрифт={new_font} twips ({float(font_var_mm.get()):.2f} мм, {float(font_var_pt.get()):.2f} pt), отступ={new_indent} twips ({float(indent_var_mm.get()):.2f} мм, {float(indent_var_pt.get()):.2f} pt), интервал={new_spacing} twips ({float(spacing_var_mm.get()):.2f} мм, {float(spacing_var_pt.get()):.2f} pt)")
                messagebox.showinfo("Сохранено", "Допуски успешно сохранены.")
                settings_win.destroy()
            except ValueError as e:
                messagebox.showerror("Ошибка", f"Некорректное значение: {e}")

        ttk.Button(btn_frame, text="Сохранить", command=save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=settings_win.destroy).pack(side=tk.LEFT, padx=5)

    def open_ignore_settings(self):
        """Открывает окно настроек игнорирования типов ошибок."""
        # Создаем дочернее окно
        ignore_win = tk.Toplevel(self.root)
        ignore_win.title("Настройки игнорирования ошибок")
        ignore_win.geometry("500x500")
        ignore_win.resizable(False, False)
        ignore_win.transient(self.root)
        ignore_win.grab_set()

        # Получаем текущие настройки из контроллера
        current_settings = self.controller.get_ignore_settings()
        current_categories = current_settings.get('categories', [])
        current_severities = current_settings.get('severities', [])

        # Фрейм для категорий
        cat_frame = ttk.LabelFrame(ignore_win, text="Категории ошибок", padding=10)
        cat_frame.pack(fill=tk.X, padx=10, pady=(10,5))

        # Список категорий
        categories = [
            ("FONT", "Шрифт"),
            ("PARAGRAPH", "Абзац"),
            ("STYLE", "Стиль"),
            ("TABLE", "Таблица"),
            ("HEADER_FOOTER", "Колонтитулы"),
            ("FIGURE", "Рисунки"),
            ("PAGE", "Настройки страницы")
        ]
        cat_vars = {}
        for i, (code, label) in enumerate(categories):
            var = tk.BooleanVar(master=ignore_win, value=(code in current_categories))
            cb = ttk.Checkbutton(cat_frame, text=label, variable=var)
            cb.grid(row=i//2, column=i%2, sticky=tk.W, padx=10, pady=5)
            cat_vars[code] = var

        # Фрейм для уровней важности
        sev_frame = ttk.LabelFrame(ignore_win, text="Уровни важности", padding=10)
        sev_frame.pack(fill=tk.X, padx=10, pady=5)

        severities = [
            (Severity.CRITICAL, "❌ Критические"),
            (Severity.WARNING, "⚠️ Предупреждения"),
            (Severity.INFO, "ℹ️ Информационные")
        ]
        sev_vars = {}
        for i, (sev, label) in enumerate(severities):
            var = tk.BooleanVar(master=ignore_win, value=(sev.value in current_severities))
            cb = ttk.Checkbutton(sev_frame, text=label, variable=var)
            cb.grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            sev_vars[sev] = var

        # Пояснение
        ttk.Label(ignore_win, text="Выбранные категории и уровни важности будут игнорироваться при аудите.",
                  font=("Arial", 9), foreground="gray").pack(pady=10)

        # Кнопки сохранения/отмены
        btn_frame = ttk.Frame(ignore_win)
        btn_frame.pack(pady=10)

        def save_ignore_settings():
            selected_categories = [code for code, var in cat_vars.items() if var.get()]
            selected_severities = [sev for sev, var in sev_vars.items() if var.get()]
            # Применяем настройки в контроллере
            self.controller.set_ignore_settings(selected_categories, selected_severities)
            self.log(f"Настройки игнорирования обновлены: категории={selected_categories}, важности={[s.value for s in selected_severities]}")
            messagebox.showinfo("Сохранено", "Настройки игнорирования успешно сохранены.")
            ignore_win.destroy()

        ttk.Button(btn_frame, text="Сохранить", command=save_ignore_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=ignore_win.destroy).pack(side=tk.LEFT, padx=5)

    def _export_report_txt(self, file_path):
        """Генерирует текстовый отчёт."""
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("ОТЧЁТ АУДИТА ДОКУМЕНТА\n")
            f.write("=" * 50 + "\n")
            f.write(f"Документ: {os.path.basename(self.controller.doc_path)}\n")
            f.write(f"Конфигурация: {self.controller.config_path}\n")
            f.write(f"Дата аудита: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("\n")
            
            # Сводка
            if self.controller.audit_engine:
                summary = self.controller.audit_engine.get_summary()
                f.write("СВОДКА:\n")
                f.write(f"  Всего проблем: {len(self.controller.current_issues)}\n")
                f.write(f"  Критических: {summary['CRITICAL']}\n")
                f.write(f"  Предупреждений: {summary['WARNING']}\n")
                f.write(f"  Информационных: {summary['INFO']}\n")
                f.write("\n")
            
            # Детали проблем
            f.write("ДЕТАЛИ ПРОБЛЕМ:\n")
            f.write("-" * 50 + "\n")
            for idx, issue in enumerate(self.controller.current_issues, 1):
                f.write(f"{idx}. ID: {issue.id}\n")
                f.write(f"   Важность: {issue.severity.value}\n")
                f.write(f"   Категория: {issue.category}\n")
                f.write(f"   Стиль: {issue.element_type}\n")
                f.write(f"   Локация: стр. ~{issue.location.get('page', '?')}\n")
                f.write(f"   Описание: {issue.description}\n")
                f.write(f"   Текущее значение: {issue.current_value}\n")
                f.write(f"   Ожидаемое значение: {issue.expected_value}\n")
                f.write(f"   Автоисправление: {'Да' if issue.auto_fixable else 'Нет'}\n")
                f.write(f"   Рекомендация: {issue.recommendation if issue.recommendation else '—'}\n")
                f.write("\n")

    def _export_report_html(self, file_path):
        """Генерирует HTML отчёт (заглушка)."""
        # Пока просто создадим простой HTML
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
            f.write(os.path.basename(self.controller.doc_path))
            f.write("""</p>
    <p><strong>Конфигурация:</strong> """)
            f.write(self.controller.config_path)
            f.write("""</p>
    <p><strong>Дата аудита:</strong> """)
            f.write(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            f.write("""</p>
    
    <h2>Сводка</h2>
    <ul>
        <li>Всего проблем: """)
            f.write(str(len(self.controller.current_issues)))
            f.write("""</li>
        <li>Критических: """)
            if self.controller.audit_engine:
                summary = self.controller.audit_engine.get_summary()
                f.write(str(summary['CRITICAL']))
            f.write("""</li>
        <li>Предупреждений: """)
            if self.controller.audit_engine:
                f.write(str(summary['WARNING']))
            f.write("""</li>
        <li>Информационных: """)
            if self.controller.audit_engine:
                f.write(str(summary['INFO']))
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
            for idx, issue in enumerate(self.controller.current_issues, 1):
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
                f.write(f"                <td>{issue.recommendation if issue.recommendation else '—'}</td>\n")
                f.write("            </tr>\n")
            f.write("""        </tbody>
    </table>
</body>
</html>""")

if __name__ == "__main__":
    root = None
    # Применение темы (если есть ttkthemes, иначе стандарт)
    try:
        from ttkthemes import ThemedTk
        root = ThemedTk(theme="arc")
    except ImportError:
        root = tk.Tk()
        
    app = GOSTFormatterGUI(root)
    root.mainloop()

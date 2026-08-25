"""
WizardMixin — пошаговый мастер MD → DOCX.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import time
import logging

logger = logging.getLogger(__name__)


class WizardMixin:
    """Миксин для вкладки 'MD → DOCX' — пошаговый мастер."""

    def _create_wizard_tab(self, parent):
        """Создаёт содержимое вкладки 'MD → DOCX'."""
        main = ttk.Frame(parent, padding=10)

        # === Индикатор шагов ===
        steps_frame = ttk.Frame(main)
        steps_frame.pack(fill=tk.X, pady=(0, 15))

        self._step_vars = {}
        step_labels = ['1. Файл', '2. Конфиг', '3. Результат']
        for i, label in enumerate(step_labels):
            lbl = ttk.Label(steps_frame, text=label, font=('', 10, 'bold'))
            lbl.pack(side=tk.LEFT, padx=20)
            self._step_vars[i] = lbl

        ttk.Separator(steps_frame, orient='horizontal').pack(fill=tk.X, pady=5)

        # === Шаг 1: Файл ===
        file_frame = ttk.LabelFrame(main, text="📄 Шаг 1: Исходный файл", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(file_frame, text="Markdown файл:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self._wiz_md_path = ttk.Entry(file_frame, width=60)
        self._wiz_md_path.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Обзор...", command=self._wiz_browse_md).grid(row=0, column=2, pady=5)

        # === Шаг 2: Настройки ===
        config_frame = ttk.LabelFrame(main, text="⚙ Шаг 2: Настройки", padding=10)
        config_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(config_frame, text="Конфиг:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self._wiz_config_path = ttk.Entry(config_frame, width=60)
        self._wiz_config_path.insert(0, "configs/active/config.yaml")
        self._wiz_config_path.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(config_frame, text="Обзор...", command=self._wiz_browse_config).grid(row=0, column=2, pady=5)
        ttk.Button(config_frame, text="Ред.", command=self._wiz_open_config_editor).grid(row=0, column=3, padx=(5,0), pady=5)

        ttk.Label(config_frame, text="Шаблон:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self._wiz_template_path = ttk.Entry(config_frame, width=60)
        self._wiz_template_path.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(config_frame, text="Обзор...", command=self._wiz_browse_template).grid(row=1, column=2, pady=5)

        # === Шаг 3: Действия ===
        action_frame = ttk.LabelFrame(main, text="▶ Шаг 3: Действия", padding=10)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        btn_row = ttk.Frame(action_frame)
        btn_row.pack(fill=tk.X)

        self._wiz_btn_convert = ttk.Button(btn_row, text="🔄 Конвертировать", command=self._wiz_convert)
        self._wiz_btn_convert.pack(side=tk.LEFT, padx=5)

        self._wiz_btn_audit = ttk.Button(btn_row, text="🔍 Аудит", command=self._wiz_audit, state=tk.DISABLED)
        self._wiz_btn_audit.pack(side=tk.LEFT, padx=5)

        self._wiz_btn_save = ttk.Button(btn_row, text="💾 Сохранить", command=self._wiz_save, state=tk.DISABLED)
        self._wiz_btn_save.pack(side=tk.LEFT, padx=5)

        # === Прогресс ===
        self._wiz_progress_var = tk.DoubleVar()
        self._wiz_progress = ttk.Progressbar(action_frame, variable=self._wiz_progress_var,
                                              maximum=100, mode='determinate')
        self._wiz_progress.pack(fill=tk.X, pady=(10, 5))

        self._wiz_status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(action_frame, textvariable=self._wiz_status_var).pack(fill=tk.X)

        # === Лог ===
        log_frame = ttk.LabelFrame(main, text="📋 Лог", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        from tkinter import scrolledtext
        self._wiz_log = scrolledtext.ScrolledText(log_frame, height=6, state=tk.DISABLED)
        self._wiz_log.pack(fill=tk.BOTH, expand=True)

        # Состояние
        self._wiz_docx_path = None
        self._wiz_converted = False

        return main

    def _wiz_log_msg(self, msg):
        """Добавляет сообщение в лог мастера."""
        self._wiz_log.config(state=tk.NORMAL)
        self._wiz_log.insert(tk.END, f"{time.strftime('%H:%M')} - {msg}\n")
        self._wiz_log.see(tk.END)
        self._wiz_log.config(state=tk.DISABLED)

    def _wiz_set_status(self, msg):
        """Обновляет статус."""
        self._wiz_status_var.set(msg)

    def _wiz_update_steps(self, active=0):
        """Подсвечивает текущий шаг."""
        colors = {0: 'black', 1: 'black', 2: 'black'}
        colors[active] = '#0066CC'
        for i, lbl in self._step_vars.items():
            lbl.config(foreground=colors[i])

    # === Браузеры ===

    def _wiz_browse_md(self):
        path = filedialog.askopenfilename(
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt"), ("All", "*.*")],
            title="Выберите Markdown файл"
        )
        if path:
            self._wiz_md_path.delete(0, tk.END)
            self._wiz_md_path.insert(0, path)
            self._wiz_update_steps(1)

    def _wiz_browse_config(self):
        path = filedialog.askopenfilename(
            filetypes=[("YAML", "*.yaml *.yml"), ("JSON", "*.json"), ("All", "*.*")],
            title="Выберите конфиг"
        )
        if path:
            self._wiz_config_path.delete(0, tk.END)
            self._wiz_config_path.insert(0, path)
            self._wiz_update_steps(1)

    def _wiz_browse_template(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx"), ("All", "*.*")],
            title="Выберите шаблон"
        )
        if path:
            self._wiz_template_path.delete(0, tk.END)
            self._wiz_template_path.insert(0, path)

    def _wiz_open_config_editor(self):
        config_path = self._wiz_config_path.get().strip()
        if config_path and os.path.exists(config_path):
            try:
                from src.config_editor.dialog import ConfigEditorDialog
                ConfigEditorDialog(self.root, config_path)
            except Exception as e:
                self._wiz_log_msg(f"Ошибка: {e}")

    # === Конвертация ===

    def _wiz_convert(self):
        md_path = self._wiz_md_path.get().strip()
        if not md_path or not os.path.exists(md_path):
            messagebox.showwarning("Внимание", "Выберите Markdown файл.")
            return

        config_path = self._wiz_config_path.get().strip()
        if not config_path or not os.path.exists(config_path):
            config_path = "configs/active/config.yaml"

        # Выбор пути сохранения
        default_name = os.path.splitext(os.path.basename(md_path))[0] + "_formatted.docx"
        docx_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word", "*.docx")],
            title="Сохранить как"
        )
        if not docx_path:
            return

        self._wiz_update_steps(2)
        self._wiz_set_status("Конвертация...")
        self._wiz_btn_convert.config(state=tk.DISABLED)
        self._wiz_progress_var.set(0)

        def convert():
            try:
                from md_to_docx import RI2013Converter
                converter = RI2013Converter(config_path)
                converter.convert(md_path, docx_path)
                self._wiz_docx_path = docx_path
                self._wiz_converted = True
                self.root.after(0, lambda: (
                    self._wiz_log_msg(f"Конвертация завершена: {os.path.basename(docx_path)}"),
                    self._wiz_set_status("Конвертация завершена"),
                    self._wiz_progress_var.set(100),
                    self._wiz_btn_convert.config(state=tk.NORMAL),
                    self._wiz_btn_audit.config(state=tk.NORMAL),
                    self._wiz_btn_save.config(state=tk.NORMAL),
                    self._wiz_update_steps(2),
                ))
            except Exception as e:
                self.root.after(0, lambda: (
                    self._wiz_log_msg(f"Ошибка: {e}"),
                    self._wiz_set_status("Ошибка конвертации"),
                    self._wiz_btn_convert.config(state=tk.NORMAL),
                ))

        threading.Thread(target=convert, daemon=True).start()

    # === Аудит ===

    def _wiz_audit(self):
        if not self._wiz_docx_path:
            return
        self._wiz_set_status("Аудит...")
        self._wiz_btn_audit.config(state=tk.DISABLED)

        def audit():
            try:
                from src.core.config_loader import ConfigLoader
                from src.core.audit_engine import AuditEngine

                config_path = self._wiz_config_path.get().strip()
                loader = ConfigLoader(config_path)
                loader.load()

                engine = AuditEngine(loader)
                issues = engine.audit_document(self._wiz_docx_path)

                self._wiz_issues = issues
                crits = sum(1 for i in issues if i.severity.value == 'CRITICAL')
                warns = sum(1 for i in issues if i.severity.value == 'WARNING')

                self.root.after(0, lambda: (
                    self._wiz_log_msg(f"Аудит: {crits} критических, {warns} предупреждений"),
                    self._wiz_set_status(f"Аудит завершён: {crits} крит., {warns} пред."),
                    self._wiz_btn_audit.config(state=tk.NORMAL),
                    self._wiz_btn_fix.config(state=tk.NORMAL if issues else tk.DISABLED),
                ))
            except Exception as e:
                self.root.after(0, lambda: (
                    self._wiz_log_msg(f"Ошибка аудита: {e}"),
                    self._wiz_set_status("Ошибка аудита"),
                    self._wiz_btn_audit.config(state=tk.NORMAL),
                ))

        threading.Thread(target=audit, daemon=True).start()

    # === Исправление ===

    def _wiz_fix(self):
        if not self._wiz_docx_path or not hasattr(self, '_wiz_issues'):
            return
        self._wiz_set_status("Исправление...")

        def fix():
            try:
                from src.core.config_loader import ConfigLoader
                from src.apply.apply_orchestrator import ApplyOrchestrator

                config_path = self._wiz_config_path.get().strip()
                loader = ConfigLoader(config_path)
                loader.load()

                orchestrator = ApplyOrchestrator(loader)
                fixed = orchestrator.apply_all(self._wiz_docx_path, self._wiz_issues)

                self.root.after(0, lambda: (
                    self._wiz_log_msg(f"Исправлено: {fixed} проблем"),
                    self._wiz_set_status(f"Исправлено {fixed} проблем"),
                ))
            except Exception as e:
                self.root.after(0, lambda: (
                    self._wiz_log_msg(f"Ошибка исправления: {e}"),
                    self._wiz_set_status("Ошибка исправления"),
                ))

        threading.Thread(target=fix, daemon=True).start()

    # === Сохранение ===

    def _wiz_save(self):
        if not self._wiz_docx_path:
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=os.path.basename(self._wiz_docx_path),
            filetypes=[("Word", "*.docx")],
            title="Сохранить результат"
        )
        if save_path:
            import shutil
            shutil.copy2(self._wiz_docx_path, save_path)
            self._wiz_log_msg(f"Сохранено: {save_path}")
            self._wiz_set_status(f"Сохранено: {os.path.basename(save_path)}")

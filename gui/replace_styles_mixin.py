"""
ReplaceStylesMixin — вкладка "Замена стилей" главного окна.
Полная замена styles.xml из DOCX-источника в целевой документ.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import logging

logger = logging.getLogger(__name__)


class ReplaceStylesMixin:
    """Миксин — вкладка полной замены стилей (ReplaceStyles)."""

    def _create_replace_styles_tab(self, parent):
        """Создаёт содержимое вкладки 'Замена стилей'."""
        main = ttk.Frame(parent, padding=10)

        # === Источник стилей ===
        src_frame = ttk.LabelFrame(main, text="Источник стилей (откуда берём)", padding=10)
        src_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(src_frame, text="DOCX-шаблон:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self._rs_src_entry = ttk.Entry(src_frame, width=65)
        self._rs_src_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(src_frame, text="Обзор...", command=self._rs_browse_source).grid(row=0, column=2, pady=5)

        btn_row = ttk.Frame(src_frame)
        btn_row.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        ttk.Button(btn_row, text="Показать стили источника", command=self._rs_show_source_styles).pack(side=tk.LEFT)
        self._rs_src_info = ttk.Label(btn_row, text="", foreground="gray")
        self._rs_src_info.pack(side=tk.LEFT, padx=(10, 0))

        self._rs_src_styles = scrolledtext.ScrolledText(src_frame, height=6, state=tk.DISABLED, font=("Consolas", 9))
        self._rs_src_styles.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=(5, 0))

        # === Целевой документ ===
        tgt_frame = ttk.LabelFrame(main, text="Целевой документ (куда применяем)", padding=10)
        tgt_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(tgt_frame, text="DOCX-файл:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self._rs_tgt_entry = ttk.Entry(tgt_frame, width=65)
        self._rs_tgt_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(tgt_frame, text="Обзор...", command=self._rs_browse_target).grid(row=0, column=2, pady=5)

        self._rs_warn = ttk.Label(
            tgt_frame,
            text="Внимание: все стили целевого документа будут заменены на стили из источника.",
            foreground="red",
            wraplength=550,
        )
        self._rs_warn.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

        # === Результат ===
        out_frame = ttk.LabelFrame(main, text="Результат", padding=10)
        out_frame.pack(fill=tk.X, pady=(0, 10))

        self._rs_output_var = tk.StringVar(value="create")
        ttk.Radiobutton(
            out_frame, text="Создать новый документ", variable=self._rs_output_var, value="create",
            command=self._rs_toggle_output,
        ).grid(row=0, column=0, sticky=tk.W, pady=3)

        out_path_frame = ttk.Frame(out_frame)
        out_path_frame.grid(row=0, column=1, columnspan=2, sticky=tk.EW, padx=(10, 0), pady=3)

        self._rs_out_entry = ttk.Entry(out_path_frame, width=55)
        self._rs_out_entry.pack(side=tk.LEFT, padx=(0, 5))
        self._rs_out_btn = ttk.Button(out_path_frame, text="Обзор...", command=self._rs_browse_output)
        self._rs_out_btn.pack(side=tk.LEFT)

        ttk.Radiobutton(
            out_frame, text="Перезаписать целевой документ", variable=self._rs_output_var, value="overwrite",
            command=self._rs_toggle_output,
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=3)

        # === Запуск ===
        run_frame = ttk.Frame(main)
        run_frame.pack(fill=tk.X, pady=(0, 5))

        self._rs_run_btn = ttk.Button(run_frame, text="Заменить стили", command=self._rs_run_replace)
        self._rs_run_btn.pack(side=tk.LEFT)

        self._rs_progress_var = tk.DoubleVar()
        self._rs_progress_bar = ttk.Progressbar(
            run_frame, variable=self._rs_progress_var, maximum=100, mode="determinate"
        )
        self._rs_progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # === Лог ===
        log_frame = ttk.LabelFrame(main, text="Результат", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self._rs_log = scrolledtext.ScrolledText(log_frame, height=6, state=tk.DISABLED, font=("Consolas", 9))
        self._rs_log.pack(fill=tk.BOTH, expand=True)

        return main

    # ------------------------------------------------------------------
    # Браузеры
    # ------------------------------------------------------------------

    def _rs_browse_source(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")],
            title="Выберите документ-источник стилей",
        )
        if path:
            self._rs_src_entry.delete(0, tk.END)
            self._rs_src_entry.insert(0, path)

    def _rs_browse_target(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")],
            title="Выберите целевой документ",
        )
        if path:
            self._rs_tgt_entry.delete(0, tk.END)
            self._rs_tgt_entry.insert(0, path)
            # Автоподстановка имени результата
            if self._rs_output_var.get() == "create":
                default_name = os.path.splitext(os.path.basename(path))[0] + "_styles.docx"
                default_dir = os.path.dirname(path)
                self._rs_out_entry.delete(0, tk.END)
                self._rs_out_entry.insert(0, os.path.join(default_dir, default_name))

    def _rs_browse_output(self):
        target = self._rs_tgt_entry.get().strip()
        initial = ""
        if target:
            initial = os.path.splitext(os.path.basename(target))[0] + "_styles.docx"
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=initial,
            filetypes=[("Word", "*.docx")],
            title="Сохранить результат",
        )
        if path:
            self._rs_out_entry.delete(0, tk.END)
            self._rs_out_entry.insert(0, path)

    def _rs_toggle_output(self):
        if self._rs_output_var.get() == "overwrite":
            self._rs_out_entry.config(state=tk.DISABLED)
            self._rs_out_btn.config(state=tk.DISABLED)
        else:
            self._rs_out_entry.config(state=tk.NORMAL)
            self._rs_out_btn.config(state=tk.NORMAL)

    # ------------------------------------------------------------------
    # Показать стили источника
    # ------------------------------------------------------------------

    def _rs_show_source_styles(self):
        path = self._rs_src_entry.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Внимание", "Укажите существующий DOCX-файл.")
            return

        def load():
            try:
                from docx import Document
                doc = Document(path)
                styles = {}
                for style in doc.styles:
                    if style.type and style.type.name == "PARAGRAPH":
                        styles[style.name] = style

                lines = [f"Параграфных стилей: {len(styles)}", ""]
                for name in sorted(styles.keys())[:50]:
                    lines.append(f"  {name}")

                text = "\n".join(lines)
                self.root.after(0, lambda: (
                    self._rs_src_styles.config(state=tk.NORMAL),
                    self._rs_src_styles.delete("1.0", tk.END),
                    self._rs_src_styles.insert("1.0", text),
                    self._rs_src_styles.config(state=tk.DISABLED),
                    self._rs_src_info.config(text=f"Найдено: {len(styles)} стилей"),
                ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

        threading.Thread(target=load, daemon=True).start()

    # ------------------------------------------------------------------
    # Запуск замены
    # ------------------------------------------------------------------

    def _rs_run_replace(self):
        source = self._rs_src_entry.get().strip()
        target = self._rs_tgt_entry.get().strip()

        if not source or not os.path.exists(source):
            messagebox.showwarning("Внимание", "Выберите документ-источник стилей.")
            return
        if not target or not os.path.exists(target):
            messagebox.showwarning("Внимание", "Выберите целевой документ.")
            return

        output_path = None
        if self._rs_output_var.get() == "create":
            output_path = self._rs_out_entry.get().strip()
            if not output_path:
                messagebox.showwarning("Внимание", "Укажите путь для результата.")
                return

        self._rs_run_btn.config(state=tk.DISABLED, text="Выполняется...")
        self._rs_progress_var.set(0)

        def run():
            try:
                from src.apply.style_applier import StyleApplier

                self.root.after(0, lambda: self._rs_progress_var.set(20))
                self._rs_log_message(f"Источник: {os.path.basename(source)}")
                self._rs_log_message(f"Цель: {os.path.basename(target)}")

                applier = StyleApplier(source)
                self.root.after(0, lambda: self._rs_progress_var.set(50))

                stats = applier.apply(
                    target,
                    output_path=output_path,
                    replace_styles_raw=True,
                )

                self.root.after(0, lambda: self._rs_progress_var.set(100))

                result_path = output_path or target
                self._rs_log_message(f"Готово! Результат: {result_path}")

                self.root.after(0, lambda: (
                    self._rs_run_btn.config(state=tk.NORMAL, text="Заменить стили"),
                    messagebox.showinfo("Готово", f"Стили заменены!\n\n{result_path}"),
                ))

            except Exception as e:
                self.root.after(0, lambda: (
                    self._rs_run_btn.config(state=tk.NORMAL, text="Заменить стили"),
                    self._rs_progress_var.set(0),
                    self._rs_log_message(f"ОШИБКА: {e}"),
                    messagebox.showerror("Ошибка", str(e)),
                ))

        threading.Thread(target=run, daemon=True).start()

    def _rs_log_message(self, msg):
        """Добавляет строку в лог вкладки."""
        def _do():
            self._rs_log.config(state=tk.NORMAL)
            self._rs_log.insert(tk.END, msg + "\n")
            self._rs_log.see(tk.END)
            self._rs_log.config(state=tk.DISABLED)
        self.root.after(0, _do)

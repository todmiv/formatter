"""
TemplateMixin — вкладка "Шаблоны".
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import logging

logger = logging.getLogger(__name__)


class TemplateMixin:
    """Миксин для вкладки 'Шаблоны'."""

    def _create_template_tab(self, parent):
        """Создаёт содержимое вкладки 'Шаблоны'."""
        main = ttk.Frame(parent, padding=10)

        # ============================================================
        # ШАГ 1 — Извлечение конфига (необязательный)
        # ============================================================
        step1 = ttk.LabelFrame(main, text="  Шаг 1 (необязательный). Создать YAML-конфиг из DOCX-шаблона  ", padding=10)
        step1.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(
            step1,
            text="Извлекает стили, поля страницы и настройки из готового DOCX-шаблона\n"
                 "и сохраняет их в YAML-файл. Этот конфиг нужен для вкладки «Аудит».\n"
                 "Если вы хотите просто отформатировать документ — перейдите к шагу 2.",
            foreground="gray",
            wraplength=600,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        ttk.Label(step1, text="DOCX-шаблон (источник стилей):").grid(row=1, column=0, sticky=tk.W, pady=4)
        self._tmpl_extract_src = ttk.Entry(step1, width=60)
        self._tmpl_extract_src.grid(row=1, column=1, padx=5, pady=4)
        ttk.Button(step1, text="Обзор...", command=self._tmpl_browse_extract).grid(row=1, column=2, pady=4)

        ttk.Label(step1, text="Сохранить YAML-конфиг как:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self._tmpl_extract_dst = ttk.Entry(step1, width=60)
        self._tmpl_extract_dst.grid(row=2, column=1, padx=5, pady=4)
        ttk.Button(step1, text="Обзор...", command=self._tmpl_browse_extract_dst).grid(row=2, column=2, pady=4)

        ttk.Button(step1, text="Извлечь и сохранить конфиг", command=self._tmpl_extract).grid(
            row=3, column=1, sticky=tk.W, pady=(8, 0)
        )

        # ============================================================
        # ШАГ 2 — Применение шаблона к документу
        # ============================================================
        step2 = ttk.LabelFrame(main, text="  Шаг 2. Применить шаблон к документу  ", padding=10)
        step2.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(
            step2,
            text="Берёт ваш DOCX-документ и переносит его текст в шаблон,\n"
                 "сохраняя стили, колонтитулы и оформление шаблона.\n"
                 "Оригинальный файл НЕ изменяется — результат сохраняется в новый файл.",
            foreground="gray",
            wraplength=600,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        ttk.Label(step2, text="Ваш документ (текст):").grid(row=1, column=0, sticky=tk.W, pady=4)
        self._tmpl_apply_src = ttk.Entry(step2, width=60)
        self._tmpl_apply_src.grid(row=1, column=1, padx=5, pady=4)
        ttk.Button(step2, text="Обзор...", command=self._tmpl_browse_apply).grid(row=1, column=2, pady=4)

        ttk.Label(step2, text="DOCX-шаблон (оформление):").grid(row=2, column=0, sticky=tk.W, pady=4)
        self._tmpl_apply_tmpl = ttk.Entry(step2, width=60)
        self._tmpl_apply_tmpl.grid(row=2, column=1, padx=5, pady=4)
        ttk.Button(step2, text="Обзор...", command=self._tmpl_browse_apply_tmpl).grid(row=2, column=2, pady=4)

        ttk.Button(step2, text="Применить шаблон → сохранить результат", command=self._tmpl_apply).grid(
            row=3, column=1, sticky=tk.W, pady=(8, 0)
        )

        # ============================================================
        # СПРАВКА — Конфиг и стили
        # ============================================================
        info = ttk.LabelFrame(main, text="  Справка  ", padding=10)
        info.pack(fill=tk.BOTH, expand=True)

        info_row = ttk.Frame(info)
        info_row.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(info_row, text="Конфиг:").pack(side=tk.LEFT)
        self._tmpl_edit_path = ttk.Entry(info_row, width=50)
        self._tmpl_edit_path.insert(0, "configs/active/config.yaml")
        self._tmpl_edit_path.pack(side=tk.LEFT, padx=5)
        ttk.Button(info_row, text="Обзор...", command=self._tmpl_browse_edit).pack(side=tk.LEFT)
        ttk.Button(info_row, text="Открыть редактор", command=self._tmpl_open_editor).pack(side=tk.LEFT, padx=(5, 0))

        self._tmpl_info = scrolledtext.ScrolledText(info, height=6, state=tk.DISABLED)
        self._tmpl_info.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        ttk.Button(info, text="Показать стили шаблона", command=self._tmpl_show_styles).pack(pady=(5, 0))

        return main

    # === Браузеры ===

    def _tmpl_browse_extract(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")],
            title="Выберите DOCX-шаблон"
        )
        if path:
            self._tmpl_extract_src.delete(0, tk.END)
            self._tmpl_extract_src.insert(0, path)
            default = "config_from_" + os.path.splitext(os.path.basename(path))[0] + ".yaml"
            self._tmpl_extract_dst.delete(0, tk.END)
            self._tmpl_extract_dst.insert(0, os.path.join("configs/active", default))

    def _tmpl_browse_extract_dst(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".yaml",
            filetypes=[("YAML", "*.yaml *.yml")],
            title="Сохранить конфиг"
        )
        if path:
            self._tmpl_extract_dst.delete(0, tk.END)
            self._tmpl_extract_dst.insert(0, path)

    def _tmpl_browse_edit(self):
        path = filedialog.askopenfilename(
            filetypes=[("YAML", "*.yaml *.yml"), ("JSON", "*.json")],
            title="Выберите конфиг"
        )
        if path:
            self._tmpl_edit_path.delete(0, tk.END)
            self._tmpl_edit_path.insert(0, path)

    def _tmpl_browse_apply(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")],
            title="Выберите ваш документ"
        )
        if path:
            self._tmpl_apply_src.delete(0, tk.END)
            self._tmpl_apply_src.insert(0, path)

    def _tmpl_browse_apply_tmpl(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")],
            title="Выберите DOCX-шаблон"
        )
        if path:
            self._tmpl_apply_tmpl.delete(0, tk.END)
            self._tmpl_apply_tmpl.insert(0, path)

    def _tmpl_apply(self):
        src = self._tmpl_apply_src.get().strip()
        tmpl = self._tmpl_apply_tmpl.get().strip()

        if not src or not os.path.exists(src):
            messagebox.showwarning("Внимание", "Выберите ваш документ.")
            return
        if not tmpl or not os.path.exists(tmpl):
            messagebox.showwarning("Внимание", "Выберите DOCX-шаблон.")
            return

        dst = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=os.path.splitext(os.path.basename(src))[0] + "_formatted.docx",
            filetypes=[("Word", "*.docx")],
            title="Сохранить результат"
        )
        if not dst:
            return

        def apply():
            try:
                import shutil
                from docx import Document

                shutil.copy2(tmpl, dst)
                src_doc = Document(src)
                dst_doc = Document(dst)

                body = dst_doc.element.body
                for child in list(body):
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if tag in ('p', 'tbl'):
                        body.remove(child)

                for para in src_doc.paragraphs:
                    if para.text.strip():
                        new_para = dst_doc.add_paragraph(para.text)
                        try:
                            new_para.style = dst_doc.styles[para.style.name]
                        except KeyError:
                            pass
                dst_doc.save(dst)
                self.root.after(0, lambda: (
                    messagebox.showinfo("Готово", f"Документ создан:\n{dst}\n\nШаблон: {os.path.basename(tmpl)}"),
                ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

        threading.Thread(target=apply, daemon=True).start()

    # === Действия ===

    def _tmpl_extract(self):
        src = self._tmpl_extract_src.get().strip()
        dst = self._tmpl_extract_dst.get().strip()

        if not src or not os.path.exists(src):
            messagebox.showwarning("Внимание", "Выберите DOCX-шаблон.")
            return
        if not dst:
            messagebox.showwarning("Внимание", "Укажите путь для сохранения конфига.")
            return

        def extract():
            try:
                from src.core.template_extractor import TemplateExtractor
                extractor = TemplateExtractor(src)
                config = extractor.extract()
                extractor.save_yaml(dst, config)
                self.root.after(0, lambda: (
                    messagebox.showinfo("Готово", f"Конфиг извлечён:\n{dst}\n\nСтилей: {len(config.get('styles', {}))}"),
                ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

        threading.Thread(target=extract, daemon=True).start()

    def _tmpl_open_editor(self):
        path = self._tmpl_edit_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Внимание", "Выберите конфиг.")
            return
        try:
            from src.config_editor.dialog import ConfigEditorDialog
            ConfigEditorDialog(self.root, path)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _tmpl_show_styles(self):
        src = self._tmpl_apply_src.get().strip() or self._tmpl_apply_tmpl.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning("Внимание", "Сначала выберите DOCX-файл (документ или шаблон).")
            return

        def load():
            try:
                from docx import Document
                doc = Document(src)
                styles = {}
                for p in doc.paragraphs:
                    sn = p.style.name
                    if sn not in styles:
                        styles[sn] = 0
                    styles[sn] += 1

                lines = [f"Файл: {os.path.basename(src)}", f"Параграфных стилей: {len(styles)}", ""]
                for name, cnt in sorted(styles.items(), key=lambda x: -x[1])[:30]:
                    lines.append(f"  {cnt:4d}  {name}")

                text = '\n'.join(lines)
                self.root.after(0, lambda: (
                    self._tmpl_info.config(state=tk.NORMAL),
                    self._tmpl_info.delete('1.0', tk.END),
                    self._tmpl_info.insert('1.0', text),
                    self._tmpl_info.config(state=tk.DISABLED),
                ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

        threading.Thread(target=load, daemon=True).start()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт преобразования Markdown в DOCX по НТД 01-2013
проектная организация - Проектно-производственный департамент
Рабочая Инструкция НТД 01-2013 "Порядок оформления проектной градостроительной документации"
Редакция 4
"""

import re
import yaml
import time
import logging
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
from src.core.page_manager import PageManager
from src.core.config_loader import ConfigLoader
from src.core.style_manager import StyleManager

# Настройка логгера для замеров времени
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
perf_logger = logging.getLogger('perf')


class RI2013Converter:
    """Конвертер Markdown в DOCX по НТД 01-2013"""
    
    def __init__(self, config_path='configs/active/config.yaml'):
        """Инициализация конвертера с загрузкой конфигурации"""
        perf_logger.info("--- ИНИЦИАЛИЗАЦИЯ RI2013Converter ---")
        
        # Этап 0: Загрузка конфигурации
        t0 = time.perf_counter()
        self.config_loader = ConfigLoader(config_path)
        try:
            self.config_loader.load()
        except FileNotFoundError:
            print(f"⚠️ Файл конфигурации {config_path} не найден. Используются настройки по умолчанию.")
            self.config = self._get_default_config()
            self.config_loader = None
        else:
            self.config = self.config_loader.get_config()
        t_config = time.perf_counter() - t0
        perf_logger.info(f"[PERF] Загрузка конфигурации: {t_config:.4f} сек")
        
        print("[DEBUG] Config keys:", self.config.keys())
        if 'page_setup' in self.config:
            print("[DEBUG] page_setup keys:", self.config['page_setup'].keys())
        
        # Создаём менеджер стилей
        if self.config_loader:
            self.style_manager = StyleManager(self.config_loader)
        else:
            self.style_manager = None
        
        # Этап: Создание документа Document()
        t0 = time.perf_counter()
        self.doc = Document()
        t_doc_create = time.perf_counter() - t0
        perf_logger.info(f"[PERF] Document(): {t_doc_create:.4f} сек")
        
        # Регулярное выражение для удаления эмодзи
        self._emoji_pattern = re.compile(
            "["
            "\U0001F600-\U0001F64F"  # эмоции
            "\U0001F300-\U0001F5FF"  # символы и пиктограммы
            "\U0001F680-\U0001F6FF"  # транспорт и карты
            "\U0001F1E0-\U0001F1FF"  # флаги
            "\U00002702-\U000027B0"  # дополнительные символы
            "\U000024C2-\U0001F251"  # прочие
            "]+",
            flags=re.UNICODE
        )
        
        # Этап: _setup_document() — PageManager + StyleManager
        t0 = time.perf_counter()
        self._setup_document()
        t_setup = time.perf_counter() - t0
        perf_logger.info(f"[PERF] _setup_document (PageManager + StyleManager): {t_setup:.4f} сек")
        
        self._current_section = 0
        self._current_chapter = 0
        
        t_total_init = t_config + t_doc_create + t_setup
        perf_logger.info(f"[PERF] ИТОГО __init__: {t_total_init:.4f} сек")
        perf_logger.info("--- ИНИЦИАЛИЗАЦИЯ ЗАВЕРШЕНА ---")
        
    def _load_config(self, config_path):
        """Загрузка конфигурации из YAML файла"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except FileNotFoundError:
            print(f"⚠️ Файл конфигурации {config_path} не найден. Используются настройки по умолчанию.")
            return self._get_default_config()
    
    def _get_default_config(self):
        """Конфигурация по умолчанию на основе НТД 01-2013"""
        return {
            'version': '4.2',
            'name': 'НТД 01-2013 Редакция 4',
            'styles': {
                'Normal': {
                    'font': {'name': 'Times New Roman', 'size': 12, 'bold': False, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 1.25,
                        'line_spacing': 1.15,
                        'space_before_pt': 0,
                        'space_after_pt': 0,
                        'alignment': 'justify'
                    }
                },
                'Heading 1': {
                    'font': {'name': 'Calibri Light', 'size': 16, 'bold': True, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 0,
                        'line_spacing': 1.25,
                        'space_before_pt': 12,
                        'space_after_pt': 3,
                        'alignment': 'left',
                        'page_break_before': True
                    }
                },
                'Heading 2': {
                    'font': {'name': 'Calibri Light', 'size': 14, 'bold': True, 'italic': True},
                    'paragraph': {
                        'first_line_indent_cm': 0,
                        'line_spacing': 1.25,
                        'space_before_pt': 12,
                        'space_after_pt': 3,
                        'alignment': 'left'
                    }
                },
                'Heading 3': {
                    'font': {'name': 'Cambria', 'size': 13, 'bold': True, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 0,
                        'line_spacing': 1.15,
                        'space_before_pt': 12,
                        'space_after_pt': 3,
                        'alignment': 'left'
                    }
                },
                'Heading 4': {
                    'font': {'name': 'Times New Roman', 'size': 12, 'bold': False, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 1.25,
                        'line_spacing': 1.25,
                        'space_before_pt': 0,
                        'space_after_pt': 0,
                        'alignment': 'left'
                    }
                },
                'Caption': {
                    'font': {'name': 'Times New Roman', 'size': 12, 'bold': False, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 0,
                        'line_spacing': 1.0,
                        'space_before_pt': 6,
                        'space_after_pt': 6,
                        'alignment': 'center'
                    }
                },
                'Table Grid': {
                    'font': {'name': 'Times New Roman', 'size': 12, 'bold': False, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 0,
                        'line_spacing': 1.0,
                        'space_before_pt': 0,
                        'space_after_pt': 6,
                        'alignment': 'left'
                    }
                },
                'Table Note': {
                    'font': {'name': 'Times New Roman', 'size': 12, 'bold': False, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 0,
                        'line_spacing': 1.0,
                        'space_before_pt': 0,
                        'space_after_pt': 6,
                        'alignment': 'left'
                    }
                },
                'List Paragraph': {
                    'font': {'name': 'Times New Roman', 'size': 12, 'bold': False, 'italic': False},
                    'paragraph': {
                        'first_line_indent_cm': 0,
                        'line_spacing': 1.15,
                        'space_before_pt': 0,
                        'space_after_pt': 0,
                        'alignment': 'left'
                    }
                }
            },
            'page_setup': {
                'portrait': {
                    'left_margin_cm': 2.5,
                    'right_margin_cm': 1.0,
                    'top_margin_cm': 2.0,
                    'bottom_margin_cm': 2.0
                },
                'landscape': {
                    'left_margin_cm': 2.0,
                    'right_margin_cm': 2.0,
                    'top_margin_cm': 2.5,
                    'bottom_margin_cm': 1.0
                },
                'paper_size': 'A4',
                'orientation_default': 'portrait'
            },
            'headers_footers': {
                'header': {
                    'enabled': True,
                    'position_from_edge_cm': 1.0,
                    'content': 'page_number',
                    'alignment': 'right'
                },
                'footer': {
                    'enabled': True,
                    'position_from_edge_cm': 1.0,
                    'content': 'section_chapter_number',
                    'alignment': 'center'
                }
            },
            'formatting_rules': {
                'tables': {
                    'header_row_height_cm': 0,
                    'header_row_height_mode': 'minimum',
                    'auto_fit_window': True,
                    'fixed_column_width': True,
                    'spacing_after_table_pt': 6,
                    'spacing_before_note_pt': 0,
                    'header_repeat_on_pages': True,
                    'column_numbering': True
                }
            }
        }
    
    def _setup_document(self):
        """Настройка документа согласно НТД 01-2013"""
        self.page_manager = PageManager(self.config)
        self.page_manager.apply_all(self.doc)
        
        if self.style_manager:
            self.style_manager.setup_styles(self.doc)
        else:
            self._setup_styles()
    
    def _setup_styles(self):
        """Настройка стилей документа"""
        styles_config = self.config['styles']
        
        # Настройка стиля Normal (п. 2.7.2)
        normal_style = self.doc.styles['Normal']
        normal_config = styles_config['Normal']
        self._apply_style_settings(normal_style, normal_config)
        
        # Настройка заголовков (Таблица 2.1)
        for heading_level in range(1, 5):
            style_name = f'Heading {heading_level}'
            if style_name in styles_config:
                heading_style = self.doc.styles[f'Heading {heading_level}']
                self._apply_style_settings(heading_style, styles_config[style_name])
        
        # Настройка дополнительных стилей
        additional_styles = ['Caption', 'Table Note', 'List Paragraph', 'Note']
        for style_name in additional_styles:
            if style_name in styles_config:
                # Проверяем, существует ли стиль в документе, если нет - создаём
                if style_name not in self.doc.styles:
                    # Создаём стиль параграфа
                    style = self.doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                else:
                    style = self.doc.styles[style_name]
                self._apply_style_settings(style, styles_config[style_name])
    
    def _apply_style_settings(self, style, config):
        """Применение настроек стиля"""
        font_config = config['font']
        style.font.name = font_config['name']
        style.font.size = Pt(font_config['size'])
        style.font.bold = font_config['bold']
        style.font.italic = font_config['italic']
        
        east_asia = font_config.get('east_asia', font_config['name'])
        style.element.rPr.rFonts.set(qn('w:eastAsia'), east_asia)
        
        # Абзац (п. 2.7.2, Таблица 2.1)
        paragraph_config = config['paragraph']
        style.paragraph_format.first_line_indent = Cm(paragraph_config['first_line_indent_cm'])
        style.paragraph_format.line_spacing = paragraph_config['line_spacing']
        style.paragraph_format.space_before = Pt(paragraph_config['space_before_pt'])
        style.paragraph_format.space_after = Pt(paragraph_config['space_after_pt'])
        
        # Выравнивание
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        style.paragraph_format.alignment = alignment_map.get(
            paragraph_config['alignment'], 
            WD_ALIGN_PARAGRAPH.JUSTIFY
        )
        
        # Разрыв страницы перед заголовком (Таблица 2.1 - РАЗДЕЛ)
        if paragraph_config.get('page_break_before', False):
            style.paragraph_format.page_break_before = True
    
    def convert(self, md_path, docx_path, progress_callback=None):
        """Конвертация Markdown файла в DOCX
        
        Args:
            md_path: путь к исходному Markdown файлу
            docx_path: путь для сохранения DOCX
            progress_callback: опциональный callback(stage, progress, total, message)
        """
        perf_logger.info("=" * 60)
        perf_logger.info(f"НАЧАЛО КОНВЕРТАЦИИ: {md_path} -> {docx_path}")
        perf_logger.info("=" * 60)
        t_start = time.perf_counter()
        
        # Этап 1: Загрузка Markdown файла
        t0 = time.perf_counter()
        print(f"📄 Чтение файла: {md_path}")
        with open(md_path, 'r', encoding='utf-8') as f:
            content = f.read()
        t_load = time.perf_counter() - t0
        perf_logger.info(f"[PERF] Загрузка Markdown файла: {t_load:.4f} сек")
        
        # Этап 2: Обработка содержимого (парсинг Markdown)
        t0 = time.perf_counter()
        print("🔄 Обработка документа...")
        self._process_content(content, progress_callback=progress_callback)
        t_process = time.perf_counter() - t0
        perf_logger.info(f"[PERF] _process_content (парсинг Markdown): {t_process:.4f} сек")
        
        # Этап 3: Сохранение DOCX
        t0 = time.perf_counter()
        print(f"💾 Сохранение файла: {docx_path}")
        if progress_callback:
            progress_callback('save', 95, 100, 'Сохранение DOCX файла...')
        self.doc.save(docx_path)
        t_save = time.perf_counter() - t0
        perf_logger.info(f"[PERF] Document.save(): {t_save:.4f} сек")
        
        t_total = time.perf_counter() - t_start
        perf_logger.info(f"[PERF] ИТОГО ВРЕМЯ КОНВЕРТАЦИИ: {t_total:.4f} сек")
        perf_logger.info(f"[PERF] Из них: load={t_load:.4f}s, process={t_process:.4f}s, save={t_save:.4f}s")
        perf_logger.info("=" * 60)
        
        print("✅ Конвертация завершена успешно!")
        if progress_callback:
            progress_callback('done', 100, 100, 'Конвертация завершена')
        
        return docx_path
    
    def _process_content(self, content, progress_callback=None):
        """Обработка содержимого Markdown
        
        Args:
            content: текст Markdown
            progress_callback: опциональный callback(stage, progress, total, message)
        """
        lines = content.split('\n')
        total_lines = len(lines)
        i = 0
        table_counter = {'current': 0}
        caption_added = False
        last_logged_percent = 0
        
        while i < total_lines:
            line = lines[i].strip()
            
            # Логирование прогресса каждые 5%
            if progress_callback and total_lines > 0:
                percent = int(i * 100 / total_lines)
                if percent >= last_logged_percent + 5:
                    last_logged_percent = percent
                    progress_callback('processing', i, total_lines, f'Обработка строки {i}/{total_lines} ({percent}%)')
            
            # Пропуск пустых строк
            if not line:
                i += 1
                continue
            
            # Список (п. 2.8) — маркированные и индентированные списки
            if re.match(r'^[\s]*[-•*]\s+', line):
                caption_added = False
                self._add_list_item(line)
                i += 1
                continue
            
            # Подпись таблицы: **Таблица X.Y.Z** или **Название**
            if re.match(r'^\*\*Таблица\s+\d+', line) or re.match(r'^\*\*Перечень\s+', line):
                self._add_caption(line)
                caption_added = True
                i += 1
                continue
            
            # Заголовок 1 уровня - РАЗДЕЛ (Таблица 2.1)
            if line.startswith('# '):
                caption_added = False
                self._add_heading(line.replace('# ', ''), level=1)
                i += 1
                continue
            
            # Заголовок 2 уровня - ГЛАВА
            if line.startswith('## '):
                caption_added = False
                self._add_heading(line.replace('## ', ''), level=2)
                i += 1
                continue
            
            # Заголовок 3 уровня - ПУНКТ
            if line.startswith('### '):
                caption_added = False
                self._add_heading(line.replace('### ', ''), level=3)
                i += 1
                continue
            
            # Заголовок 4 уровня - ПОДПУНКТ
            if line.startswith('#### '):
                caption_added = False
                self._add_heading(line.replace('#### ', ''), level=4)
                i += 1
                continue
            
            # Нумерованный список: "1. **текст**" или "1. текст"
            if re.match(r'^\d+\.\s+', line):
                caption_added = False
                self._add_list_item(line)
                i += 1
                continue
            
            # Таблица (п. 2.10)
            elif line.startswith('|') and '|' in line:
                table_lines = []
                start_i = i
                while i < len(lines) and lines[i].strip().startswith('|'):
                    table_lines.append(lines[i].strip())
                    i += 1
                
                # Проверка на разделительную строку таблицы
                if len(table_lines) > 1 and re.match(r'^\|[\s\-:|]+\|$', table_lines[1]):
                    table_counter['current'] += 1
                    self._add_table(table_lines, table_counter, self._current_section, self._current_chapter, skip_caption=caption_added)
                    caption_added = False
                else:
                    # Это не таблица, а текст с вертикальными чертами
                    # Возвращаем i на начало и обрабатываем каждую строку как обычный текст
                    i = start_i
                    while i < len(lines) and lines[i].strip().startswith('|'):
                        self._add_paragraph(lines[i].strip())
                        i += 1
                continue
            
            # Примечание (п. 2.12)
            elif line.startswith('⚠️') or line.startswith('📊') or line.startswith('✅') or line.startswith('🔍') or line.startswith('🎯'):
                self._add_note(line)
                i += 1
                continue
            
            # Примечание к таблице (п. 2.12.2)
            elif line.startswith('* ') and i > 0 and '*' in line:
                self._add_table_note(line)
                i += 1
                continue
            
            # Обычный текст (п. 2.7.2)
            else:
                caption_added = False
                self._add_paragraph(line)
                i += 1
                continue
    
    def _add_heading(self, text, level=1):
        """Добавление заголовка (Таблица 2.1) с поддержкой **bold**"""
        paragraph = self.doc.add_paragraph(style=f'Heading {level}')
        self._add_rich_text(paragraph, text)
        
        if paragraph.text.endswith('.'):
            paragraph.text = paragraph.text[:-1]
        
        if hasattr(self, 'page_manager') and self.page_manager:
            if level == 1:
                self._current_section += 1
                self._current_chapter = 0
                self.page_manager.update_footer_text(self.doc, self._current_section, self._current_chapter)
            elif level == 2:
                self._current_chapter += 1
                self.page_manager.update_footer_text(self.doc, self._current_section, self._current_chapter)
        
        return paragraph
    
    def _add_paragraph(self, text):
        """Добавление абзаца текста (п. 2.7.2) с поддержкой **bold**"""
        paragraph = self.doc.add_paragraph(style='Normal')
        self._add_rich_text(paragraph, text)
        self._disable_hyphenation(paragraph)
        return paragraph
    
    def _add_rich_text(self, paragraph, text):
        """Добавляет текст с поддержкой **bold** сегментов в параграф."""
        parts = re.split(r'(\*\*.*?\*\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                inner = self._clean_text(part[2:-2])
                run = paragraph.add_run(inner)
                run.bold = True
                run.font.name = 'Times New Roman'
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            else:
                cleaned = self._clean_text(part)
                if cleaned:
                    run = paragraph.add_run(cleaned)
                    run.font.name = 'Times New Roman'
                    rPr = run._element.get_or_add_rPr()
                    rFonts = rPr.get_or_add_rFonts()
                    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    
    def _add_caption(self, text):
        """Добавление подписи таблицы/рисунка (п. 2.10.2, 2.11.2)"""
        cleaned = self._clean_text(text)
        paragraph = self.doc.add_paragraph(cleaned, style='Caption')
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if paragraph.text.endswith('.'):
            paragraph.text = paragraph.text[:-1]
        return paragraph
    
    def _add_table(self, table_lines, table_counter, section_num, chapter_num, skip_caption=False):
        """Добавление таблицы по НТД 01-2013 (п. 2.10)"""
        data = self._parse_table_data(table_lines)
        
        if not data or len(data) < 2:
            return
        
        if not skip_caption and len(data) > 2:
            table_num = table_counter['current']
            if chapter_num > 0:
                table_title = f"Таблица {section_num}.{chapter_num}.{table_num}"
            else:
                table_title = f"Таблица {section_num}.{table_num}"
            
            caption_paragraph = self.doc.add_paragraph(table_title, style='Caption')
            caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if caption_paragraph.text.endswith('.'):
                caption_paragraph.text = caption_paragraph.text[:-1]
        
        # Определяем, есть ли строка нумерации граф после заголовка
        numbering_row_idx = None
        if len(data) > 2 and self._is_numbering_row(data[1]):
            numbering_row_idx = 1
        
        # Шапка таблицы: заголовок + опционально строка нумерации
        if numbering_row_idx is not None:
            header_data = data[:2]
        else:
            header_data = [data[0]]
        
        header_table = self.doc.add_table(rows=len(header_data), cols=len(header_data[0]))
        header_table.style = 'Table Grid'
        self._setup_table_header(header_table, header_data[0])
        if numbering_row_idx is not None:
            self._setup_numbering_row(header_table, header_data[1])
        
        # Spacer между шапкой и телом — минимальная высота
        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(0)
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.line_spacing = Pt(1)
        pPr = spacer._element.get_or_add_pPr()
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        spacing.set(qn('w:line'), '20')
        spacing.set(qn('w:lineRule'), 'exact')
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')
        run = spacer.add_run('')
        run.font.size = Pt(1)
        
        # Основная таблица
        body_start = (numbering_row_idx + 1) if numbering_row_idx is not None else 1
        main_table = self.doc.add_table(rows=len(data) - body_start, cols=len(data[0]))
        main_table.style = 'Table Grid'
        self._setup_main_table(main_table, data[body_start:], header_data[0])
        
        # Интервал после таблицы 6 пт (п. 2.10.13)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    def _is_numbering_row(self, row_data):
        """Проверяет, является ли строка строкой нумерации граф (1 | 2 | 3 | ...)"""
        for cell in row_data:
            cleaned = cell.strip().replace(' ', '')
            if not re.match(r'^\d+$', cleaned):
                return False
        return True
    
    def _setup_numbering_row(self, table, row_data):
        """Настройка строки нумерации граф (п. 2.10.3)"""
        row = table.rows[1]
        table_font_cfg = self.config.get('styles', {}).get('Table Grid', {}).get('font', {})
        font_name = table_font_cfg.get('name', 'Times New Roman')
        font_size = table_font_cfg.get('size', 12)
        
        for i, cell_text in enumerate(row_data):
            if i >= len(row.cells):
                break
            cell = row.cells[i]
            cell.text = cell_text.strip()
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.paragraph_format.first_line_indent = Cm(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    rPr = run._element.get_or_add_rPr()
                    rFonts = rPr.get_or_add_rFonts()
                    rFonts.set(qn('w:eastAsia'), font_name)
            cell.margin_left = Cm(0)
            cell.margin_right = Cm(0)
            cell.margin_top = Cm(0)
            cell.margin_bottom = Cm(0)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        self._set_table_row_height(table, 1, 'minimum')
    
    def _parse_table_data(self, table_lines):
        """Парсинг данных таблицы из Markdown"""
        data = []
        
        for line in table_lines:
            # Пропуск разделительной строки
            if re.match(r'^\|[\s\-:|]+\|$', line):
                continue
            
            # Извлечение ячеек
            cells = [cell.strip() for cell in line.split('|')]
            cells = [cell for cell in cells if cell]  # Удаление пустых
            
            if cells:
                data.append(cells)
        
        return data
    
    def _setup_table_header(self, table, headers):
        """Настройка шапки таблицы (п. 2.10.3, 2.10.4)"""
        header_row = table.rows[0]
        table_font_cfg = self.config.get('styles', {}).get('Table Grid', {}).get('font', {})
        font_name = table_font_cfg.get('name', 'Times New Roman')
        font_size = table_font_cfg.get('size', 12)
        
        for i, header_text in enumerate(headers):
            header_text = self._clean_text(header_text)
            cell = header_row.cells[i]
            cell.text = header_text
            
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.paragraph_format.first_line_indent = Cm(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                
                if header_text:
                    paragraph.text = header_text.capitalize()
                
                if paragraph.text.endswith('.'):
                    paragraph.text = paragraph.text[:-1]
                
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    rPr = run._element.get_or_add_rPr()
                    rFonts = rPr.get_or_add_rFonts()
                    rFonts.set(qn('w:eastAsia'), font_name)
            
            # Обнуление margin ячеек
            cell.margin_left = Cm(0)
            cell.margin_right = Cm(0)
            cell.margin_top = Cm(0)
            cell.margin_bottom = Cm(0)
            
            # Выравнивание по центру сверху (п. 2.10.7)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        # Настройка высоты строки 0 см, режим минимум (п. 2.10.8)
        self._set_table_row_height(table, 0, 'minimum')
        
        # Отключение нижней границы шапки (п. 2.10.3)
        self._remove_bottom_border(header_row)
        
        # Нумерация граф (п. 2.10.3)
        self._add_column_numbering(table)
    
    def _setup_main_table(self, table, data, headers):
        """Настройка основной таблицы (п. 2.10)"""
        table_font_cfg = self.config.get('styles', {}).get('Table Grid', {}).get('font', {})
        font_name = table_font_cfg.get('name', 'Times New Roman')
        font_size = table_font_cfg.get('size', 12)
        
        for row_idx, row_data in enumerate(data):
            row = table.rows[row_idx]
            
            for col_idx, cell_text in enumerate(row_data):
                if col_idx < len(row.cells):
                    cell_text = self._clean_text(cell_text)
                    cell = row.cells[col_idx]
                    cell.text = cell_text
                    
                    for paragraph in cell.paragraphs:
                        paragraph.paragraph_format.line_spacing = 1.0
                        paragraph.paragraph_format.first_line_indent = Cm(0)
                        paragraph.paragraph_format.space_before = Pt(0)
                        paragraph.paragraph_format.space_after = Pt(0)
                        
                        if col_idx == 0:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            if cell_text:
                                paragraph.text = cell_text[0].upper() + cell_text[1:] if cell_text else cell_text
                            if paragraph.text.endswith('.'):
                                paragraph.text = paragraph.text[:-1]
                        elif self._is_numeric(cell_text):
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        
                        for run in paragraph.runs:
                            run.font.name = font_name
                            run.font.size = Pt(font_size)
                            rPr = run._element.get_or_add_rPr()
                            rFonts = rPr.get_or_add_rFonts()
                            rFonts.set(qn('w:eastAsia'), font_name)
                    
                    # Обнуление margin ячеек
                    cell.margin_left = Cm(0)
                    cell.margin_right = Cm(0)
                    cell.margin_top = Cm(0)
                    cell.margin_bottom = Cm(0)
                    
                    # Выравнивание сверху (п. 2.10.7)
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            
            # Настройка высоты строки (п. 2.10.8)
            self._set_table_row_height(table, row_idx, 'minimum')
        
        # Автоподбор по ширине окна (п. 2.10.9)
        table.autofit = True
        
        # Повтор заголовков на следующих страницах (п. 2.10.5)
        self._repeat_header_on_pages(table)
    
    def _set_table_row_height(self, table, row_idx, mode='minimum'):
        """Установка высоты строки таблицы (п. 2.10.8)"""
        row = table.rows[row_idx]
        tr = row._tr
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), '0')
        trHeight.set(qn('w:hRule'), mode)
        tr.append(trHeight)
    
    def _remove_bottom_border(self, row):
        """Отключение нижней границы шапки (п. 2.10.3)"""
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = tcPr.first_child_found_in('w:tcBorders')
            if tcBorders is None:
                tcBorders = OxmlElement('w:tcBorders')
                tcPr.append(tcBorders)
            
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'nil')
            tcBorders.append(bottom)
    
    def _add_column_numbering(self, table):
        """Добавление нумерации граф (п. 2.10.3)"""
        header_row = table.rows[0]
        for i, cell in enumerate(header_row.cells):
            # Добавляем номер графы в первую ячейку если нужно
            pass
    
    def _repeat_header_on_pages(self, table):
        """Повтор заголовков на следующих страницах (п. 2.10.5)"""
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblHeader = OxmlElement('w:tblHeader')
        tblPr.append(tblHeader)
    
    def _is_numeric(self, text):
        """Проверка является ли текст числовым значением"""
        if not text:
            return False
        # Удаление пробелов, знаков +, -, %
        clean_text = text.replace(' ', '').replace('+', '').replace('−', '-')
        try:
            float(clean_text.replace(',', '.'))
            return True
        except ValueError:
            return False
    
    def _add_table_note(self, text):
        """Добавление примечания к таблице (п. 2.10.14, 2.12)"""
        text = self._clean_text(text)
        paragraph = self.doc.add_paragraph(text, style='Table Note')
        
        # Интервал перед примечанием 0 пт (п. 2.10.14, 2.12.4)
        paragraph.paragraph_format.space_before = Pt(0)
        
        # Без абзаца (п. 2.12.2)
        paragraph.paragraph_format.first_line_indent = Cm(0)
        
        # С прописной буквы (п. 2.12.2)
        if paragraph.text:
            paragraph.text = paragraph.text[0].upper() + paragraph.text[1:]
        
        return paragraph
    
    def _add_note(self, text):
        """Добавление примечания/комментария (п. 2.12)"""
        clean_text = self._clean_text(text.strip())
        
        note_config = self.config.get('styles', {}).get('Note', {}).get('paragraph', {})
        paragraph = self.doc.add_paragraph(clean_text, style='Note')
        paragraph.paragraph_format.first_line_indent = Cm(note_config.get('first_line_indent_cm', 0))
        paragraph.paragraph_format.line_spacing = note_config.get('line_spacing', 1.0)
        
        return paragraph
    
    def _add_list_item(self, text):
        """Добавление элемента списка (п. 2.8) с поддержкой **bold** и номеров"""
        paragraph = self.doc.add_paragraph(style='List Paragraph')
        
        # Определяем уровень вложенности по количеству ведущих пробелов
        stripped = text.lstrip()
        leading_spaces = len(text) - len(stripped)
        
        if leading_spaces >= 6:
            indent_level = 3
        elif leading_spaces >= 3:
            indent_level = 2
        else:
            indent_level = 1
        
        list_config = self.config.get('styles', {}).get('List Paragraph', {}).get('paragraph', {})
        indent_map = {
            1: list_config.get('left_indent_level_1_cm', 0.75),
            2: list_config.get('left_indent_level_2_cm', 1.00),
            3: list_config.get('left_indent_level_3_cm', 1.25),
        }
        paragraph.paragraph_format.left_indent = Cm(indent_map.get(indent_level, 1.25))
        
        first_indent = list_config.get('first_line_indent_cm', 0)
        paragraph.paragraph_format.first_line_indent = Cm(first_indent)
        
        # Добавляем текст с поддержкой bold
        self._add_rich_text(paragraph, stripped)
        
        return paragraph
    
    def _get_list_indent_level(self, text):
        """Определение уровня вложенности списка"""
        if text.startswith('    '):
            return 2
        elif text.startswith('        '):
            return 3
        return 1
    
    def _disable_hyphenation(self, paragraph):
        """Запрет переноса слов (п. 2.7.3)"""
        pPr = paragraph._element.get_or_add_pPr()
        hyphen = OxmlElement('w:hyphen')
        hyphen.set(qn('w:val'), '0')
        pPr.append(hyphen)

    def _clean_text(self, text):
        """Удаление markdown-артефактов: двойных звёздочек ** и эмодзи"""
        if not isinstance(text, str):
            return text
        # Удаляем **, но не трогаем одиночные *
        cleaned = re.sub(r'\*\*', '', text)
        # Удаляем эмодзи
        cleaned = self._emoji_pattern.sub('', cleaned)
        return cleaned


def main():
    """Основная функция"""
    import sys
    
    # Пути к файлам
    if len(sys.argv) > 1:
        md_path = sys.argv[1]
    else:
        md_path = '(ГЕНЕРАЦИЯ) Том II Черкесск Демография.md'
    
    if len(sys.argv) > 2:
        docx_path = sys.argv[2]
    else:
        docx_path = '(ГЕНЕРАЦИЯ) Том II Черкесск Демография.docx'
    
    # Конфигурация
    config_path = 'configs/active/config.yaml'
    
    # Конвертация
    converter = RI2013Converter(config_path)
    converter.convert(md_path, docx_path)
    
    print(f"\n📋 Отчёт о конвертации:")
    print(f"   Исходный файл: {md_path}")
    print(f"   Результат: {docx_path}")
    print(f"   Стандарт: НТД 01-2013 Редакция 4")
    print(f"   Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == '__main__':
    main()

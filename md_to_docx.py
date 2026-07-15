#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт преобразования Markdown в DOCX по НТД 01-2013
проектная организация - Проектно-производственный департамент
Рабочая Инструкция НТД 01-2013 "Порядок оформления проектной градостроительной документации"
Редакция 4
"""

import re
import os
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
logger = logging.getLogger(__name__)


class RI2013Converter:
    """Конвертер Markdown в DOCX по НТД 01-2013"""

    def __init__(self, config_path='configs/active/config.yaml', template_path=None):
        """Инициализация конвертера с загрузкой конфигурации

        Args:
            config_path: путь к YAML-конфигу
            template_path: путь к DOCX-шаблону для копирования стилей (опционально)
        """
        perf_logger.info("--- ИНИЦИАЛИЗАЦИЯ RI2013Converter ---")

        self.template_path = template_path

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
            "\U0001F600-\U0001F64F"
            "\U0001F300-\U0001F5FF"
            "\U0001F680-\U0001F6FF"
            "\U0001F1E0-\U0001F1FF"
            "\U00002702-\U000027B0"
            "\U000024C2-\U0001F251"
            "]+",
            flags=re.UNICODE
        )

        # Этап: _setup_document() — PageManager + StyleManager + копирование стилей из шаблона
        t0 = time.perf_counter()
        self._setup_document()
        t_setup = time.perf_counter() - t0
        perf_logger.info(f"[PERF] _setup_document (PageManager + StyleManager): {t_setup:.4f} сек")

        self._current_section = 0
        self._current_chapter = 0

        # Инициализация маппинга типов контента -> имена стилей из конфига
        self._init_style_map()

        t_total_init = t_config + t_doc_create + t_setup
        perf_logger.info(f"[PERF] ИТОГО __init__: {t_total_init:.4f} сек")
        perf_logger.info("--- ИНИЦИАЛИЗАЦИЯ ЗАВЕРШЕНА ---")
        
    def _init_style_map(self):
        """Инициализирует маппинг типов контента -> имена стилей из конфига.
        
        Читает секцию content_type_styles из конфига. Если секция отсутствует,
        строит маппинг по умолчанию из detection_rules и стандартных имён.
        """
        # Дефолтный маппинг на случай отсутствия секции в конфиге
        self._style_map = {
            'heading_1': 'Heading 1',
            'heading_2': 'Heading 2',
            'heading_3': 'Heading 3',
            'heading_4': 'Heading 4',
            'paragraph': 'Normal',
            'caption': 'Caption',
            'table_name': 'Caption',
            'table_note': 'Table Note',
            'table_grid': 'Table Grid',
            'list': 'List Paragraph',
            'note': 'Note',
        }
        
        # Если есть секция content_type_styles — используем её
        content_type_styles = self.config.get('content_type_styles', {})
        if content_type_styles:
            self._style_map.update(content_type_styles)
        
        # Если есть detection_rules — дополним маппинг заголовков
        detection_rules = self.config.get('detection_rules', {})
        if detection_rules:
            # Ищем стили для заголовков по паттернам
            for style_name, patterns in detection_rules.items():
                if not patterns:
                    continue
                for pattern in patterns:
                    if re.match(r'^\^?РАЗДЕЛ', pattern):
                        self._style_map.setdefault('heading_1', style_name)
                    elif re.match(r'^\^?ГЛАВА', pattern):
                        self._style_map.setdefault('heading_2', style_name)
                    elif re.match(r'^\^?\d+\\\.\d+', pattern) or re.match(r'^\^?\d+\.\d+', pattern):
                        self._style_map.setdefault('heading_3', style_name)
        
        logger.debug(f"Style map: {self._style_map}")

    def _resolve_style(self, content_type: str, fallback: str = None) -> str:
        """Резолвит имя стиля по типу контента из конфига.
        
        Args:
            content_type: тип контента ('heading_1', 'paragraph', 'caption', и т.д.)
            fallback: имя стиля по умолчанию, если тип не найден в маппинге
        
        Returns:
            Имя стиля для использования в doc.add_paragraph(style=...)
        """
        style_name = self._style_map.get(content_type, fallback or 'Normal')
        
        # Проверяем существование стиля в документе
        if style_name in self.doc.styles:
            return style_name
        
        # Fallback: стандартные имена Word
        fallback_map = {
            'heading_1': 'Heading 1',
            'heading_2': 'Heading 2',
            'heading_3': 'Heading 3',
            'heading_4': 'Heading 4',
            'paragraph': 'Normal',
            'caption': 'Normal',
            'table_name': 'Normal',
            'table_note': 'Normal',
            'table_grid': 'Table Grid',
            'list': 'List Paragraph',
            'note': 'Normal',
        }
        fallback_name = fallback_map.get(content_type, 'Normal')
        if fallback_name in self.doc.styles:
            logger.warning(f"Стиль '{style_name}' для '{content_type}' не найден, используем '{fallback_name}'")
            return fallback_name
        
        return 'Normal'

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
            'version': '4.3',
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
                    'font': {'name': 'Calibri Light', 'size': 14, 'bold': True, 'italic': False},
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
            },
            'content_type_styles': {
                'heading_1': 'Heading 1',
                'heading_2': 'Heading 2',
                'heading_3': 'Heading 3',
                'heading_4': 'Heading 4',
                'paragraph': 'Normal',
                'caption': 'Caption',
                'table_name': 'Caption',
                'table_note': 'Table Note',
                'table_grid': 'Table Grid',
                'list': 'List Paragraph',
                'note': 'Note',
            }
        }
    
    def _setup_document(self):
        """Настройка документа согласно конфигу"""
        self.page_manager = PageManager(self.config)
        self.page_manager.apply_all(self.doc)

        if self.style_manager:
            self.style_manager.setup_styles(self.doc)
        else:
            from src.core.style_manager import StyleManager
            from src.core.config_loader import ConfigLoader
            loader = ConfigLoader('')
            loader.styles = self.config.get('styles', {})
            sm = StyleManager(loader)
            sm.setup_styles(self.doc)

        # Обеспечиваем наличие всех стилей, указанных в content_type_styles
        self._ensure_all_config_styles()

        # Копирование стилей из DOCX-шаблона (если указан)
        if self.template_path and os.path.exists(self.template_path):
            self._copy_styles_from_template()

    def _ensure_all_config_styles(self):
        """Создаёт в документе все стили, перечисленные в content_type_styles конфига."""
        content_type_styles = self.config.get('content_type_styles', {})
        for content_type, style_name in content_type_styles.items():
            if style_name not in self.doc.styles:
                # Стиль не существует — создаём с настройками из секции styles конфига
                style_config = self.config.get('styles', {}).get(style_name)
                if style_config and style_config.get('enabled', True):
                    self.style_manager.ensure_style_exists(self.doc, style_name, style_config)
                else:
                    # Нет конфига — создаём пустой стиль
                    try:
                        self.doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                    except Exception:
                        pass  # Стиль уже существует или не может быть создан

    def _copy_styles_from_template(self):
        """Копирует стили из DOCX-шаблона в текущий документ"""
        perf_logger.info(f"Копирование стилей из шаблона: {self.template_path}")
        template_doc = Document(self.template_path)

        copied_count = 0
        for template_style in template_doc.styles:
            if template_style.type is None:
                continue
            try:
                style_type_name = template_style.type.name
            except AttributeError:
                continue

            if style_type_name != 'PARAGRAPH':
                continue

            style_name = template_style.name

            # Пропускаем стандартные стили
            if style_name in ('Default Paragraph Font',):
                continue

            # Создаём или обновляем стиль
            try:
                if style_name in self.doc.styles:
                    target_style = self.doc.styles[style_name]
                else:
                    target_style = self.doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

                # Копируем свойства шрифта
                if template_style.font:
                    if template_style.font.name:
                        target_style.font.name = template_style.font.name
                    if template_style.font.size:
                        target_style.font.size = template_style.font.size
                    if template_style.font.bold is not None:
                        target_style.font.bold = template_style.font.bold
                    if template_style.font.italic is not None:
                        target_style.font.italic = template_style.font.italic

                # Копируем свойства абзаца
                if template_style.paragraph_format:
                    pf = template_style.paragraph_format
                    if pf.alignment is not None:
                        target_style.paragraph_format.alignment = pf.alignment
                    if pf.first_line_indent is not None:
                        target_style.paragraph_format.first_line_indent = pf.first_line_indent
                    if pf.left_indent is not None:
                        target_style.paragraph_format.left_indent = pf.left_indent
                    if pf.space_before is not None:
                        target_style.paragraph_format.space_before = pf.space_before
                    if pf.space_after is not None:
                        target_style.paragraph_format.space_after = pf.space_after
                    if pf.line_spacing is not None:
                        target_style.paragraph_format.line_spacing = pf.line_spacing
                    if pf.page_break_before:
                        target_style.paragraph_format.page_break_before = True

                copied_count += 1
                logger.debug(f"Скопирован стиль: {style_name}")

            except Exception as e:
                logger.warning(f"Не удалось скопировать стиль '{style_name}': {e}")

        perf_logger.info(f"Скопировано стилей из шаблона: {copied_count}")
    
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
        last_caption_text = ''
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

            # Автоопределение заголовков по паттернам (если нет markdown-разметки)
            # РАЗДЕЛ X.
            if re.match(r'^РАЗДЕЛ\s+\d+', line):
                caption_added = False
                self._add_heading(line, level=1)
                i += 1
                continue

            # ГЛАВА X.
            if re.match(r'^ГЛАВА\s+\d+', line):
                caption_added = False
                self._add_heading(line, level=2)
                i += 1
                continue

            # X.Y Подзаголовок (например "3.1 Анализ...")
            if re.match(r'^\d+\.\d+\s+[А-Я]', line) and len(line) < 100:
                caption_added = False
                self._add_heading(line, level=3)
                i += 1
                continue

            # Подпись таблицы: **Таблица X.Y.Z** или Таблица X.Y.Z или **Название**
            if re.match(r'^(\*\*)?Таблица\s+\d+', line) or re.match(r'^\*\*Перечень\s+', line):
                # Извлекаем номер таблицы из подписи (если есть)
                num_match = re.search(r'Таблица\s+(\d+[\.\d]*)', line)
                if num_match:
                    existing_num = num_match.group(1)
                    parts = existing_num.split('.')
                    if len(parts) >= 3:
                        self._current_section = int(parts[0])
                        self._current_chapter = int(parts[1])
                        table_counter['current'] = int(parts[2])
                    elif len(parts) == 2:
                        self._current_section = int(parts[0])
                        table_counter['current'] = int(parts[1])
                    else:
                        table_counter['current'] = int(parts[0])
                self._add_caption(line)
                caption_added = True
                last_caption_text = line.strip()
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
                    # Проверяем, была ли подписана таблица (ищем "Таблица X.Y.Z" перед таблицей)
                    has_caption = False
                    for j in range(max(0, start_i - 5), start_i):
                        if re.match(r'^(\*\*)?Таблица\s+\d+', lines[j].strip()):
                            has_caption = True
                            break
                    self._add_table(table_lines, table_counter, self._current_section, self._current_chapter, skip_caption=has_caption)
                    caption_added = False
                    last_caption_text = ''
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

            # Примечание к таблице: "Примечание — ..."
            elif line.startswith('Примечание'):
                self._add_table_note(line)
                i += 1
                continue
            
            # Обычный текст (п. 2.7.2)
            else:
                # Если предыдущая непустая строка была подписью таблицы — это заголовок таблицы
                if caption_added and last_caption_text:
                    # Заголовок таблицы — стиль определяет выравнивание из конфига
                    table_name_style = self._resolve_style('table_name')
                    paragraph = self.doc.add_paragraph(style=table_name_style)
                    self._add_rich_text(paragraph, line)
                    paragraph.paragraph_format.first_line_indent = Cm(0)
                    caption_added = False
                    last_caption_text = ''
                else:
                    self._add_paragraph(line)
                i += 1
                continue
    
    def _add_heading(self, text, level=1):
        """Добавление заголовка с поддержкой **bold**"""
        # Резолвим стиль из конфига по уровню заголовка
        style_name = self._resolve_style(f'heading_{level}')

        paragraph = self.doc.add_paragraph(style=style_name)

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
        style_name = self._resolve_style('paragraph')
        paragraph = self.doc.add_paragraph(style=style_name)
        self._add_rich_text(paragraph, text)
        self._disable_hyphenation(paragraph)
        return paragraph
    
    def _add_rich_text(self, paragraph, text):
        """Добавляет текст с поддержкой **bold** сегментов в параграф."""
        # Определяем шрифт из конфига текущего стиля
        style_name = paragraph.style.name
        style_cfg = self.config.get('styles', {}).get(style_name, {})
        font_cfg = style_cfg.get('font', {})
        font_name = font_cfg.get('name', 'Times New Roman')
        east_asia = font_cfg.get('east_asia', font_name)

        parts = re.split(r'(\*\*.*?\*\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                inner = self._clean_text(part[2:-2])
                run = paragraph.add_run(inner)
                run.bold = True
                run.font.name = font_name
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(qn('w:eastAsia'), east_asia)
            else:
                cleaned = self._clean_text(part)
                if cleaned:
                    run = paragraph.add_run(cleaned)
                    run.font.name = font_name
                    rPr = run._element.get_or_add_rPr()
                    rFonts = rPr.get_or_add_rFonts()
                    rFonts.set(qn('w:eastAsia'), east_asia)
    
    def _add_caption(self, text):
        """Добавление подписи таблицы/рисунка (п. 2.10.2, 2.11.2)"""
        cleaned = self._clean_text(text)
        style_name = self._resolve_style('caption')
        paragraph = self.doc.add_paragraph(cleaned, style=style_name)
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
            
            caption_style = self._resolve_style('caption')
            caption_paragraph = self.doc.add_paragraph(table_title, style=caption_style)
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
        
        table_grid_style = self._resolve_style('table_grid')
        header_table = self.doc.add_table(rows=len(header_data), cols=len(header_data[0]))
        header_table.style = table_grid_style
        self._setup_table_header(header_table, header_data[0])
        if numbering_row_idx is not None:
            self._setup_numbering_row(header_table, header_data[1])
        
        # Настройка высоты строки 0 см, режим минимум (п. 2.10.8)
        self._set_table_row_height(header_table, 0, 'minimum')
        if numbering_row_idx is not None:
            self._set_table_row_height(header_table, 1, 'minimum')
        
        # Отключение нижней границы шапки (п. 2.10.3) — с последней строки
        last_header_row = header_table.rows[len(header_data) - 1]
        self._remove_bottom_border(last_header_row)

        # Вычисление и установка одинаковых ширин колонок для шапки и тела
        self._sync_column_widths(header_table, data)

        # Основная таблица (без spacer — шапка и тело идут подряд)
        body_start = (numbering_row_idx + 1) if numbering_row_idx is not None else 1
        main_table = self.doc.add_table(rows=len(data) - body_start, cols=len(data[0]))
        main_table.style = table_grid_style
        self._setup_main_table(main_table, data[body_start:], header_data[0])

        # Установка одинаковых ширин колонок для тела таблицы
        self._sync_column_widths(main_table, data)

        # Интервал после таблицы из конфига (п. 2.10.13)
        spacing_after = self.config.get('formatting_rules', {}).get('tables', {}).get('spacing_after_table_pt', 6)
        if main_table.rows:
            last_row = main_table.rows[-1]
            for cell in last_row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.space_after = Pt(spacing_after)
    
    def _is_numbering_row(self, row_data):
        """Проверяет, является ли строка строкой нумерации граф (1 | 2 | 3 | ...)"""
        for cell in row_data:
            cleaned = cell.strip().replace(' ', '')
            if not re.match(r'^\d+$', cleaned):
                return False
        return True
    
    def _get_table_cell_format(self):
        """Получает настройки ячеек таблицы из конфига (шрифт, размер, интервалы)."""
        table_grid_style = self._resolve_style('table_grid')
        style_cfg = self.config.get('styles', {}).get(table_grid_style, {})
        font_cfg = style_cfg.get('font', {})
        para_cfg = style_cfg.get('paragraph', {})
        return {
            'font_name': font_cfg.get('name', 'Times New Roman'),
            'font_size': font_cfg.get('size', 12),
            'east_asia': font_cfg.get('east_asia', font_cfg.get('name', 'Times New Roman')),
            'line_spacing': para_cfg.get('line_spacing', 1.0),
            'first_line_indent_cm': para_cfg.get('first_line_indent_cm', 0),
            'space_before_pt': para_cfg.get('space_before_pt', 0),
            'space_after_pt': para_cfg.get('space_after_pt', 0),
        }

    def _apply_cell_format(self, paragraph, cell_fmt, bold=False, alignment=None):
        """Применяет форматирование ячейки таблицы из конфига."""
        if alignment is not None:
            paragraph.alignment = alignment
        paragraph.paragraph_format.line_spacing = cell_fmt['line_spacing']
        paragraph.paragraph_format.first_line_indent = Cm(cell_fmt['first_line_indent_cm'])
        paragraph.paragraph_format.space_before = Pt(cell_fmt['space_before_pt'])
        paragraph.paragraph_format.space_after = Pt(cell_fmt['space_after_pt'])
        for run in paragraph.runs:
            run.font.name = cell_fmt['font_name']
            run.font.size = Pt(cell_fmt['font_size'])
            if bold:
                run.bold = True
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn('w:eastAsia'), cell_fmt['east_asia'])

    def _setup_numbering_row(self, table, row_data):
        """Настройка строки нумерации граф (п. 2.10.3)"""
        row = table.rows[1]
        cell_fmt = self._get_table_cell_format()

        for i, cell_text in enumerate(row_data):
            if i >= len(row.cells):
                break
            cell = row.cells[i]
            cell.text = cell_text.strip()
            for paragraph in cell.paragraphs:
                self._apply_cell_format(paragraph, cell_fmt, alignment=WD_ALIGN_PARAGRAPH.CENTER)
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
        cell_fmt = self._get_table_cell_format()

        for i, header_text in enumerate(headers):
            header_text = self._clean_text(header_text)
            cell = header_row.cells[i]
            cell.text = header_text

            for paragraph in cell.paragraphs:
                if header_text:
                    paragraph.text = header_text

                if paragraph.text.endswith('.'):
                    paragraph.text = paragraph.text[:-1]

                self._apply_cell_format(paragraph, cell_fmt, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)

            # Обнуление margin ячеек
            cell.margin_left = Cm(0)
            cell.margin_right = Cm(0)
            cell.margin_top = Cm(0)
            cell.margin_bottom = Cm(0)

            # Выравнивание по центру сверху (п. 2.10.7)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def _setup_main_table(self, table, data, headers):
        """Настройка основной таблицы (п. 2.10)"""
        cell_fmt = self._get_table_cell_format()

        for row_idx, row_data in enumerate(data):
            row = table.rows[row_idx]

            for col_idx, cell_text in enumerate(row_data):
                if col_idx < len(row.cells):
                    cell_text = self._clean_text(cell_text)
                    cell = row.cells[col_idx]
                    cell.text = cell_text

                    for paragraph in cell.paragraphs:
                        if col_idx == 0:
                            alignment = WD_ALIGN_PARAGRAPH.LEFT
                            if cell_text:
                                paragraph.text = cell_text[0].upper() + cell_text[1:] if cell_text else cell_text
                            if paragraph.text.endswith('.'):
                                paragraph.text = paragraph.text[:-1]
                        elif self._is_numeric(cell_text):
                            alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            alignment = WD_ALIGN_PARAGRAPH.LEFT

                        self._apply_cell_format(paragraph, cell_fmt, alignment=alignment)

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

    def _sync_column_widths(self, table, all_data):
        """Синхронизация ширины колонок между шапкой и телом таблицы.

        Вычисляет ширину колонок на основе самой длинной строки во всех строках
        (шапка + тело) и применяет одинаковые фиксированные ширины к таблице.
        """
        section = self.doc.sections[0]
        usable_width_cm = (section.page_width - section.left_margin - section.right_margin) / Cm(1)

        num_cols = len(all_data[0])

        # Вычисление пропорций на основе длины текста
        # Ограничиваем максимальную длину строки для расчёта ширины
        MAX_CHARS_FOR_WIDTH = 25  # максимальная длина строки для расчёта пропорций
        col_widths_chars = [0] * num_cols
        for row in all_data:
            for i, cell in enumerate(row):
                if i < num_cols:
                    # Ограничиваем длину для расчёта ширины
                    char_count = min(len(cell), MAX_CHARS_FOR_WIDTH)
                    col_widths_chars[i] = max(col_widths_chars[i], char_count)

        total_chars = sum(col_widths_chars)
        if total_chars == 0:
            total_chars = num_cols

        col_widths_cm = []
        for i in range(num_cols):
            w = (col_widths_chars[i] / total_chars) * usable_width_cm
            col_widths_cm.append(max(w, 0.8))  # минимум 0.8 см на колонку

        # Пересчитаем суммарно, чтобы точно занять всю ширину
        total_cm = sum(col_widths_cm)
        scale = usable_width_cm / total_cm
        col_widths_cm = [w * scale for w in col_widths_cm]

        # Отключение autofit и установка фиксированных ширин
        table.autofit = False
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                if i < num_cols:
                    cell.width = Cm(col_widths_cm[i])

        # Установка ширины через tblGrid для надёжности
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblGrid = tblPr.find(qn('w:tblGrid'))
        if tblGrid is None:
            tblGrid = OxmlElement('w:tblGrid')
            tblPr.append(tblGrid)

        # Удаляем существующие gridCol
        for gc in tblGrid.findall(qn('w:gridCol')):
            tblGrid.remove(gc)

        for w in col_widths_cm:
            gridCol = OxmlElement('w:gridCol')
            gridCol.set(qn('w:w'), str(int(Cm(w) / 635)))  # EMU в twips: 1 cm ≈ 567 twips
            tblGrid.append(gridCol)

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
        # С прописной буквы (п. 2.12.2)
        if text:
            text = text[0].upper() + text[1:]

        style_name = self._resolve_style('table_note')
        paragraph = self.doc.add_paragraph(style=style_name)

        # Добавляем run с нужным размером шрифта из конфига
        font_cfg = self.config.get('styles', {}).get(style_name, {}).get('font', {})
        font_size = font_cfg.get('size', 10)
        run = paragraph.add_run(text)
        run.font.size = Pt(font_size)

        return paragraph
    
    def _add_note(self, text):
        """Добавление примечания/комментария (п. 2.12)"""
        clean_text = self._clean_text(text.strip())
        
        style_name = self._resolve_style('note')
        paragraph = self.doc.add_paragraph(clean_text, style=style_name)
        
        # Применяем настройки абзаца из конфига
        note_config = self.config.get('styles', {}).get(style_name, {}).get('paragraph', {})
        if note_config:
            if 'first_line_indent_cm' in note_config:
                paragraph.paragraph_format.first_line_indent = Cm(note_config['first_line_indent_cm'])
            if 'line_spacing' in note_config:
                paragraph.paragraph_format.line_spacing = note_config['line_spacing']
        
        return paragraph
    
    def _add_list_item(self, text):
        """Добавление элемента списка (п. 2.8) с поддержкой **bold** и номеров"""
        style_name = self._resolve_style('list')
        paragraph = self.doc.add_paragraph(style=style_name)

        stripped = text.lstrip()
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

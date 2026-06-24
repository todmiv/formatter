#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль управления настройками страницы и колонтитулов.
Обеспечивает единообразное применение полей, ориентации, колонтитулов
в соответствии с НТД 01-2013.
"""

from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import logging
import time

logger = logging.getLogger(__name__)
perf_logger = logging.getLogger('perf')


class PageManager:
    """
    Менеджер настройки страницы и колонтитулов.
    """

    def __init__(self, config_loader):
        """
        :param config_loader: экземпляр ConfigLoader или словарь конфигурации
        """
        self.config = config_loader

    def _get_page_setup(self):
        """Возвращает конфигурацию page_setup из config_loader или словаря."""
        if hasattr(self.config, 'get_page_setup'):
            return self.config.get_page_setup()
        # Предполагаем, что это словарь
        return self.config.get('page_setup')

    def _get_headers_footers_config(self):
        """Возвращает конфигурацию headers_footers из config_loader или словаря."""
        if hasattr(self.config, 'get_headers_footers_config'):
            return self.config.get_headers_footers_config()
        # Предполагаем, что это словарь
        return self.config.get('headers_footers')

    def apply_page_setup(self, doc, orientation=None):
        """
        Применяет настройки страницы (поля, размер бумаги) к документу.

        :param doc: объект docx.Document
        :param orientation: 'portrait' или 'landscape'. Если None, берётся из конфига.
        """
        page_setup = self._get_page_setup()
        if not page_setup:
            logger.warning("Конфигурация page_setup отсутствует. Используются значения по умолчанию.")
            return

        # Определяем ориентацию
        if orientation is None:
            orientation = page_setup.get('orientation_default', 'portrait')

        # Выбираем поля
        margins = page_setup.get(orientation, {})
        if not margins:
            logger.warning(f"Поля для ориентации '{orientation}' не найдены. Используются поля по умолчанию.")
            margins = page_setup.get('portrait', {})

        # Применяем к первой секции (предполагаем, что документ имеет одну секцию)
        section = doc.sections[0]

        if 'left_margin_cm' in margins:
            section.left_margin = Cm(margins['left_margin_cm'])
        if 'right_margin_cm' in margins:
            section.right_margin = Cm(margins['right_margin_cm'])
        if 'top_margin_cm' in margins:
            section.top_margin = Cm(margins['top_margin_cm'])
        if 'bottom_margin_cm' in margins:
            section.bottom_margin = Cm(margins['bottom_margin_cm'])

        # Размер бумаги (A4 по умолчанию)
        paper_size = page_setup.get('paper_size', 'A4')
        if paper_size.upper() == 'A4':
            section.page_width = Cm(21.0)
            section.page_height = Cm(29.7)
        else:
            # Можно добавить другие форматы при необходимости
            pass

        logger.debug(f"Применены поля страницы: {orientation}, левое {margins.get('left_margin_cm')} см")

    def apply_headers_footers(self, doc):
        """
        Применяет настройки колонтитулов к документу.

        :param doc: объект docx.Document
        """
        t0 = time.perf_counter()
        headers_footers = self._get_headers_footers_config()
        if not headers_footers:
            logger.warning("Конфигурация headers_footers отсутствует. Колонтитулы не настроены.")
            return

        section = doc.sections[0]

        # Верхний колонтитул
        header_cfg = headers_footers.get('header', {})
        if header_cfg.get('enabled', False):
            t1 = time.perf_counter()
            self._setup_header(section, header_cfg)
            t_header = time.perf_counter() - t1
            perf_logger.info(f"[PERF] PageManager._setup_header: {t_header:.4f} сек")

        # Нижний колонтитул
        footer_cfg = headers_footers.get('footer', {})
        if footer_cfg.get('enabled', False):
            t1 = time.perf_counter()
            self._setup_footer(section, footer_cfg)
            t_footer = time.perf_counter() - t1
            perf_logger.info(f"[PERF] PageManager._setup_footer: {t_footer:.4f} сек")

        # Расстояние от края
        if 'position_from_edge_cm' in header_cfg:
            section.header_distance = Cm(header_cfg['position_from_edge_cm'])
        if 'position_from_edge_cm' in footer_cfg:
            section.footer_distance = Cm(footer_cfg['position_from_edge_cm'])

        t_total = time.perf_counter() - t0
        perf_logger.info(f"[PERF] PageManager.apply_headers_footers (всего): {t_total:.4f} сек")

    def _setup_header(self, section, config):
        """
        Настраивает верхний колонтитул.
        """
        header = section.header
        # Очищаем существующие параграфы
        for paragraph in header.paragraphs:
            paragraph.clear()

        content_type = config.get('content', 'page_number')
        alignment = config.get('alignment', 'right')

        if content_type == 'page_number':
            # Добавляем номер страницы как поле PAGE
            paragraph = header.add_paragraph()
            run = paragraph.add_run()
            fldChar = OxmlElement('w:fldChar')
            fldChar.set(qn('w:fldCharType'), 'begin')
            run._r.append(fldChar)

            instrText = OxmlElement('w:instrText')
            instrText.text = "PAGE"
            run._r.append(instrText)

            fldChar = OxmlElement('w:fldChar')
            fldChar.set(qn('w:fldCharType'), 'end')
            run._r.append(fldChar)

            # Выравнивание
            align_map = {
                'left': WD_ALIGN_PARAGRAPH.LEFT,
                'center': WD_ALIGN_PARAGRAPH.CENTER,
                'right': WD_ALIGN_PARAGRAPH.RIGHT
            }
            paragraph.alignment = align_map.get(alignment, WD_ALIGN_PARAGRAPH.RIGHT)

            # Настройка шрифта
            font_cfg = config.get('font', {})
            if font_cfg:
                run.font.name = font_cfg.get('name', 'Times New Roman')
                run.font.size = Pt(font_cfg.get('size', 12))
                run.font.bold = font_cfg.get('bold', True)
                run.font.italic = font_cfg.get('italic', False)
                if font_cfg.get('character_spacing'):
                    # Установка межсимвольного интервала (требует дополнительной обработки)
                    pass
        else:
            # Другой контент (например, текст)
            text = config.get('text', '')
            if text:
                paragraph = header.add_paragraph(text)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        logger.debug("Верхний колонтитул настроен")

    def _setup_footer(self, section, config):
        """
        Настраивает нижний колонтитул.
        """
        footer = section.footer
        for paragraph in footer.paragraphs:
            paragraph.clear()

        content_type = config.get('content', 'section_chapter_number')
        alignment = config.get('alignment', 'center')

        if content_type == 'section_chapter_number':
            paragraph = footer.add_paragraph("РАЗДЕЛ 1. ГЛАВА 1")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            font_cfg = config.get('font', {})
            if font_cfg:
                for run in paragraph.runs:
                    run.font.name = font_cfg.get('name', 'Times New Roman')
                    run.font.size = Pt(font_cfg.get('size', 12))
                    run.font.bold = font_cfg.get('bold', True)
                    run.font.italic = font_cfg.get('italic', False)
                    rFonts = run._element.get_or_add_rPr().get_or_add_rFonts()
                    rFonts.set(qn('w:eastAsia'), font_cfg.get('name', 'Times New Roman'))
        else:
            text = config.get('text', '')
            if text:
                paragraph = footer.add_paragraph(text)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        logger.debug("Нижний колонтитул настроен")

    def update_footer_text(self, doc, section_num, chapter_num):
        """
        Обновляет текст нижнего колонтитула с текущими номерами раздела и главы.
        """
        section = doc.sections[0]
        footer = section.footer
        for paragraph in footer.paragraphs:
            if paragraph.text.startswith("РАЗДЕЛ"):
                if chapter_num > 0:
                    paragraph.text = f"РАЗДЕЛ {section_num}. ГЛАВА {chapter_num}"
                else:
                    paragraph.text = f"РАЗДЕЛ {section_num}"
                break

    def apply_all(self, doc, orientation=None):
        """
        Применяет все настройки страницы и колонтитулов.

        :param doc: объект docx.Document
        :param orientation: ориентация страницы (опционально)
        """
        t0 = time.perf_counter()
        self.apply_page_setup(doc, orientation)
        t_page_setup = time.perf_counter() - t0
        perf_logger.info(f"[PERF] PageManager.apply_page_setup: {t_page_setup:.4f} сек")

        t1 = time.perf_counter()
        self.apply_headers_footers(doc)
        t_hf = time.perf_counter() - t1
        perf_logger.info(f"[PERF] PageManager.apply_headers_footers: {t_hf:.4f} сек")

        t_total = time.perf_counter() - t0
        perf_logger.info(f"[PERF] PageManager.apply_all (всего): {t_total:.4f} сек")
        logger.info("Настройки страницы и колонтитулов применены")
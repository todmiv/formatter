# style_manager.py

import docx
import time
import logging
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from typing import Dict, Any, Optional
from src.core.docx_utils import normalize_color

perf_logger = logging.getLogger('perf')


class StyleManager:
    """
    Менеджер стилей для унификации работы со стилями в генераторе и форматировщике.
    Обеспечивает создание и настройку стилей в документе согласно конфигурации.
    """

    def __init__(self, config_loader):
        """
        :param config_loader: экземпляр ConfigLoader с загруженной конфигурацией.
        """
        self.config_loader = config_loader

    def setup_styles(self, doc):
        """
        Настраивает все стили в документе согласно конфигурации.
        Аналог _setup_styles из md_to_docx.py.
        """
        t0 = time.perf_counter()
        styles_config = self.config_loader.styles

        # Настройка стиля Normal
        t1 = time.perf_counter()
        normal_style = doc.styles['Normal']
        normal_config = styles_config.get('Normal')
        if normal_config:
            self.apply_style_settings(normal_style, normal_config)
        perf_logger.info(f"[PERF] StyleManager: Normal style: {time.perf_counter() - t1:.4f} сек")

        # Настройка заголовков
        t1 = time.perf_counter()
        for heading_level in range(1, 5):
            style_name = f'Heading {heading_level}'
            if style_name in styles_config:
                heading_style = doc.styles[style_name]
                self.apply_style_settings(heading_style, styles_config[style_name])
        perf_logger.info(f"[PERF] StyleManager: Headings 1-4: {time.perf_counter() - t1:.4f} сек")

        # Настройка дополнительных стилей
        t1 = time.perf_counter()
        additional_styles = ['Caption', 'Table Note', 'List Paragraph', 'Note', 'Table Grid',
                             'Table Header', 'Table Column Header', 'Table Column Subheader',
                             'Table First Column', 'Table Numeric Cell', 'Table Text Cell',
                             'Appendix', 'Header', 'Footer']
        for style_name in additional_styles:
            if style_name in styles_config:
                if style_name not in doc.styles:
                    style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                else:
                    style = doc.styles[style_name]
                self.apply_style_settings(style, styles_config[style_name])
            elif style_name not in doc.styles:
                # Создаём стиль с настройками по умолчанию, даже если его нет в конфиге
                doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        perf_logger.info(f"[PERF] StyleManager: Additional styles ({len(additional_styles)}): {time.perf_counter() - t1:.4f} сек")

        t_total = time.perf_counter() - t0
        perf_logger.info(f"[PERF] StyleManager.setup_styles (всего): {t_total:.4f} сек")

    def apply_style_settings(self, style, config: Dict[str, Any]):
        """
        Применяет настройки стиля к объекту стиля python-docx.
        """
        if 'font' in config:
            font_cfg = config['font']
            style.font.name = font_cfg.get('name', 'Times New Roman')
            if 'size' in font_cfg:
                style.font.size = Pt(font_cfg['size'])
            style.font.bold = font_cfg.get('bold', False)
            style.font.italic = font_cfg.get('italic', False)
            color_val = font_cfg.get('color', 'black')
            color = normalize_color(color_val)
            if color is not None:
                style.font.color.rgb = color
            east_asia = font_cfg.get('east_asia', font_cfg.get('name', 'Times New Roman'))
            style._element.rPr.rFonts.set(qn('w:eastAsia'), east_asia)

        # Абзац
        if 'paragraph' in config:
            p_cfg = config['paragraph']
            para_format = style.paragraph_format

            # Выравнивание
            align_map = {
                'left': WD_ALIGN_PARAGRAPH.LEFT,
                'center': WD_ALIGN_PARAGRAPH.CENTER,
                'right': WD_ALIGN_PARAGRAPH.RIGHT,
                'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
            }
            if 'alignment' in p_cfg:
                para_format.alignment = align_map.get(p_cfg['alignment'], WD_ALIGN_PARAGRAPH.JUSTIFY)

            # Отступы
            if 'first_line_indent_cm' in p_cfg:
                para_format.first_line_indent = Cm(p_cfg['first_line_indent_cm'])
            if 'left_indent_cm' in p_cfg:
                para_format.left_indent = Cm(p_cfg['left_indent_cm'])

            # Интервалы
            if 'space_before_pt' in p_cfg:
                para_format.space_before = Pt(p_cfg['space_before_pt'])
            if 'space_after_pt' in p_cfg:
                para_format.space_after = Pt(p_cfg['space_after_pt'])

            # Межстрочный интервал
            if 'line_spacing' in p_cfg:
                val = p_cfg['line_spacing']
                if isinstance(val, float) and val < 10:  # множитель
                    para_format.line_spacing = val
                else:
                    para_format.line_spacing = Pt(val)

            # Разрыв страницы перед заголовком
            if p_cfg.get('page_break_before', False):
                para_format.page_break_before = True

        # Расширенные свойства (если есть)
        # text_transform, no_period_at_end и т.д. могут использоваться позже для валидации.
        # Здесь мы их пока только сохраняем в объекте стиля как атрибуты.
        if 'text_transform' in config:
            style.text_transform = config['text_transform']
        if 'no_period_at_end' in config:
            style.no_period_at_end = config['no_period_at_end']

    @staticmethod
    def _find_existing_base_style(styles) -> str:
        """Находит существующий базовый стиль в документе для создания нового."""
        for name in ('Normal', 'Обычный'):
            try:
                _ = styles[name]
                return name
            except KeyError:
                continue
        # Нет ни Normal, ни Обычный — берём первый доступный непользовательский стиль
        for style in styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH and not style.builtin:
                return style.name
        # Последний шанс: создаём Normal «с нуля»
        try:
            styles.add_style('Normal', WD_STYLE_TYPE.PARAGRAPH)
            return 'Normal'
        except Exception:
            # Абсолютный fallback: берём любой параграфный стиль
            for style in styles:
                if style.type == WD_STYLE_TYPE.PARAGRAPH:
                    return style.name
        raise RuntimeError("В документе нет ни одного параграфного стиля")

    def ensure_style_exists(self, doc, style_name: str, config: Optional[Dict] = None, cache: Dict = None):
        """
        Гарантирует наличие стиля в документе. Если стиля нет - создаёт.
        Если передан config, применяет настройки.
        Возвращает объект стиля.
        """
        if cache is None:
            cache = {}

        if style_name in cache:
            return cache[style_name]

        styles = doc.styles
        target_style = None
        try:
            target_style = styles[style_name]
        except KeyError:
            # Стиль не найден, создаём новый.
            # Сначала убеждаемся, что базовый стиль существует.
            base_style_name = self._find_existing_base_style(styles)
            base_style = styles[base_style_name]

            target_style = styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            target_style.base_style = base_style

        if config is None:
            # Получаем конфигурацию из config_loader
            config = self.config_loader.get_style_config(style_name)

        if config:
            self.apply_style_settings(target_style, config)

        cache[style_name] = target_style
        return target_style

    def apply_direct_overrides(self, paragraph, config: Dict):
        """
        Применяет прямое форматирование к параграфу на основе конфигурации.
        Используется для перебивки старого форматирования.
        """
        # Шрифт
        if 'font' in config:
            f_cfg = config['font']
            for run in paragraph.runs:
                if 'name' in f_cfg:
                    run.font.name = f_cfg['name']
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), f_cfg['name'])
                if 'size' in f_cfg:
                    run.font.size = Pt(f_cfg['size'])
                if 'bold' in f_cfg:
                    run.font.bold = f_cfg['bold']
                if 'italic' in f_cfg:
                    run.font.italic = f_cfg['italic']
                if 'color' in f_cfg:
                    color = normalize_color(f_cfg['color'])
                    if color is not None:
                        run.font.color.rgb = color

        # Абзац
        if 'paragraph' in config:
            p_cfg = config['paragraph']
            p_fmt = paragraph.paragraph_format

            align_map = {
                'left': WD_ALIGN_PARAGRAPH.LEFT,
                'center': WD_ALIGN_PARAGRAPH.CENTER,
                'right': WD_ALIGN_PARAGRAPH.RIGHT,
                'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
            }
            if 'alignment' in p_cfg:
                p_fmt.alignment = align_map.get(p_cfg['alignment'])

            if 'first_line_indent_cm' in p_cfg:
                p_fmt.first_line_indent = Cm(p_cfg['first_line_indent_cm'])
            if 'left_indent_cm' in p_cfg:
                p_fmt.left_indent = Cm(p_cfg['left_indent_cm'])

            # Интервалы
            if 'space_before_pt' in p_cfg:
                p_fmt.space_before = Pt(p_cfg['space_before_pt'])
            if 'space_after_pt' in p_cfg:
                p_fmt.space_after = Pt(p_cfg['space_after_pt'])

            # Межстрочный интервал
            if 'line_spacing' in p_cfg:
                val = p_cfg['line_spacing']
                if isinstance(val, float) and val < 10:  # множитель
                    p_fmt.line_spacing = val
                else:
                    p_fmt.line_spacing = Pt(val)

            # Разрыв страницы перед заголовком
            if p_cfg.get('page_break_before', False):
                p_fmt.page_break_before = True
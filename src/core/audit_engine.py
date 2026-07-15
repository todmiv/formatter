# audit_engine.py

import re
from typing import List, Dict, Any
from dataclasses import dataclass, field
from enum import Enum
import docx
from docx.oxml.ns import qn
import logging
logger = logging.getLogger(__name__)

from src.core.config_loader import ConfigLoader
from src.core.docx_utils import get_xml_attr_twips, normalize_color, has_direct_paragraph_formatting, has_direct_run_formatting

class Severity(Enum):
    INFO = "INFO"
    WARNING = "WARNING"
    CRITICAL = "CRITICAL"

@dataclass
class AuditIssue:
    """Модель проблемы, найденной при аудите."""
    id: str
    severity: Severity
    category: str  # FONT, PARAGRAPH, STYLE, PAGE
    element_type: str  # Имя стиля (напр. "Heading 1")
    location: Dict[str, Any]  # {'paragraph_index': 10, 'page_estimate': 3} или {'table_index': 0}
    description: str
    current_value: str
    expected_value: str
    recommendation: str = ""  # Рекомендация по исправлению
    auto_fixable: bool = True
    fix_payload: Dict = field(default_factory=dict)

class AuditEngine:
    """
    Движок аудита документов.
    Сканирует DOCX, сравнивает свойства с конфигом и генерирует отчет.
    """
    
    # Допуски для сравнения (twips)
    FONT_SIZE_TOLERANCE_TWIPS = 10  # ~0.5 пт
    INDENT_TOLERANCE_TWIPS = 20     # ~1 мм
    SPACING_TOLERANCE_TWIPS = 20

    def __init__(self, config_loader: ConfigLoader,
                 font_size_tolerance_twips: int = None,
                 indent_tolerance_twips: int = None,
                 spacing_tolerance_twips: int = None,
                 progress_callback = None):
        self.config = config_loader
        self.issues: List[AuditIssue] = []
        self.issue_counter = 0
        self.progress_callback = progress_callback
        
        # Устанавливаем допуски: если переданы, используем их, иначе значения по умолчанию из класса
        self.FONT_SIZE_TOLERANCE_TWIPS = font_size_tolerance_twips if font_size_tolerance_twips is not None else self.FONT_SIZE_TOLERANCE_TWIPS
        self.INDENT_TOLERANCE_TWIPS = indent_tolerance_twips if indent_tolerance_twips is not None else self.INDENT_TOLERANCE_TWIPS
        self.SPACING_TOLERANCE_TWIPS = spacing_tolerance_twips if spacing_tolerance_twips is not None else self.SPACING_TOLERANCE_TWIPS

    def _update_progress(self, stage: str, progress: int, total: int, message: str = ""):
        """Вызывает колбэк прогресса, если он установлен."""
        if self.progress_callback:
            self.progress_callback(stage, progress, total, message)

    def scan_document(self, doc_path: str = None, doc=None) -> List[AuditIssue]:
        """Запускает полный аудит документа.

        :param doc_path: путь к документу (используется если doc не передан).
        :param doc: объект Document для аудита (DI).
        :return: список найденных проблем.
        """
        self.issues = []
        self.issue_counter = 0

        if doc is None:
            if doc_path is None:
                raise ValueError("Необходимо указать doc_path или doc")
            logger.info(f"Начало аудита документа: {doc_path}")
            try:
                doc = docx.Document(doc_path)
                logger.debug(f"Документ открыт: параграфов={len(doc.paragraphs)}, таблиц={len(doc.tables)}")
            except Exception as e:
                logger.error(f"Не удалось открыть документ {doc_path}: {e}", exc_info=True)
                raise RuntimeError(f"Не удалось открыть документ: {e}")
        else:
            logger.info("Аудит переданного объекта Document")

        # Подсчёт элементов для прогресса
        total_paragraphs = len(doc.paragraphs)
        total_tables = len(doc.tables)
        total_sections = len(doc.sections)
        total_images = len(doc.inline_shapes)
        logger.debug(f"Элементы для аудита: параграфы={total_paragraphs}, таблицы={total_tables}, секции={total_sections}, изображения={total_images}")

        # 1. Аудит настроек страницы
        self._update_progress('page_setup', 0, 1, "Проверка настроек страницы...")
        try:
            self._audit_page_setup(doc)
        except Exception as e:
            logger.error(f"Ошибка при аудите настроек страницы: {e}", exc_info=True)
            # Продолжаем аудит, но добавляем проблему в лог
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="SYSTEM",
                elem_type="Page Setup",
                loc={'section': 0, 'page': 0},
                desc=f"Ошибка аудита настроек страницы: {e}",
                current="Ошибка",
                expected="Успешный аудит",
                payload={'error': str(e)},
                auto_fixable=False
            ))
        self._update_progress('page_setup', 1, 1, "Настройки страницы проверены.")

        # 2. Аудит параграфов
        self._update_progress('paragraphs', 0, total_paragraphs, "Аудит параграфов...")
        lines_per_page_estimate = 25
        current_line = 0

        for idx, para in enumerate(doc.paragraphs):
            page_est = (current_line // lines_per_page_estimate) + 1
            
            try:
                issues = self._audit_paragraph(para, idx, page_est)
                self.issues.extend(issues)
            except Exception as e:
                logger.error(f"Ошибка при аудите параграфа {idx}: {e}", exc_info=True)
                # Добавляем системную проблему
                self.issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="SYSTEM",
                    elem_type="Paragraph Audit",
                    loc={'index': idx, 'page': page_est},
                    desc=f"Ошибка аудита параграфа: {e}",
                    current="Ошибка",
                    expected="Успешный аудит",
                    payload={'error': str(e), 'paragraph_index': idx},
                    auto_fixable=False
                ))
            
            text_len = len(para.text)
            lines = max(1, text_len // 60)
            current_line += lines

            # Обновляем прогресс каждые 10 параграфов или на последнем
            if idx % 10 == 0 or idx == total_paragraphs - 1:
                self._update_progress('paragraphs', idx + 1, total_paragraphs,
                                      f"Проверено параграфов: {idx + 1}/{total_paragraphs}")

        self._update_progress('paragraphs', total_paragraphs, total_paragraphs, "Аудит параграфов завершён.")

        # 3. Аудит таблиц
        self._update_progress('tables', 0, max(1, total_tables), "Аудит таблиц...")
        try:
            self._audit_tables(doc)
        except Exception as e:
            logger.error(f"Ошибка при аудите таблиц: {e}", exc_info=True)
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="SYSTEM",
                elem_type="Table Audit",
                loc={'section': 0, 'page': 0},
                desc=f"Ошибка аудита таблиц: {e}",
                current="Ошибка",
                expected="Успешный аудит",
                payload={'error': str(e)},
                auto_fixable=False
            ))
        if total_tables > 0:
            self._update_progress('tables', total_tables, total_tables, "Аудит таблиц завершён.")
        else:
            self._update_progress('tables', 1, 1, "Таблицы отсутствуют.")

        # 4. Аудит колонтитулов
        self._update_progress('headers_footers', 0, max(1, total_sections), "Аудит колонтитулов...")
        try:
            self._audit_headers_footers(doc)
        except Exception as e:
            logger.error(f"Ошибка при аудите колонтитулов: {e}", exc_info=True)
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="SYSTEM",
                elem_type="Header/Footer Audit",
                loc={'section': 0, 'page': 0},
                desc=f"Ошибка аудита колонтитулов: {e}",
                current="Ошибка",
                expected="Успешный аудит",
                payload={'error': str(e)},
                auto_fixable=False
            ))
        if total_sections > 0:
            self._update_progress('headers_footers', total_sections, total_sections, "Аудит колонтитулов завершён.")
        else:
            self._update_progress('headers_footers', 1, 1, "Колонтитулы отсутствуют.")

        # 5. Аудит рисунков
        self._update_progress('images', 0, max(1, total_images), "Аудит рисунков...")
        try:
            self._audit_images(doc)
        except Exception as e:
            logger.error(f"Ошибка при аудите рисунков: {e}", exc_info=True)
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="SYSTEM",
                elem_type="Image Audit",
                loc={'section': 0, 'page': 0},
                desc=f"Ошибка аудита рисунков: {e}",
                current="Ошибка",
                expected="Успешный аудит",
                payload={'error': str(e)},
                auto_fixable=False
            ))
        if total_images > 0:
            self._update_progress('images', total_images, total_images, "Аудит рисунков завершён.")
        else:
            self._update_progress('images', 1, 1, "Рисунки отсутствуют.")

        self._update_progress('complete', 100, 100, "Аудит документа завершён.")
        return self.issues

    def _detect_target_style(self, para) -> str:
        """
        Определяет целевой стиль для параграфа.
        ГАРАНТИРУЕТ возврат имени стиля. Никогда не возвращает None.
        
        Логика приоритетов:
        1. Если текущий стиль явно совпадает с активным стилем из конфига -> используем его.
        2. Если текст соответствует правилам detection_rules (заголовки, списки) -> используем найденный стиль.
        3. ВО ВСЕХ ОСТАЛЬНЫХ СЛУЧАЯХ -> принудительно считаем целью стиль 'Normal' (или 'Обычный').
           Это закрывает логическую дыру: любой текст, не являющийся заголовком/списком, 
           должен быть оформлен как основной текст.
        """
        current_style_name = para.style.name
        
        # 1. Проверка явного соответствия активным стилям конфига
        if current_style_name in self.config.styles:
            if self.config.styles[current_style_name].get('enabled', True):
                return current_style_name
        
        # 2. Проверка по правилам (Regex, ключевые слова)
        text = para.text.strip()
        if text:
            rules = self.config.get_detection_rules()
            for style_name, patterns in rules.items():
                # Проверяем, существует ли такой стиль в конфиге и активен ли он
                if style_name not in self.config.styles or not self.config.styles[style_name].get('enabled', True):
                    continue
                    
                for pattern in patterns:
                    try:
                        if re.search(pattern, text):
                            return style_name
                    except re.error:
                        logger.warning(f"Неверное регулярное выражение в правиле для {style_name}: {pattern}")
                        continue
        
        # 3. FALLBACK: Все, что не опознано как спец. стиль, должно быть Normal.
        # Проверяем наличие аналогов Normal в конфиге
        if "Normal" in self.config.styles:
            return "Normal"
        if "Обычный" in self.config.styles:
            return "Обычный"

        # Крайний случай: если в конфиге вообще нет Normal (ошибка конфига),
        # возвращаем текущий стиль, чтобы не ломать выполнение, но это редкость.
        return current_style_name

    def _audit_paragraph(self, para, index: int, page_est: int) -> List[AuditIssue]:
        local_issues = []
        
        # Теперь target_style_name гарантированно строка
        target_style_name = self._detect_target_style(para)
        
        # Получаем эталонные настройки
        style_config = self.config.get_style_config(target_style_name)
        
        # Если стиль вдруг отключен в конфиге (хотя _detect_target_style это проверяет), пропускаем
        if not style_config or not style_config.get('enabled', True):
            return local_issues

        # --- Логика выявления несоответствий ---
        
        # 1. Проверка свойств (Шрифт, Абзац)
        # Сравниваем фактическое состояние параграфа с эталоном target_style_name
        if 'font' in style_config:
            issues = self._check_font(para, style_config['font'], target_style_name, index, page_est)
            local_issues.extend(issues)

        if 'paragraph' in style_config:
            issues = self._check_paragraph_formatting(para, style_config['paragraph'], target_style_name, index, page_est)
            local_issues.extend(issues)

        # 2. Проверка именования стиля (Семантическая чистота)
        # Если параграф должен быть "Normal", но у него стиль "Text Body" или "Default Paragraph Font"
        # это считается ошибкой, даже если визуально он похож на Normal.
        actual_style_name = para.style.name
        if actual_style_name != target_style_name:
            # Если это fallback-случай (ожидается Normal, а там что-то непонятное)
            if target_style_name in ["Normal", "Обычный"]:
                local_issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="STYLE",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc=f"Некорректное имя стиля для основного текста",
                    current=actual_style_name,
                    expected=target_style_name,
                    payload={'style': target_style_name, 'prop': 'style_name'}
                ))
            else:
                # Если ожидается заголовок, а стиль другой
                local_issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="STYLE",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc=f"Ожидается стиль '{target_style_name}', но используется '{actual_style_name}'",
                    current=actual_style_name,
                    expected=target_style_name,
                    payload={'style': target_style_name, 'prop': 'style_name'}
                ))

        return local_issues


    def _check_font(self, para, expected_font: Dict, style_name: str, idx: int, page: int) -> List[AuditIssue]:
        issues = []
        runs = [r for r in para.runs if r.text.strip()]
        if not runs:
            return issues

        rep_run = runs[0]

        # Имя шрифта — пропускаем если задано прямое форматирование
        if 'name' in expected_font and not has_direct_run_formatting(para, 'font_name'):
            actual_font_name = None
            rPr = rep_run._element.find(qn('w:rPr'))
            if rPr is not None:
                rFonts_elem = rPr.find(qn('w:rFonts'))
                if rFonts_elem is not None:
                    actual_font_name = rFonts_elem.get(qn('w:ascii')) or rFonts_elem.get(qn('w:hAnsi'))

            if actual_font_name is None and hasattr(para.style.font, 'name'):
                actual_font_name = para.style.font.name

            if actual_font_name and actual_font_name != expected_font['name']:
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="FONT",
                    elem_type=style_name,
                    loc={'index': idx, 'page': page},
                    desc=f"Неверный шрифт",
                    current=actual_font_name,
                    expected=expected_font['name'],
                    payload={'style': style_name, 'prop': 'font_name'}
                ))

        # Размер шрифта — пропускаем если задано прямое форматирование
        if not has_direct_run_formatting(para, 'font_size'):
            actual_size_twips = None
            if rep_run.font.size:
                actual_size_twips = rep_run.font.size.twips
            elif para.style.font.size:
                actual_size_twips = para.style.font.size.twips

            if actual_size_twips is not None:
                expected_twips = expected_font.get('size_twips', 0)
                if abs(actual_size_twips - expected_twips) > self.FONT_SIZE_TOLERANCE_TWIPS:
                    issues.append(self._create_issue(
                        severity=Severity.CRITICAL,
                        category="FONT",
                        elem_type=style_name,
                        loc={'index': idx, 'page': page},
                        desc=f"Неверный размер шрифта",
                        current=f"{actual_size_twips/20:.1f} пт",
                        expected=f"{expected_font['size']} пт",
                        payload={'style': style_name, 'prop': 'font_size'}
                    ))

        # Цвет шрифта — пропускаем если задано прямое форматирование
        if 'color' in expected_font and not has_direct_run_formatting(para, 'color'):
            expected_color_str = expected_font['color']
            expected_color = normalize_color(expected_color_str)
            actual_color = rep_run.font.color.rgb if rep_run.font.color and rep_run.font.color.rgb else None
            if actual_color != expected_color:
                actual_str = str(actual_color) if actual_color else "не задан"
                expected_str = str(expected_color) if expected_color else "не задан"
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="FONT",
                    elem_type=style_name,
                    loc={'index': idx, 'page': page},
                    desc=f"Неверный цвет шрифта",
                    current=actual_str,
                    expected=expected_str,
                    payload={'style': style_name, 'prop': 'font_color'}
                ))

        return issues

    # Метод _get_xml_attr_twips удалён, используем функцию из docx_utils

    def _check_paragraph_formatting(self, para, expected_para: Dict, style_name: str, idx: int, page: int) -> List[AuditIssue]:
        issues = []
        pPr = para._element.pPr

        # --- Отступ первой строки — пропускаем если прямое форматирование ---
        if not has_direct_paragraph_formatting(para, 'first_line'):
            expected_first_twips = expected_para.get('first_line_indent_twips', 0)
            actual_first_twips = 0

            if pPr is not None:
                indent_elem = pPr.find(qn('w:ind'))
                if indent_elem is not None:
                    val = get_xml_attr_twips(indent_elem, 'w:firstLine')
                    if val is not None:
                        actual_first_twips = val

            if pPr is None or pPr.find(qn('w:ind')) is None:
                if para.style.paragraph_format.first_line_indent:
                    actual_first_twips = para.style.paragraph_format.first_line_indent.twips

            if abs(actual_first_twips - expected_first_twips) > self.INDENT_TOLERANCE_TWIPS:
                 issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="PARAGRAPH",
                    elem_type=style_name,
                    loc={'index': idx, 'page': page},
                    desc="Неверный абзацный отступ",
                    current=f"{actual_first_twips/567:.2f} см",
                    expected=f"{expected_para.get('first_line_indent_cm', 0)} см",
                    payload={'style': style_name, 'prop': 'first_line_indent'}
                ))

        # --- Выравнивание ---
        actual_align = 'left'
        if pPr is not None:
            jc_elem = pPr.find(qn('w:jc'))
            if jc_elem is not None:
                val = jc_elem.get(qn('w:val'))
                if val:
                    map_align = {'both': 'justify', 'left': 'left', 'center': 'center', 'right': 'right'}
                    actual_align = map_align.get(val, 'left')

        # Если нет прямого форматирования — наследуем от стиля
        if not has_direct_paragraph_formatting(para, 'alignment'):
            try:
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                style_align = para.style.paragraph_format.alignment
                if style_align is not None:
                    wd_map = {
                        WD_ALIGN_PARAGRAPH.LEFT: 'left',
                        WD_ALIGN_PARAGRAPH.CENTER: 'center',
                        WD_ALIGN_PARAGRAPH.RIGHT: 'right',
                        WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify',
                        WD_ALIGN_PARAGRAPH.DISTRIBUTE: 'justify',
                    }
                    actual_align = wd_map.get(style_align, 'left')
            except Exception:
                pass

        expected_align = expected_para.get('alignment', 'left')
        if actual_align != expected_align:
                issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="PARAGRAPH",
                    elem_type=style_name,
                    loc={'index': idx, 'page': page},
                    desc="Неверное выравнивание",
                    current=actual_align,
                    expected=expected_align,
                    payload={'style': style_name, 'prop': 'alignment'}
                ))

        # --- Интервалы — пропускаем если прямое форматирование ---
        if not has_direct_paragraph_formatting(para, 'spacing'):
            spacing_elem = None
            if pPr is not None:
                spacing_elem = pPr.find(qn('w:spacing'))

            if 'space_before_twips' in expected_para:
                exp_before = expected_para['space_before_twips']
                act_before = 0
                if spacing_elem is not None:
                    val = get_xml_attr_twips(spacing_elem, 'w:before')
                    if val is not None:
                        act_before = val
                elif para.style.paragraph_format.space_before:
                    act_before = para.style.paragraph_format.space_before.twips

                if abs(act_before - exp_before) > self.SPACING_TOLERANCE_TWIPS:
                    issues.append(self._create_issue(
                        severity=Severity.WARNING,
                        category="PARAGRAPH",
                        elem_type=style_name,
                        loc={'index': idx, 'page': page},
                        desc="Неверный интервал перед абзацем",
                        current=f"{act_before/20:.0f} пт",
                        expected=f"{expected_para.get('space_before_pt', 0)} пт",
                        payload={'style': style_name, 'prop': 'space_before'}
                    ))

            if 'space_after_twips' in expected_para:
                exp_after = expected_para['space_after_twips']
                act_after = 0
                if spacing_elem is not None:
                    val = get_xml_attr_twips(spacing_elem, 'w:after')
                    if val is not None:
                        act_after = val
                elif para.style.paragraph_format.space_after:
                    act_after = para.style.paragraph_format.space_after.twips

                if abs(act_after - exp_after) > self.SPACING_TOLERANCE_TWIPS:
                    issues.append(self._create_issue(
                        severity=Severity.WARNING,
                        category="PARAGRAPH",
                        elem_type=style_name,
                        loc={'index': idx, 'page': page},
                        desc="Неверный интервал после абзаца",
                        current=f"{act_after/20:.0f} пт",
                        expected=f"{expected_para.get('space_after_pt', 0)} пт",
                        payload={'style': style_name, 'prop': 'space_after'}
                    ))
                        
        return issues

    def _audit_tables(self, doc):
        """Аудит таблиц в документе."""
        expected_style = "Table Grid"
        style_config = self.config.get_style_config(expected_style)
        if not style_config:
            logger.warning(f"Стиль таблиц '{expected_style}' не найден в конфигурации. Пропускаем аудит таблиц.")
            return

        for idx, table in enumerate(doc.tables):
            actual_style = table.style.name if table.style else "No Style"
            if actual_style != expected_style:
                self.issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="TABLE",
                    elem_type=expected_style,
                    loc={'table_index': idx, 'page': 0, 'type': 'table'},
                    desc=f"Неверный стиль таблицы",
                    current=actual_style,
                    expected=expected_style,
                    payload={'style': expected_style, 'prop': 'table_style'},
                    auto_fixable=True
                ))
            # Дополнительные проверки можно добавить здесь

    def _hf_has_content(self, hf) -> bool:
        """Проверяет, есть ли содержимое в колонтителе (текст, SDT, поля, рисунки)."""
        if hf is None:
            return False
        # Проверяем XML на уровне колонтитула (SDT могут быть родителями параграфов)
        hf_xml = hf._element.xml
        has_sdt_at_root = 'w:sdt' in hf_xml
        for para in hf.paragraphs:
            # Текст
            if para.text.strip():
                return True
            # SDT (Structured Document Tags — нумерация страниц, Автотекст)
            # Проверяем и параграф, и его родителя (SDT может оборачивать параграф)
            p_xml = para._element.xml
            parent = para._element.getparent()
            parent_xml = parent.xml if parent is not None else ''
            if 'w:sdt' in p_xml or 'w:sdt' in parent_xml:
                return True
            # Поля (PAGE, NUMPAGES и т.д.)
            if 'fldChar' in p_xml or 'instrText' in p_xml or 'fldSimple' in p_xml:
                return True
            # Рисунки / фигуры
            if 'w:drawing' in p_xml or 'w:pict' in p_xml:
                return True
        # Если SDT на уровне колонтитула, но параграфов нет — считаем что содержимое есть
        if has_sdt_at_root:
            return True
        return False

    def _audit_headers_footers(self, doc):
        """Аудит колонтитулов документа."""
        config = self.config.get_headers_footers_config()
        if not config:
            logger.debug("Конфигурация колонтитулов отсутствует, пропускаем аудит.")
            return

        header_enabled = config.get('header', {}).get('enabled', False)
        footer_enabled = config.get('footer', {}).get('enabled', False)

        for sect_idx, section in enumerate(doc.sections):
            # Аудит верхнего колонтитула
            if header_enabled:
                header = section.header
                # Пропускаем секции с связанными колонтитулами — они наследуют от предыдущей секции
                if header is not None and header.is_linked_to_previous:
                    continue
                # Пропускаем первую секцию если конфиг помечен как first_page_empty
                if sect_idx == 0 and config.get('header', {}).get('first_page_empty', False):
                    continue
                if not self._hf_has_content(header):
                    self.issues.append(self._create_issue(
                        severity=Severity.WARNING,
                        category="HEADER_FOOTER",
                        elem_type="Header",
                        loc={'section': sect_idx, 'page': 0},
                        desc="Верхний колонтитул отсутствует или пуст",
                        current="Отсутствует",
                        expected="Должен присутствовать",
                        payload={},
                        auto_fixable=False
                    ))
                else:
                    self._audit_header_footer_paragraphs(header, "Header", sect_idx)

            # Аудит нижнего колонтитула
            if footer_enabled:
                footer = section.footer
                # Пропускаем секции с связанными колонтитулами — они наследуют от предыдущей секции
                if footer is not None and footer.is_linked_to_previous:
                    continue
                # Пропускаем первую секцию если конфиг помечен как first_page_empty
                if sect_idx == 0 and config.get('footer', {}).get('first_page_empty', False):
                    continue
                if not self._hf_has_content(footer):
                    self.issues.append(self._create_issue(
                        severity=Severity.WARNING,
                        category="HEADER_FOOTER",
                        elem_type="Footer",
                        loc={'section': sect_idx, 'page': 0},
                        desc="Нижний колонтитул отсутствует или пуст",
                        current="Отсутствует",
                        expected="Должен присутствовать",
                        payload={},
                        auto_fixable=False
                    ))
                else:
                    self._audit_header_footer_paragraphs(footer, "Footer", sect_idx)

        logger.debug("Аудит колонтитулов выполнен.")

    def _audit_header_footer_paragraphs(self, hf, hf_type: str, sect_idx: int):
        """Проверяет параграфы в колонтитуле."""
        target_style_name = hf_type  # "Header" или "Footer"
        style_config = self.config.get_style_config(target_style_name)
        if not style_config or not style_config.get('enabled', True):
            logger.debug(f"Стиль {target_style_name} не настроен, пропускаем проверку.")
            return

        container = hf_type.lower()  # 'header' или 'footer'
        for para_idx, para in enumerate(hf.paragraphs):
            # Проверка стиля
            actual_style_name = para.style.name
            if actual_style_name != target_style_name:
                self.issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="HEADER_FOOTER",
                    elem_type=target_style_name,
                    loc={'section': sect_idx, 'paragraph': para_idx, 'page': 0, 'container': container},
                    desc=f"Неверный стиль колонтитула",
                    current=actual_style_name,
                    expected=target_style_name,
                    payload={'style': target_style_name, 'prop': 'style_name'}
                ))

            # Проверка шрифта
            if 'font' in style_config:
                issues = self._check_font(para, style_config['font'], target_style_name, para_idx, 0)
                for issue in issues:
                    issue.category = "HEADER_FOOTER"
                    issue.location = {'section': sect_idx, 'paragraph': para_idx, 'page': 0, 'container': container}
                    self.issues.append(issue)

            # Проверка абзаца
            if 'paragraph' in style_config:
                issues = self._check_paragraph_formatting(para, style_config['paragraph'], target_style_name, para_idx, 0)
                for issue in issues:
                    issue.category = "HEADER_FOOTER"
                    issue.location = {'section': sect_idx, 'paragraph': para_idx, 'page': 0, 'container': container}
                    self.issues.append(issue)

            # Проверка положения от края (если есть в конфиге)
            # Пока пропустим, так как требует анализа свойств секции.

    def _audit_images(self, doc):
        """Аудит рисунков в документе."""
        # Получаем конфигурацию рисунков
        formatting_rules = self.config.get_formatting_rules()
        figures_config = formatting_rules.get('figures', {}) if formatting_rules else {}
        if not figures_config:
            logger.debug("Конфигурация рисунков отсутствует, пропускаем аудит.")
            return

        # Получаем стиль для подписей (05_Номер таблицы или Caption)
        caption_style_name = '05_Номер таблицы'
        caption_style_config = self.config.get_style_config(caption_style_name)
        if not caption_style_config or not caption_style_config.get('enabled', True):
            caption_style_name = 'Caption'
            caption_style_config = self.config.get_style_config(caption_style_name)
        if not caption_style_config or not caption_style_config.get('enabled', True):
            logger.debug("Стиль подписей не настроен, пропускаем проверку подписей.")
            # Но всё равно проверим наличие подписей

        inline_shapes = doc.inline_shapes
        if not inline_shapes:
            logger.debug("В документе нет рисунков.")
            return

        # Собираем все параграфы-подписи
        caption_paragraphs = []
        for idx, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if re.search(r'^Рисунок\s+\d+', text) or re.search(r'^Рис\.\s*\d+', text, re.IGNORECASE):
                caption_paragraphs.append((idx, para))

        # Проверка подписей на соответствие стилю
        for idx, para in caption_paragraphs:
            # Проверка стиля
            actual_style_name = para.style.name
            if actual_style_name != caption_style_name:
                self.issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="FIGURE",
                    elem_type=caption_style_name,
                    loc={'paragraph_index': idx, 'page': 0},
                    desc="Подпись рисунка имеет неверный стиль",
                    current=actual_style_name,
                    expected=caption_style_name,
                    payload={'style': caption_style_name, 'prop': 'style_name'}
                ))
            # Проверка шрифта и абзаца, если есть конфигурация
            if caption_style_config:
                if 'font' in caption_style_config:
                    issues = self._check_font(para, caption_style_config['font'], caption_style_name, idx, 0)
                    for issue in issues:
                        issue.location = {'paragraph_index': idx, 'page': 0}
                        self.issues.append(issue)
                if 'paragraph' in caption_style_config:
                    issues = self._check_paragraph_formatting(para, caption_style_config['paragraph'], caption_style_name, idx, 0)
                    for issue in issues:
                        issue.location = {'paragraph_index': idx, 'page': 0}
                        self.issues.append(issue)

        # Проверка наличия подписи для каждого рисунка
        # Упрощённо: если количество рисунков больше количества подписей, выдаём предупреждение
        if len(inline_shapes) > len(caption_paragraphs):
            self.issues.append(self._create_issue(
                severity=Severity.WARNING,
                category="FIGURE",
                elem_type="Image",
                loc={'global': 0},
                desc="Не для всех рисунков есть подписи",
                current=f"{len(caption_paragraphs)} подписей",
                expected=f"не менее {len(inline_shapes)} подписей",
                auto_fixable=False
            ))

        # Проверка положения подписи (должна быть над рисунком)
        # Пропускаем из-за сложности точного определения

        # Проверка интервала после рисунка (spacing after_figure_pt)
        spacing_after_pt = figures_config.get('spacing', {}).get('after_figure_pt', 6)
        # Пропускаем, так как требует анализа позиции параграфа после рисунка

        # Проверка ссылок в тексте (references required_in_text)
        if figures_config.get('references', {}).get('required_in_text', True):
            # Собираем номера рисунков из подписей
            figure_numbers = []
            for idx, para in caption_paragraphs:
                text = para.text.strip()
                match = re.search(r'Рисунок\s+(\d+)', text) or re.search(r'Рис\.\s*(\d+)', text, re.IGNORECASE)
                if match:
                    figure_numbers.append(match.group(1))
            # Ищем упоминания в тексте
            for para in doc.paragraphs:
                text = para.text
                for num in figure_numbers[:]:  # копия для удаления
                    if re.search(rf'\bрис\.?\s*{num}\b', text, re.IGNORECASE) or \
                       re.search(rf'\bрисунок\s*{num}\b', text, re.IGNORECASE):
                        figure_numbers.remove(num)
                        break
            # Оставшиеся номера — отсутствующие ссылки
            for num in figure_numbers:
                self.issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="FIGURE",
                    elem_type="Image",
                    loc={'global': 0},
                    desc=f"Отсутствует ссылка на рисунок {num} в тексте",
                    current="Нет ссылки",
                    expected="Должна быть ссылка в тексте",
                    auto_fixable=False
                ))

        logger.debug("Аудит рисунков выполнен.")

    def _audit_page_setup(self, doc):
        """Аудит настроек страницы: поля, ориентация, размер бумаги."""
        config = self.config.get_page_setup()
        if not config:
            logger.debug("Конфигурация страницы отсутствует, пропускаем аудит.")
            return

        # Берем первую секцию (предполагаем одинаковые настройки во всем документе)
        if not doc.sections:
            return
        section = doc.sections[0]

        # Определяем ориентацию
        page_width_twips = section.page_width.twips
        page_height_twips = section.page_height.twips
        is_portrait = page_height_twips > page_width_twips
        orientation = "portrait" if is_portrait else "landscape"

        # Ожидаемая ориентация по умолчанию
        expected_orientation = config.get('orientation_default', 'portrait')
        if orientation != expected_orientation:
            self.issues.append(self._create_issue(
                severity=Severity.WARNING,
                category="PAGE",
                elem_type="Page Setup",
                loc={'section': 0, 'page': 0},
                desc=f"Неверная ориентация страницы",
                current=orientation,
                expected=expected_orientation,
                payload={},
                auto_fixable=False
            ))

        # Выбираем соответствующие поля
        margins_config = config.get('portrait' if orientation == 'portrait' else 'landscape', {})
        expected_left_cm = margins_config.get('left_margin_cm', 0)
        expected_right_cm = margins_config.get('right_margin_cm', 0)
        expected_top_cm = margins_config.get('top_margin_cm', 0)
        expected_bottom_cm = margins_config.get('bottom_margin_cm', 0)

        # Конвертация в twips (1 см = 567 twips)
        cm_to_twips = 567
        expected_left_twips = int(expected_left_cm * cm_to_twips)
        expected_right_twips = int(expected_right_cm * cm_to_twips)
        expected_top_twips = int(expected_top_cm * cm_to_twips)
        expected_bottom_twips = int(expected_bottom_cm * cm_to_twips)

        # Фактические поля в twips
        actual_left_twips = section.left_margin.twips
        actual_right_twips = section.right_margin.twips
        actual_top_twips = section.top_margin.twips
        actual_bottom_twips = section.bottom_margin.twips

        # Допуск для полей (используем INDENT_TOLERANCE_TWIPS = 20 twips ~ 0.35 мм)
        tolerance = self.INDENT_TOLERANCE_TWIPS

        # Проверка левого поля
        if abs(actual_left_twips - expected_left_twips) > tolerance:
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="PAGE",
                elem_type="Page Setup",
                loc={'section': 0, 'page': 0},
                desc="Неверное левое поле",
                current=f"{actual_left_twips/cm_to_twips:.2f} см",
                expected=f"{expected_left_cm} см",
                payload={'prop': 'left_margin'},
                auto_fixable=True
            ))

        # Проверка правого поля
        if abs(actual_right_twips - expected_right_twips) > tolerance:
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="PAGE",
                elem_type="Page Setup",
                loc={'section': 0, 'page': 0},
                desc="Неверное правое поле",
                current=f"{actual_right_twips/cm_to_twips:.2f} см",
                expected=f"{expected_right_cm} см",
                payload={'prop': 'right_margin'},
                auto_fixable=True
            ))

        # Проверка верхнего поля
        if abs(actual_top_twips - expected_top_twips) > tolerance:
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="PAGE",
                elem_type="Page Setup",
                loc={'section': 0, 'page': 0},
                desc="Неверное верхнее поле",
                current=f"{actual_top_twips/cm_to_twips:.2f} см",
                expected=f"{expected_top_cm} см",
                payload={'prop': 'top_margin'},
                auto_fixable=True
            ))

        # Проверка нижнего поля
        if abs(actual_bottom_twips - expected_bottom_twips) > tolerance:
            self.issues.append(self._create_issue(
                severity=Severity.CRITICAL,
                category="PAGE",
                elem_type="Page Setup",
                loc={'section': 0, 'page': 0},
                desc="Неверное нижнее поле",
                current=f"{actual_bottom_twips/cm_to_twips:.2f} см",
                expected=f"{expected_bottom_cm} см",
                payload={'prop': 'bottom_margin'},
                auto_fixable=True
            ))

        # Проверка размера бумаги (опционально)
        paper_size = config.get('paper_size', 'A4')
        if paper_size == 'A4':
            # Размеры A4 в мм: 210 x 297
            expected_width_mm = 210
            expected_height_mm = 297
            # Конвертация в twips (1 мм = 56.7 twips)
            mm_to_twips = 56.7
            expected_width_twips = int(expected_width_mm * mm_to_twips)
            expected_height_twips = int(expected_height_mm * mm_to_twips)
            # Допуск 2 мм (~113 twips)
            size_tolerance = 113
            if (abs(page_width_twips - expected_width_twips) > size_tolerance or
                abs(page_height_twips - expected_height_twips) > size_tolerance):
                self.issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="PAGE",
                    elem_type="Page Setup",
                    loc={'section': 0, 'page': 0},
                    desc="Неверный размер бумаги",
                    current=f"{page_width_twips/mm_to_twips:.0f}x{page_height_twips/mm_to_twips:.0f} мм",
                    expected=f"{expected_width_mm}x{expected_height_mm} мм",
                    payload={},
                    auto_fixable=False
                ))

        logger.debug("Аудит настроек страницы выполнен.")

    def _create_issue(self, severity: Severity, category: str, elem_type: str,
                      loc: Dict, desc: str, current: Any, expected: Any, payload: Dict,
                      auto_fixable: bool = True, recommendation: str = "") -> AuditIssue:
        self.issue_counter += 1
        uid = f"{category}_{self.issue_counter}"
        
        # Добавляем тип элемента в location для универсальной обработки
        location_with_type = dict(loc)  # копируем, чтобы не менять оригинальный словарь
        if 'type' not in location_with_type:
            if category == 'TABLE':
                location_with_type['type'] = 'table'
            elif category in ('PARAGRAPH', 'FONT', 'STYLE', 'HEADER', 'FOOTER'):
                location_with_type['type'] = 'paragraph'
            else:
                # Для остальных категорий по умолчанию 'paragraph'
                location_with_type['type'] = 'paragraph'
        
        return AuditIssue(
            id=uid,
            severity=severity,
            category=category,
            element_type=elem_type,
            location=location_with_type,
            description=desc,
            current_value=str(current),
            expected_value=str(expected),
            recommendation=recommendation,
            auto_fixable=auto_fixable,
            fix_payload=payload
        )

    def get_summary(self) -> Dict[str, int]:
        summary = {s.value: 0 for s in Severity}
        for issue in self.issues:
            summary[issue.severity.value] += 1
        return summary

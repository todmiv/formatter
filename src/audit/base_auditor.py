"""
BaseAuditor - базовый класс для всех аудиторов.
Определяет общий интерфейс и утилиты для проверки элементов документа.
"""

import re
import logging
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
from enum import Enum
import docx
from docx.shared import Twips, Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from src.core.config_loader import ConfigLoader
from src.core.docx_utils import (
    get_xml_attr_twips, get_paragraph_indent_twips, get_paragraph_spacing_twips,
    get_paragraph_alignment, get_font_properties, normalize_color,
    has_direct_paragraph_formatting, has_direct_run_formatting,
)

logger = logging.getLogger(__name__)


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


class BaseAuditor:
    """Базовый класс для всех аудиторов."""
    
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
        
        # Устанавливаем допуски
        self.FONT_SIZE_TOLERANCE_TWIPS = font_size_tolerance_twips if font_size_tolerance_twips is not None else self.FONT_SIZE_TOLERANCE_TWIPS
        self.INDENT_TOLERANCE_TWIPS = indent_tolerance_twips if indent_tolerance_twips is not None else self.INDENT_TOLERANCE_TWIPS
        self.SPACING_TOLERANCE_TWIPS = spacing_tolerance_twips if spacing_tolerance_twips is not None else self.SPACING_TOLERANCE_TWIPS
    
    def _update_progress(self, stage: str, progress: int, total: int, message: str = ""):
        """Вызывает колбэк прогресса, если он установлен."""
        if self.progress_callback:
            self.progress_callback(stage, progress, total, message)
    
    def _create_issue(self, severity: Severity, category: str, elem_type: str, loc: Dict,
                      desc: str, current: str, expected: str, payload: Dict = None,
                      auto_fixable: bool = True) -> AuditIssue:
        """Создаёт объект проблемы с уникальным ID."""
        self.issue_counter += 1
        issue_id = f"{category}_{self.issue_counter:04d}"
        return AuditIssue(
            id=issue_id,
            severity=severity,
            category=category,
            element_type=elem_type,
            location=loc,
            description=desc,
            current_value=current,
            expected_value=expected,
            recommendation="",
            auto_fixable=auto_fixable,
            fix_payload=payload or {}
        )
    
    def _detect_target_style(self, para) -> str:
        """
        Определяет целевой стиль для параграфа.
        ГАРАНТИРУЕТ возврат имени стиля. Никогда не возвращает None.
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
        if "Normal" in self.config.styles:
            return "Normal"
        if "Обычный" in self.config.styles:
            return "Обычный"
            
        return current_style_name
    
    def _check_font(self, para, font_config: Dict, target_style_name: str,
                    index: int, page_est: int) -> List[AuditIssue]:
        """Проверяет свойства шрифта параграфа. Пропускает свойства с прямым форматированием."""
        issues = []
        actual_font = get_font_properties(para)

        # Размер шрифта — проверяем только если нет прямого форматирования
        if 'size' in font_config and not has_direct_run_formatting(para, 'font_size'):
            expected_size_pt = font_config['size']
            expected_size_twips = Pt(expected_size_pt).twips
            actual_size_twips = actual_font.get('size_twips', 0)
            if actual_size_twips == 0:
                issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="FONT",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Не удалось определить размер шрифта",
                    current="Неизвестно",
                    expected=f"{expected_size_pt} пт",
                    payload={'property': 'size', 'expected_pt': expected_size_pt},
                    auto_fixable=False
                ))
            elif abs(actual_size_twips - expected_size_twips) > self.FONT_SIZE_TOLERANCE_TWIPS:
                actual_size_pt = actual_font.get('size_pt', 0)
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="FONT",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Размер шрифта не соответствует ГОСТ",
                    current=f"{actual_size_pt:.1f} пт",
                    expected=f"{expected_size_pt} пт",
                    payload={'property': 'size', 'expected_pt': expected_size_pt, 'actual_pt': actual_size_pt}
                ))

        # Имя шрифта — пропускаем если задано прямое форматирование
        if 'name' in font_config and not has_direct_run_formatting(para, 'font_name'):
            expected_font = font_config['name']
            actual_font_name = actual_font.get('name', '')
            if actual_font_name and actual_font_name.lower() != expected_font.lower():
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="FONT",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Название шрифта не соответствует ГОСТ",
                    current=actual_font_name,
                    expected=expected_font,
                    payload={'property': 'name', 'expected': expected_font, 'actual': actual_font_name}
                ))

        # Цвет — пропускаем если задано прямое форматирование
        if 'color' in font_config and not has_direct_run_formatting(para, 'color'):
            expected_color = normalize_color(font_config['color'])
            actual_color = actual_font.get('color', '')
            if actual_color and actual_color.lower() != expected_color.lower():
                issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="FONT",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Цвет шрифта не соответствует ГОСТ",
                    current=actual_color,
                    expected=expected_color,
                    payload={'property': 'color', 'expected': expected_color, 'actual': actual_color}
                ))

        # Полужирность — пропускаем если задано прямое форматирование
        if 'bold' in font_config and not has_direct_run_formatting(para, 'bold'):
            expected_bold = font_config['bold']
            actual_bold = actual_font.get('bold', False)
            if actual_bold != expected_bold:
                issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="FONT",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Насыщенность шрифта не соответствует ГОСТ",
                    current="Полужирный" if actual_bold else "Обычный",
                    expected="Полужирный" if expected_bold else "Обычный",
                    payload={'property': 'bold', 'expected': expected_bold, 'actual': actual_bold}
                ))

        # Курсив — пропускаем если задано прямое форматирование
        if 'italic' in font_config and not has_direct_run_formatting(para, 'italic'):
            expected_italic = font_config['italic']
            actual_italic = actual_font.get('italic', False)
            if actual_italic != expected_italic:
                issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="FONT",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Наклон шрифта не соответствует ГОСТ",
                    current="Курсив" if actual_italic else "Прямой",
                    expected="Курсив" if expected_italic else "Прямой",
                    payload={'property': 'italic', 'expected': expected_italic, 'actual': actual_italic}
                ))

        return issues
    
    def _check_paragraph_formatting(self, para, para_config: Dict, target_style_name: str,
                                    index: int, page_est: int) -> List[AuditIssue]:
        """Проверяет форматирование абзаца. Пропускает свойства с прямым форматированием."""
        issues = []

        # Выравнивание — пропускаем если задано прямое форматирование
        if 'alignment' in para_config and not has_direct_paragraph_formatting(para, 'alignment'):
            expected_align = para_config['alignment']
            actual_align = get_paragraph_alignment(para)
            if actual_align != expected_align:
                align_map = {
                    'left': 'по левому краю',
                    'center': 'по центру',
                    'right': 'по правому краю',
                    'justify': 'по ширине'
                }
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="PARAGRAPH",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Выравнивание абзаца не соответствует ГОСТ",
                    current=align_map.get(actual_align, actual_align),
                    expected=align_map.get(expected_align, expected_align),
                    payload={'property': 'alignment', 'expected': expected_align, 'actual': actual_align}
                ))

        # Отступ красной строки — пропускаем если задано прямое форматирование
        if 'first_line_indent_cm' in para_config and not has_direct_paragraph_formatting(para, 'first_line'):
            expected_cm = para_config['first_line_indent_cm']
            expected_twips = Cm(expected_cm).twips
            actual = get_paragraph_indent_twips(para)
            actual_twips = actual.get('first_line', 0) if actual else 0
            if abs(actual_twips - expected_twips) > self.INDENT_TOLERANCE_TWIPS:
                actual_cm = actual_twips / Cm(1).twips
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="PARAGRAPH",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Отступ красной строки не соответствует ГОСТ",
                    current=f"{actual_cm:.2f} см",
                    expected=f"{expected_cm:.2f} см",
                    payload={'property': 'first_line_indent', 'expected_twips': expected_twips, 'actual_twips': actual_twips}
                ))

        # Левый отступ — пропускаем если задано прямое форматирование
        if 'left_indent_cm' in para_config and not has_direct_paragraph_formatting(para, 'left_indent'):
            expected_cm = para_config['left_indent_cm']
            expected_twips = Cm(expected_cm).twips
            actual = get_paragraph_indent_twips(para)
            actual_twips = actual.get('left', 0) if actual else 0
            if abs(actual_twips - expected_twips) > self.INDENT_TOLERANCE_TWIPS:
                actual_cm = actual_twips / Cm(1).twips
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="PARAGRAPH",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Левый отступ абзаца не соответствует ГОСТ",
                    current=f"{actual_cm:.2f} см",
                    expected=f"{expected_cm:.2f} см",
                    payload={'property': 'left_indent', 'expected_twips': expected_twips, 'actual_twips': actual_twips}
                ))

        # Интервалы — пропускаем если задано прямое форматирование
        if 'space_before_pt' in para_config and not has_direct_paragraph_formatting(para, 'spacing'):
            expected_pt = para_config['space_before_pt']
            expected_twips = Pt(expected_pt).twips
            actual = get_paragraph_spacing_twips(para)
            actual_twips = actual.get('before', 0) if actual else 0
            if abs(actual_twips - expected_twips) > self.SPACING_TOLERANCE_TWIPS:
                actual_pt = actual_twips / Pt(1).twips
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="PARAGRAPH",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Интервал перед абзацем не соответствует ГОСТ",
                    current=f"{actual_pt:.1f} пт",
                    expected=f"{expected_pt:.1f} пт",
                    payload={'property': 'spacing_before', 'expected_twips': expected_twips, 'actual_twips': actual_twips}
                ))

        if 'space_after_pt' in para_config and not has_direct_paragraph_formatting(para, 'spacing'):
            expected_pt = para_config['space_after_pt']
            expected_twips = Pt(expected_pt).twips
            actual = get_paragraph_spacing_twips(para)
            actual_twips = actual.get('after', 0) if actual else 0
            if abs(actual_twips - expected_twips) > self.SPACING_TOLERANCE_TWIPS:
                actual_pt = actual_twips / Pt(1).twips
                issues.append(self._create_issue(
                    severity=Severity.CRITICAL,
                    category="PARAGRAPH",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Интервал после абзаца не соответствует ГОСТ",
                    current=f"{actual_pt:.1f} пт",
                    expected=f"{expected_pt:.1f} пт",
                    payload={'property': 'spacing_after', 'expected_twips': expected_twips, 'actual_twips': actual_twips}
                ))

        return issues
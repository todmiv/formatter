"""
ParagraphAuditor - аудитор для проверки параграфов документа.
"""

import logging
from typing import List
import docx

from src.audit.base_auditor import BaseAuditor, AuditIssue, Severity

logger = logging.getLogger(__name__)


class ParagraphAuditor(BaseAuditor):
    """Аудитор для проверки параграфов."""
    
    def audit_paragraphs(self, doc: docx.Document) -> List[AuditIssue]:
        """
        Проверяет все параграфы в документе.
        
        :param doc: объект docx.Document
        :return: список найденных проблем
        """
        issues = []
        total_paragraphs = len(doc.paragraphs)
        lines_per_page_estimate = 25
        current_line = 0
        
        for idx, para in enumerate(doc.paragraphs):
            page_est = (current_line // lines_per_page_estimate) + 1
            
            try:
                para_issues = self._audit_paragraph(para, idx, page_est)
                issues.extend(para_issues)
            except Exception as e:
                logger.error(f"Ошибка при аудите параграфа {idx}: {e}", exc_info=True)
                issues.append(self._create_issue(
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
            
            # Обновляем прогресс
            if idx % 10 == 0 or idx == total_paragraphs - 1:
                self._update_progress('paragraphs', idx + 1, total_paragraphs,
                                      f"Проверено параграфов: {idx + 1}/{total_paragraphs}")
        
        return issues
    
    def _audit_paragraph(self, para, index: int, page_est: int) -> List[AuditIssue]:
        """Проверяет отдельный параграф."""
        local_issues = []
        
        target_style_name = self._detect_target_style(para)
        style_config = self.config.get_style_config(target_style_name)
        
        if not style_config or not style_config.get('enabled', True):
            return local_issues
        
        # Проверка свойств шрифта
        if 'font' in style_config:
            font_issues = self._check_font(para, style_config['font'], target_style_name, index, page_est)
            local_issues.extend(font_issues)
        
        # Проверка форматирования абзаца
        if 'paragraph' in style_config:
            para_issues = self._check_paragraph_formatting(para, style_config['paragraph'], target_style_name, index, page_est)
            local_issues.extend(para_issues)
        
        # Проверка именования стиля
        actual_style_name = para.style.name
        if actual_style_name != target_style_name:
            if target_style_name in ["Normal", "Обычный"]:
                local_issues.append(self._create_issue(
                    severity=Severity.WARNING,
                    category="STYLE",
                    elem_type=target_style_name,
                    loc={'index': index, 'page': page_est},
                    desc="Некорректное имя стиля для основного текста",
                    current=actual_style_name,
                    expected=target_style_name,
                    payload={'property': 'style_name', 'expected': target_style_name, 'actual': actual_style_name}
                ))
        
        return local_issues
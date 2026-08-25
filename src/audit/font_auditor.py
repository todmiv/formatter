"""
FontAuditor - аудитор для проверки свойств шрифта.
"""

from src.audit.base_auditor import BaseAuditor, AuditIssue, Severity


class FontAuditor(BaseAuditor):
    """Аудитор для проверки свойств шрифта."""
    
    def audit_fonts(self, doc):
        """Проверяет свойства шрифта во всём документе."""
        # Реализация будет перенесена из audit_engine.py
        return []
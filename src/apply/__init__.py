"""
Пакет apply содержит модули для применения исправлений к документам.

Подмодули:
- base_applier: базовый класс Applier
- font_applier: применение исправлений шрифтов
- paragraph_applier: применение исправлений абзацев
- table_applier: применение исправлений таблиц
- header_footer_applier: применение исправлений колонтитулов
- apply_orchestrator: координатор ApplyEngine
"""

from .base_applier import BaseApplier
from .font_applier import FontApplier
from .paragraph_applier import ParagraphApplier
from .table_applier import TableApplier
from .header_footer_applier import HeaderFooterApplier
from .apply_orchestrator import ApplyOrchestrator

__all__ = [
    'BaseApplier',
    'FontApplier',
    'ParagraphApplier',
    'TableApplier',
    'HeaderFooterApplier',
    'ApplyOrchestrator',
]
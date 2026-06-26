#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dataclass-модели для результатов форматирования и валидации.
Заменяют Dict[str, Any] на типизированные структуры.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional


@dataclass
class FormatResult:
    """Результат форматирования документа."""
    input: str
    output: str
    mode: str  # 'config' или 'template'
    success: bool
    duration: float = 0.0
    stats: Dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None

    @property
    def styles_applied(self) -> int:
        return self.stats.get('styles_applied', 0)

    @property
    def issues_fixed(self) -> int:
        return self.stats.get('applied', 0)

    @property
    def issues_failed(self) -> int:
        return self.stats.get('failed', 0)


@dataclass
class AuditIterationResult:
    """Результат одной итерации аудита."""
    iteration: int
    issues_found: int
    issues_fixed: int
    issues_failed: int
    duration: float = 0.0


@dataclass
class AuditFormatResult:
    """Результат форматирования с аудитом."""
    input: str
    output: str
    mode: str
    success: bool
    duration: float = 0.0
    iterations: List[AuditIterationResult] = field(default_factory=list)
    final_stats: Dict[str, Any] = field(default_factory=dict)
    content_issues: List[Dict[str, Any]] = field(default_factory=list)
    error: Optional[str] = None

    @property
    def total_iterations(self) -> int:
        return len(self.iterations)

    @property
    def remaining_issues(self) -> int:
        return self.final_stats.get('remaining_issues', 0)


@dataclass
class ValidationResult:
    """Результат валидации документа."""
    issues: List[Dict[str, Any]] = field(default_factory=list)

    @property
    def count(self) -> int:
        return len(self.issues)

    @property
    def has_critical(self) -> bool:
        return any(i.get('severity') == 'CRITICAL' for i in self.issues)

    @property
    def has_warnings(self) -> bool:
        return any(i.get('severity') == 'WARNING' for i in self.issues)


@dataclass
class TransformResult:
    """Результат трансформации содержимого."""
    transforms_applied: int = 0
    paragraphs_processed: int = 0


@dataclass
class DiffResult:
    """Результат сравнения документов."""
    missing_styles: List[str] = field(default_factory=list)
    different_styles: List[str] = field(default_factory=list)
    page_setup_diff: bool = False

    @property
    def has_differences(self) -> bool:
        return bool(self.missing_styles or self.different_styles or self.page_setup_diff)


@dataclass
class BatchResult:
    """Результат пакетной обработки."""
    results: List[FormatResult] = field(default_factory=list)

    @property
    def total(self) -> int:
        return len(self.results)

    @property
    def success_count(self) -> int:
        return sum(1 for r in self.results if r.success)

    @property
    def failed_count(self) -> int:
        return self.total - self.success_count

"""
MainController - центральный контроллер приложения, отделяющий бизнес-логику от GUI.
Отвечает за управление документами, аудит, применение исправлений, генерацию отчётов.
"""

import os
import logging
import threading
import datetime
import shutil
from typing import List, Optional, Dict, Any, Callable

import docx
from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine, AuditIssue, Severity
from src.apply.apply_orchestrator import ApplyOrchestrator
from src.core.report_manager import ReportManager
from md_to_docx import RI2013Converter
from src.core.highlight_engine import HighlightEngine
from src.core.document_model import DocumentModel
from src.core.optimized_application import OptimizedApplyEngine

logger = logging.getLogger(__name__)


class MainController:
    def __init__(self, config_path: str = "configs/active/config.yaml"):
        self.config_path = config_path
        self.config_loader: Optional[ConfigLoader] = None
        self.audit_engine: Optional[AuditEngine] = None
        self.apply_engine: Optional[ApplyOrchestrator] = None
        self.report_manager: Optional[ReportManager] = None
        self.highlight_engine = HighlightEngine

        # Состояние
        self.doc_path: Optional[str] = None
        self.document_model: Optional[DocumentModel] = None
        self.current_issues: List[AuditIssue] = []
        self.last_fixed_path: Optional[str] = None

        # Допуски (значения по умолчанию в twips)
        self.font_tolerance = 10   # twips
        self.indent_tolerance = 20
        self.spacing_tolerance = 20

        # Настройки игнорирования ошибок
        self.ignored_categories: List[str] = []  # Категории для игнорирования (FONT, PARAGRAPH, STYLE, TABLE, HEADER_FOOTER, FIGURE, PAGE)
        self.ignored_severities: List[Severity] = []  # Уровни важности для игнорирования

        # Прогресс-колбэки
        self.progress_callback: Optional[Callable] = None
        self.status_callback: Optional[Callable] = None
        self.log_callback: Optional[Callable] = None

        # Блокировка для потокобезопасного доступа к состоянию
        self._lock = threading.Lock()

        self._load_config()

    def _load_config(self):
        """Загружает конфигурацию из файла."""
        if os.path.exists(self.config_path):
            try:
                self.config_loader = ConfigLoader(self.config_path)
                self.config_loader.load()
                self._log("Конфигурация загружена успешно.")
            except Exception as e:
                self._log(f"Ошибка загрузки конфига: {e}", level=logging.ERROR)
                raise
        else:
            self._log(f"Файл конфигурации не найден. Используйте стандартный или выберите другой.", level=logging.WARNING)

    def load_config(self, config_path: str):
        """Загружает конфигурацию из указанного файла."""
        self.config_path = config_path
        self._load_config()

    def set_callbacks(self, progress_cb=None, status_cb=None, log_cb=None):
        """Устанавливает колбэки для обновления UI."""
        self.progress_callback = progress_cb
        self.status_callback = status_cb
        self.log_callback = log_cb

    def _log(self, message: str, level=logging.INFO):
        """Логирует сообщение и передаёт в UI, если есть колбэк."""
        logger.log(level, message)
        if self.log_callback:
            self.log_callback(message)

    def _update_status(self, message: str):
        """Обновляет статусную строку UI."""
        if self.status_callback:
            self.status_callback(message)

    def _update_progress(self, stage: str, progress: int, total: int, message: str = ""):
        """Обновляет прогресс аудита."""
        if self.progress_callback:
            self.progress_callback(stage, progress, total, message)

    def set_document(self, doc_path: str):
        """Устанавливает текущий документ для работы."""
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"Файл не найден: {doc_path}")
        with self._lock:
            self.doc_path = doc_path
            self.document_model = DocumentModel(doc_path)
            self.current_issues = []
            self.last_fixed_path = None
        self._log(f"Документ установлен: {os.path.basename(doc_path)}")

    def get_document_structure(self) -> List[Dict]:
        """Возвращает структуру документа для отображения в дереве."""
        if not self.document_model:
            return []
        try:
            structure = []
            # Упрощённая реализация - можно расширить
            for idx, para in enumerate(self.document_model.paragraphs()):
                if para.text.strip():
                    structure.append({
                        'index': idx,
                        'text': para.text[:50],
                        'style': para.style.name,
                        'type': 'paragraph'
                    })
            return structure
        except Exception as e:
            self._log(f"Ошибка получения структуры документа: {e}", level=logging.ERROR)
            return []

    def start_audit(self) -> List[AuditIssue]:
        """Запускает аудит документа в текущем потоке (блокирующий)."""
        from src.core.formatter_errors import FormatterError, ErrorKind

        if not self.doc_path:
            raise FormatterError(ErrorKind.DOCX_OPEN_ERROR, detail="документ не выбран. Используйте set_document().")
        if not self.config_loader:
            raise FormatterError(ErrorKind.CONFIG_NOT_FOUND, detail="конфигурация не загружена. Используйте load_config().")

        self._log("Запуск аудита документа...")
        self.audit_engine = AuditEngine(
            self.config_loader,
            font_size_tolerance_twips=self.font_tolerance,
            indent_tolerance_twips=self.indent_tolerance,
            spacing_tolerance_twips=self.spacing_tolerance,
            progress_callback=self._update_progress
        )
        issues = self.audit_engine.scan_document(self.doc_path)
        # Применяем фильтры игнорирования
        filtered_issues = self._apply_ignore_filters(issues)
        self.current_issues = filtered_issues
        self.report_manager = ReportManager(filtered_issues, self.doc_path, self.config_path)
        self._log(f"Аудит завершён. Найдено проблем: {len(filtered_issues)} (игнорировано {len(issues) - len(filtered_issues)})")
        return filtered_issues

    def start_audit_async(self, on_complete: Callable[[List[AuditIssue]], None], on_error: Callable[[Exception], None] = None):
        """Запускает аудит в отдельном потоке."""
        def audit_thread():
            try:
                issues = self.start_audit()
                if on_complete:
                    on_complete(issues)
            except Exception as e:
                self._log(f"Ошибка аудита: {e}", level=logging.ERROR)
                if on_error:
                    on_error(e)

        thread = threading.Thread(target=audit_thread, daemon=True)
        thread.start()

    def apply_fixes(self, issue_ids: List[str]) -> Dict[str, int]:
        """Применяет исправления к выбранным проблемам."""
        return self.apply_fixes_with_options(issue_ids, apply_styles=True, apply_direct_overrides=True, clear_direct_formatting=False)

    def apply_fixes_with_options(self, issue_ids: List[str], apply_styles: bool = True,
                                  apply_direct_overrides: bool = True,
                                  clear_direct_formatting: bool = False,
                                  iterative: bool = False, max_iterations: int = 3,
                                  optimized: bool = False) -> Dict[str, int]:
        """
        Применяет исправления к выбранным проблемам с указанными опциями форматирования.

        :param issue_ids: Список идентификаторов проблем для исправления.
        :param apply_styles: Применять ли стили к параграфам.
        :param apply_direct_overrides: Применять ли прямое форматирование из конфига.
        :param clear_direct_formatting: Очищать ли прямое форматирование перед применением стилей.
        :param iterative: Использовать ли итеративное применение.
        :param max_iterations: Максимальное количество итераций (если iterative=True).
        :param optimized: Использовать ли оптимизированный алгоритм с анализом зависимостей.
        :return: Статистика { 'applied': int, 'failed': int } (для итеративного режима возвращает статистику последней итерации).
        """
        from src.core.formatter_errors import FormatterError, ErrorKind

        if not self.doc_path:
            raise FormatterError(ErrorKind.DOCX_OPEN_ERROR, detail="документ не выбран. Используйте set_document().")
        if not self.config_loader:
            raise FormatterError(ErrorKind.CONFIG_NOT_FOUND, detail="конфигурация не загружена. Используйте load_config().")

        selected_issues = [issue for issue in self.current_issues if issue.id in issue_ids]
        if not selected_issues:
            return {'applied': 0, 'failed': 0}

        self._log(f"Применение исправлений для {len(selected_issues)} проблем...")
        base_name = os.path.splitext(self.doc_path)[0]
        output_path = f"{base_name}_fixed.docx"
        self.last_fixed_path = output_path

        # Используем оптимизированный алгоритм, если включен
        if optimized:
            self._log("Использование оптимизированного алгоритма с анализом зависимостей...")
            optimized_engine = OptimizedApplyEngine(self.config_loader)
            results = optimized_engine.apply_optimized(
                self.doc_path, selected_issues, output_path,
                max_iterations=max_iterations if iterative else 1,
                apply_styles=apply_styles,
                apply_direct_overrides=apply_direct_overrides,
                clear_direct_formatting=clear_direct_formatting
            )
            stats = {'applied': results['total_applied'], 'failed': results['total_failed']}
            self._log(f"Оптимизированное применение завершено. Групп: {len(results.get('group_results', []))}")
        elif iterative:
            self._log(f"Итеративное применение (максимум {max_iterations} итераций)...")
            engine = ApplyOrchestrator(self.config_loader)
            results = engine.apply_fixes_iterative(self.doc_path, selected_issues, output_path,
                                                   max_iterations=max_iterations,
                                                   apply_styles=apply_styles,
                                                   apply_direct_overrides=apply_direct_overrides,
                                                   clear_direct_formatting=clear_direct_formatting)
            # Возвращаем статистику последней итерации
            stats = results[-1] if results else {'applied': 0, 'failed': len(selected_issues)}
            self._log(f"Итеративное применение завершено. Итераций: {len(results)}")
        else:
            engine = ApplyOrchestrator(self.config_loader)
            stats = engine.apply_fixes(self.doc_path, selected_issues, output_path,
                                       apply_styles=apply_styles,
                                       apply_direct_overrides=apply_direct_overrides,
                                       clear_direct_formatting=clear_direct_formatting)
        
        applied = stats['applied']
        failed = stats['failed']

        # Удаляем исправленные проблемы из текущего списка
        self.current_issues = [issue for issue in self.current_issues if issue.id not in issue_ids]
        if self.report_manager:
            self.report_manager.exclude_issues(issue_ids)

        self._log(f"Исправления применены. Успешно: {applied}, Ошибок: {failed}")
        return stats

    def apply_styles_only(self, issue_ids: List[str]) -> Dict[str, int]:
        """
        Применяет только стили к выбранным проблемам, без прямого форматирования.
        Эквивалент кнопки "Исправить только стили".
        """
        return self.apply_fixes_with_options(issue_ids, apply_styles=True,
                                             apply_direct_overrides=False,
                                             clear_direct_formatting=False)

    def apply_all_fixes(self) -> Dict[str, int]:
        """Применяет исправления ко всем найденным проблемам."""
        issue_ids = [issue.id for issue in self.current_issues]
        return self.apply_fixes(issue_ids)

    def export_report(self, format_choice: str, file_path: str) -> bool:
        """Экспортирует отчёт в указанный формат."""
        if not self.report_manager:
            raise ValueError("Отчёт не сформирован")
        try:
            if format_choice == 'txt':
                self.report_manager.export_txt(file_path)
            elif format_choice == 'html':
                self.report_manager.export_html(file_path)
            elif format_choice == 'csv':
                self.report_manager.export_csv(file_path)
            else:
                raise ValueError(f"Неподдерживаемый формат: {format_choice}")
            self._log(f"Отчёт экспортирован в {file_path}")
            return True
        except Exception as e:
            self._log(f"Ошибка экспорта отчёта: {e}", level=logging.ERROR)
            return False

    def highlight_issues(self, output_path: Optional[str] = None) -> Optional[str]:
        """Создаёт копию документа с подсветкой проблемных элементов."""
        if not self.doc_path:
            raise ValueError("Документ не выбран")
        if not self.current_issues:
            self._log("Нет проблем для подсветки", level=logging.WARNING)
            return None

        if output_path is None:
            base_name = os.path.splitext(self.doc_path)[0]
            output_path = f"{base_name}_highlighted.docx"

        try:
            highlighted_path = HighlightEngine.highlight_issues(
                self.doc_path, self.current_issues, output_path
            )
            self._log(f"Документ с подсветкой создан: {highlighted_path}")
            return highlighted_path
        except Exception as e:
            self._log(f"Ошибка подсветки: {e}", level=logging.ERROR)
            return None

    def import_markdown(self, md_path: str, docx_output_path: str, progress_callback: Optional[Callable] = None) -> bool:
        """Конвертирует Markdown в DOCX и загружает как текущий документ.
        
        Args:
            md_path: путь к исходному Markdown файлу
            docx_output_path: путь для сохранения DOCX
            progress_callback: опциональный callback(stage, progress, total, message)
        """
        if not os.path.exists(md_path):
            raise FileNotFoundError(f"Markdown файл не найден: {md_path}")

        config_path = self.config_path if os.path.exists(self.config_path) else "configs/active/config.yaml"
        try:
            converter = RI2013Converter(config_path)
            # Передаём callback прогресса конвертации (отдельный от аудита)
            converter.convert(md_path, docx_output_path, progress_callback=progress_callback)
            self.set_document(docx_output_path)
            self._log(f"Markdown сконвертирован и загружен: {docx_output_path}")
            return True
        except Exception as e:
            self._log(f"Ошибка конвертации Markdown: {e}", level=logging.ERROR)
            return False

    def set_tolerances(self, font_twips: int, indent_twips: int, spacing_twips: int):
        """Устанавливает допуски для аудита."""
        self.font_tolerance = font_twips
        self.indent_tolerance = indent_twips
        self.spacing_tolerance = spacing_twips
        self._log(f"Допуски обновлены: шрифт={font_twips}, отступ={indent_twips}, интервал={spacing_twips}")

    def set_ignore_settings(self, categories: List[str], severities: List[Severity]):
        """Устанавливает настройки игнорирования ошибок."""
        self.ignored_categories = categories
        self.ignored_severities = severities
        self._log(f"Настройки игнорирования обновлены: категории={categories}, важности={[s.value for s in severities]}")

    def get_ignore_settings(self) -> Dict[str, Any]:
        """Возвращает текущие настройки игнорирования."""
        return {
            'categories': self.ignored_categories,
            'severities': [s.value for s in self.ignored_severities]
        }

    def _apply_ignore_filters(self, issues: List[AuditIssue]) -> List[AuditIssue]:
        """Применяет фильтры игнорирования к списку проблем."""
        if not self.ignored_categories and not self.ignored_severities:
            return issues
        filtered = []
        for issue in issues:
            if issue.category in self.ignored_categories:
                continue
            if issue.severity in self.ignored_severities:
                continue
            filtered.append(issue)
        ignored_count = len(issues) - len(filtered)
        if ignored_count > 0:
            self._log(f"Игнорировано проблем: {ignored_count} (категории: {self.ignored_categories}, важности: {[s.value for s in self.ignored_severities]})")
        return filtered

    def get_summary(self) -> Dict[str, int]:
        """Возвращает сводку по найденным проблемам."""
        if not self.current_issues:
            return {s.value: 0 for s in Severity}
        summary = {s.value: 0 for s in Severity}
        for issue in self.current_issues:
            summary[issue.severity.value] += 1
        return summary

    def create_backup(self, source_path: str, backup_dir: str = "backups") -> Optional[str]:
        """
        Создаёт резервную копию файла перед сохранением.
        
        Args:
            source_path: путь к исходному файлу
            backup_dir: имя подпапки для бэкапов (относительно директории исходного файла)
        
        Returns:
            Путь к созданной резервной копии или None в случае ошибки.
        """
        if not os.path.exists(source_path):
            self._log(f"Исходный файл не существует: {source_path}", level=logging.WARNING)
            return None
        
        try:
            source_dir = os.path.dirname(source_path)
            backup_path = os.path.join(source_dir, backup_dir)
            os.makedirs(backup_path, exist_ok=True)
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.basename(source_path)
            backup_name = f"{timestamp}_{base_name}"
            backup_full = os.path.join(backup_path, backup_name)
            
            shutil.copy2(source_path, backup_full)
            self._log(f"Создана резервная копия: {backup_full}")
            return backup_full
        except Exception as e:
            self._log(f"Ошибка создания резервной копии: {e}", level=logging.ERROR)
            return None

    def apply_fixes_with_optimization(self, issue_ids: List[str], max_iterations: int = 3,
                                     apply_styles: bool = True, apply_direct_overrides: bool = True,
                                     clear_direct_formatting: bool = False) -> Dict[str, Any]:
        """
        Применяет исправления с оптимизированным алгоритмом и валидацией.
        
        :param issue_ids: Список идентификаторов проблем для исправления.
        :param max_iterations: Максимальное количество итераций.
        :param apply_styles: Применять ли стили к параграфам.
        :param apply_direct_overrides: Применять ли прямое форматирование из конфига.
        :param clear_direct_formatting: Очищать ли прямое форматирование перед применением стилей.
        :return: Детальный отчёт с результатами применения.
        """
        if not self.doc_path:
            raise ValueError("Документ не выбран")
        if not self.config_loader:
            raise ValueError("Конфигурация не загружена")

        selected_issues = [issue for issue in self.current_issues if issue.id in issue_ids]
        if not selected_issues:
            return {'total_applied': 0, 'total_failed': 0, 'iterations': []}

        self._log(f"Оптимизированное применение исправлений для {len(selected_issues)} проблем...")
        optimized_engine = OptimizedApplyEngine(self.config_loader)
        
        base_name = os.path.splitext(self.doc_path)[0]
        output_path = f"{base_name}_optimized.docx"
        self.last_fixed_path = output_path

        # Применяем с валидацией
        results = optimized_engine.apply_with_validation(
            self.doc_path, selected_issues, output_path, max_iterations=max_iterations
        )

        # Удаляем исправленные проблемы из текущего списка
        # (в реальности нужно обновить список на основе оставшихся проблем)
        # Для простоты удаляем все выбранные проблемы
        self.current_issues = [issue for issue in self.current_issues if issue.id not in issue_ids]
        if self.report_manager:
            self.report_manager.exclude_issues(issue_ids)

        self._log(f"Оптимизированное применение завершено. Эффективность: {results['efficiency']:.1%}")
        return results

    def analyze_dependencies(self, issue_ids: List[str]) -> Dict[str, Any]:
        """
        Анализирует зависимости между выбранными проблемами.
        
        :param issue_ids: Список идентификаторов проблем для анализа.
        :return: Отчёт с анализом зависимостей.
        """
        selected_issues = [issue for issue in self.current_issues if issue.id in issue_ids]
        if not selected_issues:
            return {'total_issues': 0, 'cycles_found': 0, 'groups': []}
        
        optimized_engine = OptimizedApplyEngine(self.config_loader)
        analysis = optimized_engine.analyze_dependencies(selected_issues)
        
        self._log(f"Анализ зависимостей: {analysis['group_count']} групп, {analysis['cycles_found']} циклов")
        return analysis

    def apply_all_fixes_with_optimization(self, max_iterations: int = 3) -> Dict[str, Any]:
        """
        Применяет исправления ко всем найденным проблемам с оптимизированным алгоритмом.
        
        :param max_iterations: Максимальное количество итераций.
        :return: Детальный отчёт с результатами.
        """
        issue_ids = [issue.id for issue in self.current_issues]
        return self.apply_fixes_with_optimization(issue_ids, max_iterations=max_iterations)
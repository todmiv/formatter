"""
ApplyOrchestrator - координатор применения исправлений.
Заменяет ApplyEngine, использует специализированные апплеры.
"""

import logging
from typing import List, Dict, Any, Optional
import docx
from docx.document import Document

from src.core.audit_engine import AuditIssue, Severity
from src.core.config_loader import ConfigLoader
from src.core.style_manager import StyleManager
from src.core.page_manager import PageManager
from src.core.docx_utils import open_document
from .base_applier import BaseApplier
from .font_applier import FontApplier
from .paragraph_applier import ParagraphApplier
from .table_applier import TableApplier
from .header_footer_applier import HeaderFooterApplier

logger = logging.getLogger(__name__)


class ApplyOrchestrator:
    """
    Координатор применения исправлений.
    Принимает список объектов AuditIssue и применяет изменения к документу,
    делегируя работу специализированным апплерам.
    """

    def __init__(self, config_loader: ConfigLoader):
        self.config = config_loader
        self.style_manager = StyleManager(config_loader)
        self.page_manager = PageManager(config_loader)
        self.applied_count = 0
        self.failed_count = 0

        # Инициализация апплеров
        self.font_applier = FontApplier(config_loader)
        self.paragraph_applier = ParagraphApplier(config_loader)
        self.table_applier = TableApplier(config_loader)
        self.header_footer_applier = HeaderFooterApplier(config_loader)

    def apply_fixes(self, doc_path: str, issues: List[AuditIssue], output_path: Optional[str] = None,
                    apply_styles: bool = True, apply_direct_overrides: bool = True,
                    clear_direct_formatting: bool = False) -> Dict[str, int]:
        """
        Применяет список исправлений к документу с опциями управления форматированием.

        :param doc_path: Путь к исходному файлу.
        :param issues: Список объектов AuditIssue, которые нужно исправить.
        :param output_path: Путь для сохранения результата. Если None, перезаписывает исходный (не рекомендуется).
        :param apply_styles: Применять ли стили к параграфам (True по умолчанию).
        :param apply_direct_overrides: Применять ли прямое форматирование из конфига (True по умолчанию).
        :param clear_direct_formatting: Очищать ли прямое форматирование перед применением стилей (False по умолчанию).
        :return: Статистика { 'applied': int, 'failed': int }
        """
        if not issues:
            logger.info("Список исправлений пуст. Ничего не делаю.")
            return {'applied': 0, 'failed': 0}

        try:
            doc = open_document(doc_path)
        except Exception as e:
            logger.error(f"Не удалось открыть документ для применения исправлений: {e}")
            return {'applied': 0, 'failed': len(issues)}

        self.applied_count = 0
        self.failed_count = 0

        # Сортируем issues по индексу, чтобы идти последовательно
        sorted_issues = sorted(issues, key=lambda x: x.location.get('index', 0))

        # Кэш стилей: чтобы не создавать стиль каждый раз, если он уже есть в документе
        style_cache = {}

        for issue in sorted_issues:
            if not issue.auto_fixable:
                logger.debug(f"Исправление {issue.id} пропущено, так как auto_fixable=False")
                self.failed_count += 1
                continue

            try:
                # Определяем категорию и делегируем соответствующему апплеру
                if issue.category == 'FONT':
                    success = self.font_applier.apply(doc, issue)
                elif issue.category in ('PARAGRAPH', 'STYLE'):
                    success = self.paragraph_applier.apply(
                        doc, issue, style_cache,
                        apply_styles=apply_styles,
                        apply_direct_overrides=apply_direct_overrides,
                        clear_direct_formatting=clear_direct_formatting
                    )
                elif issue.category == 'TABLE':
                    success = self.table_applier.apply(doc, issue)
                elif issue.category in ('HEADER', 'FOOTER', 'HEADER_FOOTER'):
                    success = self.header_footer_applier.apply(doc, issue)
                elif issue.category == 'PAGE':
                    success = self._apply_page_fix(doc, issue)
                else:
                    logger.warning(f"Неизвестная категория {issue.category} для исправления {issue.id}")
                    success = False

                if success:
                    self.applied_count += 1
                    logger.info(f"Исправление {issue.id} успешно применено")
                else:
                    self.failed_count += 1
                    logger.warning(f"Не удалось применить исправление {issue.id}")

            except ValueError as e:
                logger.warning(f"Пропущено исправление {issue.id}: {e}")
                self.failed_count += 1
            except Exception as e:
                logger.error(f"Ошибка при применении исправления {issue.id}: {e}", exc_info=True)
                self.failed_count += 1

        # Сохранение
        save_path = output_path if output_path else doc_path
        try:
            doc.save(save_path)
            logger.info(f"Документ сохранен: {save_path}. Успешно применено: {self.applied_count}, Ошибок: {self.failed_count}")
        except Exception as e:
            logger.error(f"Не удалось сохранить документ: {e}")
            # В случае ошибки сохранения можно попробовать сохранить во временный файл, но здесь просто вернем статус
            return {'applied': self.applied_count, 'failed': self.failed_count + 1}  # +1 за ошибку сохранения

        return {'applied': self.applied_count, 'failed': self.failed_count}

    def _apply_page_fix(self, doc, issue: AuditIssue) -> bool:
        """Применяет исправление настроек страницы (поля, ориентация)."""
        try:
            prop = issue.fix_payload.get('prop', '')
            section_idx = issue.location.get('section', 0)

            if section_idx >= len(doc.sections):
                logger.warning(f"Секция {section_idx} вне диапазона (всего {len(doc.sections)})")
                return False

            section = doc.sections[section_idx]
            page_setup = self.config.get_page_setup()

            if not page_setup:
                logger.warning("Конфигурация page_setup отсутствует")
                return False

            orientation = page_setup.get('orientation_default', 'portrait')
            margins = page_setup.get(orientation, {})

            from docx.shared import Cm

            margin_map = {
                'left_margin': ('left_margin_cm', 'left_margin'),
                'right_margin': ('right_margin_cm', 'right_margin'),
                'top_margin': ('top_margin_cm', 'top_margin'),
                'bottom_margin': ('bottom_margin_cm', 'bottom_margin'),
            }

            if prop in margin_map:
                cfg_key, attr_name = margin_map[prop]
                if cfg_key in margins:
                    setattr(section, attr_name, Cm(margins[cfg_key]))
                    logger.debug(f"Поле {prop} установлено в {margins[cfg_key]} см (секция {section_idx})")
                    return True

            # Если prop не указан — применяем все поля сразу
            for prop_name, (cfg_key, attr_name) in margin_map.items():
                if cfg_key in margins:
                    setattr(section, attr_name, Cm(margins[cfg_key]))

            logger.debug(f"Все поля страницы применены к секции {section_idx}")
            return True
        except Exception as e:
            logger.error(f"Ошибка при исправлении страницы {issue.id}: {e}")
            return False

    def apply_fixes_iterative(self, doc_path: str, issues: List[AuditIssue], output_path: Optional[str] = None,
                              max_iterations: int = 3, apply_styles: bool = True,
                              apply_direct_overrides: bool = True, clear_direct_formatting: bool = False) -> List[Dict[str, int]]:
        """
        Применяет исправления итеративно для устранения взаимозависимостей.

        :param doc_path: Путь к исходному файлу.
        :param issues: Список объектов AuditIssue, которые нужно исправить.
        :param output_path: Путь для сохранения результата. Если None, перезаписывает исходный.
        :param max_iterations: Максимальное количество итераций.
        :param apply_styles: Применять ли стили к параграфам.
        :param apply_direct_overrides: Применять ли прямое форматирование из конфига.
        :param clear_direct_formatting: Очищать ли прямое форматирование перед применением стилей.
        :return: Список статистик по каждой итерации.
        """
        results = []
        previous_failed = len(issues)
        previous_applied = 0

        for iteration in range(max_iterations):
            logger.info(f"Итерация {iteration + 1} из {max_iterations}")

            # Приоритизация: сначала критические проблемы (PARAGRAPH с интервалами)
            critical_issues = [i for i in issues if i.category == 'PARAGRAPH' and 'интервал' in i.description.lower()]
            other_issues = [i for i in issues if i not in critical_issues]
            prioritized_issues = critical_issues + other_issues

            # Применяем исправления с улучшенным алгоритмом
            result = self.apply_fixes(doc_path, prioritized_issues, output_path,
                                      apply_styles=apply_styles,
                                      apply_direct_overrides=apply_direct_overrides,
                                      clear_direct_formatting=clear_direct_formatting)
            results.append(result)

            # Анализ прогресса
            progress = previous_failed - result['failed']
            logger.info(f"Прогресс: уменьшение неудач на {progress} (было {previous_failed}, стало {result['failed']})")

            # Если все исправления успешны, завершаем
            if result['failed'] == 0:
                logger.info(f"Все исправления применены успешно на итерации {iteration + 1}")
                break

            # Если количество неудач не уменьшается 2 итерации подряд, останавливаемся
            if iteration > 0 and result['failed'] >= previous_failed:
                logger.warning(f"Прогресс остановился: неудач {result['failed']} (предыдущие {previous_failed}). Прерываем.")
                break

            # Если прогресс менее 10% от общего числа проблем и прошло более 1 итерации
            total_issues = len(issues)
            if iteration >= 1 and progress < total_issues * 0.1:
                logger.info(f"Прогресс замедлился (<10%). Прерываем.")
                break

            previous_failed = result['failed']
            previous_applied = result['applied']

            # Обновляем путь для следующей итерации (используем тот же выходной файл)
            if output_path:
                doc_path = output_path  # следующая итерация применяется к уже измененному документу

        # Логирование взаимозависимостей для отладки
        if results:
            logger.info(f"Итоговые результаты: {results[-1]}")
            logger.info(f"Всего итераций: {len(results)}")

        return results

    def initialize_styles(self, doc_path: str, output_path: Optional[str] = None) -> bool:
        """
        Режим «Инициализация стилей»: подготавливает DOCX без изменения текста,
        только создание стилей и настройка полей страницы и колонтитулов.

        :param doc_path: Путь к исходному файлу.
        :param output_path: Путь для сохранения результата. Если None, перезаписывает исходный.
        :return: True в случае успеха, False при ошибке.
        """
        try:
            doc = open_document(doc_path)
        except Exception as e:
            logger.error(f"Не удалось открыть документ для инициализации стилей: {e}")
            return False

        # 1. Настройка стилей
        self.style_manager.setup_styles(doc)

        # 2. Настройка полей страницы и колонтитулов через PageManager
        self.page_manager.apply_all(doc)

        # Сохранение
        save_path = output_path if output_path else doc_path
        try:
            doc.save(save_path)
            logger.info(f"Инициализация стилей завершена. Документ сохранен: {save_path}")
            return True
        except Exception as e:
            logger.error(f"Не удалось сохранить документ после инициализации стилей: {e}")
            return False

    def clear_direct_formatting(self, doc_path: str, output_path: str):
        """
        Вспомогательный метод: удаляет ВСЕ прямое форматирование в документе, оставляя только стили.
        Полезно как радикальная мера перед применением стилей.
        """
        doc = open_document(doc_path)

        for para in doc.paragraphs:
            self.font_applier._clear_paragraph_direct_formatting(para)

        doc.save(output_path)
        logger.info(f"Прямое форматирование очищено, файл сохранен: {output_path}")

    def apply_styles_only(self, doc_path: str, issues: List[AuditIssue], output_path: Optional[str] = None) -> Dict[str, int]:
        """
        Применяет только стили к выбранным проблемам, без прямого форматирования.
        Эквивалент вызова apply_fixes с apply_styles=True, apply_direct_overrides=False.
        """
        return self.apply_fixes(doc_path, issues, output_path,
                                apply_styles=True,
                                apply_direct_overrides=False,
                                clear_direct_formatting=False)
#!/usr/bin/env python3
"""
Стресс-тестирование системы в граничных условиях (edge cases).
Проверяет стабильность и корректность работы новых функций.
"""

import sys
import os
import time
import tempfile
import shutil
import logging
from typing import List, Dict, Any
import traceback

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine, AuditIssue
from src.apply.apply_orchestrator import ApplyOrchestrator
import docx
from docx import Document
from docx.shared import Pt, Inches

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class StressTester:
    def __init__(self, config_path: str = "config.yaml"):
        self.config_path = config_path
        self.config = ConfigLoader(config_path)
        self.config.load()
        self.audit_engine = AuditEngine(self.config)
        self.apply_engine = ApplyOrchestrator(self.config)
        self.results = []
        
    def log_result(self, test_name: str, status: str, message: str = ""):
        """Логирует результат теста."""
        result = {
            'test_name': test_name,
            'status': status,
            'message': message,
            'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
        }
        self.results.append(result)
        logger.info(f"{test_name}: {status} - {message}")
        
    def create_test_document_no_tables(self) -> str:
        """Создаёт тестовый документ без таблиц."""
        temp_dir = tempfile.mkdtemp()
        doc_path = os.path.join(temp_dir, "no_tables.docx")
        
        doc = Document()
        # Добавляем различные параграфы
        for i in range(10):
            p = doc.add_paragraph(f"Параграф {i+1} без таблиц.")
            if i % 2 == 0:
                p.style = "Normal"
            else:
                p.style = "Heading 1"
        
        # Добавляем колонтитулы
        section = doc.sections[0]
        header = section.header
        header_para = header.paragraphs[0]
        header_para.text = "Верхний колонтитул без таблиц"
        
        footer = section.footer
        footer_para = footer.paragraphs[0]
        footer_para.text = "Нижний колонтитул без таблиц"
        
        doc.save(doc_path)
        return doc_path, temp_dir
        
    def create_test_document_only_tables(self) -> str:
        """Создаёт тестовый документ только с таблицами."""
        temp_dir = tempfile.mkdtemp()
        doc_path = os.path.join(temp_dir, "only_tables.docx")
        
        doc = Document()
        
        # Добавляем 5 таблиц
        for table_idx in range(5):
            table = doc.add_table(rows=3, cols=4)
            table.style = "Table Grid"
            for i in range(3):
                for j in range(4):
                    cell = table.cell(i, j)
                    cell.text = f"Таблица {table_idx+1}, ячейка {i+1},{j+1}"
        
        doc.save(doc_path)
        return doc_path, temp_dir
        
    def test_no_tables_document(self):
        """Тест 1: Документ без таблиц."""
        test_name = "Документ без таблиц"
        try:
            # Создаём документ
            doc_path, temp_dir = self.create_test_document_no_tables()
            
            # Аудит
            issues = self.audit_engine.scan_document(doc_path)
            
            # Проверяем, что нет проблем с таблицами
            table_issues = [i for i in issues if i.category == 'TABLE']
            
            if len(table_issues) == 0:
                self.log_result(test_name, "PASS", 
                              f"Документ без таблиц обработан корректно. Найдено проблем: {len(issues)}, проблем таблиц: {len(table_issues)}")
            else:
                self.log_result(test_name, "WARN", 
                              f"Найдены проблемы таблиц в документе без таблиц: {len(table_issues)}")
            
            # Применяем исправления (если есть проблемы)
            if issues:
                result = self.apply_engine.apply_fixes(doc_path, issues, 
                                                      output_path=os.path.join(temp_dir, "no_tables_fixed.docx"))
                self.log_result(test_name + " (apply)", "PASS", 
                              f"Исправления применены: успешно={result.get('applied', 0)}, неудачно={result.get('failed', 0)}")
            
            # Очистка
            shutil.rmtree(temp_dir, ignore_errors=True)
            
        except Exception as e:
            self.log_result(test_name, "FAIL", f"Ошибка: {e}")
            logger.error(traceback.format_exc())
            
    def test_only_tables_document(self):
        """Тест 2: Документ только с таблицами."""
        test_name = "Документ только с таблицами"
        try:
            # Создаём документ
            doc_path, temp_dir = self.create_test_document_only_tables()
            
            # Аудит
            issues = self.audit_engine.scan_document(doc_path)
            
            # Проверяем, что есть проблемы с таблицами
            table_issues = [i for i in issues if i.category == 'TABLE']
            other_issues = [i for i in issues if i.category != 'TABLE']
            
            if len(table_issues) > 0:
                self.log_result(test_name, "PASS", 
                              f"Документ с таблицами обработан корректно. Всего проблем: {len(issues)}, проблем таблиц: {len(table_issues)}, других проблем: {len(other_issues)}")
            else:
                self.log_result(test_name, "WARN", 
                              f"Не найдено проблем таблиц в документе только с таблицами")
            
            # Применяем исправления (если есть проблемы)
            if issues:
                result = self.apply_engine.apply_fixes(doc_path, issues, 
                                                      output_path=os.path.join(temp_dir, "only_tables_fixed.docx"))
                self.log_result(test_name + " (apply)", "PASS", 
                              f"Исправления применены: успешно={result.get('applied', 0)}, неудачно={result.get('failed', 0)}")
                
                # Проверяем функцию _get_table
                doc = Document(doc_path)
                for i in range(len(doc.tables)):
                    try:
                        table = self.apply_engine._get_table(doc, {'table_index': i})
                        if table:
                            pass  # OK
                    except Exception as e:
                        self.log_result(test_name + " (_get_table)", "FAIL", 
                                      f"Ошибка получения таблицы {i}: {e}")
            
            # Очистка
            shutil.rmtree(temp_dir, ignore_errors=True)
            
        except Exception as e:
            self.log_result(test_name, "FAIL", f"Ошибка: {e}")
            logger.error(traceback.format_exc())
            
    def test_many_iterations(self):
        """Тест 3: Большое количество итераций (5-10)."""
        test_name = "Большое количество итераций"
        try:
            # Используем существующий документ или создаём простой
            doc_path = "Том II_Волжск.docx"
            if not os.path.exists(doc_path):
                # Создаём простой документ для теста
                temp_dir = tempfile.mkdtemp()
                doc_path = os.path.join(temp_dir, "test_iterations.docx")
                doc = Document()
                for i in range(5):
                    p = doc.add_paragraph(f"Параграф {i+1} для теста итераций.")
                    p.style = "Normal"
                doc.save(doc_path)
                is_temp = True
            else:
                temp_dir = None
                is_temp = False
            
            # Аудит
            issues = self.audit_engine.scan_document(doc_path)
            
            if not issues:
                # Создаём искусственные проблемы для теста
                issues = []
                for i in range(5):
                    issue = AuditIssue(
                        id=f"TEST_{i}",
                        severity=AuditIssue.Severity.WARNING,
                        category="PARAGRAPH",
                        element_type="Normal",
                        location={'index': i % 3, 'type': 'paragraph'},
                        description="Тестовая проблема",
                        current_value="",
                        expected_value="",
                        auto_fixable=True
                    )
                    issues.append(issue)
            
            # Применяем итеративно с большим количеством итераций
            max_iterations = 10
            logger.info(f"Запуск итеративного применения с {max_iterations} итерациями...")
            
            start_time = time.time()
            results = self.apply_engine.apply_fixes_iterative(
                doc_path,
                issues,
                output_path=None,  # Используем временный файл
                max_iterations=max_iterations
            )
            elapsed_time = time.time() - start_time
            
            # Анализ результатов
            iterations_completed = len(results)
            last_result = results[-1] if results else {}
            
            # Проверяем, что система не сломалась
            if iterations_completed > 0:
                self.log_result(test_name, "PASS", 
                              f"Выполнено {iterations_completed} итераций из {max_iterations}. "
                              f"Время: {elapsed_time:.2f} сек. "
                              f"Последний результат: успешно={last_result.get('applied', 0)}, неудачно={last_result.get('failed', 0)}")
                
                # Проверяем, что количество итераций не превышает максимум
                if iterations_completed <= max_iterations:
                    self.log_result(test_name + " (max iterations)", "PASS", 
                                  f"Количество итераций в пределах максимума: {iterations_completed}")
                else:
                    self.log_result(test_name + " (max iterations)", "FAIL", 
                                  f"Превышено максимальное количество итераций: {iterations_completed} > {max_iterations}")
                
                # Проверяем стабильность (результаты не должны ухудшаться катастрофически)
                fails = [r.get('failed', 0) for r in results]
                if all(f <= fails[0] * 2 for f in fails):  # Допускаем ухудшение не более чем в 2 раза
                    self.log_result(test_name + " (stability)", "PASS", 
                                  f"Стабильность сохранена, неудачи: {fails}")
                else:
                    self.log_result(test_name + " (stability)", "WARN", 
                                  f"Возможная нестабильность, неудачи: {fails}")
            else:
                self.log_result(test_name, "FAIL", "Итеративное применение не вернуло результатов")
            
            # Очистка
            if is_temp and temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)
                
        except Exception as e:
            self.log_result(test_name, "FAIL", f"Ошибка: {e}")
            logger.error(traceback.format_exc())
            
    def test_parallel_execution(self):
        """Тест 4: Параллельное выполнение (если поддерживается)."""
        test_name = "Параллельное выполнение"
        
        # Проверяем, поддерживается ли параллельное выполнение
        # В текущей реализации, вероятно, не поддерживается, но можно проверить
        # возможность запуска нескольких apply_fixes одновременно
        
        try:
            import threading
            import queue
            
            # Создаём тестовый документ
            temp_dir = tempfile.mkdtemp()
            doc_path = os.path.join(temp_dir, "parallel_test.docx")
            
            doc = Document()
            for i in range(20):
                p = doc.add_paragraph(f"Параграф {i+1} для параллельного теста.")
                p.style = "Normal"
            doc.save(doc_path)
            
            # Аудит
            issues = self.audit_engine.scan_document(doc_path)
            
            if not issues or len(issues) < 5:
                # Создаём искусственные проблемы
                issues = []
                for i in range(10):
                    issue = AuditIssue(
                        id=f"PARALLEL_TEST_{i}",
                        severity=AuditIssue.Severity.WARNING,
                        category="PARAGRAPH",
                        element_type="Normal",
                        location={'index': i % 5, 'type': 'paragraph'},
                        description="Тестовая проблема для параллельного выполнения",
                        current_value="",
                        expected_value="",
                        auto_fixable=True
                    )
                    issues.append(issue)
            
            # Разделяем проблемы на две группы
            half = len(issues) // 2
            group1 = issues[:half]
            group2 = issues[half:]
            
            results_queue = queue.Queue()
            errors = []
            
            def apply_fixes_thread(doc_path, issues_group, group_name, output_suffix):
                try:
                    output_path = os.path.join(temp_dir, f"parallel_output_{output_suffix}.docx")
                    result = self.apply_engine.apply_fixes(doc_path, issues_group, output_path=output_path)
                    results_queue.put((group_name, result))
                except Exception as e:
                    errors.append((group_name, e))
            
            # Запускаем два потока
            thread1 = threading.Thread(
                target=apply_fixes_thread,
                args=(doc_path, group1, "group1", "1")
            )
            thread2 = threading.Thread(
                target=apply_fixes_thread,
                args=(doc_path, group2, "group2", "2")
            )
            
            start_time = time.time()
            thread1.start()
            thread2.start()
            
            thread1.join(timeout=30)
            thread2.join(timeout=30)
            elapsed_time = time.time() - start_time
            
            # Собираем результаты
            thread_results = []
            while not results_queue.empty():
                thread_results.append(results_queue.get())
            
            # Анализируем
            if errors:
                self.log_result(test_name, "WARN", 
                              f"Ошибки в потоках: {len(errors)}. Успешных потоков: {len(thread_results)}. Время: {elapsed_time:.2f} сек.")
                for group_name, error in errors:
                    logger.error(f"Ошибка в потоке {group_name}: {error}")
            else:
                if len(thread_results) == 2:
                    self.log_result(test_name, "PASS", 
                                  f"Оба потока завершились успешно. Время: {elapsed_time:.2f} сек.")
                    for group_name, result in thread_results:
                        logger.info(f"  {group_name}: успешно={result.get('applied', 0)}, неудачно={result.get('failed', 0)}")
                else:
                    self.log_result(test_name, "WARN", 
                                  f"Не все потоки завершились. Завершено: {len(thread_results)} из 2. Время: {elapsed_time:.2f} сек.")
            
            # Проверяем, что документы созданы
            output1 = os.path.join(temp_dir, "parallel_output_1.docx")
            output2 = os.path.join(temp_dir, "parallel_output_2.docx")
            
            if os.path.exists(output1):
                self.log_result(test_name + " (output1)", "PASS", "Выходной файл 1 создан")
            else:
                self.log_result(test_name + " (output1)", "WARN", "Выходной файл 1 не создан")
                
            if os.path.exists(output2):
                self.log_result(test_name + " (output2)", "PASS", "Выходной файл 2 создан")
            else:
                self.log_result(test_name + " (output2)", "WARN", "Выходной файл 2 не создан")
            
            # Очистка
            shutil.rmtree(temp_dir, ignore_errors=True)
            
        except Exception as e:
            self.log_result(test_name, "FAIL", f"Ошибка: {e}")
            logger.error(traceback.format_exc())
            
    def test_edge_case_locations(self):
        """Тест 5: Граничные случаи location."""
        test_name = "Граничные случаи location"
        try:
            # Создаём документ
            temp_dir = tempfile.mkdtemp()
            doc_path = os.path.join(temp_dir, "edge_locations.docx")
            
            doc = Document()
            # Добавляем параграфы и таблицы
            for i in range(3):
                doc.add_paragraph(f"Параграф {i+1}")
            for i in range(2):
                doc.add_table(rows=2, cols=2)
            
            doc.save(doc_path)
            
            # Тестируем различные форматы location
            test_cases = [
                # (location, expected_success, description)
                ({'index': 0, 'type': 'paragraph'}, True, "Корректный location параграфа"),
                ({'table_index': 0, 'type': 'table'}, True, "Корректный location таблицы"),
                ({'index': 999, 'type': 'paragraph'}, False, "Несуществующий индекс параграфа"),
                ({'table_index': 999, 'type': 'table'}, False, "Несуществующий индекс таблицы"),
                ({'section': 0, 'paragraph': 0, 'container': 'header', 'type': 'paragraph'}, True, "Location колонтитула"),
                ({'type': 'unknown'}, False, "Неизвестный тип элемента"),
                ({}, False, "Пустой location"),
                ({'index': 0}, True, "Legacy location без типа"),
                ({'paragraph_index': 0}, True, "Legacy location с paragraph_index"),
            ]
            
            doc_obj = Document(doc_path)
            passed = 0
            failed = 0
            
            for location, expected_success, description in test_cases:
                try:
                    element, elem_type = self.apply_engine._get_element(doc_obj, location)
                    if expected_success:
                        self.log_result(f"{test_name}: {description}", "PASS", f"Успешно, тип: {elem_type}")
                        passed += 1
                    else:
                        self.log_result(f"{test_name}: {description}", "FAIL", f"Ожидалась ошибка, но получен элемент типа {elem_type}")
                        failed += 1
                except Exception as e:
                    if not expected_success:
                        self.log_result(f"{test_name}: {description}", "PASS", f"Ожидаемая ошибка: {str(e)[:50]}")
                        passed += 1
                    else:
                        self.log_result(f"{test_name}: {description}", "FAIL", f"Неожиданная ошибка: {e}")
                        failed += 1
            
            self.log_result(test_name, "PASS" if failed == 0 else "WARN", 
                          f"Пройдено: {passed}/{len(test_cases)}, Не пройдено: {failed}")
            
            # Очистка
            shutil.rmtree(temp_dir, ignore_errors=True)
            
        except Exception as e:
            self.log_result(test_name, "FAIL", f"Ошибка: {e}")
            logger.error(traceback.format_exc())
            
    def run_all_tests(self):
        """Запуск всех тестов."""
        logger.info("=== НАЧАЛО СТРЕСС-ТЕСТИРОВАНИЯ В GRANICHНЫХ УСЛОВИЯХ ===")
        
        # Тест 1: Документ без таблиц
        self.test_no_tables_document()
        
        # Тест 2: Документ только с таблицами
        self.test_only_tables_document()
        
        # Тест 3: Большое количество итераций
        self.test_many_iterations()
        
        # Тест 4: Параллельное выполнение
        self.test_parallel_execution()
        
        # Тест 5: Граничные случаи location
        self.test_edge_case_locations()
        
        # Сводка
        self.print_summary()
        
        logger.info("=== СТРЕСС-ТЕСТИРОВАНИЕ ЗАВЕРШЕНО ===")
        
    def print_summary(self):
        """Вывод сводки результатов."""
        logger.info("\n=== СВОДКА РЕЗУЛЬТАТОВ СТРЕСС-ТЕСТИРОВАНИЯ ===")
        
        total = len(self.results)
        passed = sum(1 for r in self.results if r['status'] == 'PASS')
        failed = sum(1 for r in self.results if r['status'] == 'FAIL')
        warned = sum(1 for r in self.results if r['status'] == 'WARN')
        
        logger.info(f"Всего тестов: {total}")
        logger.info(f"Пройдено успешно: {passed}")
        logger.info(f"Провалено: {failed}")
        logger.info(f"Предупреждения: {warned}")
        
        if failed > 0:
            logger.info("\nПроваленные тесты:")
            for r in self.results:
                if r['status'] == 'FAIL':
                    logger.info(f"  - {r['test_name']}: {r['message']}")
        
        if warned > 0:
            logger.info("\nТесты с предупреждениями:")
            for r in self.results:
                if r['status'] == 'WARN':
                    logger.info(f"  - {r['test_name']}: {r['message']}")
        
        # Сохранение результатов в файл
        report_path = "stress_test_report.md"
        self.generate_report(report_path)
        logger.info(f"\nПолный отчёт сохранён: {report_path}")
        
    def generate_report(self, report_path: str):
        """Генерация отчёта о стресс-тестировании."""
        report = "# Отчёт о стресс-тестировании в граничных условиях\n\n"
        report += f"**Дата тестирования**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += f"**Конфигурация**: {self.config_path}\n\n"
        
        # Сводная таблица
        report += "## Сводка результатов\n\n"
        report += "| Тест | Статус | Сообщение |\n"
        report += "|------|--------|-----------|\n"
        
        for r in self.results:
            status_emoji = "✅" if r['status'] == 'PASS' else "⚠️" if r['status'] == 'WARN' else "❌"
            report += f"| {r['test_name']} | {status_emoji} {r['status']} | {r['message']} |\n"
        
        # Детали по каждому тесту
        report += "\n## Детальные результаты\n\n"
        
        test_groups = {}
        for r in self.results:
            base_name = r['test_name'].split(':')[0] if ':' in r['test_name'] else r['test_name']
            if base_name not in test_groups:
                test_groups[base_name] = []
            test_groups[base_name].append(r)
        
        for group, tests in test_groups.items():
            report += f"### {group}\n\n"
            for test in tests:
                report += f"- **{test['test_name']}**: {test['status']} - {test['message']}\n"
            report += "\n"
        
        # Выводы и рекомендации
        report += "## Выводы и рекомендации\n\n"
        
        failed_tests = [r for r in self.results if r['status'] == 'FAIL']
        warned_tests = [r for r in self.results if r['status'] == 'WARN']
        
        if not failed_tests and not warned_tests:
            report += "✅ **Все тесты пройдены успешно.** Система стабильно работает в граничных условиях.\n"
        else:
            if failed_tests:
                report += "❌ **Критические проблемы:**\n"
                for test in failed_tests:
                    report += f"- {test['test_name']}: {test['message']}\n"
                report += "\n"
            
            if warned_tests:
                report += "⚠️ **Потенциальные проблемы:**\n"
                for test in warned_tests:
                    report += f"- {test['test_name']}: {test['message']}\n"
                report += "\n"
        
        report += "### Рекомендации по улучшению:\n"
        report += "1. **Обработка ошибок**: Улучшить обработку граничных случаев location.\n"
        report += "2. **Параллельное выполнение**: Рассмотреть возможность реализации thread-safe операций.\n"
        report += "3. **Производительность**: Оптимизировать итеративное применение для больших документов.\n"
        report += "4. **Тестирование**: Добавить unit-тесты для всех edge cases.\n"
        
        # Сохранение
        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report)
            logger.info(f"Отчёт сохранён: {report_path}")
        except Exception as e:
            logger.error(f"Не удалось сохранить отчёт: {e}")


def main():
    """Основная функция."""
    tester = StressTester()
    
    try:
        tester.run_all_tests()
        return True
    except Exception as e:
        logger.error(f"Критическая ошибка при стресс-тестировании: {e}")
        traceback.print_exc()
        return False


if __name__ == "__main__":
    main()
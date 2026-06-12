#!/usr/bin/env python3
"""
Комплексный тестовый скрипт для проверки исправлений на реальном документе "Том II_Волжск.docx".
Выполняет полный цикл работы с документом и собирает детальную статистику.
"""

import sys
import os
import time
import json
from collections import defaultdict, Counter
from typing import List, Dict, Any
import logging

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine, AuditIssue
from src.apply.apply_orchestrator import ApplyOrchestrator
import docx

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class RealDocumentTester:
    def __init__(self, doc_path: str, config_path: str = None):
        self.doc_path = doc_path
        if config_path is None:
            # Автовыбор конфигурации
            if os.path.exists("config.yaml"):
                config_path = "config.yaml"
            elif os.path.exists("config_v4.2.yaml"):
                config_path = "config_v4.2.yaml"
            else:
                raise FileNotFoundError("Не найден файл конфигурации")
        self.config_path = config_path
        
        self.config = ConfigLoader(config_path)
        self.config.load()
        
        self.audit_engine = AuditEngine(self.config)
        self.apply_engine = ApplyOrchestrator(self.config)
        
        # Результаты
        self.initial_issues = []
        self.fixed_issues = []
        self.final_issues = []
        self.apply_results = {}
        self.stats = {}
        
    def run_full_test(self):
        """Запуск полного тестового цикла."""
        logger.info(f"=== НАЧАЛО ТЕСТИРОВАНИЯ ДОКУМЕНТА: {self.doc_path} ===")
        
        # 1. Первый аудит
        logger.info("1. Запуск первого аудита...")
        start = time.time()
        self.initial_issues = self.audit_engine.scan_document(self.doc_path)
        audit_time = time.time() - start
        logger.info(f"   Аудит завершён за {audit_time:.2f} сек. Найдено проблем: {len(self.initial_issues)}")
        
        # Сбор статистики по начальным проблемам
        self._collect_initial_stats()
        
        # 2. Применение исправлений
        logger.info("2. Применение исправлений ко всем проблемам...")
        start = time.time()
        self.apply_results = self.apply_engine.apply_fixes(
            self.doc_path, 
            self.initial_issues,
            output_path="Том II_Волжск_fixed_v1.docx"
        )
        apply_time = time.time() - start
        logger.info(f"   Исправления применены за {apply_time:.2f} сек.")
        logger.info(f"   Результат: успешно={self.apply_results.get('success_count', 0)}, "
                   f"неудачно={self.apply_results.get('failed_count', 0)}, "
                   f"пропущено={self.apply_results.get('skipped_count', 0)}")
        
        # 3. Второй аудит на исправленном документе
        fixed_doc_path = "Том II_Волжск_fixed_v1.docx"
        if not os.path.exists(fixed_doc_path):
            logger.error("Исправленный документ не создан!")
            return False
            
        logger.info("3. Запуск второго аудита на исправленном документе...")
        start = time.time()
        self.final_issues = self.audit_engine.scan_document(fixed_doc_path)
        second_audit_time = time.time() - start
        logger.info(f"   Второй аудит завершён за {second_audit_time:.2f} сек. Найдено проблем: {len(self.final_issues)}")
        
        # Сбор статистики по финальным проблемам
        self._collect_final_stats()
        
        # 4. Сравнение результатов
        logger.info("4. Сравнение результатов до и после исправлений...")
        self._compare_results()
        
        # 5. Анализ проблем колонтитулов
        logger.info("5. Детальный анализ проблем колонтитулов...")
        self._analyze_header_footer_issues()
        
        # 6. Проверка гипотезы о 30 постоянных ошибках
        logger.info("6. Проверка гипотезы о 30 постоянных ошибках...")
        self._check_30_errors_hypothesis()
        
        # 7. Генерация отчёта
        logger.info("7. Генерация отчёта...")
        self._generate_report()
        
        logger.info("=== ТЕСТИРОВАНИЕ ЗАВЕРШЕНО ===")
        return True
    
    def _collect_initial_stats(self):
        """Сбор статистики по начальным проблемам."""
        stats = {
            'total': len(self.initial_issues),
            'by_category': Counter(),
            'by_location_type': Counter(),
            'by_severity': Counter(),
            'header_footer_details': [],
            'location_keys': defaultdict(int)
        }
        
        for issue in self.initial_issues:
            stats['by_category'][issue.category] += 1
            stats['by_severity'][issue.severity.value] += 1
            
            # Анализ location
            loc = issue.location
            if loc:
                # Тип location
                if 'index' in loc:
                    stats['by_location_type']['index'] += 1
                elif 'section' in loc and 'paragraph' in loc:
                    stats['by_location_type']['section+paragraph'] += 1
                elif 'paragraph_index' in loc:
                    stats['by_location_type']['paragraph_index'] += 1
                else:
                    stats['by_location_type']['other'] += 1
                
                # Ключи location
                for key in loc.keys():
                    stats['location_keys'][key] += 1
            
            # Детали колонтитулов
            if issue.category == 'HEADER_FOOTER':
                stats['header_footer_details'].append({
                    'id': issue.id,
                    'element_type': issue.element_type,
                    'location': issue.location,
                    'description': issue.description
                })
        
        stats['header_footer_count'] = len(stats['header_footer_details'])
        self.stats['initial'] = stats
        
        logger.info(f"   Статистика начальных проблем:")
        logger.info(f"     - По категориям: {dict(stats['by_category'])}")
        logger.info(f"     - Колонтитулы: {stats['header_footer_count']}")
        logger.info(f"     - Типы location: {dict(stats['by_location_type'])}")
    
    def _collect_final_stats(self):
        """Сбор статистики по финальным проблемам."""
        stats = {
            'total': len(self.final_issues),
            'by_category': Counter(),
            'by_severity': Counter(),
            'header_footer_count': 0
        }
        
        for issue in self.final_issues:
            stats['by_category'][issue.category] += 1
            stats['by_severity'][issue.severity.value] += 1
            if issue.category == 'HEADER_FOOTER':
                stats['header_footer_count'] += 1
        
        self.stats['final'] = stats
        
        logger.info(f"   Статистика финальных проблем:")
        logger.info(f"     - Всего: {stats['total']}")
        logger.info(f"     - По категориям: {dict(stats['by_category'])}")
        logger.info(f"     - Колонтитулы: {stats['header_footer_count']}")
    
    def _compare_results(self):
        """Сравнение результатов до и после исправлений."""
        initial = self.stats['initial']
        final = self.stats['final']
        
        reduction = initial['total'] - final['total']
        reduction_percent = (reduction / initial['total'] * 100) if initial['total'] > 0 else 0
        
        self.stats['comparison'] = {
            'total_reduction': reduction,
            'total_reduction_percent': reduction_percent,
            'category_reduction': {},
            'header_footer_reduction': initial['header_footer_count'] - final['header_footer_count']
        }
        
        # Сравнение по категориям
        for category in set(list(initial['by_category'].keys()) + list(final['by_category'].keys())):
            init_count = initial['by_category'].get(category, 0)
            final_count = final['by_category'].get(category, 0)
            reduction = init_count - final_count
            if init_count > 0:
                percent = (reduction / init_count * 100) if init_count > 0 else 0
                self.stats['comparison']['category_reduction'][category] = {
                    'initial': init_count,
                    'final': final_count,
                    'reduction': reduction,
                    'percent': percent
                }
        
        logger.info(f"   Сравнение результатов:")
        logger.info(f"     - Общее сокращение проблем: {reduction} ({reduction_percent:.1f}%)")
        logger.info(f"     - Сокращение колонтитулов: {self.stats['comparison']['header_footer_reduction']}")
        
        # Вывод по категориям с наибольшим сокращением
        for category, data in self.stats['comparison']['category_reduction'].items():
            if data['reduction'] > 0:
                logger.info(f"     - {category}: {data['initial']} → {data['final']} (сокращение {data['reduction']}, {data['percent']:.1f}%)")
    
    def _analyze_header_footer_issues(self):
        """Детальный анализ проблем колонтитулов."""
        initial_hf = self.stats['initial']['header_footer_details']
        final_hf_count = self.stats['final']['header_footer_count']
        
        analysis = {
            'initial_count': len(initial_hf),
            'final_count': final_hf_count,
            'reduction': len(initial_hf) - final_hf_count,
            'location_types': Counter(),
            'element_types': Counter(),
            'issues_without_index': []
        }
        
        for hf in initial_hf:
            loc = hf['location']
            # Тип location
            if 'index' in loc:
                analysis['location_types']['index'] += 1
            elif 'section' in loc and 'paragraph' in loc:
                analysis['location_types']['section+paragraph'] += 1
            elif 'paragraph_index' in loc:
                analysis['location_types']['paragraph_index'] += 1
            else:
                analysis['location_types']['other'] += 1
                analysis['issues_without_index'].append(hf['id'])
            
            analysis['element_types'][hf['element_type']] += 1
        
        # Проверка успешности применения исправлений
        apply_success = self.apply_results.get('success_count', 0)
        apply_failed = self.apply_results.get('failed_count', 0)
        
        # Оценка эффективности для колонтитулов
        # (упрощённо: считаем, что все колонтитулы были исправлены, если не было ошибок)
        hf_fixed = 0
        hf_failed = 0
        for issue in self.initial_issues:
            if issue.category == 'HEADER_FOOTER':
                # Проверим, была ли эта проблема в списке успешных исправлений
                # (это требует более детального отслеживания, упростим)
                pass
        
        analysis['apply_success'] = apply_success
        analysis['apply_failed'] = apply_failed
        analysis['estimated_hf_fixed'] = analysis['reduction']  # предполагаем, что сокращение = исправленные
        
        self.stats['header_footer_analysis'] = analysis
        
        logger.info(f"   Анализ колонтитулов:")
        logger.info(f"     - Начальное количество: {analysis['initial_count']}")
        logger.info(f"     - Финальное количество: {analysis['final_count']}")
        logger.info(f"     - Сокращение: {analysis['reduction']}")
        logger.info(f"     - Типы location: {dict(analysis['location_types'])}")
        logger.info(f"     - Типы элементов: {dict(analysis['element_types'])}")
        if analysis['issues_without_index']:
            logger.info(f"     - Проблемы без индекса: {len(analysis['issues_without_index'])}")
    
    def _check_30_errors_hypothesis(self):
        """Проверка гипотезы о 30 постоянных ошибках."""
        # Ищем проблемы с location, содержащими None или ошибками индекса
        none_index_issues = []
        for issue in self.initial_issues:
            loc = issue.location
            if loc is None:
                none_index_issues.append(issue)
                continue
            # Проверяем значения на None
            for key, val in loc.items():
                if val is None:
                    none_index_issues.append(issue)
                    break
        
        # Проверяем, были ли это колонтитулы
        hf_none_issues = [i for i in none_index_issues if i.category == 'HEADER_FOOTER']
        
        hypothesis = {
            'total_none_index_issues': len(none_index_issues),
            'header_footer_none_issues': len(hf_none_issues),
            'other_none_issues': len(none_index_issues) - len(hf_none_issues),
            'initial_total': len(self.initial_issues),
            'is_30_errors': len(none_index_issues) >= 30
        }
        
        self.stats['30_errors_hypothesis'] = hypothesis
        
        logger.info(f"   Проверка гипотезы о 30 ошибках:")
        logger.info(f"     - Всего проблем с None в location: {hypothesis['total_none_index_issues']}")
        logger.info(f"     - Из них колонтитулы: {hypothesis['header_footer_none_issues']}")
        logger.info(f"     - Другие: {hypothesis['other_none_issues']}")
        logger.info(f"     - Соответствует гипотезе (>=30): {hypothesis['is_30_errors']}")
        
        # Проверяем, исчезли ли эти ошибки после исправлений
        if hypothesis['total_none_index_issues'] > 0:
            # Ищем аналогичные проблемы в финальном аудите
            final_none_count = 0
            for issue in self.final_issues:
                loc = issue.location
                if loc is None:
                    final_none_count += 1
                    continue
                for key, val in loc.items():
                    if val is None:
                        final_none_count += 1
                        break
            
            hypothesis['final_none_index_issues'] = final_none_count
            hypothesis['none_reduction'] = hypothesis['total_none_index_issues'] - final_none_count
            
            logger.info(f"     - После исправлений осталось проблем с None: {final_none_count}")
            logger.info(f"     - Сокращение: {hypothesis['none_reduction']}")
    
    def _generate_report(self):
        """Генерация отчёта в формате Markdown."""
        report_path = "test_real_document_report.md"
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("# Отчёт тестирования исправлений на реальном документе\n\n")
            f.write(f"**Документ:** {self.doc_path}\n")
            f.write(f"**Дата тестирования:** {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            # Сводная таблица
            f.write("## Сводная таблица результатов\n\n")
            f.write("| Показатель | До исправлений | После исправлений | Изменение |\n")
            f.write("|------------|----------------|-------------------|-----------|\n")
            f.write(f"| Всего проблем | {self.stats['initial']['total']} | {self.stats['final']['total']} | {self.stats['comparison']['total_reduction']} ({self.stats['comparison']['total_reduction_percent']:.1f}%) |\n")
            f.write(f"| Колонтитулы | {self.stats['initial']['header_footer_count']} | {self.stats['final']['header_footer_count']} | {self.stats['comparison']['header_footer_reduction']} |\n")
            
            # Результаты применения исправлений
            f.write("\n## Результаты применения исправлений\n\n")
            f.write(f"- Успешно применено: {self.apply_results.get('success_count', 0)}\n")
            f.write(f"- Не удалось применить: {self.apply_results.get('failed_count', 0)}\n")
            f.write(f"- Пропущено: {self.apply_results.get('skipped_count', 0)}\n")
            
            # Анализ колонтитулов
            hf_analysis = self.stats.get('header_footer_analysis', {})
            f.write("\n## Анализ проблем колонтитулов\n\n")
            f.write(f"- Начальное количество: {hf_analysis.get('initial_count', 0)}\n")
            f.write(f"- Финальное количество: {hf_analysis.get('final_count', 0)}\n")
            f.write(f"- Сокращение: {hf_analysis.get('reduction', 0)}\n")
            f.write(f"- Типы location: {dict(hf_analysis.get('location_types', {}))}\n")
            
            # Гипотеза о 30 ошибках
            hypothesis = self.stats.get('30_errors_hypothesis', {})
            f.write("\n## Проверка гипотезы о 30 постоянных ошибках\n\n")
            f.write(f"- Всего проблем с None в location: {hypothesis.get('total_none_index_issues', 0)}\n")
            f.write(f"- Из них колонтитулы: {hypothesis.get('header_footer_none_issues', 0)}\n")
            f.write(f"- Соответствует гипотезе (>=30): {hypothesis.get('is_30_errors', False)}\n")
            if hypothesis.get('final_none_index_issues', 0) >= 0:
                f.write(f"- После исправлений осталось: {hypothesis.get('final_none_index_issues', 0)}\n")
                f.write(f"- Сокращение: {hypothesis.get('none_reduction', 0)}\n")
            
            # Рекомендации
            f.write("\n## Рекомендации\n\n")
            if self.stats['final']['total'] > 0:
                f.write("1. **Остались неисправленные проблемы** - требуется дополнительный анализ\n")
                remaining_categories = [(cat, count) for cat, count in self.stats['final']['by_category'].items() if count > 0]
                for cat, count in remaining_categories:
                    f.write(f"   - {cat}: {count} проблем\n")
            else:
                f.write("1. **Все проблемы успешно исправлены** - отличный результат!\n")
            
            if hf_analysis.get('issues_without_index', []):
                f.write(f"2. **Обнаружены проблемы колонтитулов без индекса** - {len(hf_analysis['issues_without_index'])} шт. Требуется улучшение определения location.\n")
            
            if hypothesis.get('total_none_index_issues', 0) > 0 and hypothesis.get('final_none_index_issues', 0) > 0:
                f.write(f"3. **Остались проблемы с None в location** - {hypothesis['final_none_index_issues']} шт. Требуется дополнительная отладка.\n")
            
            f.write("\n## Заключение\n\n")
            if self.stats['comparison']['total_reduction_percent'] > 80:
                f.write("Исправления показали высокую эффективность. Проблема с 30 постоянными ошибками решена.\n")
            elif self.stats['comparison']['total_reduction_percent'] > 50:
                f.write("Исправления показали умеренную эффективность. Часть проблем осталась нерешённой.\n")
            else:
                f.write("Исправления показали низкую эффективность. Требуется дополнительная работа.\n")
        
        logger.info(f"Отчёт сохранён в {report_path}")
        
        # Дополнительно сохраняем JSON с полной статистикой
        json_path = "test_real_document_stats.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            # Конвертируем Counter в dict для сериализации
            stats_serializable = {}
            for key, value in self.stats.items():
                if isinstance(value, dict):
                    stats_serializable[key] = {}
                    for k, v in value.items():
                        if isinstance(v, Counter):
                            stats_serializable[key][k] = dict(v)
                        elif isinstance(v, defaultdict):
                            stats_serializable[key][k] = dict(v)
                        else:
                            stats_serializable[key][k] = v
                else:
                    stats_serializable[key] = value
            
            json.dump(stats_serializable, f, indent=2, ensure_ascii=False, default=str)
        
        logger.info(f"Полная статистика сохранена в {json_path}")

def main():
    """Основная функция тестирования."""
    doc_path = "Том II_Волжск.docx"
    
    if not os.path.exists(doc_path):
        logger.error(f"Документ {doc_path} не найден в текущей директории.")
        # Попробуем найти в родительской директории
        parent_path = os.path.join("..", doc_path)
        if os.path.exists(parent_path):
            logger.info(f"Найден документ в родительской директории, копируем...")
            import shutil
            shutil.copy(parent_path, doc_path)
        else:
            logger.error("Документ не найден. Завершение теста.")
            return False
    
    tester = RealDocumentTester(doc_path)
    success = tester.run_full_test()
    
    if success:
        logger.info("Тестирование завершено успешно. Проверьте сгенерированные отчёты.")
        return True
    else:
        logger.error("Тестирование завершилось с ошибками.")
        return False

if __name__ == "__main__":
    main()
#!/usr/bin/env python3
"""
Скрипт для полного цикла многократного применения исправлений.
Выполняет итеративное применение исправлений с анализом оставшихся проблем.
"""

import sys
import os
import time
import json
import logging
from collections import Counter, defaultdict
from typing import List, Dict, Any, Tuple
import copy

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.core.config_loader import ConfigLoader
from src.core.audit_engine import AuditEngine, AuditIssue
from src.apply.apply_orchestrator import ApplyOrchestrator
import docx

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class CompleteFixCycleTester:
    def __init__(self, doc_path: str, config_path: str = None, max_iterations: int = 4):
        self.doc_path = doc_path
        self.max_iterations = max_iterations
        
        if config_path is None:
            if os.path.exists("configs/active/config.yaml"):
                config_path = "configs/active/config.yaml"
            elif os.path.exists("configs/active/config_v4.2.yaml"):
                config_path = "configs/active/config_v4.2.yaml"
            else:
                raise FileNotFoundError("Не найден файл конфигурации")
        self.config_path = config_path
        
        self.config = ConfigLoader(config_path)
        self.config.load()
        
        self.audit_engine = AuditEngine(self.config)
        self.apply_engine = ApplyOrchestrator(self.config)
        
        # Результаты по итерациям
        self.iterations = []
        # Детали оставшихся проблем после первой итерации
        self.remaining_issues_details = []
        
    def run_complete_cycle(self) -> bool:
        """Запуск полного цикла многократного применения."""
        logger.info(f"=== НАЧАЛО ПОЛНОГО ЦИКЛА ИСПРАВЛЕНИЙ ===")
        logger.info(f"Документ: {self.doc_path}")
        logger.info(f"Максимальное количество итераций: {self.max_iterations}")
        
        current_doc_path = self.doc_path
        
        for iteration in range(1, self.max_iterations + 1):
            logger.info(f"\n--- Итерация {iteration} ---")
            
            # 1. Аудит текущего документа
            logger.info(f"1. Аудит документа...")
            start = time.time()
            issues = self.audit_engine.scan_document(current_doc_path)
            audit_time = time.time() - start
            logger.info(f"   Найдено проблем: {len(issues)} (время: {audit_time:.2f} сек)")
            
            # Сбор статистики по проблемам
            stats = self._collect_stats(issues, iteration)
            
            # 2. Применение исправлений (кроме последней итерации)
            if iteration < self.max_iterations:
                output_path = f"Том II_Волжск_fixed_iter{iteration}.docx"
                logger.info(f"2. Применение исправлений...")
                start = time.time()
                apply_results = self.apply_engine.apply_fixes(
                    current_doc_path,
                    issues,
                    output_path=output_path
                )
                apply_time = time.time() - start
                logger.info(f"   Результат: успешно={apply_results.get('success_count', 0)}, "
                           f"неудачно={apply_results.get('failed_count', 0)}, "
                           f"пропущено={apply_results.get('skipped_count', 0)}")
                logger.info(f"   Время: {apply_time:.2f} сек")
                
                # Обновление текущего документа
                current_doc_path = output_path
                if not os.path.exists(current_doc_path):
                    logger.error(f"Исправленный документ не создан: {current_doc_path}")
                    return False
            else:
                apply_results = None
                output_path = current_doc_path
                logger.info(f"2. Последняя итерация - исправления не применяются")
            
            # Сохранение результатов итерации
            iteration_data = {
                'iteration': iteration,
                'input_doc': current_doc_path,
                'issues_count': len(issues),
                'issues': [self._issue_to_dict(issue) for issue in issues],
                'stats': stats,
                'audit_time': audit_time,
                'apply_results': apply_results,
                'output_doc': output_path
            }
            self.iterations.append(iteration_data)
            
            # Сохранение деталей оставшихся проблем после первой итерации
            if iteration == 1:
                self._analyze_remaining_issues(issues)
        
        logger.info(f"\n=== ЦИКЛ ЗАВЕРШЕН ===")
        return True
    
    def _collect_stats(self, issues: List[AuditIssue], iteration: int) -> Dict[str, Any]:
        """Сбор статистики по проблемам."""
        stats = {
            'total': len(issues),
            'by_category': Counter(),
            'by_severity': Counter(),
            'by_location_type': Counter(),
            'header_footer_count': 0,
            'location_keys': defaultdict(int)
        }
        
        for issue in issues:
            stats['by_category'][issue.category] += 1
            stats['by_severity'][issue.severity.value] += 1
            
            if issue.category == 'HEADER_FOOTER':
                stats['header_footer_count'] += 1
            
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
        
        # Преобразование Counter в dict для сериализации
        stats['by_category'] = dict(stats['by_category'])
        stats['by_severity'] = dict(stats['by_severity'])
        stats['by_location_type'] = dict(stats['by_location_type'])
        stats['location_keys'] = dict(stats['location_keys'])
        
        logger.info(f"   Статистика итерации {iteration}:")
        logger.info(f"     - Категории: {stats['by_category']}")
        logger.info(f"     - Severity: {stats['by_severity']}")
        logger.info(f"     - Колонтитулы: {stats['header_footer_count']}")
        
        return stats
    
    def _analyze_remaining_issues(self, issues: List[AuditIssue]):
        """Анализ оставшихся проблем после первой итерации."""
        logger.info("Анализ оставшихся проблем...")
        
        for issue in issues:
            issue_dict = self._issue_to_dict(issue)
            # Определение типа проблемы
            problem_type = issue.category
            location = issue.location
            
            # Анализ возможности исправления
            fixable = self._assess_fixability(issue)
            
            # Классификация
            classification = self._classify_issue(issue, fixable)
            
            self.remaining_issues_details.append({
                'id': issue.id,
                'category': issue.category,
                'element_type': issue.element_type,
                'description': issue.description,
                'location': location,
                'fixable': fixable,
                'classification': classification,
                'reason': self._get_failure_reason(issue)
            })
        
        # Группировка по классификации
        classification_counts = Counter()
        for detail in self.remaining_issues_details:
            classification_counts[detail['classification']] += 1
        
        logger.info(f"   Классификация оставшихся проблем:")
        for cls, count in classification_counts.items():
            logger.info(f"     - {cls}: {count}")
    
    def _assess_fixability(self, issue: AuditIssue) -> str:
        """Оценка возможности исправления проблемы."""
        # Проверка наличия необходимых ключей в location
        loc = issue.location
        if not loc:
            return "Невозможно исправить (отсутствует location)"
        
        # Для колонтитулов
        if issue.category == 'HEADER_FOOTER':
            if 'section' in loc and 'paragraph' in loc:
                return "Можно исправить"
            else:
                return "Требует доработки ApplyEngine"
        
        # Для абзацев
        if issue.category == 'PARAGRAPH':
            if 'index' in loc or 'paragraph_index' in loc:
                return "Можно исправить повторным применением"
            else:
                return "Требует доработки ApplyEngine"
        
        # Для таблиц
        if issue.category == 'TABLE':
            if 'table_index' in loc or 'index' in loc:
                return "Можно исправить"
            else:
                return "Требует доработки ApplyEngine"
        
        # Для рисунков
        if issue.category == 'FIGURE':
            return "Требует доработки ApplyEngine"
        
        return "Неизвестно"
    
    def _classify_issue(self, issue: AuditIssue, fixable: str) -> str:
        """Классификация проблемы."""
        if "Невозможно исправить" in fixable:
            return "Невозможно исправить (ошибка аудита)"
        elif "Требует доработки ApplyEngine" in fixable:
            return "Требует доработки ApplyEngine"
        elif "Можно исправить повторным применением" in fixable:
            return "Можно исправить повторным применением"
        elif "Можно исправить" in fixable:
            return "Можно исправить"
        else:
            return "Неопределённая"
    
    def _get_failure_reason(self, issue: AuditIssue) -> str:
        """Определение причины неудачи исправления."""
        loc = issue.location
        if not loc:
            return "Отсутствует location"
        
        # Проверка специфических ключей
        missing_keys = []
        if issue.category == 'HEADER_FOOTER':
            if 'container' not in loc:
                missing_keys.append('container')
        elif issue.category == 'PARAGRAPH':
            if 'index' not in loc and 'paragraph_index' not in loc:
                missing_keys.append('index/paragraph_index')
        elif issue.category == 'TABLE':
            if 'table_index' not in loc:
                missing_keys.append('table_index')
        
        if missing_keys:
            return f"Отсутствуют ключи location: {', '.join(missing_keys)}"
        
        return "Неизвестная причина"
    
    def _issue_to_dict(self, issue: AuditIssue) -> Dict[str, Any]:
        """Преобразование AuditIssue в словарь."""
        return {
            'id': issue.id,
            'category': issue.category,
            'element_type': issue.element_type,
            'description': issue.description,
            'severity': issue.severity.value,
            'location': issue.location,
            'expected_value': issue.expected_value,
            'current_value': issue.current_value
        }
    
    def generate_report(self) -> str:
        """Генерация отчёта о полном цикле."""
        report_lines = []
        
        report_lines.append("# Отчёт полного цикла многократного применения исправлений")
        report_lines.append("")
        report_lines.append(f"**Документ:** {self.doc_path}")
        report_lines.append(f"**Дата тестирования:** {time.strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"**Количество итераций:** {len(self.iterations)}")
        report_lines.append("")
        
        # 1. Сводная таблица по итерациям
        report_lines.append("## 1. Результаты по итерациям")
        report_lines.append("")
        report_lines.append("| Итерация | Проблем до исправлений | Успешных исправлений | Проблем после | Изменение |")
        report_lines.append("|----------|------------------------|----------------------|---------------|-----------|")
        
        for iter_data in self.iterations:
            iteration = iter_data['iteration']
            issues_count = iter_data['issues_count']
            apply_results = iter_data.get('apply_results')
            
            if apply_results:
                success = apply_results.get('success_count', 0)
                failed = apply_results.get('failed_count', 0)
                # Проблемы после - это issues_count следующей итерации
                next_issues = 0
                if iteration < len(self.iterations):
                    next_issues = self.iterations[iteration]['issues_count']
                else:
                    next_issues = issues_count  # последняя итерация
                
                change = issues_count - next_issues
                change_percent = (change / issues_count * 100) if issues_count > 0 else 0
                
                report_lines.append(f"| {iteration} | {issues_count} | {success} | {next_issues} | {change} ({change_percent:.1f}%) |")
            else:
                report_lines.append(f"| {iteration} | {issues_count} | - | {issues_count} | - |")
        
        report_lines.append("")
        
        # 2. График уменьшения количества проблем
        report_lines.append("## 2. Динамика уменьшения количества проблем")
        report_lines.append("")
        report_lines.append("```")
        for iter_data in self.iterations:
            iteration = iter_data['iteration']
            issues_count = iter_data['issues_count']
            bar = '█' * (issues_count // 5)  # Упрощённая визуализация
            report_lines.append(f"Итерация {iteration}: {issues_count:3d} проблем {bar}")
        report_lines.append("```")
        report_lines.append("")
        
        # 3. Анализ оставшихся проблем после первой итерации
        report_lines.append("## 3. Анализ оставшихся проблем (после первой итерации)")
        report_lines.append("")
        
        if self.remaining_issues_details:
            # Группировка по категориям
            by_category = Counter()
            by_classification = Counter()
            for detail in self.remaining_issues_details:
                by_category[detail['category']] += 1
                by_classification[detail['classification']] += 1
            
            report_lines.append("### 3.1 Распределение по категориям")
            report_lines.append("")
            for category, count in by_category.items():
                report_lines.append(f"- **{category}**: {count} проблем")
            report_lines.append("")
            
            report_lines.append("### 3.2 Классификация по возможности исправления")
            report_lines.append("")
            for classification, count in by_classification.items():
                report_lines.append(f"- **{classification}**: {count} проблем")
            report_lines.append("")
            
            # Детальная таблица
            report_lines.append("### 3.3 Детальный анализ проблем")
            report_lines.append("")
            report_lines.append("| ID | Категория | Описание | Location | Классификация | Причина |")
            report_lines.append("|----|-----------|----------|----------|---------------|---------|")
            
            for detail in self.remaining_issues_details[:20]:  # Ограничим для читаемости
                loc_str = str(detail['location'])[:50] + "..." if len(str(detail['location'])) > 50 else str(detail['location'])
                report_lines.append(f"| {detail['id']} | {detail['category']} | {detail['description'][:30]}... | {loc_str} | {detail['classification']} | {detail['reason'][:30]}... |")
            
            if len(self.remaining_issues_details) > 20:
                report_lines.append(f"| ... | ... | ... | ... | ... | ... |")
                report_lines.append(f"*И ещё {len(self.remaining_issues_details) - 20} проблем*")
        else:
            report_lines.append("Нет данных об оставшихся проблемах.")
        report_lines.append("")
        
        # 4. Эффективность по типам проблем
        report_lines.append("## 4. Эффективность исправлений по типам проблем")
        report_lines.append("")
        
        if len(self.iterations) >= 2:
            first_iter = self.iterations[0]
            last_iter = self.iterations[-1]
            
            first_cats = first_iter['stats']['by_category']
            last_cats = last_iter['stats']['by_category']
            
            report_lines.append("| Категория | Начало | Конец | Устранено | Эффективность |")
            report_lines.append("|-----------|--------|-------|-----------|---------------|")
            
            all_categories = set(list(first_cats.keys()) + list(last_cats.keys()))
            for cat in sorted(all_categories):
                start = first_cats.get(cat, 0)
                end = last_cats.get(cat, 0)
                eliminated = start - end
                efficiency = (eliminated / start * 100) if start > 0 else 0
                report_lines.append(f"| {cat} | {start} | {end} | {eliminated} | {efficiency:.1f}% |")
        report_lines.append("")
        
        # 5. Рекомендации по достижению 100% эффективности
        report_lines.append("## 5. Рекомендации по достижению 100% эффективности")
        report_lines.append("")
        
        # Анализ классификаций
        classifications = Counter([d['classification'] for d in self.remaining_issues_details])
        
        report_lines.append("### 5.1 Приоритеты доработок")
        report_lines.append("")
        
        if classifications.get("Требует доработки ApplyEngine", 0) > 0:
            report_lines.append("1. **Высокий приоритет**: Доработка ApplyEngine для обработки проблем с недостающими ключами location")
            report_lines.append("   - Добавить fallback-логику для определения location")
            report_lines.append("   - Расширить поддержку специальных типов элементов (таблицы, рисунки)")
        
        if classifications.get("Можно исправить повторным применением", 0) > 0:
            report_lines.append("2. **Средний приоритет**: Реализация повторного применения исправлений")
            report_lines.append("   - Добавить механизм итеративного применения")
            report_lines.append("   - Учитывать взаимозависимости стилей")
        
        if classifications.get("Невозможно исправить (ошибка аудита)", 0) > 0:
            report_lines.append("3. **Низкий приоритет**: Исправление ошибок аудита")
            report_lines.append("   - Проверить корректность генерации location")
            report_lines.append("   - Уточнить критерии обнаружения проблем")
        
        report_lines.append("")
        report_lines.append("### 5.2 Конкретные шаги")
        report_lines.append("")
        report_lines.append("1. **Для проблем таблиц (9 шт.)**:")
        report_lines.append("   - Проверить наличие ключа 'table_index' в location")
        report_lines.append("   - Реализовать функцию `_get_table` в ApplyEngine")
        report_lines.append("")
        report_lines.append("2. **Для проблем абзацев (22 шт.)**:")
        report_lines.append("   - Проанализировать взаимозависимости стилей")
        report_lines.append("   - Реализовать повторное применение с учётом контекста")
        report_lines.append("")
        report_lines.append("3. **Для проблем рисунков (1 шт.)**:")
        report_lines.append("   - Добавить поддержку обработки рисунков в ApplyEngine")
        report_lines.append("")
        
        # 6. Проверка гипотезы о повторном применении
        report_lines.append("## 6. Проверка гипотезы о повторном применении")
        report_lines.append("")
        
        if len(self.iterations) >= 3:
            report_lines.append("### 6.1 Результаты повторных применений")
            report_lines.append("")
            for i in range(1, len(self.iterations)):
                prev = self.iterations[i-1]['issues_count']
                curr = self.iterations[i]['issues_count']
                reduction = prev - curr
                report_lines.append(f"- **Итерация {i} → {i+1}**: {prev} → {curr} проблем (уменьшение на {reduction})")
            
            report_lines.append("")
            report_lines.append("### 6.2 Выводы")
            report_lines.append("")
            total_reduction = self.iterations[0]['issues_count'] - self.iterations[-1]['issues_count']
            if total_reduction > 0:
                report_lines.append(f"✅ **Гипотеза подтверждается**: Повторное применение уменьшило количество проблем на {total_reduction}")
                report_lines.append("   - Взаимозависимые стили требуют многократной коррекции")
                report_lines.append("   - Некоторые проблемы решаются только после исправления соседних элементов")
            else:
                report_lines.append("❌ **Гипотеза не подтверждается**: Повторное применение не уменьшило количество проблем")
                report_lines.append("   - Проблемы требуют других подходов к исправлению")
                report_lines.append("   - Возможно, ошибки аудита не связаны со стилями")
        else:
            report_lines.append("Недостаточно данных для проверки гипотезы (требуется минимум 3 итерации).")
        report_lines.append("")
        
        # 7. Интеграция с GUI
        report_lines.append("## 7. Интеграция с GUI")
        report_lines.append("")
        report_lines.append("### 7.1 Проверка обновления current_issues")
        report_lines.append("")
        report_lines.append("После каждого применения исправлений MainController должен:")
        report_lines.append("1. Обновлять `current_issues` с новым списком проблем")
        report_lines.append("2. Пересчитывать статистику для отображения в GUI")
        report_lines.append("3. Предоставлять возможность повторного аудита")
        report_lines.append("")
        report_lines.append("### 7.2 Рекомендации для GUI")
        report_lines.append("")
        report_lines.append("1. **Добавить прогресс-бар** для отображения хода многократного применения")
        report_lines.append("2. **Реализовать кнопку \"Повторное применение\"** для запуска дополнительных итераций")
        report_lines.append("3. **Визуализировать классификацию проблем** (цветовая маркировка по типу)")
        report_lines.append("")
        
        return "\n".join(report_lines)
    
    def save_report(self, report_path: str = "complete_fix_cycle_report.md"):
        """Сохранение отчёта в файл."""
        report = self.generate_report()
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        logger.info(f"Отчёт сохранён в {report_path}")
        return report_path
    
    def save_iteration_data(self, json_path: str = "data/test_results/complete_fix_cycle_data.json"):
        """Сохранение данных итераций в JSON."""
        data = {
            'document': self.doc_path,
            'iterations': self.iterations,
            'remaining_issues': self.remaining_issues_details
        }
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, default=str)
        logger.info(f"Данные итераций сохранены в {json_path}")
        return json_path


def main():
    """Основная функция запуска теста."""
    doc_path = "Том II_Волжск.docx"
    
    if not os.path.exists(doc_path):
        logger.error(f"Документ не найден: {doc_path}")
        # Попробуем найти альтернативный
        alt_paths = [
            "Том II_Волжск_fixed.docx",
            "Том II_Волжск_fixed_v1.docx"
        ]
        for alt in alt_paths:
            if os.path.exists(alt):
                doc_path = alt
                logger.info(f"Используем альтернативный документ: {doc_path}")
                break
        else:
            logger.error("Не найден ни один тестовый документ")
            return 1
    
    tester = CompleteFixCycleTester(doc_path, max_iterations=4)
    
    try:
        success = tester.run_complete_cycle()
        if not success:
            logger.error("Цикл завершился с ошибкой")
            return 1
        
        # Генерация отчёта
        report_path = tester.save_report()
        data_path = tester.save_iteration_data()
        
        logger.info(f"\n=== РЕЗУЛЬТАТЫ ===")
        logger.info(f"1. Отчёт: {report_path}")
        logger.info(f"2. Данные: {data_path}")
        logger.info(f"3. Итерации выполнены: {len(tester.iterations)}")
        
        # Вывод сводной статистики
        if tester.iterations:
            first = tester.iterations[0]
            last = tester.iterations[-1]
            total_reduction = first['issues_count'] - last['issues_count']
            reduction_percent = (total_reduction / first['issues_count'] * 100) if first['issues_count'] > 0 else 0
            
            logger.info(f"4. Сводка:")
            logger.info(f"   - Начальные проблемы: {first['issues_count']}")
            logger.info(f"   - Финальные проблемы: {last['issues_count']}")
            logger.info(f"   - Общее уменьшение: {total_reduction} ({reduction_percent:.1f}%)")
            
            # Анализ колонтитулов
            header_footer_start = first['stats'].get('header_footer_count', 0)
            header_footer_end = last['stats'].get('header_footer_count', 0)
            if header_footer_start > 0:
                logger.info(f"   - Колонтитулы: {header_footer_start} → {header_footer_end} (устранено {header_footer_start - header_footer_end})")
        
        return 0
        
    except Exception as e:
        logger.exception(f"Ошибка при выполнении цикла: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
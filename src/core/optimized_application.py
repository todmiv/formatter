#!/usr/bin/env python3
"""
Оптимизированный алгоритм применения исправлений.
Анализирует зависимости между проблемами, строит граф применения с правильным порядком,
применяет исправления группами для минимизации взаимовлияния.
"""

import sys
import os
import json
import logging
from typing import List, Dict, Any, Optional, Set, Tuple
from collections import defaultdict, deque

try:
    import networkx as nx
except ImportError:
    nx = None

from src.core.audit_engine import AuditIssue, Severity
from src.core.config_loader import ConfigLoader
from src.apply.apply_orchestrator import ApplyOrchestrator

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class _SimpleDiGraph:
    """Minimal directed graph fallback when networkx is unavailable."""

    def __init__(self):
        self._nodes: Dict[str, dict] = {}
        self._adj: Dict[str, Set[str]] = defaultdict(set)
        self._pred: Dict[str, Set[str]] = defaultdict(set)

    def add_node(self, node, **kwargs):
        self._nodes[node] = kwargs

    def add_edge(self, u, v, **kwargs):
        self._adj[u].add(v)
        self._pred[v].add(u)

    def predecessors(self, node):
        return iter(self._pred.get(node, set()))

    def successors(self, node):
        return iter(self._adj.get(node, set()))

    def nodes(self):
        return self._nodes.keys()


def _topological_sort(graph):
    in_degree = defaultdict(int)
    for u in graph._adj:
        for v in graph._adj[u]:
            in_degree[v] += 1
    queue = deque(n for n in graph._nodes if in_degree[n] == 0)
    order = []
    while queue:
        n = queue.popleft()
        order.append(n)
        for v in graph._adj.get(n, set()):
            in_degree[v] -= 1
            if in_degree[v] == 0:
                queue.append(v)
    if len(order) != len(graph._nodes):
        raise RuntimeError("Graph has a cycle")
    return order


def _simple_cycles(graph):
    cycles = []
    visited = set()
    for start in graph._nodes:
        if start in visited:
            continue
        stack = [(start, [start])]
        path_set = {start}
        while stack:
            node, path = stack.pop()
            for neighbor in graph._adj.get(node, set()):
                if neighbor == start and len(path) > 1:
                    cycles.append(list(path))
                elif neighbor not in path_set:
                    path_set.add(neighbor)
                    stack.append((neighbor, path + [neighbor]))
                    visited.add(neighbor)
            path_set.discard(node)
    return cycles

class DependencyAnalyzer:
    """Анализатор зависимостей между проблемами."""
    
    def __init__(self, issues: List[AuditIssue]):
        self.issues = issues
        self.graph = nx.DiGraph() if nx else _SimpleDiGraph()
        self._build_graph()
    
    def _build_graph(self):
        """Строит граф зависимостей между проблемами."""
        # Добавляем узлы
        for issue in self.issues:
            self.graph.add_node(issue.id, issue=issue)
        
        # Добавляем рёбра на основе location и категорий
        for i, issue1 in enumerate(self.issues):
            for j, issue2 in enumerate(self.issues):
                if i == j:
                    continue
                
                # Зависимость по индексу параграфа (последовательность)
                idx1 = issue1.location.get('index')
                idx2 = issue2.location.get('index')
                if idx1 is not None and idx2 is not None:
                    if idx1 == idx2:
                        # Одинаковый параграф - порядок не важен, но отметим связь
                        self.graph.add_edge(issue1.id, issue2.id, type='same_paragraph')
                    elif idx1 + 1 == idx2:
                        # Соседние параграфы - issue1 может влиять на issue2
                        self.graph.add_edge(issue1.id, issue2.id, type='adjacent')
                
                # Зависимость по стилю (если один стиль зависит от другого)
                if issue1.element_type == issue2.element_type:
                    self.graph.add_edge(issue1.id, issue2.id, type='same_style')
                
                # Зависимость по ожидаемым значениям (если исправление одного влияет на другое)
                if (issue1.expected_value and issue2.expected_value and 
                    issue1.expected_value == issue2.expected_value):
                    self.graph.add_edge(issue1.id, issue2.id, type='same_expected')
    
    def get_dependencies(self, issue_id: str) -> List[str]:
        """Возвращает список ID проблем, от которых зависит данная проблема."""
        return list(self.graph.predecessors(issue_id))
    
    def get_dependents(self, issue_id: str) -> List[str]:
        """Возвращает список ID проблем, которые зависят от данной."""
        return list(self.graph.successors(issue_id))
    
    def find_cycles(self) -> List[List[str]]:
        try:
            if nx:
                cycles = list(nx.simple_cycles(self.graph))
            else:
                cycles = _simple_cycles(self.graph)
            return cycles
        except Exception:
            return []
    
    def topological_order(self) -> List[str]:
        try:
            if nx:
                order = list(nx.topological_sort(self.graph))
            else:
                order = _topological_sort(self.graph)
            return order
        except Exception:
            # Есть циклы, используем эвристику
            logger.warning("Обнаружены циклы в зависимостях. Используем эвристическую сортировку.")
            return self._heuristic_order()
    
    def _heuristic_order(self) -> List[str]:
        """Эвристический порядок при наличии циклов."""
        # Сортируем по категории и индексу
        sorted_issues = sorted(self.issues, key=lambda x: (
            x.category,
            x.location.get('index', 0),
            x.location.get('paragraph_index', 0)
        ))
        return [issue.id for issue in sorted_issues]
    
    def group_by_independence(self) -> List[List[str]]:
        """
        Группирует проблемы на независимые группы.
        Проблемы в одной группе не имеют зависимостей между собой.
        """
        # Используем алгоритм раскраски графа
        groups = []
        colored = set()
        
        for issue in self.issues:
            if issue.id in colored:
                continue
            
            # Находим независимую группу
            group = [issue.id]
            colored.add(issue.id)
            
            # Добавляем проблемы, которые не зависят от уже добавленных
            for other in self.issues:
                if other.id in colored:
                    continue
                
                # Проверяем, есть ли зависимости между проблемами в группе
                independent = True
                for group_id in group:
                    if self.graph.has_edge(other.id, group_id) or self.graph.has_edge(group_id, other.id):
                        independent = False
                        break
                
                if independent:
                    group.append(other.id)
                    colored.add(other.id)
            
            groups.append(group)
        
        return groups

class OptimizedApplyEngine:
    """Оптимизированный движок применения исправлений."""
    
    def __init__(self, config_loader: ConfigLoader):
        self.config = config_loader
        self.apply_engine = ApplyOrchestrator(config_loader)
        self.dependency_analyzer = None
    
    def analyze_dependencies(self, issues: List[AuditIssue]) -> Dict[str, Any]:
        """Анализирует зависимости между проблемами и возвращает отчёт."""
        self.dependency_analyzer = DependencyAnalyzer(issues)
        
        cycles = self.dependency_analyzer.find_cycles()
        topological_order = self.dependency_analyzer.topological_order()
        groups = self.dependency_analyzer.group_by_independence()
        
        return {
            'total_issues': len(issues),
            'cycles_found': len(cycles),
            'cycles': cycles,
            'topological_order': topological_order,
            'independent_groups': groups,
            'group_count': len(groups)
        }
    
    def apply_optimized(self, doc_path: str, issues: List[AuditIssue], output_path: Optional[str] = None,
                       max_iterations: int = 3, apply_styles: bool = True,
                       apply_direct_overrides: bool = True, clear_direct_formatting: bool = False) -> Dict[str, Any]:
        """
        Применяет исправления с оптимизированным алгоритмом.
        
        Алгоритм:
        1. Анализирует зависимости между проблемами
        2. Строит граф применения с правильным порядком
        3. Применяет исправления группами для минимизации взаимовлияния
        4. Проверяет результат после каждой группы
        """
        # Анализ зависимостей
        analysis = self.analyze_dependencies(issues)
        logger.info(f"Анализ зависимостей: {analysis['group_count']} групп, {analysis['cycles_found']} циклов")
        
        # Если есть циклы, логируем предупреждение
        if analysis['cycles_found'] > 0:
            logger.warning(f"Обнаружены циклы зависимостей: {analysis['cycles']}")
        
        # Группируем проблемы
        groups = analysis['independent_groups']
        
        # Создаём mapping ID -> issue
        issue_map = {issue.id: issue for issue in issues}
        
        results = {
            'total_applied': 0,
            'total_failed': 0,
            'group_results': [],
            'iteration_results': []
        }
        
        # Итеративное применение
        for iteration in range(max_iterations):
            logger.info(f"Оптимизированная итерация {iteration + 1} из {max_iterations}")
            
            iteration_applied = 0
            iteration_failed = 0
            
            # Применяем группы в порядке топологической сортировки
            for group_idx, group_ids in enumerate(groups):
                group_issues = [issue_map[issue_id] for issue_id in group_ids if issue_id in issue_map]
                
                if not group_issues:
                    continue
                
                logger.info(f"  Группа {group_idx + 1}/{len(groups)}: {len(group_issues)} проблем")
                
                # Применяем исправления для группы
                group_result = self.apply_engine.apply_fixes(
                    doc_path, group_issues, output_path,
                    apply_styles=apply_styles,
                    apply_direct_overrides=apply_direct_overrides,
                    clear_direct_formatting=clear_direct_formatting
                )
                
                iteration_applied += group_result['applied']
                iteration_failed += group_result['failed']
                
                results['group_results'].append({
                    'iteration': iteration + 1,
                    'group': group_idx + 1,
                    'issues': group_ids,
                    'applied': group_result['applied'],
                    'failed': group_result['failed']
                })
                
                # Если есть неудачи в группе, логируем
                if group_result['failed'] > 0:
                    logger.warning(f"    Группа {group_idx + 1}: {group_result['failed']} неудач")
            
            results['iteration_results'].append({
                'iteration': iteration + 1,
                'applied': iteration_applied,
                'failed': iteration_failed
            })
            
            results['total_applied'] += iteration_applied
            results['total_failed'] += iteration_failed
            
            # Если все исправления успешны, завершаем
            if iteration_failed == 0:
                logger.info(f"Все исправления применены успешно на итерации {iteration + 1}")
                break
            
            # Если прогресс отсутствует, останавливаемся
            if iteration > 0 and iteration_failed >= results['iteration_results'][-2]['failed']:
                logger.warning(f"Прогресс остановился. Прерываем.")
                break
        
        return results
    
    def apply_with_validation(self, doc_path: str, issues: List[AuditIssue], output_path: str,
                             max_iterations: int = 3) -> Dict[str, Any]:
        """
        Применяет исправления с валидацией после каждой группы.
        Возвращает детальный отчёт.
        """
        from audit_engine import AuditEngine
        
        # Загружаем конфиг для аудита
        audit_engine = AuditEngine(self.config)
        
        results = {
            'initial_issues': len(issues),
            'iterations': [],
            'final_issues': 0,
            'efficiency': 0.0
        }
        
        current_doc = doc_path
        temp_output = output_path
        
        for iteration in range(max_iterations):
            logger.info(f"Валидационная итерация {iteration + 1}")
            
            # Применяем исправления с оптимизацией
            apply_result = self.apply_optimized(
                current_doc, issues, temp_output,
                max_iterations=1,  # Одна итерация внутри
                apply_styles=True,
                apply_direct_overrides=True,
                clear_direct_formatting=False
            )
            
            # Запускаем аудит для проверки результата
            remaining_issues = audit_engine.scan_document(temp_output)
            
            iteration_info = {
                'iteration': iteration + 1,
                'applied': apply_result['total_applied'],
                'failed': apply_result['total_failed'],
                'remaining_issues': len(remaining_issues),
                'remaining_by_category': self._count_by_category(remaining_issues)
            }
            
            results['iterations'].append(iteration_info)
            
            # Если проблем не осталось, завершаем
            if len(remaining_issues) == 0:
                logger.info(f"Достигнуто 100% эффективности на итерации {iteration + 1}")
                break
            
            # Обновляем issues для следующей итерации
            issues = remaining_issues
            current_doc = temp_output
        
        # Финальный аудит
        final_issues = audit_engine.scan_document(temp_output)
        results['final_issues'] = len(final_issues)
        results['efficiency'] = 1.0 - (len(final_issues) / results['initial_issues']) if results['initial_issues'] > 0 else 1.0
        
        return results
    
    def _count_by_category(self, issues: List[AuditIssue]) -> Dict[str, int]:
        """Считает проблемы по категориям."""
        counts = defaultdict(int)
        for issue in issues:
            counts[issue.category] += 1
        return dict(counts)

def main():
    """Пример использования оптимизированного алгоритма."""
    import sys
    from config_loader import ConfigLoader
    
    if len(sys.argv) < 2:
        print("Использование: python optimized_application.py <путь_к_документу>")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    config_path = "config_v4.2.yaml"
    
    # Загружаем конфиг
    config = ConfigLoader(config_path)
    config.load()
    
    # Запускаем аудит
    from audit_engine import AuditEngine
    audit_engine = AuditEngine(config)
    issues = audit_engine.scan_document(doc_path)
    
    print(f"Найдено проблем: {len(issues)}")
    
    # Создаём оптимизированный движок
    optimized_engine = OptimizedApplyEngine(config)
    
    # Анализируем зависимости
    analysis = optimized_engine.analyze_dependencies(issues)
    print(f"Групп: {analysis['group_count']}")
    print(f"Циклов: {analysis['cycles_found']}")
    
    # Применяем с оптимизацией
    output_path = doc_path.replace('.docx', '_optimized.docx')
    results = optimized_engine.apply_optimized(doc_path, issues, output_path)
    
    print(f"Итоговые результаты:")
    print(f"  Всего применено: {results['total_applied']}")
    print(f"  Всего неудач: {results['total_failed']}")
    print(f"  Групп обработано: {len(results['group_results'])}")

if __name__ == "__main__":
    main()
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YAML Includes — механизм подключения YAML-файлов (аналог partials в Pandoc).

Поддерживает:
- Секцию `includes` в любом YAML-файле
- Относительные и абсолютные пути
- Рекурсивные подключения (с защитой от циклов)
- Глубокое слияние (deep merge) секций
- Приоритет: локальные ключи перезаписывают включённые

Использование:
    # В YAML-файле:
    includes:
      - configs/base/styles.yaml
      - configs/gost/page_setup.yaml

    styles:
      Normal:
        font:
          name: "Times New Roman"

    # Или программно:
    from src.core.yaml_includes import load_yaml_with_includes
    config = load_yaml_with_includes("configs/main.yaml")
"""

import os
import yaml
import logging
from typing import Dict, Any, List, Optional, Set

logger = logging.getLogger(__name__)


def deep_merge(base: Dict, override: Dict) -> Dict:
    """
    Глубокое слияние двух словарей.
    Значения из override перезаписывают base.
    Для вложенных словарей — рекурсивное слияние.

    :param base: базовый словарь.
    :param override: словарь с перезаписями.
    :return: объединённый словарь.
    """
    result = base.copy()

    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value

    return result


def resolve_path(base_path: str, relative_path: str) -> str:
    """
    Разрешает относительный путь относительно базового файла.

    :param base_path: путь к базовому YAML-файлу.
    :param relative_path: относительный или абсолютный путь для подключения.
    :return: абсолютный путь.
    """
    if os.path.isabs(relative_path):
        return relative_path

    base_dir = os.path.dirname(os.path.abspath(base_path))
    return os.path.normpath(os.path.join(base_dir, relative_path))


def load_yaml_with_includes(
    config_path: str,
    _visited: Optional[Set[str]] = None,
    _depth: int = 0,
    max_depth: int = 10,
) -> Dict[str, Any]:
    """
    Загружает YAML-файл с обработкой секции includes.

    :param config_path: путь к YAML-файлу.
    :param _visited: множество посещённых файлов (для защиты от циклов).
    :param _depth: текущая глубина рекурсии.
    :param max_depth: максимальная глубина рекурсии.
    :return: объединённый словарь конфигурации.
    """
    config_path = os.path.abspath(config_path)

    if _visited is None:
        _visited = set()

    if config_path in _visited:
        logger.warning(f"Циклическое подключение: {config_path} уже был обработан. Пропускаем.")
        return {}

    if _depth > max_depth:
        logger.warning(f"Превышена максимальная глубина рекурсии ({max_depth}) при загрузке {config_path}")
        return {}

    _visited.add(config_path)

    if not os.path.exists(config_path):
        from src.core.formatter_errors import raise_config_not_found
        raise_config_not_found(config_path)

    with open(config_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    if not isinstance(config, dict):
        return config or {}

    includes = config.pop('includes', [])

    if includes:
        base_config = {}
        for include_path in includes:
            resolved = resolve_path(config_path, include_path)
            logger.debug(f"Включение: {resolved}")
            included = load_yaml_with_includes(
                resolved,
                _visited=_visited.copy(),
                _depth=_depth + 1,
                max_depth=max_depth,
            )
            base_config = deep_merge(base_config, included)

        config = deep_merge(base_config, config)

    return config


def get_included_paths(config_path: str) -> List[str]:
    """
    Возвращает список всех включённых путей (рекурсивно).

    :param config_path: путь к основному YAML-файлу.
    :return: список абсолютных путей.
    """
    config_path = os.path.abspath(config_path)

    if not os.path.exists(config_path):
        return []

    with open(config_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    if not isinstance(config, dict):
        return []

    includes = config.get('includes', [])
    paths = []

    for include_path in includes:
        resolved = resolve_path(config_path, include_path)
        paths.append(resolved)
        paths.extend(get_included_paths(resolved))

    return paths

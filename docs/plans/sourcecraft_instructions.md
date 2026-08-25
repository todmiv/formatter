# Инструкции по настройке SourceCraft Code Assistant для проекта

## Обзор

Этот документ содержит готовые к применению артефакты для настройки SourceCraft Code Assistant под проект «Форматировщик документов по ГОСТ».

---

## 1. Глобальные настройки VS Code (settings.json)

**Путь:** `%APPDATA%\Code\User\settings.json`

Добавьте следующие секции в ваш глобальный `settings.json`:

```json
{
    "sourcecraft-code-assist.allowedCommands": [
        "python",
        "python -m pytest",
        "python -m unittest",
        "python gui_formatter.py",
        "python md_to_docx.py",
        "pip install",
        "pip list",
        "cd",
        "dir",
        "type",
        "findstr",
        "echo",
        "git status",
        "git log",
        "git diff",
        "git show",
        "git add",
        "git commit",
        "git push",
        "git pull"
    ],
    "sourcecraft-code-assist.deniedCommands": [
        "del /f /s",
        "rmdir /s /q",
        "format",
        "diskpart",
        "reg delete",
        "shutdown",
        "taskkill /f",
        "net user",
        "net localgroup"
    ],
    "sourcecraft-code-assist.debug": false
}
```

---

## 2. Workspace-настройки VS Code

**Путь:** `.vscode/settings.json` (в корне проекта)

Создайте файл `.vscode/settings.json` со следующим содержимым:

```json
{
    "python.defaultInterpreterPath": "python",
    "python.analysis.typeCheckingMode": "basic",
    "python.analysis.autoImportCompletions": true,
    "python.analysis.completeFunctionParens": true,
    "python.formatting.provider": "none",
    "python.linting.enabled": true,
    "python.linting.pylintEnabled": true,
    "python.linting.flake8Enabled": true,
    "python.linting.mypyEnabled": false,
    "python.linting.pylintArgs": [
        "--max-line-length=120",
        "--disable=C0111,C0103,R0903,R0913,R0914,R0915,W0511"
    ],
    "python.linting.flake8Args": [
        "--max-line-length=120",
        "--ignore=E501,W503,F401"
    ],
    "editor.rulers": [120],
    "editor.formatOnSave": false,
    "files.encoding": "utf8",
    "files.exclude": {
        "**/__pycache__": true,
        "**/*.pyc": true,
        "**/.pytest_cache": true,
        "**/*.egg-info": true,
        "build/": true,
        "dist/": true
    },
    "search.exclude": {
        "build/": true,
        "dist/": true,
        "outputs/": true,
        "data/": true,
        "apply/": true,
        "audit/": true,
        "gui/": true
    }
}
```

---

## 3. Сниппеты для VS Code

**Путь:** `.vscode/gost_formatter.code-snippets`

Создайте файл `.vscode/gost_formatter.code-snippets`:

```json
{
    "AuditIssue creation": {
        "prefix": "auditissue",
        "body": [
            "self._create_issue(",
            "    severity=Severity.${1:WARNING},",
            "    category=\"${2:CATEGORY}\",",
            "    elem_type=\"${3:element_type}\",",
            "    loc={'index': ${4:index}},",
            "    description=\"${5:description}\",",
            "    current_value=\"${6:current}\",",
            "    expected_value=\"${7:expected}\"",
            ")"
        ],
        "description": "Создать объект AuditIssue"
    },
    "Applier method": {
        "prefix": "applier",
        "body": [
            "def apply(self, doc: Document, issue: AuditIssue) -> bool:",
            "    \"\"\"",
            "    Применяет исправление ${1:category} к документу.",
            "",
            "    :param doc: объект документа",
            "    :param issue: проблема для исправления",
            "    :return: True если успешно, False если ошибка",
            "    \"\"\"",
            "    try:",
            "        ${2:# логика применения}",
            "        return True",
            "    except Exception as e:",
            "        logger.error(f\"Ошибка применения исправления {issue.id}: {e}\")",
            "        return False"
        ],
        "description": "Шаблон метода applier"
    },
    "Auditor method": {
        "prefix": "auditor",
        "body": [
            "def _audit_${1:category}(self, doc: Document) -> None:",
            "    \"\"\"",
            "    Аудит ${2:описание категории}.",
            "",
            "    :param doc: объект документа для аудита",
            "    \"\"\"",
            "    ${3:# логика аудита}"
        ],
        "description": "Шаблон метода auditor"
    },
    "Test method": {
        "prefix": "testmethod",
        "body": [
            "def test_${1:name}(self):",
            "    \"\"\"Тест: ${2:описание}.\"\"\"",
            "    # Arrange",
            "    ${3:# подготовка}",
            "",
            "    # Act",
            "    ${4:# действие}",
            "",
            "    # Assert",
            "    self.${5:assertTrue}(${6:condition})"
        ],
        "description": "Шаблон тестового метода"
    },
    "Location dict": {
        "prefix": "location",
        "body": [
            "{",
            "    'index': ${1:index},",
            "    'page_estimate': ${2:page}",
            "}"
        ],
        "description": "Словарь location для параграфа"
    },
    "Location table": {
        "prefix": "loctable",
        "body": [
            "{",
            "    'table_index': ${1:index},",
            "    'cell_row': ${2:row},",
            "    'cell_col': ${3:col}",
            "}"
        ],
        "description": "Словарь location для таблицы"
    },
    "Location header": {
        "prefix": "locheader",
        "body": [
            "{",
            "    'section': ${1:0},",
            "    'paragraph': ${2:0},",
            "    'container': '${3:header}'",
            "}"
        ],
        "description": "Словарь location для колонтитула"
    },
    "New module template": {
        "prefix": "pymodule",
        "body": [
            "#!/usr/bin/env python3",
            "# -*- coding: utf-8 -*-",
            "\"\"\"",
            "${1:Краткое описание модуля}.",
            "",
            "${2:Подробное описание}.",
            "\"\"\"",
            "",
            "import logging",
            "from typing import List, Dict, Any, Optional",
            "",
            "logger = logging.getLogger(__name__)",
            "",
            "",
            "class ${3:ClassName}:",
            "    \"\"\"${4:Описание класса}.\"\"\"",
            "",
            "    def __init__(self, ${5:param}: str):",
            "        \"\"\"",
            "        Инициализация.",
            "",
            "        :param ${5:param}: описание параметра",
            "        \"\"\"",
            "        self.${5:param} = ${5:param}"
        ],
        "description": "Шаблон нового Python-модуля"
    },
    "Config style section": {
        "prefix": "configstyle",
        "body": [
            "${1:StyleName}:",
            "  enabled: true",
            "  font:",
            "    name: \"Times New Roman\"",
            "    size: ${2:12}",
            "    bold: false",
            "    italic: false",
            "  paragraph:",
            "    first_line_indent_cm: ${3:1.25}",
            "    line_spacing: ${4:1.15}",
            "    alignment: \"${5:justify}\""
        ],
        "description": "Шаблон секции стиля в YAML-конфиге"
    },
    "Logger setup": {
        "prefix": "pylogger",
        "body": [
            "import logging",
            "logger = logging.getLogger(__name__)"
        ],
        "description": "Настройка логгера"
    },
    "Test class template": {
        "prefix": "pytestclass",
        "body": [
            "\"\"\"",
            "Тесты для ${1:модуля/класса}.",
            "\"\"\"",
            "",
            "import unittest",
            "import os",
            "import tempfile",
            "",
            "",
            "class Test${2:ClassName}(unittest.TestCase):",
            "    \"\"\"Тесты класса ${2:ClassName}.\"\"\"",
            "",
            "    def setUp(self):",
            "        \"\"\"Подготовка тестового окружения.\"\"\"",
            "        self.temp_dir = tempfile.mkdtemp()",
            "",
            "    def tearDown(self):",
            "        \"\"\"Очистка после тестов.\"\"\"",
            "        import shutil",
            "        shutil.rmtree(self.temp_dir, ignore_errors=True)",
            "",
            "    def test_${3:feature}(self):",
            "        \"\"\"Тест: ${4:описание}.\"\"\"",
            "        # Arrange",
            "        ${5:# подготовка}",
            "",
            "        # Act",
            "        ${6:# действие}",
            "",
            "        # Assert",
            "        self.${7:assertTrue}(${8:condition})"
        ],
        "description": "Шаблон класса с тестами"
    }
}
```

---

## 4. Custom Instructions для SourceCraft

**Куда вставить:** В интерфейсе SourceCraft Code Assistant → настройки → Custom Instructions (системный промпт)

Скопируйте следующий текст в поле Custom Instructions:

```
Ты — SourceCraft Code Assistant для проекта «Форматировщик документов по ГОСТ».

## Язык и стиль
- Отвечай на РУССКОМ языке.
- Пиши код на Python 3.8+.
- Следуй PEP 8: отступы 4 пробела, макс. длина строки 120 символов.
- Docstring на русском, в стиле PEP 257 (тройные двойные кавычки).
- Комментарии в коде — на русском, поясняющие «почему», а не «что».
- Для строк используй одинарные кавычки, для docstring — двойные.

## Конвенции именования
- Классы: PascalCase (MainController, AuditEngine, ApplyOrchestrator)
- Методы/функции: snake_case (scan_document, apply_fixes, _load_config)
- Константы: UPPER_SNAKE_CASE (CM_TO_TWIPS, PT_TO_TWIPS)
- Приватные методы: _префикс
- Пакеты: snake_case (src.core, src.apply, src.audit)
- Enum: PascalCase, значения UPPER_CASE (Severity.INFO, Severity.WARNING)
- Dataclass: PascalCase (AuditIssue)

## Типизация
- Используй typing: List, Dict, Optional, Any, Tuple, Callable
- Аннотации типов для всех публичных методов
- Optional[X] вместо Union[X, None]

## Импорты
- Порядок: стандартная библиотека → сторонние → локальные
- Разделяй группы пустой строкой
- Избегай from module import *

## Архитектура проекта
- src/core/ — ядро: MainController, AuditEngine, ConfigLoader, DocumentModel, docx_utils, StyleManager, PageManager, ReportManager, HighlightEngine, OptimizedApplyEngine
- src/apply/ — применение исправлений: ApplyOrchestrator, BaseApplier, FontApplier, ParagraphApplier, TableApplier, HeaderFooterApplier
- src/audit/ — аудиторы: BaseAuditor, FontAuditor, ParagraphAuditor
- src/gui/ — GUI-компоненты: main_window, actions, file_handlers, table_manager
- configs/active/ — YAML-конфигурации
- tests/unit/ — модульные тесты
- tests/stress/ — нагрузочные тесты

## Ключевые типы данных
- AuditIssue: id, severity (Severity), category (str), element_type (str), location (Dict), description (str), current_value (str), expected_value (str), recommendation (str), auto_fixable (bool), fix_payload (Dict)
- Severity: INFO, WARNING, CRITICAL
- location: {'index': int} для параграфов, {'table_index': int} для таблиц, {'section': int, 'paragraph': int, 'container': 'header'|'footer'} для колонтитулов

## Константы проекта
- CM_TO_TWIPS = 567
- PT_TO_TWIPS = 20
- MM_TO_TWIPS = 56.7
- FONT_SIZE_TOLERANCE_TWIPS = 10
- INDENT_TOLERANCE_TWIPS = 20
- SPACING_TOLERANCE_TWIPS = 20

## Правила генерации кода
- Новые модули: #!/usr/bin/env python3, # -*- coding: utf-8 -*-
- Модели данных: используй @dataclass
- Фиксированные наборы: используй Enum
- Обработка ошибок: явные try/except с логированием
- Логирование: logger = logging.getLogger(__name__)
- Запрещено: global variables, wildcard imports, bare except, mutable default arguments
- Длинные операции: в отдельных потоках с колбэками прогресса
- GUI-обновления: через root.after() в главном потоке

## Тестирование
- Используй unittest.TestCase
- Тестовые методы: test_имя_функциональности
- Шаблон: Arrange → Act → Assert
- Временные файлы: через tempfile.mkdtemp() + tearDown

## Безопасность
- Исходный файл НИКОГДА не изменяется напрямую
- Все операции с документами в try/except
- Критические ошибки логировать с exc_info=True
- Валидация путей перед открытием файлов
```

---

## 5. Обновлённый .codeassistantignore

**Путь:** `.codeassistantignore` (в корне проекта)

Добавьте следующие строки в существующий файл:

```gitignore
# SourceCraft Code Assistant - исключения
.vscode/
.idea/
*.swp
*.swo
*~
.DS_Store
Thumbs.db

# Устаревшие модули (заменены на src/apply/, src/audit/, src/gui/)
apply/
audit/
gui/

# Сборки и выходные данные
build/
dist/
outputs/
*.spec
*.ico

# Данные
data/
*.docx
*.json

# Логи
*.log
formatter.log

# Временные файлы
*.tmp
*.temp
*.bak
*.backup

# Тестовые выходные файлы
test_output*.docx
test_*.docx
test_*.json
test_*.md
```

---

## 6. Пошаговая инструкция по применению

### Шаг 1: Глобальные настройки VS Code
1. Откройте VS Code
2. Нажмите `Ctrl+Shift+P` → "Preferences: Open User Settings (JSON)"
3. Добавьте секции `sourcecraft-code-assist.allowedCommands`, `sourcecraft-code-assist.deniedCommands` и `sourcecraft-code-assist.debug` из п.1
4. Сохраните файл

### Шаг 2: Workspace-настройки
1. Создайте файл `.vscode/settings.json` в корне проекта
2. Скопируйте содержимое из п.2
3. Сохраните

### Шаг 3: Сниппеты
1. Создайте файл `.vscode/gost_formatter.code-snippets` в корне проекта
2. Скопируйте содержимое из п.3
3. Сохраните
4. Перезапустите VS Code (или выполните "Developer: Reload Window")

### Шаг 4: Custom Instructions
1. Откройте интерфейс SourceCraft Code Assistant
2. Найдите настройки "Custom Instructions" или "System Prompt"
3. Скопируйте текст из п.4
4. Сохраните

### Шаг 5: .codeassistantignore
1. Откройте файл `.codeassistantignore` в корне проекта
2. Добавьте содержимое из п.5
3. Сохраните

### Шаг 6: Проверка
1. Откройте любой `.py` файл проекта
2. Начните печатать `auditissue` — должен появиться сниппет
3. Начните печатать `applier` — должен появиться сниппет
4. Проверьте, что команды `python`, `pytest`, `git` доступны
5. Проверьте, что команды `del /f /s`, `format` заблокированы
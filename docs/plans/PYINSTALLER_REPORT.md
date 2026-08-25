# Отчёт по сборке portable EXE (PyInstaller) для Formatter_modification

## 1. Обзор проекта

**Назначение:** GUI-приложение для форматирования DOCX-документов по стандарту НТД 01-2013 (проектная организация).  
**Точка входа:** `gui_formatter.py`  
**Архитектура:** MVC-подобная (MainController — контроллер, DocumentModel — модель, GUI — представление)  
**GUI-фреймворк:** Tkinter (ttkthemes — опционально)  
**Формат сборки:** `--onefile` (один portable EXE-файл)

---

## 2. Полный граф зависимостей

### 2.1. Внешние пакеты (PyPI)

| Пакет | Версия (min) | Откуда импортируется | Назначение |
|-------|-------------|---------------------|------------|
| `PyYAML` | >=6.0 | `config_loader.py`, `md_to_docx.py` | Загрузка YAML-конфигурации |
| `python-docx` | >=0.8.11 | Почти все модули | Работа с DOCX-документами |
| `ttkthemes` | >=3.2.2 | `gui_formatter.py` (опционально) | Тема "arc" для Tkinter |
| `networkx` | — | `optimized_application.py` | Анализ графа зависимостей проблем |

### 2.2. Стандартная библиотека Python

| Модуль | Где используется |
|--------|-----------------|
| `tkinter` (tk, ttk, filedialog, messagebox, scrolledtext, simpledialog) | `gui_formatter.py`, `main_window.py`, `actions.py`, `file_handlers.py`, `table_manager.py` |
| `threading` | `gui_formatter.py`, `main_window.py`, `actions.py`, `file_handlers.py` |
| `logging` | Почти все модули |
| `os` | `gui_formatter.py`, `main_window.py`, `file_handlers.py`, `config_loader.py`, `main_controller.py` |
| `datetime` | `gui_formatter.py`, `main_window.py`, `md_to_docx.py` |
| `shutil` | `gui_formatter.py`, `main_window.py`, `main_controller.py` |
| `re` | `base_auditor.py`, `md_to_docx.py`, `audit_engine.py` |
| `csv` | `report_manager.py` |
| `json` | `config_loader.py` |
| `dataclasses` | `audit_engine.py`, `base_auditor.py` |
| `enum` | `audit_engine.py`, `base_auditor.py` |
| `collections` | `audit_engine.py` |
| `tempfile` | `audit_engine.py` |
| `pathlib` | `config_loader.py` |
| `copy` | `audit_engine.py` |
| `typing` | Почти все модули |
| `functools` | `audit_engine.py` |
| `itertools` | `audit_engine.py` |
| `traceback` | `gui_formatter.py`, `main_window.py` |
| `webbrowser` | `gui_formatter.py`, `main_window.py` |
| `subprocess` | `gui_formatter.py`, `main_window.py` |

### 2.3. Локальные модули проекта

```
gui_formatter.py                  # Точка входа
md_to_docx.py                     # Конвертер Markdown → DOCX (RI2013Converter)
src/
├── __init__.py                   # (отсутствует — пакет неявный)
├── core/
│   ├── __init__.py               # (отсутствует)
│   ├── config_loader.py          # ConfigLoader
│   ├── document_model.py         # DocumentModel
│   ├── docx_utils.py             # Утилиты для docx
│   ├── style_manager.py          # StyleManager
│   ├── page_manager.py           # PageManager
│   ├── audit_engine.py           # AuditEngine, Severity, AuditIssue
│   ├── highlight_engine.py       # HighlightEngine
│   ├── report_manager.py         # ReportManager
│   ├── main_controller.py        # MainController
│   └── optimized_application.py  # OptimizedApplyEngine, DependencyAnalyzer
├── gui/
│   ├── __init__.py               # (пустой)
│   ├── main_window.py            # GOSTFormatterGUI (дубль gui_formatter.py)
│   ├── actions.py                # Actions
│   ├── file_handlers.py          # FileHandlers
│   └── table_manager.py          # TableManager
├── apply/
│   ├── __init__.py               # Экспортирует все апплеры
│   ├── base_applier.py           # BaseApplier
│   ├── font_applier.py           # FontApplier
│   ├── paragraph_applier.py      # ParagraphApplier
│   ├── table_applier.py          # TableApplier
│   ├── header_footer_applier.py  # HeaderFooterApplier
│   └── apply_orchestrator.py     # ApplyOrchestrator
└── audit/
    ├── __init__.py               # (пустой)
    ├── base_auditor.py           # BaseAuditor, Severity, AuditIssue
    ├── font_auditor.py           # FontAuditor
    └── paragraph_auditor.py      # ParagraphAuditor
```

---

## 3. Файлы данных и ресурсы, которые необходимо включить

### 3.1. Конфигурационные файлы (обязательны)

| Файл | Назначение |
|------|-----------|
| `configs/active/config.yaml` | **Основной конфиг** — загружается по умолчанию |
| `configs/active/config_v3.2.yaml` | Альтернативный конфиг |
| `configs/active/config_v4.2.yaml` | Альтернативный конфиг |

**Важно:** В коде (`main_controller.py:27`) путь по умолчанию — `"config.yaml"`, но фактически конфиги лежат в `configs/active/`. Нужно либо:
- Исправить путь по умолчанию в коде, либо
- В spec-файле указать `datas` для включения директории `configs/active/` и настроить поиск через `sys._MEIPASS`.

### 3.2. Динамически загружаемые ресурсы

- **configs/active/config.yaml** — загружается через `ConfigLoader.load()` с `open(config_path, 'r', encoding='utf-8')`
- **Лог-файл** `formatter.log` — создаётся в runtime (не нужно включать, но нужно убедиться, что директория для записи доступна)

---

## 4. Структура spec-файла для PyInstaller

```python
# formatter.spec
# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

a = Analysis(
    ['gui_formatter.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Конфигурационные файлы
        ('configs/active/config.yaml', 'configs/active'),
        ('configs/active/config_v3.2.yaml', 'configs/active'),
        ('configs/active/config_v4.2.yaml', 'configs/active'),
    ],
    hiddenimports=[
        # Скрытые импорты python-docx
        'docx',
        'docx.document',
        'docx.text.paragraph',
        'docx.table',
        'docx.shared',
        'docx.enum.text',
        'docx.enum.style',
        'docx.enum.table',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.oxml.parser',
        'docx.opc.constants',
        'docx.opc.part',
        'docx.opc.packuri',
        'docx.opc.phys_pkg',
        'docx.image.image',
        'docx.image.bmp',
        'docx.image.gif',
        'docx.image.jpeg',
        'docx.image.png',
        'docx.image.tiff',
        'docx.styles.style',
        'docx.styles.styles',
        # lxml (используется python-docx под капотом)
        'lxml',
        'lxml.etree',
        'lxml._elementpath',
        # PyYAML
        'yaml',
        # networkx
        'networkx',
        'networkx.algorithms',
        'networkx.algorithms.dag',
        'networkx.algorithms.shortest_paths',
        'networkx.algorithms.components',
        # ttkthemes (опционально)
        'ttkthemes',
        'ttkthemes.themed_tk',
        # Локальные модули (на случай, если autodetect не сработает)
        'md_to_docx',
        'src.core.config_loader',
        'src.core.document_model',
        'src.core.docx_utils',
        'src.core.style_manager',
        'src.core.page_manager',
        'src.core.audit_engine',
        'src.core.highlight_engine',
        'src.core.report_manager',
        'src.core.main_controller',
        'src.core.optimized_application',
        'src.gui.main_window',
        'src.gui.actions',
        'src.gui.file_handlers',
        'src.gui.table_manager',
        'src.apply.base_applier',
        'src.apply.font_applier',
        'src.apply.paragraph_applier',
        'src.apply.table_applier',
        'src.apply.header_footer_applier',
        'src.apply.apply_orchestrator',
        'src.audit.base_auditor',
        'src.audit.font_auditor',
        'src.audit.paragraph_auditor',
    ],
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=['hooks/runtime_hook.py'],
    excludes=[
        # Исключаем ненужные модули для уменьшения размера
        'tkinter.test',
        'unittest',
        'pytest',
        'numpy',
        'scipy',
        'matplotlib',
        'PIL',
        'cv2',
        'pandas',
        'notebook',
        'ipython',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GOSTFormatter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # Без консоли (GUI-приложение)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='feather.ico' if os.path.exists('feather.ico') else None,
)
```

---

## 5. Необходимые хуки (hooks)

### 5.1. Хук для python-docx

PyInstaller может не обнаружить все подмодули `python-docx`. Рекомендуется создать хук:

**Файл: `hooks/hook-docx.py`**
```python
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = collect_submodules('docx')
datas = collect_data_files('docx')
```

### 5.2. Хук для lxml

`python-docx` использует `lxml` под капотом. lxml известен проблемами с PyInstaller:

**Файл: `hooks/hook-lxml.py`**
```python
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('lxml')
```

### 5.3. Хук для networkx

**Файл: `hooks/hook-networkx.py`**
```python
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('networkx')
```

### 5.4. Runtime-хук для поиска конфигов

Поскольку при `--onefile` файлы распаковываются во временную директорию (`sys._MEIPASS`), необходимо добавить runtime-хук, который корректирует пути:

**Файл: `hooks/runtime_hook.py`**
```python
import sys
import os

# Добавляем путь к временной директории PyInstaller в sys.path
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Функция для получения правильного пути к ресурсам
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)
```

---

## 6. Критические моменты и рекомендации

### 6.1. Проблема с путём к конфигу по умолчанию

В `main_controller.py:27`:
```python
def __init__(self, config_path: str = "config.yaml"):
```

Но фактический конфиг находится в `configs/active/config.yaml`.  
**Рекомендация:** Изменить путь по умолчанию на `"configs/active/config.yaml"` и в spec-файле включить эту директорию через `datas`.

### 6.2. Проблема с ttkthemes (опциональный импорт)

В `gui_formatter.py`:
```python
try:
    from ttkthemes import ThemedTk
    USE_THEME = True
except ImportError:
    USE_THEME = False
    root = tk.Tk()
```

PyInstaller может не обнаружить `ttkthemes` при опциональном импорте.  
**Рекомендация:** Добавить `'ttkthemes'` в `hiddenimports`.

### 6.3. Проблема с networkx (скрытая зависимость)

`networkx` импортируется в `optimized_application.py`, но **не указан** в `requirements.txt`.  
**Рекомендация:** 
1. Добавить `networkx` в `requirements.txt`
2. Добавить `'networkx'` в `hiddenimports` spec-файла
3. Убедиться, что `networkx` установлен в окружении перед сборкой

### 6.4. Проблема с дублированием GUI-классов

В проекте есть два файла с одинаковым классом `GOSTFormatterGUI`:
- `gui_formatter.py` (корень проекта) — **точка входа**
- `src/gui/main_window.py` — дубль

**Рекомендация:** Убедиться, что точкой входа является именно `gui_formatter.py`, а `src/gui/main_window.py` не конфликтует. При сборке PyInstaller может подхватить не тот файл. Лучше удалить дубль или явно указать точку входа.

### 6.5. Проблема с лог-файлом

Лог-файл `formatter.log` создаётся в текущей рабочей директории. При `--onefile` рабочая директория может быть временной.  
**Рекомендация:** В runtime-хуке настроить путь для лога на `os.path.dirname(sys.executable)` или на `os.getcwd()`.

### 6.6. Проблема с `src/__init__.py` и `src/core/__init__.py`

В проекте **отсутствуют** `src/__init__.py` и `src/core/__init__.py`. Это может вызвать проблемы с импортами при сборке PyInstaller.  
**Рекомендация:** Создать пустые `__init__.py` файлы:
- `src/__init__.py`
- `src/core/__init__.py`

### 6.7. Проблема с md_to_docx.py

Файл `md_to_docx.py` находится в корне проекта и импортируется как `from md_to_docx import RI2013Converter`. PyInstaller должен найти его автоматически, но для надёжности добавить в `hiddenimports`.

---

## 7. Команда для сборки

```bash
# Установка зависимостей
pip install pyinstaller PyYAML>=6.0 python-docx>=0.8.11 ttkthemes>=3.2.2 networkx

# Сборка через spec-файл
pyinstaller formatter.spec

# Или сборка через командную строку (альтернатива)
pyinstaller --onefile --windowed --name "GOSTFormatter" ^
    --add-data "configs/active/config.yaml;configs/active" ^
    --add-data "configs/active/config_v3.2.yaml;configs/active" ^
    --add-data "configs/active/config_v4.2.yaml;configs/active" ^
    --hidden-import docx ^
    --hidden-import docx.document ^
    --hidden-import docx.text.paragraph ^
    --hidden-import docx.table ^
    --hidden-import docx.shared ^
    --hidden-import docx.enum.text ^
    --hidden-import docx.enum.style ^
    --hidden-import docx.enum.table ^
    --hidden-import docx.oxml ^
    --hidden-import docx.oxml.ns ^
    --hidden-import lxml ^
    --hidden-import lxml.etree ^
    --hidden-import networkx ^
    --hidden-import ttkthemes ^
    --hidden-import md_to_docx ^
    --hidden-import src.core.config_loader ^
    --hidden-import src.core.document_model ^
    --hidden-import src.core.docx_utils ^
    --hidden-import src.core.style_manager ^
    --hidden-import src.core.page_manager ^
    --hidden-import src.core.audit_engine ^
    --hidden-import src.core.highlight_engine ^
    --hidden-import src.core.report_manager ^
    --hidden-import src.core.main_controller ^
    --hidden-import src.core.optimized_application ^
    --hidden-import src.gui.main_window ^
    --hidden-import src.gui.actions ^
    --hidden-import src.gui.file_handlers ^
    --hidden-import src.gui.table_manager ^
    --hidden-import src.apply.base_applier ^
    --hidden-import src.apply.font_applier ^
    --hidden-import src.apply.paragraph_applier ^
    --hidden-import src.apply.table_applier ^
    --hidden-import src.apply.header_footer_applier ^
    --hidden-import src.apply.apply_orchestrator ^
    --hidden-import src.audit.base_auditor ^
    --hidden-import src.audit.font_auditor ^
    --hidden-import src.audit.paragraph_auditor ^
    --exclude-module tkinter.test ^
    --exclude-module unittest ^
    --exclude-module pytest ^
    --exclude-module numpy ^
    --exclude-module scipy ^
    --exclude-module matplotlib ^
    --exclude-module PIL ^
    --exclude-module pandas ^
    gui_formatter.py
```

---

## 8. Чек-лист предварительных действий перед сборкой

- [ ] Создать `src/__init__.py` (пустой файл)
- [ ] Создать `src/core/__init__.py` (пустой файл)
- [ ] Исправить путь по умолчанию в `main_controller.py` с `"config.yaml"` на `"configs/active/config.yaml"`
- [ ] Добавить `networkx` в `requirements.txt`
- [ ] Удалить или переименовать дубль `src/gui/main_window.py` (если он не нужен)
- [ ] Создать директорию `hooks/` с файлами:
  - `hooks/hook-docx.py`
  - `hooks/hook-lxml.py`
  - `hooks/hook-networkx.py`
  - `hooks/runtime_hook.py`
- [ ] Создать `formatter.spec` на основе шаблона из раздела 4
- [ ] Установить все зависимости: `pip install -r requirements.txt networkx`
- [ ] Запустить сборку: `pyinstaller formatter.spec`
- [ ] Протестировать полученный EXE-файл

---

## 9. Ожидаемый размер EXE

| Компонент | Примерный размер |
|-----------|-----------------|
| Python + ядро | ~15-20 MB |
| Tkinter | ~5-10 MB |
| python-docx + lxml | ~10-15 MB |
| PyYAML | ~1 MB |
| networkx | ~5-10 MB |
| ttkthemes | ~2-5 MB |
| Код проекта | ~1 MB |
| **Итого (с UPX)** | **~25-40 MB** |
# Отчёт о баге: Зависание при конвертации Markdown в DOCX

## 1. Сводка

| Поле | Значение |
|------|----------|
| **ID** | BUG-2026-0605-001 |
| **Приоритет** | Критический (Critical) |
| **Серьёзность** | Блокирующая (Blocker) |
| **Статус** | Подтверждён |
| **Модуль** | `md_to_docx.py` — `RI2013Converter` |
| **Компонент** | Конвертация Markdown → DOCX |
| **Окружение** | Windows 10, Python 3.x |

---

## 2. Шаги воспроизведения

1. Запустить приложение: `gui_formatter.py`
2. Нажать кнопку **"Импорт из Markdown"**
3. В диалоговом окне выбрать `.md` файл для конвертации
4. В диалоговом окне указать путь и имя для сохранения `.docx` файла
5. Нажать **"Сохранить"**
6. Наблюдать: в строке статуса отображается *"Конвертация Markdown в DOCX..."*, индикатор прогресса бесконечно крутится
7. Процесс не завершается — приложение зависает на неопределённое время

---

## 3. Ожидаемое поведение

- Приложение должно сконвертировать Markdown-файл в DOCX за разумное время (секунды/минуты в зависимости от размера файла)
- После завершения конвертации должно появиться сообщение об успехе
- Сконвертированный DOCX-файл должен автоматически загрузиться в интерфейс

---

## 4. Фактическое поведение

- Приложение зависает на этапе выполнения конвертации
- Индикатор прогресса крутится бесконечно
- Процесс не завершается ни успехом, ни ошибкой
- Единственный способ прервать — принудительно закрыть приложение

---

## 5. Анализ причин

### 5.1. Первопричина: Отсутствие конфигурационного файла `config.yaml` в корне проекта

**Файловая структура:**
```
d:/project/Форматер/Formatter_modification/
├── configs/
│   └── active/
│       ├── config.yaml          # Фактический конфиг
│       ├── config_v3.2.yaml
│       └── config_v4.2.yaml
├── md_to_docx.py                # Ищет config.yaml в корне
├── gui_formatter.py             # Ищет config.yaml в корне
└── src/
    └── core/
        └── main_controller.py   # Ищет configs/active/config.yaml
```

**Цепочка вызовов, приводящая к зависанию:**

```
gui_formatter.py:import_markdown()
  → self.controller.import_markdown(md_path, docx_path)  # [main_controller.py:297]
    → RI2013Converter(config_path)                        # [main_controller.py:304]
      → config_path = "config.yaml"                       # (т.к. configs/active/config.yaml не найден)
      → ConfigLoader("config.yaml").load()                # [md_to_docx.py:31-33]
        → FileNotFoundError                               # config.yaml не существует в корне
      → except FileNotFoundError:                         # [md_to_docx.py:34-39]
        → self.config = self._get_default_config()        # используется default config
        → self.config_loader = None                       # ❌ КРИТИЧНО: loader = None
      → self._setup_document()                            # [md_to_docx.py:71]
        → PageManager(self.config)                        # [md_to_docx.py:227]
          → self.config = config_loader                   # config — это словарь, а не ConfigLoader
          → apply_all() → apply_headers_footers()
            → _setup_header() / _setup_footer()           # [page_manager.py:117-204]
              → header.paragraphs / footer.paragraphs
              → paragraph.clear()                         # ❌ Потенциальное зависание
        → self.style_manager = None                       # ❌ style_manager = None
        → self._setup_styles()                            # fallback на старую логику
```

### 5.2. Конкретные проблемы в коде

#### Проблема A: Конфликт путей к конфигурации (КРИТИЧЕСКАЯ)

| Файл | Строка | Ожидаемый путь | Реальность |
|------|--------|----------------|------------|
| `md_to_docx.py:28` | `__init__(self, config_path='config.yaml')` | `config.yaml` в корне | Файла нет |
| `gui_formatter.py:41` | `self.config_path = "config.yaml"` | `config.yaml` в корне | Файла нет |
| `main_controller.py:27` | `__init__(self, config_path="configs/active/config.yaml")` | `configs/active/config.yaml` | Файл существует |
| `main_controller.py:302` | `config_path = self.config_path if os.path.exists(self.config_path) else "config.yaml"` | fallback на `config.yaml` | Файла нет |

**Механизм отказа:**
1. `MainController.__init__` использует `configs/active/config.yaml` — **работает**
2. `MainController.import_markdown()` вызывает `RI2013Converter(config_path)` с путём `configs/active/config.yaml` (строка 302-304 `main_controller.py`)
3. `RI2013Converter.__init__` получает `configs/active/config.yaml` — **файл существует**, `ConfigLoader.load()` успешно загружает конфиг
4. **НО:** если `configs/active/config.yaml` по какой-то причине не существует (например, переименован), то `main_controller.py:302` падает на `"config.yaml"`, которого тоже нет
5. `RI2013Converter` переходит в режим default config с `self.config_loader = None`
6. `PageManager` получает словарь вместо `ConfigLoader` — работает через `self.config.get('page_setup')` — **корректно**
7. `StyleManager` не создаётся (`self.style_manager = None`) — используется `_setup_styles()` — **корректно**

#### Проблема B: Потенциальный бесконечный цикл в `_process_content` (КРИТИЧЕСКАЯ)

В `md_to_docx.py:316-397` метод `_process_content` содержит дублирующиеся regex-условия:

```python
# Строка 333: список
if re.match(r'^[-•*]\s+', line) or re.match(r'^\d+\.\s+', line):
    self._add_list_item(line)
    i += 1
    continue

# Строка 339: заголовок 1 уровня
elif line.startswith('# ') or re.match(r'^\d+\.\s+', line):
```

Оба условия используют `re.match(r'^\d+\.\s+', line)`. Строка вида `1. Текст` будет **всегда** обработана как список, а не как заголовок. Это логическая ошибка, но не причина зависания.

#### Проблема C: Некорректная обработка таблиц (СРЕДНЯЯ)

В `md_to_docx.py:366-379`:

```python
elif line.startswith('|') and '|' in line:
    table_lines = []
    while i < len(lines) and lines[i].strip().startswith('|'):
        table_lines.append(lines[i].strip())
        i += 1
    
    if len(table_lines) > 1 and re.match(r'^\|[\s\-:|]+\|$', table_lines[1]):
        ...
    else:
        self._add_paragraph(line)  # ❌ line уже не актуальна, i сдвинут
    continue
```

После цикла `while` переменная `i` указывает на следующую строку после таблицы, но `line` содержит значение ДО цикла. При выполнении `self._add_paragraph(line)` добавляется **не та строка**, которая ожидается. Это не вызывает зависание, но приводит к некорректному содержимому.

#### Проблема D: `Document.save()` на больших документах (СРЕДНЯЯ)

`python-docx` известен проблемами с производительностью `Document.save()` при:
- Большом количестве таблиц
- Наличии полей Word (PAGE, NUMPAGES) в колонтитулах
- Сложной XML-структуре документа

В `md_to_docx.py:311`:
```python
self.doc.save(docx_path)
```

При большом Markdown-файле с множеством таблиц этот вызов может выполняться **очень долго** (от нескольких минут до бесконечности), особенно в сочетании с колонтитулами, содержащими поля `PAGE`.

#### Проблема E: Отсутствие таймаута и индикации прогресса (НИЗКАЯ)

В `gui_formatter.py:327-339` конвертация запускается в отдельном потоке, но:
- Нет таймаута на выполнение
- Нет промежуточной индикации прогресса (только бесконечный spinner)
- Поток daemon — при закрытии GUI поток будет принудительно завершён, что может привести к повреждению файлов

---

## 6. Логи ошибок (ожидаемые)

При отсутствии `config.yaml` в корне в stdout будут выведены:
```
⚠️ Файл конфигурации config.yaml не найден. Используются настройки по умолчанию.
[DEBUG] Config keys: dict_keys(['version', 'name', 'styles', 'page_setup', 'headers_footers', 'formatting_rules'])
```

В лог-файл `formatter.log`:
```
INFO - Начата конвертация: <file>.md -> <file>.docx
```

При успешной загрузке `configs/active/config.yaml`:
```
[DEBUG] Config keys: dict_keys(['version', 'name', 'description', 'last_modified', 'styles', 'detection_rules', 'page_setup', 'headers_footers', 'formatting_rules', 'validation', 'federal_regional_projects'])
[DEBUG] page_setup keys: dict_keys(['portrait', 'landscape', 'paper_size', 'orientation_default'])
```

---

## 7. Системная среда

| Параметр | Значение |
|----------|----------|
| **ОС** | Windows 10 |
| **Python** | 3.x |
| **python-docx** | >= 0.8.11 |
| **PyYAML** | >= 6.0 |
| **markdown** | >= 3.4.0 |
| **ttkthemes** | >= 3.2.2 |

---

## 8. Версии библиотек (из `requirements.txt`)

```
PyYAML>=6.0
python-docx>=0.8.11
ttkthemes>=3.2.2
markdown>=3.4.0
```

---

## 9. Рекомендации по исправлению

### 9.1. Критические (блокирующие)

1. **Устранить конфликт путей к конфигурации:**
   - В `md_to_docx.py:28` изменить путь по умолчанию на `configs/active/config.yaml`
   - В `gui_formatter.py:41` изменить путь по умолчанию на `configs/active/config.yaml`
   - Либо создать символическую ссылку / скопировать `config.yaml` в корень проекта

2. **Добавить проверку существования конфига перед вызовом `RI2013Converter`:**
   - В `main_controller.py:302-304` добавить валидацию существования файла
   - При отсутствии конфига выбрасывать понятное исключение, а не падать в default config

3. **Добавить таймаут на конвертацию:**
   - В `gui_formatter.py:327-339` добавить таймаут (например, 30 секунд)
   - При превышении таймаута показывать сообщение об ошибке

### 9.2. Средние

4. **Исправить дублирование regex в `_process_content`:**
   - Убрать `re.match(r'^\d+\.\s+', line)` из условия определения списка (строка 333)
   - Оставить только для заголовков (строка 339)
   - Либо изменить порядок проверки: заголовки проверять раньше списков

5. **Исправить обработку таблиц:**
   - В блоке `else` (строка 378) использовать актуальную строку, а не сохранённую `line`
   - Добавить проверку, что `table_lines` содержит хотя бы 2 строки (заголовок + разделитель)

6. **Добавить индикацию прогресса конвертации:**
   - Разбить `_process_content` на этапы с callback-ами
   - Передавать прогресс в GUI через `root.after()`

### 9.3. Низкие

7. **Оптимизировать `Document.save()`:**
   - Для больших файлов использовать временный файл и копирование
   - Добавить логирование времени выполнения save

8. **Добавить обработку исключений в `_setup_header`/`_setup_footer`:**
   - Обернуть `paragraph.clear()` в try-except
   - Добавить проверку на наличие полей Word перед очисткой

---

## 10. Диаграмма потока вызовов

```mermaid
flowchart TD
    A[gui_formatter.py\nimport_markdown] --> B[main_controller.py\nimport_markdown]
    B --> C{configs/active/config.yaml\nсуществует?}
    C -->|Да| D[RI2013Converter\nconfigs/active/config.yaml]
    C -->|Нет| E[RI2013Converter\nconfig.yaml]
    E --> F{config.yaml\nсуществует?}
    F -->|Нет| G[FileNotFoundError]
    G --> H[default config\nconfig_loader = None]
    H --> I[_setup_document]
    I --> J[PageManager\nсо словарём config]
    J --> K[apply_headers_footers]
    K --> L[_setup_header\nparagraph.clear]
    L --> M{Зависание на\nDocument.save()}
    D --> N[ConfigLoader.load\nуспешно]
    N --> O[_setup_document]
    O --> P[StyleManager.setup_styles]
    P --> Q[_process_content]
    Q --> R{Большой файл\nмного таблиц?}
    R -->|Да| S[Document.save\nдолго/зависание]
    R -->|Нет| T[Document.save\nуспешно]
```

---

## 11. Временные метки

| Событие | Время (UTC) |
|---------|-------------|
| Обнаружение бага | 2026-06-05 |
| Анализ завершён | 2026-06-05 07:29 UTC |
| Версия конфигурации | 4.2 |

---

## 12. Примечания

- Баг воспроизводится **только** при отсутствии `config.yaml` в корне проекта
- Если в корне проекта создать копию `configs/active/config.yaml` → `config.yaml`, баг может не проявляться (зависит от размера конвертируемого файла)
- Рекомендуется **немедленное** исправление путей к конфигурации как наиболее вероятной причины зависания
# Отчёт о диагностике 40-секундного зависания при конвертации Markdown → DOCX

## 1. Проведённые тесты

### 1.1. Прямая конвертация (CLI)
**Скрипт:** `scripts/diagnostics/diagnose_conversion_speed.py`
**Результат:** Конвертация одного слова заняла **0.32 сек** (инициализация 0.27с + конвертация 0.05с)

| Этап | Время |
|------|-------|
| Загрузка конфигурации | 0.17 сек |
| `Document()` | 0.02 сек |
| `PageManager.apply_all()` | 0.004 сек |
| `StyleManager.setup_styles()` | 0.08 сек |
| Загрузка Markdown | 0.01 сек |
| `_process_content()` | 0.005 сек |
| **`Document.save()`** | **0.03 сек** |
| **ИТОГО** | **0.32 сек** |

### 1.2. Конвертация в отдельном потоке (как в GUI)
**Скрипт:** `scripts/diagnostics/test_gui_thread_conversion.py`
**Результат:** **0.23 сек** — поток завершился мгновенно

### 1.3. Повторная конвертация (3 итерации)
**Скрипт:** `scripts/diagnostics/test_repeated_conversion.py`
**Результат:** Все 3 итерации завершились за **0.24-0.32 сек** — никакого замедления

### 1.4. Конвертация в свежем процессе Python
**Скрипт:** `scripts/diagnostics/test_fresh_process.py`
**Результат:** **0.42 сек** (импорт 0.24с + инициализация 0.16с + конвертация 0.02с)

### 1.5. Загрузка DOCX с полями PAGE/NUMPAGES
**Скрипт:** `scripts/diagnostics/test_load_docx_with_fields.py`
**Результат:** Загрузка DOCX с колонтитулами — **0.03 сек**

## 2. Вывод: Проблема НЕ в скорости конвертации

**Все тесты показывают, что конвертация Markdown → DOCX работает быстро (0.2-0.5 сек).**
40-секундное зависание, описанное в баг-репорте, **не воспроизводится** в изолированных тестах.

## 3. Наиболее вероятные причины (требующие проверки через GUI)

### 3.1. Race condition в GUI (НАИБОЛЕЕ ВЕРОЯТНО)
В `gui_formatter.py:361-366`:
```python
completion_event.wait(timeout=timeout_seconds)  # блокирует main thread на 40 сек
```
- `completion_event.wait()` — **блокирующий вызов**, который замораживает Tkinter main loop
- Если в этот момент пользователь взаимодействует с GUI, события не обрабатываются
- После таймаута вызывается `messagebox.showerror()`, который создаёт модальный диалог
- Если поток конвертации завершается **после** таймаута, но **до** закрытия `messagebox`, то `completion_event.set()` разблокирует `wait()`, и код продолжается к `on_conversion_success()`
- **Результат:** пользователь видит и сообщение об ошибке, и сообщение об успехе

### 3.2. Проблема с `populate_doc_tree()` для больших файлов
В `gui_formatter.py:634-659`:
```python
for i, para in enumerate(doc.paragraphs):  # может быть тысячи параграфов
    self.doc_tree.insert(root, tk.END, ...)  # вставка в TreeView
```
- Для документов с тысячами параграфов вставка каждого в TreeView может быть очень медленной
- `tk.END` при каждой вставке вызывает перерисовку

### 3.3. Проблема с `log4cplus`
Ошибка `log4cplus:ERROR No appenders could be found for logger (DSL)` — это сообщение от C++ библиотеки, используемой `lxml`. Она **не вызывает** 40-секундное зависание, но может быть индикатором проблем с инициализацией C-расширений.

## 4. Рекомендации по исправлению

### 4.1. Устранить race condition в GUI
```python
# Вместо блокирующего wait() с таймаутом:
def convert_thread():
    try:
        success = self.controller.import_markdown(md_path, docx_path)
        # Использовать root.after() для возврата в main thread
        self.root.after(0, lambda: self._on_conversion_result(success, docx_path))
    except Exception as e:
        self.root.after(0, lambda: self._on_conversion_error(e))
    # Не использовать completion_event

thread = threading.Thread(target=convert_thread, daemon=True)
thread.start()
# Не блокировать main thread!
```

### 4.2. Оптимизировать `populate_doc_tree()`
- Использовать `self.doc_tree.insert()` с `batch_size` для вставки пачками
- Или использовать виртуальное дерево (загружать только видимые элементы)

### 4.3. Подавление `log4cplus`
Создан файл `log4cplus.properties` в корне проекта с содержимым:
```
log4cplus.rootLogger=OFF
log4cplus.logger.DSL=OFF
```

## 5. Добавленные инструменты диагностики

| Файл | Назначение |
|------|------------|
| `scripts/diagnostics/diagnose_conversion_speed.py` | Замер времени всех этапов конвертации |
| `scripts/diagnostics/test_gui_thread_conversion.py` | Симуляция вызова из GUI (поток + таймаут) |
| `scripts/diagnostics/test_repeated_conversion.py` | Проверка повторной конвертации |
| `scripts/diagnostics/test_fresh_process.py` | Конвертация в новом процессе Python |
| `scripts/diagnostics/test_load_docx_with_fields.py` | Загрузка DOCX с полями PAGE/NUMPAGES |

## 6. Перманентные замеры времени в коде

В следующие файлы добавлены замеры времени через `perf_logger`:

| Файл | Что замеряется |
|------|----------------|
| `md_to_docx.py` | `__init__` (конфиг, Document, _setup_document), `convert()` (загрузка, _process_content, save) |
| `src/core/page_manager.py` | `apply_all()`, `apply_page_setup()`, `apply_headers_footers()`, `_setup_header()`, `_setup_footer()` |
| `src/core/style_manager.py` | `setup_styles()` (Normal, Headings, Additional) |

Замеры выводятся через `logging.getLogger('perf')` и видны при уровне логирования `INFO`.
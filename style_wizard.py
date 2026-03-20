# formatter\style_wizard.py
"""
Мастер пошаговой настройки стилей ГОСТ
Помогает не-ИТ специалистам легко конфигурировать стили

Функции:
    - StyleWizard: Класс мастера настройки
    - launch_wizard: Функция запуска мастера

Пресеты:
    - ГОСТ Р 21.101-2020 (СПДС)
    - ГОСТ 7.32-2017 (НИР)
    - ГОСТ 2.105-2019 (ЕСКД)
    - ВУЗ (Типовой)

Версия: 4.0
Дата: 2024-01-15
"""


# =============================================================
# ИМПОРТ НЕОБХОДИМЫХ МОДУЛЕЙ
# =============================================================


# Импорт модуля для работы с операционной системой
import os
# Импорт модуля для работы с JSON файлами
import json
# Импорт библиотеки Tkinter для GUI
import tkinter as tk
# Импорт виджетов ttk для современного интерфейса
from tkinter import ttk, messagebox, colorchooser
# Импорт класса GOSTFormatter из core модуля
from formatter_core import GOSTFormatter


# =============================================================
# СЛОВАРЬ ПРЕДУСТАНОВЛЕННЫХ КОНФИГУРАЦИЙ (ПРЕСЕТЫ)
# =============================================================


# Словарь с 4 предустановленными конфигурациями ГОСТ
GOST_PRESETS = {
    # Пресет 1: Для проектной документации в строительстве
    "ГОСТ Р 21.101-2020 (СПДС)": {
        # Описание пресета
        "description": "Для проектной документации в строительстве",
        # Размер шрифта основного текста
        "main_font_size": 14,
        # Размер шрифта таблиц
        "table_font_size": 12,
        # Межстрочный интервал
        "line_spacing": 1.5,
        # Отступ первой строки
        "first_line_indent": 1.25
    },
    # Пресет 2: Для отчётов по научно-исследовательским работам
    "ГОСТ 7.32-2017 (НИР)": {
        # Описание пресета
        "description": "Для отчётов по научно-исследовательским работам",
        # Размер шрифта основного текста
        "main_font_size": 14,
        # Размер шрифта таблиц
        "table_font_size": 12,
        # Межстрочный интервал
        "line_spacing": 1.5,
        # Отступ первой строки
        "first_line_indent": 1.25
    },
    # Пресет 3: Для конструкторской документации
    "ГОСТ 2.105-2019 (ЕСКД)": {
        # Описание пресета
        "description": "Для конструкторской документации",
        # Размер шрифта основного текста
        "main_font_size": 14,
        # Размер шрифта таблиц
        "table_font_size": 12,
        # Межстрочный интервал
        "line_spacing": 1.5,
        # Отступ первой строки
        "first_line_indent": 1.25
    },
    # Пресет 4: Стандартные требования большинства вузов
    "ВУЗ (Типовой)": {
        # Описание пресета
        "description": "Стандартные требования большинства вузов",
        # Размер шрифта основного текста
        "main_font_size": 14,
        # Размер шрифта таблиц
        "table_font_size": 12,
        # Межстрочный интервал
        "line_spacing": 1.5,
        # Отступ первой строки
        "first_line_indent": 1.25
    }
}


# =============================================================
# КЛАСС МАСТЕРА НАСТРОЙКИ СТИЛЕЙ
# =============================================================


class StyleWizard:
    """
    Пошаговый мастер настройки стилей
    
    Атрибуты:
        parent (tk.Tk): Родительское окно
        config_path (str): Путь к файлу конфигурации
        config (dict): Текущая конфигурация
        current_step (int): Текущий шаг мастера
        widgets (dict): Словарь виджетов для доступа
    
    Шаги мастера:
        1. Выбор стандарта (select_preset)
        2. Основной текст (main_text)
        3. Заголовки (headings)
        4. Таблицы и подписи (tables_captions)
        5. Сохранение (save)
    """
    
    # =============================================================
    # СПИСОК ШАГОВ МАСТЕРА
    # =============================================================
    
    
    # Список кортежей (название шага, идентификатор)
    STEPS = [
        # Шаг 1: Выбор пресета
        ("📋 Выбор стандарта", "select_preset"),
        # Шаг 2: Настройка основного текста
        ("🔤 Основной текст", "main_text"),
        # Шаг 3: Настройка заголовков
        ("📝 Заголовки", "headings"),
        # Шаг 4: Настройка таблиц и подписей
        ("📊 Таблицы и подписи", "tables_captions"),
        # Шаг 5: Сохранение конфигурации
        ("✅ Сохранение", "save")
    ]
    
    # =============================================================
    # МЕТОД ИНИЦИАЛИЗАЦИИ МАСТЕРА
    # =============================================================
    
    
    def __init__(self, parent, config_path):
        """
        Инициализация мастера настройки
        
        Args:
            parent (tk.Tk): Родительское окно GUI
            config_path (str): Путь к файлу конфигурации
        
        Returns:
            None: Метод создаёт окно мастера
        """
        # Сохранение ссылки на родительское окно
        self.parent = parent
        # Сохранение пути к конфигурации
        self.config_path = config_path
        # Загрузка текущей конфигурации
        self.config = self._load_config()
        # Установка начального шага (0 = первый шаг)
        self.current_step = 0
        # Инициализация словаря для хранения виджетов
        self.widgets = {}
        
        # Вызов метода создания окна мастера
        self._create_wizard_window()
    
    
    # =============================================================
    # МЕТОД ЗАГРУЗКИ ТЕКУЩЕЙ КОНФИГУРАЦИИ
    # =============================================================
    
    
    def _load_config(self):
        """
        Загружает текущую конфигурацию из файла
        
        Returns:
            dict: Словарь с конфигурацией или конфиг по умолчанию
        
        Note:
            При ошибке загрузки возвращает конфигурацию по умолчанию
        """
        # Попытка загрузки конфигурации
        try:
            # Открытие файла конфигурации на чтение
            with open(self.config_path, 'r', encoding='utf-8') as f:
                # Загрузка JSON данных из файла
                return json.load(f)
        # Обработка любой ошибки загрузки
        except:
            # Возврат конфигурации по умолчанию
            return self._get_default_config()
    
    
    # =============================================================
    # МЕТОД ПОЛУЧЕНИЯ КОНФИГУРАЦИИ ПО УМОЛЧАНИЮ
    # =============================================================
    
    
    def _get_default_config(self):
        """
        Возвращает конфигурацию по умолчанию
        
        Returns:
            dict: Словарь с минимальной конфигурацией
        
        Note:
            Используется если файл конфигурации не найден
        """
        # Возврат словаря с базовой конфигурацией
        return {
            # Версия конфигурации
            "version": "4.0",
            # Название конфигурации
            "name": "Пользовательская конфигурация",
            # Стили (пустой словарь)
            "styles": {},
            # Правила обнаружения (пустой словарь)
            "detection_rules": {}
        }
    
    
    # =============================================================
    # МЕТОД СОЗДАНИЯ ОКНА МАСТЕРА
    # =============================================================
    
    
    def _create_wizard_window(self):
        """
        Создаёт окно мастера настройки
        
        Returns:
            None: Метод создаёт все элементы окна мастера
        
        Note:
            Окно модальное (блокирует родительское)
        """
        # Создание дочернего окна (Toplevel)
        self.wizard = tk.Toplevel(self.parent)
        # Установка заголовка окна
        self.wizard.title("🎨 Мастер настройки стилей")
        # Установка размера окна
        self.wizard.geometry("750x600")
        # Запрет изменения размера окна
        self.wizard.resizable(False, False)
        # Установка родительского окна
        self.wizard.transient(self.parent)
        # Перехват фокуса (модальное окно)
        self.wizard.grab_set()
        
        # Вызов метода создания заголовка
        self._create_header()
        # Вызов метода создания области контента
        self._create_content_area()
        # Вызов метода создания навигации
        self._create_navigation()
        # Вызов метода отображения текущего шага
        self._show_step()
    
    
    # =============================================================
    # МЕТОД СОЗДАНИЯ ЗАГОЛОВКА МАСТЕРА
    # =============================================================
    
    
    def _create_header(self):
        """
        Создаёт заголовок мастера с индикатором прогресса
        
        Returns:
            None: Метод добавляет виджеты в окно мастера
        """
        # Создание фрейма для заголовка
        header_frame = ttk.Frame(self.wizard, padding=20)
        # Размещение фрейма в верхней части
        header_frame.pack(fill="x")
        
        # Создание метки заголовка
        ttk.Label(header_frame, text="🎨 Мастер настройки стилей ГОСТ", 
                 font=("Segoe UI", 16, "bold")).pack()
        # Создание метки подзаголовка
        ttk.Label(header_frame, text="Пошаговая настройка параметров форматирования", 
                 font=("Segoe UI", 10), foreground="gray").pack()
        
        # Создание фрейма для индикаторов шагов
        self.progress_frame = ttk.Frame(self.wizard, padding=10)
        # Размещение фрейма под заголовком
        self.progress_frame.pack(fill="x")
        
        # Инициализация списка индикаторов шагов
        self.step_indicators = []
        # Перебор всех шагов мастера
        for i, (title, _) in enumerate(self.STEPS):
            # Создание метки индикатора шага
            indicator = ttk.Label(self.progress_frame, text=f"{i+1}. {title.split()[0]}", 
                                 font=("Segoe UI", 9))
            # Размещение индикатора
            indicator.pack(side="left", padx=5)
            # Добавление в список
            self.step_indicators.append(indicator)
        
        # Создание прогресс-бара
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode="determinate", 
                                           length=600)
        # Размещение прогресс-бара
        self.progress_bar.pack(fill="x", pady=10)
    
    
    # =============================================================
    # МЕТОД СОЗДАНИЯ ОБЛАСТИ КОНТЕНТА
    # =============================================================
    
    
    def _create_content_area(self):
        """
        Создаёт область контента для текущего шага
        
        Returns:
            None: Метод создаёт фрейм для динамического контента
        """
        # Создание фрейма для контента
        self.content_frame = ttk.LabelFrame(self.wizard, text="Настройки", padding=20)
        # Размещение фрейма в центре окна
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    
    # =============================================================
    # МЕТОД СОЗДАНИЯ НАВИГАЦИИ
    # =============================================================
    
    
    def _create_navigation(self):
        """
        Создаёт кнопки навигации (Назад, Далее, Готово)
        
        Returns:
            None: Метод добавляет кнопки в нижнюю часть окна
        """
        # Создание фрейма для навигации
        nav_frame = ttk.Frame(self.wizard, padding=20)
        # Размещение фрейма в нижней части
        nav_frame.pack(fill="x")
        
        # Создание кнопки Назад
        self.back_button = ttk.Button(nav_frame, text="← Назад", command=self._prev_step, 
                                      state="disabled")
        # Размещение кнопки слева
        self.back_button.pack(side="left")
        
        # Создание кнопки Далее
        self.next_button = ttk.Button(nav_frame, text="Далее →", command=self._next_step)
        # Размещение кнопки справа
        self.next_button.pack(side="right")
        
        # Создание кнопки Готово
        self.finish_button = ttk.Button(nav_frame, text="✅ Готово", command=self._finish, 
                                        state="disabled")
        # Размещение кнопки справа от Далее
        self.finish_button.pack(side="right", padx=10)
    
    
    # =============================================================
    # МЕТОД ОБНОВЛЕНИЯ НАВИГАЦИИ
    # =============================================================
    
    
    def _update_navigation(self):
        """
        Обновляет состояние кнопок навигации
        
        Returns:
            None: Метод изменяет state кнопок и индикаторов
        """
        # Включение кнопки Назад если не первый шаг
        self.back_button.config(state="normal" if self.current_step > 0 else "disabled")
        
        # Проверка: последний ли шаг
        if self.current_step < len(self.STEPS) - 1:
            # Включение кнопки Далее
            self.next_button.config(state="normal")
            # Отключение кнопки Готово
            self.finish_button.config(state="disabled")
        else:
            # Отключение кнопки Далее
            self.next_button.config(state="disabled")
            # Включение кнопки Готово
            self.finish_button.config(state="normal")
        
        # Обновление индикаторов шагов
        for i, indicator in enumerate(self.step_indicators):
            # Проверка: пройден ли шаг
            if i < self.current_step:
                # Зелёный цвет для пройденных
                indicator.config(foreground="green")
            # Проверка: текущий ли шаг
            elif i == self.current_step:
                # Синий жирный для текущего
                indicator.config(foreground="blue", font=("Segoe UI", 9, "bold"))
            else:
                # Серый для будущих
                indicator.config(foreground="gray")
        
        # Расчёт процента прогресса
        progress = int((self.current_step + 1) / len(self.STEPS) * 100)
        # Установка значения прогресс-бара
        self.progress_bar['value'] = progress
    
    
    # =============================================================
    # МЕТОД ОТОБРАЖЕНИЯ ТЕКУЩЕГО ШАГА
    # =============================================================
    
    
    def _show_step(self):
        """
        Отображает контент текущего шага мастера
        
        Returns:
            None: Метод очищает и заполняет content_frame
        """
        # Очистка всех виджетов из фрейма контента
        for widget in self.content_frame.winfo_children():
            # Удаление каждого виджета
            widget.destroy()
        
        # Получение имени текущего шага
        step_name = self.STEPS[self.current_step][1]
        
        # Проверка: какой шаг отображать
        if step_name == "select_preset":
            # Отображение шага выбора пресета
            self._show_preset_step()
        elif step_name == "main_text":
            # Отображение шага основного текста
            self._show_main_text_step()
        elif step_name == "headings":
            # Отображение шага заголовков
            self._show_headings_step()
        elif step_name == "tables_captions":
            # Отображение шага таблиц
            self._show_tables_captions_step()
        elif step_name == "save":
            # Отображение шага сохранения
            self._show_save_step()
        
        # Обновление кнопок навигации
        self._update_navigation()
    
    
    # =============================================================
    # МЕТОД ОТОБРАЖЕНИЯ ШАГА ВЫБОРА ПРЕСЕТА
    # =============================================================
    
    
    def _show_preset_step(self):
        """
        Шаг 1: Выбор готового пресета ГОСТ
        
        Returns:
            None: Метод добавляет radio buttons с пресетами
        """
        # Создание метки инструкции
        ttk.Label(self.content_frame, text="Выберите готовый стандарт или настройте вручную:", 
                 font=("Segoe UI", 11)).pack(anchor="w", pady=10)
        
        # Создание переменной для хранения выбранного пресета
        self.preset_var = tk.StringVar(value="ГОСТ Р 21.101-2020 (СПДС)")
        
        # Перебор всех пресетов
        for preset_name, preset_data in GOST_PRESETS.items():
            # Создание фрейма для каждого пресета
            frame = ttk.Frame(self.content_frame, padding=5)
            # Размещение фрейма
            frame.pack(fill="x", pady=2)
            
            # Создание radio button
            ttk.Radiobutton(frame, text=preset_name, variable=self.preset_var, 
                           value=preset_name).pack(side="left")
            # Создание метки описания
            ttk.Label(frame, text=f"  ({preset_data['description']})", 
                     foreground="gray").pack(side="left")
        
        # Создание метки подсказки
        ttk.Label(self.content_frame, text="\n💡 Выбранный пресет можно будет изменить на следующих шагах", 
                 font=("Segoe UI", 9), foreground="blue").pack(anchor="w", pady=10)
    
    
    # =============================================================
    # МЕТОД ОТОБРАЖЕНИЯ ШАГА ОСНОВНОГО ТЕКСТА
    # =============================================================
    
    
    def _show_main_text_step(self):
        """
        Шаг 2: Настройка параметров основного текста
        
        Returns:
            None: Метод добавляет поля для шрифта и абзаца
        """
        # Создание метки заголовка шага
        ttk.Label(self.content_frame, text="Параметры основного текста:", 
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=10)
        
        # =========================================================
        # НАСТРОЙКИ ШРИФТА
        # =========================================================
        
        # Создание фрейма для настроек шрифта
        font_frame = ttk.LabelFrame(self.content_frame, text="Шрифт", padding=10)
        # Размещение фрейма
        font_frame.pack(fill="x", pady=5)
        
        # Создание метки названия шрифта
        ttk.Label(font_frame, text="Название:").grid(row=0, column=0, sticky="w", pady=5)
        # Создание combobox с выбором шрифта
        self.main_font_name = ttk.Combobox(font_frame, values=["Times New Roman", "Arial", "Calibri"], 
                                           width=30)
        # Установка значения по умолчанию
        self.main_font_name.set(self.config.get("styles", {}).get("Normal", {}).get("font", {}).get("name", "Times New Roman"))
        # Размещение combobox
        self.main_font_name.grid(row=0, column=1, padx=10, pady=5)
        
        # Создание метки размера шрифта
        ttk.Label(font_frame, text="Размер (пт):").grid(row=1, column=0, sticky="w", pady=5)
        # Создание spinbox для размера
        self.main_font_size = ttk.Spinbox(font_frame, from_=8, to=24, width=10)
        # Установка значения по умолчанию
        self.main_font_size.set(self.config.get("styles", {}).get("Normal", {}).get("font", {}).get("size", 14))
        # Размещение spinbox
        self.main_font_size.grid(row=1, column=1, padx=10, pady=5)
        
        # =========================================================
        # НАСТРОЙКИ АБЗАЦА
        # =========================================================
        
        # Создание фрейма для настроек абзаца
        para_frame = ttk.LabelFrame(self.content_frame, text="Абзац", padding=10)
        # Размещение фрейма
        para_frame.pack(fill="x", pady=5)
        
        # Создание метки отступа первой строки
        ttk.Label(para_frame, text="Отступ первой строки (см):").grid(row=0, column=0, sticky="w", pady=5)
        # Создание spinbox для отступа
        self.main_first_indent = ttk.Spinbox(para_frame, from_=0, to=5, increment=0.25, width=10)
        # Установка значения по умолчанию
        self.main_first_indent.set(self.config.get("styles", {}).get("Normal", {}).get("paragraph", {}).get("first_line_indent_cm", 1.25))
        # Размещение spinbox
        self.main_first_indent.grid(row=0, column=1, padx=10, pady=5)
        
        # Создание метки межстрочного интервала
        ttk.Label(para_frame, text="Межстрочный интервал:").grid(row=1, column=0, sticky="w", pady=5)
        # Создание combobox для интервала
        self.main_line_spacing = ttk.Combobox(para_frame, values=["1.0", "1.15", "1.5", "2.0"], width=10)
        # Установка значения по умолчанию
        self.main_line_spacing.set(str(self.config.get("styles", {}).get("Normal", {}).get("paragraph", {}).get("line_spacing", 1.5)))
        # Размещение combobox
        self.main_line_spacing.grid(row=1, column=1, padx=10, pady=5)
        
        # Создание метки выравнивания
        ttk.Label(para_frame, text="Выравнивание:").grid(row=2, column=0, sticky="w", pady=5)
        # Создание combobox для выравнивания
        self.main_alignment = ttk.Combobox(para_frame, values=["left", "center", "right", "justify"], width=10)
        # Установка значения по умолчанию
        self.main_alignment.set(self.config.get("styles", {}).get("Normal", {}).get("paragraph", {}).get("alignment", "justify"))
        # Размещение combobox
        self.main_alignment.grid(row=2, column=1, padx=10, pady=5)
    
    
    # =============================================================
    # МЕТОД ОТОБРАЖЕНИЯ ШАГА ЗАГОЛОВКОВ
    # =============================================================
    
    
    def _show_headings_step(self):
        """
        Шаг 3: Настройка заголовков (4 уровня)
        
        Returns:
            None: Метод добавляет вкладки для каждого уровня
        """
        # Создание метки заголовка шага
        ttk.Label(self.content_frame, text="Параметры заголовков:", 
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=10)
        
        # Создание notebook (вкладок) для заголовков
        notebook = ttk.Notebook(self.content_frame)
        # Размещение notebook
        notebook.pack(fill="both", expand=True)
        
        # Инициализация словаря виджетов заголовков
        self.heading_widgets = {}
        
        # Перебор 4 уровней заголовков
        for i in range(1, 5):
            # Создание фрейма для уровня
            frame = ttk.Frame(notebook, padding=10)
            # Добавление вкладки
            notebook.add(frame, text=f"Уровень {i}")
            
            # Формирование имени стиля
            style_name = f"Heading {i}"
            # Получение конфигурации стиля
            style_config = self.config.get("styles", {}).get(style_name, {})
            
            # Чекбокс включения стиля
            enabled_var = tk.BooleanVar(value=style_config.get("enabled", True))
            ttk.Checkbutton(frame, text="Включить стиль", variable=enabled_var).pack(anchor="w", pady=5)
            self.heading_widgets[f"h{i}_enabled"] = enabled_var
            
            # Чекбокс жирного начертания
            bold_var = tk.BooleanVar(value=style_config.get("font", {}).get("bold", True))
            ttk.Checkbutton(frame, text="Жирный", variable=bold_var).pack(anchor="w", pady=2)
            self.heading_widgets[f"h{i}_bold"] = bold_var
            
            # Чекбокс всех заглавных
            caps_var = tk.BooleanVar(value=style_config.get("font", {}).get("all_caps", i <= 2))
            ttk.Checkbutton(frame, text="Все заглавные буквы", variable=caps_var).pack(anchor="w", pady=2)
            self.heading_widgets[f"h{i}_caps"] = caps_var
            
            # Поле отступа до
            ttk.Label(frame, text="Отступ до (пт):").pack(anchor="w", pady=(10, 2))
            space_before = ttk.Spinbox(frame, from_=0, to=50, width=10)
            space_before.set(style_config.get("paragraph", {}).get("space_before_pt", 12))
            space_before.pack(anchor="w")
            self.heading_widgets[f"h{i}_space_before"] = space_before
            
            # Поле отступа после
            ttk.Label(frame, text="Отступ после (пт):").pack(anchor="w", pady=(5, 2))
            space_after = ttk.Spinbox(frame, from_=0, to=50, width=10)
            space_after.set(style_config.get("paragraph", {}).get("space_after_pt", 6))
            space_after.pack(anchor="w")
            self.heading_widgets[f"h{i}_space_after"] = space_after
    
    
    # =============================================================
    # МЕТОД ОТОБРАЖЕНИЯ ШАГА ТАБЛИЦ И ПОДПИСЕЙ
    # =============================================================
    
    
    def _show_tables_captions_step(self):
        """
        Шаг 4: Настройка таблиц и подписей
        
        Returns:
            None: Метод добавляет поля для таблиц и правил
        """
        # Создание метки заголовка шага
        ttk.Label(self.content_frame, text="Параметры таблиц и подписей:", 
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=10)
        
        # =========================================================
        # НАСТРОЙКИ ТАБЛИЦ
        # =========================================================
        
        # Создание фрейма для таблиц
        table_frame = ttk.LabelFrame(self.content_frame, text="Таблицы", padding=10)
        # Размещение фрейма
        table_frame.pack(fill="x", pady=5)
        
        # Метка размера шрифта таблиц
        ttk.Label(table_frame, text="Размер шрифта (пт):").grid(row=0, column=0, sticky="w", pady=5)
        # Spinbox размера шрифта
        self.table_font_size = ttk.Spinbox(table_frame, from_=8, to=24, width=10)
        # Установка значения
        self.table_font_size.set(self.config.get("styles", {}).get("Table Grid", {}).get("font", {}).get("size", 12))
        # Размещение spinbox
        self.table_font_size.grid(row=0, column=1, padx=10, pady=5)
        
        # =========================================================
        # НАСТРОЙКИ ПОДПИСЕЙ
        # =========================================================
        
        # Создание фрейма для подписей
        caption_frame = ttk.LabelFrame(self.content_frame, text="Подписи (таблицы, рисунки)", padding=10)
        # Размещение фрейма
        caption_frame.pack(fill="x", pady=5)
        
        # Метка размера шрифта подписей
        ttk.Label(caption_frame, text="Размер шрифта (пт):").grid(row=0, column=0, sticky="w", pady=5)
        # Spinbox размера шрифта
        self.caption_font_size = ttk.Spinbox(caption_frame, from_=8, to=24, width=10)
        # Установка значения
        self.caption_font_size.set(self.config.get("styles", {}).get("Caption", {}).get("font", {}).get("size", 12))
        # Размещение spinbox
        self.caption_font_size.grid(row=0, column=1, padx=10, pady=5)
        
        # Чекбокс курсива
        italic_var = tk.BooleanVar(value=self.config.get("styles", {}).get("Caption", {}).get("font", {}).get("italic", True))
        ttk.Checkbutton(caption_frame, text="Курсив", variable=italic_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        self.caption_italic = italic_var
        
        # =========================================================
        # ПРАВИЛА ОБНАРУЖЕНИЯ
        # =========================================================
        
        # Создание фрейма для правил
        rules_frame = ttk.LabelFrame(self.content_frame, text="Правила обнаружения стилей", padding=10)
        # Размещение фрейма
        rules_frame.pack(fill="x", pady=5)
        
        # Метка для разделов
        ttk.Label(rules_frame, text="Заголовки разделов начинаются с:").pack(anchor="w", pady=2)
        # Поле ввода для разделов
        self.rule_section = ttk.Entry(rules_frame, width=50)
        # Установка значения
        self.rule_section.insert(0, "РАЗДЕЛ")
        # Размещение поля
        self.rule_section.pack(anchor="w", pady=2)
        
        # Метка для глав
        ttk.Label(rules_frame, text="Заголовки глав начинаются с:").pack(anchor="w", pady=2)
        # Поле ввода для глав
        self.rule_chapter = ttk.Entry(rules_frame, width=50)
        # Установка значения
        self.rule_chapter.insert(0, "ГЛАВА")
        # Размещение поля
        self.rule_chapter.pack(anchor="w", pady=2)
        
        # Метка для подписей
        ttk.Label(rules_frame, text="Подписи начинаются с (через запятую):").pack(anchor="w", pady=2)
        # Поле ввода для подписей
        self.rule_caption = ttk.Entry(rules_frame, width=50)
        # Установка значения
        self.rule_caption.insert(0, "Таблица, Рисунок, Примечание")
        # Размещение поля
        self.rule_caption.pack(anchor="w", pady=2)
    
    
    # =============================================================
    # МЕТОД ОТОБРАЖЕНИЯ ШАГА СОХРАНЕНИЯ
    # =============================================================
    
    
    def _show_save_step(self):
        """
        Шаг 5: Предпросмотр и сохранение конфигурации
        
        Returns:
            None: Метод показывает превью и кнопку сохранения
        """
        # Создание метки заголовка шага
        ttk.Label(self.content_frame, text="Сохранение конфигурации:", 
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=10)
        
        # =========================================================
        # ПРЕДПРОСМОТР
        # =========================================================
        
        # Создание фрейма для предпросмотра
        preview_frame = ttk.LabelFrame(self.content_frame, text="Предпросмотр конфигурации", padding=10)
        # Размещение фрейма
        preview_frame.pack(fill="both", expand=True, pady=5)
        
        # Создание текстового поля для предпросмотра
        self.preview_text = tk.Text(preview_frame, height=15, wrap="word")
        # Размещение поля
        self.preview_text.pack(fill="both", expand=True)
        
        # Генерация предпросмотра
        self._generate_preview()
        
        # =========================================================
        # ПУТЬ СОХРАНЕНИЯ
        # =========================================================
        
        # Создание фрейма для пути
        path_frame = ttk.Frame(self.content_frame, padding=5)
        # Размещение фрейма
        path_frame.pack(fill="x", pady=10)
        
        # Метка пути
        ttk.Label(path_frame, text="Сохранить в:").pack(side="left")
        # Поле ввода пути
        self.save_path = ttk.Entry(path_frame, width=60)
        # Установка значения
        self.save_path.insert(0, self.config_path)
        # Размещение поля
        self.save_path.pack(side="left", padx=10)
        # Кнопка обзора
        ttk.Button(path_frame, text="Обзор...", command=self._browse_save_path).pack(side="left")
    
    
    # =============================================================
    # МЕТОД ГЕНЕРАЦИИ ПРЕДПРОСМОТРА
    # =============================================================
    
    
    def _generate_preview(self):
        """
        Генерирует текстовый предпросмотр конфигурации
        
        Returns:
            None: Метод вставляет текст в preview_text
        """
        # Очистка текстового поля
        self.preview_text.delete("1.0", tk.END)
        
        # Словарь с данными для предпросмотра
        preview = {
            # Секция основного текста
            "Основной текст": {
                "Шрифт": f"{self.main_font_name.get()}, {self.main_font_size.get()} пт",
                "Интервал": self.main_line_spacing.get(),
                "Отступ": f"{self.main_first_indent.get()} см"
            },
            # Секция заголовков
            "Заголовки": "Настроены для уровней 1-4",
            # Секция таблиц
            "Таблицы": f"Шрифт {self.table_font_size.get()} пт",
            # Секция подписей
            "Подписи": f"Шрифт {self.caption_font_size.get()} пт" + (" (курсив)" if self.caption_italic.get() else "")
        }
        
        # Перебор секций предпросмотра
        for section, settings in preview.items():
            # Вставка заголовка секции
            self.preview_text.insert(tk.END, f"📌 {section}\n", "bold")
            # Проверка: словарь или строка
            if isinstance(settings, dict):
                # Перебор настроек
                for key, value in settings.items():
                    # Вставка настройки
                    self.preview_text.insert(tk.END, f"   • {key}: {value}\n")
            else:
                # Вставка строки
                self.preview_text.insert(tk.END, f"   {settings}\n")
            # Добавление пустой строки
            self.preview_text.insert(tk.END, "\n")
        
        # Настройка тега bold
        self.preview_text.tag_config("bold", font=("Segoe UI", 10, "bold"))
        # Установка только для чтения
        self.preview_text.config(state="disabled")
    
    
    # =============================================================
    # МЕТОД ВЫБОРА ПУТИ СОХРАНЕНИЯ
    # =============================================================
    
    
    def _browse_save_path(self):
        """
        Открывает диалог выбора пути сохранения
        
        Returns:
            None: Метод устанавливает путь в save_path
        """
        # Открытие диалога сохранения
        filename = filedialog.asksaveasfilename(
            # Заголовок диалога
            title="Сохранить конфигурацию",
            # Расширение по умолчанию
            defaultextension=".json",
            # Фильтр файлов
            filetypes=[("JSON файлы", "*.json")]
        )
        # Проверка: выбран ли путь
        if filename:
            # Очистка поля
            self.save_path.delete(0, tk.END)
            # Установка пути
            self.save_path.insert(0, filename)
    
    
    # =============================================================
    # МЕТОД ПЕРЕХОДА К ПРЕДЫДУЩЕМУ ШАГУ
    # =============================================================
    
    
    def _prev_step(self):
        """
        Переход к предыдущему шагу мастера
        
        Returns:
            None: Метод уменьшает current_step и обновляет
        """
        # Проверка: не первый ли шаг
        if self.current_step > 0:
            # Уменьшение номера шага
            self.current_step -= 1
            # Обновление отображения
            self._show_step()
    
    
    # =============================================================
    # МЕТОД ПЕРЕХОДА К СЛЕДУЮЩЕМУ ШАГУ
    # =============================================================
    
    
    def _next_step(self):
        """
        Переход к следующему шагу мастера
        
        Returns:
            None: Метод увеличивает current_step и обновляет
        """
        # Проверка: не последний ли шаг
        if self.current_step < len(self.STEPS) - 1:
            # Увеличение номера шага
            self.current_step += 1
            # Обновление отображения
            self._show_step()
    
    
    # =============================================================
    # МЕТОД ЗАВЕРШЕНИЯ МАСТЕРА
    # =============================================================
    
    
    def _finish(self):
        """
        Завершение мастера и сохранение конфигурации
        
        Returns:
            None: Метод собирает данные и сохраняет в JSON
        """
        # Сбор данных конфигурации
        new_config = self._collect_config_data()
        
        # Получение пути сохранения
        save_path = self.save_path.get()
        
        # Попытка сохранения
        try:
            # Открытие файла на запись
            with open(save_path, 'w', encoding='utf-8') as f:
                # Запись JSON с поддержкой кириллицы
                json.dump(new_config, f, ensure_ascii=False, indent=2)
            
            # Показ сообщения об успехе
            messagebox.showinfo("Успех", 
                               f"✅ Конфигурация сохранена!\n\n"
                               f"📁 Путь: {save_path}\n\n"
                               f"Теперь вы можете использовать эти настройки для форматирования документов.")
            
            # Закрытие окна мастера
            self.wizard.destroy()
            
        # Обработка ошибки сохранения
        except Exception as e:
            # Показ сообщения об ошибке
            messagebox.showerror("Ошибка", f"Не удалось сохранить конфигурацию:\n{e}")
    
    
    # =============================================================
    # МЕТОД СБОРА ДАННЫХ КОНФИГУРАЦИИ
    # =============================================================
    
    
    def _collect_config_data(self):
        """
        Собирает всю конфигурацию из данных мастера
        
        Returns:
            dict: Полный словарь конфигурации для сохранения
        """
        # Импорт datetime для метки времени
        from datetime import datetime
        
        # Создание словаря конфигурации
        config = {
            # Версия
            "version": "4.0",
            # Название из выбранного пресета
            "name": self.preset_var.get(),
            # Описание из пресета
            "description": GOST_PRESETS.get(self.preset_var.get(), {}).get("description", ""),
            # Дата изменения
            "last_modified": str(datetime.now().date()),
            # Стили
            "styles": {
                # Normal
                "Normal": {
                    "enabled": True,
                    "font": {
                        "name": self.main_font_name.get(),
                        "size": int(self.main_font_size.get()),
                        "bold": False,
                        "italic": False,
                        "all_caps": False
                    },
                    "paragraph": {
                        "first_line_indent_cm": float(self.main_first_indent.get()),
                        "line_spacing": float(self.main_line_spacing.get()),
                        "space_before_pt": 0,
                        "space_after_pt": 0,
                        "alignment": self.main_alignment.get()
                    }
                },
                # Обычный
                "Обычный": {
                    "enabled": True,
                    "font": {
                        "name": self.main_font_name.get(),
                        "size": int(self.main_font_size.get()),
                        "bold": False,
                        "italic": False,
                        "all_caps": False
                    },
                    "paragraph": {
                        "first_line_indent_cm": float(self.main_first_indent.get()),
                        "line_spacing": float(self.main_line_spacing.get()),
                        "space_before_pt": 0,
                        "space_after_pt": 0,
                        "alignment": self.main_alignment.get()
                    }
                },
                # Heading 1
                "Heading 1": {
                    "enabled": self.heading_widgets.get("h1_enabled", tk.BooleanVar(value=True)).get(),
                    "font": {
                        "name": self.main_font_name.get(),
                        "size": int(self.main_font_size.get()),
                        "bold": self.heading_widgets.get("h1_bold", tk.BooleanVar(value=True)).get(),
                        "italic": False,
                        "all_caps": self.heading_widgets.get("h1_caps", tk.BooleanVar(value=True)).get()
                    },
                    "paragraph": {
                        "first_line_indent_cm": 0,
                        "line_spacing": 1.0,
                        "space_before_pt": int(self.heading_widgets.get("h1_space_before", ttk.Spinbox()).get()),
                        "space_after_pt": int(self.heading_widgets.get("h1_space_after", ttk.Spinbox()).get()),
                        "alignment": "left"
                    }
                },
                # Caption
                "Caption": {
                    "enabled": True,
                    "font": {
                        "name": self.main_font_name.get(),
                        "size": int(self.caption_font_size.get()),
                        "bold": False,
                        "italic": self.caption_italic.get(),
                        "all_caps": False
                    },
                    "paragraph": {
                        "first_line_indent_cm": 0,
                        "line_spacing": 1.0,
                        "space_before_pt": 6,
                        "space_after_pt": 12,
                        "alignment": "center"
                    }
                },
                # Table Grid
                "Table Grid": {
                    "enabled": True,
                    "font": {
                        "name": self.main_font_name.get(),
                        "size": int(self.table_font_size.get()),
                        "bold": False,
                        "italic": False,
                        "all_caps": False
                    },
                    "paragraph": {
                        "first_line_indent_cm": 0,
                        "line_spacing": 1.0,
                        "space_before_pt": 0,
                        "space_after_pt": 0,
                        "alignment": "center"
                    }
                }
            },
            # Правила обнаружения
            "detection_rules": {
                "Heading 1": [self.rule_section.get()],
                "Heading 2": [self.rule_chapter.get()],
                "Heading 3": ["^\\d+\\.\\d+\\s"],
                "Caption": [s.strip() for s in self.rule_caption.get().split(",")],
                "List Paragraph": ["^\\d+\\.\\s", "•", "–", "−"]
            }
        }
        
        # Добавление заголовков 2-4
        for i in range(2, 5):
            # Имя стиля
            style_name = f"Heading {i}"
            # Добавление в конфиг
            config["styles"][style_name] = {
                "enabled": self.heading_widgets.get(f"h{i}_enabled", tk.BooleanVar(value=True)).get(),
                "font": {
                    "name": self.main_font_name.get(),
                    "size": int(self.main_font_size.get()),
                    "bold": self.heading_widgets.get(f"h{i}_bold", tk.BooleanVar(value=True)).get(),
                    "italic": False,
                    "all_caps": self.heading_widgets.get(f"h{i}_caps", tk.BooleanVar(value=i<=2)).get()
                },
                "paragraph": {
                    "first_line_indent_cm": 1.25 if i == 4 else 0,
                    "line_spacing": 1.0,
                    "space_before_pt": int(self.heading_widgets.get(f"h{i}_space_before", ttk.Spinbox()).get()),
                    "space_after_pt": int(self.heading_widgets.get(f"h{i}_space_after", ttk.Spinbox()).get()),
                    "alignment": "left"
                }
            }
        
        # Возврат собранной конфигурации
        return config


# =============================================================
# ФУНКЦИЯ ЗАПУСКА МАСТЕРА
# =============================================================


def launch_wizard(parent, config_path):
    """
    Запускает мастер настройки
    
    Args:
        parent (tk.Tk): Родительское окно
        config_path (str): Путь к конфигурации
    
    Returns:
        None: Функция создаёт и показывает мастер
    """
    # Создание объекта мастера
    wizard = StyleWizard(parent, config_path)
    # Ожидание закрытия окна мастера
    parent.wait_window(wizard.wizard)


# =============================================================
# ТОЧКА ВХОДА ПРИ ЗАПУСКЕ КАК СКРИПТ
# =============================================================


if __name__ == "__main__":
    # Создание корневого окна
    root = tk.Tk()
    # Скрытие корневого окна
    root.withdraw()
    # Получение пути к конфигурации
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    # Запуск мастера
    launch_wizard(root, config_path)
    # Уничтожение окна
    root.destroy()

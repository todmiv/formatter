# formatter\formatter_core.py
"""
Ядро форматировщика документов по ГОСТ
Этот модуль содержит основную логику применения стилей к документам DOCX

Функции:
    - format_file(): Форматирование файла с проверкой структуры
    - validate_file(): Проверка файла на соответствие структуре
    - save_config(): Сохранение конфигурации
    - GOSTFormatter: Класс для работы с документами ГОСТ

Версия: 4.0
Дата: 2024-01-15
"""


# =============================================================
# ИМПОРТ НЕОБХОДИМЫХ МОДУЛЕЙ
# =============================================================


# Импорт модуля для работы с операционной системой (пути, файлы)
import os
# Импорт модуля для работы с регулярными выражениями (поиск паттернов)
import re
# Импорт модуля для работы с JSON файлами (конфигурация)
import json
# Импорт модуля для копирования файлов (создание бэкапов)
import shutil
# Импорт модуля для логгирования событий
import logging
# Импорт модуля для работы с датой и временем
from datetime import datetime
# Импорт класса Document из python-docx для работы с DOCX файлами
from docx import Document
# Импорт классов для работы с размерами (пункты, сантиметры)
from docx.shared import Pt, Cm
# Импорт перечисления для выравнивания текста
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Импорт перечисления для типов стилей
from docx.enum.style import WD_STYLE_TYPE
# Импорт функции для работы с XML пространствами имен (кириллица)
from docx.oxml.ns import qn


# =============================================================
# НАСТРОЙКА ЛОГГЕРА ДЛЯ ЗАПИСИ СОБЫТИЙ
# =============================================================


# Настройка базовой конфигурации логгера
logging.basicConfig(
    # Уровень логгирования (INFO показывает основную информацию)
    level=logging.INFO,
    # Формат сообщений логов (дата, уровень, сообщение)
    format='%(asctime)s - %(levelname)s - %(message)s',
    # Обработчики логов (файл и консоль)
    handlers=[
        # Запись логов в файл formatter.log с кодировкой UTF-8
        logging.FileHandler('formatter.log', encoding='utf-8', mode='a'),
        # Вывод логов в консоль (стандартный поток)
        logging.StreamHandler()
    ]
)
# Создание объекта логгера для текущего модуля
logger = logging.getLogger(__name__)


# =============================================================
# СЛОВАРЬ ДЛЯ ПРЕОБРАЗОВАНИЯ ВЫРАВНИВАНИЯ ИЗ СТРОКИ В КОНСТАНТУ
# =============================================================


# Словарь маппинга строковых значений выравнивания в константы docx
ALIGNMENT_MAP = {
    # Выравнивание по левому краю
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    # Выравнивание по центру
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    # Выравнивание по правому краю
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    # Выравнивание по ширине (justify)
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
}


# =============================================================
# КЛАСС GOSTFormatter - ОСНОВНОЙ КЛАСС ДЛЯ РАБОТЫ С ДОКУМЕНТАМИ
# =============================================================


class GOSTFormatter:
    """
    Класс для форматирования документов по ГОСТ
    
    Атрибуты:
        config_path (str): Путь к файлу конфигурации JSON
        config (dict): Словарь с загруженной конфигурацией
    
    Методы:
        __init__(): Инициализация с загрузкой конфигурации
        _load_config(): Загрузка JSON конфигурации
        _get_alignment(): Преобразование строки выравнивания в константу
        set_font_for_cyrillic(): Настройка шрифта для кириллицы
        apply_style_properties(): Применение настроек к стилю
        clear_direct_formatting(): Очистка прямого форматирования
        apply_heading_case(): Преобразование заголовков в верхний регистр
        detect_paragraph_style(): Определение стиля параграфа
        validate_document_structure(): Проверка структуры документа
        update_style_definitions(): Обновление стилей в документе
        format_document(): Основная функция форматирования
        validate_document(): Валидация документа
        save_config(): Сохранение конфигурации
    """
    
    # =============================================================
    # МЕТОД ИНИЦИАЛИЗАЦИИ КЛАССА
    # =============================================================
    
    
    def __init__(self, config_path=None):
        """
        Инициализация объекта GOSTFormatter с загрузкой конфигурации
        
        Args:
            config_path (str, optional): Путь к JSON файлу конфигурации.
                                         Если None, используется конфигурация по умолчанию.
        
        Returns:
            None: Метод не возвращает значение, инициализирует объект
        
        Example:
            >>> formatter = GOSTFormatter('config.json')
            >>> formatter = GOSTFormatter()  # С конфигурацией по умолчанию
        """
        # Проверка: указан ли путь к конфигурации и существует ли файл
        if config_path and os.path.exists(config_path):
            # Загрузка конфигурации из указанного файла
            self.config = self._load_config(config_path)
        else:
            # Использование конфигурации по умолчанию если файл не найден
            self.config = self._get_default_config()
        
        # Запись в лог информации о загруженной конфигурации
        logger.info(f"Конфигурация загружена: {self.config.get('name', 'Без имени')}")
    
    
    # =============================================================
    # МЕТОД ПОЛУЧЕНИЯ КОНФИГУРАЦИИ ПО УМОЛЧАНИЮ
    # =============================================================
    
    
    def _get_default_config(self):
        """
        Возвращает конфигурацию по умолчанию если файл не найден
        
        Returns:
            dict: Словарь с базовой конфигурацией стилей ГОСТ
        
        Note:
            Используется как fallback при отсутствии файла конфигурации
        """
        # Возврат словаря с минимальной конфигурацией
        return {
            # Версия конфигурации
            "version": "4.0",
            # Название конфигурации
            "name": "ГОСТ по умолчанию",
            # Стили по умолчанию (пустой словарь, будет заполнен)
            "styles": {},
            # Правила обнаружения стилей (пустой словарь)
            "detection_rules": {}
        }
    
    
    # =============================================================
    # МЕТОД ЗАГРУЗКИ JSON КОНФИГУРАЦИИ ИЗ ФАЙЛА
    # =============================================================
    
    
    def _load_config(self, config_path):
        """
        Загружает JSON конфигурацию из указанного файла
        
        Args:
            config_path (str): Полный путь к файлу конфигурации JSON
        
        Returns:
            dict: Словарь с загруженной конфигурацией
        
        Raises:
            FileNotFoundError: Если файл конфигурации не найден
            json.JSONDecodeError: Если файл содержит невалидный JSON
        
        Note:
            При ошибке загрузки возвращается конфигурация по умолчанию
        """
        # Попытка загрузки конфигурации из файла
        try:
            # Открытие файла конфигурации на чтение с кодировкой UTF-8
            with open(config_path, 'r', encoding='utf-8') as f:
                # Загрузка JSON данных из файла в словарь
                return json.load(f)
        # Обработка ошибки: файл конфигурации не найден
        except FileNotFoundError:
            # Запись ошибки в лог с указанием пути к файлу
            logger.error(f"Файл конфигурации не найден: {config_path}")
            # Возврат конфигурации по умолчанию
            return {"styles": {}}
        # Обработка ошибки: невалидный JSON в файле
        except json.JSONDecodeError as e:
            # Запись ошибки в лог с деталями проблемы
            logger.error(f"Ошибка JSON в конфигурации: {e}")
            # Возврат конфигурации по умолчанию
            return {"styles": {}}
    
    
    # =============================================================
    # МЕТОД ПРЕОБРАЗОВАНИЯ СТРОКИ ВЫРАВНИВАНИЯ В КОНСТАНТУ DOCX
    # =============================================================
    
    
    def _get_alignment(self, align_str):
        """
        Преобразует строковое значение выравнивания в константу docx
        
        Args:
            align_str (str): Строка выравнивания ('left', 'center', 'right', 'justify')
        
        Returns:
            WD_ALIGN_PARAGRAPH: Константа выравнивания из python-docx
        
        Example:
            >>> formatter._get_alignment('left')
            <WD_ALIGN_PARAGRAPH.LEFT: 1>
        """
        # Получение константы из словаря маппинга с fallback на justify
        return ALIGNMENT_MAP.get(align_str.lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)
    
    
    # =============================================================
    # МЕТОД НАСТРОЙКИ ШРИФТА ДЛЯ ПОДДЕРЖКИ КИРИЛЛИЦЫ
    # =============================================================
    
    
    def set_font_for_cyrillic(self, font, font_name):
        """
        Настраивает шрифт с полной поддержкой кириллицы через XML атрибуты
        
        Args:
            font: Объект шрифта из python-docx
            font_name (str): Название шрифта для установки
        
        Returns:
            None: Метод устанавливает атрибуты напрямую
        
        Note:
            Критически важно для корректного отображения кириллицы в Word
            Устанавливает шрифт для всех языковых диапазонов XML
        """
        # Установка основного имени шрифта
        font.name = font_name
        
        # Попытка настройки XML атрибутов для полной совместимости
        try:
            # Получение элемента свойств шрифта (rPr)
            rPr = font._element.rPr
            # Получение элемента настроек шрифтов (rFonts)
            rFonts = rPr.rFonts
            # Установка шрифта для восточноазиатских символов (включая кириллицу)
            rFonts.set(qn('w:eastAsia'), font_name)
            # Установка шрифта для латинских символов (ASCII)
            rFonts.set(qn('w:ascii'), font_name)
            # Установка шрифта для высокого ANSI диапазона
            rFonts.set(qn('w:hAnsi'), font_name)
            # Установка шрифта для сложных скриптов (кириллица, арабский)
            rFonts.set(qn('w:cs'), font_name)
        # Обработка любой ошибки при настройке XML
        except Exception as e:
            # Запись предупреждения в лог при неудаче
            logger.warning(f"Не удалось настроить XML шрифта: {e}")
            # Повторная установка имени шрифта как fallback
            font.name = font_name
    
    
    # =============================================================
    # МЕТОД ПРИМЕНЕНИЯ НАСТРОЕК СТИЛЯ К ОБЪЕКТУ СТИЛЯ
    # =============================================================
    
    
    def apply_style_properties(self, style, properties):
        """
        Применяет настройки шрифта и абзаца к объекту стиля
        
        Args:
            style: Объект стиля из python-docx
            properties (dict): Словарь с настройками стиля из конфигурации
        
        Returns:
            None: Метод модифицирует объект стиля напрямую
        
        Note:
            Пропускает отключенные стили (enabled: false)
        """
        # Проверка: включен ли стиль в конфигурации
        if not properties.get("enabled", True):
            # Запись в лог о пропуске отключенного стиля
            logger.debug(f"Стиль пропущен (отключён): {style.name}")
            # Выход из метода без применения настроек
            return
        
        # Получение настроек шрифта из свойств стиля
        font_props = properties.get("font", {})
        # Проверка: есть ли настройки шрифта
        if font_props:
            # Получение объекта шрифта из стиля
            font = style.font
            # Проверка: указано ли имя шрифта
            if "name" in font_props:
                # Применение шрифта с поддержкой кириллицы
                self.set_font_for_cyrillic(font, font_props["name"])
            # Проверка: указан ли размер шрифта
            if "size" in font_props:
                # Установка размера шрифта в пунктах
                font.size = Pt(font_props["size"])
            # Проверка: указано ли жирное начертание
            if "bold" in font_props:
                # Установка жирного начертания (True/False)
                font.bold = font_props["bold"]
            # Проверка: указано ли курсивное начертание
            if "italic" in font_props:
                # Установка курсивного начертания (True/False)
                font.italic = font_props["italic"]
        
        # Получение настроек абзаца из свойств стиля
        para_props = properties.get("paragraph", {})
        # Проверка: есть ли настройки абзаца
        if para_props:
            # Получение объекта формата абзаца из стиля
            pf = style.paragraph_format
            # Проверка: указан ли отступ первой строки
            if "first_line_indent_cm" in para_props:
                # Установка отступа первой строки в сантиметрах
                pf.first_line_indent = Cm(para_props["first_line_indent_cm"])
            # Проверка: указан ли левый отступ
            if "left_indent_cm" in para_props:
                # Установка левого отступа в сантиметрах
                pf.left_indent = Cm(para_props["left_indent_cm"])
            # Проверка: указан ли межстрочный интервал
            if "line_spacing" in para_props:
                # Установка межстрочного интервала
                pf.line_spacing = para_props["line_spacing"]
            # Проверка: указан ли интервал перед абзацем
            if "space_before_pt" in para_props:
                # Установка интервала перед абзацем в пунктах
                pf.space_before = Pt(para_props["space_before_pt"])
            # Проверка: указан ли интервал после абзаца
            if "space_after_pt" in para_props:
                # Установка интервала после абзаца в пунктах
                pf.space_after = Pt(para_props["space_after_pt"])
            # Проверка: указано ли выравнивание
            if "alignment" in para_props:
                # Преобразование строки выравнивания в константу и установка
                pf.alignment = self._get_alignment(para_props["alignment"])
    
    
    # =============================================================
    # МЕТОД ОЧИСТКИ ПРЯМОГО ФОРМАТИРОВАНИЯ ПАРАГРАФА
    # =============================================================
    
    
    def clear_direct_formatting(self, paragraph, is_table_cell=False):
        """
        Безопасно очищает прямое форматирование параграфа
        
        Args:
            paragraph: Объект параграфа из python-docx
            is_table_cell (bool): Флаг, находится ли параграф в ячейке таблицы
        
        Returns:
            None: Метод модифицирует параграф напрямую
        
        Note:
            Не сбрасывает имя и размер шрифта (они берутся из стиля)
            Для ячеек таблиц очистка более аккуратная
        """
        # Перебор всех runs (отрезков текста) в параграфе
        for run in paragraph.runs:
            # Сброс жирного начертания (будет взято из стиля)
            run.font.bold = None
            # Сброс курсивного начертания (будет взято из стиля)
            run.font.italic = None
            # Сброс подчёркивания (будет взято из стиля)
            run.font.underline = None
            # Сброс зачёркивания (будет взято из стиля)
            run.font.strike = None
            # Сброс цвета выделения текста
            run.font.highlight_color = None
            # Сброс цвета текста RGB
            run.font.color.rgb = None
            # Примечание: name и size НЕ сбрасываем - они будут взяты из стиля!
            
            # Проверка: находится ли параграф в ячейке таблицы
            if is_table_cell:
                # Для таблиц сохраняем больше форматирования
                # Восстановление жирного начертания
                run.font.bold = run.font.bold
                # Восстановление курсивного начертания
                run.font.italic = run.font.italic
    
    
    # =============================================================
    # МЕТОД ПРЕОБРАЗОВАНИЯ ЗАГОЛОВКОВ В ВЕРХНИЙ РЕГИСТР
    # =============================================================
    
    
    def apply_heading_case(self, paragraph, style_name):
        """
        Преобразует текст заголовков в ВЕРХНИЙ РЕГИСТР программно
        
        Args:
            paragraph: Объект параграфа из python-docx
            style_name (str): Название применяемого стиля
        
        Returns:
            None: Метод модифицирует текст параграфа напрямую
        
        Note:
            Используется вместо all_caps=True который ломает кириллицу
            Применяется только к Heading 1 и Heading 2
        """
        # Проверка: является ли стиль заголовком раздела или главы
        if style_name in ("Heading 1", "Heading 2"):
            # Преобразование всего текста параграфа в верхний регистр
            paragraph.text = paragraph.text.upper()
    
    
    # =============================================================
    # МЕТОД ОПРЕДЕЛЕНИЯ СТИЛЯ ПАРАГРАФА ПО СОДЕРЖИМОМУ
    # =============================================================
    
    
    def detect_paragraph_style(self, text, doc):
        """
        Определяет стиль параграфа на основе его текстового содержимого
        
        Args:
            text (str): Текст параграфа для анализа
            doc: Объект документа DOCX для проверки доступных стилей
        
        Returns:
            str: Название стиля для применения к параграфу
        
        Note:
            Проверки идут от более специфичных к более общим
            Возвращает "Обычный" или "Normal" если стиль не найден
        """
        # Удаление ведущих и замыкающих пробелов из текста
        text_stripped = text.strip()
        
        # Проверка: пустой ли параграф (нет текста)
        if not text_stripped:
            # Возврат русского имени стиля если есть, иначе английского
            return "Обычный" if "Обычный" in doc.styles else "Normal"
        
        # Получение правил обнаружения стилей из конфигурации
        rules = self.config.get("detection_rules", {})
        
        # Проверка: начинается ли текст с "РАЗДЕЛ" (заголовок уровня 1)
        if text_stripped.startswith("РАЗДЕЛ"):
            # Возврат Heading 1 если стиль есть в документе, иначе Normal
            return "Heading 1" if "Heading 1" in doc.styles else "Normal"
        
        # Проверка: начинается ли текст с "ГЛАВА" (заголовок уровня 2)
        if text_stripped.startswith("ГЛАВА"):
            # Возврат Heading 2 если стиль есть в документе, иначе Normal
            return "Heading 2" if "Heading 2" in doc.styles else "Normal"
        
        # Проверка: соответствует ли текст паттерну пункта (5.1, 5.2)
        if re.match(r'^\d+\.\d+\s', text_stripped):
            # Возврат Heading 3 если стиль есть в документе, иначе Normal
            return "Heading 3" if "Heading 3" in doc.styles else "Normal"
        
        # Проверка: начинается ли текст с ключевых слов подписей
        if text_stripped.startswith(("Таблица", "Рисунок", "[!Рисунок", "Примечание")):
            # Возврат Caption если стиль есть в документе, иначе Normal
            return "Caption" if "Caption" in doc.styles else "Normal"
        
        # Проверка: является ли текст элементом списка (цифра с точкой или маркер)
        if re.match(r'^\d+\.\s', text_stripped) or text_stripped.startswith(("•", "–", "−")):
            # Возврат List Paragraph если стиль есть в документе, иначе Normal
            return "List Paragraph" if "List Paragraph" in doc.styles else "Normal"
        
        # Проверка: является ли текст заголовком подпункта (короткий, с заглавной)
        if (len(text_stripped) < 80 and 
            not text_stripped.endswith(('.', '!', '?')) and
            text_stripped[0].isupper() and 
            text_stripped.count(' ') < 8):
            # Возврат Heading 4 если стиль есть в документе, иначе Normal
            return "Heading 4" if "Heading 4" in doc.styles else "Normal"
        
        # Возврат стиля по умолчанию (основной текст)
        return "Обычный" if "Обычный" in doc.styles else "Normal"
    
    
    # =============================================================
    # МЕТОД ПРОВЕРКИ СТРУКТУРЫ ДОКУМЕНТА НА СООТВЕТСТВИЕ ГОСТ
    # =============================================================
    
    
    def validate_document_structure(self, doc):
        """
        Проверяет структуру документа на соответствие ожидаемой
        
        Args:
            doc: Объект документа DOCX для проверки
        
        Returns:
            dict: Словарь с результатами валидации структуры
        
        Note:
            Проверяет наличие необходимых стилей в документе
            Проверяет наличие таблиц, рисунков, заголовков
        """
        # Инициализация словаря результатов валидации
        result = {
            # Флаг: соответствует ли структура требованиям
            "valid": True,
            # Список предупреждений о проблемах
            "warnings": [],
            # Список найденных элементов
            "info": []
        }
        
        # Получение списка необходимых стилей из конфигурации
        required_styles = ["Normal", "Обычный", "Heading 1", "Heading 2", "Heading 3"]
        
        # Проверка: есть ли хотя бы один из основных стилей
        has_basic_style = any(style in doc.styles for style in ["Normal", "Обычный"])
        
        # Проверка: отсутствует ли базовый стиль
        if not has_basic_style:
            # Установка флага валидации в False
            result["valid"] = False
            # Добавление предупреждения о проблеме
            result["warnings"].append("❌ Отсутствует базовый стиль (Normal/Обычный)")
        
        # Подсчёт количества параграфов в документе
        paragraph_count = len(doc.paragraphs)
        
        # Проверка: слишком ли мало параграфов (возможно пустой файл)
        if paragraph_count < 5:
            # Добавление предупреждения о малом количестве параграфов
            result["warnings"].append(f"⚠️ Мало параграфов: {paragraph_count}")
        
        # Добавление информации о количестве параграфов
        result["info"].append(f"ℹ️ Найдено параграфов: {paragraph_count}")
        
        # Проверка: есть ли таблицы в документе
        if len(doc.tables) > 0:
            # Добавление информации о количестве таблиц
            result["info"].append(f"ℹ️ Найдено таблиц: {len(doc.tables)}")
        
        # Проверка: есть ли изображения в документе
        if len(doc.inline_shapes) > 0:
            # Добавление информации о количестве изображений
            result["info"].append(f"ℹ️ Найдено изображений: {len(doc.inline_shapes)}")
        
        # Проверка: защищён ли документ от редактирования
        if hasattr(doc.settings, 'document_protection') and doc.settings.document_protection:
            # Установка флага валидации в False
            result["valid"] = False
            # Добавление предупреждения о защите документа
            result["warnings"].append("❌ Файл защищён от редактирования")
        
        # Возврат результатов валидации
        return result
    
    
    # =============================================================
    # МЕТОД ОБНОВЛЕНИЯ ОПРЕДЕЛЕНИЙ СТИЛЕЙ В ДОКУМЕНТЕ
    # =============================================================
    
    
        def update_style_definitions(self, doc):
            """Обновляет или создаёт стили в документе"""
        # ✅ Получаем стили из конфига с проверкой типа
        styles_config = self.config.get("styles", {})
        
        # ✅ Валидация: styles_config должен быть словарём
        if not isinstance(styles_config, dict):
            logger.error(f"❌ Ошибка: styles_config должен быть dict, получено: {type(styles_config)}")
            styles_config = {}  # Используем пустой словарь как fallback
        
        updated_count = 0
        created_count = 0
        
        for style_name, properties in styles_config.items():
            # ✅ Проверка: properties должен быть словарём
            if not isinstance(properties, dict):
                logger.warning(f"⚠️ Пропущен стиль '{style_name}': properties не dict (тип: {type(properties)})")
                continue
            
            # Пропускаем отключенные стили
            if not properties.get("enabled", True):
                continue
            
            try:
                style = doc.styles[style_name]
                self.apply_style_properties(style, properties)
                updated_count += 1
                logger.debug(f"Стиль обновлён: {style_name}")
            except KeyError:
                try:
                    style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                    self.apply_style_properties(style, properties)
                    created_count += 1
                    logger.debug(f"Стиль создан: {style_name}")
                except Exception as e:
                    logger.warning(f"Не удалось создать стиль {style_name}: {e}")
        
        logger.info(f"Стили обновлены: {updated_count}, создано: {created_count}")
        return updated_count, created_count
    
    
    # =============================================================
    # ОСНОВНОЙ МЕТОД ФОРМАТИРОВАНИЯ ДОКУМЕНТА
    # =============================================================
    
    
    def format_document(self, input_path, output_path, create_backup=True, progress_callback=None):
        """
        Основная функция форматирования документа по ГОСТ
        
        Args:
            input_path (str): Путь к входному DOCX файлу
            output_path (str): Путь для сохранения выходного файла
            create_backup (bool): Флаг создания резервной копии
            progress_callback (callable, optional): Функция обратного вызова для прогресса
        
        Returns:
            dict: Словарь со статистикой выполнения
        
        Raises:
            FileNotFoundError: Если входной файл не найден
            PermissionError: Если нет доступа к файлу
            ValueError: Если файл имеет неверное расширение
        
        Note:
            Создаёт бэкап перед обработкой если create_backup=True
            Восстанавливает из бэкапа при критической ошибке
        """
        # Инициализация словаря статистики выполнения
        stats = {
            # Количество обработанных параграфов
            "paragraphs_processed": 0,
            # Словарь: стиль -> количество применений
            "styles_applied": {},
            # Количество обработанных таблиц
            "tables_processed": 0,
            # Список ошибок
            "errors": [],
            # Флаг успешного выполнения
            "success": False
        }
        
        # Попытка выполнения форматирования
        try:
            # Проверка: существует ли входной файл
            if not os.path.exists(input_path):
                # Вызов ошибки если файл не найден
                raise FileNotFoundError(f"Файл не найден: {input_path}")
            
            # Проверка: имеет ли файл расширение .docx
            if not input_path.lower().endswith('.docx'):
                # Вызов ошибки если расширение неверное
                raise ValueError("Входной файл должен иметь расширение .docx")
            
            # Проверка: запрошено ли создание бэкапа
            if create_backup:
                # Формирование пути к файлу бэкапа
                backup_path = input_path + '.backup'
                # Копирование входного файла в бэкап
                shutil.copy2(input_path, backup_path)
                # Запись в лог о создании бэкапа
                logger.info(f"Бэкап создан: {backup_path}")
            
            # Получение директории для выходного файла
            output_dir = os.path.dirname(output_path) or '.'
            
            # Проверка: существует ли директория вывода
            if output_dir and output_dir != '.' and not os.path.exists(output_dir):
                # Создание директории если не существует
                os.makedirs(output_dir)
                # Запись в лог о создании директории
                logger.info(f"Директория создана: {output_dir}")
            
            # Проверка: есть ли права на чтение входного файла
            if not os.access(input_path, os.R_OK):
                # Вызов ошибки если нет прав на чтение
                raise PermissionError(f"Нет доступа на чтение: {input_path}")
            
            # Запись в лог об открытии файла
            logger.info(f"Открытие файла: {input_path}")
            
            # Открытие документа DOCX для обработки
            doc = Document(input_path)
            
            # =========================================================
            # ШАГ 1: ОБНОВЛЕНИЕ ОПРЕДЕЛЕНИЙ СТИЛЕЙ В ДОКУМЕНТЕ
            # =========================================================
            
            # Проверка: есть ли функция обратного вызова прогресса
            if progress_callback:
                # Вызов функции прогресса (10% выполнено)
                progress_callback(10, "Обновление определений стилей...")
            
            # Запись в лог о начале обновления стилей
            logger.info("Обновление определений стилей...")
            
            # Вызов метода обновления стилей в документе
            self.update_style_definitions(doc)
            
            # =========================================================
            # ШАГ 2: ПРИМЕНЕНИЕ СТИЛЕЙ К КАЖДОМУ ПАРАГРАФУ
            # =========================================================
            
            # Проверка: есть ли функция обратного вызова прогресса
            if progress_callback:
                # Вызов функции прогресса (30% выполнено)
                progress_callback(30, "Применение стилей к параграфам...")
            
            # Запись в лог о начале применения стилей
            logger.info("Применение стилей к параграфам...")
            
            # Получение общего количества параграфов в документе
            total_paragraphs = len(doc.paragraphs)
            
            # Перебор всех параграфов документа с индексом
            for i, paragraph in enumerate(doc.paragraphs):
                # Получение текста текущего параграфа
                text = paragraph.text
                
                # Проверка: находится ли параграф в таблице (через XML)
                is_in_table = paragraph._element.xpath('./ancestor::w:tbl')
                
                # Очистка прямого форматирования перед применением стиля
                self.clear_direct_formatting(paragraph, is_table_cell=is_in_table)
                
                # Определение подходящего стиля для параграфа
                style_name = self.detect_paragraph_style(text, doc)
                
                # Проверка: существует ли стиль в документе
                if style_name in doc.styles:
                    # Применение стиля к параграфу
                    paragraph.style = style_name
                    # Преобразование заголовков в верхний регистр
                    self.apply_heading_case(paragraph, style_name)
                    # Обновление статистики применённых стилей
                    stats["styles_applied"][style_name] = stats["styles_applied"].get(style_name, 0) + 1
                
                # Увеличение счётчика обработанных параграфов
                stats["paragraphs_processed"] += 1
                
                # Проверка: нужно ли обновлять прогресс (каждые 10 параграфов)
                if progress_callback and (i + 1) % 10 == 0:
                    # Расчёт процента выполнения (30-90%)
                    progress = 30 + int(((i + 1) / total_paragraphs) * 60)
                    # Вызов функции прогресса с сообщением
                    progress_callback(progress, f"Обработано {i + 1}/{total_paragraphs} параграфов")
            
            # =========================================================
            # ШАГ 3: ОБРАБОТКА ТАБЛИЦ ОТДЕЛЬНО
            # =========================================================
            
            # Проверка: есть ли функция обратного вызова прогресса
            if progress_callback:
                # Вызов функции прогресса (обработка таблиц)
                progress_callback(90, "Обработка таблиц...")
            
            # Запись в лог о начале обработки таблиц
            logger.info("Обработка таблиц...")
            
            # Перебор всех таблиц в документе
            for table in doc.tables:
                # Увеличение счётчика обработанных таблиц
                stats["tables_processed"] += 1
                
                # Перебор всех строк в таблице
                for row in table.rows:
                    # Перебор всех ячеек в строке
                    for cell in row.cells:
                        # Перебор всех параграфов в ячейке
                        for paragraph in cell.paragraphs:
                            # Проверка: есть ли стиль Table Grid в документе
                            if "Table Grid" in doc.styles:
                                # Применение стиля таблицы к параграфу
                                paragraph.style = "Table Grid"
                            # Очистка форматирования для ячеек таблицы
                            self.clear_direct_formatting(paragraph, is_table_cell=True)
            
            # =========================================================
            # ШАГ 4: СОХРАНЕНИЕ ОБРАБОТАННОГО ФАЙЛА
            # =========================================================
            
            # Проверка: есть ли функция обратного вызова прогресса
            if progress_callback:
                # Вызов функции прогресса (95% выполнено)
                progress_callback(95, "Сохранение файла...")
            
            # Запись в лог о сохранении файла
            logger.info(f"Сохранение в файл: {output_path}")
            
            # Сохранение обработанного документа в выходной файл
            doc.save(output_path)
            
            # Проверка: есть ли функция обратного вызова прогресса
            if progress_callback:
                # Вызов функции прогресса (100% выполнено)
                progress_callback(100, "Готово!")
            
            # Установка флага успешного выполнения
            stats["success"] = True
            # Сохранение абсолютного пути к выходному файлу
            stats["output_path"] = os.path.abspath(output_path)
            
            # Запись в лог об успешном завершении
            logger.info("✅ Успешно выполнено!")
            # Запись в лог пути к выходному файлу
            logger.info(f"📁 Выходной файл: {stats['output_path']}")
            # Запись в лог количества обработанных параграфов
            logger.info(f"📊 Обработано параграфов: {stats['paragraphs_processed']}")
            
        # Обработка ошибки: файл не найден
        except FileNotFoundError as e:
            # Запись ошибки в лог
            logger.error(f"❌ Файл не найден: {e}")
            # Добавление ошибки в статистику
            stats["errors"].append(str(e))
            # Повторное выбрасывание ошибки
            raise
        
        # Обработка ошибки: нет доступа к файлу
        except PermissionError as e:
            # Запись ошибки в лог
            logger.error(f"❌ Ошибка доступа: {e}")
            # Добавление ошибки в статистику
            stats["errors"].append(str(e))
            # Повторное выбрасывание ошибки
            raise
        
        # Обработка любой другой ошибки
        except Exception as e:
            # Запись критической ошибки в лог с трассировкой
            logger.error(f"❌ Критическая ошибка: {e}", exc_info=True)
            # Добавление ошибки в статистику
            stats["errors"].append(str(e))
            
            # Проверка: запрошено ли создание бэкапа и существует ли он
            if create_backup and os.path.exists(input_path + '.backup'):
                # Восстановление исходного файла из бэкапа
                shutil.copy2(input_path + '.backup', input_path)
                # Запись в лог о восстановлении файла
                logger.info("🔄 Файл восстановлен из бэкапа")
            
            # Повторное выбрасывание ошибки
            raise
        
        # Возврат статистики выполнения (в блоке finally гарантированно)
        return stats
    
    
    # =============================================================
    # МЕТОД ВАЛИДАЦИИ ДОКУМЕНТА ПЕРЕД ФОРМАТИРОВАНИЕМ
    # =============================================================
    
    
    def validate_document(self, input_path):
        """
        Проверяет документ перед форматированием на соответствие структуре
        
        Args:
            input_path (str): Путь к входному DOCX файлу
        
        Returns:
            dict: Словарь с результатами валидации
        
        Note:
            Проверяет доступность файла, наличие стилей, защиту
        """
        # Инициализация списка предупреждений
        warnings = []
        # Инициализация списка информации
        info = []
        
        # Попытка проверки документа
        try:
            # Открытие документа для проверки
            doc = Document(input_path)
            
            # Оценка количества страниц (параграфы / 30)
            page_estimate = len(doc.paragraphs) // 30
            
            # Проверка: большой ли документ (>100 страниц)
            if page_estimate > 100:
                # Добавление предупреждения о большом документе
                warnings.append(f"⚠️ Большой документ (~{page_estimate} стр.)")
            
            # Проверка: есть ли таблицы в документе
            if len(doc.tables) > 0:
                # Добавление информации о таблицах
                info.append(f"ℹ️ Найдено таблиц: {len(doc.tables)}")
            
            # Проверка: есть ли изображения в документе
            if len(doc.inline_shapes) > 0:
                # Добавление информации об изображениях
                info.append(f"ℹ️ Найдено изображений: {len(doc.inline_shapes)}")
            
            # Проверка: защищён ли документ от редактирования
            if hasattr(doc.settings, 'document_protection') and doc.settings.document_protection:
                # Добавление предупреждения о защите
                warnings.append("❌ Файл защищён от редактирования")
            
            # Проверка структуры документа на соответствие ГОСТ
            structure_result = self.validate_document_structure(doc)
            
            # Добавление предупреждений из проверки структуры
            warnings.extend(structure_result.get("warnings", []))
            
            # Добавление информации из проверки структуры
            info.extend(structure_result.get("info", []))
            
        # Обработка любой ошибки при проверке
        except Exception as e:
            # Добавление предупреждения об ошибке чтения
            warnings.append(f"❌ Ошибка чтения файла: {e}")
        
        # Возврат результатов валидации
        return {
            # Флаг валидности (нет критических предупреждений)
            "valid": len([w for w in warnings if "❌" in w]) == 0,
            # Список всех предупреждений
            "warnings": warnings,
            # Список всей информации
            "info": info
        }
    
    
    # =============================================================
    # МЕТОД СОХРАНЕНИЯ КОНФИГУРАЦИИ В JSON ФАЙЛ
    # =============================================================
    
    
    def save_config(self, config_path, config_data):
        """
        Сохраняет конфигурацию в JSON файл
        
        Args:
            config_path (str): Путь к файлу конфигурации
            config_data (dict): Данные конфигурации для сохранения
        
        Returns:
            None: Метод сохраняет файл напрямую
        
        Note:
            Используется мастером настройки для сохранения изменений
        """
        # Открытие файла конфигурации на запись с кодировкой UTF-8
        with open(config_path, 'w', encoding='utf-8') as f:
            # Запись данных конфигурации в формате JSON с отступами
            json.dump(config_data, f, ensure_ascii=False, indent=2)
        
        # Запись в лог о сохранении конфигурации
        logger.info(f"Конфигурация сохранена: {config_path}")


# =============================================================
# ФУНКЦИИ-ОБЁРТКИ ДЛЯ УДОБНОГО ИМПОРТА В GUI
# =============================================================


def format_file(input_path, output_path, config_path, create_backup=True, progress_callback=None):
    """
    Обёртка для вызова форматирования из GUI
    
    Args:
        input_path (str): Путь к входному файлу
        output_path (str): Путь к выходному файлу
        config_path (str): Путь к файлу конфигурации
        create_backup (bool): Флаг создания бэкапа
        progress_callback (callable): Функция обратного вызова прогресса
    
    Returns:
        dict: Статистика выполнения форматирования
    """
    # Создание объекта форматировщика с конфигурацией
    formatter = GOSTFormatter(config_path)
    # Вызов метода форматирования документа
    return formatter.format_document(input_path, output_path, create_backup, progress_callback)


def validate_file(input_path, config_path):
    """
    Обёртка для вызова валидации из GUI
    
    Args:
        input_path (str): Путь к входному файлу
        config_path (str): Путь к файлу конфигурации
    
    Returns:
        dict: Результаты валидации документа
    """
    # Создание объекта форматировщика с конфигурацией
    formatter = GOSTFormatter(config_path)
    # Вызов метода валидации документа
    return formatter.validate_document(input_path)


def save_config(config_path, config_data):
    """
    Обёртка для сохранения конфигурации
    
    Args:
        config_path (str): Путь к файлу конфигурации
        config_data (dict): Данные конфигурации для сохранения
    
    Returns:
        None: Функция сохраняет файл напрямую
    """
    # Создание объекта форматировщика с конфигурацией
    formatter = GOSTFormatter(config_path)
    # Вызов метода сохранения конфигурации
    return formatter.save_config(config_path, config_data)


# =============================================================
# ТОЧКА ВХОДА ПРИ ЗАПУСКЕ КАК СКРИПТ
# =============================================================


if __name__ == "__main__":
    # Вывод сообщения об успешной загрузке модуля
    print("✅ Модуль formatter_core загружен успешно")
    # Вывод информации о доступных классах
    print(f"   Классы: GOSTFormatter")
    # Вывод информации о доступных функциях
    print(f"   Функции: format_file, validate_file, save_config")
    # Вывод версии модуля
    print(f"   Версия: 4.0 (исправленная)")
    
    # Попытка тестового импорта функций
    try:
        # Импорт функций из текущего модуля
        from formatter_core import format_file, validate_file, save_config, GOSTFormatter
        # Вывод сообщения об успешном импорте
        print("\n✅ Все импорты успешны!")
    # Обработка ошибки импорта
    except ImportError as e:
        # Вывод сообщения об ошибке импорта
        print(f"\n❌ Ошибка импорта: {e}")

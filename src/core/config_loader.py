# config_loader.py

import json
import os
import yaml
from typing import Dict, Any, Optional

class ConfigLoader:
    """
    Загрузчик и валидатор конфигурации ГОСТ.
    Преобразует человеческие единицы (см, пт) в единицы Word (twips).
    1 см = 567 twips
    1 пт = 20 twips
    """
    
    CM_TO_TWIPS = 567
    PT_TO_TWIPS = 20

    def __init__(self, config_path: str):
        self.config_path = config_path
        self.config: Dict[str, Any] = {}
        self.styles: Dict[str, Any] = {}
        self.detection_rules: Dict[str, Any] = {}
        self.page_setup: Dict[str, Any] = {}
        self.headers_footers: Dict[str, Any] = {}
        self.formatting_rules: Dict[str, Any] = {}
        self.validation: Dict[str, Any] = {}
        self.federal_regional_projects: Dict[str, Any] = {}

    def load(self) -> bool:
        """Загружает конфиг с поддержкой includes и проверяет наличие обязательных полей."""
        if not os.path.exists(self.config_path):
            from src.core.formatter_errors import raise_config_not_found
            raise_config_not_found(self.config_path)

        from src.core.yaml_includes import load_yaml_with_includes

        ext = os.path.splitext(self.config_path)[1].lower()
        try:
            if ext == '.json':
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            elif ext in ('.yaml', '.yml'):
                self.config = load_yaml_with_includes(self.config_path)
            else:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                try:
                    self.config = yaml.safe_load(content)
                except yaml.YAMLError:
                    self.config = json.loads(content)
        except (json.JSONDecodeError, yaml.YAMLError) as e:
            from src.core.formatter_errors import raise_config_parse_error
            raise_config_parse_error(self.config_path, str(e))

        # Валидация структуры
        required_keys = ['styles', 'detection_rules']
        for key in required_keys:
            if key not in self.config:
                from src.core.formatter_errors import raise_config_missing_key
                raise_config_missing_key(self.config_path, key)

        self.styles = self.config.get('styles', {})
        self.detection_rules = self.config.get('detection_rules', {})
        self.page_setup = self.config.get('page_setup', {})
        self.headers_footers = self.config.get('headers_footers', {})
        self.formatting_rules = self.config.get('formatting_rules', {})
        self.validation = self.config.get('validation', {})
        self.federal_regional_projects = self.config.get('federal_regional_projects', {})

        # Предварительная обработка стилей (конвертация единиц)
        self._process_styles()
        return True

    def _process_styles(self):
        """Конвертирует значения отступов и размеров в twips для быстрого сравнения."""
        for style_name, style_props in self.styles.items():
            if not style_props.get('enabled', True):
                continue

            # Обработка шрифта
            if 'font' in style_props:
                font = style_props['font']
                if 'size' in font:
                    # Храним и в пт (для отчета), и в твипсах (для сравнения)
                    font['size_twips'] = int(font['size'] * self.PT_TO_TWIPS)

            # Обработка абзаца
            if 'paragraph' in style_props:
                para = style_props['paragraph']
                
                if 'first_line_indent_cm' in para:
                    val = para['first_line_indent_cm']
                    para['first_line_indent_twips'] = int(val * self.CM_TO_TWIPS)
                
                if 'left_indent_cm' in para:
                    val = para['left_indent_cm']
                    para['left_indent_twips'] = int(val * self.CM_TO_TWIPS)
                    
                if 'space_before_pt' in para:
                    val = para['space_before_pt']
                    para['space_before_twips'] = int(val * self.PT_TO_TWIPS)
                    
                if 'space_after_pt' in para:
                    val = para['space_after_pt']
                    para['space_after_twips'] = int(val * self.PT_TO_TWIPS)

    def get_style_config(self, style_name: str) -> Optional[Dict]:
        """Получает настройки конкретного стиля."""
        return self.styles.get(style_name)

    def get_detection_rules(self) -> Dict:
        """Получает правила обнаружения стилей."""
        return self.detection_rules

    def get_page_setup(self) -> Dict:
        """Получает настройки страницы."""
        return self.page_setup

    def get_headers_footers_config(self) -> Dict:
        """Получает настройки колонтитулов."""
        return self.headers_footers

    def get_formatting_rules(self) -> Dict:
        """Получает правила форматирования."""
        return self.formatting_rules

    def get_validation_config(self) -> Dict:
        """Получает настройки валидации."""
        return self.validation

    def get_federal_regional_projects_config(self) -> Dict:
        """Получает настройки для проектов федерального/регионального значения."""
        return self.federal_regional_projects

    def get_config(self) -> Dict:
        """Возвращает полный конфиг."""
        return self.config

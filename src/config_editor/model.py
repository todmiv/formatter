"""Модель данных конфигурации: FieldDescriptor, StyleDescriptor, ConfigModel."""

import copy
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional

import yaml


class FieldType(Enum):
    STRING = "string"
    NUMBER = "number"
    BOOLEAN = "boolean"
    ENUM = "enum"
    COLOR = "color"
    ALIGNMENT = "alignment"
    TEXT_TRANSFORM = "text_transform"
    LINE_SPACING = "line_spacing"
    DICTIONARY = "dictionary"
    LIST = "list"


ALIGNMENT_VALUES = ["left", "center", "right", "justify"]
TEXT_TRANSFORM_VALUES = ["capitalize", "uppercase", "lowercase", "none"]


@dataclass
class FieldDescriptor:
    key: str
    label: str
    field_type: FieldType
    value: Any
    default: Any = None
    enum_values: Optional[List[str]] = None
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    unit: Optional[str] = None
    description: str = ""
    required: bool = True
    children: Optional[Dict[str, "FieldDescriptor"]] = None


@dataclass
class StyleDescriptor:
    name: str
    enabled: bool
    font: Dict[str, FieldDescriptor] = field(default_factory=dict)
    paragraph: Dict[str, FieldDescriptor] = field(default_factory=dict)
    table: Optional[Dict[str, FieldDescriptor]] = None
    position: Optional[Dict[str, FieldDescriptor]] = None
    content: Optional[Dict[str, FieldDescriptor]] = None


def _make_field(key: str, label: str, value: Any, **kwargs) -> FieldDescriptor:
    if "field_type" in kwargs:
        ft = kwargs.pop("field_type")
    elif isinstance(value, bool):
        ft = FieldType.BOOLEAN
    elif isinstance(value, (int, float)):
        ft = FieldType.NUMBER
    elif isinstance(value, dict):
        ft = FieldType.DICTIONARY
    elif isinstance(value, list):
        ft = FieldType.LIST
    else:
        ft = FieldType.STRING
    return FieldDescriptor(key=key, label=label, field_type=ft, value=value, **kwargs)


def _build_font_fields(font_data: Dict) -> Dict[str, FieldDescriptor]:
    if not font_data:
        return {}
    return {
        "name": FieldDescriptor(
            key="name", label="Имя шрифта", field_type=FieldType.STRING,
            value=font_data.get("name", "Times New Roman"),
        ),
        "size": FieldDescriptor(
            key="size", label="Размер (пт)", field_type=FieldType.NUMBER,
            value=font_data.get("size", 12), min_value=1, max_value=100, unit="пт",
        ),
        "bold": FieldDescriptor(
            key="bold", label="Жирный", field_type=FieldType.BOOLEAN,
            value=font_data.get("bold", False),
        ),
        "italic": FieldDescriptor(
            key="italic", label="Курсив", field_type=FieldType.BOOLEAN,
            value=font_data.get("italic", False),
        ),
        "color": FieldDescriptor(
            key="color", label="Цвет", field_type=FieldType.COLOR,
            value=font_data.get("color", "black"),
        ),
    }


def _build_paragraph_fields(para_data: Dict) -> Dict[str, FieldDescriptor]:
    if not para_data:
        return {}
    return {
        "first_line_indent_cm": FieldDescriptor(
            key="first_line_indent_cm", label="Отступ 1-й строки (см)",
            field_type=FieldType.NUMBER, value=para_data.get("first_line_indent_cm", 0),
            min_value=0, max_value=10, unit="см",
        ),
        "line_spacing": FieldDescriptor(
            key="line_spacing", label="Межстрочный интервал",
            field_type=FieldType.NUMBER, value=para_data.get("line_spacing", 1.0),
            min_value=0.5, max_value=5.0,
        ),
        "space_before_pt": FieldDescriptor(
            key="space_before_pt", label="Интервал перед (пт)",
            field_type=FieldType.NUMBER, value=para_data.get("space_before_pt", 0),
            min_value=0, max_value=100, unit="пт",
        ),
        "space_after_pt": FieldDescriptor(
            key="space_after_pt", label="Интервал после (пт)",
            field_type=FieldType.NUMBER, value=para_data.get("space_after_pt", 0),
            min_value=0, max_value=100, unit="пт",
        ),
        "alignment": FieldDescriptor(
            key="alignment", label="Выравнивание",
            field_type=FieldType.ENUM, value=para_data.get("alignment", "left"),
            enum_values=ALIGNMENT_VALUES,
        ),
    }


def _build_table_fields(table_data: Optional[Dict]) -> Optional[Dict[str, FieldDescriptor]]:
    if not table_data:
        return None
    fields = {}
    field_defs = [
        ("auto_fit", "Автоподбор", FieldType.STRING, None),
        ("fixed_column_width", "Фикс. ширина", FieldType.BOOLEAN, None),
        ("row_height_cm", "Высота строки (см)", FieldType.NUMBER, "см"),
        ("row_height_mode", "Режим высоты", FieldType.STRING, None),
        ("border_style", "Стиль границ", FieldType.STRING, None),
        ("border_width_pt", "Ширина границ (пт)", FieldType.NUMBER, "пт"),
    ]
    for key, label, ft, unit in field_defs:
        if key in table_data:
            fields[key] = FieldDescriptor(
                key=key, label=label, field_type=ft,
                value=table_data[key], unit=unit,
            )
    return fields if fields else None


class ConfigModel:
    """Центральная модель конфигурации."""

    def __init__(self, config_path: str):
        self.config_path = config_path
        self.raw_config: Dict[str, Any] = {}
        self.styles: Dict[str, StyleDescriptor] = {}
        self.sections: Dict[str, Any] = {}
        self._modified = False

    @property
    def modified(self) -> bool:
        return self._modified

    def load(self) -> bool:
        with open(self.config_path, "r", encoding="utf-8") as f:
            self.raw_config = yaml.safe_load(f) or {}
        self._parse_styles()
        self._parse_sections()
        self._modified = False
        return True

    def _parse_styles(self):
        self.styles.clear()
        for name, data in self.raw_config.get("styles", {}).items():
            if not isinstance(data, dict):
                continue
            font = _build_font_fields(data.get("font", {}))
            para = _build_paragraph_fields(data.get("paragraph", {}))
            table = _build_table_fields(data.get("table"))
            self.styles[name] = StyleDescriptor(
                name=name,
                enabled=data.get("enabled", True),
                font=font,
                paragraph=para,
                table=table,
            )

    def _parse_sections(self):
        for key in ("detection_rules", "page_setup", "headers_footers",
                     "formatting_rules", "validation", "federal_regional_projects"):
            self.sections[key] = self.raw_config.get(key, {})

    def get_style(self, name: str) -> Optional[StyleDescriptor]:
        return self.styles.get(name)

    def set_style_enabled(self, name: str, enabled: bool):
        if name in self.styles:
            self.styles[name].enabled = enabled
            self.raw_config["styles"][name]["enabled"] = enabled
            self._modified = True

    def update_style_field(self, style_name: str, section: str, field_key: str, value: Any):
        if style_name not in self.styles:
            return
        style = self.styles[style_name]
        if section == "font" and field_key in style.font:
            style.font[field_key].value = value
            self.raw_config["styles"][style_name].setdefault("font", {})[field_key] = value
            self._modified = True
        elif section == "paragraph" and field_key in style.paragraph:
            style.paragraph[field_key].value = value
            self.raw_config["styles"][style_name].setdefault("paragraph", {})[field_key] = value
            self._modified = True
        elif section == "table" and style.table and field_key in style.table:
            style.table[field_key].value = value
            self.raw_config["styles"][style_name].setdefault("table", {})[field_key] = value
            self._modified = True

    def add_style(self, name: str):
        default_style = {
            "enabled": True,
            "font": {"name": "Times New Roman", "size": 12, "bold": False, "italic": False},
            "paragraph": {"first_line_indent_cm": 0, "line_spacing": 1.0, "alignment": "left"},
        }
        self.raw_config.setdefault("styles", {})[name] = default_style
        self._parse_styles()
        self._modified = True

    def delete_style(self, name: str):
        if name in self.raw_config.get("styles", {}):
            del self.raw_config["styles"][name]
            self._parse_styles()
            self._modified = True

    def update_section(self, section: str, data: Any):
        self.raw_config[section] = data
        self.sections[section] = data
        self._modified = True

    def to_dict(self) -> Dict[str, Any]:
        return copy.deepcopy(self.raw_config)

    def save(self, path: str = None) -> bool:
        target = path or self.config_path
        with open(target, "w", encoding="utf-8") as f:
            yaml.dump(self.raw_config, f, default_flow_style=False,
                      allow_unicode=True, sort_keys=False)
        self._modified = False
        return True

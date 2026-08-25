"""Валидаторы значений конфигурации."""

import re
from dataclasses import dataclass
from typing import List, Literal

from .model import FieldType


@dataclass
class ValidationError:
    field_path: str
    message: str
    severity: Literal["error", "warning"] = "error"


class Validators:
    @staticmethod
    def validate_string(value, field_desc) -> List[ValidationError]:
        errors = []
        if field_desc.required and (value is None or str(value).strip() == ""):
            errors.append(ValidationError(field_desc.key, "Поле обязательно для заполнения"))
        return errors

    @staticmethod
    def validate_number(value, field_desc) -> List[ValidationError]:
        errors = []
        try:
            num = float(value)
        except (TypeError, ValueError):
            errors.append(ValidationError(field_desc.key, f"Значение '{value}' не является числом"))
            return errors
        if field_desc.min_value is not None and num < field_desc.min_value:
            errors.append(ValidationError(
                field_desc.key,
                f"Значение {num} меньше допустимого минимума {field_desc.min_value}",
            ))
        if field_desc.max_value is not None and num > field_desc.max_value:
            errors.append(ValidationError(
                field_desc.key,
                f"Значение {num} больше допустимого максимума {field_desc.max_value}",
            ))
        return errors

    @staticmethod
    def validate_enum(value, field_desc) -> List[ValidationError]:
        errors = []
        if field_desc.enum_values and value not in field_desc.enum_values:
            errors.append(ValidationError(
                field_desc.key,
                f"Значение '{value}' не входит в список допустимых: {field_desc.enum_values}",
            ))
        return errors

    @staticmethod
    def validate_color(value) -> List[ValidationError]:
        errors = []
        if value is None:
            return errors
        s = str(value).strip()
        if s.startswith("#"):
            if not re.match(r"^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$", s):
                errors.append(ValidationError("color", f"Некорректный HEX-цвет: {s}"))
        elif not re.match(r"^[a-zA-Z]+$", s):
            errors.append(ValidationError("color", f"Некорректный цвет: {s}"))
        return errors

    @staticmethod
    def validate_all(model) -> List[ValidationError]:
        errors = []
        for name, style in model.styles.items():
            prefix = f"styles.{name}"
            for section_name in ("font", "paragraph", "table"):
                section = getattr(style, section_name, None)
                if not section:
                    continue
                for key, fd in section.items():
                    path = f"{prefix}.{section_name}.{key}"
                    fd.key = path
                    if fd.field_type == FieldType.STRING:
                        errors.extend(Validators.validate_string(fd.value, fd))
                    elif fd.field_type == FieldType.NUMBER:
                        errors.extend(Validators.validate_number(fd.value, fd))
                    elif fd.field_type == FieldType.ENUM:
                        errors.extend(Validators.validate_enum(fd.value, fd))
                    elif fd.field_type == FieldType.COLOR:
                        errors.extend(Validators.validate_color(fd.value))
                    fd.key = key
        return errors

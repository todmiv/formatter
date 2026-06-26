#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FormatterError — типизированные ошибки с кодами выхода и контекстными подсказками.
Аналог PandocError в Pandoc: каждый тип ошибки имеет уникальный exit code
и справочное сообщение для пользователя.

Использование:
    from src.core.formatter_errors import FormatterError, ErrorKind

    raise FormatterError(ErrorKind.CONFIG_NOT_FOUND, path="config.yaml")

    # Или с контекстом
    try:
        formatter.format(...)
    except FormatterError as e:
        print(e)  # сообщение с подсказкой
        sys.exit(e.exit_code)
"""

import sys
from enum import IntEnum, Enum
from typing import Optional, Any


class ErrorKind(IntEnum):
    """
    Коды ошибок Formatter.
    Уникальные для каждого типа, как exit codes в Pandoc.
    """
    OK = 0

    FILE_NOT_FOUND = 1
    FILE_READ_ERROR = 2
    FILE_WRITE_ERROR = 3

    CONFIG_NOT_FOUND = 10
    CONFIG_PARSE_ERROR = 11
    CONFIG_MISSING_KEY = 12
    CONFIG_INVALID_VALUE = 13

    TEMPLATE_NOT_FOUND = 20
    TEMPLATE_EXTRACT_ERROR = 21
    TEMPLATE_APPLY_ERROR = 22

    DOCX_OPEN_ERROR = 30
    DOCX_INVALID = 31
    DOCX_SECTION_ERROR = 32

    STYLE_NOT_FOUND = 40
    STYLE_CREATE_ERROR = 41
    STYLE_APPLY_ERROR = 42

    PARAGRAPH_INDEX_ERROR = 50
    TABLE_INDEX_ERROR = 51
    HEADER_FOOTER_ERROR = 52
    LOCATION_INVALID = 53

    AUDIT_ERROR = 60
    APPLY_ERROR = 61

    CONTENT_VALIDATION_ERROR = 70
    CONTENT_TRANSFORM_ERROR = 71

    BATCH_PARTIAL_FAILURE = 80

    INTERNAL_ERROR = 99


class FormatterError(Exception):
    """
    Типизированная ошибка Formatter с кодом выхода и контекстной подсказкой.

    :param kind: тип ошибки (ErrorKind)
    :param message: описание ошибки (опционально, перезаписывает стандартное)
    :param **kwargs: контекстные параметры для подсказки
    """

    HINTS = {
        ErrorKind.FILE_NOT_FOUND: "Убедитесь, что путь указан верно и файл существует. Используйте --verbose для диагностики.",
        ErrorKind.FILE_READ_ERROR: "Проверьте права доступа к файлу и не повреждён ли он.",
        ErrorKind.FILE_WRITE_ERROR: "Проверьте права доступа к директории назначения. Директория может не существовать.",

        ErrorKind.CONFIG_NOT_FOUND: "Укажите путь к конфигу через -c/--config или используйте configs/active/config.yaml.",
        ErrorKind.CONFIG_PARSE_ERROR: "Проверьте синтаксис YAML/JSON. Используйте YAML-валидатор для диагностики.",
        ErrorKind.CONFIG_MISSING_KEY: "Добавьте недостающий раздел в конфиг. Пример: см. configs/active/config.yaml.",
        ErrorKind.CONFIG_INVALID_VALUE: "Проверьте тип значения в конфиге. Ожидается число, строка или булево значение.",

        ErrorKind.TEMPLATE_NOT_FOUND: "Укажите путь к DOCX-шаблону через -t/--template. См. docs/user_guides/USER_GUIDE.md.",
        ErrorKind.TEMPLATE_EXTRACT_ERROR: "Убедитесь, что DOCX-шаблон не повреждён и содержит стили.",
        ErrorKind.TEMPLATE_APPLY_ERROR: "Шаблон может содержать несовместимые стили. Проверьте целевой документ.",

        ErrorKind.DOCX_OPEN_ERROR: "Файл может быть повреждён, зашифрован или открыт в другом приложении.",
        ErrorKind.DOCX_INVALID: "Файл не является валидным DOCX. Проверьте расширение и формат.",
        ErrorKind.DOCX_SECTION_ERROR: "Документ не содержит секций. Возможно, повреждена структура XML.",

        ErrorKind.STYLE_NOT_FOUND: "Стиль отсутствует в документе. Используйте initialize_styles для его создания.",
        ErrorKind.STYLE_CREATE_ERROR: "Не удалось создать стиль. Проверьте, не является ли имя зарезервированным.",
        ErrorKind.STYLE_APPLY_ERROR: "Не удалось применить стиль к элементу. Возможно, стиль несовместим с типом элемента.",

        ErrorKind.PARAGRAPH_INDEX_ERROR: "Индекс параграфа вне диапазона. Проверьте структуру документа через --verbose.",
        ErrorKind.TABLE_INDEX_ERROR: "Индекс таблицы вне диапазона. Проверьте количество таблиц в документе.",
        ErrorKind.HEADER_FOOTER_ERROR: "Ошибка доступа к колонтитулу. Убедитесь, что секция существует.",
        ErrorKind.LOCATION_INVALID: "Формат location не распознан. Допустимые ключи: index, paragraph_index, section, container.",

        ErrorKind.AUDIT_ERROR: "Ошибка при сканировании документа. Проверьте логи для деталей.",
        ErrorKind.APPLY_ERROR: "Ошибка при применении исправлений. Возможно, документ повреждён.",

        ErrorKind.CONTENT_VALIDATION_ERROR: "Найдены нарушения правил форматирования. Используйте detail() для просмотра.",
        ErrorKind.CONTENT_TRANSFORM_ERROR: "Ошибка при трансформации текста. Проверьте правила в конфиге.",

        ErrorKind.BATCH_PARTIAL_FAILURE: "Не все документы обработаны успешно. Проверьте логи для деталей.",

        ErrorKind.INTERNAL_ERROR: "Внутренняя ошибка. Создайте issue: https://github.com/your-repo/issues",
    }

    def __init__(
        self,
        kind: ErrorKind,
        message: Optional[str] = None,
        **context,
    ):
        self.kind = kind
        self.exit_code = int(kind)
        self.context = context
        self._message = message
        super().__init__(self._build_message())

    def _build_message(self) -> str:
        if self._message:
            msg = self._message
        else:
            msg = self.kind.name.replace('_', ' ').title()

        parts = [msg]

        if self.context:
            ctx_parts = []
            for k, v in self.context.items():
                if k == 'hint':
                    continue
                ctx_parts.append(f"{k}={v!r}")
            if ctx_parts:
                parts.append(f"({', '.join(ctx_parts)})")

        hint = self.context.get('hint') or self.HINTS.get(self.kind)
        if hint:
            parts.append(f"\n  → {hint}")

        return ' '.join(parts[:1]) + ''.join(parts[1:])

    def __str__(self) -> str:
        return self._build_message()

    def __repr__(self) -> str:
        return f"FormatterError({self.kind.name}, context={self.context})"


def raise_file_not_found(path: str) -> None:
    raise FormatterError(ErrorKind.FILE_NOT_FOUND, path=path)


def raise_config_not_found(path: str) -> None:
    raise FormatterError(ErrorKind.CONFIG_NOT_FOUND, path=path)


def raise_config_parse_error(path: str, detail: str) -> None:
    raise FormatterError(ErrorKind.CONFIG_PARSE_ERROR, path=path, detail=detail)


def raise_config_missing_key(path: str, key: str) -> None:
    raise FormatterError(ErrorKind.CONFIG_MISSING_KEY, path=path, key=key)


def raise_template_not_found(path: str) -> None:
    raise FormatterError(ErrorKind.TEMPLATE_NOT_FOUND, path=path)


def raise_docx_open_error(path: str, detail: str) -> None:
    raise FormatterError(ErrorKind.DOCX_OPEN_ERROR, path=path, detail=detail)


def raise_docx_invalid(path: str) -> None:
    raise FormatterError(ErrorKind.DOCX_INVALID, path=path)


def raise_style_not_found(style_name: str, doc_path: str) -> None:
    raise FormatterError(ErrorKind.STYLE_NOT_FOUND, style=style_name, doc=doc_path)


def raise_paragraph_index_error(index: int, total: int) -> None:
    raise FormatterError(
        ErrorKind.PARAGRAPH_INDEX_ERROR,
        index=index,
        total=total,
    )


def raise_table_index_error(index: int, total: int) -> None:
    raise FormatterError(
        ErrorKind.TABLE_INDEX_ERROR,
        index=index,
        total=total,
    )


def raise_location_invalid(location: dict) -> None:
    raise FormatterError(ErrorKind.LOCATION_INVALID, location=location)


def raise_header_footer_error(section: int, detail: str) -> None:
    raise FormatterError(
        ErrorKind.HEADER_FOOTER_ERROR,
        section=section,
        detail=detail,
    )


def exit_with_error(error: FormatterError) -> None:
    """Выводит ошибку и завершает процесс с соответствующим exit code."""
    print(f"Ошибка [{error.kind.name}]: {error}", file=sys.stderr)
    sys.exit(error.exit_code)

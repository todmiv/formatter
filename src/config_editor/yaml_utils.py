"""Утилиты для загрузки/сохранения YAML с сохранением комментариев."""

import re
from typing import Dict, List, Tuple

import yaml


def _extract_inline_comment(line: str) -> Tuple[str, str]:
    """Извлекает inline-комментарий из строки YAML."""
    stripped = line.rstrip()
    if not stripped or stripped.lstrip().startswith("#"):
        return line, ""
    in_string = False
    quote_char = None
    for i, ch in enumerate(stripped):
        if in_string:
            if ch == quote_char and (i == 0 or stripped[i - 1] != "\\"):
                in_string = False
        else:
            if ch in ('"', "'"):
                in_string = True
                quote_char = ch
            elif ch == "#":
                return stripped[:i].rstrip(), stripped[i:].rstrip()
    return stripped, ""


class YamlUtils:
    @staticmethod
    def load_with_comments(path: str) -> Tuple[Dict, Dict[str, str]]:
        """Загружает YAML и извлекает комментарии по ключевым путям."""
        comments: Dict[str, str] = {}
        with open(path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        full_comment_lines = []
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("#"):
                full_comment_lines.append(stripped)

        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}

        comment_text = "\n".join(full_comment_lines)
        if comment_text:
            comments["__header__"] = comment_text

        return data, comments

    @staticmethod
    def save_with_comments(path: str, data: Dict, comments: Dict[str, str] = None) -> bool:
        """Сохраняет YAML. Если есть комментарии — вставляет header."""
        with open(path, "w", encoding="utf-8") as f:
            header = (comments or {}).get("__header__", "")
            if header:
                for line in header.split("\n"):
                    f.write(line + "\n")
                f.write("\n")
            yaml.dump(data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        return True

    @staticmethod
    def extract_comments(yaml_text: str) -> Dict[str, str]:
        """Извлекает комментарии из текста YAML."""
        result = {}
        for line in yaml_text.split("\n"):
            stripped = line.strip()
            if stripped.startswith("#"):
                result[f"line_{len(result)}"] = stripped
        return result

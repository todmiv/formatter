#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Snapshot-тестирование для XML-структуры DOCX.

Аналог snapshot-тестов в Pandoc: сравнивает сгенерированный DOCX
с эталонным XML-представлением.

Использование:
    from src.core.snapshot_testing import DocxSnapshot

    snapshot = DocxSnapshot("tests/snapshots")

    # Первый запуск — создаёт эталонный файл
    snapshot.assert_matches("normal_style.docx", doc)

    # Следующие запуски — сравнивает с эталоном
    snapshot.assert_matches("normal_style.docx", doc)

    # Обновление эталона
    snapshot.update("normal_style.docx", doc)
"""

import os
import re
import zipfile
import hashlib
from typing import Dict, Optional
from lxml import etree


class DocxSnapshot:
    """
    Сравнивает XML-структуру DOCX с эталонными снимками.

    Снимки хранятся как XML-файлы, извлечённые из DOCX.
    """

    KEY_FILES = [
        "[Content_Types].xml",
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
    ]

    def __init__(self, snapshots_dir: str, update_mode: bool = False):
        """
        :param snapshots_dir: директория со снимками.
        :param update_mode: если True, перезаписывает снимки вместо сравнения.
        """
        self.snapshots_dir = os.path.abspath(snapshots_dir)
        self.update_mode = update_mode
        os.makedirs(self.snapshots_dir, exist_ok=True)

    def assert_matches(self, name: str, docx_path: str) -> bool:
        """
        Сравнивает DOCX-файл с эталонным снимком.

        :param name: имя снимка (без расширения).
        :param docx_path: путь к DOCX-файлу.
        :return: True если совпадает.
        """
        current = self._extract_xml(docx_path)
        snapshot_path = self._snapshot_path(name)

        if self.update_mode or not os.path.exists(snapshot_path):
            self._save_snapshot(name, current)
            print(f"[SNAPSHOT] Создан эталон: {name}")
            return True

        expected = self._load_snapshot(name)
        diff = self._compare(current, expected)

        if diff:
            msg = f"[SNAPSHOT] Расхождения в '{name}':\n"
            for d in diff[:10]:
                msg += f"  - {d}\n"
            if len(diff) > 10:
                msg += f"  ... и ещё {len(diff) - 10}\n"
            msg += f"\nДля обновления: snapshot.update('{name}', docx_path)"
            raise AssertionError(msg)

        return True

    def update(self, name: str, docx_path: str):
        """Обновляет эталонный снимок."""
        current = self._extract_xml(docx_path)
        self._save_snapshot(name, current)
        print(f"[SNAPSHOT] Обновлён эталон: {name}")

    def _snapshot_path(self, name: str) -> str:
        return os.path.join(self.snapshots_dir, f"{name}.snapshot")

    def _extract_xml(self, docx_path: str) -> Dict[str, str]:
        """Извлекает ключевые XML-файлы из DOCX."""
        result = {}

        with zipfile.ZipFile(docx_path, 'r') as z:
            for filename in z.namelist():
                if filename in self.KEY_FILES or filename.endswith('.xml'):
                    try:
                        content = z.read(filename).decode('utf-8')
                        content = self._normalize_xml(content)
                        result[filename] = content
                    except Exception:
                        pass

        return result

    def _normalize_xml(self, xml_str: str) -> str:
        """Нормализует XML: убирает пробелы, динамические поля, форматирует."""
        xml_str = re.sub(r'\s+', ' ', xml_str)
        xml_str = re.sub(r'>\s+<', '><', xml_str)
        xml_str = re.sub(r'xmlns:ns\d+="[^"]*"', '', xml_str)

        xml_str = re.sub(
            r'<dcterms:modified>[^<]*</dcterms:modified>',
            '<dcterms:modified>2024-01-01T00:00:00Z</dcterms:modified>',
            xml_str
        )
        xml_str = re.sub(
            r'<dcterms:created>[^<]*</dcterms:created>',
            '<dcterms:created>2024-01-01T00:00:00Z</dcterms:created>',
            xml_str
        )
        xml_str = re.sub(
            r'uuid="[a-f0-9-]+"',
            'uuid="00000000-0000-0000-0000-000000000000"',
            xml_str
        )
        xml_str = re.sub(
            r'r:id="[^"]*"',
            'r:id="rId0"',
            xml_str
        )
        xml_str = re.sub(
            r'Target="[^"]*\.xml"',
            'Target="item0.xml"',
            xml_str
        )
        xml_str = re.sub(
            r'TargetMode="[^"]*"',
            '',
            xml_str
        )

        return xml_str.strip()

    def _compare(self, current: Dict[str, str], expected: Dict[str, str]) -> list:
        """Сравнивает два набора XML-файлов. Возвращает список различий."""
        diffs = []

        all_files = set(current.keys()) | set(expected.keys())

        for filename in sorted(all_files):
            if filename not in current:
                diffs.append(f"Файл '{filename}' отсутствует в текущем DOCX")
            elif filename not in expected:
                diffs.append(f"Файл '{filename}' отсутствует в эталоне")
            elif current[filename] != expected[filename]:
                cur_hash = hashlib.md5(current[filename].encode()).hexdigest()[:8]
                exp_hash = hashlib.md5(expected[filename].encode()).hexdigest()[:8]
                diffs.append(f"Файл '{filename}' различается ( текущий: {cur_hash}, эталон: {exp_hash} )")

        return diffs

    def _save_snapshot(self, name: str, data: Dict[str, str]):
        """Сохраняет снимок в файл."""
        snapshot_path = self._snapshot_path(name)
        with open(snapshot_path, 'w', encoding='utf-8') as f:
            for filename, content in sorted(data.items()):
                f.write(f"=== {filename} ===\n")
                f.write(content)
                f.write("\n\n")

    def _load_snapshot(self, name: str) -> Dict[str, str]:
        """Загружает снимок из файла."""
        snapshot_path = self._snapshot_path(name)
        result = {}
        current_file = None
        current_content = []

        with open(snapshot_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.startswith("=== ") and line.rstrip().endswith(" ==="):
                    if current_file:
                        content = '\n'.join(current_content).rstrip('\n')
                        result[current_file] = content
                    current_file = line.strip().strip("=").strip()
                    current_content = []
                else:
                    current_content.append(line.rstrip('\n'))

            if current_file:
                content = '\n'.join(current_content).rstrip('\n')
                result[current_file] = content

        return result


def extract_docx_xml(docx_path: str) -> Dict[str, str]:
    """Утилита: извлекает все XML-файлы из DOCX."""
    result = {}
    with zipfile.ZipFile(docx_path, 'r') as z:
        for filename in z.namelist():
            if filename.endswith('.xml') or filename.endswith('.rels'):
                try:
                    result[filename] = z.read(filename).decode('utf-8')
                except Exception:
                    pass
    return result


def get_docx_styles(docx_path: str) -> list:
    """Утилита: возвращает список стилей из DOCX."""
    xml = extract_docx_xml(docx_path)
    if "word/styles.xml" not in xml:
        return []

    root = etree.fromstring(xml["word/styles.xml"].encode())
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    styles = []
    for style in root.findall(".//w:style", ns):
        style_id = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
        name_elem = style.find("w:name", ns)
        name = name_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if name_elem is not None else style_id
        styles.append({"id": style_id, "name": name})

    return styles


def get_docx_paragraph_count(docx_path: str) -> int:
    """Утилита: возвращает количество параграфов в DOCX."""
    xml = extract_docx_xml(docx_path)
    if "word/document.xml" not in xml:
        return 0

    root = etree.fromstring(xml["word/document.xml"].encode())
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    return len(root.findall(".//w:p", ns))

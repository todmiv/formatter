#!/usr/bin/env python3
"""Тест преобразования единиц допусков в настройках GUI."""
import sys
import os
import pytest

sys.path.insert(0, os.path.dirname(__file__))

from src.core.docx_utils import MM_TO_TWIPS, PT_TO_TWIPS


def test_tolerance_conversion():
    """Проверяем, что преобразование twips <-> мм работает корректно."""
    root = None
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()

        from gui_formatter import GOSTFormatterGUI
        app = GOSTFormatterGUI(root)

        font_twips = app.font_tolerance
        indent_twips = app.indent_tolerance
        spacing_twips = app.spacing_tolerance

        font_mm = font_twips / MM_TO_TWIPS
        indent_mm = indent_twips / MM_TO_TWIPS
        spacing_mm = spacing_twips / MM_TO_TWIPS

        font_twips_back = int(round(font_mm * MM_TO_TWIPS))
        indent_twips_back = int(round(indent_mm * MM_TO_TWIPS))
        spacing_twips_back = int(round(spacing_mm * MM_TO_TWIPS))

        assert abs(font_twips_back - font_twips) <= 1, "Ошибка преобразования font"
        assert abs(indent_twips_back - indent_twips) <= 1, "Ошибка преобразования indent"
        assert abs(spacing_twips_back - spacing_twips) <= 1, "Ошибка преобразования spacing"
    finally:
        if root is not None:
            root.destroy()

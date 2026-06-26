#!/usr/bin/env python3
"""Тест логики преобразования мм <-> pt <-> twips."""


def test_conversion():
    MM_TO_TWIPS = 56.7
    PT_TO_TWIPS = 20

    font_twips = 10
    indent_twips = 20
    spacing_twips = 20

    font_mm = font_twips / MM_TO_TWIPS
    font_pt = font_twips / PT_TO_TWIPS
    indent_mm = indent_twips / MM_TO_TWIPS
    indent_pt = indent_twips / PT_TO_TWIPS
    spacing_mm = spacing_twips / MM_TO_TWIPS
    spacing_pt = spacing_twips / PT_TO_TWIPS

    font_twips_back = int(round(font_mm * MM_TO_TWIPS))
    indent_twips_back = int(round(indent_mm * MM_TO_TWIPS))
    spacing_twips_back = int(round(spacing_mm * MM_TO_TWIPS))

    assert font_twips_back == font_twips, f"Ошибка font: {font_twips_back} != {font_twips}"
    assert indent_twips_back == indent_twips, f"Ошибка indent: {indent_twips_back} != {indent_twips}"
    assert spacing_twips_back == spacing_twips, f"Ошибка spacing: {spacing_twips_back} != {spacing_twips}"

    font_twips_from_pt = int(round(font_pt * PT_TO_TWIPS))
    indent_twips_from_pt = int(round(indent_pt * PT_TO_TWIPS))
    spacing_twips_from_pt = int(round(spacing_pt * PT_TO_TWIPS))

    assert font_twips_from_pt == font_twips, f"Ошибка font из pt: {font_twips_from_pt} != {font_twips}"
    assert indent_twips_from_pt == indent_twips, f"Ошибка indent из pt: {indent_twips_from_pt} != {indent_twips}"
    assert spacing_twips_from_pt == spacing_twips, f"Ошибка spacing из pt: {spacing_twips_from_pt} != {spacing_twips}"

#!/usr/bin/env python3
"""Анализ проблем конвертации"""

from docx import Document
from docx.shared import Pt, Cm

# Загружаем конвертированный файл
converted_path = r'd:\project\Demographics\Наборы данных\Миасский ГО\result\ОБРАЗЕЦ_Том II. Материалы по обоснованию_Миасский_converted_v2.docx'
template_path = r'd:\project\Demographics\Наборы данных\Миасский ГО\result\ОБРАЗЕЦ_Том II. Материалы по обоснованию_Миасский.docx'

print("=" * 70)
print("ПРОБЛЕМА 1: Выравнивание подписей и заголовков таблиц")
print("=" * 70)

doc = Document(converted_path)
print("\nКонвертированный файл - подписи таблиц:")
for i, p in enumerate(doc.paragraphs):
    if 'Таблица' in p.text and len(p.text) < 20:
        align = p.alignment
        print(f"  [{i:3d}] '{p.text}' alignment={align}")

t_doc = Document(template_path)
print("\nШаблон - подписи таблиц:")
for i, p in enumerate(t_doc.paragraphs):
    if 'Таблица' in p.text and len(p.text) < 20:
        align = p.alignment
        print(f"  [{i:3d}] '{p.text}' alignment={align}")

print("\n" + "=" * 70)
print("ПРОБЛЕМА 2: Пустые абзацы после таблиц")
print("=" * 70)

doc = Document(converted_path)
print("\nПустые абзацы:")
for i, p in enumerate(doc.paragraphs):
    if not p.text.strip() and i > 0:
        prev_text = doc.paragraphs[i-1].text[:40] if doc.paragraphs[i-1].text else "(пусто)"
        print(f"  [{i:3d}] после: '{prev_text}'")

print("\n" + "=" * 70)
print("ПРОБЛЕМА 3: Шрифт примечаний")
print("=" * 70)

doc = Document(converted_path)
print("\nПримечания (Примечание —):")
for i, p in enumerate(doc.paragraphs):
    if 'Примечание' in p.text and '—' in p.text:
        for run in p.runs:
            if run.font.size:
                print(f"  [{i:3d}] '{p.text[:50]}' size={run.font.size.pt}pt")
                break
        else:
            print(f"  [{i:3d}] '{p.text[:50]}' size=стиль")

print("\n" + "=" * 70)
print("ПРОБЛЕМА 4: Ширина колонок таблиц 3.3.7 и 3.3.10")
print("=" * 70)

doc = Document(converted_path)
# Найдём таблицы 3.3.7 и 3.3.10
for i, p in enumerate(doc.paragraphs):
    if 'Таблица 3.3.7' in p.text or 'Таблица 3.3.10' in p.text:
        # Найдём следующую таблицу
        for j in range(i+1, min(i+20, len(doc.paragraphs))):
            # Проверяем таблицы после этой подписи
            break
    
# Проверим все таблицы
print("\nШирины колонок всех таблиц:")
for t_idx, table in enumerate(doc.tables[:10]):
    if table.rows:
        widths = []
        for cell in table.rows[0].cells:
            width = cell.width
            if width:
                widths.append(f"{width/914400*2.54:.1f}см")
            else:
                widths.append("auto")
        print(f"  Таблица {t_idx+1}: {widths}")

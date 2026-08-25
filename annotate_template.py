"""
Копирует DOCX-шаблон и добавляет аннотации форматирования ко ВСЕМ элементам.
Включая абзацы без стиля (с прямым форматированием) и таблицы.
"""
import shutil, os, yaml, zipfile, xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

SRC = r'd:\project\Demographics\Наборы данных\Миасский ГО\result\ОБРАЗЕЦ_Том II. Материалы по обоснованию_Миасский.docx'
DST = os.path.join(os.path.dirname(__file__), 'dist', 'templates', 'reference_format.docx')
CONFIG = os.path.join(os.path.dirname(__file__), 'dist', 'configs', 'active', 'config_from_Миасс.yaml')

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ALIGN_MAP = {'both': 'по ширине', 'left': 'лево', 'center': 'центр', 'right': 'право', None: 'лево'}


def load_config(path):
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def get_style_map(docx_path):
    """Читает mapping styleId -> styleName из styles.xml."""
    smap = {}
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/styles.xml') as f:
            for s in ET.fromstring(f.read()).iter(NS + 'style'):
                sid = s.get(NS + 'styleId')
                nm = s.find(NS + 'name')
                if sid and nm is not None:
                    smap[sid] = nm.get(NS + 'val')
    return smap


def describe_paragraph(p_elem, style_name, config):
    """Генерирует описание форматирования абзаца по его XML."""
    parts = []

    # Стиль
    if style_name and style_name != '(none)':
        s_cfg = config.get('styles', {}).get(style_name, {})
        font = s_cfg.get('font', {})
        if font:
            fname = font.get('name', 'TNR')
            fsize = font.get('size', 12)
            fbold = ' bold' if font.get('bold') else ''
            parts.append(f'Стиль: {style_name} — {fname} {fsize}pt{fbold}')
        else:
            parts.append(f'Стиль: {style_name}')
    else:
        parts.append('Стиль: Normal (нет явного стиля)')

    # Прямое форматирование абзаца
    pPr = p_elem.find(NS + 'pPr')
    if pPr is not None:
        jc = pPr.find(NS + 'jc')
        if jc is not None:
            align = ALIGN_MAP.get(jc.get(NS + 'val'), jc.get(NS + 'val'))
            parts.append(f'  Выравнивание: {align}')

        ind = pPr.find(NS + 'ind')
        if ind is not None:
            fl = ind.get(NS + 'firstLine')
            left = ind.get(NS + 'left')
            if fl:
                parts.append(f'  Красная строка: {int(fl)/567:.2f} см')
            if left and int(left) > 0:
                parts.append(f'  Левый отступ: {int(left)/567:.2f} см')

        sp = pPr.find(NS + 'spacing')
        if sp is not None:
            before = sp.get(NS + 'before')
            after = sp.get(NS + 'after')
            line = sp.get(NS + 'line')
            if before and int(before) > 0:
                parts.append(f'  Интервал перед: {int(before)/20:.0f} пт')
            if after and int(after) > 0:
                parts.append(f'  Интервал после: {int(after)/20:.0f} пт')
            if line:
                parts.append(f'  Межстрочный: {int(line)/240:.2f}')

    # Прямое форматирование шрифта (из первого run)
    for r_elem in p_elem.iter(NS + 'r'):
        rPr = r_elem.find(NS + 'rPr')
        if rPr is not None:
            sz = rPr.find(NS + 'sz')
            b = rPr.find(NS + 'b')
            color = rPr.find(NS + 'color')
            rFonts = rPr.find(NS + 'rFonts')
            if rFonts is not None:
                fname = rFonts.get(NS + 'ascii') or rFonts.get(NS + 'hAnsi')
                if fname:
                    parts.append(f'  Прямой шрифт: {fname}')
            if sz is not None:
                parts.append(f'  Прямой размер: {int(sz.get(NS + "val"))//2}pt')
            if b is not None:
                parts.append('  Прямой bold')
            if color is not None:
                c = color.get(NS + 'val')
                if c and c.lower() != 'auto':
                    parts.append(f'  Прямой цвет: {c}')
        break  # только первый run

    return '\n'.join(parts)


def is_all_caps(text):
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return False
    return all(c.isupper() for c in letters)


def get_text_from_elem(p_elem):
    """Извлекает текст из XML-элемента абзаца."""
    parts = []
    for t in p_elem.iter(NS + 't'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts).strip()


def make_annotation(text, desc, is_caps=False):
    """Формирует строку аннотации."""
    if is_caps:
        desc += '\n  ★ Текст ЗАГЛАВНЫМИ буквами'
    return '── ФОРМАТИРОВАНИЕ ──\n' + desc


def annotate_table(doc, tbl_elem, config):
    """Аннотирует таблицу — добавляет описание перед ней."""
    desc = 'Таблица: Table Grid стиль\n  Сетка:所有 ячейки с рамками\n  Шрифт ячеек: TNR 10-12pt'

    new_p = doc.add_paragraph()
    run = new_p.add_run('── ФОРМАТИРОВАНИЕ ──\n' + desc)
    run.font.name = 'Consolas'
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x00, 0x80, 0x00)
    run.font.italic = True
    new_p.paragraph_format.space_before = Pt(1)
    new_p.paragraph_format.space_after = Pt(2)

    # Вставляем перед таблицей
    tbl_elem.addprevious(new_p._element)


def process():
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    shutil.copy2(SRC, DST)

    config = load_config(CONFIG)
    smap = get_style_map(SRC)
    doc = Document(DST)

    body = doc.element.body

    # Собираем все элементы тела (абзацы + таблицы) в порядке следования
    elements = []
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            elements.append(('para', child))
        elif tag == 'tbl':
            elements.append(('table', child))

    # Обходим в обратном порядке
    count = 0
    for kind, elem in reversed(elements):
        if kind == 'table':
            annotate_table(doc, elem, config)
            count += 1
            continue

        # Определяем стиль
        pPr = elem.find(NS + 'pPr')
        sid = None
        if pPr is not None:
            ps = pPr.find(NS + 'pStyle')
            if ps is not None:
                sid = ps.get(NS + 'val')
        style_name = smap.get(sid, '(none)') if sid else '(none)'

        text = get_text_from_elem(elem)
        if not text:
            continue

        caps = is_all_caps(text)

        desc = describe_paragraph(elem, style_name, config)
        annotation = make_annotation(text, desc, is_caps=caps)

        new_p = doc.add_paragraph()
        run = new_p.add_run(annotation)
        run.font.name = 'Consolas'
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x00, 0x80, 0x00)
        run.font.italic = True
        new_p.paragraph_format.space_before = Pt(1)
        new_p.paragraph_format.space_after = Pt(3)

        para = Paragraph(elem, doc)
        elem.addnext(new_p._element)
        count += 1

    doc.save(DST)
    print(f'Аннотировано элементов: {count}')
    print(f'Файл: {DST}')


if __name__ == '__main__':
    process()

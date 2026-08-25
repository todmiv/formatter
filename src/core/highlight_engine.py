"""
Модуль для визуальной подсветки проблемных элементов в документе.
Создает копию DOCX с цветовой заливкой (shading) параграфов, таблиц и других элементов
в соответствии с важностью проблемы (CRITICAL, WARNING, INFO).
"""
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from src.core.audit_engine import AuditIssue, Severity
import os
import tempfile
import logging

logger = logging.getLogger(__name__)

class HighlightEngine:
    """Движок подсветки проблемных мест в документе."""
    
    # Цвета заливки в формате RRGGBB (без префикса #)
    COLOR_CRITICAL = "FFCCCC"   # светло-красный
    COLOR_WARNING = "FFF3CD"    # светло-желтый
    COLOR_INFO = "D1ECF1"       # светло-голубой
    
    @classmethod
    def highlight_issues(cls, doc_path: str, issues: list[AuditIssue], output_path: str = None) -> str:
        """
        Создает копию документа с подсветкой проблемных элементов.
        
        :param doc_path: путь к исходному документу
        :param issues: список объектов AuditIssue
        :param output_path: путь для сохранения (если None, создается временный файл)
        :return: путь к созданному файлу
        """
        if not issues:
            logger.warning("Список проблем пуст, подсветка не требуется.")
            return doc_path
        
        try:
            doc = docx.Document(doc_path)
        except Exception as e:
            logger.error(f"Не удалось открыть документ {doc_path}: {e}")
            raise
        
        # Группируем проблемы по типу элемента и локации
        for issue in issues:
            cls._apply_highlight(doc, issue)
        
        if output_path is None:
            # Создаем временный файл
            fd, output_path = tempfile.mkstemp(suffix='_highlighted.docx', dir=os.path.dirname(doc_path))
            os.close(fd)
        
        try:
            doc.save(output_path)
            logger.info(f"Документ с подсветкой сохранен: {output_path}")
        except Exception as e:
            logger.error(f"Ошибка сохранения документа с подсветкой: {e}")
            raise
        
        return output_path
    
    @classmethod
    def _apply_highlight(cls, doc: docx.Document, issue: AuditIssue):
        """Применяет подсветку к конкретному элементу на основе его локации."""
        location = issue.location
        severity = issue.severity
        
        # Определяем цвет в зависимости от важности
        if severity == Severity.CRITICAL:
            fill_color = cls.COLOR_CRITICAL
        elif severity == Severity.WARNING:
            fill_color = cls.COLOR_WARNING
        else:
            fill_color = cls.COLOR_INFO
        
        # Обработка параграфов
        if 'index' in location:
            para_index = location['index']
            if 0 <= para_index < len(doc.paragraphs):
                para = doc.paragraphs[para_index]
                cls._add_shading_to_paragraph(para, fill_color)
                return
        
        # Обработка таблиц
        if 'table_index' in location:
            table_index = location['table_index']
            if 0 <= table_index < len(doc.tables):
                table = doc.tables[table_index]
                cls._add_shading_to_table(table, fill_color)
                return
        
        # Обработка колонтитулов (секций)
        if 'section' in location:
            # Пока пропускаем, сложная реализация
            pass
        
        logger.debug(f"Не удалось применить подсветку для проблемы {issue.id}, локация {location}")
    
    @staticmethod
    def _add_shading_to_paragraph(paragraph, fill_color: str):
        """Добавляет заливку (shading) к параграфу."""
        pPr = paragraph._element.get_or_add_pPr()
        # Удаляем существующий shd, если есть
        for elem in pPr.iterchildren():
            if elem.tag == qn('w:shd'):
                pPr.remove(elem)
                break
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), fill_color)
        shd.set(qn('w:val'), 'clear')
        pPr.append(shd)
    
    @staticmethod
    def _add_shading_to_table(table, fill_color: str):
        """Добавляет заливку ко всей таблице (к каждой ячейке)."""
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._element.get_or_add_tcPr()
                # Удаляем существующий shd
                for elem in tcPr.iterchildren():
                    if elem.tag == qn('w:shd'):
                        tcPr.remove(elem)
                        break
                shd = OxmlElement('w:shd')
                shd.set(qn('w:fill'), fill_color)
                shd.set(qn('w:val'), 'clear')
                tcPr.append(shd)

if __name__ == "__main__":
    # Пример использования
    import sys
    if len(sys.argv) < 2:
        print("Использование: python highlight_engine.py <docx_path>")
        sys.exit(1)
    # Заглушка: создаем тестовые проблемы
    from audit_engine import Severity
    test_issues = [
        AuditIssue(
            id="test_1",
            severity=Severity.CRITICAL,
            category="FONT",
            element_type="Normal",
            location={'index': 0},
            description="Test",
            current_value="Arial",
            expected_value="Times New Roman"
        )
    ]
    output = HighlightEngine.highlight_issues(sys.argv[1], test_issues)
    print(f"Создан файл с подсветкой: {output}")
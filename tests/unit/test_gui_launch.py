import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import tkinter as tk
from gui_formatter import GOSTFormatterGUI

def test_gui():
    root = tk.Tk()
    root.withdraw()  # скрыть окно
    app = GOSTFormatterGUI(root)
    print("GUI инициализирован успешно")
    # Проверяем атрибуты (doc_path теперь в контроллере)
    assert app.controller.doc_path is None
    # config_loader находится в контроллере
    assert app.controller.config_loader is not None
    print("Атрибуты в порядке")
    root.destroy()
    print("Тест пройден")

if __name__ == "__main__":
    test_gui()
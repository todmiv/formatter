"""Компонент предпросмотра стиля на Canvas."""

import tkinter as tk
from tkinter import font as tkfont
from typing import Optional


class StylePreview(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, bg="white", height=200)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self._current_style = None

    def update_preview(self, style_descriptor):
        self._current_style = style_descriptor
        self.canvas.delete("all")
        if not style_descriptor:
            return
        self._draw_font_sample(style_descriptor)
        self._draw_alignment_indicator(style_descriptor)
        self._draw_indent_ruler(style_descriptor)

    def _draw_font_sample(self, style):
        font_data = style.font
        font_name = font_data.get("name", {}).value if "name" in font_data else "Times New Roman"
        font_size = font_data.get("size", {}).value if "size" in font_data else 12
        font_bold = font_data.get("bold", {}).value if "bold" in font_data else False
        font_italic = font_data.get("italic", {}).value if "italic" in font_data else False
        font_color = font_data.get("color", {}).value if "color" in font_data else "black"

        weight = "bold" if font_bold else "normal"
        slant = "italic" if font_italic else "roman"
        try:
            f = tkfont.Font(family=font_name, size=int(font_size), weight=weight, slant=slant)
        except Exception:
            f = tkfont.Font(family="Times New Roman", size=12)

        self.canvas.create_text(
            20, 30, text="AaBbCc 123", anchor=tk.W, font=f, fill=font_color,
        )
        self.canvas.create_text(
            20, 60, text=f"{font_name}, {font_size}пт",
            anchor=tk.W, fill="gray",
        )

    def _draw_alignment_indicator(self, style):
        para = style.paragraph
        alignment = "left"
        if "alignment" in para:
            alignment = para["alignment"].value

        y_base = 100
        self.canvas.create_text(20, y_base - 10, text="Выравнивание:", anchor=tk.W, fill="gray")

        lines = [
            (100, 180), (120, 160), (90, 190), (150, 130),
        ]
        for i, (x_start, width) in enumerate(lines):
            y = y_base + 5 + i * 14
            if alignment == "left":
                x = 30
            elif alignment == "center":
                x = 30 + (180 - width) // 2
            elif alignment == "right":
                x = 30 + 180 - width
            else:
                x = 30
            self.canvas.create_rectangle(x, y, x + width, y + 8, fill="lightblue", outline="")

    def _draw_indent_ruler(self, style):
        para = style.paragraph
        indent_cm = 0
        if "first_line_indent_cm" in para:
            indent_cm = para["first_line_indent_cm"].value

        y = 170
        self.canvas.create_text(20, y, text="Красная строка:", anchor=tk.W, fill="gray")
        self.canvas.create_line(30, y + 15, 230, y + 15, fill="black")
        for i in range(11):
            x = 30 + i * 20
            self.canvas.create_line(x, y + 12, x, y + 18, fill="black")

        indent_px = int(indent_cm * 20)
        if indent_px > 0:
            self.canvas.create_rectangle(
                30, y + 20, 30 + indent_px, y + 28, fill="red", outline="",
            )
            self.canvas.create_text(
                30 + indent_px + 5, y + 24, text=f"{indent_cm} см", anchor=tk.W, fill="red",
            )

    def clear(self):
        self.canvas.delete("all")
        self._current_style = None

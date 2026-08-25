"""Типизированные виджеты для редактирования полей конфигурации."""

import tkinter as tk
from tkinter import ttk, colorchooser
from typing import Any, Callable, Optional

from .model import FieldDescriptor, FieldType


class FieldWidgetFactory:
    @staticmethod
    def create(parent, descriptor: FieldDescriptor, on_change: Callable = None) -> tk.Widget:
        ft = descriptor.field_type

        if ft == FieldType.BOOLEAN:
            return FieldWidgetFactory._create_boolean(parent, descriptor, on_change)
        elif ft == FieldType.NUMBER:
            return FieldWidgetFactory._create_number(parent, descriptor, on_change)
        elif ft == FieldType.ENUM:
            return FieldWidgetFactory._create_enum(parent, descriptor, on_change)
        elif ft == FieldType.COLOR:
            return FieldWidgetFactory._create_color(parent, descriptor, on_change)
        elif ft == FieldType.STRING:
            return FieldWidgetFactory._create_string(parent, descriptor, on_change)
        elif ft == FieldType.LIST:
            return FieldWidgetFactory._create_list(parent, descriptor, on_change)
        else:
            return FieldWidgetFactory._create_string(parent, descriptor, on_change)

    @staticmethod
    def _create_string(parent, desc: FieldDescriptor, on_change: Callable) -> ttk.Entry:
        var = tk.StringVar(value=str(desc.value or ""))
        entry = ttk.Entry(parent, textvariable=var, width=40)
        if on_change:
            var.trace_add("write", lambda *_: on_change(desc.key, var.get()))
        entry._field_var = var
        return entry

    @staticmethod
    def _create_number(parent, desc: FieldDescriptor, on_change: Callable) -> ttk.Entry:
        var = tk.StringVar(value=str(desc.value if desc.value is not None else ""))
        entry = ttk.Entry(parent, textvariable=var, width=20)
        entry._field_var = var
        entry._field_desc = desc
        if on_change:
            def _on_num_change(*_):
                raw = var.get().strip()
                try:
                    val = float(raw) if "." in raw else int(raw)
                    on_change(desc.key, val)
                except ValueError:
                    pass
            var.trace_add("write", _on_num_change)
        return entry

    @staticmethod
    def _create_boolean(parent, desc: FieldDescriptor, on_change: Callable) -> ttk.Checkbutton:
        var = tk.BooleanVar(value=bool(desc.value))
        cb = ttk.Checkbutton(parent, variable=var)
        if on_change:
            var.trace_add("write", lambda *_: on_change(desc.key, var.get()))
        cb._field_var = var
        return cb

    @staticmethod
    def _create_enum(parent, desc: FieldDescriptor, on_change: Callable) -> ttk.Combobox:
        values = desc.enum_values or []
        var = tk.StringVar(value=str(desc.value or ""))
        cb = ttk.Combobox(parent, textvariable=var, values=values, state="readonly", width=25)
        if on_change:
            var.trace_add("write", lambda *_: on_change(desc.key, var.get()))
        cb._field_var = var
        return cb

    @staticmethod
    def _create_color(parent, desc: FieldDescriptor, on_change: Callable) -> ttk.Frame:
        frame = ttk.Frame(parent)
        var = tk.StringVar(value=str(desc.value or "black"))
        entry = ttk.Entry(frame, textvariable=var, width=20)
        entry.pack(side=tk.LEFT)

        color_preview = tk.Label(frame, width=3, height=1, bg=str(desc.value or "black"),
                                 relief="sunken")
        color_preview.pack(side=tk.LEFT, padx=(5, 0))

        def _pick_color():
            color = colorchooser.askcolor(initialcolor=var.get())
            if color and color[1]:
                var.set(color[1])
                color_preview.config(bg=color[1])

        btn = ttk.Button(frame, text="...", command=_pick_color, width=3)
        btn.pack(side=tk.LEFT, padx=(2, 0))

        if on_change:
            var.trace_add("write", lambda *_: (
                on_change(desc.key, var.get()),
                color_preview.config(bg=var.get()) if var.get().startswith("#") else None,
            ))

        frame._field_var = var
        frame._color_preview = color_preview
        return frame

    @staticmethod
    def _create_list(parent, desc: FieldDescriptor, on_change: Callable) -> tk.Frame:
        frame = tk.Frame(parent)
        listbox = tk.Listbox(frame, height=5, width=35)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for item in (desc.value or []):
            listbox.insert(tk.END, str(item))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(2, 0))

        def _add():
            val = entry_var.get().strip()
            if val:
                listbox.insert(tk.END, val)
                entry_var.set("")
                if on_change:
                    on_change(desc.key, list(listbox.get(0, tk.END)))

        def _remove():
            sel = listbox.curselection()
            if sel:
                listbox.delete(sel[0])
                if on_change:
                    on_change(desc.key, list(listbox.get(0, tk.END)))

        entry_var = tk.StringVar()
        ttk.Entry(btn_frame, textvariable=entry_var, width=15).pack(pady=1)
        ttk.Button(btn_frame, text="+", command=_add, width=3).pack(pady=1)
        ttk.Button(btn_frame, text="-", command=_remove, width=3).pack(pady=1)

        frame._field_listbox = listbox
        return frame

    @staticmethod
    def get_value(widget: tk.Widget, descriptor: FieldDescriptor) -> Any:
        if hasattr(widget, "_field_var"):
            var = widget._field_var
            if isinstance(var, tk.BooleanVar):
                return var.get()
            raw = var.get().strip()
            if descriptor.field_type == FieldType.NUMBER:
                try:
                    return float(raw) if "." in raw else int(raw)
                except ValueError:
                    return raw
            return raw
        if hasattr(widget, "_field_listbox"):
            return list(widget._field_listbox.get(0, tk.END))
        return None

    @staticmethod
    def set_value(widget: tk.Widget, descriptor: FieldDescriptor, value: Any):
        if hasattr(widget, "_field_var"):
            var = widget._field_var
            if isinstance(var, tk.BooleanVar):
                var.set(bool(value))
            else:
                var.set(str(value) if value is not None else "")
            if hasattr(widget, "_color_preview") and str(value).startswith("#"):
                widget._color_preview.config(bg=str(value))

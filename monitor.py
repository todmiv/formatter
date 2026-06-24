"""
MonitorWindow — панель мониторинга в реальном времени.
Отображает статус обработки документов и потребление системных ресурсов.
"""

import tkinter as tk
from tkinter import ttk
import threading
import time
import os
try:
    import psutil
except ImportError:
    psutil = None
import logging
from collections import deque
from datetime import datetime

logger = logging.getLogger(__name__)

# Цвета
COLOR_BG = "#1e1e2e"
COLOR_FG = "#cdd6f4"
COLOR_GREEN = "#a6e3a1"
COLOR_RED = "#f38ba8"
COLOR_YELLOW = "#f9e2af"
COLOR_BLUE = "#89b4fa"
COLOR_SURFACE = "#313244"
COLOR_OVERLAY = "#45475a"


class CircularBuffer:
    """Кольцевой буфер для хранения истории значений."""

    def __init__(self, maxlen=60):
        self.data = deque(maxlen=maxlen)
        self.maxlen = maxlen

    def append(self, value):
        self.data.append(value)

    def get_list(self):
        return list(self.data)

    def last(self):
        return self.data[-1] if self.data else 0


class MonitorWindow:
    """Окно мониторинга в реальном времени."""

    def __init__(self, root, controller=None):
        self.root = root
        self.controller = controller
        self.running = False
        self.update_interval = 1000  # ms

        # Данные
        self.cpu_history = CircularBuffer(60)
        self.ram_history = CircularBuffer(60)
        self.process_cpu_history = CircularBuffer(60)
        self.process_ram_history = CircularBuffer(60)
        self.audit_issues_history = CircularBuffer(60)
        self.events = deque(maxlen=100)

        self.window = tk.Toplevel(root)
        self.window.title("Мониторинг системы")
        self.window.geometry("900x650")
        self.window.minsize(750, 500)
        self.window.configure(bg=COLOR_BG)

        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self.start()

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Monitor.TFrame", background=COLOR_BG)
        style.configure("Monitor.TLabel", background=COLOR_BG, foreground=COLOR_FG, font=("Consolas", 10))
        style.configure("MonitorTitle.TLabel", background=COLOR_BG, foreground=COLOR_FG, font=("Consolas", 12, "bold"))
        style.configure("Card.TFrame", background=COLOR_SURFACE)
        style.configure("Card.TLabel", background=COLOR_SURFACE, foreground=COLOR_FG, font=("Consolas", 10))
        style.configure("CardTitle.TLabel", background=COLOR_SURFACE, foreground=COLOR_BLUE, font=("Consolas", 10, "bold"))
        style.configure("Value.TLabel", background=COLOR_SURFACE, foreground=COLOR_GREEN, font=("Consolas", 18, "bold"))
        style.configure("Unit.TLabel", background=COLOR_SURFACE, foreground=COLOR_FG, font=("Consolas", 9))
        style.configure("Green.TLabel", background=COLOR_SURFACE, foreground=COLOR_GREEN, font=("Consolas", 10))
        style.configure("Red.TLabel", background=COLOR_SURFACE, foreground=COLOR_RED, font=("Consolas", 10))
        style.configure("Yellow.TLabel", background=COLOR_SURFACE, foreground=COLOR_YELLOW, font=("Consolas", 10))

        main = ttk.Frame(self.window, style="Monitor.TFrame")
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Верхняя панель: карточки метрик ---
        cards_frame = ttk.Frame(main, style="Monitor.TFrame")
        cards_frame.pack(fill=tk.X, pady=(0, 10))

        self.card_cpu = self._create_card(cards_frame, "CPU", "0%", COLOR_GREEN)
        self.card_cpu.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.card_ram = self._create_card(cards_frame, "RAM", "0 MB", COLOR_BLUE)
        self.card_ram.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.card_proc_cpu = self._create_card(cards_frame, "APP CPU", "0%", COLOR_YELLOW)
        self.card_proc_cpu.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.card_proc_ram = self._create_card(cards_frame, "APP RAM", "0 MB", COLOR_GREEN)
        self.card_proc_ram.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # --- Графики ---
        charts_frame = ttk.Frame(main, style="Monitor.TFrame")
        charts_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.cpu_canvas = self._create_chart(charts_frame, "CPU Usage (%)")
        self.cpu_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.ram_canvas = self._create_chart(charts_frame, "RAM Usage (MB)")
        self.ram_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # --- Нижняя панель: статус + лог ---
        bottom_frame = ttk.Frame(main, style="Monitor.TFrame")
        bottom_frame.pack(fill=tk.BOTH, expand=True)

        # Статус обработки
        status_frame = ttk.LabelFrame(bottom_frame, text=" Статус обработки ", style="Card.TFrame")
        status_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.lbl_status = ttk.Label(status_frame, text="Ожидание...", style="Card.TLabel")
        self.lbl_status.pack(anchor=tk.W, padx=5, pady=2)

        self.lbl_issues = ttk.Label(status_frame, text="Проблем: 0", style="Card.TLabel")
        self.lbl_issues.pack(anchor=tk.W, padx=5, pady=2)

        self.lbl_applied = ttk.Label(status_frame, text="Исправлено: 0", style="Card.TLabel")
        self.lbl_applied.pack(anchor=tk.W, padx=5, pady=2)

        self.lbl_docs = ttk.Label(status_frame, text="Документов: 0", style="Card.TLabel")
        self.lbl_docs.pack(anchor=tk.W, padx=5, pady=2)

        self.lbl_uptime = ttk.Label(status_frame, text="Время работы: 0:00", style="Card.TLabel")
        self.lbl_uptime.pack(anchor=tk.W, padx=5, pady=2)

        # Лог событий
        log_frame = ttk.LabelFrame(bottom_frame, text=" Лог событий ", style="Card.TFrame")
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.log_text = tk.Text(log_frame, bg=COLOR_OVERLAY, fg=COLOR_FG,
                                font=("Consolas", 9), wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)

        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.start_time = time.time()

    def _create_card(self, parent, title, value, color):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=10)
        ttk.Label(frame, text=title, style="CardTitle.TLabel").pack(anchor=tk.W)
        self_val = ttk.Label(frame, text=value, style="Value.TLabel")
        self_val.pack(anchor=tk.W)
        self_unit = ttk.Label(frame, text="", style="Unit.TLabel")
        self_unit.pack(anchor=tk.W)
        frame._value_label = self_val
        frame._unit_label = self_unit
        frame._color = color
        return frame

    def _create_chart(self, parent, title):
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=title, style="CardTitle.TLabel").pack(anchor=tk.W, padx=5, pady=(5, 0))
        canvas = tk.Canvas(frame, bg=COLOR_OVERLAY, highlightthickness=0)
        canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        canvas._frame = frame
        return canvas

    def _draw_chart(self, canvas, data, max_val=100, color=COLOR_GREEN):
        canvas.delete("all")
        w = canvas.winfo_width()
        h = canvas.winfo_height()
        if w < 10 or h < 10:
            return

        # Сетка
        for i in range(5):
            y = int(h * i / 4)
            canvas.create_line(0, y, w, y, fill=COLOR_OVERLAY, dash=(2, 4))
            val = int(max_val * (4 - i) / 4)
            canvas.create_text(5, y + 2, text=str(val), fill=COLOR_FG, anchor=tk.W, font=("Consolas", 8))

        if not data:
            return

        # Линия графика
        points = []
        n = len(data)
        for i, val in enumerate(data):
            x = int(w * i / max(n - 1, 1))
            y = int(h * (1 - min(val, max_val) / max_val))
            points.append((x, y))

        if len(points) >= 2:
            flat = []
            for x, y in points:
                flat.extend([x, y])
            canvas.create_line(*flat, fill=color, width=2, smooth=True)

            # Заливка под линией
            fill_points = flat + [w, h, 0, h]
            canvas.create_polygon(*fill_points, fill=color, outline="", stipple="gray25")

    def _log_event(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.events.append(f"[{timestamp}] {message}")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _update_metrics(self):
        """Собирает и обновляет метрики."""
        try:
            if not psutil:
                self.cpu_history.append(0)
                self.ram_history.append(0)
                self.process_cpu_history.append(0)
                self.process_ram_history.append(0)
                return

            cpu = psutil.cpu_percent(interval=0)
            ram = psutil.virtual_memory()
            ram_mb = ram.used / (1024 * 1024)

            self.cpu_history.append(cpu)
            self.ram_history.append(ram_mb)

            proc = psutil.Process(os.getpid())
            proc_cpu = proc.cpu_percent(interval=0)
            proc_mem = proc.memory_info().rss / (1024 * 1024)

            self.process_cpu_history.append(proc_cpu)
            self.process_ram_history.append(proc_mem)

            self.card_cpu._value_label.config(text=f"{cpu:.1f}%")
            self.card_cpu._unit_label.config(text=f"cores: {psutil.cpu_count()}")

            self.card_ram._value_label.config(text=f"{ram_mb:.0f} MB")
            self.card_ram._unit_label.config(text=f"total: {ram.total // (1024*1024)} MB")

            self.card_proc_cpu._value_label.config(text=f"{proc_cpu:.1f}%")
            self.card_proc_cpu._unit_label.config(text=f"pid: {os.getpid()}")

            self.card_proc_ram._value_label.config(text=f"{proc_mem:.0f} MB")
            self.card_proc_ram._unit_label.config(text=f"peak: {proc.memory_info().peak_wset // (1024*1024) if hasattr(proc.memory_info(), 'peak_wset') else 'N/A'} MB")

            # Цвет карточек по нагрузке
            if cpu > 80:
                self.card_cpu._value_label.config(foreground=COLOR_RED)
            elif cpu > 50:
                self.card_cpu._value_label.config(foreground=COLOR_YELLOW)
            else:
                self.card_cpu._value_label.config(foreground=COLOR_GREEN)

            # Графики
            self._draw_chart(self.cpu_canvas, self.cpu_history.get_list(), 100, COLOR_GREEN)
            self._draw_chart(self.ram_canvas, self.ram_history.get_list(), max_val=max(self.ram_history.last() * 1.2, 100), color=COLOR_BLUE)

            # Статус обработки
            if self.controller:
                self._update_processing_status()

            # Uptime
            elapsed = int(time.time() - self.start_time)
            mins, secs = divmod(elapsed, 60)
            hours, mins = divmod(mins, 60)
            self.lbl_uptime.config(text=f"Время работы: {hours}:{mins:02d}:{secs:02d}")

        except Exception as e:
            logger.debug(f"Ошибка обновления метрик: {e}")

    def _update_processing_status(self):
        """Обновляет статус обработки из контроллера."""
        try:
            ctrl = self.controller
            if hasattr(ctrl, 'current_issues') and ctrl.current_issues:
                total = len(ctrl.current_issues)
                self.lbl_issues.config(text=f"Проблем: {total}")

                by_cat = {}
                for issue in ctrl.current_issues:
                    cat = issue.category
                    by_cat[cat] = by_cat.get(cat, 0) + 1
                details = ", ".join(f"{k}:{v}" for k, v in sorted(by_cat.items()))
                self.lbl_issues.config(text=f"Проблем: {total} ({details})")
            else:
                self.lbl_issues.config(text="Проблем: 0")

            if hasattr(ctrl, 'doc_path') and ctrl.doc_path:
                self.lbl_docs.config(text=f"Документ: {os.path.basename(ctrl.doc_path)}")
            else:
                self.lbl_docs.config(text="Документ: не загружен")

            if hasattr(ctrl, 'apply_engine') and ctrl.apply_engine:
                applied = getattr(ctrl.apply_engine, 'applied_count', 0)
                failed = getattr(ctrl.apply_engine, 'failed_count', 0)
                self.lbl_applied.config(text=f"Исправлено: {applied}, ошибок: {failed}")
            else:
                self.lbl_applied.config(text="Исправлено: 0")

        except Exception as e:
            logger.debug(f"Ошибка обновления статуса: {e}")

    def _tick(self):
        """Периодическое обновление."""
        if not self.running:
            return
        self._update_metrics()
        self.window.after(self.update_interval, self._tick)

    def start(self):
        """Запуск мониторинга."""
        self.running = True
        self._log_event("Мониторинг запущен")
        self.window.after(500, self._tick)

    def stop(self):
        """Остановка мониторинга."""
        self.running = False
        self._log_event("Мониторинг остановлен")

    def log(self, message):
        """Публичный метод для добавления сообщения в лог."""
        self.window.after(0, lambda: self._log_event(message))

    def update_status(self, message):
        """Публичный метод для обновления статуса."""
        self.window.after(0, lambda: self.lbl_status.config(text=message))

    def _on_close(self):
        self.stop()
        self.window.destroy()


# Глобальная ссылка для доступа из GUI
_monitor_instance = None


def open_monitor(root, controller=None):
    """Открывает окно мониторинга (или фокусирует существующее)."""
    global _monitor_instance
    if _monitor_instance and _monitor_instance.window.winfo_exists():
        _monitor_instance.window.lift()
        _monitor_instance.window.focus_force()
        return _monitor_instance
    _monitor_instance = MonitorWindow(root, controller)
    return _monitor_instance


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    monitor = open_monitor(root)
    root.mainloop()

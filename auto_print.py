"""
Авто-печать Excel файлов из папки Загрузки.

Мониторит папку Загрузки. При появлении нового файла Excel,
начинающегося с "ОС-2" или "М11", показывает диалог подтверждения.
При согласии — печатает на A4, альбомная ориентация, вписать на одну страницу.
"""

import os
import sys
import time
import queue
import threading
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ─── Настройки ───────────────────────────────────────────────────────────────

DOWNLOADS_FOLDER = str(Path.home() / "Downloads")

# Префиксы файлов для отслеживания
PREFIXES = ("ОС-2", "М11")

# Допустимые расширения Excel
EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".xlsb")

# ─── Вспомогательные функции ─────────────────────────────────────────────────


def is_matching_file(filepath: str) -> bool:
    """Проверяет, что файл — Excel и начинается с одного из PREFIX."""
    name = os.path.basename(filepath)
    return (
        any(name.startswith(p) for p in PREFIXES)
        and any(name.lower().endswith(ext) for ext in EXTENSIONS)
    )


def print_excel_file(filepath: str) -> None:
    """Открывает Excel через COM и печатает файл с заданными параметрами."""
    try:
        import win32com.client  # noqa: PLC0415  (доступно только на Windows)
    except ImportError:
        messagebox.showerror(
            "Ошибка",
            "Не удалось импортировать win32com.client.\n"
            "Установите пакет pywin32:\n  pip install pywin32",
        )
        return

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(filepath))

        for i in range(1, wb.Sheets.Count + 1):
            ps = wb.Sheets(i).PageSetup
            ps.PaperSize = 9        # xlPaperA4
            ps.Orientation = 2      # xlLandscape
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False

        wb.PrintOut()
        wb.Close(False)

    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Ошибка при печати", str(exc))

    finally:
        try:
            excel.Quit()
        except Exception:  # noqa: BLE001
            pass


# ─── Очередь файлов (поток наблюдателя → главный поток Tkinter) ──────────────

_file_queue: queue.Queue = queue.Queue()


class ExcelFileHandler(FileSystemEventHandler):
    """Обработчик событий файловой системы для watchdog."""

    def _enqueue(self, filepath: str) -> None:
        if is_matching_file(filepath):
            _file_queue.put(filepath)

    def on_created(self, event):
        """Новый файл создан (в т.ч. скачан напрямую)."""
        if not event.is_directory:
            # Небольшая пауза, чтобы файл успел полностью записаться
            time.sleep(1.5)
            self._enqueue(event.src_path)

    def on_moved(self, event):
        """Файл переименован (браузеры: .crdownload → .xlsx и т.п.)."""
        if not event.is_directory:
            time.sleep(0.3)
            self._enqueue(event.dest_path)


# ─── Опрос очереди в главном потоке ──────────────────────────────────────────


def _poll_queue(root: tk.Tk) -> None:
    """Вызывается каждые 500 мс из главного цикла Tkinter."""
    try:
        while True:
            filepath = _file_queue.get_nowait()
            _ask_and_print(root, filepath)
    except queue.Empty:
        pass
    finally:
        root.after(500, _poll_queue, root)


def _ask_and_print(root: tk.Tk, filepath: str) -> None:
    """Показывает диалог поверх всех окон и при согласии запускает печать."""
    filename = os.path.basename(filepath)

    # Вспомогательное окно-«родитель» для диалога, чтобы он был поверх всех окон
    top = tk.Toplevel(root)
    top.attributes("-topmost", True)
    top.withdraw()

    answer = messagebox.askyesno(
        "Печать файла",
        f"Обнаружен новый файл:\n{filename}\n\nРаспечатать?",
        parent=top,
    )
    top.destroy()

    if answer:
        # Печать в отдельном потоке, чтобы не блокировать интерфейс
        threading.Thread(
            target=print_excel_file,
            args=(filepath,),
            daemon=True,
        ).start()


# ─── Точка входа ─────────────────────────────────────────────────────────────


def main() -> None:
    # Инициализация Tkinter (скрытое главное окно)
    root = tk.Tk()
    root.withdraw()
    root.title("Авто-печать Excel")

    # Запуск наблюдателя за папкой Загрузки
    observer = Observer()
    observer.schedule(ExcelFileHandler(), DOWNLOADS_FOLDER, recursive=False)
    observer.start()

    # Запуск опроса очереди
    root.after(500, _poll_queue, root)

    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass
    finally:
        observer.stop()
        observer.join()


if __name__ == "__main__":
    main()

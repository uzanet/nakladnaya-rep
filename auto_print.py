"""
Авто-печать Excel файлов из папки Загрузки.

Мониторит папку Загрузки. При появлении нового файла Excel,
начинающегося с "ОС-2" или "М11", показывает диалог подтверждения.
При согласии — печатает на A4, альбомная ориентация, вписать на одну страницу.
"""

import ctypes
import ctypes.wintypes
import os
import sys
import time
import queue
import threading
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

import pystray
from PIL import Image, ImageDraw

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ─── Настройки ───────────────────────────────────────────────────────────────

# Префиксы файлов для отслеживания
PREFIXES = ("ОС_2", "М11")

# Допустимые расширения Excel
EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".xlsb")

# Окно подавления дублирующих событий watchdog для одного файла (секунды)
_DEDUP_WINDOW = 3.0

# ─── Вспомогательные функции ─────────────────────────────────────────────────


def _get_downloads_folder() -> str:
    """
    Возвращает путь к папке Загрузки через Windows API (не зависит от локали).
    Fallback — ~/Downloads.
    """
    FOLDERID_Downloads = "{374DE290-123F-4565-9164-39C4925E467B}"
    try:
        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        guid = GUID()
        ctypes.windll.ole32.CLSIDFromString(FOLDERID_Downloads, ctypes.byref(guid))

        path_ptr = ctypes.c_wchar_p()
        result = ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(guid), 0, None, ctypes.byref(path_ptr)
        )
        path = path_ptr.value if result == 0 else None
        ctypes.windll.ole32.CoTaskMemFree(path_ptr)
        if path:
            return path
    except Exception:  # noqa: BLE001
        pass

    return str(Path.home() / "Downloads")


DOWNLOADS_FOLDER = _get_downloads_folder()


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
        import pythoncom
        import win32com.client  # noqa: PLC0415  (доступно только на Windows)
    except ImportError:
        _file_queue.put(("error", "Не удалось импортировать win32com.\nУстановите pywin32."))
        return

    # COM должен быть инициализирован в каждом новом потоке
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(filepath))

        for i in range(1, wb.Sheets.Count + 1):
            ps = wb.Sheets(i).PageSetup
            # Zoom = False обязательно до FitToPages-свойств
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1   # 1 = вписать; False/0 сбрасывает режим
            ps.PaperSize = 9        # xlPaperA4
            ps.Orientation = 2      # xlLandscape

        wb.PrintOut()
        wb.Close(False)

    except Exception as exc:  # noqa: BLE001
        _file_queue.put(("error", str(exc)))

    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                pass
        pythoncom.CoUninitialize()


# ─── Очередь файлов (поток наблюдателя → главный поток Tkinter) ──────────────

_file_queue: queue.Queue = queue.Queue()


class ExcelFileHandler(FileSystemEventHandler):
    """Обработчик событий файловой системы для watchdog."""

    def __init__(self) -> None:
        super().__init__()
        self._recently_enqueued: dict[str, float] = {}
        self._lock = threading.Lock()

    def _enqueue(self, filepath: str) -> None:
        if not is_matching_file(filepath):
            return
        now = time.monotonic()
        with self._lock:
            if now - self._recently_enqueued.get(filepath, 0.0) < _DEDUP_WINDOW:
                return  # дубликат события — пропускаем
            self._recently_enqueued[filepath] = now
            # Чистим старые записи, чтобы словарь не рос бесконечно
            cutoff = now - 60.0
            self._recently_enqueued = {
                k: v for k, v in self._recently_enqueued.items() if v > cutoff
            }
        _file_queue.put(filepath)

    def on_created(self, event):
        """Новый файл создан (в т.ч. скачан напрямую)."""
        if not event.is_directory:
            time.sleep(1.5)
            self._enqueue(event.src_path)

    def on_moved(self, event):
        """Файл переименован (браузеры: .crdownload → .xlsx и т.п.)."""
        if not event.is_directory:
            time.sleep(0.3)
            self._enqueue(event.dest_path)

    def on_modified(self, event):
        """Файл перезаписан (браузер скачивает поверх существующего файла)."""
        if not event.is_directory:
            time.sleep(1.5)
            self._enqueue(event.src_path)


# ─── Опрос очереди в главном потоке ──────────────────────────────────────────


def _poll_queue(root: tk.Tk) -> None:
    """Вызывается каждые 500 мс из главного цикла Tkinter."""
    try:
        while True:
            item = _file_queue.get_nowait()
            if isinstance(item, tuple) and item[0] == "error":
                messagebox.showerror("Ошибка при печати", item[1])
            else:
                _ask_and_print(root, item)
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


# ─── Системный трей ───────────────────────────────────────────────────────────


def _create_tray_image() -> Image.Image:
    """Зелёный квадрат 64×64 с белым крестом — иконка в трее."""
    img = Image.new("RGB", (64, 64), color="#217346")  # Excel green
    draw = ImageDraw.Draw(img)
    m, w, lw = 12, 64, 8
    draw.line([(m, m), (w - m, w - m)], fill="white", width=lw)
    draw.line([(w - m, m), (m, w - m)], fill="white", width=lw)
    return img


def _create_tray_icon(root: tk.Tk) -> pystray.Icon:
    """Создаёт иконку в области уведомлений и запускает её в фоновом потоке."""

    def on_quit(icon: pystray.Icon, item: pystray.MenuItem) -> None:
        icon.stop()
        root.after(0, root.quit)

    menu = pystray.Menu(
        pystray.MenuItem("Авто-печать Excel", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Выход", on_quit),
    )
    icon = pystray.Icon(
        "auto_print_excel",
        _create_tray_image(),
        "Авто-печать Excel",
        menu,
    )
    threading.Thread(target=icon.run, daemon=True).start()
    return icon


# ─── Точка входа ─────────────────────────────────────────────────────────────


def main() -> None:
    # Инициализация Tkinter (скрытое главное окно)
    root = tk.Tk()
    root.withdraw()
    root.title("Авто-печать Excel")

    # Иконка в системном трее
    _create_tray_icon(root)

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

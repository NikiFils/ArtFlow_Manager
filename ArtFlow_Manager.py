# -*- coding: utf-8 -*-
"""
ArtFlow Manager v1.0
Массовая обработка папок и файлов по списку артикулов.
© 2026 NikiFils
"""

import os
import sys
import json
import shutil
import threading
import queue
import datetime
import fnmatch
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    from send2trash import send2trash
    HAS_TRASH = True
except ImportError:
    HAS_TRASH = False

try:
    from openpyxl import load_workbook
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ============================================================
#  КОНСТАНТЫ ПРОГРАММЫ
# ============================================================
APP_NAME = "ArtFlow Manager"
APP_VERSION = "1.0"
APP_COPYRIGHT = "© 2026 NikiFils"
APP_DESCRIPTION = ("Массовая обработка папок и файлов\n"
                   "по списку артикулов.")
APP_FEATURES = [
    "• Сканирование папок по маскам и фильтрам",
    "• Сравнение шаблонов (есть _pr, нет _01)",
    "• Копирование, перемещение, удаление в корзину",
    "• Переименование с заменой подстрок",
    "• Работа с папками и файлами внутри — отдельно или вместе",
    "• Фильтр по списку артикулов (txt/csv/xlsx/docx)",
    "• Многопоточная обработка файлов",
    "• Откат последней операции",
]


# ============================================================
#  ПУТИ И КОНСТАНТЫ
# ============================================================
if getattr(sys, 'frozen', False):
    APPLICATION_PATH = os.path.dirname(sys.executable)
else:
    APPLICATION_PATH = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(APPLICATION_PATH, "settings.json")
JOURNAL_FILE = os.path.join(APPLICATION_PATH, "last_operation.json")

VK_V, VK_C, VK_X, VK_A = 86, 67, 88, 65
VK_INSERT, VK_DELETE = 45, 46

EXT_ALL = "Все файлы"
EXT_IMAGES = "Только изображения"
EXT_JPG = "Только .jpg / .jpeg"
EXT_PNG = "Только .png"
EXT_CUSTOM = "Свои расширения..."

IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.tiff', '.tif'}

DEFAULT_SCAN = {
    'mask': '*_01', 'compare_mask': '*_01', 'compare': False,
    'compare_mode': 'missing_pair', 'find_mode': 'not_exists',
    'ext_choice': EXT_ALL, 'custom_ext': '.jpg, .png', 'exts': None,
    'use_list': False, 'list_articles': [], 'list_match': 'exact',
}

DEFAULT_GENERAL = {
    'use_thread_pool': False, 'thread_count': 4,
    'auto_open_last_path': False, 'save_editor_session': True,
    'save_scan_settings': True, 'show_tooltips': True,
}

LIST_CONTAINS_WARN_LIMIT = 500


# ============================================================
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================
def norm(p):
    if not p: return p
    return os.path.normpath(p)


def make_matcher(mask):
    m = mask.strip()
    if not m: return lambda name, ext: False
    if m.startswith('*') and '*' not in m[1:]:
        suffix = m[1:].lower()
        if '.' not in suffix:
            return lambda name, ext: name.lower().endswith(suffix)
        return lambda name, ext: (name + ext).lower().endswith(suffix)
    return lambda name, ext: fnmatch.fnmatch((name + ext).lower(), m.lower())


def make_suffix(mask):
    return mask.replace('*', '').rsplit('.', 1)[0].lower()


def extract_article(name, suffix):
    if suffix and name.lower().endswith(suffix):
        return name[:len(name) - len(suffix)]
    return name


def parse_date(s):
    s = s.strip()
    if not s: return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try: return datetime.datetime.strptime(s, fmt).timestamp()
        except ValueError: continue
    return None


# ============================================================
#  ВСПЛЫВАЮЩИЕ ПОДСКАЗКИ
# ============================================================
class ToolTip:
    _current = None

    def __init__(self, widget, text, app_ref):
        self.widget = widget
        self.text = text
        self.app = app_ref
        self.tip_window = None
        self._after_id = None
        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, event=None):
        if not getattr(self.app, 'tooltips_enabled', True): return
        if self._after_id:
            self.widget.after_cancel(self._after_id)
        self._after_id = self.widget.after(700, self._show)

    def _on_leave(self, event=None):
        if self._after_id:
            try: self.widget.after_cancel(self._after_id)
            except Exception: pass
            self._after_id = None
        self._hide()

    def _show(self):
        if self.tip_window or not self.text: return
        if not getattr(self.app, 'tooltips_enabled', True): return
        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        except Exception: return
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        try: tw.attributes("-topmost", True)
        except Exception: pass
        frame = tk.Frame(tw, background="#ffffe0",
                         relief=tk.SOLID, borderwidth=1)
        frame.pack()
        tk.Label(frame, text=self.text, justify=tk.LEFT,
                 background="#ffffe0", foreground="#333",
                 font=("Segoe UI", 9), padx=8, pady=4).pack()

    def _hide(self):
        if self.tip_window:
            try: self.tip_window.destroy()
            except Exception: pass
            self.tip_window = None


# ============================================================
#  ЖУРНАЛ ОПЕРАЦИЙ
# ============================================================
def save_journal(data):
    data2 = dict(data)
    data2['files'] = [{'from': norm(it['from']), 'to': norm(it['to'])}
                      for it in data.get('files', [])]
    data2['folders'] = [{'from': norm(it['from']), 'to': norm(it['to'])}
                        for it in data.get('folders', [])]
    data2['deleted'] = [{'path': norm(it['path'])}
                        for it in data.get('deleted', [])]
    try:
        with open(JOURNAL_FILE, 'w', encoding='utf-8') as f:
            json.dump(data2, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Ошибка записи журнала: {e}")


def load_journal():
    if not os.path.exists(JOURNAL_FILE): return None
    try:
        with open(JOURNAL_FILE, 'r', encoding='utf-8') as f:
            j = json.load(f)
        j['files'] = [{'from': norm(it['from']), 'to': norm(it['to'])}
                      for it in j.get('files', [])]
        j['folders'] = [{'from': norm(it['from']), 'to': norm(it['to'])}
                        for it in j.get('folders', [])]
        j['deleted'] = [{'path': norm(it['path'])}
                        for it in j.get('deleted', [])]
        return j
    except Exception as e:
        print(f"Ошибка чтения журнала: {e}")
        return None


# ============================================================
#  ЗАГРУЗКА СПИСКОВ
# ============================================================
def load_list_from_txt(path):
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        return [line.strip() for line in f if line.strip()]


def load_list_from_csv(path, column=0):
    result = []
    with open(path, 'r', encoding='utf-8', errors='ignore', newline='') as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) > column:
                v = row[column].strip()
                if v: result.append(v)
    return result


def load_list_from_xlsx(path, column=0):
    if not HAS_XLSX:
        raise ImportError('openpyxl не установлен. Установите: pip install openpyxl')
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    result = []
    for row in ws.iter_rows(values_only=True):
        if row is not None and len(row) > column:
            v = row[column]
            if v is not None:
                s = str(v).strip()
                if s: result.append(s)
    wb.close()
    return result


def load_list_from_docx(path):
    if not HAS_DOCX:
        raise ImportError('python-docx не установлен. Установите: pip install python-docx')
    doc = Document(path)
    result = []
    if doc.tables:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    v = cell.text.strip()
                    if v: result.append(v)
    if not result:
        for p in doc.paragraphs:
            v = p.text.strip()
            if v: result.append(v)
    return result


# ============================================================
#  ТЕКСТ СПРАВКИ
# ============================================================
HELP_TEXT = """\
ArtFlow Manager — краткое руководство
═══════════════════════════════════════════════════════════

ОБЩЕЕ
─────
Программа предназначена для массовой обработки папок и
файлов по списку артикулов: сканирование, копирование,
перемещение, переименование и удаление в корзину.

Работает в два этапа:
  1. Сканирование — находит нужные папки и файлы.
  2. Редактор — выполняет действия с результатами.

═══════════════════════════════════════════════════════════

ВЕРХНЯЯ ПАНЕЛЬ
──────────────
  Папка      — корневая папка для сканирования.
  Обзор...   — выбор папки в диалоге.
  ⚙ Настройки — общие настройки программы.
  📖 Справка  — это окно.
  ? О программе — сведения о программе.

═══════════════════════════════════════════════════════════

БЛОК «АУДИТ (СКАНИРОВАНИЕ)»
──────────────────────────
  🔍 Сканировать      — запустить обход папки.
  ⏹ Остановить        — прервать сканирование.
  ⚙ Настройки скана   — открыть параметры поиска.

Под строкой кнопок — информация о текущих настройках скана.

───────────────────────────────────────────────────────────
НАСТРОЙКИ СКАНИРОВАНИЯ
───────────────────────────────────────────────────────────
Шаблон поиска
  Маска имени файла. Примеры:
    *_01         — имена, оканчивающиеся на _01
    *_pr         — имена, оканчивающиеся на _pr
    *.psd        — расширение .psd
    abc*         — начинается с abc
    *test*       — содержит test

Расширения
  Все файлы          — без фильтра.
  Только изображения — jpg, png, webp, bmp, gif, tiff.
  Только .jpg/.jpeg  — только эти два.
  Только .png        — только png.
  Свои расширения... — список через запятую.

Сравнение с другим шаблоном
  Позволяет найти папки, где есть файлы одного шаблона,
  но нет файлов другого. Примеры:
    Есть *_pr, НЕТ *_01
    Есть *_01, НЕТ *_pr
  Удобно для поиска артикулов, где нет второй версии файла.

Без сравнения — что показывать
  Где они ЕСТЬ    — показывать папки, где найден шаблон.
  Где их НЕТ      — показывать папки БЕЗ шаблона.

Фильтр по списку артикулов
  Ограничивает поиск только теми артикулами, которые есть
  в списке. Список можно:
    • вставить руками;
    • загрузить из .txt (одна строка — один артикул);
    • загрузить из .csv (укажите номер столбца);
    • загрузить из .xlsx (укажите номер столбца);
    • загрузить из .docx (из таблиц или абзацев).

  Режим совпадения:
    Точное   — имя папки полностью совпадает с артикулом.
    Содержит — артикул является частью имени папки.

═══════════════════════════════════════════════════════════

БЛОК «РЕДАКТОР»
───────────────
Источник
  Отмеченное из таблицы скана — работаем с выбранными
    папками из результатов скана.
  Папка вручную — указать папку для обработки без скана.

  📋 Открыть таблицу — отметить папки для обработки.

───────────────────────────────────────────────────────────
РАБОТА С ПАПКАМИ
───────────────────────────────────────────────────────────
  Галочка «Обрабатывать папки» — включить обработку самих
  папок (переименование / копирование / перемещение /
  удаление).

  Действие:
    Копировать       — создаётся копия папки.
    Переместить      — папка переезжает в новое место.
    Удалить (в корзину) — папка целиком удаляется в Корзину.

  Изменить имя:
    Заменить [X] на [Y] — все вхождения X заменяются на Y.
    Если оставить поле «на» пустым — подстрока удаляется.
    Если поле «Заменить» пустое, а «на» заполнено —
      подстрока добавляется в начало или в конец имени.

───────────────────────────────────────────────────────────
РАБОТА С ФАЙЛАМИ ВНУТРИ
───────────────────────────────────────────────────────────
  Галочка «Обрабатывать файлы» — включить обработку файлов
  внутри папок.

  Что:
    Все файлы         — обрабатывать все.
    Файлы по фильтру  — только соответствующие фильтру
                        (расширения, маска, размер, даты).

  Действие — как для папок (копировать / переместить /
  удалить).

  Изменить имя — как для папок.

  ВАЖНО: если папка выбрана на удаление, файлы внутри не
  обрабатываются — они удалятся вместе с папкой.

───────────────────────────────────────────────────────────
КУДА (общее для папок и файлов)
───────────────────────────────────────────────────────────
  В другую папку — указать путь назначения.
  В ту же папку  — работать на месте (только для
                   переименования).

  Если «В ту же папку» + Копировать/Переместить без
  изменения имени — действие запрещено (копия себя же).

───────────────────────────────────────────────────────────
ОПЦИИ
───────────────────────────────────────────────────────────
  Пропускать существующие — не перезаписывать файлы и
    папки, которые уже есть в целевой папке.
  Обновлять дату файлов — после копирования/перемещения
    проставить файлам текущую дату.
  Сохранять структуру подпапок — при копировании файлов
    в другую папку сохранять вложенность.

═══════════════════════════════════════════════════════════

КНОПКИ ЗАПУСКА
──────────────
  ▶ Применить — выполнить операцию.
  ⏹ Остановить — прервать операцию.
  ↶ Отменить последнюю операцию — откатить последнюю
    операцию (по журналу last_operation.json).

  Журнал — сведения о последней операции.

═══════════════════════════════════════════════════════════

ВКЛАДКИ ВНИЗУ
─────────────
  Сканирование — отчёт — лог сканирования.
  Редактор — отчёт     — лог операций редактора.

  Кнопки «Очистить» и «💾 Сохранить отчёт» — управление
  каждым логом отдельно.

═══════════════════════════════════════════════════════════

ЛОГИКА РАБОТЫ
─────────────
  1. Папка → Обзор → выбрать корень.
  2. Настроить скан (маска, расширения, опции).
  3. 🔍 Сканировать.
  4. Открыть таблицу → отметить нужные папки.
  5. Настроить редактор (источник, действие, куда).
  6. ▶ Применить.
  7. При необходимости — ↶ Отменить.

При обработке СНАЧАЛА применяются действия к файлам внутри,
ЗАТЕМ — к самим папкам. Это сделано для сохранения путей.

═══════════════════════════════════════════════════════════

ОТКАТ ОПЕРАЦИИ
──────────────
  Программа ведёт журнал всех операций (last_operation.json).
  Кнопка «↶ Отменить последнюю операцию»:
    • Копирование   → удаляет созданные копии.
    • Перемещение   → возвращает объекты на место.
    • Удаление      → НЕ ОТКАТЫВАЕТСЯ, восстановите
                       вручную из Корзины Windows.
  После успешного отката журнал очищается.

═══════════════════════════════════════════════════════════

ГОРЯЧИЕ КЛАВИШИ
───────────────
  Ctrl+C / Ctrl+Insert — копировать выделенное.
  Ctrl+V / Shift+Insert — вставить.
  Ctrl+X / Shift+Delete — вырезать.
  Ctrl+A — выделить всё.
  ПКМ — контекстное меню в полях ввода.

═══════════════════════════════════════════════════════════

ВСПЛЫВАЮЩИЕ ПОДСКАЗКИ
─────────────────────
  При наведении мыши на элементы управления появляются
  краткие пояснения.
  Отключить/включить можно в ⚙ Настройках.

═══════════════════════════════════════════════════════════

© 2026 NikiFils
"""


# ============================================================
#  ОКНО СПРАВКИ
# ============================================================
class HelpDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(f"Справка — {APP_NAME}")
        self.geometry("720x640")
        self.minsize(500, 400)
        self.transient(parent)

        head = tk.Frame(self)
        head.pack(fill=tk.X, padx=10, pady=(10, 4))
        tk.Label(head, text="📖 Справка", font=("Segoe UI", 14, "bold"),
                 fg="#2e7d32").pack(side=tk.LEFT)
        tk.Label(head, text=f"   {APP_NAME} v{APP_VERSION}",
                 font=("Segoe UI", 10), fg="#666").pack(side=tk.LEFT, pady=(4, 0))

        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, padx=10, pady=(0, 6))

        text_frame = tk.Frame(self)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        self.text = scrolledtext.ScrolledText(
            text_frame, wrap=tk.WORD, font=("Consolas", 10),
            bg="#fafafa", fg="#222", padx=12, pady=8)
        self.text.pack(fill=tk.BOTH, expand=True)
        self.text.insert(1.0, HELP_TEXT)
        self.text.config(state=tk.DISABLED)

        btn = tk.Frame(self)
        btn.pack(fill=tk.X, side=tk.BOTTOM, pady=10, padx=10)
        tk.Button(btn, text="Закрыть", command=self.destroy,
                  font=("Segoe UI", 10), padx=15, pady=4,
                  bg="#e1f5fe").pack(side=tk.RIGHT)

        self.bind('<Escape>', lambda e: self.destroy())

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


# ============================================================
#  О ПРОГРАММЕ
# ============================================================
class AboutDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("О программе")
        self.geometry("480x480")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        head = tk.Frame(self)
        head.pack(fill=tk.X, padx=20, pady=(20, 5))
        tk.Label(head, text=APP_NAME, font=("Segoe UI", 16, "bold"),
                 fg="#2e7d32").pack(anchor=tk.W)
        tk.Label(head, text=f"Версия {APP_VERSION}", font=("Segoe UI", 11),
                 fg="#555").pack(anchor=tk.W, pady=(2, 0))

        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)

        desc_frame = tk.Frame(self)
        desc_frame.pack(fill=tk.X, padx=20)
        tk.Label(desc_frame, text=APP_DESCRIPTION,
                 font=("Segoe UI", 10), justify=tk.LEFT, anchor=tk.W
                 ).pack(anchor=tk.W)

        tk.Label(self, text="Возможности:",
                 font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, padx=20, pady=(12, 4))
        feat_frame = tk.Frame(self)
        feat_frame.pack(fill=tk.X, padx=30)
        for line in APP_FEATURES:
            tk.Label(feat_frame, text=line, font=("Segoe UI", 9),
                     fg="#333", anchor=tk.W, justify=tk.LEFT).pack(anchor=tk.W)

        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, padx=20, pady=12)
        tk.Label(self, text=APP_COPYRIGHT, font=("Segoe UI", 9),
                 fg="#777").pack(anchor=tk.W, padx=20)

        # Кнопка OK — крупная, читаемая
        btn = tk.Frame(self)
        btn.pack(fill=tk.X, side=tk.BOTTOM, pady=15, padx=20)
        tk.Button(btn, text="OK", command=self.destroy,
                  font=("Segoe UI", 10), padx=20, pady=5,
                  bg="#e1f5fe").pack(side=tk.RIGHT)

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

        self.bind('<Return>', lambda e: self.destroy())
        self.bind('<Escape>', lambda e: self.destroy())


# ============================================================
#  ДИАЛОГ СПИСКА АРТИКУЛОВ
# ============================================================
class ArticleListDialog(tk.Toplevel):
    def __init__(self, parent, current_list, current_match):
        super().__init__(parent)
        self.title('Список артикулов')
        self.geometry('640x640')
        self.transient(parent); self.grab_set()
        self.result = None
        self.match_mode = tk.StringVar(value=current_match)
        self._build(); self._load(current_list)
        self.protocol('WM_DELETE_WINDOW', self._cancel)
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f'+{x}+{y}')

    def _build(self):
        top = tk.Frame(self); top.pack(fill=tk.X, padx=10, pady=(10, 0))
        tk.Label(top, text='Вставьте список — один артикул на строку:',
                 font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W)
        self.text = scrolledtext.ScrolledText(self, wrap=tk.NONE, font=('Consolas', 10))
        self.text.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)
        bar = tk.Frame(self); bar.pack(fill=tk.X, padx=10, pady=2)
        tk.Button(bar, text='📁 .txt', command=lambda: self._load_file('txt'), width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text='📊 .csv', command=lambda: self._load_file('csv'), width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text='📗 .xlsx', command=lambda: self._load_file('xlsx'), width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text='📘 .docx', command=lambda: self._load_file('docx'), width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text='🗑 Очистить', command=self._clear, width=12).pack(side=tk.LEFT, padx=2)
        opt = tk.Frame(self); opt.pack(fill=tk.X, padx=10, pady=4)
        tk.Label(opt, text='Совпадение:', font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT)
        tk.Radiobutton(opt, text='Точное', variable=self.match_mode, value='exact').pack(side=tk.LEFT, padx=6)
        tk.Radiobutton(opt, text='Содержит', variable=self.match_mode, value='contains').pack(side=tk.LEFT, padx=6)
        tk.Label(opt, text='(если артикул является частью названия папки)',
                 fg='#888', font=('Segoe UI', 7)).pack(side=tk.LEFT, padx=8)
        self.lbl_count = tk.Label(self, text='', anchor=tk.W, fg='#333', font=('Segoe UI', 9))
        self.lbl_count.pack(fill=tk.X, padx=10, pady=(4, 2))
        btn = tk.Frame(self); btn.pack(fill=tk.X, side=tk.BOTTOM, pady=10, padx=10)
        tk.Button(btn, text='Отмена', command=self._cancel, width=12).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn, text='OK', command=self._ok, width=12, bg='#e1f5fe').pack(side=tk.RIGHT)
        self.text.bind('<KeyRelease>', lambda e: self._update_count())
        self.text.bind('<<Paste>>', lambda e: self.after(50, self._update_count))

    def _update_count(self):
        self.lbl_count.config(text=f'Артикулов: {len(self._get_items())}')

    def _get_items(self):
        raw = self.text.get(1.0, tk.END)
        result, seen = [], set()
        for line in raw.splitlines():
            v = line.strip()
            if v and v not in seen:
                seen.add(v); result.append(v)
        return result

    def _load(self, items):
        self.text.delete(1.0, tk.END)
        if items: self.text.insert(1.0, '\n'.join(items))
        self._update_count()

    def _clear(self):
        self.text.delete(1.0, tk.END); self._update_count()

    def _load_file(self, kind):
        items = []
        try:
            if kind == 'txt':
                fp = filedialog.askopenfilename(title='Выберите текстовый файл',
                    filetypes=[('Текстовые файлы', '*.txt'), ('Все файлы', '*.*')])
                if not fp: return
                items = load_list_from_txt(fp)
            elif kind == 'csv':
                fp = filedialog.askopenfilename(title='Выберите CSV файл',
                    filetypes=[('CSV файлы', '*.csv'), ('Все файлы', '*.*')])
                if not fp: return
                col = self._ask_column('CSV', 'номер столбца (0 — первый)')
                if col is None: return
                items = load_list_from_csv(fp, col)
            elif kind == 'xlsx':
                if not HAS_XLSX:
                    messagebox.showerror('Нет библиотеки', 'Для .xlsx требуется openpyxl.\n\nУстановите:\npip install openpyxl')
                    return
                fp = filedialog.askopenfilename(title='Выберите Excel файл',
                    filetypes=[('Excel', '*.xlsx'), ('Все файлы', '*.*')])
                if not fp: return
                col = self._ask_column('Excel', 'номер столбца (0 = A, 1 = B, ...)')
                if col is None: return
                items = load_list_from_xlsx(fp, col)
            elif kind == 'docx':
                if not HAS_DOCX:
                    messagebox.showerror('Нет библиотеки', 'Для .docx требуется python-docx.\n\nУстановите:\npip install python-docx')
                    return
                fp = filedialog.askopenfilename(title='Выберите Word файл',
                    filetypes=[('Word', '*.docx'), ('Все файлы', '*.*')])
                if not fp: return
                items = load_list_from_docx(fp)
        except Exception as e:
            messagebox.showerror('Ошибка загрузки', f'{type(e).__name__}:\n{e}'); return
        if not items:
            messagebox.showinfo('Пусто', 'Не найдено ни одного артикула.'); return
        seen, uniq = set(), []
        for v in items:
            if v not in seen:
                seen.add(v); uniq.append(v)
        self._load(uniq)

    def _ask_column(self, kind, hint):
        dlg = tk.Toplevel(self)
        dlg.title(f'Столбец {kind}')
        dlg.geometry('340x180'); dlg.transient(self); dlg.grab_set()
        tk.Label(dlg, text=f'Из какого столбца брать данные?\n{hint}',
                 justify=tk.LEFT, font=('Segoe UI', 9)).pack(padx=10, pady=10)
        e = tk.Entry(dlg, width=10); e.pack(pady=4); e.insert(0, '0'); e.focus()
        result = {'v': None}
        def ok():
            try: result['v'] = int(e.get().strip())
            except ValueError:
                messagebox.showerror('Ошибка', 'Введите число'); return
            dlg.destroy()
        def cancel(): dlg.destroy()
        row = tk.Frame(dlg); row.pack(pady=8)
        tk.Button(row, text='OK', command=ok, width=10).pack(side=tk.LEFT, padx=4)
        tk.Button(row, text='Отмена', command=cancel, width=10).pack(side=tk.LEFT, padx=4)
        dlg.bind('<Return>', lambda e: ok())
        self.wait_window(dlg)
        return result['v']

    def _ok(self):
        items = self._get_items()
        mode = self.match_mode.get()
        if mode == 'contains' and len(items) > LIST_CONTAINS_WARN_LIMIT:
            if not messagebox.askyesno('Внимание',
                f'В списке {len(items)} артикулов.\nРежим «Содержит» может работать медленно.\n\nПродолжить?'):
                return
        self.result = (items, mode); self.destroy()

    def _cancel(self):
        self.result = None; self.destroy()


# ============================================================
#  ДИАЛОГ НАСТРОЕК СКАНИРОВАНИЯ
# ============================================================
class ScanSettingsDialog(tk.Toplevel):
    def __init__(self, parent, current, app_ref):
        super().__init__(parent)
        self.app = app_ref
        self.title("Настройки сканирования")
        self.geometry("560x680"); self.resizable(False, False)
        self.transient(parent); self.grab_set()
        self.result = None
        self.current = dict(current)
        self._build(); self._load()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _tt(self, widget, text):
        ToolTip(widget, text, self.app)

    def _build(self):
        main = tk.Frame(self); main.pack(fill=tk.BOTH, expand=True)
        content = tk.Frame(main); content.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
        tk.Label(content, text="Шаблон поиска:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        tk.Label(content, text="Пример: *_01 или *_pr или *.psd", fg="#666", font=("Segoe UI", 8)).pack(anchor=tk.W)
        self.entry_mask = tk.Entry(content, font=("Consolas", 10)); self.entry_mask.pack(fill=tk.X, pady=(2, 6))
        self._tt(self.entry_mask, "Маска имени файла.\n*_01 — оканчивается на _01\n*_pr — оканчивается на _pr\n*.psd — расширение psd\n*test* — содержит test")
        tk.Label(content, text="Расширения:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        self.ext_choice = tk.StringVar(value=EXT_ALL)
        for ext in [EXT_ALL, EXT_IMAGES, EXT_JPG, EXT_PNG, EXT_CUSTOM]:
            rb = tk.Radiobutton(content, text=ext, variable=self.ext_choice, value=ext,
                                command=self._on_ext)
            rb.pack(anchor=tk.W, padx=10)
            self._tt(rb, "Ограничение по расширениям файлов")
        self.entry_custom_ext = tk.Entry(content, font=("Consolas", 10))
        self.entry_custom_ext.pack(fill=tk.X, pady=(2, 2))
        self.entry_custom_ext.config(state=tk.DISABLED)
        tk.Label(content, text="Через запятую: .jpg, .png, .psd", fg="#666", font=("Segoe UI", 8)).pack(anchor=tk.W, pady=(0, 6))
        ttk.Separator(content, orient='horizontal').pack(fill=tk.X, pady=4)
        self.compare_var = tk.BooleanVar(value=False)
        chk_cmp = tk.Checkbutton(content, text="Сравнить с другим шаблоном",
                                 variable=self.compare_var, command=self._on_compare,
                                 font=("Segoe UI", 9, "bold"))
        chk_cmp.pack(anchor=tk.W)
        self._tt(chk_cmp, "Найти папки, где есть файлы одного шаблона,\nно нет файлов другого.\nПример: есть *_pr, нет *_01")
        self.frame_compare = tk.Frame(content); self.frame_compare.pack(fill=tk.X, pady=4)
        tk.Label(self.frame_compare, text="Сравнить с:").grid(row=0, column=0, sticky=tk.W)
        self.entry_compare = tk.Entry(self.frame_compare, font=("Consolas", 10))
        self.entry_compare.grid(row=0, column=1, sticky=tk.EW, padx=5)
        self.frame_compare.columnconfigure(1, weight=1)
        self._tt(self.entry_compare, "Второй шаблон для сравнения")
        self.compare_mode = tk.StringVar(value="missing_pair")
        tk.Radiobutton(self.frame_compare, text="Есть основной, НЕТ парного",
                       variable=self.compare_mode, value="missing_pair"
                       ).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)
        tk.Radiobutton(self.frame_compare, text="Есть парный, НЕТ основного",
                       variable=self.compare_mode, value="missing_main"
                       ).grid(row=2, column=0, columnspan=2, sticky=tk.W)
        ttk.Separator(content, orient='horizontal').pack(fill=tk.X, pady=4)
        tk.Label(content, text="Без сравнения — что показывать:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        self.find_mode = tk.StringVar(value="not_exists")
        tk.Radiobutton(content, text="Где они ЕСТЬ", variable=self.find_mode, value="exists").pack(anchor=tk.W, padx=10)
        tk.Radiobutton(content, text="Где их НЕТ", variable=self.find_mode, value="not_exists").pack(anchor=tk.W, padx=10)
        ttk.Separator(content, orient='horizontal').pack(fill=tk.X, pady=4)
        self.use_list_var = tk.BooleanVar(value=False)
        chk_list = tk.Checkbutton(content, text="Использовать фильтр по списку артикулов",
                                  variable=self.use_list_var, font=("Segoe UI", 9, "bold"))
        chk_list.pack(anchor=tk.W, pady=(4, 0))
        self._tt(chk_list, "Искать только артикулы из указанного списка")
        row_list = tk.Frame(content); row_list.pack(fill=tk.X, pady=(4, 0), padx=20)
        self.lbl_list_info = tk.Label(row_list, text="Список: пуст", anchor=tk.W,
                                      fg="#444", font=("Segoe UI", 8))
        self.lbl_list_info.pack(side=tk.LEFT)
        tk.Button(row_list, text="📝 Редактор списка", command=self._edit_list).pack(side=tk.RIGHT)
        btn = tk.Frame(main); btn.pack(fill=tk.X, side=tk.BOTTOM, pady=10, padx=10)
        tk.Button(btn, text="Отмена", command=self._cancel, width=12).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn, text="OK", command=self._ok, width=12, bg="#e1f5fe").pack(side=tk.RIGHT)

    def _edit_list(self):
        dlg = ArticleListDialog(self, self.current.get('list_articles', []),
                                self.current.get('list_match', 'exact'))
        self.wait_window(dlg)
        if dlg.result is not None:
            items, mode = dlg.result
            self.current['list_articles'] = items
            self.current['list_match'] = mode
            self._update_list_info()

    def _update_list_info(self):
        items = self.current.get('list_articles', [])
        mode = self.current.get('list_match', 'exact')
        mode_str = 'точное' if mode == 'exact' else 'содержит'
        self.lbl_list_info.config(text=f'Список: {len(items)} арт. ({mode_str})')

    def _on_ext(self):
        st = tk.NORMAL if self.ext_choice.get() == EXT_CUSTOM else tk.DISABLED
        self.entry_custom_ext.config(state=st)

    def _on_compare(self):
        st = tk.NORMAL if self.compare_var.get() else tk.DISABLED
        for c in self.frame_compare.winfo_children():
            try: c.config(state=st)
            except tk.TclError: pass

    def _load(self):
        s = self.current
        self.entry_mask.insert(0, s.get('mask', '*_01'))
        self.entry_compare.insert(0, s.get('compare_mask', '*_01'))
        self.ext_choice.set(s.get('ext_choice', EXT_ALL))
        self.entry_custom_ext.insert(0, s.get('custom_ext', '.jpg, .png'))
        self.compare_var.set(s.get('compare', False))
        self.compare_mode.set(s.get('compare_mode', 'missing_pair'))
        self.find_mode.set(s.get('find_mode', 'not_exists'))
        self.use_list_var.set(s.get('use_list', False))
        self._on_ext(); self._on_compare(); self._update_list_info()

    def _ok(self):
        mask = self.entry_mask.get().strip()
        if not mask:
            messagebox.showerror("Ошибка", "Укажите шаблон поиска"); return
        if self.compare_var.get():
            cmask = self.entry_compare.get().strip()
            if not cmask:
                messagebox.showerror("Ошибка", "Укажите шаблон для сравнения"); return
        else: cmask = ""
        ch = self.ext_choice.get()
        if ch == EXT_ALL: exts = None
        elif ch == EXT_IMAGES: exts = IMAGE_EXTS
        elif ch == EXT_JPG: exts = {'.jpg', '.jpeg'}
        elif ch == EXT_PNG: exts = {'.png'}
        else:
            raw = self.entry_custom_ext.get().lower()
            exts = set()
            for item in raw.replace(';', ',').split(','):
                item = item.strip()
                if item:
                    if not item.startswith('.'): item = '.' + item
                    exts.add(item)
            if not exts:
                messagebox.showerror("Ошибка", "Укажите расширения"); return
        self.result = {'mask': mask, 'compare_mask': cmask,
            'compare': self.compare_var.get(), 'compare_mode': self.compare_mode.get(),
            'find_mode': self.find_mode.get(), 'ext_choice': ch,
            'custom_ext': self.entry_custom_ext.get(), 'exts': exts,
            'use_list': self.use_list_var.get(),
            'list_articles': list(self.current.get('list_articles', [])),
            'list_match': self.current.get('list_match', 'exact')}
        self.destroy()

    def _cancel(self):
        self.result = None; self.destroy()


# ============================================================
#  ДИАЛОГ ОБЩИХ НАСТРОЕК
# ============================================================
class GeneralSettingsDialog(tk.Toplevel):
    def __init__(self, parent, current, app_ref):
        super().__init__(parent)
        self.app = app_ref
        self.title("Общие настройки")
        self.geometry("520x560"); self.resizable(False, False)
        self.transient(parent); self.grab_set()
        self.result = None
        self.current = dict(current)
        self._build(); self._load()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _tt(self, widget, text):
        ToolTip(widget, text, self.app)

    def _build(self):
        pad = {'padx': 10, 'pady': 6}
        tk.Label(self, text="Производительность", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, **pad)
        self.pool_var = tk.BooleanVar(value=False)
        chk_pool = tk.Checkbutton(self, text="Использовать пул потоков при обработке файлов",
                                  variable=self.pool_var)
        chk_pool.pack(anchor=tk.W, padx=20)
        self._tt(chk_pool, "Многопоточная обработка ускоряет работу\nпри большом количестве файлов.")
        tk.Label(self, text="Применяется только к файлам (не к папкам).",
                 fg="#888", font=("Segoe UI", 8)).pack(anchor=tk.W, padx=30)
        row_tc = tk.Frame(self); row_tc.pack(anchor=tk.W, padx=20, pady=(2, 0))
        tk.Label(row_tc, text="Потоков:").pack(side=tk.LEFT)
        self.spin_threads = tk.Spinbox(row_tc, from_=2, to=16, width=4)
        self.spin_threads.pack(side=tk.LEFT, padx=5)
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, padx=10, pady=10)
        tk.Label(self, text="Поведение", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, **pad)
        self.auto_open_var = tk.BooleanVar(value=False)
        chk_auto = tk.Checkbutton(self, text="Открывать последнюю папку при запуске",
                                  variable=self.auto_open_var)
        chk_auto.pack(anchor=tk.W, padx=20)
        self._tt(chk_auto, "Автоматически подставлять в поле «Папка»\nпоследний использованный путь.")
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, padx=10, pady=10)
        tk.Label(self, text="Интерфейс", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, **pad)
        self.tooltips_var = tk.BooleanVar(value=True)
        chk_tt = tk.Checkbutton(self, text="Показывать всплывающие подсказки",
                                variable=self.tooltips_var)
        chk_tt.pack(anchor=tk.W, padx=20)
        self._tt(chk_tt, "Показывать краткие пояснения при наведении\nмыши на элементы управления.")
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, padx=10, pady=10)
        tk.Label(self, text="Сохранение состояния", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, **pad)
        self.save_editor_var = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text="Сохранять сессию редактора",
                       variable=self.save_editor_var).pack(anchor=tk.W, padx=20)
        self.save_scan_var = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text="Сохранять настройки скана",
                       variable=self.save_scan_var).pack(anchor=tk.W, padx=20)
        btn = tk.Frame(self); btn.pack(fill=tk.X, side=tk.BOTTOM, pady=10, padx=10)
        tk.Button(btn, text="Отмена", command=self._cancel, width=12).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn, text="OK", command=self._ok, width=12, bg="#e1f5fe").pack(side=tk.RIGHT)

    def _load(self):
        s = self.current
        self.pool_var.set(s.get('use_thread_pool', False))
        self.spin_threads.delete(0, tk.END)
        self.spin_threads.insert(0, str(s.get('thread_count', 4)))
        self.auto_open_var.set(s.get('auto_open_last_path', False))
        self.tooltips_var.set(s.get('show_tooltips', True))
        self.save_editor_var.set(s.get('save_editor_session', True))
        self.save_scan_var.set(s.get('save_scan_settings', True))

    def _ok(self):
        try:
            tc = int(self.spin_threads.get() or 4)
            tc = max(2, min(16, tc))
        except ValueError: tc = 4
        self.result = {'use_thread_pool': self.pool_var.get(), 'thread_count': tc,
            'auto_open_last_path': self.auto_open_var.get(),
            'show_tooltips': self.tooltips_var.get(),
            'save_editor_session': self.save_editor_var.get(),
            'save_scan_settings': self.save_scan_var.get()}
        self.destroy()

    def _cancel(self):
        self.result = None; self.destroy()


# ============================================================
#  ДИАЛОГ ТАБЛИЦЫ РЕЗУЛЬТАТОВ СКАНА
# ============================================================
class ScanTableDialog(tk.Toplevel):
    def __init__(self, parent, records, checked_indices):
        super().__init__(parent)
        self.title('Отметьте папки для обработки')
        self.geometry('900x520'); self.transient(parent); self.grab_set()
        self.records = records
        self.checked = set(checked_indices)
        self.result = None
        self._build(); self._populate()
        self.protocol('WM_DELETE_WINDOW', self._cancel)
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f'+{x}+{y}')

    def _build(self):
        bar = tk.Frame(self); bar.pack(fill=tk.X, padx=8, pady=(8, 4))
        tk.Button(bar, text='Все', command=self._all, width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text='Ничего', command=self._none, width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text='Инвертировать', command=self._invert, width=14).pack(side=tk.LEFT, padx=2)
        cols = ('check', 'article', 'folder', 'files')
        self.tree = ttk.Treeview(self, columns=cols, show='headings', selectmode='none')
        self.tree.heading('check', text='☑')
        self.tree.heading('article', text='Артикул')
        self.tree.heading('folder', text='Папка')
        self.tree.heading('files', text='Файлов')
        self.tree.column('check', width=50, anchor=tk.CENTER)
        self.tree.column('article', width=200, anchor=tk.W)
        self.tree.column('folder', width=500, anchor=tk.W)
        self.tree.column('files', width=90, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.tree.bind('<Button-1>', self._on_click)
        self.lbl_count = tk.Label(self, text='', anchor=tk.W, fg='#333', font=('Segoe UI', 9))
        self.lbl_count.pack(fill=tk.X, padx=10)
        btn = tk.Frame(self); btn.pack(fill=tk.X, side=tk.BOTTOM, pady=10, padx=10)
        tk.Button(btn, text='Отмена', command=self._cancel, width=12).pack(side=tk.RIGHT, padx=4)
        tk.Button(btn, text='OK', command=self._ok, width=12, bg='#e1f5fe').pack(side=tk.RIGHT)

    def _populate(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for i, rec in enumerate(self.records):
            mark = '☑' if i in self.checked else '☐'
            self.tree.insert('', tk.END, iid=str(i),
                             values=(mark, rec['article'], rec['folder_rel'], len(rec['files'])))
        self._update_count()

    def _update_count(self):
        self.lbl_count.config(text=f'Отмечено: {len(self.checked)} из {len(self.records)}')

    def _on_click(self, event):
        if self.tree.identify('region', event.x, event.y) != 'cell': return
        if self.tree.identify_column(event.x) != '#1': return
        item = self.tree.identify_row(event.y)
        if not item: return
        idx = int(item)
        if idx in self.checked:
            self.checked.discard(idx); self.tree.set(item, 'check', '☐')
        else:
            self.checked.add(idx); self.tree.set(item, 'check', '☑')
        self._update_count()

    def _all(self):
        self.checked = set(range(len(self.records))); self._populate()
    def _none(self):
        self.checked = set(); self._populate()
    def _invert(self):
        self.checked = set(range(len(self.records))) - self.checked; self._populate()
    def _ok(self):
        self.result = set(self.checked); self.destroy()
    def _cancel(self):
        self.result = None; self.destroy()


# ============================================================
#  ОСНОВНОЕ ОКНО
# ============================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1000x900")
        self.root.minsize(900, 780)

        self.log_queue = queue.Queue()
        self.is_processing = False
        self.stop_requested = False
        self.anim_dots = 0
        self.anim_running = False
        self.anim_prefix = "Обработка"

        self.tooltips_enabled = True

        self.scan_settings = dict(DEFAULT_SCAN)
        self.general_settings = dict(DEFAULT_GENERAL)
        self.last_scan_results = []
        self.checked_scan_indices = set()

        self.editor_source = tk.StringVar(value="scan")
        self.editor_manual_path = tk.StringVar(value="")

        self.work_folders = tk.BooleanVar(value=True)
        self.folder_action = tk.StringVar(value="copy")
        self.folder_replace_from = tk.StringVar(value="")
        self.folder_replace_to = tk.StringVar(value="")
        self.folder_add_position = tk.StringVar(value="suffix")

        self.work_files = tk.BooleanVar(value=False)
        self.file_action = tk.StringVar(value="copy")
        self.file_replace_from = tk.StringVar(value="")
        self.file_replace_to = tk.StringVar(value="")
        self.file_add_position = tk.StringVar(value="suffix")
        self.file_filter_mode = tk.StringVar(value="all")

        self.editor_filter_exts = tk.StringVar(value="")
        self.editor_filter_mask = tk.StringVar(value="")
        self.editor_filter_size_min = tk.StringVar(value="")
        self.editor_filter_size_max = tk.StringVar(value="")
        self.editor_filter_date_create_from = tk.StringVar(value="")
        self.editor_filter_date_create_to = tk.StringVar(value="")
        self.editor_filter_date_mod_from = tk.StringVar(value="")
        self.editor_filter_date_mod_to = tk.StringVar(value="")

        self.editor_where = tk.StringVar(value="other")
        self.editor_dest_path = tk.StringVar(value="")

        self.editor_skip_existing = tk.BooleanVar(value=True)
        self.editor_update_date = tk.BooleanVar(value=False)
        self.editor_keep_structure = tk.BooleanVar(value=True)

        self.status_var = tk.StringVar(value="Готово.")

        self.load_settings()
        self.tooltips_enabled = self.general_settings.get('show_tooltips', True)
        self.build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.after(100, self.poll_queue)

    def _tt(self, widget, text):
        ToolTip(widget, text, self)

    # ---------- Настройки ----------
    def load_settings(self):
        data = {}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception as e:
                print(f"Ошибка загрузки настроек: {e}")
        if 'scan_settings' in data:
            s = dict(data['scan_settings'])
            ch = s.get('ext_choice', EXT_ALL)
            if ch == EXT_IMAGES: s['exts'] = IMAGE_EXTS
            elif ch == EXT_ALL: s['exts'] = None
            elif ch == EXT_JPG: s['exts'] = {'.jpg', '.jpeg'}
            elif ch == EXT_PNG: s['exts'] = {'.png'}
            elif ch == EXT_CUSTOM:
                raw = s.get('custom_ext', '').lower()
                exts = set()
                for it in raw.replace(';', ',').split(','):
                    it = it.strip()
                    if it:
                        if not it.startswith('.'): it = '.' + it
                        exts.add(it)
                s['exts'] = exts
            self.scan_settings.update(s)
        if 'general_settings' in data:
            self.general_settings.update(data['general_settings'])
        self.path_var = tk.StringVar(value=data.get('last_path', ''))
        if self.general_settings.get('save_editor_session', True):
            es = data.get('editor_session', {})
            self.editor_source.set(es.get('source', 'scan'))
            self.editor_manual_path.set(es.get('manual_path', ''))
            self.work_folders.set(es.get('work_folders', True))
            self.folder_action.set(es.get('folder_action', 'copy'))
            self.folder_replace_from.set(es.get('folder_replace_from', ''))
            self.folder_replace_to.set(es.get('folder_replace_to', ''))
            self.folder_add_position.set(es.get('folder_add_position', 'suffix'))
            self.work_files.set(es.get('work_files', False))
            self.file_action.set(es.get('file_action', 'copy'))
            self.file_replace_from.set(es.get('file_replace_from', ''))
            self.file_replace_to.set(es.get('file_replace_to', ''))
            self.file_add_position.set(es.get('file_add_position', 'suffix'))
            self.file_filter_mode.set(es.get('file_filter_mode', 'all'))
            self.editor_where.set(es.get('where', 'other'))
            self.editor_dest_path.set(es.get('dest_path', ''))
            self.editor_skip_existing.set(es.get('skip_existing', True))
            self.editor_update_date.set(es.get('update_date', False))
            self.editor_keep_structure.set(es.get('keep_structure', True))
            self.editor_filter_exts.set(es.get('filter_exts', ''))
            self.editor_filter_mask.set(es.get('filter_mask', ''))
            self.editor_filter_size_min.set(es.get('filter_size_min', ''))
            self.editor_filter_size_max.set(es.get('filter_size_max', ''))
            self.editor_filter_date_create_from.set(es.get('filter_dc_from', ''))
            self.editor_filter_date_create_to.set(es.get('filter_dc_to', ''))
            self.editor_filter_date_mod_from.set(es.get('filter_dm_from', ''))
            self.editor_filter_date_mod_to.set(es.get('filter_dm_to', ''))

    def save_settings(self):
        scan_save = {k: v for k, v in self.scan_settings.items() if k != 'exts'}
        data = {'last_path': self.path_var.get().strip() if hasattr(self, 'path_var') else '',
                'scan_settings': scan_save,
                'general_settings': dict(self.general_settings)}
        if self.general_settings.get('save_editor_session', True):
            data['editor_session'] = {
                'source': self.editor_source.get(), 'manual_path': self.editor_manual_path.get(),
                'work_folders': self.work_folders.get(), 'folder_action': self.folder_action.get(),
                'folder_replace_from': self.folder_replace_from.get(),
                'folder_replace_to': self.folder_replace_to.get(),
                'folder_add_position': self.folder_add_position.get(),
                'work_files': self.work_files.get(), 'file_action': self.file_action.get(),
                'file_replace_from': self.file_replace_from.get(),
                'file_replace_to': self.file_replace_to.get(),
                'file_add_position': self.file_add_position.get(),
                'file_filter_mode': self.file_filter_mode.get(),
                'where': self.editor_where.get(), 'dest_path': self.editor_dest_path.get(),
                'skip_existing': self.editor_skip_existing.get(),
                'update_date': self.editor_update_date.get(),
                'keep_structure': self.editor_keep_structure.get(),
                'filter_exts': self.editor_filter_exts.get(),
                'filter_mask': self.editor_filter_mask.get(),
                'filter_size_min': self.editor_filter_size_min.get(),
                'filter_size_max': self.editor_filter_size_max.get(),
                'filter_dc_from': self.editor_filter_date_create_from.get(),
                'filter_dc_to': self.editor_filter_date_create_to.get(),
                'filter_dm_from': self.editor_filter_date_mod_from.get(),
                'filter_dm_to': self.editor_filter_date_mod_to.get()}
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")

    def on_closing(self):
        self.save_settings(); self.root.destroy()

    # ---------- UI ----------
    def build_ui(self):
        # === Строка 1 (верхняя): справа кнопки Настройки, Справка, О программе ===
        top_buttons = tk.Frame(self.root)
        top_buttons.pack(fill=tk.X, padx=10, pady=(10, 2))

        # Спайсер слева
        tk.Frame(top_buttons).pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_gen = tk.Button(top_buttons, text="⚙ Настройки", command=self.open_general_settings)
        btn_gen.pack(side=tk.LEFT, padx=(0, 5))
        self._tt(btn_gen, "Общие настройки программы")
        btn_help = tk.Button(top_buttons, text="📖 Справка", command=self.show_help)
        btn_help.pack(side=tk.LEFT, padx=(0, 5))
        self._tt(btn_help, "Открыть подробную инструкцию")
        btn_about = tk.Button(top_buttons, text="? О программе", command=self.show_about)
        btn_about.pack(side=tk.LEFT)
        self._tt(btn_about, "Сведения о программе")

        # === Строка 2: Папка + поле пути + Обзор ===
        top_path = tk.Frame(self.root)
        top_path.pack(fill=tk.X, padx=10, pady=(2, 8))

        tk.Label(top_path, text="Папка:").pack(side=tk.LEFT)
        self.entry_path = tk.Entry(top_path, textvariable=self.path_var)
        self.entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self._bind_entry(self.entry_path)
        self._tt(self.entry_path, "Корневая папка для сканирования.\nМожно ввести путь вручную или выбрать через Обзор.")

        btn_browse = tk.Button(top_path, text="Обзор...", command=self.browse_folder)
        btn_browse.pack(side=tk.LEFT)
        self._tt(btn_browse, "Выбрать корневую папку в диалоге")

        # === Аудит ===
        box_scan = tk.LabelFrame(self.root, text=" Аудит (сканирование) ",
                                 font=("Segoe UI", 9, "bold"), fg="#e65100", padx=8, pady=6)
        box_scan.pack(fill=tk.X, padx=10, pady=(0, 5))
        row_scan = tk.Frame(box_scan); row_scan.pack(fill=tk.X)
        self.btn_scan = tk.Button(row_scan, text="🔍 Сканировать",
                                  command=self.start_scan, bg="#fff3e0", width=16)
        self.btn_scan.pack(side=tk.LEFT, padx=(0, 5))
        self._tt(self.btn_scan, "Начать сканирование выбранной папки")
        self.btn_stop_scan = tk.Button(row_scan, text="⏹ Остановить",
                                       command=self.request_stop, bg="#ffcdd2",
                                       width=16, state=tk.DISABLED)
        self.btn_stop_scan.pack(side=tk.LEFT, padx=(0, 5))
        self._tt(self.btn_stop_scan, "Прервать сканирование")
        self.btn_scan_settings = tk.Button(row_scan, text="⚙ Настройки скана",
                                           command=self.open_scan_settings, width=18)
        self.btn_scan_settings.pack(side=tk.LEFT, padx=(0, 5))
        self._tt(self.btn_scan_settings, "Шаблон поиска, расширения, сравнение,\nфильтр по списку артикулов")
        self.lbl_scan_info = tk.Label(box_scan, text="", anchor=tk.W,
                                      fg="#444", font=("Segoe UI", 8))
        self.lbl_scan_info.pack(fill=tk.X, pady=(6, 0))

        self._build_editor_block()
        self._build_notebook()

        status = tk.Frame(self.root, bg="#eceff1")
        status.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(status, textvariable=self.status_var, anchor=tk.W,
                 bg="#eceff1", fg="#333", font=("Segoe UI", 8)
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=3)
        tk.Label(status, text=f"{APP_NAME} v{APP_VERSION}",
                 anchor=tk.E, bg="#eceff1", fg="#888", font=("Segoe UI", 8)
                 ).pack(side=tk.RIGHT, padx=8, pady=3)

        for var in (self.folder_replace_from, self.folder_replace_to, self.folder_add_position,
                    self.file_replace_from, self.file_replace_to, self.file_add_position,
                    self.editor_dest_path, self.editor_update_date):
            var.trace_add("write", lambda *a: self._update_examples())

        self.update_scan_info()
        self.update_editor_state()

    def _build_editor_block(self):
        box = tk.LabelFrame(self.root, text=" Редактор ",
                            font=("Segoe UI", 9, "bold"), fg="#2e7d32",
                            padx=8, pady=6)
        box.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 5))

        row_src = tk.Frame(box); row_src.pack(fill=tk.X, pady=(0, 3))
        tk.Label(row_src, text="Источник:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        rb_scan = tk.Radiobutton(row_src, text="Отмеченное из таблицы скана",
                                 variable=self.editor_source, value="scan",
                                 command=self.update_editor_state)
        rb_scan.pack(side=tk.LEFT, padx=(0, 10))
        self._tt(rb_scan, "Работать с папками, отмеченными в таблице скана")
        self.btn_open_table = tk.Button(row_src, text="📋 Открыть таблицу",
                                        command=self.open_scan_table, width=22)
        self.btn_open_table.pack(side=tk.LEFT, padx=(0, 5))
        self._tt(self.btn_open_table, "Отметить папки для обработки")
        self.lbl_checked_info = tk.Label(row_src, text="", fg="#444", font=("Segoe UI", 8))
        self.lbl_checked_info.pack(side=tk.LEFT, padx=(0, 15))
        rb_manual = tk.Radiobutton(row_src, text="Папка вручную", variable=self.editor_source,
                                   value="manual", command=self.update_editor_state)
        rb_manual.pack(side=tk.LEFT)
        self._tt(rb_manual, "Работать с папкой без сканирования")

        row_manual = tk.Frame(box); row_manual.pack(fill=tk.X, pady=(0, 5))
        tk.Label(row_manual, text="", width=15).pack(side=tk.LEFT)
        self.entry_manual = tk.Entry(row_manual, textvariable=self.editor_manual_path)
        self.entry_manual.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self._bind_entry(self.entry_manual)
        self._tt(self.entry_manual, "Папка для обработки без скана")
        self.btn_manual = tk.Button(row_manual, text="Обзор...", command=self.browse_manual)
        self.btn_manual.pack(side=tk.LEFT)
        self._tt(self.btn_manual, "Выбрать папку вручную")

        # Работа с ПАПКАМИ
        box_f = tk.LabelFrame(box, text=" Работа с папками ",
                              font=("Segoe UI", 9, "bold"), fg="#1565c0", padx=8, pady=4)
        box_f.pack(fill=tk.X, pady=(0, 4))
        row_fh = tk.Frame(box_f); row_fh.pack(fill=tk.X)
        self.chk_work_folders = tk.Checkbutton(row_fh, text="Обрабатывать папки",
            variable=self.work_folders, command=self.update_editor_state,
            font=("Segoe UI", 9, "bold"))
        self.chk_work_folders.pack(side=tk.LEFT)
        self._tt(self.chk_work_folders, "Включить обработку самих папок")

        row_fa = tk.Frame(box_f); row_fa.pack(fill=tk.X, pady=(3, 2))
        tk.Label(row_fa, text="Действие:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        for val, label, tip in [
            ('copy', 'Копировать', "Создать копию папки"),
            ('move', 'Переместить', "Переместить папку"),
            ('delete', 'Удалить (в корзину)', "Удалить папку в Корзину Windows")]:
            rb = tk.Radiobutton(row_fa, text=label, variable=self.folder_action, value=val,
                                command=self.update_editor_state)
            rb.pack(side=tk.LEFT, padx=(0, 15))
            self._tt(rb, tip)

        row_fn = tk.Frame(box_f); row_fn.pack(fill=tk.X, pady=(2, 0))
        tk.Label(row_fn, text="Изменить имя:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(row_fn, text="Заменить").pack(side=tk.LEFT)
        self.entry_folder_from = tk.Entry(row_fn, textvariable=self.folder_replace_from, width=12)
        self.entry_folder_from.pack(side=tk.LEFT, padx=4)
        self._bind_entry(self.entry_folder_from)
        self._tt(self.entry_folder_from, "Что заменить в имени папки")
        tk.Label(row_fn, text="на").pack(side=tk.LEFT)
        self.entry_folder_to = tk.Entry(row_fn, textvariable=self.folder_replace_to, width=12)
        self.entry_folder_to.pack(side=tk.LEFT, padx=4)
        self._bind_entry(self.entry_folder_to)
        self._tt(self.entry_folder_to, "На что заменить (можно оставить пустым —\nтогда подстрока удаляется)")
        self.frame_folder_add_pos = tk.Frame(row_fn)
        rb_fp = tk.Radiobutton(self.frame_folder_add_pos, text="В начало",
                               variable=self.folder_add_position, value="prefix",
                               command=self._update_examples)
        rb_fp.pack(side=tk.LEFT, padx=(4, 4))
        self._tt(rb_fp, "Добавить подстроку в начало имени папки")
        rb_fs = tk.Radiobutton(self.frame_folder_add_pos, text="В конец",
                               variable=self.folder_add_position, value="suffix",
                               command=self._update_examples)
        rb_fs.pack(side=tk.LEFT)
        self._tt(rb_fs, "Добавить подстроку в конец имени папки")

        row_fi = tk.Frame(box_f); row_fi.pack(fill=tk.X, pady=(2, 0))
        tk.Label(row_fi, text="", width=15).pack(side=tk.LEFT)
        fi_inner = tk.Frame(row_fi); fi_inner.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.lbl_folder_example = tk.Label(fi_inner, text="", fg="#1565c0",
                                           font=("Consolas", 9), anchor=tk.W, justify=tk.LEFT)
        self.lbl_folder_example.pack(anchor=tk.W)
        self.lbl_folder_explain = tk.Label(fi_inner, text="", fg="#888",
                                           font=("Segoe UI", 8), anchor=tk.W, justify=tk.LEFT)
        self.lbl_folder_explain.pack(anchor=tk.W)

        # Работа с ФАЙЛАМИ
        box_ff = tk.LabelFrame(box, text=" Работа с файлами внутри ",
                               font=("Segoe UI", 9, "bold"), fg="#6a1b9a", padx=8, pady=4)
        box_ff.pack(fill=tk.X, pady=(4, 4))
        row_ffh = tk.Frame(box_ff); row_ffh.pack(fill=tk.X)
        self.chk_work_files = tk.Checkbutton(row_ffh, text="Обрабатывать файлы",
            variable=self.work_files, command=self.update_editor_state,
            font=("Segoe UI", 9, "bold"))
        self.chk_work_files.pack(side=tk.LEFT)
        self._tt(self.chk_work_files, "Включить обработку файлов внутри папок.\nБлокируется, если папки выбраны на удаление.")
        row_ffw = tk.Frame(box_ff); row_ffw.pack(fill=tk.X, pady=(3, 2))
        tk.Label(row_ffw, text="Что:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        self.rb_file_all = tk.Radiobutton(row_ffw, text="Все файлы",
                       variable=self.file_filter_mode, value="all",
                       command=self.update_editor_state)
        self.rb_file_all.pack(side=tk.LEFT, padx=(0, 10))
        self._tt(self.rb_file_all, "Обрабатывать все файлы внутри папок")
        self.rb_file_filtered = tk.Radiobutton(row_ffw, text="Файлы по фильтру",
                       variable=self.file_filter_mode, value="filtered",
                       command=self.update_editor_state)
        self.rb_file_filtered.pack(side=tk.LEFT)
        self._tt(self.rb_file_filtered, "Обрабатывать только файлы,\nподходящие под фильтр")
        self.frame_filters = tk.LabelFrame(box_ff, text=" Фильтр файлов ",
                                           font=("Segoe UI", 8), padx=6, pady=4)
        self._build_filters(self.frame_filters)
        row_ffa = tk.Frame(box_ff); row_ffa.pack(fill=tk.X, pady=(3, 2))
        tk.Label(row_ffa, text="Действие:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        for val, label, tip in [
            ('copy', 'Копировать', "Создать копию файла"),
            ('move', 'Переместить', "Переместить файл"),
            ('delete', 'Удалить (в корзину)', "Удалить файл в Корзину Windows")]:
            rb = tk.Radiobutton(row_ffa, text=label, variable=self.file_action, value=val,
                                command=self.update_editor_state)
            rb.pack(side=tk.LEFT, padx=(0, 15))
            self._tt(rb, tip)
        row_ffn = tk.Frame(box_ff); row_ffn.pack(fill=tk.X, pady=(2, 0))
        tk.Label(row_ffn, text="Изменить имя:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(row_ffn, text="Заменить").pack(side=tk.LEFT)
        self.entry_file_from = tk.Entry(row_ffn, textvariable=self.file_replace_from, width=12)
        self.entry_file_from.pack(side=tk.LEFT, padx=4)
        self._bind_entry(self.entry_file_from)
        self._tt(self.entry_file_from, "Что заменить в имени файла")
        tk.Label(row_ffn, text="на").pack(side=tk.LEFT)
        self.entry_file_to = tk.Entry(row_ffn, textvariable=self.file_replace_to, width=12)
        self.entry_file_to.pack(side=tk.LEFT, padx=4)
        self._bind_entry(self.entry_file_to)
        self._tt(self.entry_file_to, "На что заменить (можно оставить пустым —\nтогда подстрока удаляется)")
        self.frame_file_add_pos = tk.Frame(row_ffn)
        rb_fp2 = tk.Radiobutton(self.frame_file_add_pos, text="В начало",
                                variable=self.file_add_position, value="prefix",
                                command=self._update_examples)
        rb_fp2.pack(side=tk.LEFT, padx=(4, 4))
        self._tt(rb_fp2, "Добавить подстроку в начало имени файла")
        rb_fs2 = tk.Radiobutton(self.frame_file_add_pos, text="В конец",
                                variable=self.file_add_position, value="suffix",
                                command=self._update_examples)
        rb_fs2.pack(side=tk.LEFT)
        self._tt(rb_fs2, "Добавить подстроку в конец имени файла")
        row_ffi = tk.Frame(box_ff); row_ffi.pack(fill=tk.X, pady=(2, 0))
        tk.Label(row_ffi, text="", width=15).pack(side=tk.LEFT)
        ffi_inner = tk.Frame(row_ffi); ffi_inner.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.lbl_file_example = tk.Label(ffi_inner, text="", fg="#6a1b9a",
                                         font=("Consolas", 9), anchor=tk.W, justify=tk.LEFT)
        self.lbl_file_example.pack(anchor=tk.W)
        self.lbl_file_explain = tk.Label(ffi_inner, text="", fg="#888",
                                         font=("Segoe UI", 8), anchor=tk.W, justify=tk.LEFT)
        self.lbl_file_explain.pack(anchor=tk.W)

        # Куда
        box_where = tk.LabelFrame(box, text=" Куда ", font=("Segoe UI", 9, "bold"),
                                  fg="#333", padx=8, pady=4)
        box_where.pack(fill=tk.X, pady=(4, 4))
        row_wh = tk.Frame(box_where); row_wh.pack(fill=tk.X)
        rb_other = tk.Radiobutton(row_wh, text="В другую папку", variable=self.editor_where,
                                  value="other", command=self.update_editor_state)
        rb_other.pack(side=tk.LEFT, padx=(0, 10))
        self._tt(rb_other, "Копировать/переместить в другую папку")
        self.entry_dst = tk.Entry(row_wh, textvariable=self.editor_dest_path, width=30)
        self.entry_dst.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self._bind_entry(self.entry_dst)
        self._tt(self.entry_dst, "Целевая папка назначения")
        self.btn_dst = tk.Button(row_wh, text="Обзор...", command=self.browse_dst)
        self.btn_dst.pack(side=tk.LEFT, padx=(0, 10))
        self._tt(self.btn_dst, "Выбрать целевую папку")
        rb_same = tk.Radiobutton(row_wh, text="В ту же папку", variable=self.editor_where,
                                 value="same", command=self.update_editor_state)
        rb_same.pack(side=tk.LEFT)
        self._tt(rb_same, "Работать в исходной папке (только переименование)")

        # Опции
        row_opt = tk.Frame(box); row_opt.pack(fill=tk.X, pady=(4, 0))
        tk.Label(row_opt, text="Опции:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        self.chk_skip = tk.Checkbutton(row_opt, text="Пропускать существующие",
                                       variable=self.editor_skip_existing)
        self.chk_skip.pack(side=tk.LEFT, padx=(0, 15))
        self._tt(self.chk_skip, "Не перезаписывать, если объект\nс таким именем уже есть")
        self.chk_date = tk.Checkbutton(row_opt, text="Обновлять дату файлов",
                                       variable=self.editor_update_date)
        self.chk_date.pack(side=tk.LEFT, padx=(0, 15))
        self._tt(self.chk_date, "После операции проставить текущую дату\nна файлы")
        self.chk_struct = tk.Checkbutton(row_opt, text="Сохранять структуру подпапок",
                                         variable=self.editor_keep_structure)
        self.chk_struct.pack(side=tk.LEFT)
        self._tt(self.chk_struct, "При копировании в другую папку\nсохранять вложенность папок")

        # Кнопки
        row_run = tk.Frame(box); row_run.pack(fill=tk.X, pady=(6, 0))
        self.btn_apply = tk.Button(row_run, text="▶ Применить",
                                   command=self.start_editor, bg="#c8e6c9", width=18)
        self.btn_apply.pack(side=tk.LEFT, padx=(0, 6))
        self._tt(self.btn_apply, "Запустить операцию редактора")
        self.btn_stop_editor = tk.Button(row_run, text="⏹ Остановить",
                                         command=self.request_stop, bg="#ffcdd2",
                                         width=16, state=tk.DISABLED)
        self.btn_stop_editor.pack(side=tk.LEFT, padx=(0, 10))
        self._tt(self.btn_stop_editor, "Прервать операцию")
        self.btn_undo = tk.Button(row_run, text="↶ Отменить последнюю операцию",
                                  command=self.undo_last_operation, width=30)
        self.btn_undo.pack(side=tk.LEFT)
        self._tt(self.btn_undo, "Откатить последнюю операцию по журналу.\nУдаление в Корзину не откатывается.")
        self.lbl_journal = tk.Label(row_run, text="", fg="#666", font=("Segoe UI", 8))
        self.lbl_journal.pack(side=tk.LEFT, padx=10)

        self.progress = ttk.Progressbar(box, mode='determinate', maximum=100)
        self.progress.pack(fill=tk.X, pady=(6, 0))
        self.lbl_progress = tk.Label(box, text="", anchor=tk.W,
                                     fg="#444", font=("Segoe UI", 8))
        self.lbl_progress.pack(fill=tk.X)
        self.lbl_editor_info = tk.Label(box, text="Записей: 0", anchor=tk.E,
                                        fg="#444", font=("Segoe UI", 8))
        self.lbl_editor_info.pack(fill=tk.X)

        self.update_journal_label()

    def _build_filters(self, parent):
        r1 = tk.Frame(parent); r1.pack(fill=tk.X, pady=1)
        tk.Label(r1, text="Расширения:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        e = tk.Entry(r1, textvariable=self.editor_filter_exts, width=18); e.pack(side=tk.LEFT, padx=4)
        self._bind_entry(e); self._tt(e, "Список расширений через запятую")
        tk.Label(r1, text="Маска:").pack(side=tk.LEFT, padx=(10, 0))
        e = tk.Entry(r1, textvariable=self.editor_filter_mask, width=18); e.pack(side=tk.LEFT, padx=4)
        self._bind_entry(e); self._tt(e, "Маска имени файла (например, *_01)")
        r2 = tk.Frame(parent); r2.pack(fill=tk.X, pady=1)
        tk.Label(r2, text="Размер КБ:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(r2, text="от").pack(side=tk.LEFT)
        e = tk.Entry(r2, textvariable=self.editor_filter_size_min, width=10); e.pack(side=tk.LEFT, padx=3)
        self._bind_entry(e)
        tk.Label(r2, text="до").pack(side=tk.LEFT)
        e = tk.Entry(r2, textvariable=self.editor_filter_size_max, width=10); e.pack(side=tk.LEFT, padx=3)
        self._bind_entry(e)
        r3 = tk.Frame(parent); r3.pack(fill=tk.X, pady=1)
        tk.Label(r3, text="Дата создания:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(r3, text="от").pack(side=tk.LEFT)
        e = tk.Entry(r3, textvariable=self.editor_filter_date_create_from, width=13); e.pack(side=tk.LEFT, padx=3)
        self._bind_entry(e)
        tk.Label(r3, text="до").pack(side=tk.LEFT)
        e = tk.Entry(r3, textvariable=self.editor_filter_date_create_to, width=13); e.pack(side=tk.LEFT, padx=3)
        self._bind_entry(e)
        r4 = tk.Frame(parent); r4.pack(fill=tk.X, pady=1)
        tk.Label(r4, text="Дата изменения:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(r4, text="от").pack(side=tk.LEFT)
        e = tk.Entry(r4, textvariable=self.editor_filter_date_mod_from, width=13); e.pack(side=tk.LEFT, padx=3)
        self._bind_entry(e)
        tk.Label(r4, text="до").pack(side=tk.LEFT)
        e = tk.Entry(r4, textvariable=self.editor_filter_date_mod_to, width=13); e.pack(side=tk.LEFT, padx=3)
        self._bind_entry(e)

    def _build_notebook(self):
        nb_frame = tk.Frame(self.root); nb_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 5))
        self.notebook = ttk.Notebook(nb_frame); self.notebook.pack(fill=tk.BOTH, expand=True)
        tab1 = tk.Frame(self.notebook); self.notebook.add(tab1, text=" Сканирование — отчёт ")
        bar1 = tk.Frame(tab1); bar1.pack(fill=tk.X, pady=(4, 0))
        tk.Button(bar1, text="Очистить", command=lambda: self.clear_log(self.log_scan)).pack(side=tk.LEFT, padx=2)
        tk.Button(bar1, text="💾 Сохранить отчёт",
                  command=lambda: self.save_log(self.log_scan, "scan_report.txt")).pack(side=tk.LEFT, padx=2)
        self.log_scan = scrolledtext.ScrolledText(tab1, wrap=tk.NONE, state=tk.DISABLED,
                                                  bg="#fafafa", font=("Consolas", 10))
        self.log_scan.pack(fill=tk.BOTH, expand=True, pady=4)
        tab2 = tk.Frame(self.notebook); self.notebook.add(tab2, text=" Редактор — отчёт ")
        bar2 = tk.Frame(tab2); bar2.pack(fill=tk.X, pady=(4, 0))
        tk.Button(bar2, text="Очистить", command=lambda: self.clear_log(self.log_edit)).pack(side=tk.LEFT, padx=2)
        tk.Button(bar2, text="💾 Сохранить отчёт",
                  command=lambda: self.save_log(self.log_edit, "editor_report.txt")).pack(side=tk.LEFT, padx=2)
        self.log_edit = scrolledtext.ScrolledText(tab2, wrap=tk.NONE, state=tk.DISABLED,
                                                  bg="#f7fbe7", font=("Consolas", 10))
        self.log_edit.pack(fill=tk.BOTH, expand=True, pady=4)

    # ---------- Логи ----------
    def clear_log(self, widget):
        widget.config(state=tk.NORMAL); widget.delete(1.0, tk.END); widget.config(state=tk.DISABLED)

    def save_log(self, widget, default_name):
        fp = filedialog.asksaveasfilename(defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
            initialfile=default_name)
        if fp:
            try:
                content = widget.get(1.0, tk.END)
                with open(fp, "w", encoding="utf-8") as f: f.write(content)
                messagebox.showinfo("Успех", f"Сохранено:\n{fp}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось: {e}")

    def log_message(self, which, msg):
        widget = self.log_scan if which == "scan" else self.log_edit
        widget.config(state=tk.NORMAL)
        widget.insert(tk.END, msg + "\n"); widget.see(tk.END); widget.config(state=tk.DISABLED)

    # ---------- Горячие клавиши ----------
    def _on_ctrl_key(self, event):
        w, c = event.widget, event.keycode
        if c == VK_V: return self._paste(w)
        if c == VK_C: return self._copy(w)
        if c == VK_X: return self._cut(w)
        if c == VK_A: return self._select_all(w)
        return None

    def _on_shift_key(self, event):
        w, c = event.widget, event.keycode
        if c == VK_INSERT: return self._paste(w)
        if c == VK_DELETE: return self._cut(w)
        return None

    def _paste(self, widget, event=None):
        try:
            text = self.root.clipboard_get()
            if not isinstance(text, str): return "break"
            text = text.replace("\r", "").replace("\n", "").strip()
            try: widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError: pass
            widget.insert(tk.INSERT, text)
        except Exception: pass
        return "break"

    def _copy(self, widget, event=None):
        try:
            sel = widget.selection_get()
            self.root.clipboard_clear(); self.root.clipboard_append(sel)
        except Exception: pass
        return "break"

    def _cut(self, widget, event=None):
        try:
            sel = widget.selection_get()
            self.root.clipboard_clear(); self.root.clipboard_append(sel)
            widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except Exception: pass
        return "break"

    def _select_all(self, widget, event=None):
        try: widget.select_range(0, tk.END); widget.icursor(tk.END)
        except Exception: pass
        return "break"

    def _context_menu(self, event):
        w = event.widget
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="Вырезать", command=lambda: self._cut(w))
        m.add_command(label="Копировать", command=lambda: self._copy(w))
        m.add_command(label="Вставить", command=lambda: self._paste(w))
        m.add_separator()
        m.add_command(label="Выделить всё", command=lambda: self._select_all(w))
        try: m.tk_popup(event.x_root, event.y_root)
        finally: m.grab_release()
        return "break"

    def _bind_entry(self, entry):
        entry.bind("<Control-KeyPress>", self._on_ctrl_key)
        entry.bind("<Shift-KeyPress>", self._on_shift_key)
        entry.bind("<Button-3>", self._context_menu)

    # ---------- Диалоги ----------
    def show_about(self):
        AboutDialog(self.root)

    def show_help(self):
        HelpDialog(self.root)

    def browse_folder(self):
        f = filedialog.askdirectory()
        if f: self.path_var.set(f); self.save_settings()
    def browse_manual(self):
        f = filedialog.askdirectory()
        if f: self.editor_manual_path.set(f)
    def browse_dst(self):
        f = filedialog.askdirectory()
        if f: self.editor_dest_path.set(f)

    def open_scan_settings(self):
        dlg = ScanSettingsDialog(self.root, self.scan_settings, self)
        self.root.wait_window(dlg)
        if dlg.result:
            self.scan_settings.update(dlg.result)
            self.update_scan_info(); self.save_settings(); self._update_examples()

    def open_general_settings(self):
        dlg = GeneralSettingsDialog(self.root, self.general_settings, self)
        self.root.wait_window(dlg)
        if dlg.result:
            self.general_settings.update(dlg.result)
            self.tooltips_enabled = self.general_settings.get('show_tooltips', True)
            self.save_settings()

    def open_scan_table(self):
        if not self.last_scan_results:
            messagebox.showinfo("Нет данных", "Результат скана пуст. Сначала выполните сканирование.")
            return
        dlg = ScanTableDialog(self.root, self.last_scan_results, self.checked_scan_indices)
        self.root.wait_window(dlg)
        if dlg.result is not None:
            self.checked_scan_indices = set(dlg.result); self.update_editor_state()

    # ---------- Информация ----------
    def update_scan_info(self):
        s = self.scan_settings
        exts = s.get('exts')
        ext_str = "все файлы" if exts is None else ", ".join(sorted(exts))
        if s.get('compare'):
            mode = "нет парного" if s.get('compare_mode') == 'missing_pair' else "нет основного"
            txt = f"Скан: сравнение | {s['mask']} ↔ {s['compare_mask']} | {mode} | {ext_str}"
        else:
            mode = "ЕСТЬ" if s.get('find_mode') == 'exists' else "НЕТ"
            txt = f"Скан: поиск | {s['mask']} | где {mode} | {ext_str}"
        if s.get('use_list'):
            n = len(s.get('list_articles', []))
            mm = 'точн' if s.get('list_match') == 'exact' else 'содерж'
            txt += f" | список: {n} ({mm})"
        self.lbl_scan_info.config(text=txt)

    def _get_sample_article_from_list(self):
        for a in self.scan_settings.get('list_articles', []):
            if a.isdigit(): return a
        return "art001"

    def _get_sample_ext(self):
        ch = self.scan_settings.get('ext_choice', EXT_ALL)
        if ch == EXT_JPG: return ".jpg"
        elif ch == EXT_PNG: return ".png"
        elif ch == EXT_CUSTOM:
            raw = self.scan_settings.get('custom_ext', '')
            first = raw.split(',')[0].strip()
            if first:
                if not first.startswith('.'): first = '.' + first
                return first
            return ".jpg"
        return ".jpg"

    def _get_scan_sample_filename(self):
        if not self.last_scan_results: return None
        idx = min(self.checked_scan_indices) if self.checked_scan_indices else 0
        if not (0 <= idx < len(self.last_scan_results)): return None
        rec = self.last_scan_results[idx]
        for fname in rec.get('files', []):
            base, ext = os.path.splitext(fname)
            if ext: return fname
        return None

    def _get_scan_sample_folder(self):
        if not self.last_scan_results: return None
        idx = min(self.checked_scan_indices) if self.checked_scan_indices else 0
        if not (0 <= idx < len(self.last_scan_results)): return None
        return os.path.basename(self.last_scan_results[idx]['folder_abs'])

    def _get_sample_folder_name(self):
        name = self._get_scan_sample_folder()
        if name: return name
        return self._get_sample_article_from_list()

    def _get_sample_file_name(self):
        name = self._get_scan_sample_filename()
        if name: return name
        return f"{self._get_sample_article_from_list()}{self._get_sample_ext()}"

    def _apply_rename_all(self, base, ren_from, ren_to, add_position):
        if ren_from:
            return base.replace(ren_from, ren_to)
        elif ren_to:
            if add_position == "prefix": return ren_to + base
            return base + ren_to
        return base

    def _update_folder_example(self):
        if not hasattr(self, 'lbl_folder_example'): return
        if not self.work_folders.get():
            self.lbl_folder_example.config(text="")
            self.lbl_folder_explain.config(text="")
            return
        action = self.folder_action.get()
        where = self.editor_where.get()
        ren_from = self.folder_replace_from.get()
        ren_to = self.folder_replace_to.get()
        add_pos = self.folder_add_position.get()
        if hasattr(self, 'frame_folder_add_pos'):
            if not ren_from and ren_to:
                self.frame_folder_add_pos.pack(side=tk.LEFT, padx=(4, 0))
            else:
                self.frame_folder_add_pos.pack_forget()
        base = self._get_sample_folder_name()
        if action == "delete":
            self.lbl_folder_example.config(text=f"Пример: {base} → (в корзину)")
            self.lbl_folder_explain.config(text="Папка будет удалена вместе с содержимым.")
            return
        new_base = self._apply_rename_all(base, ren_from, ren_to, add_pos)
        changed = (new_base != base)
        if where == "other" and changed:
            dest_dir = self.editor_dest_path.get().strip()
            short = dest_dir.rstrip("\\/").replace("/", "\\").split("\\")[-1] if dest_dir else "dest"
            example = f"Пример: {base} → ...\\{short}\\{new_base}"
        else:
            example = f"Пример: {base} → {new_base}"
        if not ren_from and not ren_to:
            example += "   (имя папки не изменится)"
        self.lbl_folder_example.config(text=example)
        if action == "copy" and where == "same": act = "Создастся копия папки. Оригинал останется."
        elif action == "move" and where == "same": act = "Папка переименуется на месте. Оригинал исчезнет."
        elif action == "copy" and where == "other": act = "Копия папки появится в целевой папке. Оригинал останется."
        elif action == "move" and where == "other": act = "Папка переедет в целевую папку. Оригинал исчезнет."
        else: act = ""
        if not ren_from and not ren_to:
            name = "Имя папки не изменится."
        elif ren_from and not ren_to:
            name = f"Все «{ren_from}» удалятся из имени."
        elif not ren_from and ren_to:
            pos = "в начало" if add_pos == "prefix" else "в конец"
            name = f"«{ren_to}» добавится {pos} имени."
        else:
            name = f"Все «{ren_from}» заменятся на «{ren_to}»."
        if not changed and ren_from:
            name = f"«{ren_from}» не найдено в имени папки."
        if where == "same" and not changed and action in ("copy", "move"):
            act = "Имя не изменится. Копия в ту же папку невозможна — папка уже на месте."
            name = ""
        self.lbl_folder_explain.config(text=(act + " " + name).strip())

    def _update_file_example(self):
        if not hasattr(self, 'lbl_file_example'): return
        if not self.work_files.get():
            self.lbl_file_example.config(text="")
            self.lbl_file_explain.config(text="")
            return
        action = self.file_action.get()
        where = self.editor_where.get()
        ren_from = self.file_replace_from.get()
        ren_to = self.file_replace_to.get()
        add_pos = self.file_add_position.get()
        upd_date = self.editor_update_date.get()
        if hasattr(self, 'frame_file_add_pos'):
            if not ren_from and ren_to:
                self.frame_file_add_pos.pack(side=tk.LEFT, padx=(4, 0))
            else:
                self.frame_file_add_pos.pack_forget()
        src_full = self._get_sample_file_name()
        base, ext = os.path.splitext(src_full)
        if action == "delete":
            self.lbl_file_example.config(text=f"Пример: {src_full} → (в корзину)")
            self.lbl_file_explain.config(text="Файл будет удалён в Корзину Windows.")
            return
        new_base = self._apply_rename_all(base, ren_from, ren_to, add_pos)
        dst_full = new_base + ext
        changed = (new_base != base)
        if where == "other" and changed:
            dest_dir = self.editor_dest_path.get().strip()
            short = dest_dir.rstrip("\\/").replace("/", "\\").split("\\")[-1] if dest_dir else "dest"
            example = f"Пример: {src_full} → ...\\{short}\\{dst_full}"
        else:
            example = f"Пример: {src_full} → {dst_full}"
        if not ren_from and not ren_to:
            example += "   (имя файла не изменится)"
        self.lbl_file_example.config(text=example)
        if action == "copy" and where == "same": act = "Создастся копия файла. Оригинал останется."
        elif action == "move" and where == "same": act = "Файл переименуется на месте. Оригинал исчезнет."
        elif action == "copy" and where == "other": act = "Копия появится в целевой папке. Оригинал останется."
        elif action == "move" and where == "other": act = "Файл переедет в целевую папку. Оригинал исчезнет."
        else: act = ""
        if not ren_from and not ren_to:
            name = "Имя файла не изменится."
        elif ren_from and not ren_to:
            name = f"Все «{ren_from}» удалятся из имени."
        elif not ren_from and ren_to:
            pos = "в начало" if add_pos == "prefix" else "в конец"
            name = f"«{ren_to}» добавится {pos} имени."
        else:
            name = f"Все «{ren_from}» заменятся на «{ren_to}»."
        if not changed and ren_from:
            name = f"«{ren_from}» не найдено в имени файла."
        if where == "same" and not changed and action in ("copy", "move"):
            if upd_date:
                act = "Имя не изменится. Копия не создастся — файл уже на месте, но дата файла будет обновлена."
            else:
                act = "Имя не изменится. Копия в ту же папку невозможна — файл уже на месте."
            name = ""
        self.lbl_file_explain.config(text=(act + " " + name).strip())

    def _update_examples(self):
        self._update_folder_example()
        self._update_file_example()

    def update_editor_state(self):
        folder_ok = self.work_folders.get()
        folder_action = self.folder_action.get()
        st_f = tk.NORMAL if folder_ok else tk.DISABLED
        for w in [self.entry_folder_from, self.entry_folder_to]:
            w.config(state=st_f)

        files_blocked_by_delete = (folder_ok and folder_action == 'delete')
        if files_blocked_by_delete: self.work_files.set(False)
        file_ok = self.work_files.get() and not files_blocked_by_delete
        st_ff = tk.NORMAL if file_ok else tk.DISABLED
        for w in [self.entry_file_from, self.entry_file_to]:
            w.config(state=st_ff)
        self.rb_file_all.config(state=st_ff)
        self.rb_file_filtered.config(state=st_ff)
        if files_blocked_by_delete:
            self.chk_work_files.config(state=tk.DISABLED)
        else:
            self.chk_work_files.config(state=tk.NORMAL)

        if file_ok and self.file_filter_mode.get() == "filtered":
            self.frame_filters.pack(fill=tk.X, pady=(3, 3))
        else:
            self.frame_filters.pack_forget()

        manual = self.editor_source.get() == "manual"
        st = tk.NORMAL if manual else tk.DISABLED
        self.entry_manual.config(state=st)
        self.btn_manual.config(state=st)
        if manual:
            self.btn_open_table.config(state=tk.DISABLED)
            self.lbl_checked_info.config(text="")
        else:
            n_total = len(self.last_scan_results)
            n_chk = len(self.checked_scan_indices)
            if n_total == 0:
                self.btn_open_table.config(state=tk.DISABLED, text="📋 Таблица (пусто)")
                self.lbl_checked_info.config(text="")
            else:
                if n_chk == 0: note = "не отмечено"
                elif n_chk == n_total: note = f"выбрано всё ({n_total})"
                else: note = f"выбрано {n_chk} из {n_total}"
                self.btn_open_table.config(state=tk.NORMAL, text=f"📋 Открыть таблицу ({n_total})")
                self.lbl_checked_info.config(text=f"({note})")

        other = self.editor_where.get() == "other"
        st = tk.NORMAL if other else tk.DISABLED
        self.entry_dst.config(state=st)
        self.btn_dst.config(state=st)

        block_reason = None
        if not self.work_folders.get() and not self.work_files.get():
            block_reason = "Включите хотя бы одно действие — папки или файлы"
        if self.editor_source.get() == "scan" and self.last_scan_results and not self.checked_scan_indices:
            block_reason = "Отметьте хотя бы одну папку в таблице"
        if block_reason is None and self.editor_where.get() == "same":
            if self.work_folders.get() and self.folder_action.get() in ('copy', 'move'):
                if not self.folder_replace_from.get() and not self.folder_replace_to.get():
                    block_reason = "Папки: имя не меняется — копия/перемещение в ту же папку невозможны"
            if block_reason is None and self.work_files.get() and self.file_action.get() in ('copy', 'move'):
                if not self.file_replace_from.get() and not self.file_replace_to.get():
                    block_reason = "Файлы: имя не меняется — копия/перемещение в ту же папку невозможны"

        if block_reason:
            self.btn_apply.config(state=tk.DISABLED)
            self.lbl_editor_info.config(text=f"⚠ {block_reason}")
        else:
            self.btn_apply.config(state=tk.NORMAL)
            self.lbl_editor_info.config(text=f"Записей: {len(self.last_scan_results)}")

        self._update_examples()

    def update_journal_label(self):
        j = load_journal()
        if j:
            n_files = len(j.get('files', []))
            n_folders = len(j.get('folders', []))
            n_deleted = len(j.get('deleted', []))
            n_total = n_files + n_folders + n_deleted
            if n_total > 0:
                action = j.get('action', '?')
                self.lbl_journal.config(text=f"Журнал: {action} ({n_total} эл.)")
                self.btn_undo.config(state=tk.NORMAL)
                return
        self.lbl_journal.config(text="Журнал пуст")
        self.btn_undo.config(state=tk.DISABLED)

    # ---------- Анимация ----------
    def start_animation(self, prefix="Обработка"):
        self.anim_running = True; self.anim_dots = 0; self.anim_prefix = prefix; self._animate()
    def stop_animation(self): self.anim_running = False
    def _animate(self):
        if not self.anim_running: return
        self.anim_dots = (self.anim_dots + 1) % 4
        if not self.stop_requested:
            self.status_var.set(f"⏳ {self.anim_prefix}{'.' * self.anim_dots}")
        self.root.after(400, self._animate)
    def request_stop(self):
        if self.is_processing:
            self.stop_requested = True; self.status_var.set("⏹ Останавливаю…")

    # ---------- Сканирование ----------
    def start_scan(self):
        path = self.path_var.get().strip().strip('"').strip("'")
        if not path or not os.path.exists(path):
            messagebox.showerror("Ошибка", "Выберите существующую папку"); return
        if self.is_processing: return
        self.is_processing = True; self.stop_requested = False
        self.btn_scan.config(state=tk.DISABLED)
        self.btn_scan_settings.config(state=tk.DISABLED)
        self.btn_apply.config(state=tk.DISABLED)
        self.btn_stop_scan.config(state=tk.NORMAL)
        self.btn_stop_editor.config(state=tk.DISABLED)
        self.start_animation("Сканирование")
        self.clear_log(self.log_scan)
        self.progress.config(mode='indeterminate'); self.progress.start(10)
        self.lbl_progress.config(text="Сканирование...")
        threading.Thread(target=self._scan_worker,
                         args=(path, dict(self.scan_settings)), daemon=True).start()

    def _scan_worker(self, root_dir, st):
        t0 = datetime.datetime.now()
        mask = st['mask']
        cmask = st.get('compare_mask', '')
        compare = st.get('compare', False)
        cmode = st.get('compare_mode', 'missing_pair')
        fmode = st.get('find_mode', 'exists')
        exts = st.get('exts')
        use_list = st.get('use_list', False)
        list_articles = st.get('list_articles', [])
        list_match = st.get('list_match', 'exact')
        list_set_lower = {a.lower() for a in list_articles} if list_articles else set()
        list_lower = [a.lower() for a in list_articles] if list_match == 'contains' else []

        def folder_matches_list(folder_name):
            if not use_list or not list_articles: return True
            n = folder_name.lower()
            if list_match == 'exact': return n in list_set_lower
            return any(a in n for a in list_lower)

        match_main = make_matcher(mask)
        match_cmp = make_matcher(cmask) if compare else None
        suf_main = make_suffix(mask); suf_cmp = make_suffix(cmask) if compare else ''
        total_folders = 0; total_main = 0; total_cmp = 0
        found = []; missing = []; results = []; skipped_by_list = 0
        interrupted = False
        stack = [(root_dir, "")]
        while stack:
            if self.stop_requested: interrupted = True; break
            abs_path, rel_path = stack.pop()
            try:
                with os.scandir(abs_path) as it:
                    dirs, main_a, cmp_a, all_files = [], set(), set(), []
                    for entry in it:
                        try:
                            if entry.is_dir(follow_symlinks=False): dirs.append(entry.name); continue
                            if not entry.is_file(follow_symlinks=False): continue
                        except OSError: continue
                        fname = entry.name; all_files.append(fname)
                        dot = fname.rfind('.')
                        if dot > 0: np, ep = fname[:dot], fname[dot:].lower()
                        else: np, ep = fname, ''
                        if exts is not None and ep not in exts: continue
                        if match_main(np, ep):
                            total_main += 1; main_a.add(extract_article(np, suf_main))
                        if compare and match_cmp(np, ep):
                            total_cmp += 1; cmp_a.add(extract_article(np, suf_cmp))
            except (PermissionError, FileNotFoundError, OSError): continue
            total_folders += 1
            folder_name = os.path.basename(abs_path)
            if not folder_matches_list(folder_name):
                skipped_by_list += 1
                for d in dirs:
                    sub_abs = os.path.join(abs_path, d)
                    sub_rel = os.path.join(rel_path, d) if rel_path else d
                    stack.append((sub_abs, sub_rel))
                continue
            in_result = False; arts_here = []
            if not compare:
                if fmode == 'exists' and main_a:
                    in_result = True; arts_here = sorted(main_a); found.append((rel_path, arts_here))
                elif fmode == 'not_exists' and not main_a:
                    in_result = True; missing.append((rel_path, []))
            else:
                if cmode == 'missing_pair':
                    diff = main_a - cmp_a
                    if diff: in_result = True; arts_here = sorted(diff); missing.append((rel_path, arts_here))
                else:
                    diff = cmp_a - main_a
                    if diff: in_result = True; arts_here = sorted(diff); missing.append((rel_path, arts_here))
            if in_result:
                results.append({'article': arts_here[0] if arts_here else os.path.basename(abs_path),
                                'folder_abs': abs_path,
                                'folder_rel': rel_path if rel_path else ".",
                                'files': all_files})
            for d in dirs:
                sub_abs = os.path.join(abs_path, d)
                sub_rel = os.path.join(rel_path, d) if rel_path else d
                stack.append((sub_abs, sub_rel))
        dur = datetime.datetime.now() - t0
        self.last_scan_results = results
        self.checked_scan_indices = set(range(len(results)))
        q = self.log_queue
        if interrupted:
            q.put(("scan", "⏹ ОПЕРАЦИЯ ПРЕРВАНА ПОЛЬЗОВАТЕЛЕМ"))
            q.put(("scan", f"Найдено до остановки: {len(results)} записей"))
            q.put(("DONE_SCAN", f"⏹ Сканирование прервано. Найдено: {len(results)} записей"))
        else:
            q.put(("scan", f"Шаблон: {mask}"))
            if compare: q.put(("scan", f"Сравнение: {cmask}"))
            if use_list:
                mm = 'точное' if list_match == 'exact' else 'содержит'
                q.put(("scan", f"Фильтр по списку: ВКЛ ({len(list_articles)} арт., {mm})"))
                q.put(("scan", f"Пропущено папок (нет в списке): {skipped_by_list}"))
            else:
                q.put(("scan", "Фильтр по списку: ВЫКЛ"))
            q.put(("scan", f"Всего папок: {total_folders}")); q.put(("scan", ""))
            if compare:
                if cmode == 'missing_pair': q.put(("scan", f"РЕЖИМ: где есть '{mask}', но НЕТ '{cmask}'"))
                else: q.put(("scan", f"РЕЖИМ: где есть '{cmask}', но НЕТ '{mask}'"))
                articles = []
                for _, arts in missing: articles.extend(arts)
                seen, uniq = set(), []
                for a in articles:
                    if a and a not in seen: seen.add(a); uniq.append(a)
                q.put(("scan", f"Найдено артикулов: {len(uniq)}"))
                q.put(("scan", "")); q.put(("scan", f"--- Артикулы ({len(uniq)}) ---")); q.put(("scan", ""))
                for a in uniq: q.put(("scan", a))
            else:
                if fmode == 'exists':
                    q.put(("scan", f"РЕЖИМ: где есть '{mask}'"))
                    articles = []
                    for _, arts in found: articles.extend(arts)
                    seen, uniq = set(), []
                    for a in articles:
                        if a and a not in seen: seen.add(a); uniq.append(a)
                    q.put(("scan", f"Найдено артикулов: {len(uniq)}"))
                    q.put(("scan", "")); q.put(("scan", f"--- Артикулы ({len(uniq)}) ---")); q.put(("scan", ""))
                    for a in uniq: q.put(("scan", a))
                else:
                    folders_only = [r for r, _ in missing if r != ""]
                    q.put(("scan", f"РЕЖИМ: где НЕТ '{mask}'"))
                    q.put(("scan", f"Найдено папок: {len(folders_only)}"))
                    q.put(("scan", "")); q.put(("scan", f"--- Папки ({len(folders_only)}) ---")); q.put(("scan", ""))
                    for r in folders_only: q.put(("scan", r))
            q.put(("scan", ""))
            q.put(("scan", f"Время сканирования: {dur}"))
            q.put(("scan", f"Найдено для редактора: {len(results)} записей"))
            q.put(("scan", f"Все {len(results)} записей отмечены по умолчанию."))
            q.put(("DONE_SCAN", f"Сканирование завершено ({len(results)} записей)"))

    # ---------- Очередь ----------
    def poll_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if isinstance(msg, tuple):
                    kind, text = msg
                    if kind == "DONE_SCAN":
                        self.is_processing = False; self.stop_animation()
                        self.progress.stop(); self.progress.config(mode='determinate')
                        self.progress['value'] = 100; self.lbl_progress.config(text="")
                        self.status_var.set(text)
                        self.btn_scan.config(state=tk.NORMAL)
                        self.btn_scan_settings.config(state=tk.NORMAL)
                        self.btn_stop_scan.config(state=tk.DISABLED)
                        self.btn_stop_editor.config(state=tk.DISABLED)
                        self.update_editor_state()
                    elif kind == "DONE_EDIT":
                        self.is_processing = False; self.stop_animation()
                        self.progress.stop(); self.progress.config(mode='determinate')
                        self.progress['value'] = 100; self.lbl_progress.config(text="")
                        self.status_var.set(text)
                        self.btn_scan.config(state=tk.NORMAL)
                        self.btn_scan_settings.config(state=tk.NORMAL)
                        self.btn_stop_scan.config(state=tk.DISABLED)
                        self.btn_stop_editor.config(state=tk.DISABLED)
                        self.update_editor_state(); self.update_journal_label()
                    elif kind == "PROGRESS":
                        try:
                            cur, total = text
                            pct = int(cur * 100 / total) if total else 0
                            self.progress['value'] = pct
                            self.lbl_progress.config(text=f"Обработано: {cur} из {total} ({pct}%)")
                        except Exception: pass
                    else:
                        self.log_message(kind, text)
                else:
                    self.status_var.set(str(msg))
        except queue.Empty:
            pass
        self.root.after(80, self.poll_queue)

    # ---------- Сбор данных ----------
    def _collect_items_from_path(self, root_dir):
        items = []
        stack = [(root_dir, "")]
        while stack:
            abs_path, rel_path = stack.pop()
            try:
                with os.scandir(abs_path) as it:
                    dirs, files = [], []
                    for entry in it:
                        try:
                            if entry.is_dir(follow_symlinks=False): dirs.append(entry.name)
                            elif entry.is_file(follow_symlinks=False): files.append(entry.name)
                        except OSError: continue
            except (PermissionError, FileNotFoundError, OSError): continue
            items.append({'article': os.path.basename(abs_path), 'folder_abs': abs_path,
                          'folder_rel': rel_path if rel_path else ".", 'files': files})
            for d in dirs:
                sub_abs = os.path.join(abs_path, d)
                sub_rel = os.path.join(rel_path, d) if rel_path else d
                stack.append((sub_abs, sub_rel))
        return items

    def _parse_filters(self):
        f = {}
        raw = self.editor_filter_exts.get().strip().lower()
        if raw:
            exts = set()
            for item in raw.replace(';', ',').split(','):
                item = item.strip()
                if item:
                    if not item.startswith('.'): item = '.' + item
                    exts.add(item)
            f['exts'] = exts
        mask = self.editor_filter_mask.get().strip()
        if mask: f['mask'] = mask
        try:
            v = self.editor_filter_size_min.get().strip()
            if v: f['size_min'] = float(v) * 1024
        except ValueError: pass
        try:
            v = self.editor_filter_size_max.get().strip()
            if v: f['size_max'] = float(v) * 1024
        except ValueError: pass
        for key, entry in [('date_create_from', self.editor_filter_date_create_from),
                           ('date_create_to', self.editor_filter_date_create_to),
                           ('date_mod_from', self.editor_filter_date_mod_from),
                           ('date_mod_to', self.editor_filter_date_mod_to)]:
            ts = parse_date(entry.get())
            if ts is not None: f[key] = ts
        return f

    def _file_matches_filters(self, filepath, filters):
        if not filters: return True
        try: st = os.stat(filepath)
        except OSError: return False
        if 'exts' in filters:
            if os.path.splitext(filepath)[1].lower() not in filters['exts']: return False
        if 'mask' in filters:
            name = os.path.basename(filepath)
            if not fnmatch.fnmatch(name.lower(), filters['mask'].lower()): return False
        if 'size_min' in filters and st.st_size < filters['size_min']: return False
        if 'size_max' in filters and st.st_size > filters['size_max']: return False
        if 'date_create_from' in filters and st.st_ctime < filters['date_create_from']: return False
        if 'date_create_to' in filters and st.st_ctime > filters['date_create_to']: return False
        if 'date_mod_from' in filters and st.st_mtime < filters['date_mod_from']: return False
        if 'date_mod_to' in filters and st.st_mtime > filters['date_mod_to']: return False
        return True

    # ---------- Редактор ----------
    def _get_source_items(self):
        if self.editor_source.get() == "scan":
            if not self.last_scan_results:
                return None, "Нет данных скана. Выполните сканирование."
            if not self.checked_scan_indices:
                return None, "Ничего не отмечено. Откройте таблицу и отметьте хотя бы одну папку."
            items = [self.last_scan_results[i]
                     for i in sorted(self.checked_scan_indices)
                     if 0 <= i < len(self.last_scan_results)]
            return items, None
        else:
            manual = self.editor_manual_path.get().strip().strip('"').strip("'")
            if not manual or not os.path.exists(manual):
                return None, "Укажите существующую папку-источник"
            return self._collect_items_from_path(manual), None

    def _count_will_change(self, source_items, is_folder, ren_from, ren_to, add_pos):
        changed = 0
        if is_folder:
            for rec in source_items:
                base = os.path.basename(rec['folder_abs'])
                new = self._apply_rename_all(base, ren_from, ren_to, add_pos)
                if new != base: changed += 1
        else:
            for rec in source_items:
                for fname in rec.get('files', []):
                    base, ext = os.path.splitext(fname)
                    new = self._apply_rename_all(base, ren_from, ren_to, add_pos)
                    if new != base: changed += 1
        return changed

    def start_editor(self):
        if self.is_processing: return
        source_items, err = self._get_source_items()
        if err:
            messagebox.showerror("Ошибка", err); return
        if not self.work_folders.get() and not self.work_files.get():
            messagebox.showerror("Ошибка", "Включите хотя бы одно действие — папки или файлы."); return

        where = self.editor_where.get()
        folder_action = self.folder_action.get() if self.work_folders.get() else None
        file_action = self.file_action.get() if self.work_files.get() else None
        folder_ren_from = self.folder_replace_from.get() if self.work_folders.get() else ""
        folder_ren_to = self.folder_replace_to.get() if self.work_folders.get() else ""
        folder_add_pos = self.folder_add_position.get()
        file_ren_from = self.file_replace_from.get() if self.work_files.get() else ""
        file_ren_to = self.file_replace_to.get() if self.work_files.get() else ""
        file_add_pos = self.file_add_position.get()
        file_filter_mode = self.file_filter_mode.get()

        if folder_action and folder_action != 'delete' and folder_ren_from and not folder_ren_to:
            base = self._get_sample_folder_name()
            sample = f"{base}{folder_ren_from}  →  {base}"
            if not messagebox.askyesno("Внимание",
                f"Папки: поле «на» пустое.\nВсе «{folder_ren_from}» будут удалены из имён.\n\nПример: {sample}\n\nПродолжить?"):
                return
        if file_action and file_action != 'delete' and file_ren_from and not file_ren_to:
            sample = self._get_sample_file_name()
            b, e = os.path.splitext(sample)
            sample2 = f"{b.replace(file_ren_from, '')}{e}"
            if not messagebox.askyesno("Внимание",
                f"Файлы: поле «на» пустое.\nВсе «{file_ren_from}» будут удалены из имён.\n\nПример: {sample} → {sample2}\n\nПродолжить?"):
                return

        total_changes = 0
        not_empty_ren = False
        if self.work_folders.get() and folder_action != 'delete' and folder_ren_from:
            total_changes += self._count_will_change(source_items, True, folder_ren_from, folder_ren_to, folder_add_pos)
            not_empty_ren = True
        if self.work_files.get() and file_action != 'delete' and file_ren_from:
            total_changes += self._count_will_change(source_items, False, file_ren_from, file_ren_to, file_add_pos)
            not_empty_ren = True
        if not_empty_ren and total_changes == 0:
            if not messagebox.askyesno("Ничего не изменится",
                "Ни в одном имени не найдено указанных подстрок для замены.\n"
                "Переименования не будет. Файлы и папки будут скопированы/перемещены\n"
                "с прежними именами (если действие «Копировать/Переместить»).\n\n"
                "Продолжить всё равно?"):
                return

        dst = ""
        if where == "other":
            need_dst = False
            if self.work_folders.get() and folder_action != 'delete': need_dst = True
            if self.work_files.get() and file_action != 'delete': need_dst = True
            if need_dst:
                dst = self.editor_dest_path.get().strip().strip('"').strip("'")
                if not dst:
                    messagebox.showerror("Ошибка", "Укажите целевую папку"); return
                try: os.makedirs(dst, exist_ok=True)
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось создать целевую папку:\n{e}"); return

        filters = None
        if self.work_files.get() and file_filter_mode == "filtered":
            filters = self._parse_filters()

        warnings = []
        if self.work_files.get() and file_action == "delete":
            warnings.append("Файлы будут удалены в Корзину Windows.")
        if folder_action == "delete":
            warnings.append("Папки будут удалены в Корзину Windows вместе с содержимым.")
        if warnings:
            txt = "\n".join(warnings) + "\n\nПродолжить?"
            if not messagebox.askyesno("Подтверждение", txt): return

        self.is_processing = True
        self.stop_requested = False
        self.btn_scan.config(state=tk.DISABLED)
        self.btn_scan_settings.config(state=tk.DISABLED)
        self.btn_apply.config(state=tk.DISABLED)
        self.btn_undo.config(state=tk.DISABLED)
        self.btn_stop_editor.config(state=tk.NORMAL)
        self.btn_stop_scan.config(state=tk.DISABLED)
        self.start_animation("Редактор")
        self.clear_log(self.log_edit)
        self.progress.config(mode='determinate', value=0, maximum=100)
        self.lbl_progress.config(text="Подготовка...")
        threading.Thread(target=self._editor_worker,
                         args=(source_items, filters), daemon=True).start()

    def _editor_worker(self, source_items, filters):
        t0 = datetime.datetime.now()
        stat = {'files_ok': 0, 'folders_ok': 0, 'errors': 0, 'deleted': 0,
                'not_renamed': 0, 'skipped_no_match': 0,
                'skipped_exists': 0, 'skipped_same': 0}
        q = self.log_queue

        work_folders = self.work_folders.get()
        folder_action = self.folder_action.get()
        folder_ren_from = self.folder_replace_from.get()
        folder_ren_to = self.folder_replace_to.get()
        folder_add_pos = self.folder_add_position.get()

        work_files = self.work_files.get()
        file_action = self.file_action.get()
        file_ren_from = self.file_replace_from.get()
        file_ren_to = self.file_replace_to.get()
        file_add_pos = self.file_add_position.get()
        file_filter_mode = self.file_filter_mode.get()

        where = self.editor_where.get()
        dst = self.editor_dest_path.get().strip().strip('"').strip("'")
        skip = self.editor_skip_existing.get()
        upd_date = self.editor_update_date.get()
        keep_struct = self.editor_keep_structure.get()

        journal = {'timestamp': t0.strftime('%Y-%m-%d %H:%M:%S'),
                   'action': 'mixed', 'files': [], 'folders': [], 'deleted': []}

        q.put(("edit", f"=== РЕДАКТОР: {t0.strftime('%Y-%m-%d %H:%M:%S')} ==="))
        src_label = 'отмеченное в скане' if self.editor_source.get() == 'scan' else self.editor_manual_path.get()
        q.put(("edit", f"Источник: {src_label}"))
        q.put(("edit", f"Куда: {'в другую папку ' + dst if where == 'other' else 'в ту же папку'}"))
        if work_folders: q.put(("edit", f"Папки: {folder_action}"))
        if work_files and not (work_folders and folder_action == 'delete'):
            q.put(("edit", f"Файлы: {file_action}"))
        q.put(("edit", f"Всего записей: {len(source_items)}"))
        q.put(("edit", ""))

        total_ops = 0
        for rec in source_items:
            if work_files and not (work_folders and folder_action == 'delete'):
                for fname in rec.get('files', []):
                    src_file = os.path.join(rec['folder_abs'], fname)
                    if not os.path.isfile(src_file): continue
                    if file_filter_mode == 'filtered' and not self._file_matches_filters(src_file, filters):
                        continue
                    total_ops += 1
            if work_folders: total_ops += 1

        done = 0
        interrupted = False
        lock = threading.Lock()

        use_pool = (self.general_settings.get('use_thread_pool', False)
                    and work_files and file_action in ('copy', 'move'))
        pool_threads = 4
        if use_pool:
            try: pool_threads = max(2, min(16, int(self.general_settings.get('thread_count', 4))))
            except Exception: pool_threads = 4
        ex = None
        if use_pool:
            ex = ThreadPoolExecutor(max_workers=pool_threads)
            q.put(("edit", f"Многопоточность файлов: ВКЛ ({pool_threads} потоков)"))
        else:
            q.put(("edit", "Многопоточность файлов: ВЫКЛ"))

        try:
            for rec in source_items:
                if self.stop_requested:
                    interrupted = True; break
                src_folder = rec['folder_abs']
                rel_folder = rec['folder_rel']

                if work_files and not (work_folders and folder_action == 'delete'):
                    file_names = list(rec.get('files', []))
                    filtered = []
                    for fname in file_names:
                        src_file = os.path.join(src_folder, fname)
                        if not os.path.isfile(src_file): continue
                        if file_filter_mode == 'filtered' and not self._file_matches_filters(src_file, filters):
                            continue
                        filtered.append(fname)
                    if filtered:
                        if use_pool and len(filtered) > 1:
                            futs = [ex.submit(self._process_file_item,
                                              src_folder, rel_folder, fn,
                                              file_action, file_ren_from, file_ren_to, file_add_pos,
                                              where, dst, skip, upd_date, keep_struct,
                                              journal, stat, lock)
                                    for fn in filtered]
                            for _ in as_completed(futs):
                                with lock:
                                    done += 1
                                    if done % 10 == 0 or done == total_ops:
                                        q.put(("PROGRESS", (done, total_ops)))
                        else:
                            for fname in filtered:
                                if self.stop_requested:
                                    interrupted = True; break
                                self._process_file_item(src_folder, rel_folder, fname,
                                                        file_action, file_ren_from, file_ren_to, file_add_pos,
                                                        where, dst, skip, upd_date, keep_struct,
                                                        journal, stat, None)
                                done += 1
                                if done % 10 == 0 or done == total_ops:
                                    q.put(("PROGRESS", (done, total_ops)))

                if work_folders and not interrupted:
                    self._process_folder_item(rec, folder_action, folder_ren_from, folder_ren_to, folder_add_pos,
                                              where, dst, skip, journal, stat, q)
                    done += 1
                    if done % 10 == 0 or done == total_ops:
                        q.put(("PROGRESS", (done, total_ops)))
        finally:
            if ex is not None:
                ex.shutdown(wait=False)

        dur = datetime.datetime.now() - t0
        n_ok = len(journal['files']) + len(journal['folders'])
        if n_ok > 0:
            save_journal(journal)

        q.put(("edit", ""))
        q.put(("edit", "=" * 55))
        if interrupted:
            q.put(("edit", "⏹ ОПЕРАЦИЯ ПРЕРВАНА ПОЛЬЗОВАТЕЛЕМ"))
        q.put(("edit", "ИТОГОВАЯ СТАТИСТИКА:"))
        q.put(("edit", f"Время: {dur}"))
        q.put(("edit", f"Папок успешно: {stat['folders_ok']}"))
        q.put(("edit", f"Файлов успешно: {stat['files_ok']}"))
        if stat['not_renamed'] > 0:
            q.put(("edit", f"  из них без переименования: {stat['not_renamed']}"))
        q.put(("edit", f"Удалено: {stat['deleted']}"))
        q.put(("edit", f"Ошибок: {stat['errors']}"))
        total_skip = stat['skipped_no_match'] + stat['skipped_exists'] + stat['skipped_same']
        q.put(("edit", f"Пропущено всего: {total_skip}"))
        if stat['skipped_no_match'] > 0:
            q.put(("edit", f"  • «подстрока» не найдена в имени: {stat['skipped_no_match']}"))
        if stat['skipped_exists'] > 0:
            q.put(("edit", f"  • уже существует в целевой папке: {stat['skipped_exists']}"))
        if stat['skipped_same'] > 0:
            q.put(("edit", f"  • имя не изменилось, В ту же папку: {stat['skipped_same']}"))
        q.put(("edit", "=" * 55))

        total_ok = stat['files_ok'] + stat['folders_ok']
        msg = (f"Готово. Успешно: {total_ok}, удалено: {stat['deleted']}, "
               f"пропущено: {total_skip}, ошибок: {stat['errors']}")
        if interrupted:
            msg = "⏹ Прервано. " + msg
        q.put(("DONE_EDIT", msg))

    def _build_new_base(self, base, ren_from, ren_to, add_position):
        if ren_from: return base.replace(ren_from, ren_to)
        elif ren_to:
            if add_position == "prefix": return ren_to + base
            return base + ren_to
        return base

    def _rename_reason(self, base, ren_from, ren_to):
        if ren_from and ren_from not in base: return 'no_match'
        if not ren_from and not ren_to: return 'same'
        return None

    def _process_folder_item(self, rec, action, ren_from, ren_to, add_pos,
                             where, dst, skip, journal, stat, q):
        src_folder = rec['folder_abs']
        rel_folder = rec['folder_rel']
        base = os.path.basename(src_folder)
        new_name = self._build_new_base(base, ren_from, ren_to, add_pos)
        try:
            if action == 'delete':
                if HAS_TRASH: send2trash(src_folder)
                else: shutil.rmtree(src_folder)
                stat['deleted'] += 1
                journal['deleted'].append({'path': src_folder})
                q.put(("edit", f"[x] Удалена папка: {rel_folder}"))
                return
            if where == "same":
                if new_name == base:
                    reason = self._rename_reason(base, ren_from, ren_to)
                    if reason == 'no_match':
                        stat['skipped_no_match'] += 1
                        q.put(("edit", f"[~] Пропущена папка: {base} — «{ren_from}» не найдено в имени"))
                    else:
                        stat['skipped_same'] += 1
                        q.put(("edit", f"[~] Пропущена папка: {base} — имя не изменилось"))
                    return
                new_path = os.path.join(os.path.dirname(src_folder), new_name)
                if os.path.exists(new_path) and skip:
                    stat['skipped_exists'] += 1
                    q.put(("edit", f"[~] Пропущена папка: {base} — «{new_name}» уже существует"))
                    return
                if action == 'move': shutil.move(src_folder, new_path)
                else: shutil.copytree(src_folder, new_path)
                stat['folders_ok'] += 1
                journal['folders'].append({'from': src_folder, 'to': new_path})
                q.put(("edit", f"[+] Папка: {base} → {new_name}"))
            else:
                target = os.path.join(dst, rel_folder) if rel_folder not in ("", ".") else os.path.join(dst, base)
                if os.path.exists(target) and skip:
                    stat['skipped_exists'] += 1
                    q.put(("edit", f"[~] Пропущена папка: {base} — «{target}» уже существует"))
                    return
                if action == 'move': shutil.move(src_folder, target)
                else: shutil.copytree(src_folder, target, dirs_exist_ok=True)
                stat['folders_ok'] += 1
                if new_name == base: stat['not_renamed'] += 1
                journal['folders'].append({'from': src_folder, 'to': target})
                q.put(("edit", f"[+] Папка: {base} → {target}"))
        except Exception as e:
            stat['errors'] += 1
            q.put(("edit", f"[!] ОШИБКА папки {rel_folder}: {e}"))

    def _process_file_item(self, src_folder, rel_folder, fname,
                           action, ren_from, ren_to, add_pos,
                           where, dst, skip, upd_date, keep_struct,
                           journal, stat, lock):
        q = self.log_queue
        src_file = os.path.join(src_folder, fname)
        base, ext = os.path.splitext(fname)
        new_base = self._build_new_base(base, ren_from, ren_to, add_pos)
        new_name = new_base + ext

        def inc(key, delta=1):
            if lock:
                with lock: stat[key] += delta
            else: stat[key] += delta

        try:
            if action == 'delete':
                if HAS_TRASH: send2trash(src_file)
                else: os.remove(src_file)
                inc('deleted')
                if lock:
                    with lock: journal['deleted'].append({'path': src_file})
                else:
                    journal['deleted'].append({'path': src_file})
                q.put(("edit", f"[x] Удалён файл: {fname}"))
                return
            if where == "same":
                if new_name == fname:
                    reason = self._rename_reason(base, ren_from, ren_to)
                    if reason == 'no_match':
                        inc('skipped_no_match')
                        q.put(("edit", f"[~] Пропущен файл: {fname} — «{ren_from}» не найдено в имени"))
                    else:
                        inc('skipped_same')
                        q.put(("edit", f"[~] Пропущен файл: {fname} — имя не изменилось"))
                    return
                target_file = os.path.join(src_folder, new_name)
                if os.path.exists(target_file) and skip:
                    inc('skipped_exists')
                    q.put(("edit", f"[~] Пропущен файл: {fname} — «{new_name}» уже существует"))
                    return
                if action == 'move': shutil.move(src_file, target_file)
                else: shutil.copy(src_file, target_file)
                if upd_date:
                    now = datetime.datetime.now().timestamp()
                    os.utime(target_file, (now, now))
                inc('files_ok')
                if lock:
                    with lock: journal['files'].append({'from': src_file, 'to': target_file})
                else:
                    journal['files'].append({'from': src_file, 'to': target_file})
                q.put(("edit", f"[+] Файл: {fname} → {new_name}"))
            else:
                if keep_struct:
                    target_file = os.path.join(dst, rel_folder if rel_folder not in ("", ".") else "", new_name)
                else:
                    target_file = os.path.join(dst, new_name)
                os.makedirs(os.path.dirname(target_file), exist_ok=True)
                if os.path.exists(target_file) and skip:
                    inc('skipped_exists')
                    q.put(("edit", f"[~] Пропущен файл: {fname} — «{target_file}» уже существует"))
                    return
                if action == 'move': shutil.move(src_file, target_file)
                else: shutil.copy(src_file, target_file)
                if upd_date:
                    now = datetime.datetime.now().timestamp()
                    os.utime(target_file, (now, now))
                inc('files_ok')
                if new_name == fname: inc('not_renamed')
                if lock:
                    with lock: journal['files'].append({'from': src_file, 'to': target_file})
                else:
                    journal['files'].append({'from': src_file, 'to': target_file})
                q.put(("edit", f"[+] Файл: {fname} → {target_file}"))
        except Exception as e:
            inc('errors')
            q.put(("edit", f"[!] ОШИБКА файла {fname}: {e}"))

    # ---------- Откат ----------
    def undo_last_operation(self):
        if self.is_processing:
            messagebox.showinfo("Занято", "Дождитесь завершения текущей операции."); return
        j = load_journal()
        if not j:
            messagebox.showinfo("Нет журнала", "Нет данных для отката."); return
        n_files = len(j.get('files', []))
        n_folders = len(j.get('folders', []))
        n_deleted = len(j.get('deleted', []))
        action = j.get('action', '?')
        if n_files + n_folders == 0 and n_deleted == 0:
            messagebox.showinfo("Пусто", "Журнал пуст."); return
        txt = (f"Отменить последнюю операцию?\n\nТип: {action}\n"
               f"Дата: {j.get('timestamp', '?')}\n\nБудет откачено:\n"
               f"  • Файлов: {n_files}\n  • Папок: {n_folders}\n")
        if n_deleted > 0:
            txt += f"  • Удалённых: {n_deleted} (откат НЕВОЗМОЖЕН — восстановите из Корзины)\n"
        txt += "\nПродолжить?"
        if not messagebox.askyesno("Откат операции", txt): return
        self.is_processing = True
        self.stop_requested = False
        self.btn_apply.config(state=tk.DISABLED)
        self.btn_undo.config(state=tk.DISABLED)
        self.btn_scan.config(state=tk.DISABLED)
        self.btn_scan_settings.config(state=tk.DISABLED)
        self.start_animation("Откат")
        self.log_message("edit", "")
        self.log_message("edit", f"=== ОТКАТ ОПЕРАЦИИ ОТ {j.get('timestamp', '?')} ===")
        threading.Thread(target=self._undo_worker, args=(j,), daemon=True).start()

    def _undo_worker(self, j):
        q = self.log_queue
        ok, err, skipped = 0, 0, 0
        for item in j.get('files', []):
            src = item.get('from'); dst = item.get('to')
            try:
                if os.path.exists(dst) and not os.path.exists(src):
                    os.makedirs(os.path.dirname(src), exist_ok=True)
                    shutil.move(dst, src); ok += 1
                    q.put(("edit", f"[↶] Возвращён файл: {dst} → {src}"))
                elif os.path.exists(dst) and os.path.exists(src):
                    if HAS_TRASH: send2trash(dst)
                    else: os.remove(dst)
                    ok += 1
                    q.put(("edit", f"[↶] Удалена копия файла: {dst}"))
                else:
                    skipped += 1
                    q.put(("edit", f"[?] Нет файла для отката: {dst}"))
            except Exception as e:
                err += 1
                q.put(("edit", f"[!] ОШИБКА отката файла {dst}: {e}"))
        for item in j.get('folders', []):
            src = item.get('from'); dst = item.get('to')
            try:
                if os.path.exists(dst) and not os.path.exists(src):
                    os.makedirs(os.path.dirname(src), exist_ok=True)
                    shutil.move(dst, src); ok += 1
                    q.put(("edit", f"[↶] Возвращена папка: {dst} → {src}"))
                elif os.path.exists(dst) and os.path.exists(src):
                    if HAS_TRASH: send2trash(dst)
                    else: shutil.rmtree(dst)
                    ok += 1
                    q.put(("edit", f"[↶] Удалена копия папки: {dst}"))
                else:
                    skipped += 1
                    q.put(("edit", f"[?] Нет папки для отката: {dst}"))
            except Exception as e:
                err += 1
                q.put(("edit", f"[!] ОШИБКА отката папки {dst}: {e}"))
        n_deleted = len(j.get('deleted', []))
        if n_deleted > 0:
            q.put(("edit", f"[i] Удалённых файлов/папок: {n_deleted}. Восстановите из Корзины Windows вручную."))
        try:
            if ok > 0:
                os.remove(JOURNAL_FILE)
                q.put(("edit", "[✓] Журнал очищен."))
        except Exception as e:
            q.put(("edit", f"[!] Не удалось удалить журнал: {e}"))
        q.put(("edit", ""))
        q.put(("edit", "=" * 55))
        q.put(("edit", f"ОТКАТ ЗАВЕРШЁН: успешно {ok}, пропущено {skipped}, ошибок {err}"))
        q.put(("edit", "=" * 55))
        q.put(("DONE_EDIT", f"Откат завершён: успешно {ok}, пропущено {skipped}, ошибок {err}"))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
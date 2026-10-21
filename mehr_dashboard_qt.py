#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
داشبورد قراردادهای مهر خلیج فارس — نسخهٔ دسکتاپ (PyQt5)
=========================================================

این برنامه تبدیل نسخهٔ Flask/dashboard.html به یک برنامهٔ دسکتاپ تک‌فایلی است:

* دیگر به Flask، مرورگر، dashboard.html یا فایل mehr_optimized.py نیاز ندارد.
* داده‌ها مستقیماً از همان فایل اکسل قبلی خوانده می‌شوند
  (شیت colorcode و همان شماره‌ستون‌های نسخهٔ وب) تا هیچ محاسبه‌ای تغییر نکند.
* نمودارها، فیلترها، جستجو، جزئیات قرارداد، یادداشت‌ها، خروجی اکسل، گزارش
  مدیریتی و اتصال فایل PDF همه درون همین فایل پیاده‌سازی شده‌اند.

اجرا:
    python mehr_dashboard_qt.py

پیش‌نیازها:
    pip install PyQt5 openpyxl

گزینه‌های خط فرمان:
    --excel PATH     مسیر فایل اکسل (بر موقت بر تنظیمات ذخیره‌شده اولویت دارد)
    --check          فقط صحت خواندن اکسل و آمار را بررسی کن (بدون باز کردن پنجره)
"""

from __future__ import annotations

import argparse
import html as html_module
import json
import math
import os
import sys
import traceback
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

# ---------------------------------------------------------------------------
# وارد کردن PyQt5 با پیام خطای فارسی
# ---------------------------------------------------------------------------
try:
    from PyQt5.QtCore import (
        QAbstractTableModel,
        QByteArray,
        QModelIndex,
        QPointF,
        QRectF,
        QSize,
        QThread,
        Qt,
        QTimer,
        QUrl,
        pyqtSignal,
    )
    from PyQt5.QtGui import (
        QBrush,
        QColor,
        QDesktopServices,
        QFont,
        QFontDatabase,
        QFontMetrics,
        QIcon,
        QKeySequence,
        QLinearGradient,
        QPainter,
        QPainterPath,
        QPalette,
        QPen,
        QPixmap,
    )
    from PyQt5.QtWidgets import (
        QAbstractItemView,
        QAction,
        QApplication,
        QButtonGroup,
        QCheckBox,
        QComboBox,
        QDialog,
        QDialogButtonBox,
        QDockWidget,
        QFileDialog,
        QFormLayout,
        QFrame,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLayout,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMenu,
        QMessageBox,
        QPlainTextEdit,
        QProgressBar,
        QPushButton,
        QScrollArea,
        QShortcut,
        QSizePolicy,
        QSpinBox,
        QStackedWidget,
        QAbstractSpinBox,
        QStyledItemDelegate,
        QStyleOptionViewItem,
        QStyle,
        QStyleFactory,
        QTableView,
        QTextBrowser,
        QToolButton,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:  # pragma: no cover
    sys.stderr.write(
        "خطا: کتابخانهٔ PyQt5 نصب نیست.\n"
        "لطفاً با دستور زیر نصب کنید:\n"
        "    pip install PyQt5 openpyxl\n"
        f"جزئیات: {exc}\n"
    )
    raise SystemExit(1)

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font as XLFont, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    openpyxl = None  # type: ignore
    Alignment = Border = XLFont = PatternFill = Side = None  # type: ignore
    get_column_letter = None  # type: ignore


# ---------------------------------------------------------------------------
# ثابت‌ها
# ---------------------------------------------------------------------------
APP_NAME = "داشبورد قراردادهای مهر خلیج فارس"
APP_SHORT = "سامانه کندو"
APP_VERSION = "2.0.0"
SHEET_NAME_DEFAULT = "colorcode"
REQUIRED_COLUMNS = 57  # ستون‌های مورد استفاده تا ستون 55 (BB)

# مسیرهای اصلی نسخهٔ وب (فقط به‌عنوان مقدار پیش‌فرض/راهنما نگه داشته شده‌اند)
ORIGINAL_EXCEL_PATH = (
    r"\\appserver93\Engineering\INVOICE\Private\Contractor Invoice"
    r"\ژیانی\کندو\F.O.C(Mehr).xlsx"
)
ORIGINAL_PDF_FOLDER = (
    r"\\appserver93\Engineering\INVOICE\Private\Contractor Invoice"
    r"\ژیانی\کندو\PDF-MEHR"
)

EXCEL_FILTER = "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*.*)"
PDF_FILTER = "همهٔ فایل‌ها (*.*)"

# هشت وضعیت پیگیری که در ستون‌های مشخصی از اکسل با حرف M علامت خورده‌اند
FLAG_DEFS: List[Tuple[str, str, int]] = [
    ("temporary", "تحویل موقت", 11),
    ("prefinal", "صورت‌وضعیت ماقبل قطعی", 12),
    ("balance", "بالانس متریال", 16),
    ("statement", "صورت‌وضعیت قطعی", 21),
    ("mofasa", "مفاصاحساب بیمه", 23),
    ("affidavit", "اقرارنامه", 24),
    ("delivery", "تحویل قطعی", 28),
    ("financial", "تسویهٔ مالی نهایی", 30),
]
FLAG_LABELS = {key: label for key, label, _ in FLAG_DEFS}

# فیلترهای جدول (کلید ← عنوان)
FILTER_CHOICES: List[Tuple[str, str]] = [
    ("all", "همهٔ قراردادها"),
    ("deductions", "دارای کسورات اسناد اختتام"),
    ("bookmarked", "فقط نشان‌شده‌ها"),
    ("expired", "قراردادهای منقضی‌شده"),
    ("mofasa", "نیازمند مفاصاحساب بیمه"),
    ("financial", "نیازمند تسویهٔ مالی"),
    ("balance", "نیازمند بالانس متریال"),
    ("temporary", "نیازمند تحویل موقت"),
    ("delivery", "نیازمند تحویل قطعی"),
    ("prefinal", "نیازمند صورت‌وضعیت ماقبل قطعی"),
    ("statement", "نیازمند صورت‌وضعیت قطعی"),
    ("affidavit", "نیازمند اقرارنامه"),
]

# رنگ نمودارها، مطابق پالت نسخهٔ وب
CHART_COLORS: Dict[str, List[str]] = {
    "type": [
        "#0d6b63", "#1f8a7e", "#39a99a", "#6ec7b8", "#a7ded4",
        "#f0b45a", "#d98e5f", "#b06a6a", "#7b8fa8", "#9a8ab5",
    ],
    "status": ["#0f766e", "#2f9e8f", "#6ee7c7", "#a7f3d0", "#d1fae5"],
    "progress": ["#92400e", "#c96f4a", "#e3a06b", "#f3c98b", "#fbe3c0"],
    "completion": ["#b45309", "#d97706", "#f59e0b", "#fcd34d"],
    "annex": [
        "#8b5cf6", "#10b981", "#0ea5e9", "#f59e0b",
        "#ef4444", "#ec4899", "#6366f1", "#14b8a6",
    ],
}

CHART_TITLES: Dict[str, str] = {
    "type": "نوع قرارداد",
    "status": "وضعیت پکیج آزادسازی",
    "progress": "پیشرفت کارکرد",
    "completion": "پیشرفت اختتام",
    "annex": "الحاقیه / زمان باقی‌مانده",
}


# ---------------------------------------------------------------------------
# ابزارهای عمومی متن و عدد
# ---------------------------------------------------------------------------
_PERSIAN_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
_LATIN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def fa(value: Any) -> str:
    """ارقام لاتین را برای نمایش به فارسی تبدیل می‌کند."""
    return str(value).translate(_PERSIAN_DIGITS)


def bidi(text: str) -> str:
    """اگر متن با رقم/درصد شروع شود، جهت پاراگراف را راست‌به‌چپ می‌کند.

    دلیل: مترجم Qt برای رشته‌ای که با رقم شروع شود جهت پایه را چپ‌به‌راست
    در نظر می‌گیرد و عدد در سمت چپ کادر می‌افتد. با افزودن علامت RLM، عدد در
    سمت راست نمایش داده می‌شود (هماهنگ با چیدمان فارسی).
    """
    if not text:
        return text
    first = text[0]
    if first.isdigit() or first in "۰۱۲۳۴۵۶۷۸۹٪%—-":
        return "\u200f" + text
    return text


def fa_num(value: Any) -> str:
    """عدد را با جداکنندهٔ هزارگان و ارقام فارسی برمی‌گرداند."""
    try:
        return fa(f"{int(value):,}")
    except Exception:
        return fa(value)


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, float):
        return math.isnan(value) or math.isinf(value)
    return False


def txt(value: Any) -> str:
    """تبدیل امن سلول اکسل به متن تمیز (بدون nan و بدون .0 اضافه)."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, datetime):
        return jalali_datetime(value)
    if isinstance(value, date):
        return jalali_datetime(datetime(value.year, value.month, value.day), with_time=False)
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return f"{value:.10f}".rstrip("0").rstrip(".")
    return str(value).strip()


def num(value: Any) -> Optional[float]:
    """تلاش برای تبدیل سلول به عدد (ارقام فارسی و جداکنندهٔ هزارگان هم پذیرفته می‌شود)."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    text = str(value).strip().translate(_LATIN_DIGITS).replace(",", "").replace("٬", "")
    if text == "":
        return None
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def normalize_per(text: Any) -> str:
    """استانداردسازی حروف فارسی برای جستجوی دقیق‌تر (همان تابع نسخهٔ وب)."""
    if not text:
        return ""
    return (
        str(text)
        .replace("ي", "ی")
        .replace("ك", "ک")
        .replace("‌", " ")
        .lower()
        .strip()
    )


def normalize_lookup_key(value: Any) -> str:
    """کلید مقایسهٔ شمارهٔ قرارداد با نام فایل PDF (همان منطق نسخهٔ وب)."""
    if value is None:
        return ""
    raw = str(value).translate(_LATIN_DIGITS)
    cleaned = "".join(c for c in raw.strip() if c.isalnum() or c in "/-_. ")
    return (
        cleaned.replace("/", "")
        .replace("-", "")
        .replace("_", "")
        .replace(" ", "")
        .replace(".", "")
    )


def digits_only(value: Any) -> str:
    return "".join(c for c in str(value).translate(_LATIN_DIGITS) if c.isdigit())


def normalize_progress_value(val: Any) -> Optional[str]:
    """ستون ۵۳: پیشرفت کارکرد (همان نگاشت نسخهٔ وب)."""
    if is_blank(val):
        return None
    number = num(val)
    if number is None:
        return None
    mapping = {
        1: "0%",
        2: "کمتر از %75",
        3: "100٪ - 75٪",
        4: "125%-100%",
        5: "بيشتر از %125",
    }
    return mapping.get(int(number)) if float(number).is_integer() else None


def normalize_completion_value(val: Any) -> Optional[str]:
    """ستون ۵۴: پیشرفت اختتام."""
    if is_blank(val):
        return None
    number = num(val)
    if number is None:
        return None
    mapping = {1: "0%", 2: "کمتر از 90%", 3: "100%-90%", 4: "100%"}
    return mapping.get(int(number)) if float(number).is_integer() else None


def normalize_annex_value(val: Any) -> Optional[str]:
    """ستون ۵۵: وضعیت الحاقیه / زمان باقی‌مانده."""
    if is_blank(val):
        return None
    raw = str(val).strip().upper()
    mapping = {
        "A": "صورت وضعیت قطعی ثبت شده",
        "B": "سنواتي",
        "C": "قرارداد جدید",
        "D": "منقضی شده",
        "E": "15 روز تا انقضا",
        "F": "30 روز تا انقضا",
        "G": "مدت دارد",
    }
    if raw in mapping:
        return mapping[raw]
    number = num(raw)
    if number is not None and float(number).is_integer():
        return {1: "کمتر از 10 روز تا پایان", 2: "کمتر از 30 روز تا پایان"}.get(int(number))
    return None


def value_counts(values: Iterable[str]) -> Tuple[List[str], List[int]]:
    """شمارش مقادیر با ترتیب نزولی تعداد (پایدار برای مقادیر هم‌تعداد)."""
    items = [v for v in values if v not in (None, "")]
    if not items:
        return [], []
    first_seen: Dict[str, int] = {}
    for index, item in enumerate(items):
        first_seen.setdefault(item, index)
    counter = Counter(items)
    ordered = sorted(counter.items(), key=lambda kv: (-kv[1], first_seen[kv[0]]))
    return [key for key, _ in ordered], [count for _, count in ordered]


# ---------------------------------------------------------------------------
# تاریخ شمسی (بدون نیاز به کتابخانهٔ jdatetime)
# ---------------------------------------------------------------------------
def gregorian_to_jalali(gy: int, gm: int, gd: int) -> Tuple[int, int, int]:
    """تبدیل تاریخ میلادی به هجری شمسی (الگوریتم استاندارد)."""
    g_days_in_month = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = gy + 1 if gm > 2 else gy
    days = (
        355666
        + (365 * gy)
        + ((gy2 + 3) // 4)
        - ((gy2 + 99) // 100)
        + ((gy2 + 399) // 400)
        + gd
        + g_days_in_month[gm - 1]
    )
    jy = -1595 + (33 * (days // 12053))
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm = 1 + days // 31
        jd = 1 + days % 31
    else:
        jm = 7 + (days - 186) // 30
        jd = 1 + (days - 186) % 30
    return jy, jm, jd


def jalali_date_str(dt: date) -> str:
    jy, jm, jd = gregorian_to_jalali(dt.year, dt.month, dt.day)
    return fa(f"{jy:04d}/{jm:02d}/{jd:02d}")


def jalali_datetime(dt: datetime, with_time: bool = True) -> str:
    jy, jm, jd = gregorian_to_jalali(dt.year, dt.month, dt.day)
    text = fa(f"{jy:04d}/{jm:02d}/{jd:02d}")
    if with_time:
        text += " " + fa(f"{dt.hour:02d}:{dt.minute:02d}")
    return text


def today_jalali(with_time: bool = True) -> str:
    return jalali_datetime(datetime.now(), with_time=with_time)


def jalali_filename_stamp() -> str:
    jy, jm, jd = gregorian_to_jalali(*date.today().timetuple()[:3])
    return f"{jy:04d}-{jm:02d}-{jd:02d}"


def human_size(size: int) -> str:
    if not size:
        return "—"
    value = float(size)
    for unit in ("بایت", "کیلوبایت", "مگابایت", "گیگابایت"):
        if value < 1024 or unit == "گیگابایت":
            return f"{fa(f'{value:.0f}')} {unit}"
        value /= 1024
    return fa(size)


# ---------------------------------------------------------------------------
# پوشه‌های برنامه و تنظیمات
# ---------------------------------------------------------------------------
def app_dir() -> Path:
    """پوشهٔ فایل اجرایی/اسکریپت."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


_DATA_DIR_CACHE: Optional[Path] = None


def data_dir() -> Path:
    """پوشه‌ای برای ذخیرهٔ تنظیمات و یادداشت‌ها (کنار برنامه یا پوشهٔ کاربر)."""
    global _DATA_DIR_CACHE
    if _DATA_DIR_CACHE is not None:
        return _DATA_DIR_CACHE
    candidates = [app_dir(), Path.home() / ".mehr_dashboard"]
    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            probe = candidate / ".write_test"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink()
            _DATA_DIR_CACHE = candidate
            return candidate
        except Exception:
            continue
    _DATA_DIR_CACHE = Path.home()
    return _DATA_DIR_CACHE


class AppConfig:
    """تنظیمات سادهٔ JSON (قابل انتقال و قابل خواندن)."""

    FILENAME = "mehr_qt_config.json"
    DEFAULTS: Dict[str, Any] = {
        "excel_path": "",
        "sheet_name": SHEET_NAME_DEFAULT,
        "pdf_folder": "",
        "theme": "light",
        "font_size": 10,
        "auto_refresh_minutes": 5,
        "columns": {},
        "last_author": "",
        "window_geometry": "",
    }

    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = Path(path) if path else (data_dir() / self.FILENAME)
        self.values: Dict[str, Any] = dict(self.DEFAULTS)
        self.load()

    def load(self) -> None:
        try:
            if self.path.exists():
                with open(self.path, "r", encoding="utf-8") as handle:
                    stored = json.load(handle)
                if isinstance(stored, dict):
                    for key, value in stored.items():
                        if key in self.DEFAULTS:
                            self.values[key] = value
        except Exception as exc:  # تنظیمات خراب نباید مانع اجرا شود
            print(f"هشدار: تنظیمات خوانده نشد ({exc})")

    def save(self) -> None:
        try:
            tmp = self.path.with_suffix(".json.tmp")
            with open(tmp, "w", encoding="utf-8") as handle:
                json.dump(self.values, handle, ensure_ascii=False, indent=2)
            os.replace(tmp, self.path)
        except Exception as exc:
            print(f"هشدار: ذخیرهٔ تنظیمات ممکن نشد ({exc})")

    def get(self, key: str, default: Any = None) -> Any:
        return self.values.get(key, self.DEFAULTS.get(key, default))

    def set(self, key: str, value: Any, save: bool = True) -> None:
        self.values[key] = value
        if save:
            self.save()


def find_default_excel(config: Optional[AppConfig] = None) -> str:
    """پیدا کردن فایل اکسل: متغیر محیطی، تنظیمات، کنار برنامه و در آخر مسیر شبکه."""
    env_path = os.environ.get("MEHR_EXCEL_PATH", "").strip()
    if env_path and os.path.exists(env_path):
        return env_path

    if config is not None:
        saved = str(config.get("excel_path") or "").strip()
        if saved and os.path.exists(saved):
            return saved

    candidates = [
        app_dir() / "F.O.C(Mehr).xlsx",
        app_dir() / "F.O.C (Mehr).xlsx",
        app_dir() / "F.O.C-Mehr.xlsx",
        app_dir() / "mehr.xlsx",
        app_dir() / "F.O.C(Mehr).xlsm",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    if os.path.exists(ORIGINAL_EXCEL_PATH):
        return ORIGINAL_EXCEL_PATH
    return "" if config is None else str(config.get("excel_path") or "")


def default_pdf_folder(excel_path: str, config: Optional[AppConfig] = None) -> str:
    if config is not None:
        saved = str(config.get("pdf_folder") or "").strip()
        if saved:
            return saved
    if excel_path:
        beside_excel = Path(excel_path).parent / "PDF-MEHR"
        if beside_excel.is_dir():
            return str(beside_excel)
    local = app_dir() / "PDF-MEHR"
    if local.is_dir():
        return str(local)
    if os.path.isdir(ORIGINAL_PDF_FOLDER):
        return ORIGINAL_PDF_FOLDER
    return ""


# ---------------------------------------------------------------------------
# یادداشت‌ها و نشان‌ها (جایگزین noteهای سمت مرورگر)
# ---------------------------------------------------------------------------
class NotesStore:
    """ذخیرهٔ یادداشت‌ها و نشان‌ها در یک فایل JSON کنار برنامه."""

    FILENAME = "mehr_notes.json"

    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = Path(path) if path else (data_dir() / self.FILENAME)
        self.data: Dict[str, Dict[str, Any]] = {}
        self.load()

    def load(self) -> None:
        try:
            if self.path.exists():
                with open(self.path, "r", encoding="utf-8") as handle:
                    stored = json.load(handle)
                if isinstance(stored, dict):
                    self.data = stored
        except Exception as exc:
            print(f"هشدار: یادداشت‌ها خوانده نشد ({exc})")
            self.data = {}

    def save(self) -> None:
        try:
            tmp = self.path.with_suffix(".json.tmp")
            with open(tmp, "w", encoding="utf-8") as handle:
                json.dump(self.data, handle, ensure_ascii=False, indent=2)
            os.replace(tmp, self.path)
        except Exception as exc:
            print(f"هشدار: ذخیرهٔ یادداشت‌ها ممکن نشد ({exc})")

    def entry(self, key: str) -> Dict[str, Any]:
        item = self.data.get(key)
        if not isinstance(item, dict):
            item = {}
            self.data[key] = item
        item.setdefault("bookmarked", False)
        notes = item.get("notes")
        if not isinstance(notes, list):
            item["notes"] = []
        return item

    def notes_for(self, key: str) -> List[Dict[str, Any]]:
        return list(self.entry(key).get("notes") or [])

    def add_note(self, key: str, author: str, text: str, followup: str = "") -> Dict[str, Any]:
        entry = self.entry(key)
        note = {
            "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "author": author.strip(),
            "text": text.strip(),
            "followup": followup.strip(),
            "created": datetime.now().isoformat(timespec="seconds"),
        }
        entry.setdefault("notes", []).append(note)
        self.save()
        return note

    def delete_notes(self, key: str) -> None:
        if key in self.data:
            self.data[key]["notes"] = []
            if not self.data[key].get("bookmarked"):
                self.data.pop(key, None)
            self.save()

    def is_bookmarked(self, key: str) -> bool:
        item = self.data.get(key)
        return bool(item and item.get("bookmarked"))

    def toggle_bookmark(self, key: str) -> bool:
        entry = self.entry(key)
        entry["bookmarked"] = not bool(entry.get("bookmarked"))
        if not entry["bookmarked"] and not entry.get("notes"):
            self.data.pop(key, None)
        self.save()
        return bool(entry.get("bookmarked"))

    def bookmarked_keys(self) -> set:
        return {key for key, item in self.data.items() if isinstance(item, dict) and item.get("bookmarked")}

    def note_count(self, key: str) -> int:
        item = self.data.get(key)
        if not isinstance(item, dict):
            return 0
        notes = item.get("notes")
        return len(notes) if isinstance(notes, list) else 0


# ---------------------------------------------------------------------------
# خواندن اکسل و ساخت داده‌های داشبورد
# ---------------------------------------------------------------------------
class ExcelDataError(Exception):
    """خطای خواندن یا ساختار نامعتبر فایل اکسل."""


@dataclass
class ContractRecord:
    """یک قرارداد (یک ردیف داده در شیت colorcode)."""

    row_index: int
    excel_row: int
    col1: str
    col2: str
    col3: str
    col4: str
    deductions: str
    f_value: str
    as_status: str
    progress_status: str
    completion_status: str
    annex_status: str
    progress_table: str
    completion_table: str
    e_value: str
    aw: str
    ax: str
    ay: str
    pdf_file: str
    flags: Dict[str, bool]

    @property
    def key(self) -> str:
        return normalize_lookup_key(self.col3) or f"row-{self.excel_row}"

    @property
    def is_expired(self) -> bool:
        return "منقضی" in (self.annex_status or "")

    @property
    def open_flags(self) -> List[str]:
        return [FLAG_LABELS[key] for key in FLAG_LABELS if self.flags.get(key)]

    @property
    def has_deductions(self) -> bool:
        return bool(self.deductions.strip())

    def search_text(self) -> str:
        parts = [
            self.col1, self.col2, self.col3, self.col4, self.deductions,
            self.f_value, self.as_status, self.progress_status,
            self.completion_status, self.annex_status, self.aw, self.ax, self.ay,
        ]
        return normalize_per(" ".join(parts))

    def as_dict(self) -> Dict[str, Any]:
        data = {
            "col1": self.col1, "col2": self.col2, "col3": self.col3, "col4": self.col4,
            "deductions": self.deductions, "f_value": self.f_value,
            "as_status": self.as_status, "progress_status": self.progress_status,
            "completion_status": self.completion_status, "annex_status": self.annex_status,
            "progress_table_value": self.progress_table,
            "completion_table_value": self.completion_table,
            "e_value": self.e_value, "aw": self.aw, "ax": self.ax, "ay": self.ay,
            "pdf_link": self.pdf_file or None, "excel_row": self.excel_row,
        }
        for key in FLAG_LABELS:
            data[key] = bool(self.flags.get(key))
        return data


@dataclass
class Distribution:
    labels: List[str]
    values: List[int]

    @property
    def total(self) -> int:
        return int(sum(self.values))

    def as_list(self) -> List[Tuple[str, int]]:
        return list(zip(self.labels, self.values))


@dataclass
class DashboardData:
    excel_path: str
    sheet_name: str
    pdf_folder: str
    records: List[ContractRecord]
    progress_percent: float
    progress_percent_text: str
    progress_raw_text: str
    distributions: Dict[str, Distribution]
    top_types: List[Dict[str, Any]]
    progress_headers: List[str]
    header_labels: List[str]
    loaded_at: datetime
    source_mtime: Optional[datetime]
    source_size: int
    warnings: List[str]

    # ---- آمار کلی ----
    @property
    def total_contracts(self) -> int:
        return len(self.records)

    @property
    def contracts_with_deductions(self) -> int:
        return sum(1 for rec in self.records if rec.has_deductions)

    @property
    def flag_counts(self) -> Dict[str, int]:
        return {key: sum(1 for rec in self.records if rec.flags.get(key)) for key in FLAG_LABELS}

    @property
    def package_done(self) -> int:
        return sum(1 for rec in self.records if rec.as_status.strip() == "انجام شده")

    @property
    def expired_count(self) -> int:
        return sum(1 for rec in self.records if rec.is_expired)

    @property
    def source_mtime_text(self) -> str:
        return jalali_datetime(self.source_mtime) if self.source_mtime else "نامشخص"

    @property
    def loaded_at_text(self) -> str:
        return jalali_datetime(self.loaded_at)

    def dist(self, name: str) -> Distribution:
        return self.distributions.get(name, Distribution([], []))

    def status_dist(self) -> Distribution:
        if self.records:
            return Distribution(["انجام شده", "در جریان", "انجام نشده"], [
                sum(1 for r in self.records if r.as_status == "انجام شده"),
                sum(1 for r in self.records if r.as_status in ("در جریان", "در حال انجام")),
                sum(1 for r in self.records
                    if r.as_status not in ("انجام شده", "در جریان", "در حال انجام", "")),
            ])
        return self.dist("status")

    def empty(self) -> bool:
        return not self.records

    def find(self, key: str) -> Optional[ContractRecord]:
        for rec in self.records:
            if rec.key == key:
                return rec
        return None


def _cell_text(row: Sequence[Any], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return txt(row[index])


def _progress_like(value: Any) -> str:
    """سلول درصدی: اگر عدد باشد به درصد تبدیل می‌شود (مثل نسخهٔ وب)."""
    text = txt(value)
    if not text:
        return ""
    number = num(text)
    if number is not None:
        return fa(f"{number * 100:.2f}") + "٪"
    return text.replace("%", "٪")


def read_excel_workbook(
    path: str,
    sheet_name: str = SHEET_NAME_DEFAULT,
    pdf_folder: str = "",
) -> DashboardData:
    """خواندن اکسل و ساخت کامل داده‌های داشبورد بدون هیچ‌گونه کد وب."""
    if openpyxl is None:
        raise ExcelDataError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    if not path:
        raise ExcelDataError("مسیر فایل اکسل مشخص نشده است.")
    if not os.path.exists(path):
        raise ExcelDataError(f"فایل اکسل پیدا نشد:\n{path}")
    if str(path).lower().endswith(".xls"):
        raise ExcelDataError(
            "فرمت قدیمی ‎.xls‎ پشتیبانی نمی‌شود.\n"
            "فایل را در اکسل با فرمت ‎.xlsx‎ ذخیره کنید و دوباره انتخاب کنید."
        )

    warnings: List[str] = []
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except PermissionError as exc:
        raise ExcelDataError(
            f"دسترسی به فایل ممکن نیست (احتمالاً در حال استفاده است):\n{path}\n{exc}"
        ) from exc
    except Exception as exc:
        raise ExcelDataError(f"فایل اکسل باز نشد:\n{path}\n{exc}") from exc

    try:
        sheet_names = list(workbook.sheetnames)
        target = sheet_name if sheet_name in sheet_names else None
        if target is None:
            lowered = {name.strip().lower(): name for name in sheet_names}
            target = lowered.get((sheet_name or "").strip().lower())
        if target is None:
            sheet_label = sheet_name or SHEET_NAME_DEFAULT
            raise ExcelDataError(
                f"شیت «{sheet_label}» در فایل پیدا نشد.\n"
                f"شیت‌های موجود: {'، '.join(sheet_names)}"
            )
        sheet = workbook[target]
        raw_rows = [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        try:
            workbook.close()
        except Exception:
            pass

    if len(raw_rows) < 4:
        raise ExcelDataError("ساختار شیت غیرمنتظره است (کمتر از ۴ ردیف دارد).")

    source_width = max((len(row) for row in raw_rows), default=0)
    if source_width < 12:
        raise ExcelDataError(
            f"ساختار شیت نامعتبر است: فقط {source_width} ستون خوانده شد "
            f"(حداقل ۵۷ ستون انتظار می‌رفت)."
        )
    if source_width < REQUIRED_COLUMNS:
        warnings.append(
            f"شیت فقط {fa(source_width)} ستون دارد؛ ستون‌های ۵۶ و ۵۷ (وضعیت الحاقیه/پیشرفت) "
            f"ناقص خوانده می‌شوند."
        )
    width = max(source_width, REQUIRED_COLUMNS)
    for row in raw_rows:
        if len(row) < width:
            row.extend([None] * (width - len(row)))

    # ردیف ۲ (ایندکس ۱): درصد پیشرفت کل در ستون ۳۷
    progress_raw = raw_rows[1][36] if len(raw_rows) > 1 and len(raw_rows[1]) > 36 else None
    progress_number = num(progress_raw)
    if progress_number is not None:
        progress_percent = round(progress_number * 100, 2)
        progress_percent_text = fa(f"{progress_percent:.2f}")
    else:
        progress_percent = 0.0
        progress_percent_text = "—"
    progress_raw_text = txt(progress_raw)

    # ردیف ۳ (ایندکس ۲): عنوان ستون‌های کسورات (ستون‌های ۱۲ تا ۳۴)
    deduction_headers = [
        txt(raw_rows[2][index]) if index < width else ""
        for index in range(11, 34)
    ]
    progress_headers = [h for h in deduction_headers if h]

    # ردیف ۴ به بعد: داده‌ها
    records: List[ContractRecord] = []
    # ردیف‌های انتهایی کاملاً خالی حذف می‌شوند (مثل رفتار pandas در نسخهٔ وب)
    last_filled = 3
    for row_number in range(len(raw_rows) - 1, 2, -1):
        if any(not is_blank(value) for value in raw_rows[row_number][: REQUIRED_COLUMNS]):
            last_filled = row_number
            break

    for row_number in range(3, last_filled + 1):
        row = raw_rows[row_number]
        core_values = [row[i] if i < len(row) else None for i in range(width)]
        if all(is_blank(value) for value in core_values[: REQUIRED_COLUMNS]):
            continue  # ردیف کاملاً خالی

        deductions = " | ".join(
            deduction_headers[offset]
            for offset in range(23)
            if normalize_per(txt(row[11 + offset])).upper() == "M" and deduction_headers[offset]
        )

        flags = {
            key: normalize_per(txt(row[column] if column < len(row) else None)).upper() == "M"
            for key, _label, column in FLAG_DEFS
        }

        contract_number = _cell_text(row, 4)
        record = ContractRecord(
            row_index=len(records),
            excel_row=row_number + 1,
            col1=_cell_text(row, 0),
            col2=_cell_text(row, 3),
            col3=contract_number,
            col4=_cell_text(row, 6),
            deductions=deductions,
            f_value=_cell_text(row, 5),
            as_status=_cell_text(row, 44),
            progress_status=normalize_progress_value(row[53] if width > 53 else None) or "",
            completion_status=normalize_completion_value(row[54] if width > 54 else None) or "",
            annex_status=normalize_annex_value(row[55] if width > 55 else None) or "",
            progress_table=_progress_like(row[10] if width > 10 else None),
            completion_table=_progress_like(row[36] if width > 36 else None),
            e_value=_cell_text(row, 4),
            aw=_cell_text(row, 47).upper(),
            ax=_cell_text(row, 48).upper(),
            ay=_cell_text(row, 49).upper(),
            pdf_file="",
            flags=flags,
        )
        records.append(record)

    # اتصال فایل‌های PDF (همان منطق نسخهٔ وب: تطبیق شمارهٔ قرارداد با نام فایل)
    pdf_index = build_pdf_index(pdf_folder)
    matched_pdfs = 0
    for record in records:
        key = normalize_lookup_key(record.col3)
        filename = pdf_index.get(key) if key else None
        if filename:
            record.pdf_file = filename
            matched_pdfs += 1

    # ---- توزیع‌ها ----
    type_labels, type_values = value_counts([
        rec.f_value for rec in records if rec.f_value
    ])

    status_values = [rec.as_status for rec in records if rec.as_status]
    if not status_values:
        status_values = ["بدون وضعیت"]
    status_labels, status_count_values = value_counts(status_values)

    progress_values = [rec.progress_status for rec in records if rec.progress_status]
    if not progress_values:
        progress_values = ["بدون داده"]
    progress_labels, progress_count_values = value_counts(progress_values)

    completion_values = [rec.completion_status for rec in records if rec.completion_status]
    if not completion_values:
        completion_labels, completion_count_values = ["بدون داده"], [1]
    else:
        completion_labels, completion_count_values = value_counts(completion_values)

    annex_values = [rec.annex_status for rec in records if rec.annex_status]
    if annex_values:
        annex_labels, annex_count_values = value_counts(annex_values)
    else:
        annex_labels, annex_count_values = ["داده‌ای موجود نیست"], [1]

    # ۱۰ نوع قرارداد پرتکرار
    type_total = max(1, len([rec for rec in records if rec.f_value]))
    top_types = [
        {
            "label": label,
            "count": count,
            "percent": round(count / type_total * 100, 1),
        }
        for label, count in list(zip(type_labels, type_values))[:10]
    ]

    source_mtime = None
    source_size = 0
    try:
        stat = os.stat(path)
        source_mtime = datetime.fromtimestamp(stat.st_mtime)
        source_size = stat.st_size
    except Exception:
        pass

    if not records:
        warnings.append("هیچ ردیف داده‌ای در شیت پیدا نشد.")

    header_labels = [
        "ردیف", "نام شرکت", "شماره قرارداد", "موضوع قرارداد",
        "تعداد موارد باز", "کسورات اسناد اختتام", "نوع قرارداد",
        "پیشرفت کارکرد", "وضعیت پکیج آزادسازی", "الحاقیه",
    ]

    return DashboardData(
        excel_path=os.path.abspath(path),
        sheet_name=target,
        pdf_folder=pdf_folder,
        records=records,
        progress_percent=progress_percent,
        progress_percent_text=progress_percent_text,
        progress_raw_text=progress_raw_text,
        distributions={
            "type": Distribution(type_labels, type_values),
            "status": Distribution(status_labels, status_count_values),
            "progress": Distribution(progress_labels, progress_count_values),
            "completion": Distribution(completion_labels, completion_count_values),
            "annex": Distribution(annex_labels, annex_count_values),
        },
        top_types=top_types,
        progress_headers=progress_headers,
        header_labels=header_labels,
        loaded_at=datetime.now(),
        source_mtime=source_mtime,
        source_size=source_size,
        warnings=warnings,
    )


_PDF_INDEX_CACHE: Dict[str, Any] = {"folder": None, "mtime": None, "index": {}}


def build_pdf_index(folder: str, limit: int = 20000) -> Dict[str, str]:
    """ایندکس نام فایل‌های PDF بر اساس کلید نرمال‌شده (مثل نسخهٔ وب)."""
    index: Dict[str, str] = {}
    if not folder or not os.path.isdir(folder):
        return index
    try:
        folder_mtime = os.path.getmtime(folder)
    except OSError:
        folder_mtime = None
    cached = _PDF_INDEX_CACHE
    if cached["folder"] == folder and cached["mtime"] == folder_mtime and cached["index"]:
        return dict(cached["index"])  # type: ignore[arg-type]
    try:
        with os.scandir(folder) as entries:
            for count, entry in enumerate(entries):
                if count >= limit:
                    break
                try:
                    if not entry.is_file() or not entry.name.lower().endswith(".pdf"):
                        continue
                except OSError:
                    continue
                key = normalize_lookup_key(os.path.splitext(entry.name)[0])
                if key and key not in index:
                    index[key] = entry.name
    except Exception as exc:
        print(f"هشدار در ایندکس پوشهٔ PDF: {exc}")
    _PDF_INDEX_CACHE.update({"folder": folder, "mtime": folder_mtime, "index": dict(index)})
    return index


def empty_dashboard(message: str = "") -> DashboardData:
    """دادهٔ خالی برای زمانی که هنوز اکسل بارگذاری نشده است."""
    return DashboardData(
        excel_path="",
        sheet_name=SHEET_NAME_DEFAULT,
        pdf_folder="",
        records=[],
        progress_percent=0.0,
        progress_percent_text="—",
        progress_raw_text="",
        distributions={
            "type": Distribution([], []),
            "status": Distribution([], []),
            "progress": Distribution([], []),
            "completion": Distribution([], []),
            "annex": Distribution([], []),
        },
        top_types=[],
        progress_headers=[],
        header_labels=[
            "ردیف", "نام شرکت", "شماره قرارداد", "موضوع قرارداد",
            "تعداد موارد باز", "کسورات اسناد اختتام", "نوع قرارداد",
            "پیشرفت کارکرد", "وضعیت پکیج آزادسازی", "الحاقیه",
        ],
        loaded_at=datetime.now(),
        source_mtime=None,
        source_size=0,
        warnings=[message] if message else [],
    )


# ---------------------------------------------------------------------------
# پوسته و رنگ‌ها
# ---------------------------------------------------------------------------
THEMES: Dict[str, Dict[str, str]] = {
    "light": {
        "bg": "#eef2f0",
        "surface": "#ffffff",
        "surface_alt": "#f6faf8",
        "surface_soft": "#e8f0ed",
        "border": "#d9e2df",
        "border_soft": "#e7eeeb",
        "text": "#1b2724",
        "text_muted": "#66756f",
        "primary": "#0d6b63",
        "primary_dark": "#0a504a",
        "primary_soft": "#dcefe9",
        "accent": "#14b8a6",
        "sidebar": "#173a34",
        "sidebar_alt": "#1e4841",
        "sidebar_text": "#e6f4ef",
        "sidebar_muted": "#9ec9bd",
        "danger": "#b91c1c",
        "danger_soft": "#fde8e8",
        "warning": "#b45309",
        "warning_soft": "#fdf0d8",
        "success": "#0f766e",
        "success_soft": "#dcf5ec",
        "info": "#0e7490",
        "info_soft": "#dceff5",
        "header": "#ffffff",
        "shadow": "#9fb4ad",
        "row_alt": "#f7fbfa",
        "selection": "#d6ebe5",
        "chart_grid": "#dfe8e5",
    },
    "dark": {
        "bg": "#0f1614",
        "surface": "#16211e",
        "surface_alt": "#1a2723",
        "surface_soft": "#1f312c",
        "border": "#2a3b36",
        "border_soft": "#233430",
        "text": "#e9f2ef",
        "text_muted": "#9db2ab",
        "primary": "#2bb3a1",
        "primary_dark": "#1d8a7c",
        "primary_soft": "#1b3c36",
        "accent": "#5eead4",
        "sidebar": "#0b1412",
        "sidebar_alt": "#13211e",
        "sidebar_text": "#e3f2ee",
        "sidebar_muted": "#8fb5ab",
        "danger": "#f87171",
        "danger_soft": "#3a1d1d",
        "warning": "#fbbf24",
        "warning_soft": "#3a2e17",
        "success": "#34d399",
        "success_soft": "#17352c",
        "info": "#38bdf8",
        "info_soft": "#15313d",
        "header": "#10201c",
        "shadow": "#000000",
        "row_alt": "#182420",
        "selection": "#1f4a42",
        "chart_grid": "#263632",
    },
}

FONT_CANDIDATES = [
    "Vazirmatn",
    "Vazirmatn UI",
    "IRANSans",
    "IRANSansX",
    "Sahel",
    "Shabnam",
    "Noto Sans Arabic",
    "Segoe UI",
    "Tahoma",
    "DejaVu Sans",
]


_LOADED_FONTS: List[str] = []


def load_custom_fonts() -> List[str]:
    """بارگذاری فونت‌های فارسی کنار برنامه (اختیاری، مثلاً Vazirmatn)."""
    if _LOADED_FONTS:
        return _LOADED_FONTS
    for directory in (app_dir(), app_dir() / "fonts", data_dir() / "fonts"):
        try:
            if not directory.is_dir():
                continue
            for entry in sorted(directory.iterdir()):
                if entry.suffix.lower() in (".ttf", ".otf"):
                    font_id = QFontDatabase.addApplicationFont(str(entry))
                    if font_id >= 0:
                        _LOADED_FONTS.extend(QFontDatabase.applicationFontFamilies(font_id))
        except Exception:
            continue
    return _LOADED_FONTS


def pick_font_family() -> str:
    available = set(QFontDatabase().families())
    for name in FONT_CANDIDATES:
        if name in available:
            return name
    return QApplication.font().family() if QApplication.instance() else "Sans Serif"


def build_stylesheet(c: Dict[str, str], base_size: int = 10) -> str:
    """ساخت QSS برای کل برنامه بر اساس پالت انتخابی."""
    return f"""
    * {{ font-family: '{pick_font_family()}'; }}
    QWidget {{
        background: {c['bg']};
        color: {c['text']};
        font-size: {base_size}pt;
    }}
    QMainWindow, QDialog {{ background: {c['bg']}; }}
    QToolTip {{
        background: {c['surface']};
        color: {c['text']};
        border: 1px solid {c['border']};
        padding: 6px 10px;
        border-radius: 6px;
    }}
    QScrollArea, QScrollArea > QWidget > QWidget {{ background: transparent; }}
    QScrollBar:vertical {{
        background: transparent; width: 10px; margin: 2px;
    }}
    QScrollBar::handle:vertical {{
        background: {c['border']}; min-height: 30px; border-radius: 5px;
    }}
    QScrollBar::handle:vertical:hover {{ background: {c['primary']}; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
    QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
    QScrollBar::handle:horizontal {{ background: {c['border']}; min-width: 30px; border-radius: 5px; }}
    QScrollBar::handle:horizontal:hover {{ background: {c['primary']}; }}

    QFrame#Card {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 14px;
    }}
    QFrame#HeaderBar {{
        background: {c['surface']};
        border-bottom: 1px solid {c['border']};
    }}
    QFrame#Sidebar {{
        background: {c['sidebar']};
        border: none;
    }}
    QFrame#Banner {{
        background: {c['warning_soft']};
        border: 1px solid {c['warning']};
        border-radius: 10px;
    }}
    QFrame#Banner QLabel {{ color: {c['warning']}; font-weight: 600; }}
    QFrame#Banner QToolButton#BannerClose {{
        color: {c['warning']}; background: transparent; border: none;
        font-size: 15pt; font-weight: 700; padding: 0 6px;
    }}
    QFrame#Banner QToolButton#BannerClose:hover {{ color: {c['text']}; }}
    QPushButton#BannerAction {{
        background: {c['surface']}; color: {c['warning']};
        border: 1px solid {c['warning']}; border-radius: 8px;
        padding: 5px 12px; font-weight: 700;
    }}
    QPushButton#BannerAction:hover {{ background: {c['warning']}; color: #ffffff; }}
    QLabel {{ background: transparent; }}
    QLabel#Brand {{ color: {c['sidebar_text']}; font-weight: 800; font-size: 13pt; }}
    QLabel#BrandSub {{ color: {c['sidebar_muted']}; font-size: 8.5pt; }}
    QLabel#PageTitle {{ font-size: 15pt; font-weight: 800; color: {c['text']}; }}
    QLabel#SectionTitle {{ font-size: 11.5pt; font-weight: 800; color: {c['text']}; }}
    QLabel#Muted {{ color: {c['text_muted']}; }}
    QLabel#CardTitle {{ color: {c['text_muted']}; font-size: 9pt; font-weight: 600; }}
    QLabel#CardValue {{ color: {c['text']}; font-size: 19pt; font-weight: 800; }}
    QLabel#Chip {{
        background: {c['primary_soft']}; color: {c['primary_dark']};
        border-radius: 9px; padding: 3px 9px; font-size: 8.5pt; font-weight: 700;
    }}

    QPushButton {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 9px;
        padding: 7px 14px;
        color: {c['text']};
        font-weight: 600;
    }}
    QPushButton:hover {{ background: {c['surface_alt']}; border-color: {c['primary']}; }}
    QPushButton:pressed {{ background: {c['surface_soft']}; }}
    QPushButton:disabled {{ color: {c['text_muted']}; border-color: {c['border_soft']}; }}
    QPushButton#Primary {{
        background: {c['primary']}; color: #ffffff; border: 1px solid {c['primary']};
        font-weight: 700;
    }}
    QPushButton#Primary:hover {{ background: {c['primary_dark']}; }}
    QPushButton#Danger {{ color: {c['danger']}; border-color: {c['danger']}; }}
    QPushButton#Danger:hover {{ background: {c['danger_soft']}; }}
    QPushButton#NavButton {{
        background: transparent; border: none; color: {c['sidebar_text']};
        text-align: right; padding: 10px 14px; border-radius: 9px;
        font-weight: 600; font-size: 10pt;
    }}
    QPushButton#NavButton:hover {{ background: {c['sidebar_alt']}; }}
    QPushButton#NavButton:checked {{
        background: {c['sidebar_alt']}; color: #ffffff;
        border-right: 3px solid {c['accent']}; border-radius: 6px;
    }}
    QPushButton#NavSection {{
        background: transparent; border: none; color: {c['sidebar_muted']};
        text-align: right; padding: 4px 14px; font-size: 8pt; font-weight: 700;
        letter-spacing: 1px;
    }}
    QPushButton#ToolButton {{ padding: 6px 10px; border-radius: 8px; }}

    QLineEdit, QComboBox, QDateEdit, QPlainTextEdit, QTextEdit, QTextBrowser {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 9px;
        padding: 6px 10px;
        selection-background-color: {c['primary']};
        selection-color: #ffffff;
    }}
    QLineEdit:focus, QComboBox:focus, QDateEdit:focus,
    QPlainTextEdit:focus, QTextEdit:focus {{
        border: 1px solid {c['primary']};
    }}
    /* QSpinBox عمداً استایل نمی‌شود تا دکمه‌های بومی بال/پایین دیده شوند */
    QSpinBox {{
        padding: 4px 8px; font-size: {base_size}pt;
        selection-background-color: {c['primary']}; selection-color: #ffffff;
    }}
    QLineEdit#Search {{
        border-radius: 18px; padding: 8px 16px; background: {c['surface_alt']};
    }}
    QComboBox {{ padding: 6px 10px 6px 30px; min-height: 22px; }}
    QSpinBox {{ padding: 6px 10px; }}
    QCheckBox {{ spacing: 8px; }}
    QGroupBox {{
        background: {c['surface']}; border: 1px solid {c['border']};
        border-radius: 12px; margin-top: 16px; padding: 14px 12px 12px 12px;
        font-weight: 700;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin; subcontrol-position: top right;
        margin-right: 14px; padding: 0 6px; color: {c['primary']};
    }}

    QTableView {{
        background: {c['surface']};
        alternate-background-color: {c['row_alt']};
        border: 1px solid {c['border']};
        border-radius: 12px;
        gridline-color: {c['border_soft']};
        selection-background-color: {c['selection']};
        selection-color: {c['text']};
        outline: none;
    }}
    QTableView::item {{ padding: 4px 8px; border: none; }}
    QHeaderView::section {{
        background: {c['surface_alt']};
        color: {c['text_muted']};
        border: none;
        border-bottom: 1px solid {c['border']};
        border-left: 1px solid {c['border_soft']};
        padding: 9px 8px;
        font-weight: 700;
        font-size: 9pt;
    }}
    QTableCornerButton::section {{ background: {c['surface_alt']}; border: none; }}

    QTabWidget::pane {{
        border: 1px solid {c['border']}; border-radius: 12px; background: {c['surface']};
        top: -1px;
    }}
    QTabBar::tab {{
        background: transparent; color: {c['text_muted']};
        padding: 8px 16px; margin-left: 2px; border-radius: 8px;
        font-weight: 600;
    }}
    QTabBar::tab:selected {{ background: {c['primary_soft']}; color: {c['primary_dark']}; }}
    QTabBar::tab:hover {{ color: {c['primary']}; }}

    QListWidget {{
        background: {c['surface']}; border: 1px solid {c['border']}; border-radius: 10px;
        outline: none;
    }}
    QListWidget::item {{ padding: 8px 10px; border-radius: 8px; }}
    QListWidget::item:selected {{ background: {c['primary_soft']}; color: {c['text']}; }}
    QListWidget::item:hover {{ background: {c['surface_alt']}; }}

    QDockWidget {{ color: {c['text']}; titlebar-close-icon: none; }}
    QDockWidget::title {{
        background: {c['surface_alt']}; padding: 8px 12px; font-weight: 700;
        border-bottom: 1px solid {c['border']};
    }}
    QStatusBar {{ background: {c['surface']}; border-top: 1px solid {c['border']}; }}
    QStatusBar QLabel {{ color: {c['text_muted']}; padding: 0 8px; }}
    QProgressBar {{
        background: {c['surface_soft']}; border: none; border-radius: 5px;
        height: 10px; text-align: center; color: transparent;
    }}
    QProgressBar::chunk {{ background: {c['primary']}; border-radius: 5px; }}
    QMenuBar {{
        background: {c['surface']}; color: {c['text']}; padding: 2px 8px;
        border-bottom: 1px solid {c['border_soft']};
    }}
    QMenuBar::item {{ padding: 4px 10px; border-radius: 6px; }}
    QMenuBar::item:selected {{ background: {c['primary_soft']}; color: {c['primary_dark']}; }}
    QMenu {{
        background: {c['surface']}; border: 1px solid {c['border']}; border-radius: 8px; padding: 5px;
    }}
    QMenu::item {{ padding: 7px 24px 7px 16px; border-radius: 6px; }}
    QMenu::item:selected {{ background: {c['primary_soft']}; }}
    QSplitter::handle {{ background: {c['border_soft']}; }}
    """


def apply_theme(app: QApplication, theme: str, base_size: int) -> Dict[str, str]:
    colors = THEMES.get(theme, THEMES["light"])
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setLayoutDirection(Qt.RightToLeft)
    style = build_stylesheet(colors, base_size)
    app.setStyleSheet(style)

    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(colors["bg"]))
    palette.setColor(QPalette.WindowText, QColor(colors["text"]))
    palette.setColor(QPalette.Base, QColor(colors["surface"]))
    palette.setColor(QPalette.AlternateBase, QColor(colors["row_alt"]))
    palette.setColor(QPalette.Text, QColor(colors["text"]))
    palette.setColor(QPalette.Button, QColor(colors["surface"]))
    palette.setColor(QPalette.ButtonText, QColor(colors["text"]))
    palette.setColor(QPalette.Highlight, QColor(colors["primary"]))
    palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    palette.setColor(QPalette.ToolTipBase, QColor(colors["surface"]))
    palette.setColor(QPalette.ToolTipText, QColor(colors["text"]))
    app.setPalette(palette)
    return colors


# ---------------------------------------------------------------------------
# آیکون‌های برداری سبک (بدون فایل تصویری)
# ---------------------------------------------------------------------------
def draw_icon(painter: QPainter, kind: str, rect: QRectF, color: QColor) -> None:
    """رسم آیکون‌های سادهٔ برداری؛ مستقل از فونت و فایل‌های تصویری."""
    painter.save()
    painter.setRenderHint(QPainter.Antialiasing, True)
    pen_width = max(1.4, rect.width() * 0.085)
    pen = QPen(color, pen_width, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)
    r = rect.adjusted(pen_width, pen_width, -pen_width, -pen_width)
    cx, cy = r.center().x(), r.center().y()
    w, h = r.width(), r.height()

    if kind == "document":
        path = QPainterPath()
        path.moveTo(r.left() + w * 0.18, r.top())
        path.lineTo(r.left() + w * 0.66, r.top())
        path.lineTo(r.right(), r.top() + h * 0.26)
        path.lineTo(r.right(), r.bottom())
        path.lineTo(r.left() + w * 0.18, r.bottom())
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(QPointF(r.left() + w * 0.34, r.top() + h * 0.45),
                         QPointF(r.left() + w * 0.82, r.top() + h * 0.45))
        painter.drawLine(QPointF(r.left() + w * 0.34, r.top() + h * 0.65),
                         QPointF(r.left() + w * 0.82, r.top() + h * 0.65))
    elif kind == "scissors":
        painter.drawEllipse(QPointF(r.left() + w * 0.26, r.bottom() - h * 0.16), w * 0.15, h * 0.15)
        painter.drawEllipse(QPointF(r.left() + w * 0.74, r.bottom() - h * 0.16), w * 0.15, h * 0.15)
        painter.drawLine(QPointF(r.left() + w * 0.36, r.bottom() - h * 0.28),
                         QPointF(r.right() - w * 0.1, r.top()))
        painter.drawLine(QPointF(r.right() - w * 0.1, r.top()),
                         QPointF(r.left() + w * 0.64, r.bottom() - h * 0.28))
    elif kind == "shield":
        path = QPainterPath()
        path.moveTo(cx, r.top())
        path.lineTo(r.right() - w * 0.08, r.top() + h * 0.16)
        path.lineTo(r.right() - w * 0.14, r.top() + h * 0.62)
        path.quadTo(cx, r.bottom(), r.left() + w * 0.14, r.top() + h * 0.62)
        path.lineTo(r.left() + w * 0.08, r.top() + h * 0.16)
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(QPointF(cx - w * 0.16, cy + h * 0.02), QPointF(cx - w * 0.03, cy + h * 0.16))
        painter.drawLine(QPointF(cx - w * 0.03, cy + h * 0.16), QPointF(cx + w * 0.18, cy - h * 0.14))
    elif kind == "cube":
        top = QPainterPath()
        top.moveTo(cx, r.top())
        top.lineTo(r.right(), r.top() + h * 0.26)
        top.lineTo(cx, r.top() + h * 0.52)
        top.lineTo(r.left(), r.top() + h * 0.26)
        top.closeSubpath()
        painter.drawPath(top)
        painter.drawLine(QPointF(r.left(), r.top() + h * 0.26), QPointF(r.left(), r.bottom() - h * 0.2))
        painter.drawLine(QPointF(r.right(), r.top() + h * 0.26), QPointF(r.right(), r.bottom() - h * 0.2))
        painter.drawLine(QPointF(r.left(), r.bottom() - h * 0.2), QPointF(cx, r.bottom()))
        painter.drawLine(QPointF(r.right(), r.bottom() - h * 0.2), QPointF(cx, r.bottom()))
        painter.drawLine(QPointF(cx, r.top() + h * 0.52), QPointF(cx, r.bottom()))
    elif kind == "truck":
        body = QRectF(r.left(), r.top() + h * 0.24, w * 0.58, h * 0.42)
        painter.drawRect(body)
        cab = QPainterPath()
        cab.moveTo(body.right(), body.top() + h * 0.1)
        cab.lineTo(r.right() - w * 0.06, body.top() + h * 0.1)
        cab.lineTo(r.right(), body.bottom())
        cab.lineTo(body.right(), body.bottom())
        cab.closeSubpath()
        painter.drawPath(cab)
        painter.drawEllipse(QPointF(r.left() + w * 0.22, r.bottom() - h * 0.12), w * 0.11, h * 0.11)
        painter.drawEllipse(QPointF(r.right() - w * 0.2, r.bottom() - h * 0.12), w * 0.11, h * 0.11)
    elif kind == "check":
        painter.drawEllipse(QPointF(cx, cy), w * 0.46, h * 0.46)
        painter.drawLine(QPointF(cx - w * 0.24, cy + h * 0.02), QPointF(cx - w * 0.05, cy + h * 0.2))
        painter.drawLine(QPointF(cx - w * 0.05, cy + h * 0.2), QPointF(cx + w * 0.26, cy - h * 0.18))
    elif kind == "clock":
        painter.drawEllipse(QPointF(cx, cy), w * 0.44, h * 0.44)
        painter.drawLine(QPointF(cx, cy), QPointF(cx, cy - h * 0.24))
        painter.drawLine(QPointF(cx, cy), QPointF(cx + w * 0.19, cy + h * 0.1))
    elif kind == "star":
        path = QPainterPath()
        points = 5
        outer, inner = w * 0.48, w * 0.2
        for i in range(points * 2):
            angle = -math.pi / 2 + i * math.pi / points
            radius = outer if i % 2 == 0 else inner
            point = QPointF(cx + radius * math.cos(angle), cy + radius * math.sin(angle))
            if i == 0:
                path.moveTo(point)
            else:
                path.lineTo(point)
        path.closeSubpath()
        painter.drawPath(path)
    elif kind == "chart":
        painter.drawLine(QPointF(r.left() + w * 0.08, r.top()), QPointF(r.left() + w * 0.08, r.bottom()))
        painter.drawLine(QPointF(r.left() + w * 0.08, r.bottom()), QPointF(r.right(), r.bottom()))
        painter.drawRect(QRectF(r.left() + w * 0.22, r.bottom() - h * 0.36, w * 0.16, h * 0.36))
        painter.drawRect(QRectF(r.left() + w * 0.46, r.bottom() - h * 0.62, w * 0.16, h * 0.62))
        painter.drawRect(QRectF(r.left() + w * 0.7, r.bottom() - h * 0.9, w * 0.16, h * 0.9))
    elif kind == "bell":
        path = QPainterPath()
        path.moveTo(r.left() + w * 0.16, r.bottom() - h * 0.26)
        path.quadTo(r.left() + w * 0.24, r.top() + h * 0.16, cx, r.top() + h * 0.1)
        path.quadTo(r.right() - w * 0.24, r.top() + h * 0.16, r.right() - w * 0.16, r.bottom() - h * 0.26)
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(QPointF(cx, r.bottom() - h * 0.06), QPointF(cx, r.bottom() - h * 0.18))
    elif kind == "search":
        painter.drawEllipse(QPointF(cx + w * 0.06, cy - h * 0.06), w * 0.32, h * 0.32)
        painter.drawLine(QPointF(cx + w * 0.3, cy + h * 0.18), QPointF(r.right() - w * 0.02, r.bottom()))
    elif kind == "refresh":
        painter.drawArc(QRectF(r.left() + w * 0.1, r.top() + h * 0.1, w * 0.8, h * 0.8), 40 * 16, 280 * 16)
        painter.drawLine(QPointF(r.right() - w * 0.24, r.top() + h * 0.06), QPointF(r.right() - w * 0.06, r.top() + h * 0.2))
        painter.drawLine(QPointF(r.right() - w * 0.06, r.top() + h * 0.2), QPointF(r.right() - w * 0.22, r.top() + h * 0.3))
    elif kind in ("sun", "moon"):
        if kind == "sun":
            painter.drawEllipse(QPointF(cx, cy), w * 0.22, h * 0.22)
            for i in range(8):
                angle = i * math.pi / 4
                p1 = QPointF(cx + w * 0.33 * math.cos(angle), cy + h * 0.33 * math.sin(angle))
                p2 = QPointF(cx + w * 0.46 * math.cos(angle), cy + h * 0.46 * math.sin(angle))
                painter.drawLine(p1, p2)
        else:
            path = QPainterPath()
            path.addEllipse(QRectF(r.left() + w * 0.12, r.top() + h * 0.08, w * 0.76, h * 0.84))
            cut = QPainterPath()
            cut.addEllipse(QRectF(r.left() + w * 0.34, r.top(), w * 0.76, h * 0.84))
            painter.drawPath(path.subtracted(cut))
    elif kind == "download":
        painter.drawLine(QPointF(cx, r.top() + h * 0.06), QPointF(cx, r.bottom() - h * 0.4))
        painter.drawLine(QPointF(cx - w * 0.2, r.bottom() - h * 0.56), QPointF(cx, r.bottom() - h * 0.36))
        painter.drawLine(QPointF(cx + w * 0.2, r.bottom() - h * 0.56), QPointF(cx, r.bottom() - h * 0.36))
        painter.drawLine(QPointF(r.left() + w * 0.12, r.bottom() - h * 0.14), QPointF(r.right() - w * 0.12, r.bottom() - h * 0.14))
    elif kind == "gear":
        painter.drawEllipse(QPointF(cx, cy), w * 0.22, h * 0.22)
        painter.drawEllipse(QPointF(cx, cy), w * 0.38, h * 0.38)
        for i in range(8):
            angle = i * math.pi / 4
            p1 = QPointF(cx + w * 0.38 * math.cos(angle), cy + h * 0.38 * math.sin(angle))
            p2 = QPointF(cx + w * 0.48 * math.cos(angle), cy + h * 0.48 * math.sin(angle))
            painter.drawLine(p1, p2)
    elif kind == "help":
        painter.drawEllipse(QPointF(cx, cy), w * 0.45, h * 0.45)
        font = QFont()
        font.setPixelSize(int(rect.height() * 0.55))
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(r, Qt.AlignCenter, "?")
    elif kind in ("table", "grid"):
        painter.drawRoundedRect(r.adjusted(w * 0.04, h * 0.1, -w * 0.04, -h * 0.1), 2, 2)
        painter.drawLine(QPointF(r.left() + w * 0.04, cy), QPointF(r.right() - w * 0.04, cy))
        painter.drawLine(QPointF(cx - w * 0.1, r.top() + h * 0.1), QPointF(cx - w * 0.1, r.bottom() - h * 0.1))
    elif kind in ("pdf", "file-pdf"):
        painter.drawRoundedRect(QRectF(r.left() + w * 0.16, r.top(), w * 0.68, h), 3, 3)
        font = QFont()
        font.setPixelSize(max(6, int(rect.height() * 0.3)))
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(QRectF(r.left() + w * 0.16, r.top(), w * 0.68, h), Qt.AlignCenter, "PDF")
    elif kind == "bookmark":
        path = QPainterPath()
        path.moveTo(r.left() + w * 0.24, r.top())
        path.lineTo(r.right() - w * 0.24, r.top())
        path.lineTo(r.right() - w * 0.24, r.bottom())
        path.lineTo(cx, r.bottom() - h * 0.24)
        path.lineTo(r.left() + w * 0.24, r.bottom())
        path.closeSubpath()
        painter.drawPath(path)
    elif kind == "note":
        painter.drawRoundedRect(QRectF(r.left() + w * 0.08, r.top() + h * 0.14, w * 0.84, h * 0.72), 3, 3)
        for frac in (0.34, 0.52, 0.7):
            painter.drawLine(QPointF(r.left() + w * 0.24, r.top() + h * frac),
                             QPointF(r.right() - w * 0.24, r.top() + h * frac))
    elif kind == "folder":
        path = QPainterPath()
        path.moveTo(r.left() + w * 0.06, r.top() + h * 0.2)
        path.lineTo(r.left() + w * 0.44, r.top() + h * 0.2)
        path.lineTo(r.left() + w * 0.52, r.top() + h * 0.32)
        path.lineTo(r.right() - w * 0.06, r.top() + h * 0.32)
        path.lineTo(r.right() - w * 0.06, r.bottom() - h * 0.1)
        path.lineTo(r.left() + w * 0.06, r.bottom() - h * 0.1)
        path.closeSubpath()
        painter.drawPath(path)
    elif kind == "users":
        painter.drawEllipse(QPointF(cx - w * 0.14, r.top() + h * 0.3), w * 0.2, h * 0.18)
        painter.drawArc(QRectF(r.left() + w * 0.04, r.top() + h * 0.5, w * 0.62, h * 0.5), 0, 180 * 16)
        painter.drawEllipse(QPointF(cx + w * 0.32, r.top() + h * 0.34), w * 0.14, h * 0.13)
    elif kind == "alert":
        path = QPainterPath()
        path.moveTo(cx, r.top() + h * 0.08)
        path.lineTo(r.right(), r.bottom() - h * 0.06)
        path.lineTo(r.left(), r.bottom() - h * 0.06)
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(QPointF(cx, r.top() + h * 0.34), QPointF(cx, r.top() + h * 0.62))
        painter.drawPoint(QPointF(cx, r.top() + h * 0.74))
    elif kind == "close":
        painter.drawLine(r.topLeft(), r.bottomRight())
        painter.drawLine(r.topRight(), r.bottomLeft())
    elif kind == "menu":
        for frac in (0.2, 0.5, 0.8):
            painter.drawLine(QPointF(r.left(), r.top() + h * frac), QPointF(r.right(), r.top() + h * frac))
    elif kind == "lock":
        painter.drawRoundedRect(QRectF(r.left() + w * 0.14, r.top() + h * 0.42, w * 0.72, h * 0.5), 4, 4)
        painter.drawArc(QRectF(r.left() + w * 0.28, r.top() + h * 0.08, w * 0.44, h * 0.6), 0, 180 * 16)
    elif kind == "filter":
        path = QPainterPath()
        path.moveTo(r.left() + w * 0.05, r.top() + h * 0.14)
        path.lineTo(r.right() - w * 0.05, r.top() + h * 0.14)
        path.lineTo(cx + w * 0.12, cy + h * 0.04)
        path.lineTo(cx + w * 0.12, r.bottom() - h * 0.06)
        path.lineTo(cx - w * 0.12, r.bottom() - h * 0.2)
        path.lineTo(cx - w * 0.12, cy + h * 0.04)
        path.closeSubpath()
        painter.drawPath(path)
    else:
        painter.drawEllipse(QPointF(cx, cy), w * 0.4, h * 0.4)
        painter.drawLine(QPointF(cx - w * 0.16, cy), QPointF(cx + w * 0.16, cy))
    painter.restore()


class Combo(QComboBox):
    """کمبوی بومی با فلش رسم‌شده.

    Qt در حالت استایل‌شده و راست‌به‌چپ فلش پیش‌فرض را نشان نمی‌دهد و تصویر
    data-URI را هم در QSS بارگذاری نمی‌کند؛ بنابراین فلش را خودمان می‌کشیم.
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.arrow_color = QColor("#66756f")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setMinimumHeight(34)

    def set_arrow_color(self, color: str) -> None:
        self.arrow_color = QColor(color)
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        # پس‌زمینه و متن را با استایل معمولی می‌کشیم، سپس فلش را خودمان اضافه می‌کنیم
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        # پوشاندن فلش ناقص بومی در سمت چپ
        painter.setPen(Qt.NoPen)
        surface = self.palette().color(QPalette.Base)
        painter.setBrush(surface)
        painter.drawRect(QRectF(1.5, 1.5, 25, self.height() - 3))
        color = self.arrow_color
        if not self.isEnabled():
            color = QColor(color.red(), color.green(), color.blue(), 110)
        pen = QPen(color, 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        painter.setPen(pen)
        cx = 15.0
        cy = self.height() / 2.0
        path = QPainterPath()
        path.moveTo(cx - 4.5, cy - 2.0)
        path.lineTo(cx, cy + 2.6)
        path.lineTo(cx + 4.5, cy - 2.0)
        painter.drawPath(path)


class CheckBox(QCheckBox):
    """چک‌باکس بومی با تنظیم رنگ‌های پوسته (رسم توسط خود Qt انجام می‌شود)."""

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        # رنگ‌ها از پالت برنامه گرفته می‌شوند؛ نیازی به رسم دستی نیست.
        palette = self.palette()
        palette.setColor(QPalette.Base, QColor(colors["surface"]))
        palette.setColor(QPalette.Text, QColor(colors["text"]))
        palette.setColor(QPalette.WindowText, QColor(colors["text"]))
        palette.setColor(QPalette.Highlight, QColor(colors["primary"]))
        self.setPalette(palette)
        self.update()


class SpinBox(QSpinBox):
    """عدد‌شمار ساده؛ عمداً از QSS استایل نمی‌شود تا دکمه‌های بومی دیده شوند."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setButtonSymbols(QAbstractSpinBox.UpDownArrows)
        self.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.setMinimumHeight(32)


class IconWidget(QWidget):
    """نمایش یک آیکون برداری با رنگ دلخواه."""

    def __init__(self, kind: str, color: str = "#0d6b63", size: int = 20, parent=None) -> None:
        super().__init__(parent)
        self.kind = kind
        self.color = QColor(color)
        self._size = size
        self.setFixedSize(size, size)

    def set_color(self, color: str) -> None:
        self.color = QColor(color)
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        draw_icon(painter, self.kind, QRectF(0, 0, self.width(), self.height()), self.color)


class IconBadge(QWidget):
    """آیکون داخل یک مینیتور گرد رنگی."""

    def __init__(self, kind: str, color: str, background: str, size: int = 38, parent=None) -> None:
        super().__init__(parent)
        self.kind = kind
        self.color = QColor(color)
        self.background = QColor(background)
        self.setFixedSize(size, size)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        painter.setBrush(self.background)
        painter.drawRoundedRect(QRectF(0, 0, self.width(), self.height()), 11, 11)
        margin = self.width() * 0.24
        draw_icon(
            painter, self.kind,
            QRectF(margin, margin, self.width() - 2 * margin, self.height() - 2 * margin),
            self.color,
        )


class ClickableFrame(QFrame):
    """قاب قابل کلیک با حالت هاور."""

    clicked = pyqtSignal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._hover = False
        self.setCursor(Qt.PointingHandCursor)
        self.setAttribute(Qt.WA_Hover, True)

    def enterEvent(self, event) -> None:  # noqa: N802
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._hover = False
        self.update()
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit()
        super().mouseReleaseEvent(event)


# ---------------------------------------------------------------------------
# نمودارهای بومی (بدون Chart.js)
# ---------------------------------------------------------------------------
def _fmt_percent(part: int, total: int) -> str:
    if total <= 0:
        return "۰٪"
    return fa(f"{part / total * 100:.1f}%").replace(".0٪", "٪")


class PieChartWidget(QWidget):
    """نمودار دونات با تعامل ماوس و راهنمای داخلی."""

    sliceClicked = pyqtSignal(str)

    def __init__(self, colors: Sequence[str], parent=None) -> None:
        super().__init__(parent)
        self.labels: List[str] = []
        self.values: List[int] = []
        self.colors = [QColor(c) for c in colors]
        self.hover_index = -1
        self.center_caption = "قرارداد"
        self.text_color = QColor("#1b2724")
        self.muted_color = QColor("#66756f")
        self.surface_color = QColor("#ffffff")
        self.setMouseTracking(True)
        self.setMinimumSize(150, 150)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._slices: List[Tuple[float, float, str]] = []  # (شروع، اندازه، برچسب)

    def set_data(self, labels: Sequence[str], values: Sequence[int],
                 colors: Optional[Sequence[str]] = None,
                 center_caption: str = "قرارداد") -> None:
        self.labels = list(labels)
        self.values = [int(v) for v in values]
        if colors:
            self.colors = [QColor(c) for c in colors]
        self.center_caption = center_caption
        self.hover_index = -1
        self.update()

    def set_colors_for_theme(self, text: str, muted: str, surface: str) -> None:
        self.text_color = QColor(text)
        self.muted_color = QColor(muted)
        self.surface_color = QColor(surface)
        self.update()

    @property
    def total(self) -> int:
        return int(sum(self.values))

    def _color(self, index: int) -> QColor:
        if not self.colors:
            return QColor("#0d6b63")
        color = QColor(self.colors[index % len(self.colors)])
        # روشن‌تر کردن تدریجی برای تکرار رنگ‌ها
        cycle = index // len(self.colors)
        if cycle:
            color = color.lighter(100 + 12 * cycle)
        return color

    def _compute_slices(self) -> List[Tuple[float, float, str]]:
        total = self.total
        if total <= 0 or not self.values:
            return []
        result: List[Tuple[float, float]] = []
        start = 0.0
        for value in self.values:
            span = 360.0 * value / total
            result.append((start, span))
            start += span
        return [(s, sp, lab) for (s, sp), lab in zip(result, self.labels)]

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        side = min(self.width(), self.height())
        margin = 16
        box = QRectF(
            (self.width() - side) / 2 + margin,
            (self.height() - side) / 2 + margin,
            side - 2 * margin,
            side - 2 * margin,
        )
        outer = box.width() / 2
        thickness = max(20.0, outer * 0.30)
        inner = outer - thickness
        center = box.center()

        self._slices = self._compute_slices()
        if not self._slices:
            painter.setPen(QPen(self.muted_color))
            painter.drawText(self.rect(), Qt.AlignCenter, "داده‌ای برای نمایش نیست")
            return

        total = self.total
        for index, (start_deg, span_deg, _label) in enumerate(self._slices):
            color = self._color(index)
            hovered = index == self.hover_index
            if hovered:
                color = color.lighter(112)
            painter.setPen(QPen(self.surface_color, 1.5))
            painter.setBrush(QBrush(color))
            adjust = 4 if hovered else 0
            rect = QRectF(
                center.x() - outer + adjust, center.y() - outer + adjust,
                2 * (outer - adjust), 2 * (outer - adjust),
            )
            painter.drawPie(rect, int(round((90 - start_deg) * 16)), int(round(-span_deg * 16)))

        # حلقهٔ وسط
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(self.surface_color))
        painter.drawEllipse(center, inner, inner)

        # متن وسط
        if self.hover_index >= 0:
            value = self.values[self.hover_index]
            caption = self.labels[self.hover_index]
            big_text = fa_num(value)
            small_text = f"{caption} • {_fmt_percent(value, total)}"
            big_font = QFont(self.font())
            big_font.setPointSizeF(max(13.0, inner * 0.16))
            big_font.setBold(True)
            painter.setFont(big_font)
            painter.setPen(QPen(self.text_color))
            painter.drawText(
                QRectF(center.x() - inner, center.y() - inner * 0.55, inner * 2, inner * 0.6),
                Qt.AlignCenter, big_text,
            )
            small_font = QFont(self.font())
            small_font.setPointSizeF(max(7.5, inner * 0.075))
            painter.setFont(small_font)
            painter.setPen(QPen(self.muted_color))
            metrics = QFontMetrics(small_font)
            painter.drawText(
                QRectF(center.x() - inner, center.y() + inner * 0.08, inner * 2, inner * 0.4),
                Qt.AlignCenter,
                metrics.elidedText(small_text, Qt.ElideRight, int(inner * 1.7)),
            )
        else:
            big_font = QFont(self.font())
            big_font.setPointSizeF(max(16.0, inner * 0.22))
            big_font.setBold(True)
            painter.setFont(big_font)
            painter.setPen(QPen(self.text_color))
            painter.drawText(
                QRectF(center.x() - inner, center.y() - inner * 0.55, inner * 2, inner * 0.62),
                Qt.AlignCenter, fa_num(total),
            )
            small_font = QFont(self.font())
            small_font.setPointSizeF(max(7.5, inner * 0.08))
            painter.setFont(small_font)
            painter.setPen(QPen(self.muted_color))
            painter.drawText(
                QRectF(center.x() - inner, center.y() + inner * 0.1, inner * 2, inner * 0.4),
                Qt.AlignCenter, self.center_caption,
            )

    def _index_at(self, pos: QPointF) -> int:
        if not self._slices:
            return -1
        side = min(self.width(), self.height())
        margin = 16
        box = QRectF(
            (self.width() - side) / 2 + margin,
            (self.height() - side) / 2 + margin,
            side - 2 * margin,
            side - 2 * margin,
        )
        center = box.center()
        dx = pos.x() - center.x()
        dy = pos.y() - center.y()
        distance = math.hypot(dx, dy)
        outer = box.width() / 2
        inner = outer - max(20.0, outer * 0.30)
        if distance > outer or distance < inner:
            return -1
        angle = (math.degrees(math.atan2(dy, dx)) + 90) % 360
        for index, (start_deg, span_deg, _label) in enumerate(self._slices):
            if start_deg <= angle < start_deg + span_deg:
                return index
        return -1

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        index = self._index_at(event.pos())
        if index != self.hover_index:
            self.hover_index = index
            self.setCursor(Qt.PointingHandCursor if index >= 0 else Qt.ArrowCursor)
            self.update()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        if self.hover_index != -1:
            self.hover_index = -1
            self.update()
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton:
            index = self._index_at(event.pos())
            if index >= 0:
                self.sliceClicked.emit(self.labels[index])
        super().mouseReleaseEvent(event)


class ChartLegend(QWidget):
    """راهنمای نمودار؛ هر ردیف قابل کلیک است.

    برای خوانایی در عرض‌های کم، هر ردیف یک‌خطی رسم می‌شود: برچسب (راست)،
    درصد و در صورت وجود فضا تعداد (چپ). تعداد و درصد با مختصات مشخص رسم
    می‌شوند تا در متن راست‌به‌چپ جابه‌جا نشوند.
    """

    itemClicked = pyqtSignal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.items: List[Tuple[str, int, int]] = []  # (برچسب، تعداد، درصد)
        self.colors: List[QColor] = []
        self.text_color = QColor("#1b2724")
        self.muted_color = QColor("#66756f")
        self.hover = -1
        self.max_rows = 12
        self.setMouseTracking(True)
        self.setMinimumWidth(120)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    # ---------- داده و رنگ ----------
    def set_data(self, labels: Sequence[str], values: Sequence[int], colors: Sequence[str]) -> None:
        total = sum(int(v) for v in values)
        self.items = [
            (str(label), int(value), int(round(value / total * 100)) if total else 1)
            for label, value in zip(labels, values)
        ]
        self.colors = [QColor(c) for c in colors]
        self.hover = -1
        self.update()

    def set_colors_for_theme(self, text: str, muted: str) -> None:
        self.text_color = QColor(text)
        self.muted_color = QColor(muted)
        self.update()

    def _color(self, index: int) -> QColor:
        if not self.colors:
            return QColor("#0d6b63")
        color = QColor(self.colors[index % len(self.colors)])
        cycle = index // len(self.colors)
        return color.lighter(100 + 12 * cycle) if cycle else color

    # ---------- چیدمان ----------
    def _row_height(self) -> int:
        return 25

    def _rows(self) -> List[Tuple[str, int, int]]:
        fits = max(1, int((self.height() - 8) // self._row_height()))
        limit = min(self.max_rows, fits)
        if limit < len(self.items):
            limit = max(1, limit - 1)  # جای یک ردیف برای «و N مورد دیگر»
        return self.items[: max(1, limit)]

    def sizeHint(self) -> QSize:  # noqa: N802
        rows = max(1, len(self._rows()))
        return QSize(150, 6 + rows * self._row_height())

    def minimumSizeHint(self) -> QSize:  # noqa: N802
        return QSize(120, 40)

    # ---------- رسم ----------
    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        font = QFont(self.font())
        font.setPointSizeF(max(8.0, min(9.5, font.pointSizeF() - 0.3)))
        painter.setFont(font)
        metrics = QFontMetrics(font)
        rows = self._rows()
        row_height = self._row_height()
        top = 4
        width = self.width()
        show_count = width >= 168

        percent_texts = [fa(f"{percent}٪") for _l, _v, percent in rows]
        percent_width = max([metrics.horizontalAdvance(t) for t in percent_texts] or [0]) + 2
        count_widths = [metrics.horizontalAdvance(bidi(fa_num(value))) for _l, value, _p in rows]
        count_width = (max(count_widths) if count_widths else 0) + 6 if show_count else 0
        dot_space = 13
        gap = 7

        for index, (label, value, percent) in enumerate(rows):
            y = top + index * row_height
            if index == self.hover:
                painter.setPen(Qt.NoPen)
                painter.setBrush(QBrush(QColor(
                    self.muted_color.red(), self.muted_color.green(),
                    self.muted_color.blue(), 24,
                )))
                painter.drawRoundedRect(QRectF(0, y + 1, width, row_height - 2), 6, 6)

            painter.setPen(Qt.NoPen)
            painter.setBrush(QBrush(self._color(index)))
            painter.drawEllipse(QPointF(width - 5, y + row_height / 2), 4.5, 4.5)

            label_left = percent_width + count_width + gap
            label_width = max(16, int(width - label_left - dot_space - 4))
            painter.setFont(font)
            painter.setPen(QPen(self.text_color))
            painter.drawText(
                QRectF(label_left, y, label_width, row_height),
                Qt.AlignRight | Qt.AlignVCenter,
                metrics.elidedText(label, Qt.ElideRight, label_width),
            )
            painter.setPen(QPen(self.muted_color))
            painter.drawText(
                QRectF(0, y, percent_width, row_height),
                Qt.AlignLeft | Qt.AlignVCenter, percent_texts[index],
            )
            if show_count:
                painter.drawText(
                    QRectF(percent_width + gap - 4, y, count_width, row_height),
                    Qt.AlignLeft | Qt.AlignVCenter, bidi(fa_num(value)),
                )

        if len(self.items) > len(rows):
            painter.setPen(QPen(self.muted_color))
            painter.drawText(
                QRectF(0, top + len(rows) * row_height, width, row_height),
                Qt.AlignCenter,
                f"و {fa(len(self.items) - len(rows))} مورد دیگر",
            )

    # ---------- تعامل ----------
    def _row_at(self, y: int) -> int:
        row = (y - 4) // self._row_height()
        return int(row) if 0 <= row < len(self._rows()) else -1

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        row = self._row_at(event.pos().y())
        if row != self.hover:
            self.hover = row
            self.setCursor(Qt.PointingHandCursor if row >= 0 else Qt.ArrowCursor)
            if 0 <= row < len(self._rows()):
                label, value, percent = self._rows()[row]
                self.setToolTip(
                    f"{label}\n{fa_num(value)} قرارداد ({fa(percent)}٪) — کلیک برای فیلتر"
                )
            else:
                self.setToolTip("")
            self.update()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self.hover = -1
        self.update()
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        row = self._row_at(event.pos().y())
        if 0 <= row < len(self._rows()):
            self.itemClicked.emit(self._rows()[row][0])
        super().mouseReleaseEvent(event)


class BarChartWidget(QWidget):
    """نمودار میله‌ای افقی برای پرتکرارترین انواع قرارداد."""

    barClicked = pyqtSignal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.items: List[Dict[str, Any]] = []
        self.text_color = QColor("#1b2724")
        self.muted_color = QColor("#66756f")
        self.bar_color = QColor("#0d6b63")
        self.grid_color = QColor("#dfe8e5")
        self.hover = -1
        self.setMouseTracking(True)
        self.setMinimumHeight(150)

    def set_data(self, items: Sequence[Dict[str, Any]]) -> None:
        self.items = list(items)
        self.hover = -1
        self.update()

    def set_colors_for_theme(self, text: str, muted: str, bar: str, grid: str) -> None:
        self.text_color = QColor(text)
        self.muted_color = QColor(muted)
        self.bar_color = QColor(bar)
        self.grid_color = QColor(grid)
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        font = QFont(self.font())
        painter.setFont(font)
        metrics = QFontMetrics(font)
        if not self.items:
            painter.setPen(QPen(self.muted_color))
            painter.drawText(self.rect(), Qt.AlignCenter, "داده‌ای برای نمایش نیست")
            return

        row_height = int(max(20, min(38, (self.height() - 10) / max(1, len(self.items)))))
        max_count = max(1, max(int(item["count"]) for item in self.items))
        sample_value = bidi(f"{fa_num(max_count)} • {fa(100)}٪")
        value_width = metrics.horizontalAdvance(sample_value) + 14
        label_width = max(90, min(int(self.width() * 0.40), self.width() - value_width - 80))
        top = (self.height() - row_height * len(self.items)) / 2 + 4
        track_right = self.width() - label_width - 8
        track_left = value_width + 8
        track_width = max(40, track_right - track_left)

        for index, item in enumerate(self.items):
            y = top + index * row_height
            bar_y = y + row_height * 0.28
            bar_h = max(8, row_height * 0.42)
            ratio = item["count"] / max_count
            width = max(4, track_width * ratio)
            hovered = index == self.hover
            painter.setPen(Qt.NoPen)
            painter.setBrush(QBrush(QColor(self.grid_color.red(), self.grid_color.green(),
                                           self.grid_color.blue(), 90)))
            painter.drawRoundedRect(QRectF(track_right - track_width, bar_y, track_width, bar_h),
                                    bar_h / 2, bar_h / 2)

            color = QColor(self.bar_color)
            if hovered:
                color = color.lighter(112)
            gradient = QLinearGradient(track_right - width, 0, track_right, 0)
            gradient.setColorAt(0, color.lighter(135))
            gradient.setColorAt(1, color)
            painter.setBrush(QBrush(gradient))
            painter.drawRoundedRect(QRectF(track_right - width, bar_y, width, bar_h),
                                    bar_h / 2, bar_h / 2)

            painter.setPen(QPen(self.text_color))
            painter.drawText(
                QRectF(track_right + 8, y, label_width, row_height),
                Qt.AlignRight | Qt.AlignVCenter,
                metrics.elidedText(str(item["label"]), Qt.ElideMiddle, label_width),
            )
            painter.setPen(QPen(self.muted_color))
            value_text = bidi(f"{fa_num(item['count'])} • {fa(item['percent'])}٪")
            painter.drawText(
                QRectF(0, y, value_width, row_height),
                Qt.AlignLeft | Qt.AlignVCenter,
                metrics.elidedText(value_text, Qt.ElideRight, int(value_width)),
            )

    def _row_at(self, y: int) -> int:
        if not self.items:
            return -1
        row_height = int(max(20, min(38, (self.height() - 10) / max(1, len(self.items)))))
        top = (self.height() - row_height * len(self.items)) / 2 + 4
        row = int((y - top) // row_height)
        return row if 0 <= row < len(self.items) else -1

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        row = self._row_at(event.pos().y())
        if row != self.hover:
            self.hover = row
            self.setCursor(Qt.PointingHandCursor if row >= 0 else Qt.ArrowCursor)
            self.update()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self.hover = -1
        self.update()
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        row = self._row_at(event.pos().y())
        if 0 <= row < len(self.items):
            self.barClicked.emit(str(self.items[row]["label"]))
        super().mouseReleaseEvent(event)


class ProgressBarWidget(QWidget):
    """نوار پیشرفت با انیمیشن و مقدار درصد."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.value = 0.0
        self.display = 0.0
        self.track = QColor("#e3ebe8")
        self.fill = QColor("#0d6b63")
        self.text_color = QColor("#1b2724")
        self.muted = QColor("#66756f")
        self.setMinimumHeight(20)
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def set_value(self, value: float, animate: bool = True) -> None:
        self.value = max(0.0, min(100.0, float(value)))
        if not animate:
            self.display = self.value
            self.update()
            return
        self._timer.stop()
        self._timer.start(16)

    def _tick(self) -> None:
        delta = self.value - self.display
        if abs(delta) < 0.4:
            self.display = self.value
            self._timer.stop()
        else:
            self.display += delta * 0.18
        self.update()

    def set_colors_for_theme(self, track: str, fill: str, text: str, muted: str) -> None:
        self.track = QColor(track)
        self.fill = QColor(fill)
        self.text_color = QColor(text)
        self.muted = QColor(muted)
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        height = 10
        track = QRectF(0, (self.height() - height) / 2, self.width(), height)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(self.track))
        painter.drawRoundedRect(track, height / 2, height / 2)
        if self.display > 0:
            width = max(height, track.width() * self.display / 100.0)
            gradient = QLinearGradient(track.right() - width, 0, track.right(), 0)
            gradient.setColorAt(0, self.fill.lighter(128))
            gradient.setColorAt(1, self.fill)
            painter.setBrush(QBrush(gradient))
            painter.drawRoundedRect(
                QRectF(track.right() - width, track.top(), width, height),
                height / 2, height / 2,
            )
        painter.setPen(QPen(self.muted))
        font = QFont(self.font())
        font.setPointSizeF(max(7.5, font.pointSizeF() - 1))
        painter.setFont(font)
        painter.drawText(track, Qt.AlignCenter, fa(f"{self.value:.0f}%") if self.value else "")


# ---------------------------------------------------------------------------
# کارت‌های آماری و کارت نمودار
# ---------------------------------------------------------------------------
class StatCard(ClickableFrame):
    """کارت یک شاخص آماری؛ کلیک روی آن جدول را فیلتر می‌کند."""

    def __init__(self, title: str, icon: str, color: str, soft: str,
                 filter_key: Optional[str] = None, subtitle: str = "", parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        self.filter_key = filter_key
        self.default_sub = subtitle
        self._accent = QColor(color)
        self._soft = QColor(soft)
        self._hover = False
        self.setMinimumHeight(92)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(11)

        self.badge = IconBadge(icon, color, soft, 40)
        layout.addWidget(self.badge, 0, Qt.AlignTop)

        texts = QVBoxLayout()
        texts.setSpacing(1)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("CardTitle")
        self.title_label.setWordWrap(True)
        self.value_label = QLabel("—")
        self.value_label.setObjectName("CardValue")
        self.sub_label = QLabel(subtitle)
        self.sub_label.setObjectName("Muted")
        font = self.sub_label.font()
        font.setPointSizeF(max(7.5, font.pointSizeF() - 1.5))
        self.sub_label.setFont(font)
        texts.addWidget(self.title_label)
        texts.addWidget(self.value_label)
        texts.addWidget(self.sub_label)
        texts.addStretch(1)
        layout.addLayout(texts, 1)

    def set_value(self, value: Any, sub: str = "") -> None:
        text = fa_num(value) if isinstance(value, (int, float)) else fa(value)
        self.value_label.setText(bidi(text))
        caption = sub or self.default_sub
        self.sub_label.setText(bidi(caption))
        self.sub_label.setVisible(bool(caption))

    def set_accent(self, color: str, soft: str) -> None:
        self._accent = QColor(color)
        self._soft = QColor(soft)
        self.badge.color = QColor(color)
        self.badge.background = QColor(soft)
        self.badge.update()
        self.update()

    def enterEvent(self, event) -> None:  # noqa: N802
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._hover = False
        self.update()
        super().leaveEvent(event)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        painter.setPen(QPen(self._soft if not self._hover else self._accent, 1))
        painter.setBrush(QBrush(self._soft if self._hover else QColor(
            self.palette().color(QPalette.Base))))
        painter.drawRoundedRect(rect, 14, 14)
        # نوار رنگی سمت راست
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(self._accent))
        painter.drawRoundedRect(QRectF(rect.right() - 3, rect.top() + 14, 3.5,
                                       max(10.0, rect.height() - 28)), 2, 2)


class ChartCard(QFrame):
    """قاب نمودار: عنوان، نمودار و راهنما."""

    def __init__(self, title: str, subtitle: str, chart: QWidget, legend: QWidget,
                 parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 12)
        layout.setSpacing(6)

        header = QHBoxLayout()
        header.setSpacing(8)
        title_label = QLabel(title)
        title_label.setObjectName("SectionTitle")
        title_label.setMinimumWidth(40)
        title_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        header.addWidget(title_label, 1)
        header.addStretch(1)
        self.subtitle_label = QLabel(subtitle)
        self.subtitle_label.setObjectName("Chip")
        header.addWidget(self.subtitle_label, 0)
        layout.addLayout(header)

        body = QHBoxLayout()
        body.setSpacing(12)
        chart.setMaximumWidth(205)
        chart.setMinimumWidth(150)
        legend.setMinimumWidth(150)
        body.addWidget(chart, 0, Qt.AlignCenter)
        body.addWidget(legend, 1)
        layout.addLayout(body, 1)
        self.chart = chart
        self.legend = legend


class AlertBanner(QFrame):
    """نوار هشدار بالای صفحه."""

    actionClicked = pyqtSignal()
    closeClicked = pyqtSignal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("Banner")
        self.setVisible(False)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(10)
        self.icon = IconWidget("alert", "#b45309", 18)
        layout.addWidget(self.icon)
        self.label = QLabel("")
        self.label.setWordWrap(True)
        layout.addWidget(self.label, 1)
        self.action_button = QPushButton("مشاهده")
        self.action_button.setObjectName("BannerAction")
        self.action_button.clicked.connect(self.actionClicked.emit)
        layout.addWidget(self.action_button)
        self.close_button = QToolButton()
        self.close_button.setObjectName("BannerClose")
        self.close_button.setText("×")
        self.close_button.setToolTip("بستن هشدار")
        self.close_button.clicked.connect(self._hide)
        layout.addWidget(self.close_button)

    def _hide(self) -> None:
        self.setVisible(False)
        self.closeClicked.emit()

    def show_message(self, text: str, action_text: str = "") -> None:
        self.label.setText(text)
        self.action_button.setVisible(bool(action_text))
        self.action_button.setText(action_text)
        self.setVisible(True)


# ---------------------------------------------------------------------------
# مدل جدول قراردادها
# ---------------------------------------------------------------------------
@dataclass
class ColumnSpec:
    title: str
    width: int
    getter: Any
    align: int = int(Qt.AlignRight | Qt.AlignVCenter)
    role: str = "text"  # text | chip | pdf | bookmark | number


def _truncate(text: str, limit: int = 70) -> str:
    text = text.strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


CHIP_STYLES: Dict[str, Dict[str, Tuple[str, str]]] = {
    "annex": {
        "منقضی شده": ("#b91c1c", "#fde8e8"),
        "15 روز تا انقضا": ("#b45309", "#fdf0d8"),
        "30 روز تا انقضا": ("#92400e", "#fdf6e3"),
        "صورت وضعیت قطعی ثبت شده": ("#0f766e", "#dcf5ec"),
        "قرارداد جدید": ("#0e7490", "#dceff5"),
        "سنواتي": ("#6d28d9", "#ece7fb"),
        "مدت دارد": ("#475569", "#eaeef2"),
        "کمتر از 10 روز تا پایان": ("#b91c1c", "#fde8e8"),
        "کمتر از 30 روز تا پایان": ("#b45309", "#fdf0d8"),
    },
    "progress": {
        "0%": ("#b91c1c", "#fde8e8"),
        "کمتر از %75": ("#b45309", "#fdf0d8"),
        "100٪ - 75٪": ("#0f766e", "#dcf5ec"),
        "125%-100%": ("#1d4ed8", "#e0e9fd"),
        "بيشتر از %125": ("#6d28d9", "#ece7fb"),
    },
    "completion": {
        "0%": ("#b91c1c", "#fde8e8"),
        "کمتر از 90%": ("#b45309", "#fdf0d8"),
        "100%-90%": ("#0f766e", "#dcf5ec"),
        "100%": ("#15803d", "#dcfce7"),
    },
    "status": {
        "انجام شده": ("#15803d", "#dcfce7"),
        "در جریان": ("#b45309", "#fdf0d8"),
        "در حال انجام": ("#b45309", "#fdf0d8"),
        "انجام نشده": ("#b91c1c", "#fde8e8"),
    },
    "type": {},
    "count": {},
}


class ContractTableModel(QAbstractTableModel):
    """مدل جدول با فیلتر، جستجو و مرتب‌سازی داخلی."""

    COLUMNS: List[ColumnSpec] = [
        ColumnSpec("ردیف", 60, lambda r, i: str(i + 1), int(Qt.AlignCenter), "number"),
        ColumnSpec("نام شرکت", 210, lambda r, i: r.col2),
        ColumnSpec("شماره قرارداد", 140, lambda r, i: r.col3, int(Qt.AlignCenter), "code"),
        ColumnSpec("موضوع قرارداد", 300, lambda r, i: r.col4),
        ColumnSpec("نوع قرارداد", 150, lambda r, i: r.f_value, int(Qt.AlignCenter), "chip:type"),
        ColumnSpec("موارد باز", 112, lambda r, i: str(len(r.open_flags)), int(Qt.AlignCenter), "flags"),
        ColumnSpec("کسورات اسناد اختتام", 320, lambda r, i: r.deductions),
        ColumnSpec("پیشرفت کارکرد", 120, lambda r, i: r.progress_status, int(Qt.AlignCenter), "chip:progress"),
        ColumnSpec("الحاقیه", 155, lambda r, i: r.annex_status, int(Qt.AlignCenter), "chip:annex"),
        ColumnSpec("وضعیت پکیج", 125, lambda r, i: r.as_status, int(Qt.AlignCenter), "chip:status"),
        ColumnSpec("فایل", 80, lambda r, i: "PDF" if r.pdf_file else "", int(Qt.AlignCenter), "pdf"),
        ColumnSpec("نشان", 60, lambda r, i: "", int(Qt.AlignCenter), "bookmark"),
    ]

    def __init__(self, notes: NotesStore, parent=None) -> None:
        super().__init__(parent)
        self.notes = notes
        self._all: List[ContractRecord] = []
        self._rows: List[ContractRecord] = []
        self._search = ""
        self._filter_key = "all"
        self._bookmarks_only = False
        self._bookmarks: set = set()
        self._sort_column = -1
        self._sort_order = Qt.AscendingOrder
        self._colors: Dict[str, str] = dict(THEMES["light"])

    # ---- داده‌های پایه ----
    def set_records(self, records: Sequence[ContractRecord]) -> None:
        self.beginResetModel()
        self._all = list(records)
        self._bookmarks = self.notes.bookmarked_keys()
        self._apply()
        self.endResetModel()

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self._colors = colors
        self.dataChanged.emit(self.index(0, 0), self.index(max(0, len(self._rows) - 1),
                                                          len(self.COLUMNS) - 1))

    def refresh_bookmarks(self) -> None:
        self._bookmarks = self.notes.bookmarked_keys()
        if self._rows:
            top = self.index(0, 0)
            bottom = self.index(len(self._rows) - 1, len(self.COLUMNS) - 1)
            self.dataChanged.emit(top, bottom)
        if self._bookmarks_only:
            self.apply_view(self._search, self._filter_key, True)

    def all_records(self) -> List[ContractRecord]:
        return list(self._all)

    def visible_records(self) -> List[ContractRecord]:
        return list(self._rows)

    def record_at(self, row: int) -> Optional[ContractRecord]:
        if 0 <= row < len(self._rows):
            return self._rows[row]
        return None

    def source_count(self) -> int:
        return len(self._all)

    def visible_count(self) -> int:
        return len(self._rows)

    # ---- فیلتر و جستجو ----
    def _matches_search(self, record: ContractRecord, query: str) -> bool:
        if not query:
            return True
        haystack = record.search_text()
        for token in query.split():
            token = token.strip()
            if not token:
                continue
            if token not in haystack:
                return False
        return True

    def _matches_filter(self, record: ContractRecord) -> bool:
        key = self._filter_key
        if key in ("all", "", None):
            return True
        if key == "deductions":
            return record.has_deductions
        if key == "bookmarked":
            return record.key in self._bookmarks
        if key == "expired":
            return record.is_expired
        if key in FLAG_LABELS:
            return bool(record.flags.get(key))
        if ":" in key:
            field, _, value = key.partition(":")
            if field == "type":
                return record.f_value == value
            if field == "status":
                return record.as_status == value
            if field == "progress":
                return record.progress_status == value
            if field == "completion":
                return record.completion_status == value
            if field == "annex":
                return record.annex_status == value
            if field == "bookmarked":
                return record.key == value
        return True

    def _apply(self) -> None:
        query = normalize_per(self._search)
        rows = [rec for rec in self._all if self._matches_filter(rec) and self._matches_search(rec, query)]
        if self._bookmarks_only:
            rows = [rec for rec in rows if rec.key in self._bookmarks]
        if 0 <= self._sort_column < len(self.COLUMNS):
            spec = self.COLUMNS[self._sort_column]
            reverse = self._sort_order == Qt.DescendingOrder

            def sort_key(record: ContractRecord):
                if spec.role == "number":
                    return record.row_index
                if spec.role == "flags":
                    return len(record.open_flags)
                value = spec.getter(record, record.row_index)
                return normalize_per(value)

            try:
                rows.sort(key=sort_key, reverse=reverse)
            except Exception:
                pass
        self._rows = rows

    def apply_view(self, search: str = "", filter_key: str = "all",
                   bookmarks_only: bool = False) -> None:
        self.beginResetModel()
        self._search = search
        self._filter_key = filter_key
        self._bookmarks_only = bookmarks_only
        self._bookmarks = self.notes.bookmarked_keys()
        self._apply()
        self.endResetModel()

    def set_sort(self, column: int, order: int) -> None:
        if column < 0 or column >= len(self.COLUMNS):
            return
        self.beginResetModel()
        self._sort_column = column
        self._sort_order = order
        self._apply()
        self.endResetModel()

    # ---- API جدول ----
    def rowCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent=QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.COLUMNS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):  # noqa: N802
        if orientation == Qt.Horizontal and 0 <= section < len(self.COLUMNS):
            if role == Qt.DisplayRole:
                return self.COLUMNS[section].title
            if role == Qt.TextAlignmentRole:
                return int(Qt.AlignCenter)
        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return fa(section + 1)
        return None

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        record = self._rows[index.row()]
        spec = self.COLUMNS[index.column()]

        if role == Qt.DisplayRole:
            if spec.role == "code":
                return "\u200e" + fa(spec.getter(record, index.row()))
            if spec.role == "bookmark":
                return "★" if record.key in self._bookmarks else ""
            if spec.role == "flags":
                count = len(record.open_flags)
                return f"{fa(count)} مورد" if count else "بدون مورد"
            if spec.role == "number":
                return fa(index.row() + 1)
            return fa(spec.getter(record, index.row()))

        if role == Qt.TextAlignmentRole:
            return spec.align

        if role == Qt.UserRole:
            return record.key
        if role == Qt.UserRole + 1:
            return spec.role
        if role == Qt.UserRole + 2:
            if spec.role == "flags":
                return 1 if record.open_flags else 0
            if spec.role == "bookmark":
                return 1 if record.key in self._bookmarks else 0
            if spec.role.startswith("chip:"):
                return spec.role.split(":", 1)[1]
            return ""

        if role == Qt.ToolTipRole:
            if index.column() == 5:
                return "\n".join(record.open_flags) or "موردی باز نیست"
            if index.column() == 6:
                return record.deductions or "کسوراتی ثبت نشده است"
            if index.column() == 10:
                return f"فایل: {record.pdf_file}" if record.pdf_file else "فایل PDF یافت نشد"
            if index.column() == 9:
                return record.as_status or "بدون وضعیت"
            if index.column() == 7:
                return record.progress_status or "بدون داده"
            if index.column() == 8:
                return record.annex_status or "بدون داده"
            return str(spec.getter(record, index.row())) or None

        if role == Qt.ForegroundRole and spec.role == "number":
            return QBrush(QColor(self._colors.get("text_muted", "#66756f")))
        if role == Qt.FontRole and spec.role == "number":
            font = QFont()
            font.setBold(True)
            return font
        return None


class ContractDelegate(QStyledItemDelegate):
    """نمایندهٔ ترسیم ردیف‌ها: متن‌های معمولی و چیپ‌های وضعیت."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._colors: Dict[str, str] = dict(THEMES["light"])

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self._colors = colors

    def paint(self, painter, option, index) -> None:
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        colors = self._colors
        rect = option.rect
        selected = bool(option.state & QStyle.State_Selected)
        hovered = bool(option.state & QStyle.State_MouseOver)
        alternate = bool(option.features & QStyleOptionViewItem.Alternate)

        if selected:
            background = QColor(colors["selection"])
        elif hovered:
            background = QColor(colors["surface_soft"])
        elif alternate:
            background = QColor(colors["row_alt"])
        else:
            background = QColor(colors["surface"])
        painter.fillRect(rect, background)
        painter.setPen(QPen(QColor(colors["border_soft"])))
        painter.drawLine(rect.bottomLeft(), rect.bottomRight())

        text = index.data(Qt.DisplayRole)
        role = index.data(Qt.UserRole + 1) or ""
        chip_kind = index.data(Qt.UserRole + 2) or ""
        align = index.data(Qt.TextAlignmentRole) or int(Qt.AlignRight | Qt.AlignVCenter)
        if text is None:
            text = ""
        text = str(text)

        font = QFont(option.font)
        metrics = QFontMetrics(font)

        if role == "bookmark":
            if text:
                star_font = QFont(font)
                star_font.setPointSizeF(font.pointSizeF() + 3)
                painter.setFont(star_font)
                painter.setPen(QPen(QColor("#d97706")))
                painter.drawText(rect, Qt.AlignCenter, "★")
            painter.restore()
            return

        if role == "pdf" and text == "PDF":
            chip_font = QFont(font)
            chip_font.setBold(True)
            chip_font.setPointSizeF(max(7.0, font.pointSizeF() - 1.2))
            painter.setFont(chip_font)
            fm = QFontMetrics(chip_font)
            width = min(rect.width() - 10, fm.horizontalAdvance("PDF") + 20)
            chip = QRectF(
                rect.center().x() - width / 2, rect.center().y() - 9, width, 18,
            )
            painter.setPen(Qt.NoPen)
            painter.setBrush(QBrush(QColor(colors["danger_soft"])))
            painter.drawRoundedRect(chip, 6, 6)
            painter.setPen(QPen(QColor(colors["danger"])))
            painter.drawText(chip, Qt.AlignCenter, "PDF")
            painter.restore()
            return

        if role == "flags":
            count = int(index.data(Qt.UserRole + 2) or 0)
            label = text
            if count:
                fg, bg = QColor("#b45309"), QColor(colors["warning_soft"])
            else:
                fg, bg = QColor(colors["text_muted"]), QColor(colors["surface_soft"])
            self._draw_chip(painter, rect, label, fg, bg, font)
            painter.restore()
            return

        if role == "code":
            painter.setFont(font)
            painter.setPen(QPen(QColor(colors["text"])))
            text_rect = rect.adjusted(8, 0, -8, 0)
            painter.drawText(text_rect, Qt.AlignCenter,
                             metrics.elidedText(text, Qt.ElideRight, max(10, text_rect.width())))
            painter.restore()
            return

        if role.startswith("chip:") and text:
            style_map = CHIP_STYLES.get(chip_kind, {})
            fg_hex, bg_hex = style_map.get(text, ("", ""))
            if fg_hex:
                fg, bg = QColor(fg_hex), QColor(bg_hex)
            else:
                fg, bg = QColor(colors["primary_dark"]), QColor(colors["primary_soft"])
            self._draw_chip(painter, rect, text, fg, bg, font)
            painter.restore()
            return

        # متن معمولی
        painter.setFont(font)
        painter.setPen(QPen(QColor(colors["text"])))
        text_rect = rect.adjusted(8, 0, -8, 0)
        elided = metrics.elidedText(text, Qt.ElideRight, max(10, text_rect.width()))
        painter.drawText(text_rect, int(align), elided)
        painter.restore()

    @staticmethod
    def _draw_chip(painter, rect, text: str, fg: QColor, bg: QColor, font: QFont) -> None:
        chip_font = QFont(font)
        chip_font.setPointSizeF(max(7.0, font.pointSizeF() - 0.8))
        chip_font.setBold(True)
        painter.setFont(chip_font)
        metrics = QFontMetrics(chip_font)
        max_width = max(20, rect.width() - 14)
        label = metrics.elidedText(text, Qt.ElideRight, max_width - 18)
        width = min(max_width, metrics.horizontalAdvance(label) + 18)
        chip = QRectF(
            rect.center().x() - width / 2,
            rect.center().y() - 10,
            width, 20,
        )
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(bg))
        painter.drawRoundedRect(chip, 7, 7)
        painter.setPen(QPen(fg))
        painter.drawText(chip, Qt.AlignCenter, label)


class ExcelLoader(QThread):
    """بارگذاری اکسل در رشتهٔ جداگانه تا رابط کاربری قفل نشود."""

    loaded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, path: str, sheet: str, pdf_folder: str, parent=None) -> None:
        super().__init__(parent)
        self.path = path
        self.sheet = sheet
        self.pdf_folder = pdf_folder

    def run(self) -> None:
        try:
            data = read_excel_workbook(self.path, self.sheet, self.pdf_folder)
        except ExcelDataError as exc:
            self.failed.emit(str(exc))
            return
        except Exception as exc:  # خطای پیش‌بینی‌نشده
            detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.failed.emit(f"خطای پیش‌بینی‌نشده در خواندن اکسل:\n{detail}")
            return
        self.loaded.emit(data)


# ---------------------------------------------------------------------------
# چیدمان جریانی (برای کارت‌های واکنش‌گرا)
# ---------------------------------------------------------------------------
class FlowLayout(QLayout):
    """چیدمان کارت‌ها که با تغییر عرض پنجره می‌شکند."""

    def __init__(self, parent=None, margin: int = 0, spacing: int = 10) -> None:
        super().__init__(parent)
        self._items: List[Any] = []
        self._spacing = spacing
        self.setContentsMargins(margin, margin, margin, margin)

    def addItem(self, item) -> None:  # noqa: N802
        self._items.append(item)

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, index: int):  # noqa: N802
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index: int):  # noqa: N802
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self):  # noqa: N802
        return Qt.Orientations(Qt.Orientation(0))

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        return self._do_layout(QRectF(0, 0, width, 0), True)

    def setGeometry(self, rect) -> None:  # noqa: N802
        super().setGeometry(rect)
        self._do_layout(QRectF(rect), False)

    def sizeHint(self) -> QSize:  # noqa: N802
        return self.minimumSize()

    def minimumSize(self) -> QSize:  # noqa: N802
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        margins = self.contentsMargins()
        return size + QSize(margins.left() + margins.right(), margins.top() + margins.bottom())

    def _do_layout(self, rect: QRectF, test_only: bool) -> int:
        margins = self.contentsMargins()
        effective = rect.adjusted(margins.left(), margins.top(), -margins.right(), -margins.bottom())
        x, y = effective.x(), effective.y()
        line_height = 0
        for item in self._items:
            widget = item.widget()
            if widget is None or widget.isHidden():
                continue
            item_size = item.sizeHint()
            if x + item_size.width() > effective.right() + 1 and line_height > 0:
                x = effective.x()
                y = y + line_height + self._spacing
                line_height = 0
            if not test_only:
                item.setGeometry(QRectF(x, y, item_size.width(), item_size.height()).toRect())
            x = x + item_size.width() + self._spacing
            line_height = max(line_height, item_size.height())
        return int(y + line_height - rect.y() + margins.bottom())


def make_section_title(text: str, subtitle: str = "") -> QWidget:
    widget = QWidget()
    layout = QHBoxLayout(widget)
    layout.setContentsMargins(2, 6, 2, 0)
    layout.setSpacing(8)
    label = QLabel(text)
    label.setObjectName("SectionTitle")
    layout.addWidget(label)
    if subtitle:
        sub = QLabel(subtitle)
        sub.setObjectName("Muted")
        layout.addWidget(sub)
    layout.addStretch(1)
    return widget


# ---------------------------------------------------------------------------
# صفحهٔ داشبورد
# ---------------------------------------------------------------------------
CARD_SPECS: List[Tuple[str, str, str, Optional[str], str]] = [
    # (عنوان، آیکون، کلید رنگ، کلید فیلتر، زیرنویس)
    ("تعداد قرارداد", "document", "primary", "all", "کل قراردادهای ثبت‌شده"),
    ("کسورات اسناد اختتام", "scissors", "warning", "deductions", "دارای کسورات"),
    ("مفاصاحساب بیمه", "shield", "danger", "mofasa", "نیازمند پیگیری"),
    ("تسویهٔ مالی نهایی", "cube", "info", "financial", "نیازمند تسویه"),
    ("بالانس متریال", "chart", "primary", "balance", "نیازمند بالانس"),
    ("تحویل موقت", "clock", "warning", "temporary", "نیازمند تحویل"),
    ("تحویل قطعی", "truck", "info", "delivery", "نیازمند تحویل قطعی"),
    ("صورت‌وضعیت ماقبل قطعی", "document", "primary", "prefinal", "نیازمند صورت‌وضعیت"),
    ("صورت‌وضعیت قطعی", "check", "primary", "statement", "نیازمند صورت‌وضعیت"),
    ("اقرارنامه", "note", "warning", "affidavit", "نیازمند اقرارنامه"),
]

COLOR_SOFT = {
    "primary": ("primary", "primary_soft"),
    "danger": ("danger", "danger_soft"),
    "warning": ("warning", "warning_soft"),
    "info": ("info", "info_soft"),
    "success": ("success", "success_soft"),
}


class PreviewTableModel(ContractTableModel):
    """مدل جدول کوچک داشبورد (محدود به چند ردیف)."""

    def __init__(self, notes: NotesStore, limit: int = 12, parent=None) -> None:
        super().__init__(notes, parent)
        self.limit = limit
        self.filtered_total = 0

    def _apply(self) -> None:
        super()._apply()
        self.filtered_total = len(self._rows)
        self._rows = self._rows[: self.limit]


class StatCardsGrid(QWidget):
    """شبکهٔ واکنش‌گرای کارت‌های آماری (تعداد ستون‌ها با عرض پنجره تغییر می‌کند)."""

    def __init__(self, spacing: int = 10, parent=None) -> None:
        super().__init__(parent)
        self.grid = QGridLayout(self)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setSpacing(spacing)
        self.cards: List[QWidget] = []
        self.columns = 0

    def add_card(self, card: QWidget) -> None:
        self.cards.append(card)

    def relayout(self, columns: int) -> None:
        columns = max(1, min(6, int(columns)))
        if columns == self.columns:
            return
        self.columns = columns
        while self.grid.count():
            item = self.grid.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(self)
        for index, card in enumerate(self.cards):
            self.grid.addWidget(card, index // columns, index % columns)
        for column in range(columns):
            self.grid.setColumnStretch(column, 1)
        for column in range(columns, 6):
            self.grid.setColumnStretch(column, 0)

    def sizeHint(self) -> QSize:  # noqa: N802
        columns = max(1, self.columns)
        rows = (len(self.cards) + columns - 1) // columns
        height = rows * 92 + max(0, rows - 1) * 10
        return QSize(400, height)

    def minimumSizeHint(self) -> QSize:  # noqa: N802
        return QSize(200, 92)


class ResponsiveRow(QWidget):
    """ردیفی از کارت‌ها که در عرض کم، به‌صورت خودکار زیر هم می‌چیند."""

    def __init__(self, min_card_width: int = 260, max_columns: int = 3,
                 spacing: int = 14, parent=None) -> None:
        super().__init__(parent)
        self.grid = QGridLayout(self)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setSpacing(spacing)
        self.cards: List[QWidget] = []
        self.min_card_width = min_card_width
        self.max_columns = max_columns
        self.columns = 0

    def add_card(self, widget: QWidget) -> None:
        self.cards.append(widget)
        widget.setMinimumWidth(self.min_card_width)

    def relayout(self, available: int) -> None:
        if available <= 0:
            return
        columns = max(1, min(self.max_columns, available // self.min_card_width))
        columns = min(columns, max(1, len(self.cards)))
        if columns == self.columns:
            return
        self.columns = columns
        while self.grid.count():
            item = self.grid.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(self)
        for index, card in enumerate(self.cards):
            self.grid.addWidget(card, index // columns, index % columns)
        for column in range(max(self.max_columns, 1)):
            self.grid.setColumnStretch(column, 1 if column < columns else 0)

    def sizeHint(self) -> QSize:  # noqa: N802
        columns = max(1, self.columns)
        rows = (len(self.cards) + columns - 1) // columns
        return QSize(self.min_card_width * columns, rows * 300)


class DashboardPage(QWidget):
    """صفحهٔ اصلی: شاخص‌ها، نمودارها و پیش‌نمایش جدول."""

    cardClicked = pyqtSignal(str)
    sliceClicked = pyqtSignal(str, str)
    showContracts = pyqtSignal()
    refreshRequested = pyqtSignal()
    pickRequested = pyqtSignal()
    recordActivated = pyqtSignal(object)

    def __init__(self, notes: NotesStore, parent=None) -> None:
        super().__init__(parent)
        self.notes = notes
        self._colors = THEMES["light"]
        self._build_ui()

    # ---------- ساخت رابط ----------
    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        outer.addWidget(self.scroll)

        container = QWidget()
        self.container = container
        self.layout_main = QVBoxLayout(container)
        self.layout_main.setContentsMargins(22, 16, 22, 22)
        self.layout_main.setSpacing(14)
        self.scroll.setWidget(container)

        # عنوان صفحه
        title_row = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(0)
        self.title_label = QLabel("داشبورد قراردادها")
        self.title_label.setObjectName("PageTitle")
        self.subtitle_label = QLabel("پروژهٔ مهر خلیج فارس — گروه مهندسی واحد هزینه و درآمد")
        self.subtitle_label.setObjectName("Muted")
        self.subtitle_label.setWordWrap(True)
        title_box.addWidget(self.title_label)
        title_box.addWidget(self.subtitle_label)
        title_row.addLayout(title_box)
        title_row.addStretch(1)
        self.source_chip = QLabel("در حال بارگذاری…")
        self.source_chip.setObjectName("Chip")
        title_row.addWidget(self.source_chip, 0, Qt.AlignTop)
        self.layout_main.addLayout(title_row)

        # نوار هشدار
        self.banner = AlertBanner()
        self.banner.actionClicked.connect(self._banner_action)
        self._banner_target = "pick"
        self.layout_main.addWidget(self.banner)

        # کارت پیشرفت
        hero = QFrame()
        hero.setObjectName("Card")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(20, 16, 20, 16)
        hero_layout.setSpacing(18)
        hero_text = QVBoxLayout()
        hero_text.setSpacing(2)
        hero_title = QLabel("پیشرفت کل اختتام")
        hero_title.setObjectName("CardTitle")
        self.progress_value = QLabel("—")
        self.progress_value.setObjectName("CardValue")
        self.progress_value.setStyleSheet("font-size: 30pt;")
        self.progress_hint = QLabel("بر اساس ستون درصد پیشرفت شیت colorcode")
        self.progress_hint.setObjectName("Muted")
        hero_text.addWidget(hero_title)
        hero_text.addWidget(self.progress_value)
        hero_text.addWidget(self.progress_hint)
        hero_layout.addLayout(hero_text, 2)

        hero_progress = QVBoxLayout()
        hero_progress.setSpacing(6)
        self.progress_bar = ProgressBarWidget()
        self.progress_bar.setMinimumHeight(26)
        hero_progress.addWidget(self.progress_bar)
        info_row = QHBoxLayout()
        self.source_info = QLabel("—")
        self.source_info.setObjectName("Muted")
        self.source_info.setWordWrap(True)
        info_row.addWidget(self.source_info, 1)
        refresh_btn = QPushButton("به‌روزرسانی داده‌ها")
        refresh_btn.setObjectName("Primary")
        refresh_btn.clicked.connect(self.refreshRequested.emit)
        info_row.addWidget(refresh_btn, 0)
        hero_progress.addLayout(info_row)
        hero_layout.addLayout(hero_progress, 3)
        self.layout_main.addWidget(hero)

        # کارت‌های آماری
        self.cards: Dict[str, StatCard] = {}
        self.cards_holder = StatCardsGrid(10)
        self.layout_main.addWidget(self.cards_holder)
        for title, icon, color_key, filter_key, subtitle in CARD_SPECS:
            fg_key, bg_key = COLOR_SOFT[color_key]
            card = StatCard(title, icon, self._colors[fg_key], self._colors[bg_key],
                            filter_key, subtitle)
            card.clicked.connect(lambda key=filter_key: self.cardClicked.emit(key or "all"))
            self.cards[filter_key or "all"] = card
            self.cards_holder.add_card(card)
        self.cards_holder.relayout(5)

        # دو کارت برجسته: پکیج آزادسازی و الحاقیه
        prominent = ResponsiveRow(min_card_width=300, max_columns=2, spacing=14)
        self.status_pie = PieChartWidget(CHART_COLORS["status"], )
        self.status_legend = ChartLegend()
        self.status_card = ChartCard("وضعیت پکیج آزادسازی", "—", self.status_pie, self.status_legend)
        self.status_card.setFixedHeight(332)
        self.status_pie.sliceClicked.connect(lambda label: self.sliceClicked.emit("status", label))
        self.status_legend.itemClicked.connect(lambda label: self.sliceClicked.emit("status", label))
        prominent.add_card(self.status_card)

        self.annex_pie = PieChartWidget(CHART_COLORS["annex"])
        self.annex_legend = ChartLegend()
        self.annex_card = ChartCard("الحاقیه قراردادها", "—", self.annex_pie, self.annex_legend)
        self.annex_card.setFixedHeight(332)
        self.annex_pie.sliceClicked.connect(lambda label: self.sliceClicked.emit("annex", label))
        self.annex_legend.itemClicked.connect(lambda label: self.sliceClicked.emit("annex", label))
        prominent.add_card(self.annex_card)
        self.prominent_row = prominent
        self.layout_main.addWidget(prominent)

        # سه نمودار اصلی
        charts_row = ResponsiveRow(min_card_width=250, max_columns=3, spacing=14)
        self.type_pie = PieChartWidget(CHART_COLORS["type"])
        self.type_legend = ChartLegend()
        self.type_card = ChartCard("نوع قرارداد", "—", self.type_pie, self.type_legend)
        self.type_card.setFixedHeight(332)
        self.type_pie.sliceClicked.connect(lambda label: self.sliceClicked.emit("type", label))
        self.type_legend.itemClicked.connect(lambda label: self.sliceClicked.emit("type", label))
        charts_row.add_card(self.type_card)

        self.progress_pie = PieChartWidget(CHART_COLORS["progress"])
        self.progress_legend = ChartLegend()
        self.progress_chart_card = ChartCard("پیشرفت کارکرد", "—", self.progress_pie, self.progress_legend)
        self.progress_chart_card.setFixedHeight(332)
        self.progress_pie.sliceClicked.connect(lambda label: self.sliceClicked.emit("progress", label))
        self.progress_legend.itemClicked.connect(lambda label: self.sliceClicked.emit("progress", label))
        charts_row.add_card(self.progress_chart_card)

        self.completion_pie = PieChartWidget(CHART_COLORS["completion"])
        self.completion_legend = ChartLegend()
        self.completion_card = ChartCard("پیشرفت اختتام", "—", self.completion_pie, self.completion_legend)
        self.completion_card.setFixedHeight(332)
        self.completion_pie.sliceClicked.connect(lambda label: self.sliceClicked.emit("completion", label))
        self.completion_legend.itemClicked.connect(lambda label: self.sliceClicked.emit("completion", label))
        charts_row.add_card(self.completion_card)
        self.charts_row = charts_row
        self.layout_main.addWidget(charts_row)

        # نمودار میله‌ای ۱۰ نوع پرتکرار
        self.bar_card = QFrame()
        self.bar_card.setObjectName("Card")
        bar_layout = QVBoxLayout(self.bar_card)
        bar_layout.setContentsMargins(16, 14, 16, 14)
        bar_layout.setSpacing(6)
        bar_header = QHBoxLayout()
        bar_title = QLabel("۱۰ نوع قرارداد پرتکرار")
        bar_title.setObjectName("SectionTitle")
        bar_header.addWidget(bar_title)
        bar_header.addStretch(1)
        bar_hint = QLabel("کلیک روی هر میله جدول را فیلتر می‌کند")
        bar_hint.setObjectName("Muted")
        bar_header.addWidget(bar_hint)
        bar_layout.addLayout(bar_header)
        self.bar_chart = BarChartWidget()
        self.bar_chart.setMinimumHeight(300)
        self.bar_chart.setMaximumHeight(340)
        self.bar_chart.barClicked.connect(lambda label: self.sliceClicked.emit("type", label))
        bar_layout.addWidget(self.bar_chart)
        self.layout_main.addWidget(self.bar_card)

        # پیش‌نمایش جدول
        preview_card = QFrame()
        preview_card.setObjectName("Card")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(16, 14, 16, 14)
        preview_layout.setSpacing(8)
        header = QHBoxLayout()
        header_title = QLabel("قراردادهای نیازمند توجه")
        header_title.setObjectName("SectionTitle")
        header.addWidget(header_title)
        header.addStretch(1)
        self.preview_count = QLabel("—")
        self.preview_count.setObjectName("Chip")
        header.addWidget(self.preview_count)
        open_all = QPushButton("مشاهدهٔ جدول کامل")
        open_all.clicked.connect(self.showContracts.emit)
        header.addWidget(open_all)
        preview_layout.addLayout(header)
        self.preview_model = PreviewTableModel(notes=self.notes, limit=12)
        self.preview_table = QTableView()
        self.preview_table.setModel(self.preview_model)
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setShowGrid(False)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.preview_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.preview_table.horizontalHeader().setStretchLastSection(True)
        self.preview_table.verticalHeader().setDefaultSectionSize(34)
        self.preview_table.setMinimumHeight(12 * 34 + 38)
        self.preview_delegate = ContractDelegate()
        self.preview_table.setItemDelegate(self.preview_delegate)
        self._configure_preview_columns()
        preview_layout.addWidget(self.preview_table)
        self.layout_main.addWidget(preview_card)

        self.layout_main.addStretch(1)

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self._update_card_widths()

    def _update_card_widths(self) -> None:
        if not hasattr(self, "cards_holder"):
            return
        viewport = self.scroll.viewport()
        available = viewport.width() - 44  # حاشیهٔ چپ/راست صفحه
        if available < 200:
            return
        columns = max(2, min(5, available // 232))
        self.cards_holder.relayout(columns)
        if hasattr(self, "prominent_row"):
            self.prominent_row.relayout(available)
        if hasattr(self, "charts_row"):
            self.charts_row.relayout(available)

    def showEvent(self, event) -> None:  # noqa: N802
        super().showEvent(event)
        QTimer.singleShot(0, self._update_card_widths)

    def _configure_preview_columns(self) -> None:
        hidden = {4, 6, 7, 8, 10, 11}  # فقط ستون‌های کلیدی در پیش‌نمایش
        for column in range(len(ContractTableModel.COLUMNS)):
            self.preview_table.setColumnHidden(column, column in hidden)
        widths = {0: 54, 1: 220, 2: 130, 3: 330, 5: 100, 9: 130}
        for column, width in widths.items():
            self.preview_table.setColumnWidth(column, width)
        self.preview_table.setMouseTracking(True)
        self.preview_table.doubleClicked.connect(self._preview_double_clicked)

    def _preview_double_clicked(self, index) -> None:
        record = self.preview_model.record_at(index.row())
        if record:
            self.recordActivated.emit(record)

    # ---------- به‌روزرسانی ----------
    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self._colors = colors
        for title, icon, color_key, filter_key, subtitle in CARD_SPECS:
            fg_key, bg_key = COLOR_SOFT[color_key]
            card = self.cards.get(filter_key or "all")
            if card:
                card.set_accent(colors[fg_key], colors[bg_key])
        for chart in (self.status_pie, self.annex_pie, self.type_pie,
                      self.progress_pie, self.completion_pie):
            chart.set_colors_for_theme(colors["text"], colors["text_muted"], colors["surface"])
        for legend in (self.status_legend, self.annex_legend, self.type_legend,
                       self.progress_legend, self.completion_legend):
            legend.set_colors_for_theme(colors["text"], colors["text_muted"])
        self.bar_chart.set_colors_for_theme(
            colors["text"], colors["text_muted"], colors["primary"], colors["chart_grid"]
        )
        self.progress_bar.set_colors_for_theme(
            colors["surface_soft"], colors["primary"], colors["text"], colors["text_muted"]
        )
        self.preview_delegate.set_theme_colors(colors)
        self.preview_model.set_theme_colors(colors)

    def update_data(self, data: DashboardData) -> None:
        self.progress_value.setText(f"{data.progress_percent_text}٪"
                                    if data.progress_percent_text != "—" else "—")
        self.progress_bar.set_value(data.progress_percent)
        self.source_info.setText(
            f"منبع داده: {os.path.basename(data.excel_path) if data.excel_path else 'انتخاب نشده'}"
            f"   •   شیت: {data.sheet_name}"
            f"   •   آخرین تغییر فایل: {data.source_mtime_text}"
            f"   •   حجم: {human_size(data.source_size)}"
        )
        self.source_chip.setText(
            f"بارگذاری: {data.loaded_at_text}" if data.records else "داده‌ای بارگذاری نشده"
        )
        counts = data.flag_counts
        self.cards["all"].set_value(
            data.total_contracts,
            f"{fa(data.flag_counts['temporary'])} تحویل موقت در جریان",
        )
        self.cards["deductions"].set_value(
            data.contracts_with_deductions,
            f"از {fa(data.total_contracts)} قرارداد",
        )
        for key, _label, _column in FLAG_DEFS:
            card = self.cards.get(key)
            if card:
                card.set_value(counts.get(key, 0))

        self.status_card.subtitle_label.setText(
            f"{fa(data.dist('status').total)} قرارداد"
        )
        self.status_pie.set_data(data.dist('status').labels, data.dist('status').values,
                                 CHART_COLORS["status"], "قرارداد")
        self.status_legend.set_data(data.dist('status').labels, data.dist('status').values,
                                    CHART_COLORS["status"])
        self.annex_card.subtitle_label.setText(f"{fa(data.dist('annex').total)} قرارداد")
        self.annex_pie.set_data(data.dist("annex").labels, data.dist("annex").values,
                                CHART_COLORS["annex"], "قرارداد")
        self.annex_legend.set_data(data.dist("annex").labels, data.dist("annex").values,
                                   CHART_COLORS["annex"])
        self.type_card.subtitle_label.setText(f"{fa(data.dist('type').total)} قرارداد")
        self.type_pie.set_data(data.dist("type").labels, data.dist("type").values,
                               CHART_COLORS["type"], "قرارداد")
        self.type_legend.set_data(data.dist("type").labels, data.dist("type").values,
                                  CHART_COLORS["type"])
        self.progress_chart_card.subtitle_label.setText(f"{fa(data.dist('progress').total)} قرارداد")
        self.progress_pie.set_data(data.dist("progress").labels, data.dist("progress").values,
                                   CHART_COLORS["progress"], "قرارداد")
        self.progress_legend.set_data(data.dist("progress").labels, data.dist("progress").values,
                                      CHART_COLORS["progress"])
        self.completion_card.subtitle_label.setText(f"{fa(data.dist('completion').total)} قرارداد")
        self.completion_pie.set_data(data.dist("completion").labels, data.dist("completion").values,
                                     CHART_COLORS["completion"], "قرارداد")
        self.completion_legend.set_data(data.dist("completion").labels,
                                        data.dist("completion").values,
                                        CHART_COLORS["completion"])
        self.bar_chart.set_data(data.top_types)
        self.preview_model.set_records(data.records)

        self._update_banner(data)
        self._refresh_preview()

    def _update_banner(self, data: DashboardData) -> None:
        if data.warnings:
            self._banner_target = "pick"
            self.banner.show_message(" • ".join(data.warnings), "انتخاب فایل اکسل")
            return
        if not data.records:
            self._banner_target = "pick"
            self.banner.show_message("داده‌ای بارگذاری نشده است. فایل اکسل را انتخاب کنید.",
                                     "انتخاب فایل")
            return
        if data.expired_count:
            self._banner_target = "expired"
            self.banner.show_message(
                f"{fa(data.expired_count)} قرارداد در وضعیت «منقضی شده» است.",
                "نمایش این قراردادها",
            )
            return
        if not data.pdf_folder or not os.path.isdir(data.pdf_folder):
            self._banner_target = "settings"
            self.banner.show_message(
                "پوشهٔ PDF تنظیم نشده است؛ ستون فایل قرارداد خالی می‌ماند.", "تنظیم پوشه"
            )
            return
        self.banner.setVisible(False)

    def _banner_action(self) -> None:
        if self._banner_target == "expired":
            self.cardClicked.emit("expired")
        elif self._banner_target == "settings":
            self.refreshRequested.emit()
        else:
            self.pickRequested.emit()

    def _refresh_preview(self) -> None:
        self.preview_count.setText(bidi(
            f"{fa(self.preview_model.filtered_total)} از "
            f"{fa(self.preview_model.source_count())} قرارداد"
        ))

    def apply_filter(self, search: str, filter_key: str, bookmarks_only: bool) -> None:
        self.preview_model.apply_view(search, filter_key, bookmarks_only)
        self._refresh_preview()

    def refresh_bookmarks(self) -> None:
        self.preview_model.refresh_bookmarks()



# ---------------------------------------------------------------------------
# صفحهٔ جدول قراردادها
# ---------------------------------------------------------------------------
class ContractsPage(QWidget):
    """جدول کامل با فیلتر، جستجو، مرتب‌سازی و خروجی."""

    recordActivated = pyqtSignal(object)
    exportRequested = pyqtSignal()
    bookmarkToggled = pyqtSignal(str)
    pdfRequested = pyqtSignal(str)
    copyRequested = pyqtSignal(str)
    filterChanged = pyqtSignal(str)
    searchChanged = pyqtSignal(str)

    def __init__(self, notes: NotesStore, parent=None) -> None:
        super().__init__(parent)
        self.notes = notes
        self._colors = THEMES["light"]
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 18)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(0)
        title = QLabel("جدول قراردادها")
        title.setObjectName("PageTitle")
        subtitle = QLabel("جستجو، فیلتر و خروجی اکسل از داده‌های همان شیت colorcode")
        subtitle.setObjectName("Muted")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        title_row.addLayout(title_box)
        title_row.addStretch(1)
        self.count_chip = QLabel("—")
        self.count_chip.setObjectName("Chip")
        title_row.addWidget(self.count_chip, 0, Qt.AlignTop)
        layout.addLayout(title_row)

        toolbar = QFrame()
        toolbar.setObjectName("Card")
        toolbar_outer = QVBoxLayout(toolbar)
        toolbar_outer.setContentsMargins(12, 10, 12, 10)
        toolbar_outer.setSpacing(8)
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(8)
        toolbar_outer.addLayout(toolbar_layout)

        self.search_edit = QLineEdit()
        self.search_edit.setObjectName("Search")
        self.search_edit.setPlaceholderText("جستجو در شمارهٔ قرارداد، پیمانکار، موضوع و کسورات…")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMinimumWidth(240)
        self.search_edit.textChanged.connect(self.searchChanged.emit)
        toolbar_layout.addWidget(self.search_edit, 2)

        self.filter_combo = Combo()
        self.filter_combo.setMinimumWidth(230)
        for key, label in FILTER_CHOICES:
            self.filter_combo.addItem(label, key)
        self.filter_combo.currentIndexChanged.connect(self._on_filter_changed)
        toolbar_layout.addWidget(self.filter_combo, 1)

        self.bookmark_check = CheckBox("فقط نشان‌شده‌ها")
        self.bookmark_check.toggled.connect(lambda _: self.filterChanged.emit(self.current_filter()))
        toolbar_layout.addWidget(self.bookmark_check)

        toolbar_layout.addStretch(1)
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(8)
        toolbar_outer.addLayout(toolbar_layout)

        self.columns_button = QToolButton()
        self.columns_button.setObjectName("ToolButton")
        self.columns_button.setText("ستون‌ها")
        self.columns_button.setPopupMode(QToolButton.InstantPopup)
        self.columns_menu = QMenu(self.columns_button)
        self.column_actions: List[QAction] = []
        for index, spec in enumerate(ContractTableModel.COLUMNS):
            action = QAction(spec.title, self.columns_menu)
            action.setCheckable(True)
            action.setChecked(True)
            action.toggled.connect(lambda checked, col=index: self._set_column_visible(col, checked))
            self.columns_menu.addAction(action)
            self.column_actions.append(action)
        self.columns_button.setMenu(self.columns_menu)
        toolbar_layout.addWidget(self.columns_button)

        clear_btn = QPushButton("پاک کردن فیلترها")
        clear_btn.clicked.connect(self.clear_filters)
        toolbar_layout.addWidget(clear_btn)
        toolbar_layout.addStretch(1)

        export_btn = QPushButton("خروجی اکسل")
        export_btn.setObjectName("Primary")
        export_btn.clicked.connect(self.exportRequested.emit)
        toolbar_layout.addWidget(export_btn)
        layout.addWidget(toolbar)

        self.table = QTableView()
        self.model = ContractTableModel(self.notes)
        self.table.setModel(self.model)
        self.delegate = ContractDelegate()
        self.table.setItemDelegate(self.delegate)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setMouseTracking(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSortingEnabled(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(36)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(False)
        header.setHighlightSections(False)
        header.setSortIndicatorShown(True)
        header.sectionClicked.connect(self._on_header_clicked)
        for index, spec in enumerate(ContractTableModel.COLUMNS):
            self.table.setColumnWidth(index, spec.width)
        self.table.doubleClicked.connect(self._on_double_click)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.selectionModel().selectionChanged.connect(self._update_count)
        layout.addWidget(self.table, 1)

        footer = QHBoxLayout()
        self.footer_label = QLabel("—")
        self.footer_label.setObjectName("Muted")
        footer.addWidget(self.footer_label)
        footer.addStretch(1)
        hint = QLabel("دو کلیک روی هر ردیف: جزئیات و یادداشت   •   Ctrl+E: خروجی اکسل")
        hint.setObjectName("Muted")
        footer.addWidget(hint)
        layout.addLayout(footer)

    # ---- رخدادها ----
    def _on_filter_changed(self, _index: int) -> None:
        self.filterChanged.emit(self.current_filter())

    def _on_header_clicked(self, section: int) -> None:
        header = self.table.horizontalHeader()
        if header.sortIndicatorSection() == section:
            order = (Qt.DescendingOrder if header.sortIndicatorOrder() == Qt.AscendingOrder
                     else Qt.AscendingOrder)
        else:
            order = Qt.AscendingOrder
        header.setSortIndicator(section, order)
        self.model.set_sort(section, order)
        self._update_count()

    def _on_double_click(self, index: QModelIndex) -> None:
        record = self.model.record_at(index.row())
        if record:
            self.recordActivated.emit(record)

    def _show_context_menu(self, pos) -> None:
        index = self.table.indexAt(pos)
        if not index.isValid():
            return
        record = self.model.record_at(index.row())
        if not record:
            return
        menu = QMenu(self)
        detail_action = menu.addAction("جزئیات و یادداشت قرارداد")
        bookmark_action = menu.addAction(
            "حذف نشان" if self.notes.is_bookmarked(record.key) else "نشان‌گذاری قرارداد"
        )
        menu.addSeparator()
        pdf_action = menu.addAction("باز کردن فایل PDF")
        pdf_action.setEnabled(bool(record.pdf_file))
        folder_action = menu.addAction("نمایش در پوشهٔ PDF")
        folder_action.setEnabled(bool(record.pdf_file))
        menu.addSeparator()
        copy_number = menu.addAction("کپی شمارهٔ قرارداد")
        copy_subject = menu.addAction("کپی موضوع قرارداد")
        chosen = menu.exec_(self.table.viewport().mapToGlobal(pos))
        if chosen is None:
            return
        if chosen == detail_action:
            self.recordActivated.emit(record)
        elif chosen == bookmark_action:
            self.bookmarkToggled.emit(record.key)
        elif chosen in (pdf_action, folder_action):
            self.pdfRequested.emit(record.pdf_file)
        elif chosen == copy_number:
            self.copyRequested.emit(record.col3)
        elif chosen == copy_subject:
            self.copyRequested.emit(record.col4)

    # ---- وضعیت ----
    def current_filter(self) -> str:
        data = self.filter_combo.currentData()
        return data if data else "all"

    def current_search(self) -> str:
        return self.search_edit.text()

    def set_filter(self, key: str, label: Optional[str] = None) -> None:
        index = self.filter_combo.findData(key)
        if index < 0 and label:
            self.filter_combo.addItem(label, key)
            index = self.filter_combo.count() - 1
        elif index < 0 and key:
            self.filter_combo.addItem(key, key)
            index = self.filter_combo.count() - 1
        if index >= 0:
            self.filter_combo.setCurrentIndex(index)

    def clear_filters(self) -> None:
        self.blockSignals(True)
        self.search_edit.clear()
        while self.filter_combo.count() > len(FILTER_CHOICES):
            self.filter_combo.removeItem(self.filter_combo.count() - 1)
        self.filter_combo.setCurrentIndex(0)
        self.bookmark_check.setChecked(False)
        self.blockSignals(False)
        self.filterChanged.emit("all")

    def _set_column_visible(self, column: int, visible: bool) -> None:
        self.table.setColumnHidden(column, not visible)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self._colors = colors
        self.delegate.set_theme_colors(colors)
        self.model.set_theme_colors(colors)

    def _update_count(self) -> None:
        visible = self.model.visible_count()
        total = self.model.source_count()
        selected = len(self.table.selectionModel().selectedRows()) if self.table.selectionModel() else 0
        self.count_chip.setText(bidi(f"{fa(visible)} از {fa(total)}"))
        footer = f"نمایش {fa(visible)} قرارداد از مجموع {fa(total)}"
        if visible != total:
            footer += " (فیلتر فعال است)"
        if selected:
            footer += f"   •   انتخاب‌شده: {fa(selected)}"
        self.footer_label.setText(footer)

    def refresh_counts(self) -> None:
        self._update_count()

    def selected_records(self) -> List[ContractRecord]:
        rows = sorted({index.row() for index in self.table.selectionModel().selectedRows()})
        records = []
        for row in rows:
            record = self.model.record_at(row)
            if record:
                records.append(record)
        return records


# ---------------------------------------------------------------------------
# خروجی اکسل
# ---------------------------------------------------------------------------
class ExportDialog(QDialog):
    """انتخاب ستون‌ها و دامنهٔ خروجی اکسل."""

    def __init__(self, column_titles: Sequence[str], visible_count: int,
                 selected_count: int, total_count: int, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("خروجی اکسل")
        self.setMinimumWidth(520)
        self.setMinimumHeight(560)
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 16)
        layout.setSpacing(10)

        title = QLabel("خروجی سفارشی اکسل")
        title.setObjectName("SectionTitle")
        layout.addWidget(title)
        subtitle = QLabel("ستون‌ها و دامنهٔ داده‌ها را انتخاب کنید.")
        subtitle.setObjectName("Muted")
        layout.addWidget(subtitle)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setFormAlignment(Qt.AlignRight | Qt.AlignTop)
        self.scope_combo = Combo()
        self.scope_combo.addItem(f"ردیف‌های فیلترشده ({fa(visible_count)} مورد)", "visible")
        if selected_count:
            self.scope_combo.addItem(f"ردیف‌های انتخاب‌شده ({fa(selected_count)} مورد)", "selected")
        self.scope_combo.addItem(f"همهٔ قراردادها ({fa(total_count)} مورد)", "all")
        form.addRow("دامنه:", self.scope_combo)
        layout.addLayout(form)

        columns_label = QLabel("ستون‌ها:")
        columns_label.setObjectName("Muted")
        layout.addWidget(columns_label)

        columns_frame = QFrame()
        columns_frame.setObjectName("Card")
        grid = QGridLayout(columns_frame)
        grid.setContentsMargins(12, 10, 12, 10)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(6)
        self.checks: List[CheckBox] = []
        for index, name in enumerate(column_titles):
            check = CheckBox(name)
            check.setChecked(True)
            self.checks.append(check)
            check.set_theme_colors(THEMES["light"] if self.palette().color(QPalette.Window).lightness() > 128 else THEMES["dark"])
            grid.addWidget(check, index // 3, index % 3)
        for column in range(3):
            grid.setColumnStretch(column, 1)
        grid.setRowStretch(grid.rowCount(), 1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(columns_frame)
        scroll.setMinimumHeight(260)
        scroll.setMaximumHeight(320)
        layout.addWidget(scroll, 1)

        buttons_row = QHBoxLayout()
        select_all = QPushButton("انتخاب همه")
        select_all.clicked.connect(lambda: [c.setChecked(True) for c in self.checks])
        select_none = QPushButton("پاک کردن")
        select_none.clicked.connect(lambda: [c.setChecked(False) for c in self.checks])
        buttons_row.addWidget(select_all)
        buttons_row.addWidget(select_none)
        buttons_row.addStretch(1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("ذخیرهٔ فایل")
        buttons.button(QDialogButtonBox.Ok).setObjectName("Primary")
        buttons.button(QDialogButtonBox.Cancel).setText("انصراف")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        buttons_row.addWidget(buttons)
        layout.addLayout(buttons_row)

    def selected_columns(self) -> List[int]:
        return [index for index, check in enumerate(self.checks) if check.isChecked()]

    def scope(self) -> str:
        return str(self.scope_combo.currentData())

    def accept(self) -> None:
        if not self.selected_columns():
            QMessageBox.warning(self, "خروجی اکسل", "حداقل یک ستون انتخاب کنید.")
            return
        super().accept()


def record_export_row(record: ContractRecord) -> List[Any]:
    """مقادیر خام یک قرارداد برای خروجی اکسل."""
    return [
        record.col1,
        record.col2,
        record.col3,
        record.col4,
        record.f_value,
        "، ".join(record.open_flags),
        record.deductions,
        record.progress_table or record.progress_status,
        record.completion_table or record.completion_status,
        record.as_status,
        record.annex_status,
        record.progress_status,
        record.completion_status,
        record.pdf_file,
        record.excel_row,
    ]


EXPORT_COLUMN_TITLES = [
    "ردیف", "نام شرکت", "شماره قرارداد", "موضوع قرارداد", "نوع قرارداد",
    "موارد باز", "کسورات اسناد اختتام", "پیشرفت کارکرد", "پیشرفت اختتام",
    "وضعیت پکیج آزادسازی", "الحاقیه", "وضعیت پیشرفت کارکرد", "وضعیت اختتام",
    "فایل PDF", "شمارهٔ ردیف اکسل",
]


def write_contracts_workbook(path: str, records: Sequence[ContractRecord],
                             columns: Sequence[int], title: str = "قراردادها") -> int:
    """نوشتن فایل اکسل خروجی و برگرداندن تعداد ردیف‌ها."""
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست؛ خروجی اکسل نیاز به آن دارد.")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = title[:31]
    sheet.sheet_view.rightToLeft = True

    header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    header_font = XLFont(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="E4E9F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    headers = [EXPORT_COLUMN_TITLES[index] for index in columns]
    for column_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=column_index, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for row_index, record in enumerate(records, start=2):
        values = record_export_row(record)
        for column_index, source_index in enumerate(columns, start=1):
            cell = sheet.cell(row=row_index, column=column_index, value=values[source_index])
            cell.border = border
            cell.alignment = center if source_index in (0, 2, 4, 8, 14) else right

    widths = {
        "ردیف": 8, "نام شرکت": 26, "شماره قرارداد": 16, "موضوع قرارداد": 40,
        "نوع قرارداد": 18, "موارد باز": 22, "کسورات اسناد اختتام": 34,
        "پیشرفت کارکرد": 14, "پیشرفت اختتام": 14, "وضعیت پکیج آزادسازی": 16,
        "الحاقیه": 18, "وضعیت پیشرفت کارکرد": 16, "وضعیت اختتام": 16,
        "فایل PDF": 26, "شمارهٔ ردیف اکسل": 12,
    }
    for column_index, header in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(column_index)].width = widths.get(header, 16)
    sheet.freeze_panes = "A2"
    if records:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(records) + 1}"
    sheet.row_dimensions[1].height = 30
    workbook.save(path)
    return len(records)


def write_management_report(path: str, data: DashboardData) -> None:
    """گزارش مدیریتی سه‌بخشی، معادل /export/management_report نسخهٔ وب."""
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست؛ خروجی گزارش نیاز به آن دارد.")
    workbook = openpyxl.Workbook()
    header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    header_font = XLFont(color="FFFFFF", bold=True, size=11)
    title_font = XLFont(bold=True, size=14, color="0F766E")
    sub_font = XLFont(size=10, color="6B7280")
    section_font = XLFont(bold=True, size=12, color="0F766E")
    bold = XLFont(bold=True)
    red_font = XLFont(bold=True, color="DC2626")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="E4E9F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    sheet = workbook.active
    sheet.title = "خلاصهٔ مدیریتی"
    sheet.sheet_view.rightToLeft = True
    sheet.merge_cells("A1:B1")
    sheet["A1"] = "گزارش مدیریتی — پروژه مهر خلیج فارس (کندو)"
    sheet["A1"].font = title_font
    sheet.merge_cells("A2:B2")
    sheet["A2"] = f"تاریخ تهیهٔ گزارش: {today_jalali()}   •   منبع: {data.excel_path}"
    sheet["A2"].font = sub_font

    sheet["A4"] = "پیشرفت کل اختتام"
    sheet["A4"].font = bold
    sheet["B4"] = f"{data.progress_percent:.1f}%"
    sheet["B4"].font = XLFont(bold=True, size=13, color="0F766E")
    sheet["A4"].alignment = right
    sheet["B4"].alignment = center

    overview_rows = [
        ("تعداد کل قرارداد", data.total_contracts),
        ("قراردادهای دارای کسورات اسناد اختتام", data.contracts_with_deductions),
        ("نیازمند مفاصاحساب بیمه", data.flag_counts["mofasa"]),
        ("نیازمند تسویهٔ مالی نهایی", data.flag_counts["financial"]),
        ("نیازمند بالانس متریال", data.flag_counts["balance"]),
        ("نیازمند تحویل موقت", data.flag_counts["temporary"]),
        ("نیازمند تحویل قطعی", data.flag_counts["delivery"]),
        ("نیازمند صورت‌وضعیت ماقبل قطعی", data.flag_counts["prefinal"]),
        ("نیازمند صورت‌وضعیت قطعی", data.flag_counts["statement"]),
        ("نیازمند اقرارنامه", data.flag_counts["affidavit"]),
        ("پکیج‌های آزادسازی انجام‌شده", data.package_done),
        ("قراردادهای منقضی‌شده", data.expired_count),
    ]
    row = 6
    sheet.cell(row=row, column=1, value="شاخص").font = header_font
    sheet.cell(row=row, column=1).fill = header_fill
    sheet.cell(row=row, column=2, value="تعداد").font = header_font
    sheet.cell(row=row, column=2).fill = header_fill
    for column in (1, 2):
        sheet.cell(row=row, column=column).alignment = center
        sheet.cell(row=row, column=column).border = border
    for label, value in overview_rows:
        row += 1
        sheet.cell(row=row, column=1, value=label).alignment = right
        sheet.cell(row=row, column=1).border = border
        sheet.cell(row=row, column=2, value=value).alignment = center
        sheet.cell(row=row, column=2).border = border
    sheet.column_dimensions["A"].width = 42
    sheet.column_dimensions["B"].width = 14
    sheet.freeze_panes = "A6"

    follow = workbook.create_sheet("نیازمند پیگیری")
    follow.sheet_view.rightToLeft = True
    headers = ["نام شرکت", "شمارهٔ قرارداد", "موضوع", "الحاقیه منقضی", "تعداد موارد باز", "موارد باز"]
    for column_index, header in enumerate(headers, start=1):
        cell = follow.cell(row=1, column=column_index, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    follow_row = 1
    for record in data.records:
        flags = record.open_flags
        expired = record.is_expired
        if not flags and not expired:
            continue
        follow_row += 1
        values = [record.col2, record.col3, record.col4, "بله" if expired else "",
                  len(flags), "، ".join(flags)]
        for column_index, value in enumerate(values, start=1):
            cell = follow.cell(row=follow_row, column=column_index, value=value)
            cell.border = border
            cell.alignment = right if column_index in (1, 3, 6) else center
            if column_index == 4 and expired:
                cell.font = red_font
    for index, width in enumerate([26, 16, 34, 12, 12, 44], start=1):
        follow.column_dimensions[get_column_letter(index)].width = width
    if follow_row > 1:
        follow.auto_filter.ref = f"A1:F{follow_row}"
    follow.freeze_panes = "A2"

    dist_sheet = workbook.create_sheet("توزیع‌ها")
    dist_sheet.sheet_view.rightToLeft = True
    sections = [
        ("پکیج آزادسازی", data.dist('status')),
        ("وضعیت الحاقیه", data.dist("annex")),
        ("پیشرفت کارکرد", data.dist("progress")),
        ("درصد پیشرفت اختتام", data.dist("completion")),
        ("نوع قرارداد", data.dist("type")),
    ]
    dist_row = 1
    for title, distribution in sections:
        dist_sheet.cell(row=dist_row, column=1, value=title).font = section_font
        dist_row += 1
        for column_index, header in ((1, "برچسب"), (2, "تعداد")):
            cell = dist_sheet.cell(row=dist_row, column=column_index, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border
        dist_row += 1
        for label, value in distribution.as_list():
            dist_sheet.cell(row=dist_row, column=1, value=str(label)).alignment = right
            dist_sheet.cell(row=dist_row, column=1).border = border
            dist_sheet.cell(row=dist_row, column=2, value=int(value)).alignment = center
            dist_sheet.cell(row=dist_row, column=2).border = border
            dist_row += 1
        dist_row += 1
    dist_sheet.column_dimensions["A"].width = 34
    dist_sheet.column_dimensions["B"].width = 14

    workbook.save(path)


# ---------------------------------------------------------------------------
# ابزارهای رابط کاربری
# ---------------------------------------------------------------------------
def make_icon(kind: str, color: str, size: int = 18) -> QIcon:
    """ساخت QIcon از آیکون برداری برای دکمه‌ها."""
    scale = 2
    pixmap = QPixmap(size * scale, size * scale)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    draw_icon(painter, kind, QRectF(0, 0, size * scale, size * scale), QColor(color))
    painter.end()
    pixmap.setDevicePixelRatio(scale)
    return QIcon(pixmap)


def make_app_icon() -> QIcon:
    """آیکون برنامه (بدون فایل خارجی)."""
    size = 128
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    gradient = QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0, QColor("#12897d"))
    gradient.setColorAt(1, QColor("#0a504a"))
    painter.setPen(Qt.NoPen)
    painter.setBrush(QBrush(gradient))
    painter.drawRoundedRect(QRectF(4, 4, size - 8, size - 8), 28, 28)
    draw_icon(painter, "chart", QRectF(28, 30, 72, 64), QColor("#ffffff"))
    painter.end()
    return QIcon(pixmap)


def configure_rtl_browser(browser: QTextBrowser) -> None:
    """راست‌به‌چپ کردن متن HTML در QTextBrowser (Qt خودش dir را نمی‌شناسد)."""
    browser.setLayoutDirection(Qt.RightToLeft)
    document = browser.document()
    option = document.defaultTextOption()
    option.setTextDirection(Qt.RightToLeft)
    option.setAlignment(Qt.AlignRight)
    document.setDefaultTextOption(option)
    browser.setStyleSheet(browser.styleSheet())


def clear_layout(layout) -> None:
    """حذف همهٔ ویجت‌های یک چیدمان."""
    if layout is None:
        return
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        if widget is not None:
            widget.setParent(None)
            widget.deleteLater()
        elif item.layout() is not None:
            clear_layout(item.layout())


def label_value_pair(title: str, value: str, colors: Dict[str, str]) -> QWidget:
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(1)
    title_label = QLabel(title)
    title_label.setObjectName("Muted")
    font = title_label.font()
    font.setPointSizeF(max(7.5, font.pointSizeF() - 1))
    title_label.setFont(font)
    value_label = QLabel(value or "—")
    value_label.setWordWrap(True)
    value_label.setStyleSheet(f"color: {colors['text']}; font-weight: 600;")
    layout.addWidget(title_label)
    layout.addWidget(value_label)
    return widget


# ---------------------------------------------------------------------------
# پنل جزئیات و یادداشت قرارداد
# ---------------------------------------------------------------------------
class ContractDetailDock(QDockWidget):
    """پنل کنار پنجره برای جزئیات، یادداشت و پیگیری هر قرارداد."""

    bookmarkToggled = pyqtSignal(str)
    pdfRequested = pyqtSignal(str)
    revealRequested = pyqtSignal(str)
    notesChanged = pyqtSignal(str)

    def __init__(self, notes: NotesStore, colors: Dict[str, str], last_author: str = "",
                 parent=None) -> None:
        super().__init__("جزئیات قرارداد", parent)
        self.notes = notes
        self.colors = colors
        self.record: Optional[ContractRecord] = None
        self.setObjectName("DetailDock")
        self.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.setMinimumWidth(330)

        container = QWidget()
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setWidget(container)
        self.setWidget(self.scroll)

        self.body = QVBoxLayout(container)
        self.body.setContentsMargins(14, 12, 14, 14)
        self.body.setSpacing(10)

        self.title_label = QLabel("قراردادی انتخاب نشده")
        self.title_label.setObjectName("SectionTitle")
        self.title_label.setWordWrap(True)
        self.subtitle_label = QLabel("برای مشاهدهٔ جزئیات، روی یک ردیف جدول دو کلیک کنید.")
        self.subtitle_label.setObjectName("Muted")
        self.subtitle_label.setWordWrap(True)
        self.body.addWidget(self.title_label)
        self.body.addWidget(self.subtitle_label)

        actions = QHBoxLayout()
        self.bookmark_button = QPushButton("نشان‌گذاری")
        self.bookmark_button.setCheckable(True)
        self.bookmark_button.clicked.connect(self._on_bookmark_clicked)
        actions.addWidget(self.bookmark_button)
        self.pdf_button = QPushButton("باز کردن PDF")
        self.pdf_button.clicked.connect(self._on_pdf_clicked)
        actions.addWidget(self.pdf_button)
        self.folder_button = QPushButton("پوشه")
        self.folder_button.clicked.connect(self._on_folder_clicked)
        actions.addWidget(self.folder_button)
        self.body.addLayout(actions)

        self.flags_container = QWidget()
        self.flags_layout = FlowLayout(self.flags_container, 0, 6)
        self.body.addWidget(self.flags_container)

        self.fields_container = QWidget()
        self.fields_layout = QGridLayout(self.fields_container)
        self.fields_layout.setContentsMargins(0, 0, 0, 0)
        self.fields_layout.setHorizontalSpacing(12)
        self.fields_layout.setVerticalSpacing(10)
        self.body.addWidget(self.fields_container)

        self.deductions_title = QLabel("کسورات اسناد اختتام")
        self.deductions_title.setObjectName("SectionTitle")
        self.deductions_label = QLabel("—")
        self.deductions_label.setWordWrap(True)
        self.deductions_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.body.addWidget(self.deductions_title)
        self.body.addWidget(self.deductions_label)

        note_title_row = QHBoxLayout()
        note_title = QLabel("یادداشت و پیگیری")
        note_title.setObjectName("SectionTitle")
        note_title_row.addWidget(note_title)
        note_title_row.addStretch(1)
        self.note_count_label = QLabel("")
        self.note_count_label.setObjectName("Chip")
        note_title_row.addWidget(self.note_count_label)
        self.body.addLayout(note_title_row)

        self.author_edit = QLineEdit(last_author)
        self.author_edit.setPlaceholderText("نام شما (ثبت‌کنندهٔ یادداشت)")
        self.body.addWidget(self.author_edit)
        self.note_edit = QPlainTextEdit()
        self.note_edit.setPlaceholderText("یادداشت، نتیجهٔ تماس یا اقدام بعدی این قرارداد…")
        self.note_edit.setFixedHeight(90)
        self.body.addWidget(self.note_edit)
        follow_row = QHBoxLayout()
        self.follow_edit = QLineEdit()
        self.follow_edit.setPlaceholderText("تاریخ پیگیری بعدی (مثلاً ۱۴۰۳/۰۶/۱۵)")
        follow_row.addWidget(self.follow_edit, 1)
        today_btn = QPushButton("امروز")
        today_btn.setObjectName("ToolButton")
        today_btn.clicked.connect(lambda: self.follow_edit.setText(today_jalali(with_time=False)))
        follow_row.addWidget(today_btn)
        self.body.addLayout(follow_row)

        note_actions = QHBoxLayout()
        save_btn = QPushButton("ذخیرهٔ یادداشت")
        save_btn.setObjectName("Primary")
        save_btn.clicked.connect(self._save_note)
        note_actions.addWidget(save_btn)
        delete_btn = QPushButton("پاک کردن یادداشت‌ها")
        delete_btn.setObjectName("Danger")
        delete_btn.clicked.connect(self._delete_notes)
        note_actions.addWidget(delete_btn)
        self.body.addLayout(note_actions)

        self.notes_list = QListWidget()
        self.notes_list.setMinimumHeight(110)
        self.body.addWidget(self.notes_list)
        self.body.addStretch(1)
        self.setEnabled(False)

    # ---- رخدادها ----
    def _on_bookmark_clicked(self) -> None:
        if self.record:
            self.bookmarkToggled.emit(self.record.key)

    def _on_pdf_clicked(self) -> None:
        if self.record and self.record.pdf_file:
            self.pdfRequested.emit(self.record.pdf_file)

    def _on_folder_clicked(self) -> None:
        if self.record and self.record.pdf_file:
            self.revealRequested.emit(self.record.pdf_file)

    def _save_note(self) -> None:
        if not self.record:
            return
        text = self.note_edit.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "یادداشت", "متن یادداشت خالی است.")
            return
        author = self.author_edit.text().strip() or "کاربر"
        self.notes.add_note(self.record.key, author, text, self.follow_edit.text())
        self.note_edit.clear()
        self.follow_edit.clear()
        self.refresh_notes()
        self.notesChanged.emit(self.record.key)

    def _delete_notes(self) -> None:
        if not self.record:
            return
        if not self.notes.notes_for(self.record.key):
            return
        answer = QMessageBox.question(
            self, "حذف یادداشت‌ها",
            "همهٔ یادداشت‌های این قرارداد حذف شود؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if answer == QMessageBox.Yes:
            self.notes.delete_notes(self.record.key)
            self.refresh_notes()
            self.notesChanged.emit(self.record.key)

    # ---- نمایش ----
    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        if self.record:
            self.set_record(self.record)

    def refresh_notes(self) -> None:
        self.notes_list.clear()
        if not self.record:
            self.note_count_label.setText("")
            return
        items = self.notes.notes_for(self.record.key)
        self.note_count_label.setText(bidi(f"{fa(len(items))} یادداشت") if items else "بدون یادداشت")
        for note in reversed(items):
            created = note.get("created", "")
            try:
                stamp = jalali_datetime(datetime.fromisoformat(created))
            except Exception:
                stamp = created
            followup = note.get("followup", "")
            text = f"{stamp} — {note.get('author') or 'کاربر'}"
            if followup:
                text += f"   •   پیگیری: {followup}"
            item = QListWidgetItem(f"{text}\n{note.get('text', '')}")
            item.setToolTip(note.get("text", ""))
            self.notes_list.addItem(item)
        if not items:
            self.notes_list.addItem(QListWidgetItem("هنوز یادداشتی ثبت نشده است."))

    def set_record(self, record: Optional[ContractRecord], pdf_exists: Optional[bool] = None) -> None:
        self.record = record
        if record is None:
            self.setEnabled(False)
            self.title_label.setText("قراردادی انتخاب نشده")
            self.subtitle_label.setText("برای مشاهدهٔ جزئیات، روی یک ردیف جدول دو کلیک کنید.")
            clear_layout(self.fields_layout)
            clear_layout(self.flags_layout)
            self.deductions_label.setText("—")
            self.notes_list.clear()
            self.note_count_label.setText("")
            return

        self.setEnabled(True)
        self.title_label.setText(record.col2 or "بدون نام شرکت")
        self.subtitle_label.setText(
            f"قرارداد شمارهٔ {record.col3 or '—'}   •   موضوع: {record.col4 or '—'}"
        )
        bookmarked = self.notes.is_bookmarked(record.key)
        self.bookmark_button.setChecked(bookmarked)
        self.bookmark_button.setText("حذف نشان" if bookmarked else "نشان‌گذاری")
        self.pdf_button.setEnabled(bool(record.pdf_file))
        self.folder_button.setEnabled(bool(record.pdf_file))
        if record.pdf_file:
            self.pdf_button.setText("باز کردن PDF")
        else:
            self.pdf_button.setText("PDF نیست")

        clear_layout(self.flags_layout)
        chip_items = [
            (f"{FLAG_LABELS[key]}", True) for key in FLAG_LABELS if record.flags.get(key)
        ]
        if not chip_items:
            chip_items = [("همهٔ موارد بسته است", False)]
        for label, active in chip_items:
            chip = QLabel(label)
            if active:
                chip.setStyleSheet(
                    f"background: {self.colors['warning_soft']}; color: {self.colors['warning']};"
                    f"border-radius: 9px; padding: 3px 9px; font-size: 8.5pt; font-weight: 700;"
                )
            else:
                chip.setStyleSheet(
                    f"background: {self.colors['success_soft']}; color: {self.colors['success']};"
                    f"border-radius: 9px; padding: 3px 9px; font-size: 8.5pt; font-weight: 700;"
                )
            self.flags_layout.addWidget(chip)

        clear_layout(self.fields_layout)
        fields = [
            ("وضعیت پکیج آزادسازی", record.as_status or "—"),
            ("پیشرفت کارکرد (ستون ۵۳)", record.progress_status or "—"),
            ("درصد پیشرفت اختتام (ستون ۵۴)", record.completion_status or "—"),
            ("الحاقیه / زمان باقی‌مانده", record.annex_status or "—"),
            ("نوع قرارداد", record.f_value or "—"),
            ("مقدار پیشرفت جدول (ستون ۱۰)", record.progress_table or "—"),
            ("درصد اختتام جدول (ستون ۳۶)", record.completion_table or "—"),
            ("شمارهٔ ردیف در اکسل", fa(record.excel_row)),
        ]
        if record.aw or record.ax or record.ay:
            fields.append(("ستون‌های AW / AX / AY", " / ".join(x for x in (record.aw, record.ax, record.ay) if x)))
        for index, (title, value) in enumerate(fields):
            widget = label_value_pair(title, value, self.colors)
            self.fields_layout.addWidget(widget, index // 2, index % 2)

        self.deductions_label.setText(record.deductions or "کسوراتی برای این قرارداد ثبت نشده است.")
        self.refresh_notes()


# ---------------------------------------------------------------------------
# صفحهٔ گزارش مدیریتی
# ---------------------------------------------------------------------------
class ReportsPage(QWidget):
    """پیش‌نمایش گزارش مدیریتی + خروجی اکسل و چاپ."""

    exportExcelRequested = pyqtSignal()
    printRequested = pyqtSignal()
    refreshRequested = pyqtSignal()

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 18)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(0)
        title = QLabel("گزارش مدیریتی")
        title.setObjectName("PageTitle")
        subtitle = QLabel("خلاصهٔ شاخص‌ها، قراردادهای نیازمند پیگیری و توزیع وضعیت‌ها")
        subtitle.setObjectName("Muted")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        title_row.addLayout(title_box)
        title_row.addStretch(1)

        refresh_btn = QPushButton("بازخوانی")
        refresh_btn.clicked.connect(self.refreshRequested.emit)
        title_row.addWidget(refresh_btn, 0, Qt.AlignTop)
        print_btn = QPushButton("چاپ / ذخیرهٔ PDF")
        print_btn.clicked.connect(self.printRequested.emit)
        title_row.addWidget(print_btn, 0, Qt.AlignTop)
        export_btn = QPushButton("خروجی اکسل گزارش")
        export_btn.setObjectName("Primary")
        export_btn.clicked.connect(self.exportExcelRequested.emit)
        title_row.addWidget(export_btn, 0, Qt.AlignTop)
        layout.addLayout(title_row)

        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(False)
        self.browser.setFrameShape(QFrame.NoFrame)
        configure_rtl_browser(self.browser)
        self.browser.setStyleSheet(
            f"background: {colors['surface']}; border: 1px solid {colors['border']};"
            f"border-radius: 12px; padding: 18px;"
        )
        layout.addWidget(self.browser, 1)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.browser.setStyleSheet(
            f"background: {colors['surface']}; border: 1px solid {colors['border']};"
            f"border-radius: 12px; padding: 18px;"
        )

    @staticmethod
    def _table(headers: Sequence[str], rows: Sequence[Sequence[Any]],
               colors: Dict[str, str], widths: Optional[Sequence[str]] = None) -> str:
        head_cells = "".join(
            f'<th bgcolor="{colors["primary_soft"]}" style="padding:7px; color:{colors["primary_dark"]};">'
            f'{html_module.escape(str(h))}</th>'
            for h in headers
        )
        body = ""
        for row in rows:
            cells = "".join(
                f'<td style="padding:6px 8px; border-bottom:1px solid {colors["border_soft"]};">'
                f'{html_module.escape(str(value))}</td>'
                for value in row
            )
            body += f"<tr>{cells}</tr>"
        colgroup = ""
        if widths:
            colgroup = "<colgroup>" + "".join(f'<col width="{w}">' for w in widths) + "</colgroup>"
        return (
            f'<table width="100%" cellspacing="0" cellpadding="0" border="0" '
            f'style="border-collapse:collapse; margin-bottom:14px;">{colgroup}'
            f"<tr>{head_cells}</tr>{body}</table>"
        )

    def update_data(self, data: DashboardData) -> None:
        colors = self.colors
        counts = data.flag_counts
        overview = [
            ("تعداد کل قرارداد", fa_num(data.total_contracts)),
            ("قراردادهای دارای کسورات اسناد اختتام", fa_num(data.contracts_with_deductions)),
            ("نیازمند مفاصاحساب بیمه", fa_num(counts["mofasa"])),
            ("نیازمند تسویهٔ مالی نهایی", fa_num(counts["financial"])),
            ("نیازمند بالانس متریال", fa_num(counts["balance"])),
            ("نیازمند تحویل موقت", fa_num(counts["temporary"])),
            ("نیازمند تحویل قطعی", fa_num(counts["delivery"])),
            ("نیازمند صورت‌وضعیت ماقبل قطعی", fa_num(counts["prefinal"])),
            ("نیازمند صورت‌وضعیت قطعی", fa_num(counts["statement"])),
            ("نیازمند اقرارنامه", fa_num(counts["affidavit"])),
            ("پکیج‌های آزادسازی انجام‌شده", fa_num(data.package_done)),
            ("قراردادهای منقضی‌شده", fa_num(data.expired_count)),
        ]
        follow_rows = []
        for record in data.records:
            flags = record.open_flags
            if not flags and not record.is_expired:
                continue
            follow_rows.append([
                record.col2 or "—",
                record.col3 or "—",
                _truncate(record.col4 or "—", 60),
                "بله" if record.is_expired else "—",
                fa(len(flags)),
                "، ".join(flags) or "—",
            ])
        shown = follow_rows[:60]
        dist_rows = []
        for name in ("status", "annex", "progress", "completion", "type"):
            distribution = data.dist('status') if name == "status" else data.dist(name)
            labels = "، ".join(
                f"{label} ({fa(value)})" for label, value in distribution.as_list()
            ) or "—"
            dist_rows.append([CHART_TITLES.get(name, name), labels])

        html_parts = [
            "<div dir='rtl' style='font-family: Tahoma; color: %s;'>" % colors["text"],
            f"<h2 style='color:{colors['primary']}; margin:0 0 4px 0;'>"
            f"گزارش مدیریتی — پروژهٔ مهر خلیج فارس (کندو)</h2>",
            f"<div style='color:{colors['text_muted']}; margin-bottom:14px;'>"
            f"تاریخ تهیه: {today_jalali()} &nbsp;•&nbsp; منبع داده: "
            f"{html_module.escape(os.path.basename(data.excel_path) or '—')} &nbsp;•&nbsp; "
            f"آخرین تغییر فایل: {data.source_mtime_text}</div>",
            f"<div style='font-size:22px; font-weight:bold; color:{colors['primary']}; margin:10px 0;'>"
            f"پیشرفت کل اختتام: {data.progress_percent_text}٪</div>",
            "<h3 style='color:%s;'>شاخص‌های کلی</h3>" % colors["primary"],
            self._table(["شاخص", "مقدار"], overview, colors, ["70%", "30%"]),
            "<h3 style='color:%s;'>قراردادهای نیازمند پیگیری</h3>" % colors["primary"],
            self._table(
                ["نام شرکت", "شمارهٔ قرارداد", "موضوع", "منقضی", "موارد باز", "شرح موارد"],
                shown, colors, ["20%", "13%", "27%", "8%", "8%", "24%"],
            ),
            (f"<div style='color:{colors['text_muted']};'>مجموع قراردادهای نیازمند پیگیری: "
             f"{fa_num(len(follow_rows))}"
             + (f" — فقط {fa(len(shown))} مورد اول نمایش داده شده است." if len(follow_rows) > len(shown) else "")
             + "</div><br>"),
            "<h3 style='color:%s;'>توزیع‌ها</h3>" % colors["primary"],
            self._table(["نمودار", "توزیع (تعداد)"], dist_rows, colors, ["25%", "75%"]),
            "</div>",
        ]
        self.browser.setHtml("".join(html_parts))
        configure_rtl_browser(self.browser)

    def html(self) -> str:
        return self.browser.toHtml()


# ---------------------------------------------------------------------------
# صفحهٔ کتابخانهٔ PDF
# ---------------------------------------------------------------------------
class PdfLibraryPage(QWidget):
    """فهرست فایل‌های PDF پوشهٔ قراردادها با جستجو و باز کردن."""

    openRequested = pyqtSignal(str)
    folderRequested = pyqtSignal()

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.folder = ""
        self._files: List[str] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 18)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(0)
        title = QLabel("کتابخانهٔ PDF قراردادها")
        title.setObjectName("PageTitle")
        self.subtitle = QLabel("پوشهٔ PDF تنظیم نشده است.")
        self.subtitle.setObjectName("Muted")
        self.subtitle.setWordWrap(True)
        title_box.addWidget(title)
        title_box.addWidget(self.subtitle)
        title_row.addLayout(title_box)
        title_row.addStretch(1)
        self.count_chip = QLabel("—")
        self.count_chip.setObjectName("Chip")
        title_row.addWidget(self.count_chip, 0, Qt.AlignTop)
        layout.addLayout(title_row)

        toolbar = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setObjectName("Search")
        self.search_edit.setPlaceholderText("جستجوی نام فایل…")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._apply_filter)
        toolbar.addWidget(self.search_edit, 1)
        refresh_btn = QPushButton("بازخوانی فهرست")
        refresh_btn.clicked.connect(self.refresh_list)
        toolbar.addWidget(refresh_btn)
        folder_btn = QPushButton("انتخاب پوشه")
        folder_btn.clicked.connect(self.folderRequested.emit)
        toolbar.addWidget(folder_btn)
        layout.addLayout(toolbar)

        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self._open_item)
        layout.addWidget(self.list_widget, 1)

        footer = QHBoxLayout()
        self.footer_label = QLabel("—")
        self.footer_label.setObjectName("Muted")
        footer.addWidget(self.footer_label, 1)
        open_btn = QPushButton("باز کردن فایل انتخاب‌شده")
        open_btn.setObjectName("Primary")
        open_btn.clicked.connect(self._open_selected)
        footer.addWidget(open_btn)
        layout.addLayout(footer)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors

    def set_folder(self, folder: str) -> None:
        self.folder = folder or ""
        self.subtitle.setText(self.folder or "پوشهٔ PDF تنظیم نشده است.")
        self.refresh_list()

    def refresh_list(self) -> None:
        self._files = []
        if self.folder and os.path.isdir(self.folder):
            try:
                with os.scandir(self.folder) as entries:
                    for entry in entries:
                        try:
                            if entry.is_file() and entry.name.lower().endswith(".pdf"):
                                self._files.append(entry.name)
                        except OSError:
                            continue
            except Exception as exc:
                self.footer_label.setText(f"خطا در خواندن پوشه: {exc}")
        self._files.sort(key=normalize_per)
        self._apply_filter()
        self.count_chip.setText(bidi(f"{fa(len(self._files))} فایل"))

    def _apply_filter(self) -> None:
        query = normalize_per(self.search_edit.text())
        self.list_widget.clear()
        shown = 0
        for name in self._files:
            if query and query not in normalize_per(name):
                continue
            shown += 1
            if shown > 5000:
                break
            item = QListWidgetItem(name)
            item.setData(Qt.UserRole, name)
            item.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
            self.list_widget.addItem(item)
        if not shown:
            message = ("پوشه‌ای انتخاب نشده است." if not self.folder
                       else "فایل PDF‌ای در این پوشه پیدا نشد.")
            self.list_widget.addItem(QListWidgetItem(message))
        self.footer_label.setText(
            f"{fa(shown)} فایل نمایش داده می‌شود" + (f" از {fa(len(self._files))}" if query else "")
        )

    def _open_selected(self) -> None:
        item = self.list_widget.currentItem()
        if item:
            self._open_item(item)

    def _open_item(self, item: QListWidgetItem) -> None:
        name = item.data(Qt.UserRole)
        if name:
            self.openRequested.emit(str(name))


# ---------------------------------------------------------------------------
# صفحهٔ تنظیمات
# ---------------------------------------------------------------------------
class SettingsPage(QWidget):
    """تنظیم مسیر اکسل/PDF، پوسته و بازخوانی خودکار."""

    saveRequested = pyqtSignal(dict)
    reloadRequested = pyqtSignal()
    openDataFolderRequested = pyqtSignal()
    browseDataFolderRequested = pyqtSignal()

    def __init__(self, config: AppConfig, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.config = config
        self.colors = colors
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 22)
        layout.setSpacing(14)

        title = QLabel("تنظیمات برنامه")
        title.setObjectName("PageTitle")
        subtitle = QLabel("مسیر منبع داده و رفتار برنامه؛ این تنظیمات در فایل JSON کنار برنامه ذخیره می‌شود.")
        subtitle.setObjectName("Muted")
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content = QWidget()
        scroll.setWidget(content)
        form_holder = QVBoxLayout(content)
        form_holder.setContentsMargins(0, 0, 0, 0)
        form_holder.setSpacing(14)

        source_group = QGroupBox("منبع داده")
        source_layout = QGridLayout(source_group)
        source_layout.setHorizontalSpacing(8)
        source_layout.setVerticalSpacing(10)

        source_layout.addWidget(QLabel("فایل اکسل:"), 0, 0)
        self.excel_edit = QLineEdit(str(config.get("excel_path") or ""))
        self.excel_edit.setPlaceholderText("مسیر فایل F.O.C(Mehr).xlsx")
        source_layout.addWidget(self.excel_edit, 0, 1)
        browse_excel = QPushButton("انتخاب…")
        browse_excel.clicked.connect(self._browse_excel)
        source_layout.addWidget(browse_excel, 0, 2)
        network_btn = QPushButton("مسیر شبکهٔ قبلی")
        network_btn.setToolTip(ORIGINAL_EXCEL_PATH)
        network_btn.clicked.connect(lambda: self.excel_edit.setText(ORIGINAL_EXCEL_PATH))
        source_layout.addWidget(network_btn, 0, 3)

        source_layout.addWidget(QLabel("نام شیت:"), 1, 0)
        self.sheet_edit = QLineEdit(str(config.get("sheet_name") or SHEET_NAME_DEFAULT))
        source_layout.addWidget(self.sheet_edit, 1, 1)

        source_layout.addWidget(QLabel("پوشهٔ PDF:"), 2, 0)
        self.pdf_edit = QLineEdit(str(config.get("pdf_folder") or ""))
        self.pdf_edit.setPlaceholderText("پوشه‌ای که فایل‌های PDF قراردادها در آن است")
        source_layout.addWidget(self.pdf_edit, 2, 1)
        browse_pdf = QPushButton("انتخاب…")
        browse_pdf.clicked.connect(self._browse_pdf)
        source_layout.addWidget(browse_pdf, 2, 2)
        pdf_network = QPushButton("مسیر شبکهٔ قبلی")
        pdf_network.setToolTip(ORIGINAL_PDF_FOLDER)
        pdf_network.clicked.connect(lambda: self.pdf_edit.setText(ORIGINAL_PDF_FOLDER))
        source_layout.addWidget(pdf_network, 2, 3)
        source_layout.setColumnStretch(1, 1)
        form_holder.addWidget(source_group)

        display_group = QGroupBox("نمایش")
        display_layout = QFormLayout(display_group)
        self.theme_combo = Combo()
        self.theme_combo.addItem("روشن", "light")
        self.theme_combo.addItem("تیره", "dark")
        current_theme = str(config.get("theme") or "light")
        index = self.theme_combo.findData(current_theme)
        self.theme_combo.setCurrentIndex(max(0, index))
        display_layout.addRow("پوسته:", self.theme_combo)

        self.font_spin = SpinBox()
        self.font_spin.setRange(8, 16)
        self.font_spin.setSuffix(" pt")
        self.font_spin.setValue(int(config.get("font_size") or 10))
        display_layout.addRow("اندازهٔ قلم:", self.font_spin)

        self.auto_spin = SpinBox()
        self.auto_spin.setRange(0, 120)
        self.auto_spin.setSuffix(" دقیقه")
        self.auto_spin.setSpecialValueText("غیرفعال")
        self.auto_spin.setValue(int(config.get("auto_refresh_minutes") or 5))
        display_layout.addRow("بازخوانی خودکار:", self.auto_spin)
        form_holder.addWidget(display_group)

        storage_group = QGroupBox("فایل‌های برنامه")
        storage_layout = QVBoxLayout(storage_group)
        self.config_path_label = QLabel(f"تنظیمات: {config.path}")
        self.config_path_label.setObjectName("Muted")
        self.config_path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.notes_path_label = QLabel(f"یادداشت‌ها: {data_dir() / NotesStore.FILENAME}")
        self.notes_path_label.setObjectName("Muted")
        self.notes_path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        storage_layout.addWidget(self.config_path_label)
        storage_layout.addWidget(self.notes_path_label)
        open_folder = QPushButton("باز کردن پوشهٔ برنامه")
        open_folder.clicked.connect(self.openDataFolderRequested.emit)
        storage_layout.addWidget(open_folder, 0, Qt.AlignRight)
        form_holder.addWidget(storage_group)

        buttons = QHBoxLayout()
        save_btn = QPushButton("ذخیره و اعمال")
        save_btn.setObjectName("Primary")
        save_btn.clicked.connect(self._emit_save)
        buttons.addWidget(save_btn)
        reload_btn = QPushButton("بارگذاری مجدد داده‌ها")
        reload_btn.clicked.connect(self.reloadRequested.emit)
        buttons.addWidget(reload_btn)
        buttons.addStretch(1)
        form_holder.addLayout(buttons)
        form_holder.addStretch(1)
        layout.addWidget(scroll, 1)

    def _browse_excel(self) -> None:
        start = self.excel_edit.text() or str(app_dir())
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل", start, EXCEL_FILTER)
        if path:
            self.excel_edit.setText(path)
            if not self.pdf_edit.text().strip():
                guess = Path(path).parent / "PDF-MEHR"
                if guess.is_dir():
                    self.pdf_edit.setText(str(guess))

    def _browse_pdf(self) -> None:
        start = self.pdf_edit.text() or str(app_dir())
        folder = QFileDialog.getExistingDirectory(self, "انتخاب پوشهٔ PDF", start)
        if folder:
            self.pdf_edit.setText(folder)

    def _emit_save(self) -> None:
        self.saveRequested.emit({
            "excel_path": self.excel_edit.text().strip(),
            "sheet_name": self.sheet_edit.text().strip() or SHEET_NAME_DEFAULT,
            "pdf_folder": self.pdf_edit.text().strip(),
            "theme": self.theme_combo.currentData(),
            "font_size": self.font_spin.value(),
            "auto_refresh_minutes": self.auto_spin.value(),
        })

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors

    def load_values(self, config: AppConfig) -> None:
        self.excel_edit.setText(str(config.get("excel_path") or ""))
        self.sheet_edit.setText(str(config.get("sheet_name") or SHEET_NAME_DEFAULT))
        self.pdf_edit.setText(str(config.get("pdf_folder") or ""))
        theme_index = self.theme_combo.findData(str(config.get("theme") or "light"))
        self.theme_combo.setCurrentIndex(max(0, theme_index))
        self.font_spin.setValue(int(config.get("font_size") or 10))
        self.auto_spin.setValue(int(config.get("auto_refresh_minutes") or 0))


# ---------------------------------------------------------------------------
# صفحهٔ راهنما
# ---------------------------------------------------------------------------
class HelpPage(QWidget):
    """راهنمای برنامه و نگاشت امکانات نسخهٔ وب به نسخهٔ دسکتاپ."""

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 18)
        title = QLabel("راهنما و توضیحات")
        title.setObjectName("PageTitle")
        layout.addWidget(title)
        self.browser = QTextBrowser()
        self.browser.setFrameShape(QFrame.NoFrame)
        configure_rtl_browser(self.browser)
        layout.addWidget(self.browser, 1)
        self.browser.setHtml(self._html())

    def _html(self) -> str:
        c = self.colors
        return f"""
        <div align="right" dir="rtl" style="font-family: Tahoma; color: {c['text']}; line-height: 170%;">
        <h2 style="color:{c['primary']};">داشبورد قراردادهای مهر خلیج فارس — نسخهٔ دسکتاپ</h2>
        <p>این برنامه نسخهٔ مستقل و تک‌فایلی داشبورد Flask است. دیگر نه سرور Flask لازم است،
        نه مرورگر و نه فایل <code>dashboard.html</code>. داده‌ها مستقیماً از فایل اکسل
        (شیت <code>colorcode</code>) خوانده می‌شوند و همان شماره‌ستون‌های نسخهٔ وب استفاده می‌شوند؛
        بنابراین محاسبات دقیقاً مانند قبل است.</p>

        <h3 style="color:{c['primary']};">شروع سریع</h3>
        <ol>
          <li>از صفحهٔ «تنظیمات» فایل اکسل را انتخاب کنید (یا فایل را کنار همین اسکریپت بگذارید).</li>
          <li>پوشهٔ PDF را انتخاب کنید تا ستون «فایل» در جدول فعال شود.</li>
          <li>دکمهٔ «به‌روزرسانی داده‌ها» یا <b>Ctrl+R</b> را بزنید.</li>
        </ol>

        <h3 style="color:{c['primary']};">نقشهٔ امکانات نسخهٔ وب ← دسکتاپ</h3>
        <table width="100%" cellspacing="0" cellpadding="6" border="1" style="border-color:{c['border']};">
        <tr bgcolor="{c['primary_soft']}"><th>نسخهٔ وب (Flask)</th><th>در این برنامه</th></tr>
        <tr><td>صفحهٔ اصلی با کارت‌های آماری</td><td>صفحهٔ «داشبورد» با کارت‌های قابل کلیک</td></tr>
        <tr><td>فیلتر با کلیک روی کارت‌ها</td><td>کلیک روی کارت یا «موارد باز» → جدول فیلتر می‌شود</td></tr>
        <tr><td>نمودارهای Chart.js</td><td>نمودارهای بومی Qt با کلیک و راهنمای تعاملی</td></tr>
        <tr><td>جدول قراردادها</td><td>صفحهٔ «جدول قراردادها» با جستجو، مرتب‌سازی و انتخاب چندگانه</td></tr>
        <tr><td>پنجرهٔ جزئیات/یادداشت سمت راست</td><td>پنل «جزئیات قرارداد» (راست پنجره)</td></tr>
        <tr><td>خروجی اکسل مرورگر (XLSX)</td><td>خروجی اکسل با انتخاب ستون‌ها و دامنهٔ داده</td></tr>
        <tr><td>/export/management_report</td><td>صفحهٔ «گزارش مدیریتی» → خروجی اکسل سه‌شیتی</td></tr>
        <tr><td>لینک فایل PDF + رمز</td><td>ستون «فایل» و پنل جزئیات؛ باز شدن با برنامهٔ پیش‌فرض</td></tr>
        <tr><td>/api/ollama-chat</td><td>حذف شده (نیازمند سرور محلی Ollama)؛ بقیهٔ امکانات مستقل کار می‌کنند</td></tr>
        <tr><td>ورود با رمز ۱۵۲۴ و کاربران شبکه</td><td>حذف شده؛ برنامه تک‌کاربرهٔ دسکتاپ است</td></tr>
        </table>

        <h3 style="color:{c['primary']};">میان‌برها</h3>
        <table cellspacing="0" cellpadding="5">
        <tr><td><kbd>Ctrl+R</kbd></td><td>بازخوانی اکسل</td></tr>
        <tr><td><kbd>Ctrl+F</kbd></td><td>جستجو</td></tr>
        <tr><td><kbd>Ctrl+E</kbd></td><td>خروجی اکسل جدول</td></tr>
        <tr><td><kbd>Ctrl+D</kbd></td><td>تغییر پوستهٔ روشن/تیره</td></tr>
        <tr><td><kbd>F1</kbd></td><td>راهنما</td></tr>
        <tr><td><kbd>Esc</kbd></td><td>بستن پنل جزئیات</td></tr>
        </table>

        <h3 style="color:{c['primary']};">فایل‌های ذخیره‌شده</h3>
        <p>تنظیمات در <code>mehr_qt_config.json</code> و یادداشت‌ها/نشان‌ها در
        <code>mehr_notes.json</code> در همان پوشهٔ برنامه (یا در صورت نداشتن دسترسی، پوشهٔ کاربر) ذخیره می‌شوند.</p>

        <h3 style="color:{c['primary']};">نصب پیش‌نیازها</h3>
        <pre style="background:{c['surface_soft']}; padding:10px; border-radius:8px;">pip install PyQt5 openpyxl</pre>
        <p style="color:{c['text_muted']};">نسخهٔ {APP_VERSION} — ساخته‌شده بر پایهٔ منطق mehr_optimized.py و dashboard.html</p>
        </div>
        """

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.browser.setHtml(self._html())
        configure_rtl_browser(self.browser)


# ---------------------------------------------------------------------------
# پنجرهٔ اصلی
# ---------------------------------------------------------------------------
NAV_ITEMS = [
    ("نمایش", [
        ("dashboard", "داشبورد", "chart"),
        ("contracts", "جدول قراردادها", "table"),
    ]),
    ("گزارش‌ها", [
        ("reports", "گزارش مدیریتی", "document"),
        ("pdfs", "کتابخانهٔ PDF", "pdf"),
    ]),
    ("سیستم", [
        ("settings", "تنظیمات", "gear"),
        ("help", "راهنما", "help"),
    ]),
]

PAGE_TITLES = {
    "dashboard": "داشبورد",
    "contracts": "جدول قراردادها",
    "reports": "گزارش مدیریتی",
    "pdfs": "کتابخانهٔ PDF",
    "settings": "تنظیمات",
    "help": "راهنما",
}

FIELD_LABELS = {
    "type": "نوع قرارداد",
    "status": "وضعیت پکیج",
    "progress": "پیشرفت کارکرد",
    "completion": "پیشرفت اختتام",
    "annex": "الحاقیه",
}


class MainWindow(QMainWindow):
    def __init__(self, config: AppConfig, notes: NotesStore) -> None:
        super().__init__()
        self.config = config
        self.notes = notes
        self.colors = THEMES.get(str(config.get("theme") or "light"), THEMES["light"])
        self.data: DashboardData = empty_dashboard()
        self.loader: Optional[ExcelLoader] = None
        self._last_error = ""
        self._current_page = "dashboard"
        self.nav_buttons: Dict[str, QPushButton] = {}

        self.setWindowTitle(f"{APP_NAME} — {APP_SHORT}")
        self.setWindowIcon(make_app_icon())
        self.setMinimumSize(1100, 700)
        self.resize(1440, 900)
        self.setAcceptDrops(True)

        self._build_ui()
        self._create_actions()
        self._restore_geometry()
        self.apply_theme(str(config.get("theme") or "light"), int(config.get("font_size") or 10),
                         persist=False)
        self._setup_auto_timer()
        QTimer.singleShot(250, self.initial_load)

    # ------------------------------------------------------------------
    # ساخت رابط
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- نوار بالایی ----
        header = QFrame()
        header.setObjectName("HeaderBar")
        header.setFixedHeight(66)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(18, 8, 18, 8)
        header_layout.setSpacing(10)

        self.page_title_label = QLabel(PAGE_TITLES["dashboard"])
        self.page_title_label.setStyleSheet("font-size: 12pt; font-weight: 800;")
        header_layout.addWidget(self.page_title_label)

        self.search_edit = QLineEdit()
        self.search_edit.setObjectName("Search")
        self.search_edit.setPlaceholderText("جستجو در قراردادها: شماره، نام شرکت، موضوع…")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMaximumWidth(430)
        self.search_edit.setMinimumWidth(240)
        self.search_action = self.search_edit.addAction(
            make_icon("search", self.colors["text_muted"], 16), QLineEdit.LeadingPosition
        )
        self.search_edit.textChanged.connect(self._on_search_text_changed)
        header_layout.addStretch(1)
        header_layout.addWidget(self.search_edit, 2)
        header_layout.addStretch(1)

        self.refresh_button = QPushButton("به‌روزرسانی")
        self.refresh_button.setObjectName("Primary")
        self.refresh_button.setToolTip("بازخوانی فایل اکسل (Ctrl+R)")
        self.refresh_button.clicked.connect(lambda: self.load_data())
        header_layout.addWidget(self.refresh_button)

        self.export_button = QPushButton("خروجی اکسل")
        self.export_button.setToolTip("خروجی اکسل از جدول (Ctrl+E)")
        self.export_button.clicked.connect(self.export_table)
        header_layout.addWidget(self.export_button)

        self.theme_button = QPushButton()
        self.theme_button.setObjectName("ToolButton")
        self.theme_button.setToolTip("تغییر پوستهٔ روشن/تیره (Ctrl+D)")
        self.theme_button.clicked.connect(self.toggle_theme)
        header_layout.addWidget(self.theme_button)
        root.addWidget(header)

        # ---- بدنه: سایدبار + صفحه‌ها ----
        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(236)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(10, 14, 10, 12)
        sidebar_layout.setSpacing(2)

        brand_row = QHBoxLayout()
        brand_row.setContentsMargins(6, 0, 6, 6)
        brand_row.setSpacing(9)
        logo = QLabel()
        logo.setPixmap(make_app_icon().pixmap(40, 40))
        logo.setFixedSize(42, 42)
        brand_row.addWidget(logo)
        brand_text = QVBoxLayout()
        brand_text.setSpacing(0)
        brand_title = QLabel(APP_SHORT)
        brand_title.setObjectName("Brand")
        brand_sub = QLabel("مهر خلیج فارس")
        brand_sub.setObjectName("BrandSub")
        brand_text.addWidget(brand_title)
        brand_text.addWidget(brand_sub)
        brand_row.addLayout(brand_text)
        brand_row.addStretch(1)
        sidebar_layout.addLayout(brand_row)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        for section, items in NAV_ITEMS:
            section_label = QLabel(section)
            section_label.setStyleSheet(
                f"color: {self.colors['sidebar_muted']}; font-size: 8pt; font-weight: 700;"
                f"padding: 10px 10px 3px 10px;"
            )
            sidebar_layout.addWidget(section_label)
            for key, label, icon in items:
                button = QPushButton(f"  {label}")
                button.setObjectName("NavButton")
                button.setCheckable(True)
                button.setIcon(make_icon(icon, self.colors["sidebar_text"], 18))
                button.setIconSize(QSize(18, 18))
                button.setMinimumHeight(38)
                button.setCursor(Qt.PointingHandCursor)
                button.clicked.connect(lambda _checked=False, page=key: self.show_page(page))
                self.nav_group.addButton(button)
                self.nav_buttons[key] = button
                sidebar_layout.addWidget(button)
        sidebar_layout.addStretch(1)

        self.sidebar_status = QLabel("در حال راه‌اندازی…")
        self.sidebar_status.setWordWrap(True)
        self.sidebar_status.setStyleSheet(
            f"color: {self.colors['sidebar_muted']}; font-size: 8pt; padding: 8px 8px 4px 8px;"
        )
        sidebar_layout.addWidget(self.sidebar_status)
        version_label = QLabel(f"نسخهٔ {fa(APP_VERSION)}")
        version_label.setStyleSheet(
            f"color: {self.colors['sidebar_muted']}; font-size: 7.5pt; padding: 0 8px 4px 8px;"
        )
        sidebar_layout.addWidget(version_label)
        body.addWidget(sidebar)

        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        root.addLayout(body, 1)

        # ---- صفحه‌ها ----
        self.dashboard = DashboardPage(self.notes)
        self.contracts = ContractsPage(self.notes)
        self.reports = ReportsPage(self.colors)
        self.pdfs = PdfLibraryPage(self.colors)
        self.settings = SettingsPage(self.config, self.colors)
        self.help = HelpPage(self.colors)
        self.pages = {
            "dashboard": self.dashboard,
            "contracts": self.contracts,
            "reports": self.reports,
            "pdfs": self.pdfs,
            "settings": self.settings,
            "help": self.help,
        }
        for page in self.pages.values():
            self.stack.addWidget(page)

        # ---- پنل جزئیات ----
        self.detail = ContractDetailDock(self.notes, self.colors,
                                         str(self.config.get("last_author") or ""), self)
        self.addDockWidget(Qt.RightDockWidgetArea, self.detail)
        self.detail.hide()
        self.detail.bookmarkToggled.connect(self.toggle_bookmark)
        self.detail.pdfRequested.connect(self.open_pdf)
        self.detail.revealRequested.connect(self.reveal_pdf)
        self.detail.notesChanged.connect(self._on_notes_changed)

        # ---- نوار وضعیت ----
        status = self.statusBar()
        self.status_source = QLabel("منبع: —")
        self.status_counts = QLabel("—")
        self.status_loaded = QLabel("—")
        self.loading_bar = QProgressBar()
        self.loading_bar.setRange(0, 0)
        self.loading_bar.setFixedWidth(110)
        self.loading_bar.setTextVisible(False)
        self.loading_bar.hide()
        status.addWidget(self.status_source, 1)
        status.addPermanentWidget(self.status_counts)
        status.addPermanentWidget(self.status_loaded)
        status.addPermanentWidget(self.loading_bar)

        # ---- اتصال رخدادها ----
        self.dashboard.cardClicked.connect(self.on_card_clicked)
        self.dashboard.sliceClicked.connect(self.on_slice_clicked)
        self.dashboard.showContracts.connect(lambda: self.show_page("contracts"))
        self.dashboard.refreshRequested.connect(self._on_refresh_requested)
        self.dashboard.pickRequested.connect(self.pick_excel_file)
        self.dashboard.recordActivated.connect(self.show_record_detail)
        self.contracts.recordActivated.connect(self.show_record_detail)
        self.contracts.searchChanged.connect(self._on_page_search)
        self.contracts.filterChanged.connect(self._on_filter_changed)
        self.contracts.exportRequested.connect(self.export_table)
        self.contracts.bookmarkToggled.connect(self.toggle_bookmark)
        self.contracts.pdfRequested.connect(self.open_pdf)
        self.contracts.copyRequested.connect(self.copy_text)
        self.reports.exportExcelRequested.connect(self.export_management)
        self.reports.printRequested.connect(self.print_report)
        self.reports.refreshRequested.connect(self._refresh_reports)
        self.pdfs.openRequested.connect(self.open_pdf)
        self.pdfs.folderRequested.connect(lambda: self.show_page("settings"))
        self.settings.saveRequested.connect(self.save_settings)
        self.settings.reloadRequested.connect(lambda: self.load_data(force_settings=True))
        self.settings.openDataFolderRequested.connect(
            lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(data_dir())))
        )
        self._set_page("dashboard")

    def _create_actions(self) -> None:
        file_menu = self.menuBar().addMenu("فایل")
        open_action = QAction("انتخاب فایل اکسل…", self)
        open_action.triggered.connect(self.pick_excel_file)
        file_menu.addAction(open_action)
        reload_action = QAction("بازخوانی داده‌ها", self)
        reload_action.setShortcut(QKeySequence("Ctrl+R"))
        reload_action.triggered.connect(lambda: self.load_data())
        file_menu.addAction(reload_action)
        file_menu.addSeparator()
        export_action = QAction("خروجی اکسل جدول…", self)
        export_action.setShortcut(QKeySequence("Ctrl+E"))
        export_action.triggered.connect(self.export_table)
        file_menu.addAction(export_action)
        report_action = QAction("خروجی اکسل گزارش مدیریتی…", self)
        report_action.triggered.connect(self.export_management)
        file_menu.addAction(report_action)
        file_menu.addSeparator()
        quit_action = QAction("خروج", self)
        quit_action.setShortcut(QKeySequence.Quit)
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        view_menu = self.menuBar().addMenu("نمایش")
        search_action = QAction("جستجو", self)
        search_action.setShortcut(QKeySequence.Find)
        search_action.triggered.connect(self.focus_search)
        view_menu.addAction(search_action)
        theme_action = QAction("تغییر پوسته", self)
        theme_action.setShortcut(QKeySequence("Ctrl+D"))
        theme_action.triggered.connect(self.toggle_theme)
        view_menu.addAction(theme_action)
        details_action = QAction("پنل جزئیات قرارداد", self)
        details_action.setShortcut(QKeySequence("Ctrl+I"))
        details_action.triggered.connect(self._toggle_detail_dock)
        view_menu.addAction(details_action)
        view_menu.addSeparator()
        for key in ("dashboard", "contracts", "reports", "pdfs", "settings", "help"):
            action = QAction(PAGE_TITLES[key], self)
            action.triggered.connect(lambda _checked=False, page=key: self.show_page(page))
            view_menu.addAction(action)

        help_menu = self.menuBar().addMenu("راهنما")
        help_action = QAction("راهنمای برنامه", self)
        help_action.setShortcut(QKeySequence.HelpContents)
        help_action.triggered.connect(lambda: self.show_page("help"))
        help_menu.addAction(help_action)
        about_action = QAction("دربارهٔ برنامه", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        escape_shortcut = QShortcut(QKeySequence("Esc"), self)
        escape_shortcut.activated.connect(self._on_escape)

    # ------------------------------------------------------------------
    # ناوبری
    # ------------------------------------------------------------------
    def _set_page(self, key: str) -> None:
        self._current_page = key
        button = self.nav_buttons.get(key)
        if button:
            button.setChecked(True)
        self.page_title_label.setText(PAGE_TITLES.get(key, ""))
        self.stack.setCurrentWidget(self.pages[key])

    def show_page(self, key: str) -> None:
        if key not in self.pages:
            return
        self._set_page(key)
        if key == "reports":
            self._refresh_reports()
        elif key == "pdfs":
            self.pdfs.set_folder(self.data.pdf_folder)
        elif key == "settings":
            self.settings.load_values(self.config)

    def _refresh_reports(self) -> None:
        self.reports.update_data(self.data)

    def focus_search(self) -> None:
        self.search_edit.setFocus()
        self.search_edit.selectAll()

    def _on_escape(self) -> None:
        if self.detail.isVisible():
            self.detail.hide()
        elif self._current_page != "dashboard":
            self.show_page("dashboard")

    def _toggle_detail_dock(self) -> None:
        if self.detail.isVisible():
            self.detail.hide()
            return
        records = self.contracts.model.visible_records()
        if records:
            self.show_record_detail(records[0])
        else:
            self.detail.show()

    # ------------------------------------------------------------------
    # داده‌ها
    # ------------------------------------------------------------------
    def initial_load(self) -> None:
        path = str(self.config.get("excel_path") or "").strip()
        if not path:
            path = find_default_excel(self.config)
            if path:
                self.config.set("excel_path", path)
                pdf_folder = default_pdf_folder(path, self.config)
                if pdf_folder and not self.config.get("pdf_folder"):
                    self.config.set("pdf_folder", pdf_folder, save=False)
                self.config.save()
        if not path:
            self.dashboard.update_data(empty_dashboard("فایل اکسل انتخاب نشده است."))
            self.settings.load_values(self.config)
            self.status_source.setText("منبع: —")
            self.status_counts.setText("داده‌ای بارگذاری نشده")
            self.sidebar_status.setText("برای شروع، فایل اکسل را از تنظیمات انتخاب کنید.")
            return
        self.load_data(silent=False)

    def load_data(self, silent: bool = False, force_settings: bool = False) -> None:
        if self._loader_running():
            return
        if force_settings:
            self._persist_settings_fields()
        path = str(self.config.get("excel_path") or "").strip()
        if not path:
            self.pick_excel_file()
            return
        sheet = str(self.config.get("sheet_name") or SHEET_NAME_DEFAULT)
        pdf_folder = str(self.config.get("pdf_folder") or "")
        self._set_loading(True)
        self.refresh_button.setEnabled(False)
        self.statusBar().showMessage("در حال خواندن فایل اکسل…", 0)
        self.loader = ExcelLoader(path, sheet, pdf_folder, self)
        self.loader.loaded.connect(lambda data: self._on_data_loaded(data, silent))
        self.loader.failed.connect(lambda message: self._on_load_failed(message, silent))
        self.loader.finished.connect(self._on_loader_finished)
        self.loader.finished.connect(self.loader.deleteLater)
        self.loader.start()

    def _persist_settings_fields(self) -> None:
        self.config.set("excel_path", self.settings.excel_edit.text().strip(), save=False)
        self.config.set("sheet_name", self.settings.sheet_edit.text().strip() or SHEET_NAME_DEFAULT,
                        save=False)
        self.config.set("pdf_folder", self.settings.pdf_edit.text().strip(), save=False)
        self.config.set("auto_refresh_minutes", self.settings.auto_spin.value(), save=False)
        self.config.save()

    def _loader_running(self) -> bool:
        loader = self.loader
        if loader is None:
            return False
        try:
            return bool(loader.isRunning())
        except RuntimeError:  # شیء C++ حذف شده است
            self.loader = None
            return False

    def _set_loading(self, loading: bool) -> None:
        self.loading_bar.setVisible(loading)
        self.refresh_button.setText("در حال خواندن…" if loading else "به‌روزرسانی")
        if loading:
            self.sidebar_status.setText("در حال خواندن اکسل…")

    def _on_loader_finished(self) -> None:
        self._set_loading(False)
        self.refresh_button.setEnabled(True)
        self.loader = None

    def _on_data_loaded(self, data: DashboardData, silent: bool) -> None:
        self.data = data
        self.dashboard.update_data(data)
        self.contracts.model.set_records(data.records)
        self.apply_view()
        self.reports.update_data(data)
        self.pdfs.set_folder(data.pdf_folder)
        self.detail.setEnabled(bool(data.records))
        self.status_source.setText(f"منبع: {data.excel_path}")
        self.status_counts.setText(bidi(f"{fa(data.total_contracts)} قرارداد"))
        self.status_loaded.setText(f"بارگذاری: {data.loaded_at_text}")
        self.sidebar_status.setText(bidi(f"{fa(data.total_contracts)} قرارداد") + f"\n{data.loaded_at_text}")
        self.statusBar().showMessage(
            f"داده‌ها با موفقیت بارگذاری شد — {fa(data.total_contracts)} قرارداد"
            + (f" | {len(data.warnings)} هشدار" if data.warnings else ""),
            6000,
        )
        self._setup_auto_timer()
        if not silent and data.warnings:
            QMessageBox.warning(self, "هشدار داده‌ها", "\n".join(data.warnings))

    def _on_load_failed(self, message: str, silent: bool) -> None:
        self._last_error = message
        self.statusBar().showMessage("بارگذاری داده‌ها ناموفق بود.", 8000)
        self.sidebar_status.setText("داده‌ها بارگذاری نشد.")
        if not self.data.records:
            empty = empty_dashboard(message.splitlines()[0] if message else "خطا در بارگذاری")
            self.dashboard.update_data(empty)
            self.reports.update_data(empty)
        if silent:
            return
        box = QMessageBox(self)
        box.setWindowTitle("خطا در خواندن داده‌ها")
        box.setIcon(QMessageBox.Critical)
        box.setText("خواندن فایل اکسل ممکن نشد.")
        box.setInformativeText(message)
        pick_button = box.addButton("انتخاب فایل اکسل", QMessageBox.AcceptRole)
        settings_button = box.addButton("تنظیمات", QMessageBox.ActionRole)
        box.addButton("بستن", QMessageBox.RejectRole)
        box.exec_()
        clicked = box.clickedButton()
        if clicked is pick_button:
            self.pick_excel_file()
        elif clicked is settings_button:
            self.show_page("settings")

    def _setup_auto_timer(self) -> None:
        if not hasattr(self, "auto_timer"):
            self.auto_timer = QTimer(self)
            self.auto_timer.timeout.connect(self._auto_refresh)
        minutes = int(self.config.get("auto_refresh_minutes") or 0)
        self.auto_timer.stop()
        if minutes > 0:
            self.auto_timer.start(minutes * 60 * 1000)

    def _auto_refresh(self) -> None:
        path = self.data.excel_path or str(self.config.get("excel_path") or "")
        if not path or not os.path.exists(path):
            return
        if self._loader_running():
            return
        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(path))
        except OSError:
            return
        if self.data.source_mtime is None or mtime > self.data.source_mtime:
            self.statusBar().showMessage("تغییر فایل اکسل شناسایی شد؛ بازخوانی…", 4000)
            self.load_data(silent=True)

    def pick_excel_file(self) -> None:
        start = (str(self.config.get("excel_path") or "") or str(app_dir()))
        if os.path.isfile(start):
            start = os.path.dirname(start)
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل", start, EXCEL_FILTER)
        if not path:
            return
        self.config.set("excel_path", path)
        pdf_folder = str(self.config.get("pdf_folder") or "")
        if not pdf_folder:
            guess = default_pdf_folder(path, self.config)
            if guess:
                self.config.set("pdf_folder", guess, save=False)
                self.config.save()
        self.settings.load_values(self.config)
        self.load_data()

    def save_settings(self, values: Dict[str, Any]) -> None:
        for key, value in values.items():
            self.config.set(key, value, save=False)
        self.config.save()
        theme = str(values.get("theme") or "light")
        font_size = int(values.get("font_size") or 10)
        self.apply_theme(theme, font_size, persist=False)
        self._setup_auto_timer()
        self.statusBar().showMessage("تنظیمات ذخیره شد.", 4000)
        self.load_data()

    def _on_refresh_requested(self) -> None:
        if not str(self.config.get("excel_path") or "").strip():
            self.show_page("settings")
            self.pick_excel_file()
            return
        self.load_data()

    # ------------------------------------------------------------------
    # جستجو و فیلتر
    # ------------------------------------------------------------------
    def _on_search_text_changed(self, text: str) -> None:
        if not hasattr(self, "search_timer"):
            self.search_timer = QTimer(self)
            self.search_timer.setSingleShot(True)
            self.search_timer.timeout.connect(self.apply_view)
        self.search_timer.start(220)
        if self.contracts.search_edit.text() != text:
            self.contracts.search_edit.blockSignals(True)
            self.contracts.search_edit.setText(text)
            self.contracts.search_edit.blockSignals(False)

    def _on_page_search(self, text: str) -> None:
        if self.search_edit.text() != text:
            self.search_edit.blockSignals(True)
            self.search_edit.setText(text)
            self.search_edit.blockSignals(False)
        self.apply_view()

    def _on_filter_changed(self, key: str) -> None:
        self.apply_view(filter_key=key)

    def apply_view(self, filter_key: Optional[str] = None) -> None:
        search = self.search_edit.text().strip() or self.contracts.current_search().strip()
        key = filter_key if filter_key is not None else self.contracts.current_filter()
        bookmarks_only = self.contracts.bookmark_check.isChecked()
        self.contracts.model.apply_view(search, key, bookmarks_only)
        self.dashboard.apply_filter(search, key, bookmarks_only)
        self.contracts.refresh_counts()

    def on_card_clicked(self, filter_key: str) -> None:
        key = filter_key or "all"
        label = dict(FILTER_CHOICES).get(key)
        self.contracts.set_filter(key, label)
        self.apply_view(filter_key=key)
        self.show_page("contracts")

    def on_slice_clicked(self, field: str, label: str) -> None:
        key = f"{field}:{label}"
        prefix = FIELD_LABELS.get(field, field)
        self.contracts.set_filter(key, f"{prefix}: {label}")
        self.apply_view(filter_key=key)
        self.show_page("contracts")

    # ------------------------------------------------------------------
    # جزئیات، یادداشت و PDF
    # ------------------------------------------------------------------
    def show_record_detail(self, record: Optional[ContractRecord]) -> None:
        if record is None:
            return
        self.detail.set_record(record)
        self.detail.show()
        self.detail.raise_()

    def toggle_bookmark(self, key: str) -> None:
        if not key:
            return
        flagged = self.notes.toggle_bookmark(key)
        self.contracts.model.refresh_bookmarks()
        self.dashboard.preview_model.refresh_bookmarks()
        record = self.data.find(key)
        if record and self.detail.record is not None and self.detail.record.key == key:
            self.detail.set_record(record)
        self.statusBar().showMessage(
            "قرارداد نشان‌گذاری شد." if flagged else "نشان قرارداد برداشته شد.", 3000
        )

    def _on_notes_changed(self, key: str) -> None:
        self.contracts.model.refresh_bookmarks()
        self.dashboard.preview_model.refresh_bookmarks()
        self.config.set("last_author", self.detail.author_edit.text().strip(), save=False)
        self.config.save()
        self.statusBar().showMessage("یادداشت ذخیره شد.", 3000)

    def _pdf_path(self, filename: str) -> str:
        folder = self.data.pdf_folder or str(self.config.get("pdf_folder") or "")
        if not folder:
            return ""
        return os.path.join(folder, filename)

    def open_pdf(self, filename: str) -> None:
        if not filename:
            return
        path = self._pdf_path(filename)
        if not path or not os.path.exists(path):
            QMessageBox.information(
                self, "فایل PDF",
                "فایل پیدا نشد.\nمسیر پوشهٔ PDF را در تنظیمات بررسی کنید.",
            )
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def reveal_pdf(self, filename: str) -> None:
        if not filename:
            return
        path = self._pdf_path(filename)
        if not path or not os.path.exists(path):
            QMessageBox.information(self, "فایل PDF", "فایل پیدا نشد.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))

    def copy_text(self, text: str) -> None:
        if not text:
            return
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage("در کلیپ‌بورد کپی شد.", 2500)

    # ------------------------------------------------------------------
    # خروجی‌ها
    # ------------------------------------------------------------------
    def export_table(self) -> None:
        if not self.data.records:
            QMessageBox.information(self, "خروجی اکسل", "هنوز داده‌ای بارگذاری نشده است.")
            return
        visible = self.contracts.model.visible_records()
        selected = self.contracts.selected_records()
        dialog = ExportDialog(EXPORT_COLUMN_TITLES,
                              len(visible), len(selected), len(self.data.records), self)
        if dialog.exec_() != QDialog.Accepted:
            return
        scope = dialog.scope()
        if scope == "all":
            records = self.data.records
        elif scope == "selected":
            records = selected
        else:
            records = visible
        if not records:
            QMessageBox.information(self, "خروجی اکسل", "ردیفی برای خروجی وجود ندارد.")
            return
        default_name = f"قراردادها-{jalali_filename_stamp()}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "ذخیرهٔ خروجی اکسل",
                                              str(app_dir() / default_name), "فایل اکسل (*.xlsx)")
        if not path:
            return
        try:
            count = write_contracts_workbook(path, records, dialog.selected_columns())
        except Exception as exc:
            QMessageBox.critical(self, "خطا در خروجی", f"نوشتن فایل ممکن نشد:\n{exc}")
            return
        self._offer_open(path, f"{fa(count)} ردیف در فایل اکسل ذخیره شد.")

    def export_management(self) -> None:
        if not self.data.records:
            QMessageBox.information(self, "گزارش مدیریتی", "هنوز داده‌ای بارگذاری نشده است.")
            return
        default_name = f"گزارش-مدیریتی-مهر-{jalali_filename_stamp()}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "ذخیرهٔ گزارش مدیریتی",
                                              str(app_dir() / default_name), "فایل اکسل (*.xlsx)")
        if not path:
            return
        try:
            write_management_report(path, self.data)
        except Exception as exc:
            QMessageBox.critical(self, "خطا در گزارش", f"نوشتن فایل ممکن نشد:\n{exc}")
            return
        self._offer_open(path, "گزارش مدیریتی با سه شیت ذخیره شد.")

    def _offer_open(self, path: str, message: str) -> None:
        box = QMessageBox(self)
        box.setWindowTitle("انجام شد")
        box.setIcon(QMessageBox.Information)
        box.setText(message)
        box.setInformativeText(path)
        open_button = box.addButton("باز کردن فایل", QMessageBox.AcceptRole)
        box.addButton("بستن", QMessageBox.RejectRole)
        box.exec_()
        if box.clickedButton() is open_button:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def print_report(self) -> None:
        if not self.data.records:
            QMessageBox.information(self, "چاپ گزارش", "هنوز داده‌ای بارگذاری نشده است.")
            return
        try:
            from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
        except Exception:
            QMessageBox.warning(self, "چاپ", "پشتیبانی چاپ در این سیستم نصب نیست.")
            return
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        dialog = QPrintDialog(printer, self)
        dialog.setWindowTitle("چاپ / ذخیرهٔ PDF گزارش مدیریتی")
        if dialog.exec_() != QDialog.Accepted:
            return
        from PyQt5.QtGui import QTextDocument
        document = QTextDocument()
        option = document.defaultTextOption()
        option.setTextDirection(Qt.RightToLeft)
        document.setDefaultTextOption(option)
        document.setHtml(self.reports.html())
        document.print_(printer)
        self.statusBar().showMessage("گزارش به چاپ ارسال شد.", 4000)

    # ------------------------------------------------------------------
    # پوسته و تنظیمات ظاهری
    # ------------------------------------------------------------------
    def apply_theme(self, theme: str, font_size: int = 10, persist: bool = True) -> None:
        app = QApplication.instance()
        self.colors = apply_theme(app, theme, font_size)
        if persist:
            self.config.set("theme", theme, save=False)
            self.config.set("font_size", font_size, save=False)
            self.config.save()
        for widget in (self.dashboard, self.contracts, self.reports,
                       self.pdfs, self.settings, self.help):
            if hasattr(widget, "set_theme_colors"):
                widget.set_theme_colors(self.colors)
        if hasattr(self, "detail"):
            self.detail.set_theme_colors(self.colors)
        for combo in self.findChildren(Combo):
            combo.set_arrow_color(self.colors["text_muted"])
        for check in self.findChildren(CheckBox):
            check.set_theme_colors(self.colors)
        self._update_header_icons()
        if hasattr(self, "nav_buttons"):
            for button in self.nav_buttons.values():
                button.setIcon(make_icon(self._nav_icon_kind(button), self.colors["sidebar_text"], 18))

    def _nav_icon_kind(self, button: QPushButton) -> str:
        for section, items in NAV_ITEMS:
            for key, label, icon in items:
                if label in button.text():
                    return icon
        return "document"

    def _update_header_icons(self) -> None:
        if not hasattr(self, "theme_button"):
            return
        is_dark = str(self.config.get("theme") or "light") == "dark"
        self.theme_button.setIcon(make_icon("sun" if is_dark else "moon", self.colors["text"], 18))
        self.theme_button.setToolTip("پوستهٔ روشن" if is_dark else "پوستهٔ تیره")
        self.refresh_button.setIcon(make_icon("refresh", "#ffffff", 16))
        self.export_button.setIcon(make_icon("download", self.colors["text"], 16))
        if hasattr(self, "search_action"):
            self.search_action.setIcon(make_icon("search", self.colors["text_muted"], 16))
        self.export_button.setIconSize(QSize(16, 16))
        self.refresh_button.setIconSize(QSize(16, 16))
        self.theme_button.setIconSize(QSize(18, 18))

    def toggle_theme(self) -> None:
        current = str(self.config.get("theme") or "light")
        new_theme = "dark" if current == "light" else "light"
        self.apply_theme(new_theme, int(self.config.get("font_size") or 10))
        self.config.set("theme", new_theme)

    def show_about(self) -> None:
        QMessageBox.about(
            self, "دربارهٔ برنامه",
            f"<div dir='rtl' style='font-family: Tahoma;'>"
            f"<h3>{APP_NAME}</h3>"
            f"<p>نسخهٔ {APP_VERSION} — برنامهٔ دسکتاپ تک‌فایلی بر پایهٔ PyQt5</p>"
            f"<p>داده‌ها مستقیماً از فایل اکسل (شیت colorcode) خوانده می‌شوند و "
            f"به Flask، مرورگر یا فایل HTML نیازی نیست.</p>"
            f"<p style='color:#66756f;'>مسیر تنظیمات: {self.config.path}</p>"
            f"</div>",
        )

    # ------------------------------------------------------------------
    # رخدادهای پنجره
    # ------------------------------------------------------------------
    def _restore_geometry(self) -> None:
        stored = str(self.config.get("window_geometry") or "")
        if stored:
            try:
                if self.restoreGeometry(QByteArray.fromBase64(stored.encode("ascii"))):
                    return
            except Exception:
                pass
        screen = QApplication.primaryScreen()
        if screen is not None:
            geometry = screen.availableGeometry()
            width = min(1480, int(geometry.width() * 0.92))
            height = min(940, int(geometry.height() * 0.92))
            self.resize(width, height)
            self.move(geometry.center().x() - width // 2, geometry.center().y() - height // 2)

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith((".xlsx", ".xlsm", ".xls")):
                    event.acceptProposedAction()
                    return
        super().dragEnterEvent(event)

    def dropEvent(self, event) -> None:  # noqa: N802
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith((".xlsx", ".xlsm", ".xls")):
                self.config.set("excel_path", path)
                self.settings.load_values(self.config)
                self.load_data()
                event.acceptProposedAction()
                return
        super().dropEvent(event)

    def closeEvent(self, event) -> None:  # noqa: N802
        try:
            self.config.set("window_geometry",
                            bytes(self.saveGeometry().toBase64()).decode("ascii"), save=False)
            self.config.set("last_author", self.detail.author_edit.text().strip(), save=False)
            self.config.save()
            self.notes.save()
        except Exception:
            pass
        loader = self.loader
        if loader is not None:
            try:
                if loader.isRunning():
                    loader.requestInterruption()
                    if not loader.wait(4000):
                        loader.terminate()
                        loader.wait(1000)
            except RuntimeError:
                pass
            self.loader = None
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# اجرای برنامه
# ---------------------------------------------------------------------------
def write_error_log(detail: str) -> Optional[Path]:
    """ذخیرهٔ خطاها در فایل لاگ کنار برنامه برای عیب‌یابی."""
    try:
        path = data_dir() / "mehr_error.log"
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n")
            handle.write(detail.rstrip() + "\n")
        return path
    except Exception:
        return None


def run_check(excel: str, sheet: str, pdf_folder: str) -> int:
    """بررسی سریع داده‌ها بدون باز کردن پنجره (برای عیب‌یابی)."""
    try:
        data = read_excel_workbook(excel, sheet, pdf_folder)
    except ExcelDataError as exc:
        print(f"❌ {exc}")
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"❌ خطای غیرمنتظره: {exc}")
        write_error_log(traceback.format_exc())
        return 3
    print("✅ خواندن اکسل موفق بود")
    print(f"مسیر: {data.excel_path}")
    print(f"شیت: {data.sheet_name}")
    print(f"تعداد قرارداد: {data.total_contracts}")
    print(f"پیشرفت کل: {data.progress_percent_text}%")
    print(f"کسورات: {data.contracts_with_deductions}")
    print("وضعیت‌ها: " + "، ".join(
        f"{FLAG_LABELS[key]}={count}" for key, count in data.flag_counts.items()
    ))
    print(f"فایل‌های PDF متصل‌شده: {sum(1 for r in data.records if r.pdf_file)}")
    for warning in data.warnings:
        print(f"⚠️ {warning}")
    print("توزیع نوع قرارداد: " + "، ".join(
        f"{label}={value}" for label, value in data.dist("type").as_list()[:5]
    ))
    return 0


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="داشبورد قراردادهای مهر خلیج فارس — نسخهٔ دسکتاپ PyQt5",
    )
    parser.add_argument("--excel", help="مسیر فایل اکسل (بر تنظیمات ذخیره‌شده اولویت دارد)")
    parser.add_argument("--sheet", help=f"نام شیت (پیش‌فرض: {SHEET_NAME_DEFAULT})")
    parser.add_argument("--pdf-folder", help="مسیر پوشهٔ فایل‌های PDF")
    parser.add_argument("--check", action="store_true",
                        help="فقط بررسی خواندن داده‌ها؛ پنجره باز نمی‌شود")
    parser.add_argument("--theme", choices=["light", "dark"], help="پوستهٔ برنامه")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {APP_VERSION}")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    # حالت بررسی: بدون نیاز به نمایشگر یا QApplication
    if args.check:
        config = AppConfig()
        if args.excel:
            config.values["excel_path"] = os.path.abspath(args.excel)
        if args.sheet:
            config.values["sheet_name"] = args.sheet
        if args.pdf_folder:
            config.values["pdf_folder"] = os.path.abspath(args.pdf_folder)
        excel = str(config.get("excel_path") or find_default_excel(config) or "")
        sheet = str(config.get("sheet_name") or SHEET_NAME_DEFAULT)
        pdf_folder = str(config.get("pdf_folder") or default_pdf_folder(excel, config))
        if not excel:
            print("❌ مسیر فایل اکسل مشخص نشده است. از --excel استفاده کنید.")
            return 2
        return run_check(excel, sheet, pdf_folder)

    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv[:1])
    load_custom_fonts()
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("MehrKando")

    config = AppConfig()
    if args.excel:
        config.values["excel_path"] = os.path.abspath(args.excel)
    if args.sheet:
        config.values["sheet_name"] = args.sheet
    if args.pdf_folder:
        config.values["pdf_folder"] = os.path.abspath(args.pdf_folder)
    if args.theme:
        config.values["theme"] = args.theme

    notes = NotesStore()

    window = MainWindow(config, notes)
    window.show()
    return app.exec_()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # pragma: no cover
        import traceback as _tb
        _tb.print_exc()
        log_path = write_error_log(_tb.format_exc())
        try:
            app = QApplication.instance() or QApplication(sys.argv[:1])
            message = f"برنامه با خطا متوقف شد:\n{exc}"
            if log_path:
                message += f"\n\nجزئیات در این فایل ذخیره شد:\n{log_path}"
            QMessageBox.critical(None, "خطای اجرای برنامه", message)
        except Exception:
            pass
        sys.exit(1)

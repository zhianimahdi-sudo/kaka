#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
داشبورد قراردادهای مهر خلیج فارس — نسخهٔ دسکتاپ (PyQt5)
=========================================================

این برنامه تبدیل نسخهٔ Flask/dashboard.html به یک برنامهٔ دسکتاپ تک‌فایلی است:

* دیگر به Flask، مرورگر، dashboard.html یا فایل mehr_optimized.py نیاز ندارد.
* داده‌ها به‌طور پیش‌فرض از دیتابیس مرکزی (MySQL) خوانده می‌شوند: کتاب کار اکسل
  با شناسهٔ مشخص (پیش‌فرض ۱) که در جدول‌های excel_workbooks / excel_original_files
  ذخیره شده است.
* اگر دیتابیس آماده نیست، دو راه بدون دردسر وجود دارد:
  ۱) «بارگذاری فایل اکسل» در صفحهٔ تنظیمات: فایل را در دیتابیس ذخیره می‌کند و
     دیتابیس/جدول‌های لازم را در صورت نبود خودکار می‌سازد.
  ۲) «اجرای محلی (بدون دیتابیس)»: فایل اکسل مستقیم از دیسک خوانده می‌شود.
* رمز دیتابیس اجباری نیست؛ اگر رمز وارد نشود یا رد شود، ورودهای بدون رمز
  (مثل root با رمز خالی) هم امتحان می‌شوند.
* نمودارها، فیلترها، جستجو، جزئیات قرارداد، یادداشت‌ها، خروجی اکسل، گزارش
  مدیریتی و اتصال فایل PDF همه درون همین فایل پیاده‌سازی شده‌اند.

اجرا:
    python mehr_dashboard_qt.py

پیش‌نیازها:
    pip install PyQt5 openpyxl mysql-connector-python

گزینه‌های خط فرمان:
    --db-host / --db-port / --db-name / --db-user / --db-password
    --workbook-id    شناسهٔ کتاب کار در دیتابیس (پیش‌فرض ۱)
    --import-excel   بارگذاری فایل اکسل در دیتابیس (ساخت دیتابیس از ابتدا) و خروج
    --local-excel    اجرای برنامه با فایل اکسل محلی، بدون دیتابیس
    --check          فقط اتصال به دیتابیس و آمار را بررسی کن (بدون باز کردن پنجره)
"""

from __future__ import annotations

import argparse
import base64
import ctypes
import faulthandler
import functools
import hashlib
import html as html_module
import io
import json
import math
import os
import platform
import subprocess
import re
import sys
import tempfile
import threading
import time
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
        QEasingCurve,
        QModelIndex,
        QPoint,
        QPointF,
        QRectF,
        QSize,
        QPropertyAnimation,
        QThread,
        Qt,
        QTimer,
        QUrl,
        pyqtProperty,
        pyqtSignal,
    )
    from PyQt5.QtGui import (
        QBrush,
        QColor,
        QConicalGradient,
        QRadialGradient,
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
        QApplication,
            QCheckBox,
        QDialog,
        QFileDialog,
            QFrame,
        QGridLayout,
        QHBoxLayout,
        QHeaderView,
        QInputDialog,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMenu,
        QMessageBox,
        QPlainTextEdit,
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
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:  # pragma: no cover
    _message = (
        "خطا: کتابخانهٔ PyQt5 نصب نیست یا بارگذاری نشد.\n\n"
        "در پنجرهٔ فرمان (CMD) این دستور را اجرا کنید:\n"
        "    pip install PyQt5 openpyxl mysql-connector-python\n\n"
        f"جزئیات فنی: {exc}"
    )
    sys.stderr.write(_message + "\n")
    try:  # پنجرهٔ خطای ویندوز تا پیام با بستن کنسول گم نشود
        if os.name == "nt":
            ctypes.windll.user32.MessageBoxW(None, _message, "داشبورد مهر — خطای اجرا", 0x10)
    except Exception:
        pass
    raise SystemExit(1)

try:
    import mysql.connector as mysql_connector
except Exception:  # pragma: no cover - نصب اختیاری
    mysql_connector = None
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
APP_VERSION = "3.5.8"
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
        "#5A4FE0", "#8E7BFF", "#B3A6FF", "#D9D2FF", "#7C6CF0",
        "#E3A33C", "#F3C570", "#3B72D9", "#0F9D76", "#C4634F",
    ],
    "status": ["#5A4FE0", "#8E7BFF", "#C3B9FF", "#E3A33C", "#2FB59B"],
    "progress": ["#D9484F", "#E3A33C", "#E8C96A", "#8E7BFF", "#5A4FE0"],
    "completion": ["#D9484F", "#E3A33C", "#8E7BFF", "#5A4FE0"],
    "annex": [
        "#5A4FE0", "#0F9D76", "#3B72D9", "#E3A33C",
        "#D9484F", "#B06A9E", "#7C6CF0", "#2FB59B",
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


class SecretStore:
    """نگهداری رمز دیتابیس به‌صورت رمزنگاری‌شده (DPAPI در ویندوز)."""

    FILENAME = "mehr_db_secret.bin"
    _KEY = b"mehr-dashboard-db-secret-v1"

    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = Path(path) if path else (data_dir() / self.FILENAME)

    @staticmethod
    def _dpapi(data: bytes, protect: bool) -> bytes:
        class DATA_BLOB(ctypes.Structure):
            _fields_ = [("cbData", ctypes.c_uint32),
                        ("pbData", ctypes.POINTER(ctypes.c_byte))]

        buffer = ctypes.create_string_buffer(data)
        blob_in = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
        blob_out = DATA_BLOB()
        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
        ok = fn(ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out))
        if not ok:
            raise OSError("DPAPI operation failed")
        try:
            return ctypes.string_at(blob_out.pbData, blob_out.cbData)
        finally:
            kernel32.LocalFree(blob_out.pbData)

    def _mask(self, data: bytes) -> bytes:
        key = self._KEY
        return bytes(byte ^ key[index % len(key)] for index, byte in enumerate(data))

    def save(self, password: str) -> None:
        if not password:
            self.clear()
            return
        raw = password.encode("utf-8")
        try:
            payload = self._dpapi(raw, protect=True) if os.name == "nt" else self._mask(raw)
            prefix = b"DPAPI1" if os.name == "nt" else b"LOCAL1"
        except Exception:
            payload = self._mask(raw)
            prefix = b"LOCAL1"
        try:
            self.path.write_bytes(prefix + base64.b64encode(payload))
        except Exception as exc:  # pragma: no cover
            print(f"هشدار: ذخیرهٔ رمز دیتابیس ممکن نشد ({exc})")

    def load(self) -> str:
        try:
            if not self.path.exists():
                return ""
            blob = self.path.read_bytes()
            prefix, payload = blob[:6], base64.b64decode(blob[6:])
            if prefix == b"DPAPI1":
                raw = self._dpapi(payload, protect=False)
            else:
                raw = self._mask(payload)
            return raw.decode("utf-8")
        except Exception:
            return ""

    def clear(self) -> None:
        try:
            self.path.unlink(missing_ok=True)
        except Exception:
            pass


class ExcelDataError(Exception):
    """خطای خواندن یا ساختار نامعتبر فایل اکسل."""


MANAGER_FOLDER_NAME = "CentralDatabaseManager"
DB_DEFAULTS: Dict[str, Any] = {
    "host": "127.0.0.1",
    "port": 3306,
    "database": "master_data",
    "user": "master_app",
    "workbook_id": 1,
}


def manager_app_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / MANAGER_FOLDER_NAME


def read_manager_settings() -> Dict[str, Any]:
    """خواندن تنظیمات اتصال از برنامهٔ «مدیریت دیتابیس مرکزی»."""
    path = manager_app_dir() / "settings.json"
    result: Dict[str, Any] = {}
    try:
        if path.exists():
            stored = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(stored, dict):
                result = {
                    "host": str(stored.get("host") or ""),
                    "port": int(stored.get("port") or 3306),
                    "database": str(stored.get("database") or ""),
                    "user": str(stored.get("app_user") or ""),
                }
    except Exception:
        return {}
    return {key: value for key, value in result.items() if value}


def read_manager_password() -> str:
    """رمز کاربر برنامه را از فایل امن مدیر دیتابیس می‌خواند (فقط ویندوز)."""
    path = manager_app_dir() / "secrets.bin"
    if os.name != "nt" or not path.exists():
        return ""
    try:
        raw = SecretStore._dpapi(path.read_bytes(), protect=False)
        stored = json.loads(raw.decode("utf-8"))
        return str(stored.get("app_password") or "")
    except Exception:
        return ""


class AppConfig:
    """تنظیمات سادهٔ JSON (قابل انتقال و قابل خواندن)."""

    FILENAME = "mehr_qt_config.json"
    DEFAULTS: Dict[str, Any] = {
        "db_host": "127.0.0.1",
        "db_port": 3306,
        "db_name": "master_data",
        "db_user": "master_app",
        "db_workbook_id": 1,
        "db_remember_password": True,
        "data_source": "database",
        "local_excel_path": "",
        "last_excel_path": "",
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
        self.secrets = SecretStore(self.path.with_name(SecretStore.FILENAME))
        self.load()
        if not self.values.get("db_password"):
            self.values["db_password"] = self.secrets.load()

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
            payload = {key: value for key, value in self.values.items()
                       if key != "db_password"}
            tmp = self.path.with_suffix(".json.tmp")
            with open(tmp, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=2)
            os.replace(tmp, self.path)
        except Exception as exc:
            print(f"هشدار: ذخیرهٔ تنظیمات ممکن نشد ({exc})")

    def get(self, key: str, default: Any = None) -> Any:
        return self.values.get(key, self.DEFAULTS.get(key, default))

    def set(self, key: str, value: Any, save: bool = True) -> None:
        self.values[key] = value
        if save:
            self.save()

    # ---- رمز دیتابیس ----
    def get_password(self) -> str:
        return str(self.values.get("db_password") or "")

    def set_password(self, password: str, remember: Optional[bool] = None) -> None:
        self.values["db_password"] = password
        keep = self.get("db_remember_password") if remember is None else bool(remember)
        if keep:
            self.secrets.save(password)
        else:
            self.secrets.clear()

    def db_settings(self) -> Dict[str, Any]:
        """تنظیمات اتصال و کتاب کار به‌صورت یک دیکشنری ساده."""
        return {
            "host": str(self.get("db_host") or "127.0.0.1").strip(),
            "port": int(self.get("db_port") or 3306),
            "database": str(self.get("db_name") or "master_data").strip(),
            "user": str(self.get("db_user") or "master_app").strip(),
            "password": self.get_password(),
            "workbook_id": int(self.get("db_workbook_id") or 1),
            "sheet_name": str(self.get("sheet_name") or SHEET_NAME_DEFAULT),
        }


LOG_FILENAME = "mehr_error.log"


def log_path() -> Path:
    """مسیر فایل لاگ برنامه."""
    try:
        return data_dir() / LOG_FILENAME
    except Exception:
        return Path.home() / LOG_FILENAME


ATTEMPT_FILENAME = "mehr_db_attempt.marker"
ATTEMPT_MAX_AGE = 20 * 60  # ثانیه


def attempt_marker_path() -> Path:
    """فایل نشانهٔ «در حال خواندن از دیتابیس» (برای تشخیص بسته‌شدن ناگهانی)."""
    try:
        return data_dir() / ATTEMPT_FILENAME
    except Exception:
        return Path.home() / ATTEMPT_FILENAME


def mark_db_attempt() -> None:
    """قبل از اتصال نوشته می‌شود؛ اگر برنامه کرش کند باقی می‌ماند."""
    try:
        attempt_marker_path().write_text(
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8"
        )
    except Exception:
        pass


def clear_db_attempt() -> None:
    try:
        attempt_marker_path().unlink()
    except Exception:
        pass


def previous_attempt_crashed() -> bool:
    """آیا اجرای قبلی وسطِ اتصال به دیتابیس ناگهانی بسته شد؟"""
    path = attempt_marker_path()
    try:
        if not path.exists():
            return False
        age = datetime.now().timestamp() - path.stat().st_mtime
        if age > ATTEMPT_MAX_AGE:
            return False
        return True
    except Exception:
        return False


def log_event(message: str) -> None:
    """ثبت یک رویداد یا خطا در فایل لاگ (برای عیب‌یابی)."""
    try:
        with open(log_path(), "a", encoding="utf-8") as handle:
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            handle.write(f"[{stamp}] {str(message).rstrip()}\n")
    except Exception:
        pass


PDF_FOLDER_NAMES = ("PDF-MEHR", "pdf-mehr", "PDF", "pdf", "PDFs", "pdfs", "پیوست‌ها")


def _expand_folder(value: Any) -> str:
    """مسیر پوشه را باز می‌کند (متغیر محیطی، ~ و مسیر نسبی)."""
    text = os.path.expandvars(os.path.expanduser(str(value or "").strip().strip('"').strip("'")))
    if not text:
        return ""
    if not os.path.isabs(text):
        candidate = app_dir() / text
        if candidate.exists():
            return str(candidate)
    return text


def _looks_like_contract_name(stem: str) -> bool:
    """آیا نام فایل شبیه شمارهٔ قرارداد است؟ (مثل MEHR-003-1402)"""
    text = str(stem or "").translate(_LATIN_DIGITS)
    if re.search(r"(mehr|مهر)", text, re.IGNORECASE):
        return True
    return len(re.findall(r"\d+", text)) >= 2


def _pdf_folder_score(folder: str) -> int:
    """تعداد PDFهایی که نامشان شبیه شمارهٔ قرارداد است."""
    try:
        names = scan_pdf_files(folder, depth=1, limit=300)
    except Exception:
        return 0
    score = 0
    for rel in names:
        stem = os.path.splitext(os.path.basename(rel))[0]
        if _looks_like_contract_name(stem):
            score += 1
    return score


def resolve_pdf_folder(config: Optional[AppConfig] = None,
                       extra_paths: Sequence[str] = ()) -> str:
    """پوشهٔ PDF را از تنظیمات، کنار اکسل، کنار برنامه یا مسیر شبکه پیدا می‌کند."""
    candidates: List[str] = []

    def add(value: Any) -> None:
        text = _expand_folder(value)
        if text and text not in candidates:
            candidates.append(text)

    if config is not None:
        add(config.get("pdf_folder"))
    for value in extra_paths:
        add(value)
    excel_dirs: List[str] = []
    if config is not None:
        for key in ("local_excel_path", "last_excel_path"):
            path = str(config.get(key) or "").strip()
            if path:
                folder = os.path.dirname(_expand_folder(path))
                if folder and folder not in excel_dirs:
                    excel_dirs.append(folder)
    for base in excel_dirs:
        for name in PDF_FOLDER_NAMES:
            add(os.path.join(base, name))
        add(base)
    for base in (str(app_dir()), str(data_dir())):
        for name in PDF_FOLDER_NAMES:
            add(os.path.join(base, name))
        add(base)
    for base in (str(Path.home() / "Documents"), str(Path.home() / "Desktop")):
        for name in PDF_FOLDER_NAMES:
            add(os.path.join(base, name))
    add(ORIGINAL_PDF_FOLDER)

    explicit = _expand_folder(config.get("pdf_folder")) if config is not None else ""
    if explicit and os.path.isdir(explicit):
        return explicit

    pool = candidates[1:] if explicit else candidates
    best = ""
    best_score = 0
    for folder in pool:
        if not os.path.isdir(folder) or not _pdf_folder_score(folder):
            continue
        score = _pdf_folder_score(folder)
        if score > best_score:
            best, best_score = folder, score
    if best:
        return best
    # هیچ نام هم‌خوانی نبود؛ پوشه‌ای که بیشترین PDF را دارد
    best_count = 0
    for folder in pool:
        if not os.path.isdir(folder):
            continue
        count = len(scan_pdf_files(folder, depth=1, limit=300))
        if count > best_count:
            best, best_count = folder, count
    return best


def default_pdf_folder(config: Optional[AppConfig] = None) -> str:
    return resolve_pdf_folder(config)


class DatabaseError(ExcelDataError):
    """خطای اتصال یا خواندن از دیتابیس مرکزی."""


# آخرین اتصال موفق (برای اینکه اگر ورودِ جایگزین جواب داد، همان ذخیره شود)
LAST_CONNECTION: Dict[str, Any] = {}


def login_candidates(settings: Dict[str, Any]) -> List[Tuple[str, str]]:
    """ترتیب تلاش برای ورود: ابتدا تنظیمات کاربر، سپس گزینه‌های بدون رمز."""
    user = str(settings.get("user") or "master_app").strip() or "master_app"
    password = str(settings.get("password") or "")
    candidates: List[Tuple[str, str]] = [(user, password)]
    for extra in (("root", ""), ("root", password), (user, "")):
        if extra not in candidates:
            candidates.append(extra)
    return candidates


CONNECTOR_ENGINE = "pure"


def connector_engine_label() -> str:
    """نام موتور اتصالی که استفاده می‌شود (برای لاگ و صفحهٔ راهنما)."""
    return "پایتون خالص (use_pure)"


def _mysql_connect(**kwargs):
    """اتصال به MySQL با موتور خالصِ پایتون (use_pure=True).

    موتور C کتابخانهٔ mysql-connector (فایل connection_cext) روی ویندوز + Python 3.13
    با خطای «access violation» کل برنامه را می‌بندد و با try/except هم گرفتنِ آن ممکن
    نیست؛ بنابراین همیشه نسخهٔ پایتونی استفاده می‌شود.
    """
    kwargs = dict(kwargs)
    kwargs["use_pure"] = True
    try:
        connection = mysql_connector.connect(**kwargs)
    except TypeError as exc:  # نسخه‌های قدیمی‌تر که این گزینه را نمی‌شناسند
        if "use_pure" not in str(exc):
            raise
        kwargs.pop("use_pure", None)
        connection = mysql_connector.connect(**kwargs)
    return connection


def _is_network_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return any(token in text for token in (
        "can't connect", "cannot connect", "2003", "10061", "timed out",
        "unknown host", "getaddrinfo", "connection refused",
    ))


def connect_database(settings: Dict[str, Any], timeout: int = 8,
                     allow_fallback: bool = True):
    """اتصال به دیتابیس؛ اگر رمز/کاربر جواب نداد، ورودهای بدون رمز هم امتحان می‌شود."""
    if mysql_connector is None:
        raise DatabaseError(
            "کتابخانهٔ mysql-connector-python نصب نیست.\n"
            "دستور نصب: pip install mysql-connector-python"
        )
    host = settings.get("host") or "127.0.0.1"
    port = int(settings.get("port") or 3306)
    database = settings.get("database") or "master_data"
    candidates = login_candidates(settings) if allow_fallback else [
        (str(settings.get("user") or "master_app"), str(settings.get("password") or ""))
    ]
    errors: List[str] = []
    for index, (user, password) in enumerate(candidates):
        log_event(
            f"اتصال به MySQL: {host}:{port}/{database} کاربر={user} "
            f"({'بدون رمز' if not password else 'با رمز'}) "
            f"| موتور {connector_engine_label()}"
        )
        try:
            connection = _mysql_connect(
                host=host,
                port=port,
                user=user,
                password=password,
                database=database,
                charset="utf8mb4",
                connection_timeout=timeout,
                autocommit=False,
            )
        except Exception as exc:
            text = f"{user}: {type(exc).__name__}: {exc}"
            errors.append(text)
            log_event("خطای اتصال به MySQL — " + text)
            if _is_network_error(exc):
                break
            continue
        log_event(
            "اتصال به MySQL برقرار شد"
            + ("" if index == 0 else f" (با ورود جایگزین {user})")
        )
        LAST_CONNECTION.clear()
        LAST_CONNECTION.update({
            "host": host, "port": port, "database": database,
            "user": user, "password": password,
        })
        return connection
    detail = "\n".join(errors[:4]) if errors else "خطای نامشخص"
    auth_only = bool(errors) and all(
        ("1045" in item) or ("access denied" in item.lower()) for item in errors
    )
    if auth_only:
        raise DatabaseError(
            f"نام کاربری یا رمز MySQL پذیرفته نشد ({host}:{port} — دیتابیس {database}).\n"
            f"{detail}\n\n"
            "راه‌حل‌ها:\n"
            "• رمزِ درستِ کاربر «master_app» را در تنظیمات وارد کنید (فیلد رمز را خالی نگذارید).\n"
            "• یا در MySQL یک‌بار رمز را تنظیم کنید:\n"
            "   ALTER USER 'master_app'@'localhost' IDENTIFIED BY 'رمز-شما';\n"
            "   FLUSH PRIVILEGES;\n"
            "• اگر رمز root را دارید، همان را با کاربر root وارد کنید.\n"
            "• اگر دیتابیس روی سرور دیگری است، آدرس آن را در فیلد «میزبان» بگذارید.\n"
            "• برای کار بدون دیتابیس: «اجرای محلی (بدون دیتابیس)» یا به‌روزرسانی محلی فایل ۰۱."
        )
    raise DatabaseError(
        f"اتصال به دیتابیس ممکن نشد ({host}:{port}).\n{detail}\n\n"
        "اگر MySQL شما کاربر بدون رمز دارد (مثل root بدون رمز)، همان را وارد کنید؛ "
        "یا از «اجرای محلی (بدون دیتابیس)» استفاده کنید."
    )


def connect_server(settings: Dict[str, Any], timeout: int = 8):
    """اتصال به سرور MySQL بدون انتخاب دیتابیس (برای ساخت اولیه یا بارگذاری)."""
    if mysql_connector is None:
        raise DatabaseError(
            "کتابخانهٔ mysql-connector-python نصب نیست.\n"
            "دستور نصب: pip install mysql-connector-python"
        )
    host = settings.get("host") or "127.0.0.1"
    port = int(settings.get("port") or 3306)
    user = str(settings.get("user") or "root")
    password = str(settings.get("password") or "")
    candidates: List[Tuple[str, str]] = [(user, password)]
    for extra in (("root", ""), ("root", password), (user, "")):
        if extra not in candidates:
            candidates.append(extra)
    errors: List[str] = []
    for index, (try_user, try_password) in enumerate(candidates):
        log_event(
            f"اتصال به سرور MySQL: {host}:{port} کاربر={try_user} "
            f"({'بدون رمز' if not try_password else 'با رمز'})"
        )
        try:
            connection = _mysql_connect(
                host=host,
                port=port,
                user=try_user,
                password=try_password,
                charset="utf8mb4",
                connection_timeout=timeout,
                autocommit=False,
            )
        except Exception as exc:
            errors.append(f"{try_user}: {type(exc).__name__}: {exc}")
            log_event(f"خطای اتصال به سرور MySQL — {try_user}: {exc}")
            if _is_network_error(exc):
                break
            continue
        log_event(
            "اتصال به سرور MySQL برقرار شد"
            + ("" if index == 0 else f" (با ورود جایگزین {try_user})")
        )
        LAST_CONNECTION.clear()
        LAST_CONNECTION.update({
            "host": host, "port": port, "database": "",
            "user": try_user, "password": try_password,
        })
        return connection
    detail = "\n".join(errors[:4]) if errors else "خطای نامشخص"
    raise DatabaseError(
        f"اتصال به سرور دیتابیس ممکن نشد ({host}:{port}).\n{detail}\n\n"
        "اگر کاربر MySQL شما رمز ندارد، فیلد رمز را خالی بگذارید "
        "(مثلاً کاربر root با رمز خالی)."
    )


WORKBOOK_SCHEMA_SQL = (
    """CREATE TABLE IF NOT EXISTS excel_workbooks (
        workbook_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        original_filename VARCHAR(255) NOT NULL,
        original_path TEXT NULL,
        file_size BIGINT UNSIGNED NOT NULL,
        sha256 CHAR(64) NOT NULL,
        imported_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        sheet_count INT NOT NULL DEFAULT 0,
        stored_cell_count BIGINT UNSIGNED NOT NULL DEFAULT 0,
        import_status VARCHAR(20) NOT NULL DEFAULT 'importing',
        workbook_metadata JSON NULL,
        KEY ix_workbook_filename (original_filename),
        KEY ix_workbook_sha256 (sha256),
        KEY ix_workbook_imported_at (imported_at)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci""",
    """CREATE TABLE IF NOT EXISTS excel_original_files (
        workbook_id BIGINT UNSIGNED NOT NULL PRIMARY KEY,
        filename VARCHAR(255) NOT NULL,
        content_type VARCHAR(100) NOT NULL DEFAULT 'application/octet-stream',
        sha256 CHAR(64) NOT NULL,
        binary_data LONGBLOB NOT NULL,
        CONSTRAINT fk_original_workbook FOREIGN KEY (workbook_id)
            REFERENCES excel_workbooks(workbook_id) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci""",
)


def _safe_identifier(name: str, label: str) -> str:
    """نام دیتابیس/جدول را برای درج در دستور SQL امن می‌کند."""
    clean = str(name or "").strip()
    if not clean or not re.fullmatch(r"[A-Za-z0-9_$]+", clean):
        raise DatabaseError(
            f"{label} «{clean}» معتبر نیست؛ فقط حروف انگلیسی، رقم و زیرخط مجاز است."
        )
    return clean


# ---------------------------------------------------------------------------
# دیتابیس نرمال‌شده (نسخهٔ ۳٫۵) — اعداد داشبورد از همین جدول‌ها ساخته می‌شوند
# ---------------------------------------------------------------------------
# هر ردیف شیت colorcode به یک سطر mehr_contracts تبدیل می‌شود: ستون‌های شناسایی
# جدا و قابل جست‌وجو، و نسخهٔ کامل مقدارهای ردیف در row_values (JSON) تا بازسازی
# داده بدون افت دقت انجام شود. سه ردیف اول شیت (عنوان، درصد کل، سرستون کسورات)
# در mehr_workbooks.meta_rows نگه داشته می‌شوند.
MEHR_SCHEMA_SQL = (
    """CREATE TABLE IF NOT EXISTS mehr_workbooks (
        workbook_id BIGINT UNSIGNED NOT NULL PRIMARY KEY,
        original_filename VARCHAR(255) NOT NULL DEFAULT '',
        sheet_name VARCHAR(64) NOT NULL DEFAULT 'colorcode',
        sheet_count INT NOT NULL DEFAULT 0,
        data_rows INT NOT NULL DEFAULT 0,
        progress_cell VARCHAR(64) NULL,
        meta_rows LONGTEXT NULL,
        source_kind VARCHAR(20) NOT NULL DEFAULT 'excel',
        source_sha CHAR(64) NOT NULL DEFAULT '',
        imported_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci""",
    """CREATE TABLE IF NOT EXISTS mehr_contracts (
        row_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        workbook_id BIGINT UNSIGNED NOT NULL,
        excel_row INT NOT NULL,
        contract_no VARCHAR(64) NOT NULL DEFAULT '',
        company VARCHAR(255) NOT NULL DEFAULT '',
        subject VARCHAR(512) NOT NULL DEFAULT '',
        employer VARCHAR(255) NOT NULL DEFAULT '',
        contract_type VARCHAR(255) NOT NULL DEFAULT '',
        deductions VARCHAR(1024) NOT NULL DEFAULT '',
        deduction_count INT NOT NULL DEFAULT 0,
        flag_temporary TINYINT NOT NULL DEFAULT 0,
        flag_prefinal TINYINT NOT NULL DEFAULT 0,
        flag_balance TINYINT NOT NULL DEFAULT 0,
        flag_statement TINYINT NOT NULL DEFAULT 0,
        flag_mofasa TINYINT NOT NULL DEFAULT 0,
        flag_affidavit TINYINT NOT NULL DEFAULT 0,
        flag_delivery TINYINT NOT NULL DEFAULT 0,
        flag_financial TINYINT NOT NULL DEFAULT 0,
        progress_status VARCHAR(64) NOT NULL DEFAULT '',
        completion_status VARCHAR(64) NOT NULL DEFAULT '',
        annex_status VARCHAR(64) NOT NULL DEFAULT '',
        progress_table VARCHAR(64) NOT NULL DEFAULT '',
        as_status VARCHAR(64) NOT NULL DEFAULT '',
        row_values LONGTEXT NOT NULL,
        UNIQUE KEY uq_workbook_row (workbook_id, excel_row),
        KEY ix_contract_no (contract_no),
        KEY ix_company (company)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci""",
)

_FLAG_COLUMNS_SQL = {
    "temporary": "flag_temporary",
    "prefinal": "flag_prefinal",
    "balance": "flag_balance",
    "statement": "flag_statement",
    "mofasa": "flag_mofasa",
    "affidavit": "flag_affidavit",
    "delivery": "flag_delivery",
    "financial": "flag_financial",
}


def _json_default(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


def _rows_to_json(rows: Sequence[Sequence[Any]]) -> str:
    return json.dumps([list(row) for row in rows], ensure_ascii=False,
                      default=_json_default)


def _rows_from_json(text: Any) -> List[List[Any]]:
    if not text:
        return []
    if isinstance(text, (bytes, bytearray)):
        text = text.decode("utf-8", "replace")
    try:
        data = json.loads(str(text))
    except Exception:
        return []
    return [list(row) for row in data if isinstance(row, list)]


def _numeric_columns(width: int) -> List[int]:
    """ستون‌هایی که در JSON باید عدد بمانند (درصد پیشرفت)."""
    return [37]


def _restore_numbers(row: List[Any]) -> List[Any]:
    """اعدادی که JSON به رشته تبدیل کرده را به عدد برمی‌گرداند."""
    for index in _numeric_columns(len(row)):
        if index >= len(row):
            break
        value = row[index]
        if isinstance(value, str):
            number = num(value)
            if number is not None:
                row[index] = number
    return row


def ensure_mehr_schema(cursor) -> None:
    """جدول‌های نرمال‌شده را (اگر نباشند) می‌سازد."""
    for statement in MEHR_SCHEMA_SQL:
        cursor.execute(statement)


def store_rows_in_database(cursor, workbook_id: int, filename: str,
                           sheet_name: str, sheet_count: int,
                           raw_rows: Sequence[Sequence[Any]],
                           source_kind: str = "excel",
                           source_sha: str = "") -> Dict[str, int]:
    """همهٔ ردیف‌های شیت را در جدول‌های نرمال‌شده می‌نویسد.

    سه ردیف اول (عنوان/درصد کل/سرستون کسورات) در meta_rows و بقیه به‌صورت
    سطرهای قرارداد ذخیره می‌شوند؛ بنابراین بازخوانی، همان ردیف‌های اصلی را
    برمی‌گرداند و اعداد داشبورد دقیقاً همان اعداد فایل می‌شود.
    """
    rows = [list(row) for row in raw_rows]
    meta_rows = rows[:3]
    data_rows = rows[3:]
    width = max((len(row) for row in rows), default=0)
    padded = [row + [None] * (width - len(row)) for row in data_rows]

    progress_cell = ""
    if len(rows) > 1 and len(rows[1]) > 36 and rows[1][36] not in (None, ""):
        progress_cell = str(rows[1][36])

    payload: List[Tuple[Any, ...]] = []
    for offset, row in enumerate(padded):
        excel_row = 4 + offset
        if all(is_blank(value) for value in row[:REQUIRED_COLUMNS]):
            continue  # ردیف کاملاً خالی، مثل پارس داشبورد نادیده گرفته می‌شود
        flags = {
            key: is_status_mark(row[column] if column < len(row) else None)
            for key, _label, column in FLAG_DEFS
        }
        marked = [
            index for index in range(23)
            if 11 + index < len(row) and is_status_mark(row[11 + index])
        ]
        deduction_headers = meta_rows[2] if len(meta_rows) > 2 else []
        deductions = " | ".join(
            txt(deduction_headers[index]) for index in marked
            if index < len(deduction_headers) and not is_blank(deduction_headers[index])
        )
        payload.append((
            int(workbook_id), excel_row,
            txt(row[4] if len(row) > 4 else "")[:64],
            txt(row[3] if len(row) > 3 else "")[:255],
            txt(row[6] if len(row) > 6 else "")[:512],
            txt(row[2] if len(row) > 2 else "")[:255],
            txt(row[5] if len(row) > 5 else "")[:255],
            deductions[:1024], len(marked),
            *(1 if flags.get(key) else 0 for key, _label, _column in FLAG_DEFS),
            (normalize_progress_value(row[53] if len(row) > 53 else None) or "")[:64],
            (normalize_completion_value(row[54] if len(row) > 54 else None) or "")[:64],
            (normalize_annex_value(row[55] if len(row) > 55 else None) or "")[:64],
            _progress_like(row[10] if len(row) > 10 else None)[:64],
            txt(row[44] if len(row) > 44 else "")[:64],
            _rows_to_json([row]),
        ))
    cursor.execute(
        "DELETE FROM mehr_contracts WHERE workbook_id=%s", (int(workbook_id),)
    )
    cursor.execute(
        "DELETE FROM mehr_workbooks WHERE workbook_id=%s", (int(workbook_id),)
    )
    cursor.execute(
        "INSERT INTO mehr_workbooks (workbook_id, original_filename, sheet_name, "
        "sheet_count, data_rows, progress_cell, meta_rows, source_kind, source_sha) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (int(workbook_id), str(filename)[:255], str(sheet_name)[:64],
         int(sheet_count), len(payload), progress_cell, _rows_to_json(meta_rows),
         str(source_kind)[:20], str(source_sha)[:64]),
    )
    if payload:
        cursor.executemany(
            "INSERT INTO mehr_contracts (workbook_id, excel_row, contract_no, company, "
            "subject, employer, contract_type, deductions, deduction_count, "
            "flag_temporary, flag_prefinal, flag_balance, flag_statement, flag_mofasa, "
            "flag_affidavit, flag_delivery, flag_financial, progress_status, "
            "completion_status, annex_status, progress_table, as_status, row_values) "
            "VALUES (" + ",".join(["%s"] * 23) + ")",
            payload,
        )
    return {"data_rows": len(payload), "width": width}


def load_rows_from_database(settings: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """ردیف‌های شیت را از جدول‌های نرمال‌شده برمی‌گرداند (یا None اگر خالی باشد)."""
    workbook_id = int(settings.get("workbook_id") or 1)
    connection = connect_database(settings)
    try:
        cursor = connection.cursor()
        try:
            cursor.execute(
                "SELECT original_filename, sheet_name, sheet_count, data_rows, meta_rows, "
                "source_kind FROM mehr_workbooks WHERE workbook_id=%s", (workbook_id,)
            )
            meta = cursor.fetchone()
            if not meta:
                return None
            cursor.execute(
                "SELECT excel_row, row_values FROM mehr_contracts "
                "WHERE workbook_id=%s ORDER BY excel_row", (workbook_id,)
            )
            rows = cursor.fetchall() or []
        finally:
            cursor.close()
    finally:
        try:
            connection.close()
        except Exception:
            pass
    if not rows:
        return None
    by_row: Dict[int, List[Any]] = {}
    for item in rows:
        decoded = _rows_from_json(item[1])
        if not decoded:
            continue
        try:
            excel_row = int(item[0] or 0)
        except (TypeError, ValueError):
            continue
        by_row[excel_row] = _restore_numbers(decoded[0])
    if not by_row:
        return None
    width = max(len(row) for row in by_row.values())
    last_row = max(by_row)
    body: List[List[Any]] = []
    for excel_row in range(4, last_row + 1):
        row = by_row.get(excel_row)
        body.append(list(row) if row else [None] * width)
    meta_rows = _rows_from_json(meta[4])
    raw_rows = [list(row) for row in meta_rows[:3]] + body
    return {
        "filename": str(meta[0] or ""),
        "sheet_name": str(meta[1] or SHEET_NAME_DEFAULT),
        "sheet_count": int(meta[2] or 0),
        "data_rows": int(meta[3] or 0),
        "raw_rows": raw_rows,
        "source_kind": str((meta[5] if len(meta) > 5 else "") or "excel"),
    }


def verify_database_numbers(settings: Dict[str, Any]) -> Dict[str, Any]:
    """اعداد محاسبه‌شده با SQL را با اعداد داشبورد مقایسه می‌کند."""
    workbook_id = int(settings.get("workbook_id") or 1)
    connection = connect_database(settings)
    try:
        cursor = connection.cursor()
        try:
            cursor.execute(
                "SELECT COUNT(*), SUM(deduction_count > 0), "
                + ", ".join(f"SUM({column})" for column in _FLAG_COLUMNS_SQL.values())
                + " FROM mehr_contracts WHERE workbook_id=%s",
                (workbook_id,),
            )
            row = cursor.fetchone() or (0,) * 9
        finally:
            cursor.close()
    finally:
        try:
            connection.close()
        except Exception:
            pass
    counts = {
        "total": int(row[0] or 0),
        "deductions": int(row[1] or 0),
    }
    for index, key in enumerate(_FLAG_COLUMNS_SQL, start=2):
        counts[key] = int(row[index] or 0)
    return counts


def import_workbook_to_database(
    settings: Dict[str, Any],
    file_path: str,
    sheet_name: Optional[str] = None,
    workbook_id: Optional[int] = None,
) -> Dict[str, Any]:
    """بارگذاری فایل اکسل در دیتابیس از ابتدا (ساخت جدول‌ها در صورت نیاز).

    این تابع بدون نیاز به «مدیر دیتابیس مرکزی» کار می‌کند: دیتابیس و جدول‌های
    لازم را می‌سازد و فایل را به‌عنوان کتاب کار شمارهٔ ۱ ذخیره می‌کند.
    """
    path = Path(str(file_path)).expanduser()
    if not path.exists() or not path.is_file():
        raise ExcelDataError(f"فایل اکسل پیدا نشد:\n{path}")
    if openpyxl is None:
        raise ExcelDataError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    blob = path.read_bytes()
    if not blob:
        raise ExcelDataError("فایل اکسل خالی است.")
    try:
        probe = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
        sheets = [str(name) for name in probe.sheetnames]
        probe.close()
    except Exception as exc:
        raise ExcelDataError(f"فایل انتخاب‌شده یک کتاب کار معتبر نیست:\n{exc}") from exc
    target_sheet = str(sheet_name or SHEET_NAME_DEFAULT).strip() or SHEET_NAME_DEFAULT
    if target_sheet not in sheets:
        lowered = {name.strip().lower(): name for name in sheets}
        target_sheet = lowered.get(target_sheet.lower(), "")
    if not target_sheet:
        raise ExcelDataError(
            f"شیت «{sheet_name or SHEET_NAME_DEFAULT}» در فایل نیست.\n"
            f"شیت‌های موجود: {'، '.join(sheets)}"
        )
    health = workbook_health(blob)
    repaired: Dict[str, Any] = {}
    if health.get("damaged"):
        # ذخیرهٔ فایلی که فرمول‌هایش مقدار ندارند داشبورد را صفر می‌کند؛ اول
        # ستون‌های شناسایی/درصد را از شیت خروجی سامانه همان فایل بازسازی می‌کنیم.
        try:
            book = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
            try:
                rows = [list(row) for row in book[target_sheet].iter_rows(values_only=True)]
                repaired = repair_rows_from_companion(book, target_sheet, rows)
            finally:
                book.close()
        except Exception as exc:
            repaired = {"filled": 0, "error": str(exc)}
        if not repaired.get("filled"):
            raise ExcelDataError(
                "این فایل سالم نیست: فرمول‌های شیت «"
                + str(health.get("sheet") or target_sheet)
                + "» مقدار محاسبه‌شده ندارند ("
                + fa(round(float(health.get("ratio") or 0) * 100))
                + "٪ خالی است) و شیت خروجی سامانه هم برای بازسازی پیدا نشد.\n\n"
                "فایل اصلی را در اکسل باز کنید، یک بار ذخیره کنید (F9 و سپس Ctrl+S) "
                "و بعد همین فایل را بارگذاری کنید."
            )
        log_event(
            "فایل ناقص پیش از ذخیره بازسازی شد: "
            f"{repaired.get('filled')} ردیف از شیت «{repaired.get('sheet')}»"
        )

    database = _safe_identifier(settings.get("database") or "master_data", "نام دیتابیس")
    wid = int(workbook_id or settings.get("workbook_id") or 1)
    size = len(blob)
    digest = hashlib.sha256(blob).hexdigest()
    filename = path.name
    log_event(f"بارگذاری اکسل در دیتابیس: {filename} ({size} بایت) → {database} #{wid}")

    stored_rows = 0
    source_kind = "generated" if is_generated_workbook(str(path)) else "excel"
    connection = connect_server(settings)
    try:
        cursor = connection.cursor()
        try:
            cursor.execute(
                f"CREATE DATABASE IF NOT EXISTS `{database}` "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
            cursor.execute(f"USE `{database}`")
            for statement in WORKBOOK_SCHEMA_SQL:
                cursor.execute(statement)
            ensure_mehr_schema(cursor)
            cursor.execute(
                "SELECT workbook_id FROM excel_workbooks WHERE workbook_id=%s", (wid,)
            )
            exists = cursor.fetchone() is not None
            if exists:
                cursor.execute(
                    "UPDATE excel_workbooks SET original_filename=%s, original_path=%s, "
                    "file_size=%s, sha256=%s, imported_at=NOW(), sheet_count=%s, "
                    "import_status='completed' WHERE workbook_id=%s",
                    (filename, str(path), size, digest, len(sheets), wid),
                )
            else:
                cursor.execute(
                    "INSERT INTO excel_workbooks (workbook_id, original_filename, "
                    "original_path, file_size, sha256, imported_at, sheet_count, "
                    "stored_cell_count, import_status) "
                    "VALUES (%s,%s,%s,%s,%s,NOW(),%s,0,'completed')",
                    (wid, filename, str(path), size, digest, len(sheets)),
                )
            cursor.execute(
                "SELECT workbook_id FROM excel_original_files WHERE workbook_id=%s", (wid,)
            )
            has_file = cursor.fetchone() is not None
            if has_file:
                cursor.execute(
                    "UPDATE excel_original_files SET filename=%s, content_type=%s, "
                    "sha256=%s, binary_data=%s WHERE workbook_id=%s",
                    (filename,
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     digest, blob, wid),
                )
            else:
                cursor.execute(
                    "INSERT INTO excel_original_files (workbook_id, filename, "
                    "content_type, sha256, binary_data) VALUES (%s,%s,%s,%s,%s)",
                    (wid, filename,
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     digest, blob),
                )
            # ردیف‌های شیت در جدول‌های نرمال‌شده هم ذخیره می‌شوند؛ داشبورد
            # از همین جدول‌ها ساخته می‌شود تا اعداد دقیقاً همان اعداد فایل باشند.
            try:
                book = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
                try:
                    raw_rows = [list(row) for row in book[target_sheet].iter_rows(values_only=True)]
                    if repaired.get("filled"):
                        repair_rows_from_companion(book, target_sheet, raw_rows)
                finally:
                    book.close()
                stored = store_rows_in_database(
                    cursor, wid, filename, target_sheet, len(sheets), raw_rows,
                    source_kind=source_kind, source_sha=digest,
                )
                stored_rows = int(stored.get("data_rows") or 0)
                log_event(f"ردیف‌ها در جدول‌های نرمال‌شده ذخیره شد: {fa(stored_rows)} سطر")
            except Exception as exc:
                log_event(f"ذخیرهٔ ردیف‌های نرمال‌شده ناموفق بود: {type(exc).__name__}: {exc}")
            connection.commit()
        finally:
            cursor.close()
    except ExcelDataError:
        raise
    except Exception as exc:
        try:
            connection.rollback()
        except Exception:
            pass
        log_event(f"خطای بارگذاری اکسل: {type(exc).__name__}: {exc}")
        raise DatabaseError(
            "بارگذاری فایل اکسل در دیتابیس ناموفق بود.\n"
            f"{exc}\n\n"
            "اگر کاربر MySQL دسترسی ساخت دیتابیس/جدول ندارد، دیتابیس و جدول‌ها را "
            "یک‌بار با «مدیریت دیتابیس مرکزی» بسازید و دوباره تلاش کنید."
        ) from exc
    finally:
        try:
            connection.close()
        except Exception:
            pass

    log_event(f"اکسل در دیتابیس ذخیره شد: {database} / workbook_id={wid}")
    return {
        "workbook_id": wid,
        "filename": filename,
        "size": size,
        "sheets": sheets,
        "sheet_name": target_sheet,
        "database": database,
        "sha256": digest,
        "data_rows": stored_rows,
    }


# ==========================================================================
# به‌روزرسانی خودکار از «فایل سامانه» (فایل ۰۱) — بدون کپی‌پیست دستی
# ==========================================================================
SYSTEM_UPDATE_FIRST_COLUMN = 12   # ستون L
SYSTEM_UPDATE_LAST_COLUMN = 36    # ستون AJ
_SYSTEM_UPDATE_OK_WORDS = {"okk", "ok"}

# علامت‌هایی که در ستون‌های وضعیت/کسورات «ثبت‌شده» حساب می‌شوند.
# مرجع، همان رفتار نسخهٔ اولیه است: فقط حرف M (و «م»/تیک).
# عدد ۱ یا x عمداً شمرده نمی‌شوند تا شاخص‌ها اشتباه نشوند؛ در عوض، هنگام
# نوشتن/ساخت کتاب کار، مقدار ۱ به «M» و ۰ به خالی تبدیل می‌شود تا کارت‌ها
# صفر نشوند.
_STATUS_MARK_TOKENS = {"m", "م", "✓", "✔"}


def is_status_mark(value: Any) -> bool:
    """آیا سلول علامت «ثبت‌شده» است؟ (فقط M / م / ✓)"""
    if value is None or isinstance(value, (bool, int, float)):
        return False
    text = normalize_per(txt(value)).strip()
    if not text:
        return False
    if text in _STATUS_MARK_TOKENS:
        return True
    compact = re.sub(r"\s+", " ", text)
    return compact in _STATUS_MARK_TOKENS


FLAG_COLUMNS_ONE_BASED = {column + 1 for _key, _label, column in FLAG_DEFS}

# ناحیهٔ ستون‌های «سند» در شیت داشبورد (colorcode): ستون L تا AH.
# این ستون‌ها در کتاب کار با OK/M پر می‌شوند و هرگز نباید عدد ۰/۱ بگیرند.
DASHBOARD_DOC_FIRST_COLUMN = 12
DASHBOARD_DOC_LAST_COLUMN = 34


def status_mark_value(value: Any) -> Any:
    """مقدار ستون وضعیت را به «علامت» تبدیل می‌کند: ۱ → M، ۰ → خالی.

    خروجی سامانه در این ستون‌ها عدد ۰/۱ دارد؛ اگر همان عدد در کتاب کار نوشته
    شود، داشبورد آن را «علامت» نمی‌شناسد و کارت‌ها صفر می‌شوند.
    """
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return "M" if value else None
    if isinstance(value, (int, float)):
        return "M" if float(value) > 0 else None
    text = normalize_per(txt(value)).strip()
    lowered = text.lower()
    if lowered in ("0", "0.0", "false", "-", "na", "none", "n/a"):
        return None
    if lowered in _STATUS_MARK_TOKENS:
        return "M"
    return value


def is_dashboard_sheet_name(name: str) -> bool:
    """آیا این شیت همان شیت داشبورد است (colorcode)؟"""
    text = normalize_per(txt(name)).strip().lower()
    return "colorcode" in text or "color code" in text or "کالرکد" in text


def repair_dashboard_marks(workbook) -> int:
    """عددهای ۰/۱ جاافتاده در ستون‌های سند شیت داشبورد را به علامت تبدیل می‌کند.

    نسخه‌های قبل هنگام به‌روزرسانی، شمارش‌های ۰/۱ فایل سامانه را عیناً در
    ستون‌های وضعیت می‌نوشتند و کارت‌ها صفر می‌شدند. این تابع همان آسیب را در
    هر شیت داشبورد (colorcode) ترمیم می‌کند: ۱ → M و ۰ → خالی.
    """
    fixed = 0
    for name in list(workbook.sheetnames):
        if not is_dashboard_sheet_name(str(name)):
            continue
        sheet = workbook[name]
        last_row = min(int(sheet.max_row or 0), 20000)
        for row_number in range(4, last_row + 1):
            for column in range(DASHBOARD_DOC_FIRST_COLUMN,
                                DASHBOARD_DOC_LAST_COLUMN + 1):
                try:
                    cell = sheet.cell(row=row_number, column=column)
                    current = getattr(cell, "value", None)
                    if isinstance(current, str) and current.startswith("="):
                        continue
                    if isinstance(current, bool) or not isinstance(current, (int, float)):
                        continue
                    if float(current) not in (0.0, 1.0):
                        continue
                    cell.value = status_mark_value(current)
                    fixed += 1
                except AttributeError:
                    continue
    return fixed


def has_amount(value: Any) -> bool:
    """آیا سلول یک مبلغ/عدد مثبت است (برای تشخیص خودکار ستون‌های کسورات)؟"""
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return value != 0
    text = normalize_per(txt(value)).strip()
    if not text or text in _STATUS_MARK_TOKENS:
        return False
    number = num(text)
    return bool(number) and abs(number) > 0


def _norm_words(value: Any) -> str:
    """متن را برای مقایسهٔ کلمه‌ای ساده می‌کند (بدون فاصله، خط تیره و نیم‌فاصله)."""
    text = normalize_per(txt(value))
    return re.sub(r"[\s\-_/\\.,،:;()\[\]{}<>«»\"']+", "", text)


def _looks_like_sub_project_header(value: Any) -> bool:
    text = _norm_words(value)
    return "subproject" in text or "زیرپروژه" in text


def _looks_like_contract_header(value: Any) -> bool:
    text = _norm_words(value)
    if not text:
        return False
    if "قرارداد" in text and ("شماره" in text or "کد" in text):
        return True
    if "contract" in text and ("no" in text or "num" in text or "code" in text):
        return True
    return text in {"شمارهقرارداد", "کدقرارداد", "contractno", "contractnumber"}


def _looks_like_contract_number(value: Any) -> bool:
    text = txt(value)
    if len(text) < 4 or len(text) > 40:
        return False
    digits = sum(1 for char in text if char.isdigit())
    letters = sum(1 for char in text if char.isalpha())
    return digits >= 3 and (letters >= 2 or "/" in text or "-" in text)


def _matches_sub_project(value: Any) -> bool:
    text = _norm_words(value)
    return "اجرا" in text and "مهر" in text


def _open_rows(path: str) -> Tuple[List[str], str, List[List[Any]]]:
    """فهرست شیت‌ها، شیت انتخاب‌شده و همهٔ ردیف‌های آن را برمی‌گرداند."""
    if openpyxl is None:
        raise ExcelDataError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    try:
        workbook = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    except Exception as exc:
        raise ExcelDataError(
            f"فایل «{Path(str(path)).name}» به‌عنوان اکسل (.xlsx) باز نشد:\n{exc}\n"
            "اگر فایل از سامانه به‌صورت .xls یا .csv گرفته شده، اول در اکسل با فرمت "
            ".xlsx ذخیره کنید."
        ) from exc
    try:
        names = [str(name) for name in workbook.sheetnames]
        chosen = ""
        for name in names:
            lowered = normalize_per(name)
            if "colorcode" in lowered or "کندو" in lowered:
                chosen = name
                break
        if not chosen:
            # بین چند شیت (مثل «Sheet 1» و «Sheet1 (2)») شیتی انتخاب می‌شود که
            # ردیف‌های «اجرا - مهر» دارد؛ وگرنه بزرگ‌ترین شیت.
            best_hits, best_size, fallback = -1, -1, ""
            for name in names:
                sheet = workbook[name]
                size = int(sheet.max_row or 0) * int(sheet.max_column or 0)
                if size > best_size:
                    best_size, fallback = size, name
                hits = 0
                for index, row in enumerate(sheet.iter_rows(values_only=True)):
                    if index > 300:
                        break
                    if any(_matches_sub_project(value) for value in row):
                        hits += 1
                if hits and hits > best_hits:
                    best_hits, chosen = hits, name
            if not chosen:
                chosen = fallback
        if not chosen and names:
            chosen = names[0]
        if not chosen:
            raise ExcelDataError("فایل انتخاب‌شده هیچ شیتی ندارد.")
        rows = [list(row) for row in workbook[chosen].iter_rows(values_only=True)]
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    return names, chosen, rows


def _detect_header_row(rows: Sequence[Sequence[Any]], limit: int = 15) -> Tuple[int, int, int]:
    """ردیف سرستون‌ها و ستون‌های «زیرپروژه» و «شمارهٔ قرارداد» را پیدا می‌کند."""
    best_row, best_score = -1, 0
    for index, row in enumerate(rows[:limit]):
        score = 0
        for value in row:
            if _looks_like_sub_project_header(value):
                score += 3
            elif _looks_like_contract_header(value):
                score += 2
        if score > best_score:
            best_row, best_score = index, score
    header_row = best_row if best_row >= 0 else 2
    sub_col, contract_col = -1, -1
    for column, value in enumerate(rows[header_row] if header_row < len(rows) else []):
        if sub_col < 0 and _looks_like_sub_project_header(value):
            sub_col = column
        if contract_col < 0 and _looks_like_contract_header(value):
            contract_col = column
    return header_row, sub_col, contract_col


def _guess_sub_project_column(rows: Sequence[Sequence[Any]], start: int) -> int:
    best_col, best_hits = -1, 0
    totals: Dict[int, int] = {}
    for index in range(start, min(len(rows), start + 500)):
        for column, value in enumerate(rows[index]):
            if _matches_sub_project(value):
                totals[column] = totals.get(column, 0) + 1
    for column, hits in totals.items():
        if hits > best_hits:
            best_col, best_hits = column, hits
    return best_col


def _guess_contract_column(rows: Sequence[Sequence[Any]], start: int,
                           known: set) -> int:
    """ستونی که بیشترین هم‌خوانی را با شماره‌های قرارداد کتاب کار دارد."""
    if not known:
        return -1
    totals: Dict[int, int] = {}
    seen: Dict[int, int] = {}
    for index in range(start, min(len(rows), start + 400)):
        row = rows[index]
        for column, value in enumerate(row):
            key = normalize_lookup_key(txt(value))
            if not key:
                continue
            seen[column] = seen.get(column, 0) + 1
            if key in known:
                totals[column] = totals.get(column, 0) + 1
    best_col, best_ratio = -1, 0.0
    for column, hits in totals.items():
        ratio = hits / max(1, seen.get(column, 1))
        if hits >= 3 and ratio > best_ratio:
            best_col, best_ratio = column, ratio
    return best_col if best_ratio >= 0.35 else -1


@dataclass
class SystemUpdatePlan:
    source_path: str
    source_sheet: str
    header_row: int
    sub_project_column: int
    contract_column: int
    target_contract_column: int
    total_rows: int
    filtered_rows: int
    matched_rows: int
    method: str
    unmatched: List[str]
    rows: List[Tuple[int, List[Any]]]
    sheet_name: str
    first_column: int
    last_column: int
    notes: List[str]
    source_sheets: List[str]


def _column_letter(index: int) -> str:
    """شمارهٔ ستون (۱-پایه) به حرف اکسل."""
    letters = ""
    value = max(1, int(index))
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def plan_system_update(source_path: str, target_blob: bytes,
                       sheet_name: str = SHEET_NAME_DEFAULT,
                       first_column: int = SYSTEM_UPDATE_FIRST_COLUMN,
                       last_column: int = SYSTEM_UPDATE_LAST_COLUMN) -> SystemUpdatePlan:
    """نقشهٔ به‌روزرسانی را می‌سازد: فیلتر زیرپروژه + تطبیق ردیف‌ها با کتاب کار."""
    notes: List[str] = []
    source_names, source_sheet, rows = _open_rows(source_path)
    if not rows:
        raise ExcelDataError("شیت فایل سامانه خالی است.")
    header_row, sub_col, contract_col = _detect_header_row(rows)
    if sub_col < 0:
        sub_col = _guess_sub_project_column(rows, header_row + 1)
    if sub_col < 0:
        sub_col = 5  # ستون F
        notes.append("ستون «sub project» در سرستون‌ها پیدا نشد؛ ستون F استفاده شد.")
    if contract_col < 0:
        contract_col = 4  # ستون E به‌عنوان حدس اولیه

    sample = rows[header_row + 1: header_row + 60]
    if contract_col < 0 or not any(
        _looks_like_contract_number(row[contract_col] if contract_col < len(row) else None)
        for row in sample
    ):
        best_col, best_hits = -1, 0
        width = min(25, max((len(row) for row in sample[:30]), default=0))
        for column in range(width):
            hits = sum(
                1 for row in sample[:200]
                if column < len(row) and _looks_like_contract_number(row[column])
            )
            if hits > best_hits:
                best_col, best_hits = column, hits
        if best_col >= 0 and best_hits >= 3:
            contract_col = best_col
            notes.append("ستون شمارهٔ قرارداد فایل سامانه از روی داده‌ها پیدا شد.")

    if not target_blob:
        raise ExcelDataError("محتوای کتاب کار مقصد خالی است.")
    try:
        target = openpyxl.load_workbook(io.BytesIO(target_blob), data_only=False)
    except Exception as exc:
        raise ExcelDataError(f"کتاب کار دیتابیس باز نشد:\n{exc}") from exc
    try:
        names = [str(name) for name in target.sheetnames]
        target_name = sheet_name if sheet_name in names else ""
        if not target_name:
            for name in names:
                if normalize_per(name) == normalize_per(sheet_name or ""):
                    target_name = name
                    break
        if not target_name:
            target_name = names[0] if names else ""
        if not target_name:
            raise ExcelDataError("کتاب کار دیتابیس هیچ شیتی ندارد.")
        target_ws = target[target_name]
        width = max(1, min(int(target_ws.max_column or 1), 90))
        target_grid = [list(r) for r in target_ws.iter_rows(max_col=width,
                                                            values_only=True)]
        target_header_row, target_sub_col, target_contract_col = _detect_header_row(
            target_grid[:15]
        )
        first_data_row = max(4, target_header_row + 2)

        # ستون شمارهٔ قرارداد: اول از روی سرستون، بعد از روی هم‌خوانی با فایل ۰۱
        source_keys = set()
        for row in rows[header_row + 1:]:
            key = normalize_lookup_key(txt(row[contract_col] if contract_col < len(row) else ""))
            if key:
                source_keys.add(key)
        if source_keys and target_grid:
            hits: Dict[int, int] = {}
            for row in target_grid[first_data_row - 1:first_data_row - 1 + 4000]:
                for column, value in enumerate(row):
                    key = normalize_lookup_key(txt(value))
                    if key and key in source_keys:
                        hits[column] = hits.get(column, 0) + 1
            if hits:
                best_col = max(hits, key=lambda c: hits[c])
                header_hits = hits.get(target_contract_col, 0) if target_contract_col >= 0 else 0
                if hits[best_col] >= max(3, header_hits + 1):
                    if target_contract_col != best_col:
                        notes.append("ستون شمارهٔ قرارداد در کتاب کار از روی داده‌ها پیدا شد.")
                    target_contract_col = best_col
        if target_contract_col < 0:
            target_contract_col = 4  # ستون E

        target_rows: Dict[str, int] = {}
        target_name_rows: Dict[str, int] = {}
        target_name_counts: Dict[str, int] = {}
        data_rows: List[int] = []
        name_col = _header_column([target_grid[target_header_row]], 0,
                                  ("نام شرکت", "subcontractor", "sub contractor",
                                   "شرکت", "company"))
        if name_col < 0 or name_col == target_contract_col:
            name_col = 3
        for offset, values in enumerate(target_grid[first_data_row - 1:], start=first_data_row):
            if all(is_blank(value) for value in values):
                continue
            data_rows.append(offset)
            value = values[target_contract_col] if target_contract_col < len(values) else None
            key = normalize_lookup_key(txt(value))
            if key:
                target_rows.setdefault(key, offset)
            name_value = values[name_col] if name_col < len(values) else None
            name_key = normalize_lookup_key(txt(name_value))
            if name_key:
                target_name_counts[name_key] = target_name_counts.get(name_key, 0) + 1
                target_name_rows.setdefault(name_key, offset)
        if not data_rows:
            raise ExcelDataError("در کتاب کار دیتابیس هیچ ردیف داده‌ای پیدا نشد.")
        target_name_rows = {
            key: row for key, row in target_name_rows.items()
            if target_name_counts.get(key) == 1
        }
    finally:
        try:
            target.close()
        except Exception:
            pass

    filtered: List[List[Any]] = []
    total = 0
    for index in range(header_row + 1, len(rows)):
        row = rows[index]
        if not any(not is_blank(value) for value in row):
            continue
        total += 1
        if _matches_sub_project(row[sub_col] if sub_col < len(row) else None):
            filtered.append(row)

    if contract_col < 0 or not any(
        _looks_like_contract_number(row[contract_col] if contract_col < len(row) else None)
        for row in filtered
    ):
        guess = _guess_contract_column(rows, header_row + 1, set(target_rows))
        if guess >= 0:
            contract_col = guess
            notes.append("ستون شمارهٔ قرارداد از روی داده‌ها شناسایی شد.")

    # ۱) تطبیق با شمارهٔ قرارداد
    contract_mapping: List[Tuple[int, List[Any]]] = []
    contract_unmatched: List[str] = []
    for row in filtered:
        key = normalize_lookup_key(txt(row[contract_col] if contract_col < len(row) else ""))
        target_row = target_rows.get(key) if key else None
        if target_row:
            contract_mapping.append((target_row, list(row[first_column - 1:last_column])))
        else:
            contract_unmatched.append(
                txt(row[contract_col] if contract_col < len(row) else "") or "—"
            )

    # ۲) تطبیق با نام شرکت (اگر ستون نام در فایل سامانه باشد)
    name_mapping: List[Tuple[int, List[Any]]] = []
    name_unmatched: List[str] = []
    if target_name_rows and filtered:
        name_col, name_hits = -1, 0
        for column in range(len(filtered[0])):
            hits = sum(
                1 for row in filtered
                if column < len(row)
                and normalize_lookup_key(txt(row[column])) in target_name_rows
            )
            if hits > name_hits:
                name_col, name_hits = column, hits
        if name_col >= 0 and name_hits >= max(3, int(len(filtered) * 0.5)):
            for row in filtered:
                key = normalize_lookup_key(txt(row[name_col] if name_col < len(row) else ""))
                target_row = target_name_rows.get(key) if key else None
                if target_row:
                    name_mapping.append((target_row, list(row[first_column - 1:last_column])))
                else:
                    name_unmatched.append(
                        txt(row[contract_col] if contract_col < len(row) else "") or "—"
                    )

    # ۳) چسباندن بلوکی به‌همان روش دستی: L..AJ از بالا تا انتها، به‌ترتیب ردیف‌ها
    order_rows = min(len(filtered), len(data_rows))
    order_mapping: List[Tuple[int, List[Any]]] = [
        (data_rows[order], list(row[first_column - 1:last_column]))
        for order, row in enumerate(filtered[:order_rows])
    ]

    filtered_count = max(1, len(filtered))
    contract_ratio = len(contract_mapping) / filtered_count
    name_ratio = len(name_mapping) / filtered_count
    equal_rows = len(filtered) == len(data_rows)

    mapping: List[Tuple[int, List[Any]]] = []
    unmatched: List[str] = []
    method = "شمارهٔ قرارداد"
    if contract_ratio >= 0.9:
        mapping, unmatched = contract_mapping, contract_unmatched
    elif name_ratio >= 0.9:
        mapping, unmatched = name_mapping, name_unmatched
        method = "نام شرکت"
        notes.append("تطبیق از روی نام شرکت انجام شد (شمارهٔ قرارداد هم‌خوان نبود).")
    elif equal_rows and order_mapping:
        mapping = order_mapping
        method = "ترتیب ردیف‌ها"
        notes.append(
            "مثل روش دستی، ستون‌های L تا AJ به‌ترتیب ردیف‌ها روی شیت مقصد نوشته می‌شود "
            "(تعداد ردیف‌های «اجرا - مهر» و شیت مقصد برابر است)."
        )
    elif contract_ratio >= 0.5:
        mapping, unmatched = contract_mapping, contract_unmatched
    elif name_ratio >= 0.5:
        mapping, unmatched = name_mapping, name_unmatched
        method = "نام شرکت"
        notes.append("تطبیق از روی نام شرکت انجام شد (شمارهٔ قرارداد هم‌خوان نبود).")
    elif order_mapping:
        mapping = order_mapping
        method = "ترتیب ردیف‌ها"
        notes.append(
            f"تعداد ردیف‌ها برابر نبود (فایل سامانه: {fa(len(filtered))}، شیت مقصد: "
            f"{fa(len(data_rows))})؛ {fa(len(order_mapping))} ردیف اول به‌ترتیب نوشته شد."
        )
    if not filtered:
        raise ExcelDataError(
            "در فایل سامانه هیچ ردیفی با «اجرا - مهر» پیدا نشد.\n"
            f"شیت خوانده‌شده: «{source_sheet}» — ستون زیرپروژه: "
            f"{(get_column_letter(sub_col + 1) if get_column_letter else 'F')} "
            f"(ردیف سرستون {fa(header_row + 1)}).\n"
            "اگر ستون Sub Project جای دیگری است، سرستون‌ها را بررسی کنید."
        )
    if not mapping:
        raise ExcelDataError(
            "هیچ ردیفی از فایل سامانه با کتاب کار دیتابیس تطبیق پیدا نکرد.\n"
            f"ردیف‌های «اجرا - مهر»: {fa(len(filtered))} — شماره‌های نمونه: "
            + "، ".join((contract_unmatched or name_unmatched or [])[:5])
        )

    return SystemUpdatePlan(
        source_path=str(source_path),
        source_sheet=source_sheet,
        header_row=header_row + 1,
        sub_project_column=sub_col,
        contract_column=contract_col,
        target_contract_column=target_contract_col,
        total_rows=total,
        filtered_rows=len(filtered),
        matched_rows=len(mapping),
        method=method,
        unmatched=unmatched[:50],
        rows=mapping,
        sheet_name=target_name,
        first_column=first_column,
        last_column=last_column,
        notes=notes,
        source_sheets=source_names,
    )


def _sheet_system_scan(workbook, sheet_name: str, source_keys: set) -> Dict[str, Any]:
    """شیت را می‌کاود: سرستون‌ها، ستون شمارهٔ قرارداد و تعداد هم‌خوانی با فایل ۰۱.

    تشخیص فقط به نام سرستون‌ها وابسته نیست؛ اگر سرستون‌ها فارسی یا متفاوت باشند،
    ستونی که بیشترین شمارهٔ قرارداد مشترک با فایل ۰۱ را دارد پیدا می‌شود.
    """
    info: Dict[str, Any] = {
        "sheet": str(sheet_name), "score": -1, "header_row": -1, "sub_col": -1,
        "contract_col": -1, "hits": 0, "rows": 0,
    }
    try:
        sheet = workbook[sheet_name]
        width = min(int(sheet.max_column or 1), 70)
        head_rows = [list(r) for r in sheet.iter_rows(
            min_row=1, max_row=15, max_col=width, values_only=True
        )]
    except Exception:
        return info
    header_row, sub_col, contract_col = _detect_header_row(head_rows)
    info.update({"header_row": header_row, "sub_col": sub_col,
                 "contract_col": contract_col})
    start = header_row + 1 if header_row >= 0 else 0
    hits: Dict[int, int] = {}
    scanned = 0
    try:
        for row in sheet.iter_rows(min_row=max(1, start + 1), max_col=width,
                                   values_only=True):
            scanned += 1
            if scanned > 4000:
                break
            for column, value in enumerate(row):
                if value is None or value == "":
                    continue
                key = normalize_lookup_key(txt(value))
                if key and key in source_keys:
                    hits[column] = hits.get(column, 0) + 1
    except Exception:
        pass
    best_hits = max(hits.values()) if hits else 0
    info["rows"] = scanned
    info["hits"] = best_hits
    score = 0
    if contract_col >= 0:
        score += 30
    if sub_col >= 0:
        score += 20
    lowered = normalize_per(str(sheet_name)).strip().lower()
    if "کندو" in str(sheet_name):
        score += 25
    elif "خروجی" in str(sheet_name):
        score += 15
    elif "sheet 1" in lowered or "sheet1" in lowered:
        score += 10
    if best_hits:
        score += min(60, best_hits)
        if contract_col < 0:
            score += 15  # ستون از روی داده پیدا شد
    if is_dashboard_sheet_name(sheet_name):
        # شیت colorcode کتاب کار نهایی است، نه جای پیست خروجی سامانه؛
        # نوشتن عددهای ۰/۱ روی آن، علامت‌های OK/M را خراب می‌کند.
        score -= 80
    if score <= 0 or (contract_col < 0 and best_hits == 0):
        score = -1
    info["score"] = score
    return info


def _sheet_system_score(workbook, sheet_name: str, source_keys: set) -> int:
    """امتیاز شباهت یک شیت به خروجی سامانه (Contract Close Documet Count Report)."""
    return int(_sheet_system_scan(workbook, sheet_name, source_keys)["score"])


def sheet_scan_report(workbook, source_keys: set) -> List[Dict[str, Any]]:
    """گزارش کوتاه از همهٔ شیت‌ها (برای پیام خطا و فایل لاگ)."""
    report = []
    for name in workbook.sheetnames:
        info = _sheet_system_scan(workbook, str(name), source_keys)
        report.append({
            "sheet": str(name),
            "score": max(0, int(info["score"])),
            "header_row": int(info["header_row"]) + 1 if info["header_row"] >= 0 else 0,
            "sub_col": int(info["sub_col"]) + 1 if info["sub_col"] >= 0 else 0,
            "contract_col": int(info["contract_col"]) + 1 if info["contract_col"] >= 0 else 0,
            "hits": int(info["hits"]),
            "rows": int(info["rows"]),
        })
    return report


def detect_system_sheet(workbook, source_keys: set) -> str:
    """شیتی از کتاب کار که خروجی سامانه در آن پیست می‌شود (مثل «خروجی کندو»)."""
    best, best_score = "", -1
    for name in workbook.sheetnames:
        score = _sheet_system_score(workbook, str(name), source_keys)
        if score > best_score:
            best, best_score = str(name), score
    return best if best_score > 0 else ""


def count_workbook_contracts(path: str, sheet_name: str = "") -> int:
    """شمارش سریع ردیف‌های دارای داده در شیت colorcode (پیش از بارگذاری)."""
    if openpyxl is None:
        return 0
    try:
        book = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    except Exception:
        return 0
    try:
        target = ""
        if sheet_name and sheet_name in book.sheetnames:
            target = sheet_name
        elif "colorcode" in book.sheetnames:
            target = "colorcode"
        else:
            target = book.sheetnames[0] if book.sheetnames else ""
        if not target:
            return 0
        count = 0
        for index, row in enumerate(book[target].iter_rows(min_row=4, max_col=7,
                                                           values_only=True)):
            if any(not is_blank(value) for value in row):
                count += 1
            if index > 20000:
                break
        return count
    except Exception:
        return 0
    finally:
        try:
            book.close()
        except Exception:
            pass


def recalculate_with_excel(path: str) -> bool:
    """فایل را با خودِ اکسل باز، محاسبه و ذخیره می‌کند (فقط ویندوز).

    اگر اکسل نصب نباشد یا COM در دسترس نباشد، False برمی‌گردد و برنامه به
    روش قبلی (باز کردن دستی فایل و Ctrl+S) برمی‌گردد.
    """
    if os.name != "nt":
        return False
    source = Path(str(path)).expanduser()
    if not source.exists():
        return False
    quoted = str(source).replace("'", "''")
    script = (
        "$ErrorActionPreference = 'Stop'\n"
        "$excel = New-Object -ComObject Excel.Application\n"
        "$excel.Visible = $false\n"
        "$excel.DisplayAlerts = $false\n"
        "$excel.AskToUpdateLinks = $false\n"
        f"$wb = $excel.Workbooks.Open('{quoted}', 0, $false)\n"
        "$excel.CalculateFullRebuild()\n"
        "$wb.Save()\n"
        "$wb.Close($false)\n"
        "$excel.Quit()\n"
    )
    script_path = data_dir() / "recalc_excel.ps1"
    try:
        script_path.write_text(script, encoding="utf-8-sig")
        before = source.stat().st_mtime
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-File", str(script_path)],
            capture_output=True, timeout=600,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if completed.returncode != 0:
            log_event(
                "محاسبهٔ خودکار با اکسل انجام نشد: "
                + (completed.stderr or b"").decode("utf-8", "replace")[-400:]
            )
            return False
        changed = source.stat().st_mtime > before
        log_event(f"فایل با اکسل بازمحاسبه و ذخیره شد: {source.name} (تغییر: {changed})")
        return True
    except Exception as exc:
        log_event(f"محاسبهٔ خودکار با اکسل ممکن نشد: {type(exc).__name__}: {exc}")
        return False
    finally:
        try:
            script_path.unlink()
        except Exception:
            pass


def validate_system_source(source_path: str) -> Dict[str, Any]:
    """فایل ۰۱ را قبل از به‌روزرسانی بررسی می‌کند.

    خروجی: تعداد ردیف‌ها، تعداد ردیف‌های «اجرا - مهر» و شیت/ستون‌های پیداشده.
    اگر فایل به‌جای خروجی سامانه، خودِ کتاب کار باشد (شیت colorcode داشته باشد)
    یا هیچ ردیف «اجرا - مهر» نداشته باشد، خطای روشن داده می‌شود.
    """
    path = str(source_path or "").strip()
    if not path or not Path(path).exists():
        raise ExcelDataError(f"فایل انتخاب‌شده پیدا نشد:\n{path or '—'}")
    if has_colorcode_sheet(path):
        raise ExcelDataError(
            "این فایل خودش «کتاب کار» است (شیت colorcode دارد)، نه خروجی سامانه.\n\n"
            "برای به‌روزرسانی، فایل خروجی سامانه را انتخاب کنید: "
            "Contract-Info-Report.xlsx (همان فایلی که ستون Sub Project دارد)."
        )
    names, sheet, rows = _open_rows(path)
    if not rows:
        raise ExcelDataError("فایل انتخاب‌شده هیچ داده‌ای ندارد.")
    header_row, sub_col, contract_col = _detect_header_row(rows)
    if sub_col < 0:
        sub_col = _guess_sub_project_column(rows, header_row + 1)
    matched = 0
    for row in rows[header_row + 1:]:
        if sub_col >= 0 and sub_col < len(row) and _matches_sub_project(row[sub_col]):
            matched += 1
    info = {
        "path": path,
        "sheet": sheet,
        "rows": len(rows),
        "matched": matched,
        "sub_col": sub_col + 1 if sub_col >= 0 else 0,
        "contract_col": contract_col + 1 if contract_col >= 0 else 0,
    }
    if matched < 20:
        raise ExcelDataError(
            "این فایل به‌نظر خروجی سامانه نیست: فقط "
            f"{fa(matched)} ردیف با «اجرا - مهر» در آن پیدا شد "
            f"(از {fa(len(rows))} ردیف، شیت «{sheet}»).\n\n"
            "فایل درست، خروجی «Contract Close Documet Count Report» است که ستون "
            "«Sub Project» آن مقدار «اجرا - مهر» دارد و حدود ۱۲۰۰ ردیف است.\n"
            "همان فایل را انتخاب کنید و دوباره «به‌روزرسانی» را بزنید."
        )
    return info


def source_contract_keys(source_path: str) -> set:
    """کلید شمارهٔ قراردادهای فایل سامانه (برای پیدا کردن شیت مقصد)."""
    try:
        _, _, rows = _open_rows(source_path)
    except Exception:
        return set()
    if not rows:
        return set()
    header_row, _sub_col, contract_col = _detect_header_row(rows)
    if contract_col < 0:
        contract_col = 4
    keys = set()
    for row in rows[header_row + 1:]:
        key = normalize_lookup_key(txt(row[contract_col] if contract_col < len(row) else ""))
        if key:
            keys.add(key)
    return keys


def prepare_system_copy(plan: SystemUpdatePlan, base_blob: bytes,
                        out_path: str) -> Dict[str, Any]:
    """نسخهٔ تازه‌ای از کتاب کار می‌سازد که فقط مقادیر L..AJ در شیت سامانه عوض شده‌اند.

    فایل اصلی هرگز بازنویسی نمی‌شود و سلول‌هایی که فرمول دارند دست‌نخورده می‌مانند؛
    بنابراین فرمول‌ها و ساختار کتاب کار سالم می‌ماند.
    """
    if openpyxl is None:
        raise ExcelDataError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(base_blob), data_only=False)
    except Exception as exc:
        raise ExcelDataError(f"کتاب کار باز نشد:\n{exc}") from exc
    try:
        if plan.sheet_name not in workbook.sheetnames:
            raise ExcelDataError(f"شیت «{plan.sheet_name}» در کتاب کار پیدا نشد.")
        sheet = workbook[plan.sheet_name]
        dashboard_target = is_dashboard_sheet_name(plan.sheet_name)
        written = 0
        formula_cells = 0
        merged_cells = 0
        cleaned = 0
        for target_row, values in plan.rows:
            for offset, value in enumerate(values):
                column = plan.first_column + offset
                if dashboard_target:
                    if column > DASHBOARD_DOC_LAST_COLUMN:
                        continue  # ستون‌های جمع/درصد جای دیگری در شیت داشبورد دارند
                    if column >= DASHBOARD_DOC_FIRST_COLUMN:
                        value = status_mark_value(value)
                cell = sheet.cell(row=int(target_row), column=int(column))
                current = getattr(cell, "value", None)
                if isinstance(current, str) and current.startswith("="):
                    formula_cells += 1
                    continue
                if is_blank(value):
                    continue
                if current == value:
                    continue
                try:
                    cell.value = value
                except AttributeError:
                    merged_cells += 1  # سلول ادغام‌شده در اکسل قابل نوشتن نیست
                    continue
                written += 1
        # آسیب نسخه‌های قبل (۰/۱ در ستون‌های وضعیت) در همهٔ شیت‌های داشبورد
        # ترمیم می‌شود؛ حتی وقتی مقصد پیست شیت دیگری باشد.
        try:
            cleaned = repair_dashboard_marks(workbook)
        except Exception as exc:
            log_event(f"ترمیم علامت‌های وضعیت ممکن نشد: {type(exc).__name__}: {exc}")
        try:
            workbook.calculation.fullCalcOnLoad = True
        except Exception:
            pass
        target = Path(str(out_path)).expanduser()
        target.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(str(target))
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    return {
        "output_path": str(target),
        "written_cells": written,
        "formula_cells": formula_cells,
        "merged_cells": merged_cells,
        "cleaned_marks": cleaned,
        "skipped_formulas_sheet": plan.sheet_name,
    }


def apply_system_update(plan: SystemUpdatePlan, target_blob: bytes) -> Tuple[bytes, int]:
    """مقادیر L..AJ را روی کتاب کار اعمال می‌کند و بایت‌های تازه را برمی‌گرداند."""
    if openpyxl is None:
        raise ExcelDataError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    workbook = openpyxl.load_workbook(io.BytesIO(target_blob), data_only=False)
    try:
        if plan.sheet_name not in workbook.sheetnames:
            raise ExcelDataError(f"شیت «{plan.sheet_name}» در کتاب کار پیدا نشد.")
        sheet = workbook[plan.sheet_name]
        updated = 0
        for target_row, values in plan.rows:
            for offset, value in enumerate(values):
                column = plan.first_column + offset
                if column > plan.last_column:
                    break
                if isinstance(value, str):
                    value = value.strip()
                    if value.lower() in _SYSTEM_UPDATE_OK_WORDS:
                        value = "ok"
                if is_dashboard_sheet_name(plan.sheet_name):
                    if column > DASHBOARD_DOC_LAST_COLUMN:
                        continue
                    if column >= DASHBOARD_DOC_FIRST_COLUMN:
                        value = status_mark_value(value)
                cell = sheet.cell(row=target_row, column=column)
                if is_blank(value):
                    # مقدار خالی در فایل سامانه، دادهٔ موجود کتاب کار را پاک نمی‌کند
                    # (همان رفتار نسخهٔ اولیه) تا اعداد داشبورد تغییر نکنند.
                    continue
                if cell.value == value:
                    continue
                cell.value = value
                updated += 1
        # اکسل بعد از ذخیره، ستون‌های خالی انتهایی را حذف می‌کند و عرض شیت به ۵۶
        # می‌رسد؛ برنامه ۵۸ ستون انتظار دارد، پس یک سلول خالی برای حفظ عرض می‌گذاریم.
        if int(sheet.max_column or 0) < REQUIRED_COLUMNS:
            sheet.cell(row=1, column=REQUIRED_COLUMNS, value="")
        buffer = io.BytesIO()
        workbook.save(buffer)
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    return buffer.getvalue(), updated


def preview_system_update(plan: SystemUpdatePlan) -> str:
    """متن خلاصهٔ عملیات برای نمایش به کاربر."""
    lines = [
        f"فایل سامانه: {Path(plan.source_path).name}",
        f"شیت: {plan.source_sheet} — ردیف سرستون: {fa(plan.header_row)}",
        f"ستون زیرپروژه: {_column_letter(plan.sub_project_column + 1)} · "
        f"ستون شمارهٔ قرارداد: {_column_letter(plan.contract_column + 1)}",
        f"ردیف‌های داده: {fa(plan.total_rows)} — «اجرا - مهر»: {fa(plan.filtered_rows)}",
        f"تطبیق‌شده با کتاب کار: {fa(plan.matched_rows)} ردیف (روش: {plan.method})",
        f"ستون‌های {_column_letter(plan.first_column)} تا {_column_letter(plan.last_column)} "
        f"در شیت «{plan.sheet_name}» به‌روزرسانی می‌شوند.",
    ]
    if plan.unmatched:
        lines.append(
            f"بدون تطبیق: {fa(len(plan.unmatched))} ردیف — نمونه: "
            + "، ".join(plan.unmatched[:4])
        )
    for note in plan.notes:
        lines.append(f"یادداشت: {note}")
    return "\n".join(lines)


def workbook_sheet_names(path: str) -> List[str]:
    """فهرست شیت‌های یک فایل اکسل (برای تشخیص فایل سامانه)."""
    if openpyxl is None:
        return []
    try:
        workbook = openpyxl.load_workbook(str(path), read_only=True)
    except Exception:
        return []
    try:
        return [str(name) for name in workbook.sheetnames]
    finally:
        try:
            workbook.close()
        except Exception:
            pass


def has_colorcode_sheet(path: str) -> bool:
    for name in workbook_sheet_names(path):
        if normalize_per(name).strip() == "colorcode":
            return True
    return False


def update_local_from_system(source_path: str, target_path: str,
                             sheet_name: str = SHEET_NAME_DEFAULT) -> Dict[str, Any]:
    """همان به‌روزرسانی فایل ۰۱، اما روی یک فایل اکسل محلی (بدون دیتابیس)."""
    target = Path(str(target_path)).expanduser()
    if not target.exists() or not target.is_file():
        raise ExcelDataError(f"فایل اکسل مقصد پیدا نشد:\n{target}")
    blob = target.read_bytes()
    plan = plan_system_update(source_path, blob, sheet_name=sheet_name)
    backup = backup_workbook(blob, target.name)
    new_blob, cells = apply_system_update(plan, blob)
    stamp = datetime.now().strftime("%Y%m%d-%H%M")
    output = target.with_name(f"{target.stem}-به‌روزرسانی-{stamp}.xlsx")
    output.write_bytes(new_blob)
    log_event(
        f"به‌روزرسانی محلی از فایل سامانه: {plan.filtered_rows} ردیف «اجرا - مهر»، "
        f"{plan.matched_rows} تطبیق، {cells} سلول → {output}"
    )
    return {
        "plan": plan,
        "updated_cells": cells,
        "backup_path": backup,
        "output_path": str(output),
        "summary": preview_system_update(plan),
    }


MANAGED_WORKBOOK_NAME = "MehrWorkbook.xlsx"
# سرستون‌هایی که برای تشخیص ستون‌های فایل سامانه جست‌وجو می‌شوند (کلید ← نام‌های ممکن)
_SOURCE_HEADER_HINTS: Dict[str, Tuple[str, ...]] = {
    "name": ("نام شرکت", "نام پیمانکار", "شرکت", "پیمانکار", "contractor", "company"),
    "subject": ("موضوع قرارداد", "موضوع", "شرح قرارداد", "subject", "description"),
    "employer": ("کارفرما", "employer", "client"),
    "progress": ("پیشرفت کارکرد", "درصد پیشرفت", "progress"),
    "contract_type": ("نوع قرارداد", "نوع پیمان", "contract type"),
}


GENERATED_WORKBOOK_MARK = "mehr-generated-from-file01"
GENERATED_WORKBOOK_TITLE = "کتاب کار داخلی برنامه"


def managed_workbook_path() -> Path:
    """مسیر کتاب کار داخلی برنامه (جایگزین فایل ۰۲ برای کاربر)."""
    return data_dir() / MANAGED_WORKBOOK_NAME


def is_generated_workbook(path: str) -> bool:
    """آیا این فایل را خودِ برنامه از فایل ۰۱ ساخته (و نه کتاب کار اصلی کاربر)؟"""
    try:
        probe = Path(str(path))
        if not probe.exists():
            return False
        workbook = openpyxl.load_workbook(str(probe), read_only=True)
    except Exception:
        return False
    try:
        props = getattr(workbook, "properties", None)
        keywords = str(getattr(props, "keywords", "") or "")
        if GENERATED_WORKBOOK_MARK in keywords:
            return True
        for sheet in workbook.worksheets:
            try:
                first = next(sheet.iter_rows(min_row=1, max_row=1, max_col=1, values_only=True))
            except Exception:
                continue
            text = str((first or [""])[0] or "")
            if text.startswith(GENERATED_WORKBOOK_TITLE):
                return True
            break
        return False
    finally:
        try:
            workbook.close()
        except Exception:
            pass


_COMPANION_HINTS = {
    "contract_id": ("Contract ID", "شناسه قرارداد"),
    "subcontractor": ("SubContractor", "Sub Contractor", "پیمانکار جزء", "نام شرکت"),
    "contract_no": ("Contract No", "شماره قرارداد", "Contract Number"),
    "subject": ("Contract Subject", "موضوع قرارداد"),
    "exec_type": ("Excution Type", "Execution Type", "نوع اجرا", "نوع قرارداد"),
    "progress": ("درصد پیشرفت مدارک اختتام", "درصد پیشرفت", "progress"),
}

_COMPANION_TARGET_COLUMNS = {
    "contract_id": 2,      # ستون C
    "subcontractor": 3,    # ستون D
    "contract_no": 4,      # ستون E
    "exec_type": 5,        # ستون F
    "subject": 6,          # ستون G
    "progress": 36,        # ستون AK
}


def _find_companion_sheet(workbook, target_name: str, target_rows: int = 0):
    """شیت خروجی سامانه (با ستون Sub Project، شمارهٔ قرارداد و ستون درصد) در همان فایل.

    بین شیت‌های کاندید، شیتی انتخاب می‌شود که تعداد ردیف‌هایش به شیت اصلی نزدیک
    باشد و ستون درصدش واقعاً کسری بین ۰ و ۱ داشته باشد.
    """
    best = (None, -1, [], -1)
    best_score = -1
    for name in workbook.sheetnames:
        if str(name) == str(target_name):
            continue
        try:
            sheet = workbook[name]
            rows = [list(r) for r in sheet.iter_rows(min_row=1, max_row=15, max_col=60,
                                                     values_only=True)]
        except Exception:
            continue
        header_row, sub_col, contract_col = _detect_header_row(rows)
        if header_row < 0 or sub_col < 0 or contract_col < 0:
            continue
        header = rows[header_row]
        if not _looks_like_sub_project_header(
            header[sub_col] if sub_col < len(header) else ""
        ):
            continue
        percent_col = _header_column([header], 0, _COMPANION_HINTS["progress"])
        if percent_col < 0:
            continue
        data_rows = 0
        fractions = 0
        numbers = 0
        try:
            for index, row in enumerate(sheet.iter_rows(min_row=header_row + 2,
                                                        max_col=percent_col + 1,
                                                        values_only=True)):
                if any(not is_blank(value) for value in row[:8]):
                    data_rows += 1
                value = row[percent_col] if percent_col < len(row) else None
                number = num(value)
                if number is not None:
                    numbers += 1
                    if 0 <= number <= 1:
                        fractions += 1
                if index > 1500:
                    break
        except Exception:
            continue
        if data_rows < 10:
            continue
        ratio = fractions / numbers if numbers else 0.0
        if ratio < 0.6:
            continue
        score = 100.0 * ratio
        if target_rows:
            score += 100.0 * max(0.0, 1.0 - abs(data_rows - target_rows) / max(1, target_rows))
        lowered = normalize_per(str(name)).strip().lower()
        if "کندو" in str(name):
            score += 30
        elif "خروجی" in str(name):
            score += 20
        elif "sheet 1" in lowered or "sheet1" in lowered:
            score += 10
        if score > best_score:
            best_score = score
            best = (sheet, header_row, header, percent_col)
    return best


def repair_rows_from_companion(workbook, target_name: str,
                               raw_rows: List[List[Any]]) -> Dict[str, Any]:
    """ستون‌های شناسایی و درصد را از شیت خروجی سامانه همان فایل بازسازی می‌کند.

    وقتی کتاب کار با ابزاری بیرون از اکسل ذخیره شده باشد، ستون‌های فرمولی
    (نام شرکت، شمارهٔ قرارداد، درصد) مقدار ندارند. این مقادیر در همان فایل و در
    شیت «خروجی کندو» موجودند؛ فقط همان‌ها در حافظه جای خالی‌ها گذاشته می‌شوند
    (هیچ فایلی نوشته نمی‌شود). درصد کل هم همان فرمول خود کتاب کار
    (میانگین ستون ACK) روی همین مقادیر است.
    """
    result: Dict[str, Any] = {"filled": 0, "progress": 0, "sheet": "", "total": None}
    if not raw_rows or len(raw_rows) < 4:
        return result
    sheet, companion_header_row, header_values, percent_col = _find_companion_sheet(
        workbook, target_name, max(0, len(raw_rows) - 3)
    )
    if sheet is None:
        return result
    result["sheet"] = str(sheet.title)
    columns: Dict[str, int] = {}
    for key, hints in _COMPANION_HINTS.items():
        if key == "progress":
            columns[key] = percent_col
            continue
        column = _header_column([header_values], 0, hints)
        if column >= 0:
            columns[key] = column
    if "contract_no" not in columns or "subcontractor" not in columns:
        return result

    companion_start = companion_header_row + 1
    target_header_row, _sub, _contract = _detect_header_row(raw_rows[:12])
    if target_header_row < 0:
        target_header_row = 2
    target_start = target_header_row + 1
    delta = companion_start - target_start
    companion_rows = [list(r) for r in sheet.iter_rows(values_only=True)]
    progress_values: List[float] = []
    filled = 0
    progress_rows = 0
    for index in range(target_start, len(raw_rows)):
        companion_index = index + delta
        if companion_index >= len(companion_rows) or companion_index < 0:
            continue
        source = companion_rows[companion_index]
        target = raw_rows[index]
        if not any(not is_blank(value) for value in target[:7]):
            continue
        changed = False
        for key, target_column in _COMPANION_TARGET_COLUMNS.items():
            if key not in columns:
                continue
            source_column = columns[key]
            value = source[source_column] if source_column < len(source) else None
            if is_blank(value):
                continue
            while len(target) <= target_column:
                target.append(None)
            if not is_blank(target[target_column]):
                continue
            number = num(value) if key == "progress" else None
            if key == "progress" and (number is None or not 0 <= number <= 1):
                continue  # ستون درصد نیست (مثلاً مبلغ)
            target[target_column] = value
            changed = True
            if key == "progress" and number is not None and number > 0:
                progress_values.append(number)
                progress_rows += 1
        if changed:
            filled += 1
    total = None
    if progress_values and len(raw_rows) > 1:
        if len(raw_rows[1]) <= 36:
            raw_rows[1].extend([None] * (37 - len(raw_rows[1])))
        if is_blank(raw_rows[1][36]):
            # همان فرمول کتاب کار: SUBTOTAL(101) یعنی میانگین همان ستون
            total = round(sum(progress_values) / len(progress_values), 6)
            raw_rows[1][36] = total
    result.update({"filled": filled, "progress": progress_rows, "total": total})
    return result


def workbook_health(blob: bytes, max_rows: int = 4000) -> Dict[str, Any]:
    """سالم بودن کتاب کار: آیا فرمول‌ها مقدار محاسبه‌شده (cached) دارند؟

    وقتی فایلی با ابزاری غیر از اکسل (مثل openpyxl) دوباره ذخیره شود، فرمول‌ها
    می‌مانند ولی مقدار محاسبه‌شده‌شان پاک می‌شود؛ نتیجه این است که همهٔ ستون‌های
    فرمولی (نام شرکت، شمارهٔ قرارداد، درصد، وضعیت) در خواندن، خالی/صفر می‌شوند.
    """
    result: Dict[str, Any] = {
        "ok": False, "formulas": 0, "missing": 0, "ratio": 0.0,
        "damaged": False, "sheet": "", "sheets": [],
    }
    if not blob or openpyxl is None:
        return result
    try:
        formulas_book = openpyxl.load_workbook(
            io.BytesIO(blob), data_only=False, read_only=True
        )
        values_book = openpyxl.load_workbook(
            io.BytesIO(blob), data_only=True, read_only=True
        )
    except Exception:
        return result
    try:
        result["sheets"] = [str(name) for name in formulas_book.sheetnames]
        target = ""
        for name in formulas_book.sheetnames:
            if str(name).strip().lower() == SHEET_NAME_DEFAULT:
                target = name
                break
        if not target and formulas_book.sheetnames:
            target = formulas_book.sheetnames[0]
        if not target:
            return result
        result["sheet"] = str(target)
        result["ok"] = True
        formulas = 0
        missing = 0
        rows_f = formulas_book[target].iter_rows(values_only=True)
        rows_v = values_book[target].iter_rows(values_only=True)
        for index, (row_f, row_v) in enumerate(zip(rows_f, rows_v)):
            if index >= max_rows:
                break
            for column, value in enumerate(row_f or ()):
                if isinstance(value, str) and value.startswith("="):
                    formulas += 1
                    other = row_v[column] if column < len(row_v or ()) else None
                    if other is None or (isinstance(other, str) and not other.strip()):
                        missing += 1
        result["formulas"] = formulas
        result["missing"] = missing
        result["ratio"] = round(missing / formulas, 3) if formulas else 0.0
        result["damaged"] = bool(formulas >= 20 and result["ratio"] >= 0.5)
        return result
    finally:
        for book in (formulas_book, values_book):
            try:
                book.close()
            except Exception:
                pass


DAMAGED_WORKBOOK_TITLE = "کتاب کار ذخیره‌شده ناقص است"


def damaged_workbook_message(health: Dict[str, Any]) -> str:
    """توضیح روشن برای کتاب کاری که فرمول‌هایش مقدار ندارند."""
    percent = fa(round(float(health.get("ratio") or 0) * 100))
    return (
        "این کتاب کار با ابزاری بیرون از اکسل دوباره ذخیره شده و "
        f"{percent}٪ فرمول‌هایش مقدار محاسبه‌شده ندارند؛ به همین دلیل ستون‌های "
        "فرمولی (نام شرکت، شمارهٔ قرارداد، درصد پیشرفت و وضعیت‌ها) خالی خوانده "
        "می‌شوند و کارت‌ها صفر می‌شوند.\n\n"
        "کتاب کار اصلی خودتان را (همان فایل ۰۲ با شیت colorcode) یک بار باز کنید "
        "و ذخیره کنید، بعد در برنامه «بارگذاری فایل اکسل» را بزنید. "
        "از این نسخه به بعد برنامه هیچ کتاب کاری را بازنویسی نمی‌کند."
    )


def inspect_workbook_bytes(blob: bytes) -> Dict[str, Any]:
    """ویژگی‌های مهم کتاب کار: ساختهٔ برنامه؟ درصد دارد؟ علامت وضعیت دارد؟"""
    info: Dict[str, Any] = {
        "generated": False, "has_progress": False, "mark_rows": 0,
        "sheets": [], "colorcode": False, "ok": False,
    }
    if not blob or openpyxl is None:
        return info
    try:
        workbook = openpyxl.load_workbook(
            io.BytesIO(blob), data_only=True, read_only=True
        )
    except Exception:
        return info
    try:
        info["sheets"] = [str(name) for name in workbook.sheetnames]
        info["ok"] = True
        props = getattr(workbook, "properties", None)
        if GENERATED_WORKBOOK_MARK in str(getattr(props, "keywords", "") or ""):
            info["generated"] = True
        target = None
        for name in workbook.sheetnames:
            if str(name).strip().lower() == SHEET_NAME_DEFAULT:
                target = name
                break
        info["colorcode"] = target is not None
        if target is None:
            return info
        sheet = workbook[target]
        marks = 0
        for index, row in enumerate(
            sheet.iter_rows(min_row=1, max_col=56, values_only=True)
        ):
            if index == 0:
                if str((row or [""])[0] or "").startswith(GENERATED_WORKBOOK_TITLE):
                    info["generated"] = True
                continue
            if index == 1 and len(row) > 36:
                number = num(row[36])
                if number is not None and number > 0:
                    info["has_progress"] = True
            if any(
                is_status_mark(row[column] if column < len(row) else None)
                for column in range(11, 31)
            ):
                marks += 1
            if info["has_progress"] and marks:
                break  # برای تصمیم‌گیری همین اندازه کافی است
            if info["generated"] and index > 1:
                break  # کتاب کار ساخته‌شده از ۰۱ نیازی به شمارش علامت‌ها ندارد
        info["mark_rows"] = marks
        return info
    finally:
        try:
            workbook.close()
        except Exception:
            pass


def inspect_workbook_file(path: str) -> Dict[str, Any]:
    """همان بازرسی، اما برای فایل روی دیسک."""
    try:
        probe = Path(str(path)).expanduser()
        if not probe.exists() or not probe.is_file():
            return {}
        return inspect_workbook_bytes(probe.read_bytes())
    except Exception:
        return {}


def find_real_workbook(source_path: str = "",
                       extra_paths: Sequence[str] = ()) -> str:
    """کتاب کار اصلی کاربر را پیدا می‌کند (نه کتاب کار ساخته‌شده از فایل ۰۱).

    اول کتاب کاری که واقعاً درصد/علامت دارد؛ اگر نبود، اولین کتاب کار واقعی
    دارای شیت colorcode.
    """
    candidates: List[str] = []
    seen = set()

    def add(path: Any) -> None:
        text = str(path or "").strip()
        if not text:
            return
        key = text.lower()
        if key in seen:
            return
        seen.add(key)
        candidates.append(text)

    for item in extra_paths:
        add(item)
    try:
        seed = find_seed_workbook(source_path, extra_paths)
        add(seed)
    except Exception:
        pass
    add(ORIGINAL_EXCEL_PATH)
    for path in candidates:
        quality = inspect_workbook_file(path)
        if quality.get("ok") and quality.get("colorcode") and not quality.get("generated"):
            if quality.get("has_progress") or quality.get("mark_rows"):
                return path
    for path in candidates:
        quality = inspect_workbook_file(path)
        if quality.get("ok") and quality.get("colorcode") and not quality.get("generated"):
            return path
    return ""


def _header_column(rows: Sequence[Sequence[Any]], header_row: int,
                   hints: Sequence[str]) -> int:
    """ستونی که سرستونش با یکی از نام‌های داده‌شده هم‌خوان است."""
    if not rows or header_row < 0 or header_row >= len(rows):
        return -1
    header = [txt(value) for value in rows[header_row]]
    normalized = [_norm_words(value) for value in header]
    for hint in hints:
        needle = _norm_words(hint)
        if not needle:
            continue
        for column, value in enumerate(normalized):
            if not value:
                continue
            if value == needle or (len(needle) > 3 and needle in value):
                return column
    return -1


def build_workbook_from_system(source_path: str, sheet_name: str = SHEET_NAME_DEFAULT,
                               first_column: int = SYSTEM_UPDATE_FIRST_COLUMN,
                               last_column: int = SYSTEM_UPDATE_LAST_COLUMN) -> Dict[str, Any]:
    """از فایل ۰۱ سامانه یک کتاب کار کامل (شیت colorcode) می‌سازد.

    فایل ۰۲ لازم نیست: ردیف‌های «اجرا - مهر» ستون Sub Project انتخاب می‌شوند،
    ستون‌های L تا AJ عیناً منتقل می‌شوند و ستون‌های شناسایی (ردیف، نام شرکت،
    شمارهٔ قرارداد، نوع، موضوع، کارفرما) از خود فایل سامانه پر می‌شوند.
    """
    if openpyxl is None:
        raise ExcelDataError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    sheet_names, source_sheet, rows = _open_rows(source_path)
    if not rows:
        raise ExcelDataError("شیت فایل سامانه خالی است.")
    header_row, sub_col, contract_col = _detect_header_row(rows)
    notes: List[str] = []
    if sub_col < 0:
        sub_col = _guess_sub_project_column(rows, header_row + 1)
    if sub_col < 0:
        sub_col = 5  # ستون F
        notes.append("ستون «sub project» در سرستون‌ها پیدا نشد؛ ستون F استفاده شد.")
    if contract_col < 0:
        # اگر سرستون «شمارهٔ قرارداد» نبود، از روی شکل داده‌ها (وجود عدد و حرف)
        # ستون قرارداد را پیدا می‌کنیم.
        for column in range(max((len(row) for row in rows[: header_row + 2]), default=0)):
            hits = sum(
                1 for row in rows[header_row + 1: header_row + 30]
                if column < len(row) and _looks_like_contract_number(row[column])
            )
            if hits >= 3:
                contract_col = column
                break
    if contract_col < 0:
        contract_col = 4  # ستون E
        notes.append("ستون شمارهٔ قرارداد پیدا نشد؛ ستون E استفاده شد.")

    name_col = _header_column(rows, header_row, _SOURCE_HEADER_HINTS["name"])
    subject_col = _header_column(rows, header_row, _SOURCE_HEADER_HINTS["subject"])
    employer_col = _header_column(rows, header_row, _SOURCE_HEADER_HINTS["employer"])
    progress_col = _header_column(rows, header_row, _SOURCE_HEADER_HINTS["progress"])
    type_col = _header_column(rows, header_row, _SOURCE_HEADER_HINTS["contract_type"])
    if name_col < 0:
        name_col = 1  # ستون B به‌عنوان حدس
        notes.append("ستون «نام شرکت» در سرستون‌ها پیدا نشد؛ ستون B استفاده شد.")

    filtered: List[List[Any]] = []
    total = 0
    for index in range(header_row + 1, len(rows)):
        row = rows[index]
        if not any(not is_blank(value) for value in row):
            continue
        total += 1
        if _matches_sub_project(row[sub_col] if sub_col < len(row) else None):
            filtered.append(row)
    if not filtered:
        raise ExcelDataError(
            "در فایل سامانه هیچ ردیفی با «اجرا - مهر» پیدا نشد.\n"
            f"شیت خوانده‌شده: «{source_sheet}» — ستون زیرپروژه: "
            f"{_column_letter(sub_col + 1)} (ردیف سرستون {fa(header_row + 1)})."
        )

    def cell(row: Sequence[Any], column: int) -> Any:
        if column < 0 or column >= len(row):
            return None
        value = row[column]
        if isinstance(value, str):
            value = value.strip() or None
        return value

    workbook = openpyxl.Workbook()
    try:
        workbook.properties.keywords = GENERATED_WORKBOOK_MARK
        workbook.properties.title = (
            f"{GENERATED_WORKBOOK_TITLE} — ساخته‌شده از {Path(str(source_path)).name}"
        )
        workbook.properties.description = (
            "این کتاب کار را برنامه از خروجی سامانه ساخت؛ کتاب کار اصلی نیست و "
            "ستون‌های وضعیت را ندارد."
        )
    except Exception:
        pass
    sheet = workbook.active
    sheet.title = (sheet_name or SHEET_NAME_DEFAULT)[:31]
    sheet.sheet_view.rightToLeft = True
    sheet.cell(row=1, column=1,
               value=f"کتاب کار داخلی برنامه — ساخته‌شده از {Path(str(source_path)).name}")
    headers = {
        1: "ردیف", 2: "کد", 3: "وضعیت", 4: "نام شرکت", 5: "شماره قرارداد",
        6: "نوع قرارداد", 7: "موضوع قرارداد", 8: "کارفرما", 11: "پیشرفت کارکرد",
    }
    for column, title in headers.items():
        sheet.cell(row=3, column=column, value=title)
    # عنوان ستون‌های کسورات (L تا W) از سرستون خود فایل سامانه گرفته می‌شود.
    for column in range(first_column, min(last_column, 34) + 1):
        title = txt(rows[header_row][column - 1]) if header_row < len(rows) else ""
        sheet.cell(row=3, column=column, value=title or f"کسورات {fa(column - first_column + 1)}")

    written = 0
    status_marks = 0
    progress_values: List[float] = []
    target_row = 4
    for order, row in enumerate(filtered, start=1):
        sheet.cell(row=target_row, column=1, value=order)
        sheet.cell(row=target_row, column=4, value=cell(row, name_col))
        sheet.cell(row=target_row, column=5, value=cell(row, contract_col))
        sheet.cell(row=target_row, column=6,
                   value=cell(row, type_col) if type_col >= 0 else cell(row, sub_col))
        sheet.cell(row=target_row, column=7, value=cell(row, subject_col))
        sheet.cell(row=target_row, column=8, value=cell(row, employer_col))
        status_marks += sum(
            1 for _key, _label, flag_column in FLAG_DEFS
            if flag_column < len(row) and is_status_mark(row[flag_column])
        )
        progress_raw = cell(row, progress_col)
        sheet.cell(row=target_row, column=11, value=progress_raw)
        progress_number = num(progress_raw) if progress_col >= 0 else None
        if progress_number is not None and 0 <= progress_number <= 100:
            fraction = progress_number / 100.0 if progress_number > 1.5 else progress_number
            progress_values.append(fraction)
            sheet.cell(row=target_row, column=SYSTEM_UPDATE_LAST_COLUMN + 1, value=fraction)
        for offset in range(max(0, last_column - first_column + 1)):
            column = first_column + offset
            value = cell(row, column - 1)
            if isinstance(value, str) and value.lower() in _SYSTEM_UPDATE_OK_WORDS:
                value = "ok"
            if DASHBOARD_DOC_FIRST_COLUMN <= column <= DASHBOARD_DOC_LAST_COLUMN:
                value = status_mark_value(value)
            if value is None:
                continue
            sheet.cell(row=target_row, column=column, value=value)
            written += 1
        target_row += 1
    if progress_values:
        # ردیف ۲ ستون ۳۷ در کتاب کار اصلی، درصد پیشرفت کل است؛ همان را می‌گذاریم.
        sheet.cell(row=2, column=37,
                   value=round(sum(progress_values) / len(progress_values), 4))
        notes.append(
            f"درصد پیشرفت از ستون «{txt(rows[header_row][progress_col])}» "
            f"فایل سامانه محاسبه شد (میانگین {fa(len(progress_values))} ردیف)."
        )
    # اکسل ستون‌های خالی انتهایی را می‌اندازد؛ یک سلول برای حفظ عرض ۵۷ می‌گذاریم.
    if int(sheet.max_column or 0) < REQUIRED_COLUMNS:
        sheet.cell(row=1, column=REQUIRED_COLUMNS, value="")
    buffer = io.BytesIO()
    workbook.save(buffer)
    try:
        workbook.close()
    except Exception:
        pass
    log_event(f"کتاب کار داخلی از فایل ۰۱ ساخته شد: {len(filtered)} ردیف، {written} مقدار")
    return {
        "blob": buffer.getvalue(),
        "source_path": str(source_path),
        "source_sheet": source_sheet,
        "source_sheets": sheet_names,
        "header_row": header_row + 1,
        "sub_project_column": sub_col,
        "contract_column": contract_col,
        "total_rows": total,
        "filtered_rows": len(filtered),
        "updated_cells": written,
        "status_marks": status_marks,
        "sheet_name": sheet.title,
        "notes": notes,
    }


def _seed_workbook_blob(seed_path: str, source_path: str) -> bytes:
    """اگر کتاب کار معتبری از قبل روی دیسک بود، کتاب کار داخلی از آن شروع می‌شود."""
    if not seed_path:
        return b""
    try:
        seed = Path(str(seed_path)).expanduser()
        if not seed.exists() or not seed.is_file():
            return b""
        if seed.resolve() == managed_workbook_path().resolve():
            return b""
        if seed.resolve() == Path(str(source_path)).expanduser().resolve():
            return b""
        if not has_colorcode_sheet(str(seed)):
            return b""
        return seed.read_bytes()
    except Exception as exc:
        log_event(f"استفاده از کتاب کار موجود به‌عنوان پایه ناموفق بود: {exc}")
        return b""


def diagnose_system_file(source_path: str,
                         sheet_name: str = SHEET_NAME_DEFAULT) -> str:
    """گزارش خوانا از ساختار فایل ۰۱/کتاب کار: شیت، سرستون، ستون‌ها و علامت‌ها.

    برای وقتی که داشبورد صفر نشان می‌دهد: مشخص می‌کند ردیف‌ها فیلتر شده‌اند،
    ستون شمارهٔ قرارداد کدام است و در ستون‌های وضعیت چند علامت M/✓ هست.
    """
    lines: List[str] = [f"فایل: {source_path}"]
    try:
        names, sheet, rows = _open_rows(source_path)
    except ExcelDataError as exc:
        return "\n".join(lines + [f"خطا: {exc}"])
    header_row, sub_col, contract_col = _detect_header_row(rows)
    lines.append(f"شیت‌های فایل: {'، '.join(names)}")
    lines.append(f"شیت خوانده‌شده: «{sheet}» — ردیف‌ها: {fa(len(rows))}")
    lines.append(f"ردیف سرستون: {fa(header_row + 1)}")
    header = rows[header_row] if header_row < len(rows) else []
    titles = []
    for column, value in enumerate(header[:60]):
        text = txt(value)
        if text:
            titles.append(f"{_column_letter(column + 1)}={text}")
    lines.append("سرستون‌ها: " + (" | ".join(titles[:24]) or "—"))
    if sub_col >= 0:
        lines.append(f"ستون «Sub Project»: {_column_letter(sub_col + 1)}")
    else:
        lines.append("ستون «Sub Project» پیدا نشد (ستون F فرض می‌شود).")
    if contract_col >= 0:
        lines.append(f"ستون «شمارهٔ قرارداد»: {_column_letter(contract_col + 1)}")
    else:
        lines.append("ستون «شمارهٔ قرارداد» از سرستون پیدا نشد.")
    matched = 0
    for row in rows[header_row + 1:]:
        if sub_col < len(row) and _matches_sub_project(row[sub_col]):
            matched += 1
    lines.append(f"ردیف‌های «اجرا - مهر»: {fa(matched)}")

    # علامت‌های وضعیت: در ستون‌های L..AE (۱۲ تا ۳۱) و کل سطر
    mark_columns: Dict[int, int] = {}
    for row in rows[header_row + 1:]:
        for column in range(11, min(31, len(row))):
            if is_status_mark(row[column]):
                mark_columns[column] = mark_columns.get(column, 0) + 1
    if mark_columns:
        detail = "، ".join(
            f"{_column_letter(column + 1)}: {fa(count)}"
            for column, count in sorted(mark_columns.items())[:12]
        )
        lines.append(f"علامت‌های وضعیت (L تا AE): {detail}")
    else:
        lines.append("علامت وضعیت (M/✓/۱) در ستون‌های L تا AE پیدا نشد.")
    allowed = sum(
        1 for row in rows[header_row + 1:]
        if any(is_status_mark(value) for value in list(row)[11:60])
    )
    lines.append(f"ردیف‌هایی که در هر ستونی علامت دارند: {fa(allowed)}")

    sample_values: List[str] = []
    kept = 0
    for row in rows[header_row + 1:]:
        if sub_col < len(row) and _matches_sub_project(row[sub_col]):
            kept += 1
            if len(sample_values) < 2:
                cells = [
                    f"{_column_letter(index + 1)}={txt(row[index])}"
                    for index in range(min(len(row), 26))
                    if not is_blank(row[index])
                ]
                sample_values.append(" | ".join(cells[:16]))
    if sample_values:
        lines.append("نمونهٔ ردیف‌های «اجرا - مهر»:")
        lines.extend("  " + sample for sample in sample_values)
    lines.append("")
    lines.append("کتاب کار داخلی: " + str(managed_workbook_path()))
    if managed_workbook_path().exists():
        try:
            data = read_local_workbook(str(managed_workbook_path()), sheet_name=sheet_name)
            lines.append(f"رکوردهای کتاب کار داخلی: {fa(len(data.records))}")
            lines.append(
                "علامت‌های وضعیت در کتاب کار داخلی: "
                + fa(sum(sum(1 for value in rec.flags.values() if value)
                         for rec in data.records))
            )
        except Exception as exc:
            lines.append(f"خواندن کتاب کار داخلی ممکن نشد: {exc}")
    return "\n".join(lines)


def find_seed_workbook(source_path: str, extra_paths: Sequence[str] = ()) -> str:
    """کتاب کار «colorcode» قبلی کاربر را پیدا می‌کند تا پایهٔ ساخت کتاب کار داخلی شود.

    اگر کاربر از قبل فایل ۰۲ داشته (یا نسخهٔ به‌روزرسانی‌شده‌ای ساخته شده) همان
    ستون‌های وضعیت/کسوراتش حفظ می‌شود؛ وگرنه کتاب کار از صفر از فایل ۰۱ ساخته می‌شود.
    """
    candidates: List[str] = []

    managed = managed_workbook_path()

    def add(path: Any) -> None:
        text = str(path or "").strip()
        if text:
            candidates.append(text)

    def usable(probe: Path) -> bool:
        """کتاب کار واقعی کاربر: شیت colorcode دارد و ساختهٔ خود برنامه نیست."""
        try:
            if not probe.exists() or not probe.is_file():
                return False
            if probe.name.startswith("~$"):
                return False
            if probe.resolve() == managed.resolve():
                return False
            if probe.resolve() == Path(str(source_path)).expanduser().resolve():
                return False
            if not has_colorcode_sheet(str(probe)):
                return False
            if is_generated_workbook(str(probe)):
                return False
            return True
        except Exception:
            return False

    for item in extra_paths:
        add(item)
    add(ORIGINAL_EXCEL_PATH)
    try:
        seen = set()
        for path in candidates:
            key = str(Path(path))
            if key in seen:
                continue
            seen.add(key)
            try:
                probe = Path(path).expanduser()
                if usable(probe):
                    log_event(f"کتاب کار قبلی به‌عنوان پایه پیدا شد: {probe}")
                    return str(probe)
            except Exception:
                continue
        # جست‌وجوی آخرین کتاب کار معتبر در پوشهٔ داده‌ها و پوشهٔ فایل ۰۱
        folders = [data_dir()]
        try:
            folders.append(Path(str(source_path)).expanduser().parent)
        except Exception:
            pass
        found: List[Tuple[float, str]] = []
        for folder in folders:
            try:
                if not Path(folder).is_dir():
                    continue
                for item in Path(folder).glob("*.xlsx"):
                    try:
                        if usable(item):
                            found.append((item.stat().st_mtime, str(item)))
                    except Exception:
                        continue
            except Exception:
                continue
        if found:
            found.sort(reverse=True)
            log_event(f"کتاب کار قبلی (جدیدترین) به‌عنوان پایه: {found[0][1]}")
            return found[0][1]
    except Exception as exc:
        log_event(f"جست‌وجوی کتاب کار قبلی ناموفق بود: {exc}")
    return ""


def update_managed_workbook(source_path: str,
                            sheet_name: str = SHEET_NAME_DEFAULT,
                            seed_path: str = "") -> Dict[str, Any]:
    """کتاب کار برنامه را از فایل ۰۱ به‌روزرسانی می‌کند.

    قاعدهٔ اصلی: **کتاب کار اصلی کاربر (فایل ۰۲ / دیتابیس) هرگز با کتاب کار
    ساخته‌شده از فایل ۰۱ جایگزین نمی‌شود**؛ چون فایل ۰۱ ستون‌های وضعیت، کسورات و
    درصد پیشرفت را ندارد و اعداد داشبورد اشتباه می‌شوند.
    """
    target = managed_workbook_path()
    exists = target.exists() and target.stat().st_size > 0
    has_sheet = exists and has_colorcode_sheet(str(target))
    generated = has_sheet and is_generated_workbook(str(target))
    seed_blob = _seed_workbook_blob(seed_path, source_path)
    seed_used = ""

    def replace_with(blob: bytes, source_name: str) -> Tuple[bytes, int, SystemUpdatePlan]:
        plan_ = plan_system_update(source_path, blob, sheet_name=sheet_name)
        new_blob_, cells_ = apply_system_update(plan_, blob)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(new_blob_)
        log_event(
            f"کتاب کار داخلی از روی «{Path(source_name).name}» به‌روزرسانی شد: "
            f"{plan_.matched_rows} ردیف، {cells_} سلول"
        )
        return new_blob_, cells_, plan_

    if seed_blob:
        # کتاب کار اصلی موجود است.
        if not has_sheet or generated:
            if exists and generated:
                backup_workbook(target.read_bytes(), target.name)
            _new_blob, cells, plan = replace_with(seed_blob, seed_path)
            seed_used = str(seed_path)
            return {
                "created": True,
                "replaced_generated": bool(generated),
                "plan": plan,
                "built": {},
                "source_path": plan.source_path,
                "updated_cells": cells,
                "filtered_rows": plan.filtered_rows,
                "matched_rows": plan.matched_rows,
                "backup_path": "",
                "output_path": str(target),
                "seeded_from": seed_used,
                "summary": preview_system_update(plan),
            }
        blob = target.read_bytes()
        plan = plan_system_update(source_path, blob, sheet_name=sheet_name)
        backup = backup_workbook(blob, target.name)
        new_blob, cells = apply_system_update(plan, blob)
        target.write_bytes(new_blob)
        log_event(
            f"کتاب کار داخلی به‌روزرسانی شد: {plan.filtered_rows} ردیف «اجرا - مهر»، "
            f"{cells} سلول → {target}"
        )
        return {
            "created": False,
            "replaced_generated": False,
            "plan": plan,
            "built": {},
            "source_path": plan.source_path,
            "updated_cells": cells,
            "filtered_rows": plan.filtered_rows,
            "matched_rows": plan.matched_rows,
            "backup_path": backup,
            "output_path": str(target),
            "summary": preview_system_update(plan),
        }

    if generated:
        # کتاب کار ساخته‌شده از فایل ۰۱ و کتاب کار اصلی پیدا نشد؛ همان را درجا
        # به‌روزرسانی می‌کنیم و صادقانه هشدار می‌دهیم که اعداد وضعیت ناقص است.
        blob = target.read_bytes()
        plan = plan_system_update(source_path, blob, sheet_name=sheet_name)
        backup = backup_workbook(blob, target.name)
        new_blob, cells = apply_system_update(plan, blob)
        target.write_bytes(new_blob)
        log_event(
            f"کتاب کار ساخته‌شده (بدون کتاب کار اصلی) درجا به‌روزرسانی شد: "
            f"{plan.matched_rows} ردیف، {cells} سلول"
        )
        return {
            "created": False,
            "replaced_generated": True,
            "no_source_workbook": True,
            "plan": plan,
            "built": {},
            "source_path": plan.source_path,
            "updated_cells": cells,
            "filtered_rows": plan.filtered_rows,
            "matched_rows": plan.matched_rows,
            "backup_path": backup,
            "output_path": str(target),
            "summary": preview_system_update(plan),
        }

    if not has_sheet:
        built = build_workbook_from_system(source_path, sheet_name=sheet_name)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(built["blob"])
        lines = [
            "کتاب کار اصلی روی این سیستم پیدا نشد؛ کتاب کار از خود فایل ۰۱ ساخته شد "
            "(شیت colorcode، فیلتر «اجرا - مهر»، ستون‌های L تا AJ).",
            f"{fa(built['filtered_rows'])} قرارداد از {fa(built['total_rows'])} ردیف فایل سامانه.",
        ]
        if built["notes"]:
            lines.append("یادداشت: " + " ".join(built["notes"]))
        return {
            "created": True,
            "no_source_workbook": True,
            "plan": None,
            "built": built,
            "source_path": str(source_path),
            "updated_cells": built["updated_cells"],
            "filtered_rows": built["filtered_rows"],
            "matched_rows": built["filtered_rows"],
            "status_marks": built.get("status_marks", 0),
            "backup_path": "",
            "output_path": str(target),
            "summary": "\n".join(lines),
        }

    blob = target.read_bytes()
    plan = plan_system_update(source_path, blob, sheet_name=sheet_name)
    backup = backup_workbook(blob, target.name)
    new_blob, cells = apply_system_update(plan, blob)
    target.write_bytes(new_blob)
    return {
        "created": False,
        "replaced_generated": False,
        "plan": plan,
        "built": {},
        "source_path": plan.source_path,
        "updated_cells": cells,
        "filtered_rows": plan.filtered_rows,
        "matched_rows": plan.matched_rows,
        "backup_path": backup,
        "output_path": str(target),
        "summary": preview_system_update(plan),
    }


def backup_workbook(blob: bytes, filename: str = "workbook") -> str:
    """نسخهٔ پشتیبان کتاب کار فعلی را در پوشهٔ data نگه می‌دارد."""
    try:
        folder = data_dir() / "backups"
        folder.mkdir(parents=True, exist_ok=True)
        stem = Path(str(filename or "workbook")).stem
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        path = folder / f"{stem}-{stamp}.xlsx"
        path.write_bytes(blob)
        return str(path)
    except Exception as exc:
        log_event(f"ذخیرهٔ نسخهٔ پشتیبان ناموفق بود: {exc}")
        return ""


def _settings_search_paths(settings: Dict[str, Any]) -> List[str]:
    """مسیرهای ذخیره‌شدهٔ کتاب کار که ممکن است پایهٔ درست باشند."""
    keys = ("source_workbook_path", "local_excel_path", "last_excel_path",
            "extra_excel_path")
    return [str(settings.get(key) or "") for key in keys]


def list_database_workbooks(settings: Dict[str, Any]) -> List[Dict[str, Any]]:
    """همهٔ کتاب کارهای ذخیره‌شده در دیتابیس را با مشخصات کوتاه برمی‌گرداند."""
    connection = connect_database(settings)
    items: List[Dict[str, Any]] = []
    try:
        cursor = connection.cursor()
        try:
            cursor.execute(
                "SELECT workbook_id, original_filename, file_size, imported_at "
                "FROM excel_workbooks ORDER BY workbook_id"
            )
            rows = list(cursor.fetchall() or [])
            counts: Dict[int, int] = {}
            try:
                cursor.execute(
                    "SELECT workbook_id, COUNT(*) FROM mehr_contracts GROUP BY workbook_id"
                )
                counts = {int(r[0]): int(r[1]) for r in (cursor.fetchall() or [])}
            except Exception:
                counts = {}
        finally:
            cursor.close()
    finally:
        try:
            connection.close()
        except Exception:
            pass
    for row in rows:
        try:
            wid = int(row[0])
        except Exception:
            continue
        items.append({
            "workbook_id": wid,
            "filename": str(row[1] or f"workbook-{wid}.xlsx"),
            "file_size": int(row[2] or 0),
            "imported_at": str(row[3] or ""),
            "rows": int(counts.get(wid, 0)),
        })
    return items


def inspect_database_workbook(settings: Dict[str, Any], workbook_id: int) -> Dict[str, Any]:
    """کیفیت یک کتاب کار ذخیره‌شده در دیتابیس (شیت‌ها، درصد، علامت‌ها)."""
    target = dict(settings)
    target["workbook_id"] = int(workbook_id)
    try:
        info = fetch_workbook_blob(target)
    except Exception as exc:
        return {"workbook_id": int(workbook_id), "ok": False, "error": str(exc)}
    blob = info.get("data") or b""
    quality = inspect_workbook_bytes(blob)
    health = workbook_health(blob)
    rows = 0
    try:
        rows = int(load_rows_from_database({**target})["data_rows"])
    except Exception:
        rows = 0
    return {
        "workbook_id": int(workbook_id),
        "ok": True,
        "rows": rows,
        "filename": str(info.get("filename") or ""),
        "size": len(blob),
        "blob": blob,
        "colorcode": bool(quality.get("colorcode")),
        "generated": bool(quality.get("generated")),
        "has_progress": bool(quality.get("has_progress")),
        "mark_rows": int(quality.get("mark_rows") or 0),
        "damaged": bool(health.get("damaged")),
    }


_WORKBOOK_SCAN_CACHE: Dict[Any, Dict[str, Any]] = {}


def best_database_workbook(settings: Dict[str, Any],
                           skip_ids: Sequence[int] = ()) -> Optional[Dict[str, Any]]:
    """بهترین کتاب کار واقعی دیتابیس به‌عنوان پایه (مثل فایل ۰۲ با شناسهٔ ۳)."""
    try:
        items = list_database_workbooks(settings)
    except Exception as exc:
        log_event(f"فهرست کتاب کارهای دیتابیس خوانده نشد: {exc}")
        return None
    skip = {int(value) for value in skip_ids}
    fingerprint = (
        str(settings.get("host") or ""), int(settings.get("port") or 3306),
        str(settings.get("database") or ""),
        tuple(sorted((int(item["workbook_id"]), str(item.get("imported_at") or ""))
                     for item in items)),
        tuple(sorted(skip)),
    )
    cached = _WORKBOOK_SCAN_CACHE.get(fingerprint)
    if cached is not None:
        return cached or None
    # برای اینکه انتخاب خودکار کند نشود، فقط جدیدترین‌ها بررسی می‌شوند.
    items = sorted(
        items, key=lambda entry: (str(entry.get("imported_at") or ""),
                                  int(entry.get("workbook_id") or 0)),
        reverse=True,
    )[:8]
    reports: List[Tuple[float, Dict[str, Any]]] = []
    for item in items:
        wid = int(item["workbook_id"])
        if wid in skip:
            continue
        report = inspect_database_workbook(settings, wid)
        if not report.get("ok"):
            log_event(
                f"کتاب کار {wid} باز نشد: {str(report.get('error') or '')[:160]}"
            )
            continue
        score = 0.0
        if report["generated"]:
            score -= 1000.0
        else:
            score += 100.0
        if report["colorcode"]:
            score += 60.0
        else:
            score -= 200.0
        if report["has_progress"]:
            score += 40.0
        else:
            score -= 80.0
        if report["damaged"]:
            score -= 150.0
        counted = int(report.get("rows") or item.get("rows") or 0)
        score += min(300.0, counted * 0.25)
        score += min(60.0, report["mark_rows"] * 0.1)
        reports.append((score, {**item, **report, "score": score, "rows": counted}))
        log_event(
            "بررسی کتاب کار دیتابیس: شناسهٔ {wid} «{name}» — colorcode={cc}، "
            "درصد={pg}، ساختهٔ برنامه={gen}، ناقص={dmg}، امتیاز={score:.1f}".format(
                wid=fa(wid), name=report.get("filename") or item.get("filename"),
                cc="بله" if report["colorcode"] else "خیر",
                pg="بله" if report["has_progress"] else "خیر",
                gen="بله" if report["generated"] else "خیر",
                dmg="بله" if report["damaged"] else "خیر",
                score=score,
            )
        )
    if not reports:
        _WORKBOOK_SCAN_CACHE[fingerprint] = {}
        return None
    reports.sort(key=lambda pair: pair[0], reverse=True)
    best = reports[0][1]
    if best["score"] <= 0:
        _WORKBOOK_SCAN_CACHE[fingerprint] = {}
        return None
    _WORKBOOK_SCAN_CACHE[fingerprint] = best
    return best


def next_database_workbook_id(settings: Dict[str, Any]) -> int:
    """شناسهٔ آزاد بعدی برای ذخیرهٔ نسخهٔ به‌روز (کتاب کار پایه دست‌نخورده می‌ماند)."""
    try:
        items = list_database_workbooks(settings)
        used = [int(item["workbook_id"]) for item in items]
    except Exception:
        used = []
    return (max(used) + 1) if used else 1


def resolve_update_base(settings: Dict[str, Any],
                        source_path: str = "") -> Dict[str, Any]:
    """کتاب کار پایه برای اعمال فایل ۰۱ را برمی‌گرداند.

    ترتیب انتخاب:
    ۱) کتاب کار تنظیم‌شده در دیتابیس (مثلاً شناسهٔ ۳ = فایل ۰۲ خودتان)،
    ۲) اگر آن ساختهٔ برنامه یا ناقص/موجود نبود، بهترین کتاب کار واقعی دیتابیس،
    ۳) اگر دیتابیس در دسترس نبود، همان کتاب کار روی دیسک.
    """
    configured = int(settings.get("workbook_id") or 1)
    info = None
    try:
        info = fetch_workbook_blob(settings)
    except Exception as exc:
        log_event(f"خواندن کتاب کار دیتابیس برای پایهٔ به‌روزرسانی ممکن نشد: {exc}")
    if info:
        blob = info.get("data") or b""
        quality = inspect_workbook_bytes(blob)
        health = workbook_health(blob)
        usable = (
            not quality.get("generated")
            and bool(quality.get("colorcode"))
            and not health.get("damaged")
        )
        if usable:
            current = inspect_database_workbook(settings, configured)
            best_choice = best_database_workbook(settings, skip_ids=[configured])
            if (
                best_choice
                and float(best_choice.get("score") or 0)
                > float(current.get("score") or 0) + 20
            ):
                log_event(
                    f"کتاب کار کامل‌تر در دیتابیس پیدا شد: شناسهٔ "
                    f"{fa(best_choice['workbook_id'])} «{best_choice.get('filename')}» "
                    f"({best_choice.get('rows')} قرارداد) به‌جای شناسهٔ "
                    f"{fa(configured)} ({current.get('rows')} قرارداد)"
                )
                return {
                    "blob": best_choice["blob"],
                    "filename": str(best_choice.get("filename") or "workbook.xlsx"),
                    "base_source": "database",
                    "base_path": "",
                    "base_workbook_id": int(best_choice["workbook_id"]),
                }
            log_event(
                f"کتاب کار پایه از دیتابیس انتخاب شد: شناسهٔ {fa(configured)} "
                f"«{info.get('filename')}»"
            )
            return {
                "blob": blob,
                "filename": str(info.get("filename") or "workbook.xlsx"),
                "base_source": "database",
                "base_path": "",
                "base_workbook_id": configured,
            }
        log_event(
            "کتاب کار تنظیم‌شده برای پایه مناسب نیست "
            f"(ساختهٔ برنامه={quality.get('generated')}، "
            f"colorcode={quality.get('colorcode')}، ناقص={health.get('damaged')})؛ "
            "بهترین کتاب کار دیتابیس جست‌وجو می‌شود"
        )
    best = best_database_workbook(settings, skip_ids=[configured])
    if best:
        log_event(
            "کتاب کار پایه از دیتابیس انتخاب شد: شناسهٔ "
            f"{fa(best['workbook_id'])} «{best.get('filename') or best.get('filename')}»"
        )
        return {
            "blob": best["blob"],
            "filename": str(best.get("filename") or "workbook.xlsx"),
            "base_source": "database",
            "base_path": "",
            "base_workbook_id": int(best["workbook_id"]),
        }
    real = find_real_workbook(source_path, extra_paths=_settings_search_paths(settings))
    if real:
        try:
            blob = Path(real).expanduser().read_bytes()
        except OSError as exc:
            raise ExcelDataError(f"خواندن کتاب کار اصلی ممکن نشد:\n{exc}") from exc
        log_event(f"کتاب کار اصلی به‌عنوان پایهٔ به‌روزرسانی انتخاب شد: {real}")
        return {
            "blob": blob,
            "filename": Path(real).name,
            "base_source": "file",
            "base_path": str(real),
            "base_workbook_id": 0,
        }
    raise ExcelDataError(
        "کتاب کار اصلی (فایل دارای شیت colorcode که ستون‌های وضعیت و درصد را دارد) "
        "پیدا نشد.\n\n"
        "فایل ۰۱ سامانه فقط ستون‌های کارکرد را دارد؛ اگر آن را جای کتاب کار اصلی "
        "بگذاریم کارت‌های وضعیت و درصد صفر می‌شوند.\n"
        "کتاب کار ۰۲ را در دیتابیس با یک شناسه (مثلاً ۳) ذخیره کنید و همان شناسه را "
        "در تنظیمات بگذارید، یا دکمهٔ «انتخاب کتاب کار از دیتابیس» را بزنید."
    )


def update_database_from_system(settings: Dict[str, Any], source_path: str,
                                sheet_name: str = SHEET_NAME_DEFAULT) -> Dict[str, Any]:
    """فایل ۰۱ را می‌خواند، فیلتر می‌کند و کتاب کار دیتابیس را به‌روز می‌کند."""
    base = resolve_update_base(settings, source_path)
    blob = base["blob"]
    plan = plan_system_update(source_path, blob, sheet_name=sheet_name)
    backup = backup_workbook(blob, str(base.get("filename") or "workbook"))
    new_blob, cells = apply_system_update(plan, blob)
    with tempfile.TemporaryDirectory(prefix="mehr-update-") as temp:
        temp_path = Path(temp) / (str(base.get("filename") or "workbook.xlsx"))
        if temp_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            temp_path = temp_path.with_suffix(".xlsx")
        temp_path.write_bytes(new_blob)
        imported = import_workbook_to_database(
            settings, str(temp_path), sheet_name=plan.sheet_name,
            workbook_id=int(settings.get("workbook_id") or 1),
        )
    log_event(
        f"به‌روزرسانی از فایل سامانه: {plan.filtered_rows} ردیف «اجرا - مهر»، "
        f"{plan.matched_rows} تطبیق، {cells} سلول"
    )
    return {
        "plan": plan,
        "updated_cells": cells,
        "backup_path": backup,
        "imported": imported,
        "summary": preview_system_update(plan),
    }


def fetch_workbook_blob(settings: Dict[str, Any]) -> Dict[str, Any]:
    """دریافت فایل اصلی کتاب کار از دیتابیس (جدول excel_original_files)."""
    workbook_id = int(settings.get("workbook_id") or 1)
    connection = connect_database(settings)
    try:
        cursor = connection.cursor()
        try:
            cursor.execute(
                "SELECT original_filename, file_size, imported_at "
                "FROM excel_workbooks WHERE workbook_id=%s",
                (workbook_id,),
            )
            meta = cursor.fetchone()
            if not meta:
                raise DatabaseError(
                    f"کتاب کار با شناسهٔ {fa(workbook_id)} در دیتابیس پیدا نشد."
                )
            cursor.execute(
                "SELECT binary_data FROM excel_original_files WHERE workbook_id=%s",
                (workbook_id,),
            )
            row = cursor.fetchone()
            if not row or row[0] is None:
                raise DatabaseError(
                    "فایل اصلی کتاب کار در دیتابیس ذخیره نشده است "
                    "(جدول excel_original_files خالی است)."
                )
            blob = row[0]
            if isinstance(blob, memoryview):
                blob = blob.tobytes()
            elif not isinstance(blob, (bytes, bytearray)):
                blob = bytes(blob)
            log_event(
                f"کتاب کار {workbook_id}: {meta[0]} — {len(blob)} بایت در دیتابیس یافت شد"
            )
            return {
                "filename": str(meta[0] or f"workbook-{workbook_id}.xlsx"),
                "size": int(meta[1] or len(blob)),
                "imported_at": meta[2],
                "data": bytes(blob),
            }
        finally:
            cursor.close()
    finally:
        try:
            connection.close()
        except Exception:
            pass


def fetch_workbook_updated_at(settings: Dict[str, Any]) -> Optional[datetime]:
    """زمان آخرین بارگذاری کتاب کار (برای بازخوانی خودکار)."""
    connection = connect_database(settings)
    try:
        cursor = connection.cursor()
        try:
            cursor.execute(
                "SELECT imported_at FROM excel_workbooks WHERE workbook_id=%s",
                (int(settings.get("workbook_id") or 1),),
            )
            row = cursor.fetchone()
        finally:
            cursor.close()
    finally:
        try:
            connection.close()
        except Exception:
            pass
    if row and isinstance(row[0], datetime):
        return row[0]
    return None


def _parse_imported_at(value: Any) -> Optional[datetime]:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    text = str(value or "").strip()
    for pattern in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text[:19], pattern)
        except Exception:
            continue
    return None


def migrate_blob_to_tables(settings: Dict[str, Any], info: Dict[str, Any]) -> None:
    """فایل کتاب کار ذخیره‌شده را به جدول‌های نرمال‌شده منتقل می‌کند."""
    workbook_id = int(settings.get("workbook_id") or 1)
    sheet_name = str(settings.get("sheet_name") or SHEET_NAME_DEFAULT)
    connection = connect_database(settings)
    try:
        cursor = connection.cursor()
        try:
            ensure_mehr_schema(cursor)
            cursor.execute(
                "SELECT COUNT(*) FROM mehr_contracts WHERE workbook_id=%s", (workbook_id,)
            )
            row = cursor.fetchone()
            if row and int(row[0] or 0) > 0:
                return
            blob = info.get("data") or b""
            book = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
            try:
                target = sheet_name if sheet_name in book.sheetnames else book.sheetnames[0]
                raw_rows = [list(r) for r in book[target].iter_rows(values_only=True)]
                sheet_count = len(book.sheetnames)
            finally:
                book.close()
            kind = "generated" if inspect_workbook_bytes(blob).get("generated") else "excel"
            store_rows_in_database(
                cursor, workbook_id, str(info.get("filename") or "workbook.xlsx"),
                target, sheet_count, raw_rows, source_kind=kind,
                source_sha=str(info.get("sha256") or ""),
            )
            connection.commit()
        finally:
            cursor.close()
    finally:
        try:
            connection.close()
        except Exception:
            pass


LAST_WORKBOOK_CHOICE: Dict[str, int] = {}


def _read_database_workbook_once(settings: Dict[str, Any],
                                 pdf_folder: str = "") -> DashboardData:
    """یک بار خواندن از دیتابیس با شناسهٔ داده‌شده (بدون انتخاب خودکار).

    اول از جدول‌های نرمال‌شده (mehr_contracts) خوانده می‌شود؛ اگر خالی بودند،
    فایل ذخیره‌شده (blob) خوانده و همان‌جا به جدول‌ها منتقل می‌شود.
    """
    log_event(
        f"خواندن داده از دیتابیس {settings.get('database')} "
        f"(workbook_id={settings.get('workbook_id')})"
    )
    workbook_id = int(settings.get("workbook_id") or 1)
    note = "دیتابیس {} • شناسهٔ {}".format(
        settings.get("database") or "master_data", fa(workbook_id)
    )
    try:
        stored = load_rows_from_database(settings)
    except DatabaseError:
        stored = None
    except Exception as exc:
        log_event(f"خواندن جدول‌های نرمال‌شده ناموفق بود: {type(exc).__name__}: {exc}")
        stored = None
    if stored and stored["raw_rows"]:
        log_event(
            f"داشبورد از جدول‌های دیتابیس ساخته می‌شود: "
            f"{stored['data_rows']} سطر / شیت {stored['sheet_name']}"
        )
        data = parse_workbook_rows(
            stored["raw_rows"],
            workbook_name=stored["filename"] or "کتاب کار دیتابیس",
            source_sheet=str(settings.get("sheet_name") or stored["sheet_name"]),
            pdf_folder=pdf_folder,
            source_time=datetime.now(),
            source_size=0,
            source_note=note + " • جدول contracts",
            source_kind=str(stored.get("source_kind") or "excel"),
        )
        _flag_damaged_rows(settings, data)
        return data
    info = fetch_workbook_blob(settings)
    note = "دیتابیس {} • شناسهٔ {}".format(
        settings.get("database") or "master_data", fa(int(settings.get("workbook_id") or 1))
    )
    quality = inspect_workbook_bytes(info["data"])
    health = workbook_health(info["data"])
    data = parse_workbook_bytes(
        info["filename"],
        info["data"],
        sheet_name=str(settings.get("sheet_name") or SHEET_NAME_DEFAULT),
        pdf_folder=pdf_folder,
        source_time=_parse_imported_at(info.get("imported_at")),
        source_size=int(info.get("size") or len(info["data"])),
        source_note=note,
        source_kind="generated" if quality.get("generated") else "excel",
    )
    if health.get("damaged"):
        data.damaged = True
        data.damaged_reason = damaged_workbook_message(health)
        log_event(
            "کتاب کار ذخیره‌شده فرمول‌هایی بدون مقدار دارد: "
            f"{health.get('missing')} از {health.get('formulas')}"
        )
    if not data.damaged:
        try:
            migrate_blob_to_tables(settings, info)
            log_event("کتاب کار ذخیره‌شده به جدول‌های نرمال‌شده منتقل شد")
        except Exception as exc:
            log_event(f"انتقال کتاب کار به جدول‌ها ناموفق بود: {type(exc).__name__}: {exc}")
    return data


def _flag_damaged_rows(settings: Dict[str, Any], data: DashboardData) -> None:
    """اگر کتاب کار ذخیره‌شده ناقص باشد، داده را علامت می‌زند (فقط در صورت نیاز)."""
    if data.progress_percent_text != "—" or any(data.flag_counts.values()):
        return
    try:
        info = fetch_workbook_blob(settings)
    except Exception:
        return
    health = workbook_health(info.get("data") or b"")
    if health.get("damaged"):
        data.damaged = True
        data.damaged_reason = damaged_workbook_message(health)
        log_event(
            "کتاب کار ذخیره‌شده فرمول‌هایی بدون مقدار دارد: "
            f"{health.get('missing')} از {health.get('formulas')}"
        )


def read_database_workbook(settings: Dict[str, Any], pdf_folder: str = "") -> DashboardData:
    """خواندن داشبورد از دیتابیس؛ اگر شناسهٔ تنظیم‌شده مناسب نبود، بهترین را برمی‌دارد.

    مهم: اگر کتاب کار ۰۲ با شناسهٔ دیگری (مثلاً ۳) ذخیره شده باشد و شناسهٔ
    تنظیمات شناسهٔ دیگری باشد، برنامه خودش کتاب کار واقعی/کامل را پیدا و
    استفاده می‌کند و شناسهٔ انتخابی را در LAST_WORKBOOK_CHOICE می‌گذارد.
    """
    configured = int(settings.get("workbook_id") or 1)
    data = _read_database_workbook_once(settings, pdf_folder)
    if data.total_contracts and not data.damaged:
        # اگر شناسهٔ فعلی درست است ولی کتاب کار کامل‌تری (مثلاً ۰۲ با شناسهٔ ۳)
        # در دیتابیس هست، همان کامل‌تر استفاده می‌شود.
        try:
            best_now = best_database_workbook(settings, skip_ids=[configured])
        except Exception:
            best_now = None
        if best_now and int(best_now.get("rows") or 0) >= int(
            data.total_contracts * 1.25 + 1
        ):
            alternative = dict(settings)
            alternative["workbook_id"] = int(best_now["workbook_id"])
            try:
                other = _read_database_workbook_once(alternative, pdf_folder)
            except Exception:
                other = None
            if other and other.total_contracts > data.total_contracts:
                log_event(
                    f"کتاب کار کامل‌تر انتخاب شد: شناسهٔ {fa(best_now['workbook_id'])} "
                    f"({other.total_contracts} قرارداد) به‌جای شناسهٔ {fa(configured)} "
                    f"({data.total_contracts} قرارداد)"
                )
                LAST_WORKBOOK_CHOICE.clear()
                LAST_WORKBOOK_CHOICE.update({
                    "workbook_id": int(best_now["workbook_id"]),
                    "switched": 1,
                    "previous": configured,
                })
                other.source_note = (other.source_note or "") + " • انتخاب خودکار"
                return other
        LAST_WORKBOOK_CHOICE.clear()
        LAST_WORKBOOK_CHOICE.update({"workbook_id": configured, "switched": 0})
        return data
    try:
        best = best_database_workbook(settings, skip_ids=[configured])
    except Exception as exc:
        log_event(f"انتخاب خودکار کتاب کار ممکن نشد: {exc}")
        best = None
    if best:
        alternative = dict(settings)
        alternative["workbook_id"] = int(best["workbook_id"])
        try:
            other = _read_database_workbook_once(alternative, pdf_folder)
        except Exception as exc:
            log_event(f"خواندن کتاب کار جایگزین ناموفق بود: {exc}")
            other = None
        if other and (
            other.total_contracts > data.total_contracts
            or (data.damaged and not other.damaged)
            or (data.damaged and other.total_contracts >= data.total_contracts)
        ):
            log_event(
                "کتاب کار جایگزین انتخاب شد: شناسهٔ "
                f"{fa(best['workbook_id'])} «{best.get('filename')}» "
                f"({other.total_contracts} قرارداد) به‌جای شناسهٔ {fa(configured)} "
                f"({data.total_contracts} قرارداد)"
            )
            LAST_WORKBOOK_CHOICE.clear()
            LAST_WORKBOOK_CHOICE.update({
                "workbook_id": int(best["workbook_id"]),
                "switched": 1,
                "previous": configured,
            })
            other.source_note = (other.source_note or "") + " • انتخاب خودکار"
            return other
    LAST_WORKBOOK_CHOICE.clear()
    LAST_WORKBOOK_CHOICE.update({"workbook_id": configured, "switched": 0})
    return data


def read_local_workbook(path: str, sheet_name: str = SHEET_NAME_DEFAULT,
                        pdf_folder: str = "") -> DashboardData:
    """خواندن مستقیم فایل اکسل از دیسک (حالت بدون دیتابیس)."""
    file_path = Path(str(path)).expanduser()
    log_event(f"خواندن فایل اکسل محلی: {file_path}")
    if not file_path.exists() or not file_path.is_file():
        raise ExcelDataError(f"فایل اکسل پیدا نشد:\n{file_path}")
    try:
        blob = file_path.read_bytes()
        stamp = datetime.fromtimestamp(file_path.stat().st_mtime)
    except OSError as exc:
        raise ExcelDataError(f"خواندن فایل اکسل ممکن نشد:\n{exc}") from exc
    note = f"فایل محلی: {file_path.name}"
    quality = inspect_workbook_bytes(blob)
    health = workbook_health(blob)
    data = parse_workbook_bytes(
        file_path.name,
        blob,
        sheet_name=str(sheet_name or SHEET_NAME_DEFAULT),
        pdf_folder=pdf_folder,
        source_time=stamp,
        source_size=len(blob),
        source_note=note,
        source_kind="generated" if quality.get("generated") else "excel",
    )
    if health.get("damaged"):
        data.damaged = True
        data.damaged_reason = damaged_workbook_message(health)
        if data.repaired:
            data.damaged_reason += (
                f"\n\n{fa(data.repaired)} ردیف از شیت سامانهٔ همین فایل بازسازی شد، "
                "اما وضعیت‌های علامت‌دار و درصد کل را فقط نسخهٔ سالم کتاب کار دارد."
            )
    return data


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
    source_name: str
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
    source_note: str = ""
    source_kind: str = "excel"
    pdf_total: int = 0
    metric_note: str = ""
    damaged: bool = False
    damaged_reason: str = ""
    repaired: int = 0

    @property
    def excel_path(self) -> str:
        """نام کتاب کار (سازگاری با کدهای قبلی که این فیلد را می‌خوانند)."""
        return self.source_name

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


def _progress_bucket(percent: float) -> str:
    """درصد پیشرفت را به یک دستهٔ خوانا برای نمودار تبدیل می‌کند."""
    if percent >= 99.5:
        return "۱۰۰٪ کامل"
    if percent >= 75:
        return "۷۵ تا ۱۰۰٪"
    if percent >= 50:
        return "۵۰ تا ۷۵٪"
    if percent >= 25:
        return "۲۵ تا ۵۰٪"
    if percent > 0:
        return "زیر ۲۵٪"
    return "شروع‌نشده"


def _progress_like(value: Any) -> str:
    """سلول درصدی: اگر عدد باشد به درصد تبدیل می‌شود (مثل نسخهٔ وب)."""
    text = txt(value)
    if not text:
        return ""
    number = num(text)
    if number is not None:
        return fa(f"{number * 100:.2f}") + "٪"
    return text.replace("%", "٪")


def parse_workbook_bytes(
    filename: str,
    blob: bytes,
    sheet_name: str = SHEET_NAME_DEFAULT,
    pdf_folder: str = "",
    source_time: Optional[datetime] = None,
    source_size: int = 0,
    source_note: str = "",
    source_kind: str = "",
) -> DashboardData:
    """ساخت کامل داده‌های داشبورد از بایت‌های کتاب کار (خوانده‌شده از دیتابیس)."""
    if openpyxl is None:
        raise ExcelDataError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    if not blob:
        raise ExcelDataError("محتوای کتاب کار خالی است.")
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
    except Exception as exc:
        raise ExcelDataError(
            f"کتاب کار ذخیره‌شده در دیتابیس باز نشد:\n{exc}"
        ) from exc
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
        repair = {"filled": 0, "progress": 0, "sheet": "", "total": None}
        if raw_rows and any(
            is_blank(raw_rows[index][column])
            for index in range(3, min(len(raw_rows), 60))
            for column in (3, 4)
            if column < len(raw_rows[index])
        ):
            try:
                repair = repair_rows_from_companion(workbook, target, raw_rows)
            except Exception as exc:
                log_event(f"بازسازی ستون‌های خالی از شیت سامانه ممکن نشد: {exc}")
        if repair.get("filled"):
            log_event(
                "ستون‌های خالی از شیت «{}» بازسازی شد: {} ردیف، {} درصد".format(
                    repair.get("sheet"), repair.get("filled"), repair.get("progress")
                )
            )
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    if repair.get("filled"):
        source_note = (source_note + " • " if source_note else "") + (
            f"ستون‌های شناسایی/درصد از شیت «{repair.get('sheet')}» همان فایل بازسازی شد"
        )
        source_kind = source_kind or "excel"
    kind = source_kind or (
        "generated" if inspect_workbook_bytes(blob).get("generated") else "excel"
    )
    data = parse_workbook_rows(
        raw_rows,
        workbook_name=str(filename or "کتاب کار"),
        source_sheet=target,
        pdf_folder=pdf_folder,
        source_time=source_time,
        source_size=source_size or len(blob),
        source_note=source_note,
        source_kind=kind,
    )
    data.repaired = int(repair.get("filled") or 0)
    return data


def parse_workbook_rows(
    raw_rows: List[List[Any]],
    workbook_name: str = "کتاب کار",
    source_sheet: str = SHEET_NAME_DEFAULT,
    pdf_folder: str = "",
    source_time: Optional[datetime] = None,
    source_size: int = 0,
    source_note: str = "",
    source_kind: str = "",
) -> DashboardData:
    """همان ساخت داشبورد، اما از ردیف‌های آمادهٔ شیت (فایل یا دیتابیس)."""
    workbook_name = str(workbook_name or "کتاب کار")
    target = source_sheet or SHEET_NAME_DEFAULT
    warnings: List[str] = []
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

        marked_offsets = [
            offset for offset in range(23)
            if 11 + offset < len(row) and is_status_mark(row[11 + offset])
        ]
        deductions = " | ".join(
            deduction_headers[offset]
            for offset in marked_offsets
            if offset < len(deduction_headers) and deduction_headers[offset]
        )

        flags = {
            key: is_status_mark(row[column] if column < len(row) else None)
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
        for key in pdf_alias_keys(record.col3):
            filename = pdf_index.get(key)
            if filename:
                record.pdf_file = filename
                matched_pdfs += 1
                break

    # ---- شمارش اسناد تکمیل‌شده (اطلاعات واقعی، بدون دست‌کاری شاخص‌ها) ----
    # اعداد داشبورد فقط از خود کتاب کار خوانده می‌شوند؛ هیچ درصدي از روی علامت‌ها
    # ساخته نمی‌شود تا با فایل اصلی کاربر اختلاف پیدا نکند.
    metric_note = ""
    flags_done_counts: List[int] = [0] * (len(FLAG_DEFS) + 1)
    tracked_keys = [key for key, _label, _column in FLAG_DEFS]
    total_flags = max(1, len(tracked_keys))
    for record in records:
        done = sum(1 for key in tracked_keys if record.flags.get(key))
        flags_done_counts[min(done, total_flags)] += 1

    if records and not any(any(record.flags.values()) for record in records):
        metric_note = (
            "در داده‌های فعلی هیچ علامت وضعیت (M) نیست؛ کارت‌های وضعیت صفر می‌مانند."
        )
        log_event("هیچ علامت وضعیت M در داده‌ها پیدا نشد")

    flags_done_labels = [f"{fa(index)} سند" for index in range(total_flags + 1)]

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

    source_mtime = source_time

    if not records:
        warnings.append("هیچ ردیف داده‌ای در شیت پیدا نشد.")
    elif not any(any(record.flags.values()) for record in records):
        warnings.append(
            "در ستون‌های وضعیت (L تا AE) هیچ علامت «M/✓» پیدا نشد؛ کارت‌های وضعیت خالی "
            "می‌مانند و درصد پیشرفت نمایش داده نمی‌شود.\n"
            "• اگر وضعیت‌ها در ستون‌های دیگری از فایل ۰۱ هستند، بگویید تا همان‌ها خوانده شود.\n"
            "• اگر وضعیت‌ها را دستی در شیت colorcode می‌زنید، همان فایل ذخیره‌شده را یک بار "
            "در پوشهٔ برنامه بگذارید تا برنامه از آن به‌عنوان پایه استفاده کند."
        )

    header_labels = [
        "ردیف", "نام شرکت", "شماره قرارداد", "موضوع قرارداد",
        "تعداد موارد باز", "کسورات اسناد اختتام", "نوع قرارداد",
        "پیشرفت کارکرد", "وضعیت پکیج آزادسازی", "الحاقیه",
    ]
    log_event(
        f"پارس شد: {workbook_name} / شیت {target} — "
        f"{len(records)} قرارداد، پیشرفت {progress_percent_text}٪"
    )

    return DashboardData(
        source_name=workbook_name,
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
            "flags_done": Distribution(flags_done_labels, flags_done_counts),
        },
        top_types=top_types,
        progress_headers=progress_headers,
        header_labels=header_labels,
        loaded_at=datetime.now(),
        source_mtime=source_mtime,
        source_size=source_size,
        warnings=warnings,
        source_note=source_note,
        source_kind=str(source_kind or "excel"),
        pdf_total=pdf_file_count(pdf_folder),
        metric_note=metric_note,
    )


_PDF_INDEX_CACHE: Dict[str, Any] = {"folder": None, "stamp": None, "time": 0.0, "index": {}}

_PDF_NOISE_WORDS = {
    "pdf", "copy", "final", "signed", "scan", "scanned", "new", "old", "doc", "document",
    "نسخه", "نسخهٔ", "کپی", "امضا", "امضاشده", "امضا شده", "نهایی", "اسکن", "اسکن‌شده",
    "پیوست", "ضمیمه", "قرارداد", "تایید", "تأیید", "ارسال", "دریافت", "شماره", "نمبر",
    "دانلود", "ویرایش", "اصل", "کامل", "بروزرسانی", "به‌روزرسانی", "پاک", "پاک‌شده",
}


def _pdf_noise_token(token: str) -> bool:
    cleaned = normalize_per(str(token)).strip(" .-_()[]{}")
    if not cleaned:
        return True
    if cleaned.isdigit():
        return False
    return cleaned in _PDF_NOISE_WORDS


def scan_pdf_files(folder: str, depth: int = 2, limit: int = 20000) -> List[str]:
    """فهرست مسیرهای نسبی PDFهای پوشه و زیرپوشه‌ها (تا دو سطح)."""
    results: List[str] = []
    if not folder or not os.path.isdir(folder):
        return results
    root = os.path.abspath(folder)
    stack: List[Tuple[str, int]] = [(root, 0)]
    while stack and len(results) < limit:
        current, level = stack.pop()
        try:
            entries = list(os.scandir(current))
        except OSError:
            continue
        for entry in entries:
            try:
                if entry.is_dir(follow_symlinks=False):
                    if level + 1 <= depth and not entry.name.startswith("."):
                        stack.append((entry.path, level + 1))
                    continue
                if not entry.is_file() or not entry.name.lower().endswith(".pdf"):
                    continue
            except OSError:
                continue
            results.append(os.path.relpath(entry.path, root))
            if len(results) >= limit:
                break
    results.sort(key=normalize_per)
    return results


def pdf_alias_keys(value: Any) -> List[str]:
    """کلیدهای ممکن برای تطبیق نام فایل PDF با شمارهٔ قرارداد."""
    if value is None:
        return []
    raw = str(value).translate(_LATIN_DIGITS)
    raw = re.sub(r"\(\s*\d+\s*\)", " ", raw)
    raw = re.sub(r"\[\s*\d+\s*\]", " ", raw)
    keys: List[str] = []

    def add(key: str) -> None:
        if key and key not in keys:
            keys.append(key)

    add(normalize_lookup_key(raw))

    tokens = [t for t in re.split(r"[\s_\-./\\()\[\]{}•·,،;:]+", raw) if t]
    while tokens and _pdf_noise_token(tokens[0]):
        tokens.pop(0)
    while tokens and _pdf_noise_token(tokens[-1]):
        tokens.pop()
    tokens = [t for t in tokens if not _pdf_noise_token(t)]
    if tokens:
        add(normalize_lookup_key(" ".join(tokens)))

    groups = re.findall(r"\d+", raw)
    if groups:
        add("".join(groups))
        add("".join((group.lstrip("0") or "0") for group in groups))
        if len(groups) > 1:
            add("".join(reversed(groups)))
            add("".join((group.lstrip("0") or "0") for group in reversed(groups)))
    return keys


def build_pdf_index(folder: str, limit: int = 20000) -> Dict[str, str]:
    """ایندکس PDFها: کلید نرمال‌شده (و کلیدهای جایگزین) → مسیر نسبی فایل."""
    if not folder or not os.path.isdir(folder):
        return {}
    try:
        stamp = os.path.getmtime(folder)
    except OSError:
        stamp = None
    cached = _PDF_INDEX_CACHE
    now = time.time()
    if (cached.get("folder") == folder and cached.get("stamp") == stamp
            and now - float(cached.get("time") or 0.0) < 20
            and cached.get("index")):
        return dict(cached["index"])  # type: ignore[arg-type]

    index: Dict[str, str] = {}
    ambiguous = set()
    for rel in scan_pdf_files(folder, limit=limit):
        stem = os.path.splitext(os.path.basename(rel))[0]
        for key in pdf_alias_keys(stem):
            if key in index and index[key] != rel:
                ambiguous.add(key)
            else:
                index[key] = rel
    for key in ambiguous:
        index.pop(key, None)
    _PDF_INDEX_CACHE.update(
        {"folder": folder, "stamp": stamp, "time": now, "index": dict(index)}
    )
    return index


def pdf_file_count(folder: str) -> int:
    """تعداد فایل‌های PDF موجود در پوشه (و زیرپوشه‌ها)."""
    return len(scan_pdf_files(folder))


def empty_dashboard(message: str = "") -> DashboardData:
    """دادهٔ خالی برای زمانی که هنوز اکسل بارگذاری نشده است."""
    return DashboardData(
        source_name="",
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
            "flags_done": Distribution([], []),
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



# ==========================================================================
# سیستم طراحی (رنگ، قلم و استایل)
# ==========================================================================
# ---------------------------------------------------------------------------
# سیستم طراحی (رنگ‌ها، فاصله‌ها، قلم و استایل)
# ---------------------------------------------------------------------------
THEMES: Dict[str, Dict[str, str]] = {
    "light": {
        "name": "روشن",
        "bg": "#F2F3FA",
        "bg_2": "#E9EBF5",
        "surface": "#FFFFFF",
        "surface_2": "#F8F9FE",
        "surface_3": "#EFF1FA",
        "surface_inv": "#141830",
        "border": "#E3E6F3",
        "border_soft": "#EEF0F9",
        "ink": "#12162B",
        "ink_2": "#464E6B",
        "ink_3": "#7B84A3",
        "brand": "#5A4FE0",
        "brand_2": "#7A6CF6",
        "brand_deep": "#3E34B8",
        "brand_soft": "#ECEAFF",
        "brand_tint": "#F6F5FF",
        "on_brand": "#FFFFFF",
        "accent": "#E3A33C",
        "accent_2": "#F3C570",
        "accent_soft": "#FCF3E2",
        "danger": "#D9484F",
        "danger_soft": "#FDEBEC",
        "warning": "#B2700F",
        "warning_soft": "#FCF2DF",
        "success": "#0F9D76",
        "success_soft": "#E4F6EF",
        "info": "#3B72D9",
        "info_soft": "#EAF1FD",
        "selection": "#E9E7FF",
        "row_alt": "#FAFBFF",
        "chart_grid": "#E8EAF5",
        "shadow": "#1A1B3A",
        "rail": "#0D1024",
        "rail_2": "#141833",
        "rail_ink": "#E9EBF9",
        "rail_muted": "#8B93B5",
        "rail_active": "#232A4E",
        "scrim": "rgba(10, 12, 28, 0.52)",
        "hero_1": "#241E6E",
        "hero_2": "#4B3FCB",
        "hero_3": "#7F5CE8",
        "hero_ink": "#F5F3FF",
        "hero_muted": "#CDC8FA",
        "paper": "#FFFFFF",
    },
    "dark": {
        "name": "تیره",
        "bg": "#07080F",
        "bg_2": "#0C0E19",
        "surface": "#111321",
        "surface_2": "#161A2A",
        "surface_3": "#1D2234",
        "surface_inv": "#EFF1FB",
        "border": "#262B42",
        "border_soft": "#1C2136",
        "ink": "#EDEFFA",
        "ink_2": "#B3BAD3",
        "ink_3": "#7F88A6",
        "brand": "#8C86FF",
        "brand_2": "#A79FFF",
        "brand_deep": "#C6C1FF",
        "brand_soft": "#1F1C44",
        "brand_tint": "#17163A",
        "on_brand": "#0C0A22",
        "accent": "#E9B45B",
        "accent_2": "#F4CC86",
        "accent_soft": "#33270F",
        "danger": "#F0797F",
        "danger_soft": "#3A1D22",
        "warning": "#EBB55F",
        "warning_soft": "#392C16",
        "success": "#45D3A8",
        "success_soft": "#12322A",
        "info": "#82B4F5",
        "info_soft": "#152840",
        "selection": "#24244A",
        "row_alt": "#131627",
        "chart_grid": "#232840",
        "shadow": "#000000",
        "rail": "#05060D",
        "rail_2": "#0A0D1B",
        "rail_ink": "#E5E8F8",
        "rail_muted": "#7B84A6",
        "rail_active": "#1B2138",
        "scrim": "rgba(0, 0, 0, 0.62)",
        "hero_1": "#14103F",
        "hero_2": "#2A2280",
        "hero_3": "#5A46C9",
        "hero_ink": "#F3F1FF",
        "hero_muted": "#BDB6F2",
        "paper": "#0F1120",
    },
}

# کلیدهای سازگاری با کدهای قدیمی (مدل جدول، دیالوگ‌ها و…)
_LEGACY_ALIASES = {
    "text": "ink",
    "text_muted": "ink_3",
    "primary": "brand",
    "primary_dark": "brand_deep",
    "primary_soft": "brand_soft",
    "surface_alt": "surface_2",
    "surface_soft": "surface_3",
    "sidebar": "rail",
    "sidebar_alt": "rail_2",
    "sidebar_text": "rail_ink",
    "sidebar_muted": "rail_muted",
    "header": "surface",
}
for _theme_colors in THEMES.values():
    for _legacy_key, _fresh_key in _LEGACY_ALIASES.items():
        _theme_colors.setdefault(_legacy_key, _theme_colors[_fresh_key])

RADIUS = {"card": 20, "inner": 14, "chip": 10, "button": 13}
SPACE = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 22, "xxl": 30}

FONT_CANDIDATES = [
    "Vazirmatn",
    "Vazirmatn UI",
    "IRANSansX",
    "IRANSans",
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
    loaded = set(_LOADED_FONTS)
    for name in FONT_CANDIDATES:
        if name in loaded or name in available:
            return name
    fallback = QApplication.font().family() if QApplication.instance() else "Sans Serif"
    return fallback or "Sans Serif"


def font_family() -> str:
    return pick_font_family()


def app_font(size: int = 10, weight: int = QFont.Normal) -> QFont:
    font = QFont(font_family())
    font.setPointSize(size)
    font.setWeight(weight)
    font.setStyleStrategy(QFont.PreferAntialias)
    font.setHintingPreference(QFont.PreferFullHinting)
    return font


def set_text_aligned(label, text: str, side: str = "right") -> None:
    """متن و تراز را با هم به‌روزرسانی می‌کند (تراز راست واقعی در RTL)."""
    label.setText(text)
    if side == "center":
        label.setAlignment(Qt.AlignCenter)
    elif side == "left":
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter | Qt.AlignAbsolute)
    else:
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter | Qt.AlignAbsolute)


def mix(color_a: str, color_b: str, ratio: float) -> QColor:
    """ترکیب دو رنگ (ratio=0 یعنی رنگ اول، ratio=1 یعنی رنگ دوم)."""
    a, b = QColor(color_a), QColor(color_b)
    ratio = max(0.0, min(1.0, ratio))
    return QColor(
        int(a.red() + (b.red() - a.red()) * ratio),
        int(a.green() + (b.green() - a.green()) * ratio),
        int(a.blue() + (b.blue() - a.blue()) * ratio),
    )


def alpha(color: str, value: int) -> QColor:
    c = QColor(color)
    c.setAlpha(max(0, min(255, value)))
    return c




_RANGE_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*[٪%]?\s*[-–—]\s*(\d+(?:\.\d+)?)\s*[٪%]?\s*$")
_LIMIT_RE = re.compile(r"^\s*(کمتر|بیشتر)\s*از\s*[٪%]?\s*(\d+(?:\.\d+)?)\s*[٪%]?\s*$")
_SINGLE_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*[٪%]\s*$")


def pretty_label(text: Any) -> str:
    """نمایش خوانا‌تر برچسب‌های درصدی/بازه‌ای فارسی (فقط برای نمایش، نه فیلتر)."""
    value = str(text or "").strip()
    if not value:
        return value
    value = value.replace("ي", "ی").replace("ك", "ک").replace("٬", "")
    normalized = value.translate(_LATIN_DIGITS_FOR_LABELS)
    match = _RANGE_RE.match(normalized)
    if match:
        first, second = float(match.group(1)), float(match.group(2))
        low, high = sorted((first, second))
        return f"{fa(fmt_number(low))} تا {fa(fmt_number(high))} درصد"
    match = _LIMIT_RE.match(normalized)
    if match:
        word = "کمتر از" if match.group(1) == "کمتر" else "بیشتر از"
        return f"{word} {fa(fmt_number(float(match.group(2))))} درصد"
    match = _SINGLE_RE.match(normalized)
    if match:
        return f"{fa(fmt_number(float(match.group(1))))} درصد"
    return value


def fmt_number(value: float) -> str:
    """عدد اعشاری بدون صفرهای انتهایی (۶۴.۱۲ / ۱۰۰ / ۷.۵)."""
    text = f"{value:.2f}".rstrip("0").rstrip(".")
    return text or "0"


_LATIN_DIGITS_FOR_LABELS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def build_stylesheet(c: Dict[str, str], base_size: int = 10) -> str:
    family = font_family()
    primary_gradient = (
        f"qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {c['brand_2']}, stop:1 {c['brand']})"
    )
    return f"""
    * {{ font-family: '{family}'; }}
    QWidget {{
        color: {c['ink']};
        font-size: {base_size}pt;
        background: transparent;
    }}
    QMainWindow, QDialog {{ background: {c['bg']}; }}
    QToolTip {{
        background: {c['surface_inv']};
        color: {c['surface']};
        border: none;
        padding: 6px 10px;
        border-radius: 8px;
        font-size: {max(8, base_size - 1)}pt;
    }}

    /* ---------- نوارهای پیمایش ---------- */
    QScrollArea {{ background: transparent; border: none; }}
    QScrollBar:vertical {{ background: transparent; width: 9px; margin: 3px 2px; }}
    QScrollBar::handle:vertical {{
        background: {c['border']}; min-height: 40px; border-radius: 3.5px;
    }}
    QScrollBar::handle:vertical:hover {{ background: {c['ink_3']}; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0; width: 0; background: none; border: none;
    }}
    QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical {{ image: none; height: 0; width: 0; }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
    QScrollBar:horizontal {{ background: transparent; height: 9px; margin: 2px 3px; }}
    QScrollBar::handle:horizontal {{
        background: {c['border']}; min-width: 40px; border-radius: 3.5px;
    }}
    QScrollBar::handle:horizontal:hover {{ background: {c['ink_3']}; }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        height: 0; width: 0; background: none; border: none;
    }}
    QScrollBar::left-arrow:horizontal, QScrollBar::right-arrow:horizontal {{ image: none; height: 0; width: 0; }}

    /* ---------- برچسب‌ها ---------- */
    QLabel {{ background: transparent; }}
    QLabel#Eyebrow {{
        color: {c['ink_3']}; font-size: {max(7.5, base_size - 2)}pt; font-weight: 700;
        letter-spacing: 1px;
    }}
    QLabel#PageTitle {{ font-size: {base_size + 7}pt; font-weight: 800; color: {c['ink']}; }}
    QLabel#PageSubtitle {{ color: {c['ink_3']}; font-size: {base_size}pt; }}
    QLabel#SectionTitle {{ font-size: {base_size + 2}pt; font-weight: 800; color: {c['ink']}; }}
    QLabel#CardTitle {{ color: {c['ink_3']}; font-size: {max(8, base_size - 1)}pt; font-weight: 600; }}
    QLabel#CardValue {{ color: {c['ink']}; }}
    QLabel#Muted {{ color: {c['ink_3']}; }}
    QLabel#Subtle {{ color: {c['ink_2']}; }}
    QLabel#Chip {{
        background: {c['surface_3']}; color: {c['ink_2']};
        border-radius: 10px; padding: 5px 12px; font-weight: 700;
        font-size: {max(8, base_size - 1)}pt;
    }}
    QLabel#ChipBrand {{
        background: {c['brand_soft']}; color: {c['brand_deep']};
        border-radius: 10px; padding: 5px 12px; font-weight: 700;
        font-size: {max(8, base_size - 1)}pt;
    }}
    QLabel#EmptyTitle {{ font-size: {base_size + 3}pt; font-weight: 800; color: {c['ink']}; }}
    QLabel#FileName {{ font-weight: 700; color: {c['ink']}; }}

    /* ---------- دکمه‌ها ---------- */
    QPushButton {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: {RADIUS['button']}px;
        padding: 8px 16px;
        color: {c['ink']};
        font-weight: 700;
        font-size: {base_size}pt;
    }}
    QPushButton:hover {{ background: {c['surface_2']}; border-color: {c['brand_2']}; color: {c['brand_deep']}; }}
    QPushButton:pressed {{ background: {c['surface_3']}; }}
    QPushButton:disabled {{ color: {c['ink_3']}; border-color: {c['border_soft']}; background: {c['surface_2']}; }}
    QPushButton#Primary {{
        background: {primary_gradient};
        color: {c['on_brand']};
        border: 1px solid transparent;
        padding: 9px 20px;
        font-weight: 800;
    }}
    QPushButton#Primary:hover {{ background: {c['brand_deep']}; color: #FFFFFF; border-color: transparent; }}
    QPushButton#Soft {{
        background: {c['brand_soft']}; color: {c['brand_deep']};
        border: 1px solid transparent; font-weight: 700;
    }}
    QPushButton#Soft:hover {{ background: {c['brand_tint']}; border-color: {c['brand_2']}; color: {c['brand_deep']}; }}
    QPushButton#Link {{
        background: transparent; border: none; color: {c['brand_deep']};
        padding: 4px 2px; font-weight: 700;
    }}
    QPushButton#Link:hover {{ color: {c['brand']}; text-decoration: underline; }}
    QPushButton#Ghost {{
        background: transparent; border: 1px solid transparent; color: {c['ink_2']};
    }}
    QPushButton#Ghost:hover {{ background: {c['surface_3']}; color: {c['brand_deep']}; border-color: transparent; }}
    QPushButton#Danger {{ color: {c['danger']}; border-color: {alpha(c['danger'], 90).name(QColor.HexArgb)}; background: transparent; }}
    QPushButton#Danger:hover {{ background: {c['danger_soft']}; }}
    QPushButton#OnHero {{
        background: rgba(255,255,255,0.16);
        color: {c['hero_ink']};
        border: 1px solid rgba(255,255,255,0.30);
        padding: 9px 16px;
    }}
    QPushButton#OnHero:hover {{ background: rgba(255,255,255,0.28); color: #FFFFFF; }}
    QPushButton#OnHeroSolid {{
        background: #FFFFFF; color: {c['brand_deep']}; border: 1px solid #FFFFFF; padding: 9px 18px;
    }}
    QPushButton#OnHeroSolid:hover {{ background: {c['brand_soft']}; color: {c['brand_deep']}; }}
    QPushButton#Chip {{
        background: {c['surface']}; border: 1px solid {c['border']};
        border-radius: 13px; padding: 6px 14px; font-weight: 600;
        font-size: {max(8, base_size - 1)}pt; color: {c['ink_2']};
    }}
    QPushButton#Chip:hover {{ border-color: {c['brand_2']}; color: {c['brand_deep']}; }}
    QPushButton#Chip:checked {{
        background: {c['brand_soft']}; border-color: {c['brand_2']};
        color: {c['brand_deep']}; font-weight: 800;
    }}

    /* ---------- ریل کنار (ناوبری) ---------- */
    QPushButton#RailButton {{
        background: transparent; border: none; color: {c['rail_muted']};
        border-radius: 14px; padding: 8px 4px; font-weight: 600;
        font-size: {max(7.5, base_size - 2)}pt;
    }}
    QPushButton#RailButton:hover {{ background: rgba(255,255,255,0.07); color: {c['rail_ink']}; }}
    QPushButton#RailButton:checked {{
        background: {c['rail_active']}; color: {c['rail_ink']};
        font-weight: 800;
    }}
    QPushButton#RailUtility {{
        background: transparent; border: none; border-radius: 12px; padding: 6px;
    }}
    QPushButton#RailUtility:hover {{ background: rgba(255,255,255,0.08); }}

    /* ---------- ورودی‌ها ---------- */
    QLineEdit, QDateEdit {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: {RADIUS['button']}px;
        padding: 8px 12px;
        color: {c['ink']};
        selection-background-color: {c['brand']};
        selection-color: #FFFFFF;
    }}
    QLineEdit:focus, QDateEdit:focus {{ border: 1px solid {c['brand_2']}; }}
    QLineEdit#Search {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 16px;
        padding: 8px 14px 8px 14px;
        font-size: {base_size}pt;
    }}
    QLineEdit#Search:focus {{ border: 1px solid {c['brand_2']}; background: {c['surface']}; }}
    QPlainTextEdit, QTextEdit, QTextBrowser {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: {RADIUS['inner']}px;
        padding: 10px 12px;
        color: {c['ink']};
        selection-background-color: {c['brand']};
        selection-color: #FFFFFF;
    }}
    QPlainTextEdit:focus, QTextEdit:focus {{ border: 1px solid {c['brand_2']}; }}
    QComboBox {{
        background: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: {RADIUS['button']}px;
        padding: 8px 12px 8px 32px;
        color: {c['ink']};
        min-height: 20px;
    }}
    QComboBox:hover {{ border-color: {c['brand_2']}; }}
    QComboBox QAbstractItemView {{
        background: {c['surface']}; border: 1px solid {c['border']};
        border-radius: 10px; padding: 4px;
        selection-background-color: {c['brand_soft']};
        selection-color: {c['brand_deep']};
        outline: none;
    }}
    /* عددشمار سفارشی */
    QPushButton#Stepper {{
        background: {c['surface_3']}; border: none; border-radius: 8px;
        padding: 0; color: {c['ink_2']}; font-weight: 800; font-size: 12pt;
        min-width: 26px; min-height: 26px; max-width: 26px; max-height: 26px;
    }}
    QPushButton#Stepper:hover {{ background: {c['brand_soft']}; color: {c['brand_deep']}; }}
    QPushButton#Stepper:pressed {{ background: {c['brand_soft']}; }}
    QPushButton#Stepper:disabled {{ color: {c['ink_3']}; background: {c['surface_2']}; }}
    QCheckBox {{ spacing: 8px; color: {c['ink_2']}; }}
    QGroupBox {{
        background: transparent; border: 1px solid {c['border']};
        border-radius: {RADIUS['card']}px; margin-top: 18px; padding: 18px 14px 14px 14px;
        font-weight: 800; color: {c['ink']};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin; subcontrol-position: top right;
        margin-right: 16px; padding: 0 8px; color: {c['brand_deep']};
    }}

    /* ---------- جدول ---------- */
    QTableView {{
        background: transparent;
        alternate-background-color: {c['row_alt']};
        border: none;
        gridline-color: transparent;
        selection-background-color: {c['selection']};
        selection-color: {c['ink']};
        outline: none;
    }}
    QTableView::item {{ border: none; padding: 0; }}
    QHeaderView {{ background: transparent; border: none; }}
    QHeaderView::section {{
        background: {c['surface_2']};
        color: {c['ink_3']};
        border: none;
        border-bottom: 1px solid {c['border']};
        padding: 11px 10px;
        font-weight: 700;
        font-size: {max(8, base_size - 1.5)}pt;
    }}
    QHeaderView::section:hover {{ color: {c['brand_deep']}; background: {c['brand_tint']}; }}
    QTableCornerButton::section {{ background: {c['surface_2']}; border: none; }}

    /* ---------- منوها ---------- */
    QMenuBar {{
        background: {c['surface']}; color: {c['ink_2']};
        padding: 2px 10px; border-bottom: 1px solid {c['border_soft']};
    }}
    QMenuBar::item {{ padding: 5px 12px; border-radius: 8px; }}
    QMenuBar::item:selected {{ background: {c['brand_soft']}; color: {c['brand_deep']}; }}
    QMenu {{
        background: {c['surface']}; border: 1px solid {c['border']};
        border-radius: 12px; padding: 6px;
    }}
    QMenu::item {{ padding: 8px 28px 8px 18px; border-radius: 8px; color: {c['ink']}; }}
    QMenu::item:selected {{ background: {c['brand_soft']}; color: {c['brand_deep']}; }}
    QMenu::separator {{ height: 1px; background: {c['border_soft']}; margin: 5px 8px; }}

    /* ---------- متفرقه ---------- */
    QListWidget {{
        background: transparent; border: none; outline: none;
    }}
    QListWidget::item {{ padding: 10px 8px; border-radius: 10px; color: {c['ink_2']}; }}
    QListWidget::item:selected {{ background: {c['brand_soft']}; color: {c['brand_deep']}; }}
    QListWidget::item:hover {{ background: {c['surface_3']}; }}
    QStatusBar {{ background: {c['surface']}; border-top: 1px solid {c['border_soft']}; min-height: 26px; }}
    QStatusBar::item {{ border: none; }}
    QStatusBar QLabel {{ color: {c['ink_3']}; padding: 0 10px; font-size: {max(7.5, base_size - 2)}pt; }}
    QStatusBar QLabel#StatusStrong {{ color: {c['ink_2']}; font-weight: 700; }}

    /* ---------- نشان‌ها و ریزبرچسب‌ها ---------- */
    QLabel#Overline {{
        color: {c['ink_3']}; font-size: {max(7.5, base_size - 2.5)}pt;
        font-weight: 700; letter-spacing: 0.6px;
    }}
    QLabel#StatusPill {{
        background: {c['success_soft']}; color: {c['success']};
        border-radius: 11px; padding: 5px 12px; font-weight: 700;
        font-size: {max(8, base_size - 1.5)}pt;
    }}
    QLabel#StatusPillOff {{ background: {c['danger_soft']}; color: {c['danger']}; }}
    QLabel#HeroKpi {{
        color: {c['hero_ink']}; font-size: {base_size + 2}pt; font-weight: 800;
    }}
    QLabel#HeroKpiHint {{ color: {c['hero_muted']}; font-size: {max(7.5, base_size - 2)}pt; }}
    QFrame#Divider {{ background: {c['border_soft']}; max-height: 1px; min-height: 1px; border: none; }}
    QFrame#RailDivider {{ background: rgba(255,255,255,0.08); max-height: 1px; min-height: 1px; border: none; }}
    QProgressBar {{
        background: {c['border_soft']}; border: none; border-radius: 4px;
        height: 8px; text-align: center; color: transparent;
    }}
    QProgressBar::chunk {{ background: {primary_gradient}; border-radius: 4px; }}
    QToolButton {{
        background: transparent; border: 1px solid {c['border']};
        border-radius: 10px; padding: 6px 12px; color: {c['ink_2']}; font-weight: 600;
    }}
    QToolButton:hover {{ background: {c['surface_2']}; border-color: {c['brand_2']}; color: {c['brand_deep']}; }}
    QToolButton::menu-indicator {{ image: none; width: 0; }}
    QDialog {{ background: {c['bg']}; }}
    """


def apply_theme(app: QApplication, theme: str, base_size: int = 10) -> Dict[str, str]:
    colors = THEMES.get(theme, THEMES["light"])
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setLayoutDirection(Qt.RightToLeft)
    app.setFont(app_font(base_size))
    app.setStyleSheet(build_stylesheet(colors, base_size))

    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(colors["bg"]))
    palette.setColor(QPalette.WindowText, QColor(colors["ink"]))
    palette.setColor(QPalette.Base, QColor(colors["surface"]))
    palette.setColor(QPalette.AlternateBase, QColor(colors["row_alt"]))
    palette.setColor(QPalette.Text, QColor(colors["ink"]))
    palette.setColor(QPalette.Button, QColor(colors["surface"]))
    palette.setColor(QPalette.ButtonText, QColor(colors["ink"]))
    palette.setColor(QPalette.Highlight, QColor(colors["brand"]))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    palette.setColor(QPalette.ToolTipBase, QColor(colors["surface_inv"]))
    palette.setColor(QPalette.ToolTipText, QColor(colors["surface"]))
    palette.setColor(QPalette.PlaceholderText, QColor(colors["ink_3"]))
    app.setPalette(palette)
    return colors



# ==========================================================================
# ویجت‌های پایه و آیکون‌ها
# ==========================================================================

# ---------------------------------------------------------------------------
# آیکون‌های برداری (سبک خطی، بدون فایل تصویری)
# ---------------------------------------------------------------------------
def _icon_paths(kind: str, painter: QPainter) -> None:
    """رسم آیکون در شبکهٔ ۲۴×۲۴ (قلم و مقیاس قبلاً تنظیم شده است)."""
    pen_round = QPen(painter.pen())
    pen_round.setCapStyle(Qt.RoundCap)
    pen_round.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen_round)

    def rect(x, y, w, h, r=2.0):
        painter.drawRoundedRect(QRectF(x, y, w, h), r, r)

    def line(x1, y1, x2, y2):
        painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))

    def poly(points, close=False):
        path = QPainterPath()
        path.moveTo(*points[0])
        for point in points[1:]:
            path.lineTo(*point)
        if close:
            path.closeSubpath()
        painter.drawPath(path)

    if kind == "dashboard":
        rect(3, 3, 7.4, 8.2, 2)
        rect(13.6, 3, 7.4, 5.4, 2)
        rect(13.6, 11.4, 7.4, 9.6, 2)
        rect(3, 14.2, 7.4, 6.8, 2)
    elif kind == "table":
        rect(3, 4.2, 18, 15.6, 2.6)
        line(3, 9.4, 21, 9.4)
        line(3, 14.6, 21, 14.6)
        line(9.2, 4.2, 9.2, 19.8)
    elif kind == "report":
        poly([(6, 2.8), (14, 2.8), (19, 7.8), (19, 21.2), (6, 21.2)], True)
        poly([(14, 2.8), (14, 7.8), (19, 7.8)])
        line(9, 12.4, 16, 12.4)
        line(9, 16.4, 16, 16.4)
    elif kind == "pdf":
        poly([(6, 2.8), (14, 2.8), (19, 7.8), (19, 21.2), (6, 21.2)], True)
        poly([(14, 2.8), (14, 7.8), (19, 7.8)])
    elif kind == "settings":
        painter.drawEllipse(QPointF(12, 12), 3.1, 3.1)
        for index in range(8):
            angle = index * math.pi / 4
            line(12 + 5.6 * math.cos(angle), 12 + 5.6 * math.sin(angle),
                 12 + 8.2 * math.cos(angle), 12 + 8.2 * math.sin(angle))
        painter.drawEllipse(QPointF(12, 12), 7.4, 7.4)
    elif kind == "help":
        painter.drawEllipse(QPointF(12, 12), 9, 9)
        path = QPainterPath()
        path.moveTo(9.4, 9.4)
        path.quadTo(9.4, 6.8, 12, 6.8)
        path.quadTo(14.8, 6.8, 14.8, 9.2)
        path.quadTo(14.8, 11.4, 12, 12.2)
        path.lineTo(12, 14.2)
        painter.drawPath(path)
        painter.drawPoint(QPointF(12, 17.2))
    elif kind == "search":
        painter.drawEllipse(QPointF(11, 11), 6.2, 6.2)
        line(15.6, 15.6, 20.4, 20.4)
    elif kind == "refresh":
        painter.drawArc(QRectF(3.8, 3.8, 16.4, 16.4), 40 * 16, 280 * 16)
        poly([(20.2, 3.2), (20.2, 8.4), (15.0, 8.4)])
    elif kind == "download":
        line(12, 3.4, 12, 15.2)
        poly([(7.6, 11.0), (12, 15.4), (16.4, 11.0)])
        poly([(4.4, 15.6), (4.4, 19.4), (19.6, 19.4), (19.6, 15.6)])
    elif kind == "upload":
        line(12, 20.6, 12, 8.8)
        poly([(7.6, 13.0), (12, 8.6), (16.4, 13.0)])
        poly([(4.4, 4.4), (19.6, 4.4)])
    elif kind == "sun":
        painter.drawEllipse(QPointF(12, 12), 4.4, 4.4)
        for index in range(8):
            angle = index * math.pi / 4
            line(12 + 7.0 * math.cos(angle), 12 + 7.0 * math.sin(angle),
                 12 + 9.4 * math.cos(angle), 12 + 9.4 * math.sin(angle))
    elif kind == "moon":
        painter.drawArc(QRectF(4, 2.6, 16, 18.8), 60 * 16, 240 * 16)
        painter.drawArc(QRectF(8.4, 1.4, 14.4, 17.6), 90 * 16, 180 * 16)
    elif kind == "close":
        line(6, 6, 18, 18)
        line(18, 6, 6, 18)
    elif kind in ("chevron-down", "chevron-up", "chevron-left", "chevron-right"):
        if kind == "chevron-down":
            poly([(6.5, 9.5), (12, 15), (17.5, 9.5)])
        elif kind == "chevron-up":
            poly([(6.5, 14.5), (12, 9), (17.5, 14.5)])
        elif kind == "chevron-left":
            poly([(14.5, 6.5), (9, 12), (14.5, 17.5)])
        else:
            poly([(9.5, 6.5), (15, 12), (9.5, 17.5)])
    elif kind == "arrow-left":
        line(20, 12, 4.6, 12)
        poly([(10.4, 5.8), (4.2, 12), (10.4, 18.2)])
    elif kind == "arrow-right":
        line(4, 12, 19.4, 12)
        poly([(13.6, 5.8), (19.8, 12), (13.6, 18.2)])
    elif kind == "folder":
        poly([(3.4, 6.8), (9.6, 6.8), (11.4, 9.0), (20.6, 9.0), (20.6, 19.4), (3.4, 19.4)], True)
    elif kind == "filter":
        poly([(3.6, 5.4), (20.4, 5.4), (13.8, 13.4), (13.8, 19.6), (10.2, 21.0), (10.2, 13.4)], True)
    elif kind == "bookmark":
        poly([(6.4, 3.4), (17.6, 3.4), (17.6, 20.8), (12, 16.4), (6.4, 20.8)], True)
    elif kind == "note":
        rect(4, 4.4, 16, 15.2, 3)
        line(8, 9.4, 16, 9.4)
        line(8, 13, 16, 13)
        line(8, 16.4, 12.6, 16.4)
    elif kind == "printer":
        poly([(6.6, 8.4), (6.6, 3.4), (17.4, 3.4), (17.4, 8.4)])
        rect(3.4, 8.4, 17.2, 8.4, 2.4)
        rect(6.6, 14.2, 10.8, 6.4, 2)
    elif kind == "external":
        poly([(13.6, 3.6), (20.4, 3.6), (20.4, 10.4)])
        line(20.4, 3.6, 11.4, 12.6)
        poly([(17.2, 14.4), (17.2, 20.4), (3.6, 20.4), (3.6, 6.8), (9.6, 6.8)])
    elif kind == "check":
        painter.drawEllipse(QPointF(12, 12), 9, 9)
        poly([(7.6, 12.2), (10.8, 15.4), (16.6, 8.8)])
    elif kind == "check-plain":
        poly([(4.6, 12.4), (9.6, 17.4), (19.4, 6.8)])
    elif kind == "alert":
        poly([(12, 3.2), (21.4, 20.4), (2.6, 20.4)], True)
        line(12, 9.4, 12, 14.6)
        painter.drawPoint(QPointF(12, 17.4))
    elif kind == "info":
        painter.drawEllipse(QPointF(12, 12), 9, 9)
        line(12, 11, 12, 16.6)
        painter.drawPoint(QPointF(12, 7.8))
    elif kind == "chart":
        line(3.6, 3.2, 3.6, 20.4)
        line(3.6, 20.4, 20.8, 20.4)
        rect(7.2, 12.4, 3.4, 8, 1.4)
        rect(12.4, 8.6, 3.4, 11.8, 1.4)
        rect(17.6, 4.8, 3.4, 15.6, 1.4)
    elif kind == "activity":
        poly([(2.6, 13.4), (7.4, 13.4), (10.4, 5.6), (14.4, 19.2), (17.0, 13.4), (21.4, 13.4)])
    elif kind == "cube":
        poly([(12, 2.8), (20.6, 7.4), (20.6, 16.6), (12, 21.2), (3.4, 16.6), (3.4, 7.4)], True)
        poly([(3.4, 7.4), (12, 12), (20.6, 7.4)])
        line(12, 12, 12, 21.2)
    elif kind == "shield":
        poly([(12, 2.8), (20.4, 6.2), (20.0, 13.2), (12, 21.2), (4.0, 13.2), (3.6, 6.2)], True)
        poly([(8.4, 11.8), (11.2, 14.6), (15.8, 9.4)])
    elif kind == "scissors":
        painter.drawEllipse(QPointF(7.0, 17.4), 2.6, 2.6)
        painter.drawEllipse(QPointF(17.0, 17.4), 2.6, 2.6)
        line(8.8, 15.6, 18.2, 3.4)
        line(15.2, 15.6, 5.8, 3.4)
    elif kind == "truck":
        rect(2.6, 7.4, 11.4, 9.6, 1.6)
        poly([(14, 10.4), (17.6, 10.4), (21.4, 14.2), (21.4, 17.0), (14, 17.0)], True)
        painter.drawEllipse(QPointF(7.0, 18.6), 2.2, 2.2)
        painter.drawEllipse(QPointF(18.2, 18.6), 2.2, 2.2)
    elif kind == "clock":
        painter.drawEllipse(QPointF(12, 12), 9, 9)
        line(12, 6.8, 12, 12.2)
        line(12, 12.2, 16.0, 14.4)
    elif kind == "star":
        path = QPainterPath()
        for index in range(10):
            angle = -math.pi / 2 + index * math.pi / 5
            radius = 9.2 if index % 2 == 0 else 3.9
            point = QPointF(12 + radius * math.cos(angle), 12 + radius * math.sin(angle))
            path.moveTo(point) if index == 0 else path.lineTo(point)
        path.closeSubpath()
        painter.drawPath(path)
    elif kind == "users":
        painter.drawEllipse(QPointF(9.0, 8.4), 3.4, 3.4)
        painter.drawArc(QRectF(2.6, 12.6, 12.8, 11.2), 0, 180 * 16)
        painter.drawEllipse(QPointF(17.6, 9.4), 2.6, 2.6)
        painter.drawArc(QRectF(13.4, 13.4, 8.6, 8.6), 0, 180 * 16)
    elif kind == "money":
        rect(2.6, 6.4, 18.8, 11.2, 2.4)
        painter.drawEllipse(QPointF(12, 12), 3.0, 3.0)
        line(6.0, 9.4, 6.0, 9.4)
        line(18.0, 14.6, 18.0, 14.6)
    elif kind == "calendar":
        rect(3.4, 5.2, 17.2, 15.6, 2.6)
        line(3.4, 10.2, 20.6, 10.2)
        line(8.2, 2.8, 8.2, 6.8)
        line(15.8, 2.8, 15.8, 6.8)
    elif kind == "menu":
        for y in (6.6, 12, 17.4):
            line(4.2, y, 19.8, y)
    elif kind == "dots":
        painter.drawEllipse(QPointF(12, 5.4), 1.5, 1.5)
        painter.drawEllipse(QPointF(12, 12), 1.5, 1.5)
        painter.drawEllipse(QPointF(12, 18.6), 1.5, 1.5)
    elif kind == "sparkle":
        poly([(12, 2.8), (13.9, 9.6), (20.8, 11.6), (13.9, 13.6), (12, 20.6), (10.1, 13.6), (3.2, 11.6), (10.1, 9.6)], True)
    elif kind == "layers":
        poly([(12, 2.8), (21.2, 7.6), (12, 12.4), (2.8, 7.6)], True)
        poly([(3.4, 12.2), (12, 16.4), (20.6, 12.2)])
        poly([(3.4, 16.6), (12, 20.8), (20.6, 16.6)])
    elif kind == "inbox":
        poly([(3.2, 12.4), (6.2, 4.0), (17.8, 4.0), (20.8, 12.4), (20.8, 20.0), (3.2, 20.0)], True)
        poly([(3.2, 12.4), (8.6, 12.4), (9.8, 15.2), (14.2, 15.2), (15.4, 12.4), (20.8, 12.4)])
    elif kind == "pin":
        painter.drawEllipse(QPointF(12, 9.4), 3.4, 3.4)
        painter.drawArc(QRectF(5.4, 2.4, 13.2, 13.2), 200 * 16, 140 * 16)
        line(12, 15.6, 12, 21.6)
    else:
        painter.drawEllipse(QPointF(12, 12), 8.4, 8.4)


def draw_icon(painter: QPainter, kind: str, rect: QRectF, color: QColor,
              width: float = 1.7, dashed: bool = False) -> None:
    """رسم آیکون در کادر داده‌شده."""
    painter.save()
    painter.setRenderHint(QPainter.Antialiasing, True)
    scale = min(rect.width(), rect.height()) / 24.0
    painter.translate(rect.x() + (rect.width() - 24 * scale) / 2,
                      rect.y() + (rect.height() - 24 * scale) / 2)
    painter.scale(scale, scale)
    pen = QPen(color, max(0.9, width / max(scale, 0.001) * scale if scale else width))
    pen.setWidthF(width)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    if dashed:
        pen.setStyle(Qt.DashLine)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)
    if kind == "pdf":
        poly = QPainterPath()
        poly.moveTo(6, 2.8)
        poly.lineTo(14, 2.8)
        poly.lineTo(19, 7.8)
        poly.lineTo(19, 21.2)
        poly.lineTo(6, 21.2)
        poly.closeSubpath()
        painter.drawPath(poly)
        folder = QPainterPath()
        folder.moveTo(14, 2.8)
        folder.lineTo(14, 7.8)
        folder.lineTo(19, 7.8)
        painter.drawPath(folder)
        label_font = QFont(font_family())
        label_font.setPixelSize(7)
        label_font.setBold(True)
        painter.setFont(label_font)
        painter.setPen(QPen(color))
        painter.drawText(QRectF(5.4, 10.4, 14.2, 8.4), Qt.AlignCenter, "PDF")
    else:
        _icon_paths(kind, painter)
    painter.restore()


class IconWidget(QWidget):
    """آیکون ساده با رنگ و اندازهٔ دلخواه."""

    def __init__(self, kind: str, color: str, size: int = 20, width: float = 1.7, parent=None) -> None:
        super().__init__(parent)
        self.kind = kind
        self.color = QColor(color)
        self.stroke = width
        self.setFixedSize(size, size)

    def set_color(self, color: str) -> None:
        self.color = QColor(color)
        self.update()

    def set_kind(self, kind: str) -> None:
        self.kind = kind
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        draw_icon(painter, self.kind, QRectF(0, 0, self.width(), self.height()),
                  self.color, self.stroke)


class IconTile(QWidget):
    """آیکون داخل یک کاشی رنگی گرد."""

    def __init__(self, kind: str, fg: str, bg: str, size: int = 44, radius: int = 14,
                 width: float = 1.8, parent=None) -> None:
        super().__init__(parent)
        self.kind = kind
        self.fg = QColor(fg)
        self.bg = QColor(bg)
        self.radius = radius
        self.stroke = width
        self.setFixedSize(size, size)

    def set_colors(self, fg: str, bg: str) -> None:
        self.fg, self.bg = QColor(fg), QColor(bg)
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        body = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        gradient = QLinearGradient(body.topLeft(), body.bottomRight())
        gradient.setColorAt(0.0, self.bg.lighter(104))
        gradient.setColorAt(1.0, self.bg)
        painter.setPen(Qt.NoPen)
        painter.setBrush(gradient)
        painter.drawRoundedRect(body, self.radius, self.radius)
        margin = body.width() * 0.24
        draw_icon(painter, self.kind,
                  body.adjusted(margin, margin, -margin, -margin), self.fg, self.stroke)


class Avatar(QWidget):
    """دایرهٔ حرف اول نام شرکت با رنگ ثابت بر اساس نام."""

    PALETTES = [
        ("#0E7A6A", "#18A492"),
        ("#2A63A6", "#4C8BD1"),
        ("#8A5A2B", "#C58A2E"),
        ("#7A4780", "#A96BB2"),
        ("#2F6E45", "#4E9C6A"),
        ("#9B4A3F", "#C96B5B"),
        ("#3E5C76", "#5D83A6"),
        ("#6B5B2E", "#A08A46"),
    ]

    def __init__(self, name: str, size: int = 34, parent=None) -> None:
        super().__init__(parent)
        self.name = name or "?"
        self.setFixedSize(size, size)
        digest = sum(ord(ch) for ch in self.name) % len(self.PALETTES)
        self.c1, self.c2 = self.PALETTES[digest]

    def set_name(self, name: str) -> None:
        """تغییر نام و بازمحاسبهٔ حرف اول و رنگ."""
        self.name = name or "?"
        digest = sum(ord(ch) for ch in self.name) % len(self.PALETTES)
        self.c1, self.c2 = self.PALETTES[digest]
        self.update()

    def initials(self) -> str:
        words = [w for w in self.name.replace("‌", " ").split() if w]
        if not words:
            return "؟"
        if len(words) == 1:
            return words[0][:2]
        return (words[0][0] + words[1][0])

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        gradient = QLinearGradient(rect.topLeft(), rect.bottomRight())
        gradient.setColorAt(0, QColor(self.c2))
        gradient.setColorAt(1, QColor(self.c1))
        painter.setPen(Qt.NoPen)
        painter.setBrush(gradient)
        painter.drawEllipse(rect)
        font = app_font(max(7, int(self.height() * 0.30)), QFont.Bold)
        painter.setFont(font)
        painter.setPen(QPen(QColor("#FFFFFF")))
        painter.drawText(rect, Qt.AlignCenter, self.initials())


def chip_label(text: str, fg: str, bg: str, radius: int = 9, padding: str = "4px 10px",
               size: int = 8, bold: bool = True) -> QLabel:
    label = QLabel(text)
    weight = 800 if bold else 600
    label.setStyleSheet(
        f"background: {bg}; color: {fg}; border-radius: {radius}px;"
        f" padding: {padding}; font-size: {size}pt; font-weight: {weight};"
    )
    label.setAlignment(Qt.AlignCenter)
    return label


# ---------------------------------------------------------------------------
# کارت با سایهٔ نرم
# ---------------------------------------------------------------------------
class Card(QFrame):
    """قاب کارت با سایهٔ چندلایه (بدون نیاز به افکت‌های سنگین)."""

    SHADOW = 6

    def __init__(self, colors: Dict[str, str], radius: int = RADIUS["card"],
                 padding: int = 16, variant: str = "surface",
                 hoverable: bool = False, parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.radius = radius
        self.padding = padding
        self.variant = variant
        self.hoverable = hoverable
        self._hover = False
        self._accent: Optional[QColor] = None
        self.setAttribute(Qt.WA_Hover, True)
        self.setLayoutDirection(Qt.RightToLeft)
        self.body = QVBoxLayout(self)
        inset = self.SHADOW
        self.body.setContentsMargins(inset + padding, inset + padding,
                                     inset + padding, inset + padding)
        self.body.setSpacing(7)

    # ---- تنظیمات ظاهری ----
    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def set_accent(self, color: Optional[str]) -> None:
        self._accent = QColor(color) if color else None
        self.update()

    def enterEvent(self, event) -> None:  # noqa: N802
        if self.hoverable:
            self._hover = True
            self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        if self.hoverable:
            self._hover = False
            self.update()
        super().leaveEvent(event)

    def body_rect(self) -> QRectF:
        return QRectF(self.rect()).adjusted(self.SHADOW, self.SHADOW,
                                            -self.SHADOW, -self.SHADOW)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        c = self.colors
        body = self.body_rect()
        radius = self.radius

        if self.variant != "flat":
            shadow = QColor(c["shadow"])
            if self.variant == "hero":
                strength = 18
            else:
                strength = 15 if self._hover else 9
            layers = 6
            for index in range(layers, 0, -1):
                ratio = index / layers
                alpha_value = int(strength * (1.0 - ratio) ** 1.6) + 2
                shadow.setAlpha(max(1, alpha_value))
                painter.setPen(QPen(shadow, 1.4))
                painter.setBrush(Qt.NoBrush)
                grow = index * 0.85
                painter.drawRoundedRect(
                    body.adjusted(-grow, -grow * 0.5 + 1.6, grow, grow + 2.2),
                    radius + grow * 0.7, radius + grow * 0.7,
                )

        if self.variant == "hero":
            gradient = QLinearGradient(body.topLeft(), body.bottomRight())
            gradient.setColorAt(0.0, QColor(c["hero_3"]))
            gradient.setColorAt(0.45, QColor(c["hero_2"]))
            gradient.setColorAt(1.0, QColor(c["hero_1"]))
            painter.setPen(Qt.NoPen)
            painter.setBrush(gradient)
            painter.drawRoundedRect(body, radius, radius)
            # نور و موج‌های تزئینی
            painter.save()
            clip = QPainterPath()
            clip.addRoundedRect(body, radius, radius)
            painter.setClipPath(clip)
            glow = QRadialGradient(
                QPointF(body.right() - body.width() * 0.16, body.top() + body.height() * 0.10),
                body.width() * 0.62,
            )
            glow.setColorAt(0.0, QColor(255, 255, 255, 40))
            glow.setColorAt(0.55, QColor(255, 255, 255, 12))
            glow.setColorAt(1.0, QColor(255, 255, 255, 0))
            painter.setPen(Qt.NoPen)
            painter.setBrush(glow)
            painter.drawRect(body)

            painter.setBrush(QColor(255, 255, 255, 12))
            painter.drawEllipse(
                QPointF(body.left() + body.width() * 0.16, body.bottom() + body.height() * 0.24),
                body.width() * 0.34, body.width() * 0.34,
            )
            painter.setBrush(QColor(255, 255, 255, 8))
            painter.drawEllipse(
                QPointF(body.right() - body.width() * 0.04, body.bottom() + body.height() * 0.10),
                body.width() * 0.30, body.width() * 0.30,
            )
            # خط نور بالای کارت
            highlight = QLinearGradient(body.topLeft(), body.topRight())
            highlight.setColorAt(0.0, QColor(255, 255, 255, 0))
            highlight.setColorAt(0.35, QColor(255, 255, 255, 70))
            highlight.setColorAt(0.75, QColor(255, 255, 255, 40))
            highlight.setColorAt(1.0, QColor(255, 255, 255, 0))
            painter.setBrush(Qt.NoBrush)
            painter.setPen(QPen(QBrush(highlight), 1))
            painter.drawLine(QPointF(body.left() + radius, body.top() + 0.5),
                             QPointF(body.right() - radius, body.top() + 0.5))
            painter.restore()
        elif self.variant == "soft":
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(c["surface_2"]))
            painter.drawRoundedRect(body, radius, radius)
            painter.setPen(QPen(QColor(c["border_soft"]), 1))
            painter.setBrush(Qt.NoBrush)
            painter.drawRoundedRect(body.adjusted(0.5, 0.5, -0.5, -0.5), radius, radius)
        elif self.variant == "flat":
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(c["surface"]))
            painter.drawRoundedRect(body, radius, radius)
        else:
            gradient = QLinearGradient(body.topLeft(), body.bottomLeft())
            gradient.setColorAt(0.0, QColor(c["surface"]))
            gradient.setColorAt(1.0, mix(c["surface"], c["surface_3"], 0.35))
            painter.setPen(Qt.NoPen)
            painter.setBrush(gradient)
            painter.drawRoundedRect(body, radius, radius)
            border = QColor(c["border"])
            if self._hover:
                border = QColor(c["brand_2"])
            painter.setPen(QPen(border, 1))
            painter.setBrush(Qt.NoBrush)
            painter.drawRoundedRect(body.adjusted(0.5, 0.5, -0.5, -0.5), radius, radius)

        if self._accent is not None:
            painter.setPen(Qt.NoPen)
            painter.setBrush(self._accent)
            painter.drawRoundedRect(
                QRectF(body.right() - 3, body.top() + radius * 0.9, 3,
                       max(10.0, body.height() - radius * 1.8)), 1.5, 1.5,
            )


class SectionHeader(QWidget):
    """سرتیتر یک بخش با عنوان، زیرنویس و اکشن سمت چپ."""

    def __init__(self, title: str, subtitle: str = "", parent=None) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 0)
        layout.setSpacing(8)
        title_label = QLabel(title)
        title_label.setObjectName("SectionTitle")
        set_text_aligned(title_label, title, "right")
        layout.addWidget(title_label)
        if subtitle:
            sub = QLabel(subtitle)
            sub.setObjectName("Muted")
            set_text_aligned(sub, subtitle, "right")
            layout.addWidget(sub)
        layout.addStretch(1)
        self.actions = QHBoxLayout()
        self.actions.setSpacing(6)
        layout.addLayout(self.actions)

    def add_action(self, widget: QWidget) -> None:
        self.actions.addWidget(widget)


class EmptyState(QWidget):
    """حالت خالی با تصویرسازی ساده و دکمهٔ اقدام."""

    actionClicked = pyqtSignal()

    def __init__(self, title: str, message: str, icon: str = "inbox",
                 action_text: str = "", colors: Optional[Dict[str, str]] = None, parent=None) -> None:
        super().__init__(parent)
        self.colors = colors or THEMES["light"]
        self.icon_name = icon
        self.action_text = action_text
        self.title = title
        self.message = message
        self.setMinimumHeight(300)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(6)
        layout.addStretch(1)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("EmptyTitle")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.message_label = QLabel(message)
        self.message_label.setObjectName("Muted")
        self.message_label.setAlignment(Qt.AlignCenter)
        self.message_label.setWordWrap(True)
        self.button = QPushButton(action_text)
        self.button.setObjectName("Primary")
        self.button.setCursor(Qt.PointingHandCursor)
        self.button.clicked.connect(self.actionClicked.emit)
        self.button.setVisible(bool(action_text))
        layout.addStretch(2)
        self.title_label.setVisible(False)
        layout.addStretch(3)
        layout.addWidget(self.message_label)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(self.button)
        row.addStretch(1)
        layout.addSpacing(10)
        layout.addLayout(row)
        layout.addStretch(4)

    def set_texts(self, title: str, message: str, action_text: str = "") -> None:
        self.title = title
        self.message = message
        self.message_label.setText(message)
        self.button.setText(action_text)
        self.button.setVisible(bool(action_text))

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        c = self.colors
        size = 96
        center_x = self.width() / 2
        top = self.height() * 0.30
        circle = QRectF(center_x - size / 2, top - size / 2, size, size)
        painter.setPen(Qt.NoPen)
        painter.setBrush(alpha(c["brand"], 22))
        painter.drawEllipse(circle)
        painter.setBrush(alpha(c["brand"], 34))
        painter.drawEllipse(circle.adjusted(14, 14, -14, -14))
        draw_icon(painter, self.icon_name,
                  circle.adjusted(28, 28, -28, -28), QColor(c["brand"]), 1.6)
        painter.setFont(app_font(11, QFont.Bold))
        painter.setPen(QPen(QColor(c["ink"])))
        painter.drawText(QRectF(20, top + size / 2 + 10, self.width() - 40, 28),
                         Qt.AlignHCenter, self.title)


class Toast(QWidget):
    """پیام شناور کوتاه در پایین پنجره."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setVisible(False)
        self.text = ""
        self.icon_name = "check"
        self.colors = THEMES["light"]
        self._opacity = 0.0
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._fade_out)
        self._anim = QPropertyAnimation(self, b"windowOpacity", self)

    def show_message(self, text: str, icon: str = "check", colors: Optional[Dict[str, str]] = None,
                     duration: int = 2200) -> None:
        if colors:
            self.colors = colors
        self.text = text
        self.icon_name = icon
        self.adjustSize()
        parent = self.parentWidget()
        if parent:
            x = (parent.width() - self.width()) // 2
            y = parent.height() - self.height() - 36
            self.move(x, y)
        self.setWindowOpacity(1.0)
        self.setVisible(True)
        self.raise_()
        self.update()
        self._timer.start(duration)

    def _fade_out(self) -> None:
        self._anim.stop()
        self._anim.setDuration(320)
        self._anim.setStartValue(1.0)
        self._anim.setEndValue(0.0)
        self._anim.finished.connect(lambda: self.setVisible(False))
        self._anim.start()

    def sizeHint(self) -> QSize:  # noqa: N802
        metrics = QFontMetrics(app_font(9, QFont.DemiBold))
        width = metrics.horizontalAdvance(self.text) + 74
        return QSize(int(width), 44)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        c = self.colors
        body = QRectF(self.rect()).adjusted(1, 1, -1, -1)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(c["surface_inv"]))
        painter.drawRoundedRect(body, 12, 12)
        draw_icon(painter, self.icon_name, QRectF(body.right() - 34, body.center().y() - 9, 18, 18),
                  QColor(c["brand_2"]), 1.5)
        painter.setFont(app_font(9, QFont.DemiBold))
        painter.setPen(QPen(QColor(c["surface"])))
        painter.drawText(body.adjusted(14, 0, -44, 0), Qt.AlignRight | Qt.AlignVCenter, self.text)



# ==========================================================================
# نمودارهای بومی
# ==========================================================================

# ---------------------------------------------------------------------------
# نمودارهای بومی (دونات، حلقهٔ پیشرفت، میله‌ای و راهنما)
# ---------------------------------------------------------------------------
class RingGauge(QWidget):
    """حلقهٔ پیشرفت با انیمیشن، گرادیان و متن مرکزی."""

    def __init__(self, colors: Dict[str, str], thickness: int = 14, parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.thickness = thickness
        self._value = 0.0
        self._display = 0.0
        self.caption = "پیشرفت کل"
        self.hero_mode = False
        self.setMinimumSize(170, 170)
        self._anim = QPropertyAnimation(self, b"display", self)
        self._anim.setDuration(1100)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def set_hero_mode(self, enabled: bool) -> None:
        self.hero_mode = enabled
        self.update()

    def get_display(self) -> float:
        return self._display

    def set_display(self, value: float) -> None:
        self._display = float(value)
        self.update()

    display = pyqtProperty(float, fget=get_display, fset=set_display)

    def set_value(self, value: float, animate: bool = True) -> None:
        self._value = max(0.0, min(100.0, float(value)))
        self._anim.stop()
        if not animate:
            self.set_display(self._value)
            return
        self._anim.setStartValue(self._display)
        self._anim.setEndValue(self._value)
        self._anim.start()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        c = self.colors
        side = min(self.width(), self.height()) - 8
        box = QRectF((self.width() - side) / 2, (self.height() - side) / 2, side, side)
        thickness = max(8.0, side * 0.085)
        arc_box = box.adjusted(thickness / 2 + 2, thickness / 2 + 2,
                               -thickness / 2 - 2, -thickness / 2 - 2)

        # حلقهٔ زمینه
        track_color = QColor(255, 255, 255, 55) if self.hero_mode else alpha(c["border"], 150)
        track_pen = QPen(track_color, thickness, Qt.SolidLine, Qt.RoundCap)
        painter.setPen(track_pen)
        painter.drawArc(arc_box, 210 * 16, -240 * 16)
        value = max(0.0, min(100.0, self._display))
        if value > 0:
            span = -240.0 * value / 100.0
            gradient = QConicalGradient(arc_box.center(), 210)
            if self.hero_mode:
                gradient.setColorAt(0.0, QColor("#FFFFFF"))
                gradient.setColorAt(0.40, QColor("#DCD6FF"))
                gradient.setColorAt(0.78, QColor("#A697FF"))
                gradient.setColorAt(1.0, QColor("#F6CE86"))
            else:
                gradient.setColorAt(0.0, QColor(c["brand_2"]))
                gradient.setColorAt(0.45, QColor(c["brand"]))
                gradient.setColorAt(0.75, QColor(c["brand_deep"]))
                gradient.setColorAt(1.0, QColor(c["accent"]))
            painter.setPen(QPen(QBrush(gradient), thickness, Qt.SolidLine, Qt.RoundCap))
            painter.drawArc(arc_box, 210 * 16, int(span * 16))
            # نقطهٔ انتهای حلقه
            angle = math.radians(210 + span)
            radius = arc_box.width() / 2
            end = QPointF(arc_box.center().x() + radius * math.cos(angle),
                          arc_box.center().y() - radius * math.sin(angle))
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(c["hero_1"] if self.hero_mode else c["surface"]))
            painter.drawEllipse(end, thickness * 0.22, thickness * 0.22)
            painter.setBrush(QColor("#FFFFFF" if self.hero_mode else c["brand"]))
            painter.drawEllipse(end, thickness * 0.12, thickness * 0.12)

        # متن مرکزی
        inner = arc_box.width() / 2 - thickness / 2 - 6
        percent_text = fa(f"{value:.1f}٪") if value > 0 else "—"
        size = max(16, int(inner * 0.40))
        while size > 11:
            probe = QFontMetrics(app_font(size, QFont.Black))
            if probe.horizontalAdvance(percent_text) <= inner * 1.86:
                break
            size -= 1
        painter.setFont(app_font(size, QFont.Black))
        painter.setPen(QPen(QColor(c["hero_ink"] if self.hero_mode else c["ink"])))
        painter.drawText(
            QRectF(arc_box.center().x() - inner, arc_box.center().y() - inner * 0.58,
                   inner * 2, inner * 0.62),
            Qt.AlignCenter, percent_text,
        )
        painter.setFont(app_font(max(8, int(inner * 0.115)), QFont.DemiBold))
        painter.setPen(QPen(QColor(c["hero_muted"] if self.hero_mode else c["ink_3"])))
        painter.drawText(
            QRectF(arc_box.center().x() - inner, arc_box.center().y() + inner * 0.10,
                   inner * 2, inner * 0.38),
            Qt.AlignCenter, self.caption,
        )

class DonutChart(QWidget):
    """نمودار دونات با برش‌های گرد، تعامل ماوس و متن مرکزی."""

    sliceClicked = pyqtSignal(str)

    def __init__(self, colors: Dict[str, str], palette: Sequence[str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.palette = [QColor(c) for c in palette]
        self.labels: List[str] = []
        self.values: List[int] = []
        self.hover = -1
        self.caption = "قرارداد"
        self._dark = QColor(colors["surface"]).lightness() < 128
        self.setMouseTracking(True)
        self.setMinimumSize(150, 150)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._slices: List[Tuple[float, float, str]] = []

    def set_theme(self, colors: Dict[str, str], palette: Optional[Sequence[str]] = None) -> None:
        self.colors = colors
        self._dark = QColor(colors["surface"]).lightness() < 128
        if palette:
            self.palette = [QColor(c) for c in palette]
        self.update()

    def set_data(self, labels: Sequence[str], values: Sequence[int],
                 caption: str = "قرارداد", palette: Optional[Sequence[str]] = None) -> None:
        self.labels = list(labels)
        self.values = [int(v) for v in values]
        if palette:
            self.palette = [QColor(c) for c in palette]
        self.caption = caption
        self.hover = -1
        self.update()

    @property
    def total(self) -> int:
        return int(sum(self.values))

    def _color(self, index: int) -> QColor:
        if not self.palette:
            return QColor(self.colors["brand"])
        color = QColor(self.palette[index % len(self.palette)])
        if self._dark:
            color = color.lighter(124)
        cycle = index // len(self.palette)
        return color.lighter(112 + 8 * cycle) if cycle else color

    def _geometry(self) -> Tuple[QRectF, float]:
        side = min(self.width(), self.height()) - 10
        box = QRectF((self.width() - side) / 2, (self.height() - side) / 2, side, side)
        thickness = max(14.0, min(30.0, side * 0.16))
        return box, thickness

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        c = self.colors
        box, thickness = self._geometry()
        arc_box = box.adjusted(thickness / 2, thickness / 2, -thickness / 2, -thickness / 2)
        total = self.total
        self._slices = []
        if total <= 0:
            painter.setPen(QPen(alpha(c["border"], 190), thickness * 0.7, Qt.SolidLine, Qt.RoundCap))
            painter.drawArc(arc_box, 0, 360 * 16)
            painter.setPen(QPen(QColor(c["ink_3"])))
            painter.setFont(app_font(9))
            painter.drawText(box, Qt.AlignCenter, "داده‌ای نیست")
            return

        gap = 1.1
        start = 90.0
        for value in self.values:
            span = 360.0 * value / total
            self._slices.append((start, span, ""))
            start -= span

        for index, (start_deg, span_deg, _label) in enumerate(self._slices):
            hovered = index == self.hover
            color = self._color(index)
            width = thickness + (4 if hovered else 0)
            pen = QPen(color, width, Qt.SolidLine, Qt.RoundCap)
            painter.setPen(pen)
            effective = max(0.6, span_deg - gap * 2)
            painter.drawArc(
                arc_box.adjusted(-(width - thickness) / 2, -(width - thickness) / 2,
                                 (width - thickness) / 2, (width - thickness) / 2),
                int((start_deg - gap / 2) * 16), int(-effective * 16),
            )

        # متن مرکزی
        inner = arc_box.width() / 2 - thickness / 2 - 4
        center = arc_box.center()
        if 0 <= self.hover < len(self.values):
            value = self.values[self.hover]
            caption = pretty_label(self.labels[self.hover])
            big_text = fa_num(value)
            percent = fa(f"{value / total * 100:.0f}٪")
            small_text = f"{caption} • {percent}"
            muted = c["ink_3"]
        else:
            big_text = fa_num(total)
            small_text = self.caption
            muted = c["ink_3"]

        font = app_font(max(13, int(inner * 0.30)), QFont.Black)
        painter.setFont(font)
        painter.setPen(QPen(QColor(c["ink"])))
        painter.drawText(QRectF(center.x() - inner, center.y() - inner * 0.58, inner * 2, inner * 0.66),
                         Qt.AlignCenter, big_text)
        painter.setFont(app_font(max(7, int(inner * 0.13)), QFont.DemiBold))
        painter.setPen(QPen(QColor(muted)))
        metrics = QFontMetrics(painter.font())
        painter.drawText(
            QRectF(center.x() - inner, center.y() + inner * 0.12, inner * 2, inner * 0.42),
            Qt.AlignCenter,
            metrics.elidedText(small_text, Qt.ElideRight, int(inner * 1.9)),
        )

    def _index_at(self, point: QPointF) -> int:
        if not self._slices:
            return -1
        box, thickness = self._geometry()
        center = box.center()
        distance = math.hypot(point.x() - center.x(), point.y() - center.y())
        outer = box.width() / 2
        inner = outer - thickness - 6
        if distance > outer + 2 or distance < inner:
            return -1
        angle = (math.degrees(math.atan2(-(point.y() - center.y()), point.x() - center.x()))) % 360
        for index, (start_deg, span_deg, _label) in enumerate(self._slices):
            beginning = (start_deg - span_deg) % 360
            if beginning <= angle <= start_deg or (start_deg < beginning and (angle >= beginning or angle <= start_deg)):
                return index
        return -1

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        index = self._index_at(event.pos())
        if index != self.hover:
            self.hover = index
            self.setCursor(Qt.PointingHandCursor if index >= 0 else Qt.ArrowCursor)
            self.update()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        if self.hover != -1:
            self.hover = -1
            self.update()
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton:
            index = self._index_at(event.pos())
            if index >= 0:
                self.sliceClicked.emit(self.labels[index])
        super().mouseReleaseEvent(event)


class LegendList(QWidget):
    """راهنمای نمودار به‌صورت ردیف‌های قابل کلیک."""

    itemClicked = pyqtSignal(str)

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.items: List[Tuple[str, int, int]] = []
        self.palette: List[QColor] = []
        self.hover = -1
        self.max_rows = 12
        self._dark = QColor(colors["surface"]).lightness() < 128
        self.setMouseTracking(True)
        self.setMinimumWidth(110)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def set_theme(self, colors: Dict[str, str], palette: Optional[Sequence[str]] = None) -> None:
        self.colors = colors
        self._dark = QColor(colors["surface"]).lightness() < 128
        if palette:
            self.palette = [QColor(c) for c in palette]
        self.update()

    def set_data(self, labels: Sequence[str], values: Sequence[int],
                 palette: Sequence[str]) -> None:
        total = sum(int(v) for v in values)
        self.items = [(str(label), int(value), int(round(value / total * 100)) if total else 0)
                      for label, value in zip(labels, values)]
        self.palette = [QColor(c) for c in palette]
        self.hover = -1
        self.update()

    def _color(self, index: int) -> QColor:
        if not self.palette:
            return QColor(self.colors["brand"])
        color = QColor(self.palette[index % len(self.palette)])
        if self._dark:
            color = color.lighter(124)
        cycle = index // len(self.palette)
        return color.lighter(112 + 8 * cycle) if cycle else color

    def _row_height(self) -> int:
        return 22

    def _rows(self) -> List[Tuple[str, int, int]]:
        fits = max(1, int((self.height() - 6) // self._row_height()))
        limit = min(self.max_rows, fits)
        if limit < len(self.items):
            limit = max(1, limit - 1)
        return self.items[: max(1, limit)]

    def sizeHint(self) -> QSize:  # noqa: N802
        return QSize(150, 8 + max(1, len(self._rows())) * self._row_height())

    def minimumSizeHint(self) -> QSize:  # noqa: N802
        return QSize(110, 40)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        c = self.colors
        font = app_font(8, QFont.DemiBold)
        painter.setFont(font)
        metrics = QFontMetrics(font)
        rows = self._rows()
        row_height = self._row_height()
        width = self.width()
        top = 3
        percent_texts = [fa(f"{percent}٪") for _l, _v, percent in rows]
        percent_width = max([metrics.horizontalAdvance(t) for t in percent_texts] or [0]) + 4

        for index, (label, value, percent) in enumerate(rows):
            y = top + index * row_height
            if index == self.hover:
                painter.setPen(Qt.NoPen)
                painter.setBrush(alpha(c["brand"], 16))
                painter.drawRoundedRect(QRectF(0, y + 1, width, row_height - 3), 7, 7)
            # نشانگر رنگی
            marker = QRectF(width - 13, y + row_height / 2 - 4, 8, 8)
            painter.setPen(Qt.NoPen)
            painter.setBrush(self._color(index))
            painter.drawRoundedRect(marker, 2.5, 2.5)
            painter.setPen(QPen(QColor(c["ink_2"])))
            label_left = percent_width + 10
            label_width = max(16, int(width - label_left - 20))
            painter.drawText(
                QRectF(label_left, y, label_width, row_height),
                Qt.AlignRight | Qt.AlignVCenter | Qt.AlignAbsolute,
                metrics.elidedText(pretty_label(label).replace(" درصد", "٪"),
                                   Qt.ElideRight, label_width),
            )
            painter.setPen(QPen(QColor(c["ink_3"])))
            painter.drawText(QRectF(0, y, percent_width, row_height),
                             Qt.AlignLeft | Qt.AlignVCenter | Qt.AlignAbsolute,
                             percent_texts[index])

        if len(self.items) > len(rows):
            painter.setPen(QPen(QColor(c["ink_3"])))
            painter.drawText(
                QRectF(0, top + len(rows) * row_height, width, row_height),
                Qt.AlignCenter, f"و {fa(len(self.items) - len(rows))} مورد دیگر",
            )

    def _row_at(self, y: int) -> int:
        row = (y - 3) // self._row_height()
        return int(row) if 0 <= row < len(self._rows()) else -1

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        row = self._row_at(event.pos().y())
        if row != self.hover:
            self.hover = row
            self.setCursor(Qt.PointingHandCursor if row >= 0 else Qt.ArrowCursor)
            if 0 <= row < len(self._rows()):
                label, value, percent = self._rows()[row]
                self.setToolTip(f"{pretty_label(label)}\n{fa_num(value)} قرارداد ({fa(percent)}٪)")
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


class BarChart(QWidget):
    """نمودار میله‌ای افقی با میله‌های گرد و راهنمای مقدار."""

    barClicked = pyqtSignal(str)

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.items: List[Dict[str, Any]] = []
        self.hover = -1
        self.setMouseTracking(True)
        self.setMinimumHeight(120)

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def set_data(self, items: Sequence[Dict[str, Any]]) -> None:
        self.items = list(items)
        self.hover = -1
        self.update()

    def _row_height(self) -> int:
        if not self.items:
            return 24
        available = max(60, self.height() - 6)
        return int(max(12, min(28, available / len(self.items))))

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        c = self.colors
        font = app_font(8, QFont.DemiBold)
        painter.setFont(font)
        metrics = QFontMetrics(font)
        if not self.items:
            painter.setPen(QPen(QColor(c["ink_3"])))
            painter.drawText(self.rect(), Qt.AlignCenter, "داده‌ای برای نمایش نیست")
            return

        row_height = self._row_height()
        max_count = max(1, max(int(item["count"]) for item in self.items))
        value_width = metrics.horizontalAdvance(bidi("۱۰۰٪")) + 12
        texts = [f"{item['label']} ({fa_num(item['count'])})" for item in self.items]
        longest = max(metrics.horizontalAdvance(text) for text in texts)
        label_width = int(min(longest + 10, max(120, self.width() * 0.42),
                              max(90, self.width() - value_width - 120)))
        track_right = self.width() - label_width - 12
        track_left = value_width + 12
        track_width = max(40, track_right - track_left)
        top = max(2.0, (self.height() - row_height * len(self.items)) / 2)

        for index, item in enumerate(self.items):
            y = top + index * row_height
            bar_height = max(9, row_height * 0.36)
            bar_y = y + (row_height - bar_height) / 2
            ratio = item["count"] / (max_count * 1.04)
            width = max(6.0, track_width * ratio)
            hovered = index == self.hover

            painter.setPen(Qt.NoPen)
            painter.setBrush(alpha(c["border"], 120))
            painter.drawRoundedRect(QRectF(track_right - track_width, bar_y, track_width, bar_height),
                                    bar_height / 2, bar_height / 2)
            gradient = QLinearGradient(track_right - width, 0, track_right, 0)
            if hovered:
                gradient.setColorAt(0, QColor(c["brand_2"]))
                gradient.setColorAt(1, QColor(c["brand_deep"]))
            else:
                gradient.setColorAt(0, QColor(c["brand_2"]))
                gradient.setColorAt(1, QColor(c["brand"]))
            painter.setBrush(gradient)
            painter.drawRoundedRect(QRectF(track_right - width, bar_y, width, bar_height),
                                    bar_height / 2, bar_height / 2)

            painter.setPen(QPen(QColor(c["ink_2"])))
            painter.drawText(
                QRectF(track_right + 10, y, label_width, row_height),
                Qt.AlignRight | Qt.AlignVCenter | Qt.AlignAbsolute,
                metrics.elidedText(
                    f"{item['label']} ({fa_num(item['count'])})", Qt.ElideRight, label_width,
                ),
            )
            value_text = bidi(f"{fa(item['percent'])}٪")
            painter.setPen(QPen(QColor(c["ink_3"])))
            painter.drawText(QRectF(0, y, value_width, row_height),
                             Qt.AlignLeft | Qt.AlignVCenter | Qt.AlignAbsolute, value_text)

    def _row_at(self, y: int) -> int:
        if not self.items:
            return -1
        row_height = self._row_height()
        top = max(2.0, (self.height() - row_height * len(self.items)) / 2)
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


class MiniBar(QWidget):
    """نوار ظریف درصد در کارت‌های آماری."""

    def __init__(self, colors: Dict[str, str], color: Optional[str] = None, parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.color = QColor(color) if color else QColor(colors["brand"])
        self.value = 0.0
        self.setFixedHeight(6)
        self.setMinimumWidth(40)

    def set_theme(self, colors: Dict[str, str], color: Optional[str] = None) -> None:
        self.colors = colors
        if color:
            self.color = QColor(color)
        self.update()

    def set_value(self, value: float) -> None:
        self.value = max(0.0, min(1.0, float(value)))
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        body = QRectF(self.rect())
        painter.setPen(Qt.NoPen)
        painter.setBrush(alpha(self.colors["ink_3"], 40))
        painter.drawRoundedRect(body, 3, 3)
        if self.value > 0:
            width = max(6.0, body.width() * self.value)
            gradient = QLinearGradient(body.right() - width, 0, body.right(), 0)
            gradient.setColorAt(0, self.color.lighter(125))
            gradient.setColorAt(1, self.color)
            painter.setBrush(gradient)
            painter.drawRoundedRect(QRectF(body.right() - width, body.top(), width, body.height()), 3, 3)


class LineSpark(QWidget):
    """خط کوچک روند برای کارت‌های شاخص (تزئینی اما مبتنی بر داده)."""

    def __init__(self, colors: Dict[str, str], color: Optional[str] = None, parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.color = QColor(color) if color else QColor(colors["brand_2"])
        self.points: List[float] = []
        self.setFixedHeight(30)
        self.setMinimumWidth(60)

    def set_theme(self, colors: Dict[str, str], color: Optional[str] = None) -> None:
        self.colors = colors
        if color:
            self.color = QColor(color)
        self.update()

    def set_values(self, values: Sequence[float]) -> None:
        self.points = [float(v) for v in values]
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        if len(self.points) < 2:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        body = QRectF(self.rect()).adjusted(1, 2, -1, -2)
        low, high = min(self.points), max(self.points)
        span = high - low if high > low else 1.0
        step = body.width() / (len(self.points) - 1)
        path = QPainterPath()
        coords = []
        for index, value in enumerate(self.points):
            x = body.right() - index * step
            y = body.bottom() - (value - low) / span * body.height()
            coords.append(QPointF(x, y))
        path.moveTo(coords[0])
        for point in coords[1:]:
            path.lineTo(point)
        fill = QPainterPath(path)
        fill.lineTo(QPointF(coords[-1].x(), body.bottom()))
        fill.lineTo(QPointF(coords[0].x(), body.bottom()))
        fill.closeSubpath()
        gradient = QLinearGradient(0, body.top(), 0, body.bottom())
        gradient.setColorAt(0, alpha(self.color.name(), 70))
        gradient.setColorAt(1, alpha(self.color.name(), 0))
        painter.setPen(Qt.NoPen)
        painter.setBrush(gradient)
        painter.drawPath(fill)
        painter.setBrush(Qt.NoBrush)
        pen = QPen(self.color, 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        painter.setPen(pen)
        painter.drawPath(path)



# ==========================================================================
# مدل و نمای جدول قراردادها
# ==========================================================================
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


def _mix(fg_hex: str, bg_hex: str, ratio: float) -> QColor:
    """ترکیب دو رنگ برای ساخت پس‌زمینهٔ نرم چیپ در پوستهٔ تیره."""
    fg = QColor(fg_hex)
    bg = QColor(bg_hex)
    ratio = max(0.0, min(1.0, ratio))
    return QColor(
        int(fg.red() * ratio + bg.red() * (1 - ratio)),
        int(fg.green() * ratio + bg.green() * (1 - ratio)),
        int(fg.blue() * ratio + bg.blue() * (1 - ratio)),
    )


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
        ColumnSpec("ردیف", 42, lambda r, i: str(i + 1), int(Qt.AlignCenter), "number"),
        ColumnSpec("نام شرکت", 146, lambda r, i: r.col2),
        ColumnSpec("شماره قرارداد", 130, lambda r, i: r.col3, int(Qt.AlignCenter), "code"),
        ColumnSpec("موضوع قرارداد", 134, lambda r, i: r.col4),
        ColumnSpec("نوع قرارداد", 100, lambda r, i: r.f_value, int(Qt.AlignCenter), "chip:type"),
        ColumnSpec("موارد باز", 62, lambda r, i: str(len(r.open_flags)), int(Qt.AlignCenter), "flags"),
        ColumnSpec("کسورات اسناد اختتام", 150, lambda r, i: r.deductions),
        ColumnSpec("پیشرفت کارکرد", 104, lambda r, i: r.progress_status, int(Qt.AlignCenter), "chip:progress"),
        ColumnSpec("الحاقیه", 134, lambda r, i: r.annex_status, int(Qt.AlignCenter), "chip:annex"),
        ColumnSpec("وضعیت پکیج", 110, lambda r, i: r.as_status, int(Qt.AlignCenter), "chip:status"),
        ColumnSpec("فایل", 42, lambda r, i: "PDF" if r.pdf_file else "", int(Qt.AlignCenter), "pdf"),
        ColumnSpec("نشان", 26, lambda r, i: "", int(Qt.AlignCenter), "bookmark"),
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

    @property
    def filtered_total(self) -> int:
        """سازگاری با کد قدیمی: تعداد ردیف‌های نمای فعلی."""
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
        if key == "pdf":
            return bool(record.pdf_file)
        if key == "needs_follow":
            return bool(record.open_flags) or record.is_expired or record.has_deductions
        if key in FLAG_LABELS:
            return bool(record.flags.get(key))
        if ":" in key:
            field, _, value = key.partition(":")
            if field == "type":
                return record.f_value == value

            if field == "status":
                return record.as_status == value
            def same(raw: str) -> bool:
                if not raw:
                    return False
                if raw == value:
                    return True
                pretty = pretty_label(raw)
                if pretty == value:
                    return True
                compact_a = pretty.replace(" درصد", "٪")
                compact_b = str(value).replace(" درصد", "٪")
                return compact_a == compact_b

            if field == "progress":
                return same(record.progress_status)
            if field == "completion":
                return same(record.completion_status)
            if field == "annex":
                return same(record.annex_status)
            if field == "docs":
                try:
                    wanted = int(str(value).strip())
                except (TypeError, ValueError):
                    return True
                done = sum(1 for flag in FLAG_LABELS if record.flags.get(flag))
                return done == wanted
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
            dark = QColor(colors["surface"]).lightness() < 128
            chip_bg = _mix(colors["danger"], colors["surface"], 0.22) if dark \
                else QColor(colors["danger_soft"])
            chip_fg = QColor(colors["danger"])
            if dark:
                chip_fg = chip_fg.lighter(132)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QBrush(chip_bg))
            painter.drawRoundedRect(chip, 6, 6)
            painter.setPen(QPen(chip_fg))
            painter.drawText(chip, Qt.AlignCenter, "PDF")
            painter.restore()
            return

        if role == "flags":
            count = int(index.data(Qt.UserRole + 2) or 0)
            label = text
            dark = QColor(colors["surface"]).lightness() < 128
            if count:
                if dark:
                    fg = QColor(colors["warning"]).lighter(135)
                    bg = _mix(colors["warning"], colors["surface"], 0.24)
                else:
                    fg, bg = QColor("#b45309"), QColor(colors["warning_soft"])
            else:
                fg = QColor(colors["text_muted"])
                bg = _mix(colors["text_muted"], colors["surface"], 0.12) if dark \
                    else QColor(colors["surface_soft"])
            self._draw_chip(painter, rect, label, fg, bg, font)
            painter.restore()
            return

        if role == "code":
            painter.setFont(font)
            painter.setPen(QPen(QColor(colors["text"])))
            text_rect = rect.adjusted(8, 0, -8, 0)
            painter.drawText(text_rect, Qt.AlignCenter,
                             metrics.elidedText(f"\u202a{text}\u202c", Qt.ElideRight,
                                                max(10, text_rect.width())))
            painter.restore()
            return

        if role.startswith("chip:") and text:
            style_map = CHIP_STYLES.get(chip_kind, {})
            fg_hex, bg_hex = style_map.get(text, ("", ""))
            if fg_hex:
                fg, bg = QColor(fg_hex), QColor(bg_hex)
                if QColor(colors["surface"]).lightness() < 128:
                    fg = QColor(fg_hex).lighter(145)
                    bg = _mix(fg_hex, colors["surface"], 0.24)
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
        label = metrics.elidedText(pretty_label(text).replace(" درصد", "٪"),
                                   Qt.ElideRight, max_width - 18)
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



# ==========================================================================
# بارگذاری داده و خروجی اکسل
# ==========================================================================
# ---------------------------------------------------------------------------
# ابزارهای داده و خروجی (بارگذاری، خروجی اکسل و گزارش)
# ---------------------------------------------------------------------------
class DataLoader(QThread):
    """خواندن کتاب کار از دیتابیس مرکزی در رشتهٔ جداگانه."""

    loaded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, settings: Dict[str, Any], pdf_folder: str, parent=None,
                 local_path: str = "") -> None:
        super().__init__(parent)
        self.settings = dict(settings)
        self.pdf_folder = pdf_folder
        self.local_path = str(local_path or "")

    def run(self) -> None:
        if self.local_path:
            log_event("DataLoader: شروع خواندن فایل اکسل محلی (بدون دیتابیس)")
        else:
            log_event("DataLoader: شروع خواندن از دیتابیس")
        try:
            if self.local_path:
                data = read_local_workbook(
                    self.local_path,
                    sheet_name=str(self.settings.get("sheet_name") or SHEET_NAME_DEFAULT),
                    pdf_folder=self.pdf_folder,
                )
            else:
                mark_db_attempt()
                data = read_database_workbook(self.settings, self.pdf_folder)
                clear_db_attempt()
        except ExcelDataError as exc:
            clear_db_attempt()
            log_event(f"DataLoader: خطای داده — {exc}")
            self.failed.emit(str(exc))
            return
        except Exception as exc:  # خطای پیش‌بینی‌نشده
            clear_db_attempt()
            detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            write_error_log(traceback.format_exc())
            log_event(f"DataLoader: خطای غیرمنتظره — {detail}")
            self.failed.emit(f"خطای پیش‌بینی‌نشده در خواندن دیتابیس:\n{detail}")
            return
        log_event(f"DataLoader: {data.total_contracts} قرارداد آماده شد")
        self.loaded.emit(data)


# سازگاری با نام قدیمی
ExcelLoader = DataLoader


class SystemUpdatePlanWorker(QThread):
    """خواندن فایل سامانه و ساخت نقشهٔ به‌روزرسانی (بدون نوشتن)."""

    planned = pyqtSignal(object, object, str)
    failed = pyqtSignal(str)

    def __init__(self, settings: Dict[str, Any], source_path: str,
                 sheet_name: str = SHEET_NAME_DEFAULT, parent=None) -> None:
        super().__init__(parent)
        self.settings = dict(settings)
        self.source_path = str(source_path)
        self.sheet_name = sheet_name

    def run(self) -> None:
        try:
            base = resolve_update_base(self.settings, self.source_path)
            blob = base["blob"]
            plan = plan_system_update(self.source_path, blob, sheet_name=self.sheet_name)
            self.planned.emit(plan, blob, str(base.get("filename") or "workbook.xlsx"))
        except ExcelDataError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            write_error_log(traceback.format_exc())
            self.failed.emit(f"خطای پیش‌بینی‌نشده در خواندن فایل سامانه:\n{exc}")


class LocalSystemUpdateWorker(QThread):
    """به‌روزرسانی فایل اکسل محلی با فایل سامانه (بدون دیتابیس)."""

    done = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, source_path: str, target_path: str,
                 sheet_name: str = SHEET_NAME_DEFAULT, parent=None) -> None:
        super().__init__(parent)
        self.source_path = str(source_path)
        self.target_path = str(target_path)
        self.sheet_name = sheet_name

    def run(self) -> None:
        try:
            result = update_local_from_system(
                self.source_path, self.target_path, sheet_name=self.sheet_name
            )
            self.done.emit(result)
        except ExcelDataError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            write_error_log(traceback.format_exc())
            self.failed.emit(f"خطای پیش‌بینی‌نشده در به‌روزرسانی محلی:\n{exc}")


class ManagedUpdateWorker(QThread):
    """کتاب کار داخلی برنامه را از فایل ۰۱ می‌سازد/به‌روزرسانی می‌کند (بدون فایل ۰۲)."""

    done = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, source_path: str, sheet_name: str = SHEET_NAME_DEFAULT,
                 parent=None, seed_path: str = "") -> None:
        super().__init__(parent)
        self.source_path = str(source_path)
        self.sheet_name = sheet_name
        self.seed_path = str(seed_path or "")

    def run(self) -> None:
        try:
            result = update_managed_workbook(self.source_path, sheet_name=self.sheet_name,
                                             seed_path=self.seed_path)
            self.done.emit(result)
        except ExcelDataError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            write_error_log(traceback.format_exc())
            self.failed.emit(f"خطای پیش‌بینی‌نشده در ساخت کتاب کار داخلی:\n{exc}")


class SystemCopyWorker(QThread):
    """نسخهٔ به‌روزشدهٔ کتاب کار را از فایل ۰۱ می‌سازد (بدون دست‌زدن به فایل اصلی)."""

    done = pyqtSignal(dict)
    failed = pyqtSignal(str)
    sheet_needed = pyqtSignal(dict)

    def __init__(self, settings: Dict[str, Any], source_path: str, parent=None,
                 forced_sheet: str = "", auto_import: bool = True) -> None:
        super().__init__(parent)
        self.settings = dict(settings)
        self.source_path = str(source_path)
        self.forced_sheet = str(forced_sheet or "")
        self.auto_import = bool(auto_import)

    def run(self) -> None:
        try:
            log_event(f"فایل ۰۱ انتخاب‌شده: {self.source_path}")
            check = validate_system_source(self.source_path)
            log_event(
                "فایل ۰۱ سالم است: شیت «{sheet}» — {rows} ردیف، {matched} ردیف "
                "«اجرا - مهر»، ستون زیرپروژه {sub_col}، ستون قرارداد {contract_col}".format(
                    sheet=check["sheet"], rows=fa(check["rows"]),
                    matched=fa(check["matched"]), sub_col=fa(check["sub_col"]),
                    contract_col=fa(check["contract_col"]),
                )
            )
            base = resolve_update_base(self.settings, self.source_path)
            if check.get("path") and str(base.get("base_path") or "") == str(check["path"]):
                raise ExcelDataError(
                    "فایل «۰۱» و کتاب کار پایه یکی هستند؛ چیزی برای به‌روزرسانی نیست.\n"
                    "فایل خروجی سامانه (Contract-Info-Report) را انتخاب کنید."
                )
            keys = source_contract_keys(self.source_path)
            report: List[Dict[str, Any]] = []
            try:
                probe = openpyxl.load_workbook(io.BytesIO(base["blob"]), data_only=False)
                try:
                    if self.forced_sheet and self.forced_sheet in probe.sheetnames:
                        target_sheet = self.forced_sheet
                    else:
                        target_sheet = detect_system_sheet(probe, keys)
                    if not target_sheet:
                        report = sheet_scan_report(probe, keys)
                finally:
                    probe.close()
            except Exception as exc:
                raise ExcelDataError(f"کتاب کار باز نشد:\n{exc}") from exc
            if not target_sheet:
                for item in report:
                    log_event(
                        "بررسی شیت «{sheet}»: امتیاز {score} | سرستون ردیف {header_row} | "
                        "ستون زیرپروژه {sub_col} | ستون قرارداد {contract_col} | "
                        "هم‌خوانی {hits} از {rows} ردیف".format(**item)
                    )
                self.sheet_needed.emit({
                    "base_source": base.get("base_source"),
                    "base_filename": base.get("filename"),
                    "sheets": [item["sheet"] for item in report],
                    "report": report,
                })
                return
            plan = plan_system_update(self.source_path, base["blob"],
                                      sheet_name=target_sheet)
            if not plan.matched_rows:
                raise ExcelDataError(
                    "در شیت «" + str(target_sheet) + "» هیچ ردیفی با فایل ۰۱ تطبیق "
                    "پیدا نکرد (هیچ شمارهٔ قرارداد یا نام شرکتی مشترک نبود).\n\n"
                    "یعنی این شیت همان جایی نیست که خروجی سامانه در آن پیست می‌شود. "
                    "شیت مقصد را عوض کنید یا کتاب کار اصلی (دارای شیت colorcode) را "
                    "بارگذاری کنید."
                )
            stamp = datetime.now().strftime("%Y%m%d-%H%M")
            stem = Path(str(base.get("filename") or "workbook.xlsx")).stem
            out_path = data_dir() / f"{stem}-به‌روز-{stamp}.xlsx"
            info = prepare_system_copy(plan, base["blob"], str(out_path))
            log_event(
                f"نسخهٔ به‌روز ساخته شد: {info['output_path']} — "
                f"{plan.matched_rows} ردیف، {info['written_cells']} سلول، "
                f"{info['formula_cells']} سلول فرمولی دست‌نخورده"
            )
            recalculated = False
            healthy = False
            imported: Dict[str, Any] = {}
            if self.auto_import:
                recalculated = recalculate_with_excel(info["output_path"])
                try:
                    health = workbook_health(Path(info["output_path"]).read_bytes())
                    healthy = not health.get("damaged")
                except Exception:
                    healthy = False
                log_event(
                    f"نسخهٔ به‌روز: محاسبه با اکسل={recalculated}، سالم={healthy}"
                )
                if recalculated and healthy and str(self.settings.get("host") or "").strip():
                    try:
                        target_id = next_database_workbook_id(self.settings)
                        imported = import_workbook_to_database(
                            self.settings, info["output_path"],
                            sheet_name=str(self.settings.get("sheet_name") or ""),
                            workbook_id=target_id,
                        )
                        imported["workbook_id"] = target_id
                        imported["base_workbook_id"] = int(
                            base.get("base_workbook_id") or 0
                        )
                        log_event(
                            "نسخهٔ به‌روز خودکار در دیتابیس ذخیره شد: شناسهٔ "
                            f"{target_id} — {imported.get('data_rows')} ردیف "
                            "(کتاب کار پایه دست‌نخورده ماند)"
                        )
                    except Exception as exc:
                        log_event(f"ذخیرهٔ خودکار نسخهٔ به‌روز در دیتابیس ناموفق بود: {exc}")
                        imported = {"error": str(exc)}
            self.done.emit({
                "plan": plan,
                "output_path": info["output_path"],
                "written_cells": info["written_cells"],
                "formula_cells": info["formula_cells"],
                "sheet": target_sheet,
                "base_source": base.get("base_source"),
                "base_filename": base.get("filename"),
                "base_workbook_id": int(base.get("base_workbook_id") or 0),
                "recalculated": recalculated,
                "cleaned_marks": info.get("cleaned_marks", 0),
                "healthy": healthy,
                "imported": imported,
            })
        except ExcelDataError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            write_error_log(traceback.format_exc())
            self.failed.emit(f"خطای پیش‌بینی‌نشده در ساخت نسخهٔ به‌روز:\n{exc}")


class SystemUpdateApplyWorker(QThread):
    """اعمال نقشه روی کتاب کار، گرفتن نسخهٔ پشتیبان و ذخیره در دیتابیس."""

    done = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, settings: Dict[str, Any], plan: SystemUpdatePlan, blob: bytes,
                 filename: str, parent=None) -> None:
        super().__init__(parent)
        self.settings = dict(settings)
        self.plan = plan
        self.blob = blob
        self.filename = filename

    def run(self) -> None:
        try:
            backup = backup_workbook(self.blob, self.filename)
            new_blob, cells = apply_system_update(self.plan, self.blob)
            with tempfile.TemporaryDirectory(prefix="mehr-update-") as temp:
                temp_path = Path(temp) / self.filename
                if temp_path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    temp_path = temp_path.with_suffix(".xlsx")
                temp_path.write_bytes(new_blob)
                imported = import_workbook_to_database(
                    self.settings, str(temp_path), sheet_name=self.plan.sheet_name,
                    workbook_id=int(self.settings.get("workbook_id") or 1),
                )
            log_event(
                f"به‌روزرسانی از فایل سامانه: {self.plan.filtered_rows} ردیف، "
                f"{self.plan.matched_rows} تطبیق، {cells} سلول"
            )
            self.done.emit({
                "plan": self.plan,
                "updated_cells": cells,
                "backup_path": backup,
                "imported": imported,
            })
        except ExcelDataError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            write_error_log(traceback.format_exc())
            self.failed.emit(f"خطای پیش‌بینی‌نشده در به‌روزرسانی:\n{exc}")


class ImportWorker(QThread):
    """بارگذاری فایل اکسل در دیتابیس در رشتهٔ جداگانه (بدون قفل‌شدن پنجره)."""

    succeeded = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, settings: Dict[str, Any], file_path: str,
                 sheet_name: str = SHEET_NAME_DEFAULT, parent=None) -> None:
        super().__init__(parent)
        self.settings = dict(settings)
        self.file_path = file_path
        self.sheet_name = sheet_name

    def run(self) -> None:
        log_event(f"ImportWorker: شروع بارگذاری {self.file_path}")
        try:
            info = import_workbook_to_database(
                self.settings, self.file_path, sheet_name=self.sheet_name
            )
        except ExcelDataError as exc:
            log_event(f"ImportWorker: خطا — {exc}")
            self.failed.emit(str(exc))
            return
        except Exception as exc:  # خطای پیش‌بینی‌نشده
            detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            write_error_log(traceback.format_exc())
            log_event(f"ImportWorker: خطای غیرمنتظره — {detail}")
            self.failed.emit(f"خطای پیش‌بینی‌نشده در بارگذاری اکسل:\n{detail}")
            return
        self.succeeded.emit(info)


# ---------------------------------------------------------------------------
# چیدمان جریانی (برای کارت‌های واکنش‌گرا)
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# خروجی‌های اکسل
# ---------------------------------------------------------------------------
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
    sheet["A2"] = (
        f"تاریخ تهیهٔ گزارش: {today_jalali()}   •   منبع: {data.excel_path}"
        + (f"   •   {data.source_note}" if data.source_note else "")
    )
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



# ==========================================================================


def make_app_icon() -> QIcon:
    """آیکون برنامه بدون نیاز به فایل خارجی."""
    size = 128
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    gradient = QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0, QColor("#18A492"))
    gradient.setColorAt(1, QColor("#0A5A4F"))
    painter.setPen(Qt.NoPen)
    painter.setBrush(QBrush(gradient))
    painter.drawRoundedRect(QRectF(4, 4, size - 8, size - 8), 30, 30)
    draw_icon(painter, "layers", QRectF(30, 32, 68, 64), QColor("#FFFFFF"), 2.0)
    painter.setPen(Qt.NoPen)
    painter.setBrush(QColor("#F2C777"))
    painter.drawEllipse(QPointF(96, 32), 7, 7)
    painter.end()
    return QIcon(pixmap)


def configure_rtl_browser(browser: QTextBrowser) -> None:
    """راست‌به‌چپ کردن متن HTML در QTextBrowser."""
    browser.setLayoutDirection(Qt.RightToLeft)
    browser.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    document = browser.document()
    option = document.defaultTextOption()
    option.setTextDirection(Qt.RightToLeft)
    option.setAlignment(Qt.AlignRight)
    document.setDefaultTextOption(option)


def label_value_pair(title: str, value: str, colors: Dict[str, str]) -> QWidget:
    """یک جفت «عنوان/مقدار» برای پنل جزئیات."""
    widget = QWidget()
    layout = QVBoxLayout(widget)
    layout.setContentsMargins(0, 2, 0, 2)
    layout.setSpacing(1)
    caption = QLabel(title)
    caption.setStyleSheet(f"color: {colors['ink_3']}; font-size: 7.5pt; font-weight: 700;")
    content = QLabel(value or "—")
    content.setWordWrap(True)
    content.setTextInteractionFlags(Qt.TextSelectableByMouse)
    content.setStyleSheet("font-weight: 700; font-size: 9pt;")
    layout.addWidget(caption)
    layout.addWidget(content)
    return widget


# ---------------------------------------------------------------------------
# ریل ناوبری
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


def run_check(settings: Dict[str, Any], pdf_folder: str) -> int:
    """بررسی سریع اتصال به دیتابیس و داده‌ها بدون باز کردن پنجره."""
    try:
        data = read_database_workbook(settings, pdf_folder)
    except ExcelDataError as exc:
        print(f"❌ {exc}")
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"❌ خطای غیرمنتظره: {exc}")
        write_error_log(traceback.format_exc())
        return 3
    print("✅ خواندن داده از دیتابیس موفق بود")
    print(f"دیتابیس: {settings.get('database')} روی {settings.get('host')}:{settings.get('port')}")
    print(f"کتاب کار: {data.source_name} (شناسهٔ {fa(int(settings.get('workbook_id') or 1))})")
    if data.source_note:
        print(f"منبع: {data.source_note}")
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
        description="داشبورد قراردادهای مهر خلیج فارس — نسخهٔ دسکتاپ PyQt5 (دیتابیس مرکزی)",
    )
    parser.add_argument("--db-host", help="میزبان MySQL (پیش‌فرض: 127.0.0.1)")
    parser.add_argument("--db-port", type=int, help="پورت MySQL (پیش‌فرض: 3306)")
    parser.add_argument("--db-name", help="نام دیتابیس (پیش‌فرض: master_data)")
    parser.add_argument("--db-user", help="نام کاربری دیتابیس (پیش‌فرض: master_app)")
    parser.add_argument("--db-password", help="رمز دیتابیس (در تنظیمات هم قابل ذخیره است)")
    parser.add_argument("--workbook-id", type=int,
                        help="شناسهٔ کتاب کار در دیتابیس (پیش‌فرض: 1)")
    parser.add_argument("--sheet", help=f"نام شیت (پیش‌فرض: {SHEET_NAME_DEFAULT})")
    parser.add_argument("--pdf-folder", help="مسیر پوشهٔ فایل‌های PDF")
    parser.add_argument("--local-excel", metavar="PATH",
                        help="اجرای برنامه با فایل اکسل محلی، بدون دیتابیس")
    parser.add_argument("--import-excel", metavar="PATH",
                        help="بارگذاری فایل اکسل در دیتابیس (ساخت دیتابیس و جدول‌ها از ابتدا) و خروج")
    parser.add_argument("--update-excel", metavar="PATH",
                        help="به‌روزرسانی کتاب کار دیتابیس از فایل سامانه (فایل ۰۱) و خروج")
    parser.add_argument("--check", action="store_true",
                        help="فقط بررسی اتصال و خواندن داده‌ها؛ پنجره باز نمی‌شود")
    parser.add_argument("--theme", choices=["light", "dark"], help="پوستهٔ برنامه")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {APP_VERSION}")
    return parser.parse_args(argv)


def _apply_args(config: AppConfig, args: argparse.Namespace) -> None:
    mapping = {
        "db_host": args.db_host,
        "db_port": args.db_port,
        "db_name": args.db_name,
        "db_user": args.db_user,
        "db_workbook_id": args.workbook_id,
        "sheet_name": args.sheet,
    }
    for key, value in mapping.items():
        if value not in (None, ""):
            config.values[key] = value
    if args.db_password:
        config.values["db_password"] = args.db_password
    if args.pdf_folder:
        config.values["pdf_folder"] = os.path.abspath(args.pdf_folder)
    if args.local_excel:
        config.values["data_source"] = "local"
        config.values["local_excel_path"] = os.path.abspath(args.local_excel)
    if args.theme:
        config.values["theme"] = args.theme


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    install_crash_logging()
    log_event(
        f"شروع برنامه — نسخهٔ {APP_VERSION} | python {sys.version.split()[0]} | "
        f"{platform.platform()}"
    )
    log_event(
        "mysql-connector: " + ("نصب است" if mysql_connector is not None else "نصب نیست")
        + (f" | موتور: {connector_engine_label()}" if mysql_connector is not None else "")
    )

    if args.check:
        config = AppConfig()
        _apply_args(config, args)
        settings = config.db_settings()
        settings["sheet_name"] = str(config.get("sheet_name") or SHEET_NAME_DEFAULT)
        pdf_folder = str(config.get("pdf_folder") or default_pdf_folder(config))
        return run_check(settings, pdf_folder)

    if args.import_excel:
        config = AppConfig()
        _apply_args(config, args)
        try:
            info = import_workbook_to_database(
                config.db_settings(),
                args.import_excel,
                sheet_name=str(config.get("sheet_name") or SHEET_NAME_DEFAULT),
            )
        except ExcelDataError as exc:
            print(f"خطا: {exc}")
            return 1
        config.save()
        print("فایل اکسل با موفقیت در دیتابیس ذخیره شد.")
        print(f"دیتابیس: {info['database']} — شناسهٔ کتاب کار: {fa(info['workbook_id'])}")
        print(f"فایل: {info['filename']} ({fa(info['size'])} بایت)")
        print(f"شیت هدف: {info['sheet_name']}")
        print("شیت‌های موجود: " + "، ".join(info["sheets"]))
        return 0

    if args.update_excel:
        config = AppConfig()
        _apply_args(config, args)
        sheet = str(config.get("sheet_name") or SHEET_NAME_DEFAULT)
        try:
            result = update_database_from_system(
                config.db_settings(), args.update_excel, sheet_name=sheet,
            )
            target_label = "کتاب کار دیتابیس"
        except ExcelDataError as exc:
            print(f"دیتابیس در دسترس نیست ({exc}).")
            print("کتاب کار داخلی برنامه به‌روزرسانی می‌شود…")
            try:
                result = update_managed_workbook(
                    args.update_excel, sheet_name=sheet,
                    seed_path=str(config.get("local_excel_path") or ""),
                )
            except ExcelDataError as inner:
                print(f"خطا: {inner}")
                return 1
            target_label = "کتاب کار داخلی"
        config.save()
        print(str(result.get("summary") or ""))
        print(f"سلول‌های به‌روزرسانی‌شده: {fa(result['updated_cells'])}")
        if result.get("backup_path"):
            print(f"نسخهٔ پشتیبان: {result['backup_path']}")
        print(f"{target_label} با موفقیت به‌روزرسانی شد: {result.get('output_path') or ''}")
        return 0

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
    _apply_args(config, args)
    notes = NotesStore()

    window = MainWindow(config, notes)
    window.show()
    return app.exec_()



# ==========================================================================
# اجزای مشترک رابط
# ==========================================================================
# ---------------------------------------------------------------------------
# اجزای مشترک رابط (کاشی آمار، کارت نمودار، دکمه‌ها و ورودی‌ها)
# ---------------------------------------------------------------------------
def icon_pixmap(kind: str, color: str, size: int = 16, width: float = 1.7) -> QPixmap:
    """ساخت آیکون به‌صورت پیکس‌مپ برای استفاده در دکمه‌ها و منوها."""
    ratio = 2
    pixmap = QPixmap(size * ratio, size * ratio)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    draw_icon(painter, kind, QRectF(0, 0, size * ratio, size * ratio), QColor(color), width)
    painter.end()
    pixmap.setDevicePixelRatio(ratio)
    return pixmap


class IconButton(QPushButton):
    """دکمهٔ متنی/آیکونی با پشتیبانی از تغییر رنگ در تم تاریک."""

    def __init__(self, text: str = "", kind: str = "", icon_color: str = "", parent=None) -> None:
        super().__init__(text, parent)
        self._kind = kind
        self._icon_color = icon_color
        if kind:
            self._apply_icon()

    def set_icon_kind(self, kind: str, color: Optional[str] = None) -> None:
        self._kind = kind
        if color:
            self._icon_color = color
        self._apply_icon()

    def set_icon_color(self, color: str) -> None:
        self._icon_color = color
        self._apply_icon()

    def _apply_icon(self) -> None:
        if self._kind:
            self.setIcon(QIcon(icon_pixmap(self._kind, self._icon_color or "#0E7A6A", 16)))
            self.setIconSize(QSize(16, 16))


def make_icon(kind: str, color: str, size: int = 18) -> QIcon:
    return QIcon(icon_pixmap(kind, color, size))


class SearchBox(QLineEdit):
    """کادر جستجو با آیکون ذره‌بین و دکمهٔ پاک‌کردن."""

    def __init__(self, placeholder: str = "جستجو…", parent=None) -> None:
        super().__init__(parent)
        self.setPlaceholderText(placeholder)
        self.setClearButtonEnabled(True)
        self.setMinimumHeight(38)
        self._color = "#7A8C86"
        self._sync_action()

    def set_icon_color(self, color: str) -> None:
        self._color = color
        self._sync_action()

    def _sync_action(self) -> None:
        icon = QIcon(icon_pixmap("search", self._color, 15, 1.8))
        action = self.actions()[0] if self.actions() else None
        if action is None:
            self.addAction(icon, QLineEdit.LeadingPosition)
        else:
            action.setIcon(icon)


class PillTabs(QWidget):
    """نوار تب‌های قرصی‌شکل (فیلترها) با امکان انتخاب یکی."""

    changed = pyqtSignal(str)

    def __init__(self, items: Sequence[Tuple[str, str]], colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self._buttons: Dict[str, QPushButton] = {}
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(6)
        self._current = ""
        for key, label in items:
            button = QPushButton(label)
            button.setObjectName("Chip")
            button.setCheckable(True)
            button.setCursor(Qt.PointingHandCursor)
            button.clicked.connect(lambda _checked=False, k=key: self.select(k))
            self.layout.addWidget(button)
            self._buttons[key] = button
        self.layout.addStretch(1)
        if items:
            self.select(items[0][0])

    def current(self) -> str:
        return self._current

    def select(self, key: str) -> None:
        if key not in self._buttons:
            return
        changed = key != self._current
        self._current = key
        for button_key, button in self._buttons.items():
            button.setChecked(button_key == key)
        if changed:
            self.changed.emit(key)

    def set_current(self, key: str) -> None:
        self.select(key)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors


class StatTile(Card):
    """کاشی آمار: عنوان، مقدار بزرگ، توضیح و نوار نسبت."""

    clicked = pyqtSignal(str)

    def __init__(self, key: str, title: str, kind: str, accent: str,
                 colors: Dict[str, str], parent=None) -> None:
        super().__init__(colors, radius=16, padding=13, hoverable=True, parent=parent)
        self.key = key
        self.accent = accent
        self.title_text = title
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(116)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        header = QHBoxLayout()
        header.setSpacing(6)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("CardTitle")
        self.title_label.setStyleSheet(
            f"font-size: 9pt; font-weight: 700; color: {colors['ink_2']};"
        )
        set_text_aligned(self.title_label, title, "right")
        self.tile = IconTile(kind, accent, alpha(accent, 24).name(QColor.HexArgb),
                             size=27, radius=9)
        header.addWidget(self.tile, 0, Qt.AlignTop)
        header.addWidget(self.title_label, 1)
        self.body.addLayout(header)

        value_row = QHBoxLayout()
        value_row.setSpacing(8)
        self.value_label = QLabel("—")
        self.value_label.setObjectName("CardValue")
        self.value_label.setStyleSheet("font-size: 16pt; font-weight: 800;")
        value_row.addWidget(self.value_label, 1)
        self.badge = QLabel("")
        self.badge.setObjectName("Chip")
        self.badge.setVisible(False)
        value_row.addWidget(self.badge, 0, Qt.AlignBottom)
        self.body.addLayout(value_row)

        self.hint_label = QLabel("")
        self.hint_label.setObjectName("Muted")
        self.hint_label.setStyleSheet("font-size: 8pt;")
        self.body.addWidget(self.hint_label)

        self.body.addStretch(1)
        self.bar = MiniBar(colors, accent)
        self.body.addWidget(self.bar)

    def set_theme(self, colors: Dict[str, str]) -> None:
        super().set_theme(colors)
        self.colors = colors
        self.tile.set_colors(self.accent, alpha(self.accent, 24).name(QColor.HexArgb))
        self.bar.set_theme(colors, self.accent)
        self.title_label.setStyleSheet(
            f"font-size: 9pt; font-weight: 700; color: {colors['ink_2']};"
        )

    def set_values(self, value: str, hint: str = "", ratio: Optional[float] = None,
                   badge: str = "") -> None:
        self.value_label.setText(value)
        set_text_aligned(self.value_label, value, "right")
        self.hint_label.setText(hint)
        set_text_aligned(self.hint_label, hint, "right")
        if ratio is not None:
            self.bar.set_value(ratio)
            self.bar.setVisible(True)
        else:
            self.bar.setVisible(False)
        self.badge.setText(badge)
        self.badge.setVisible(bool(badge))

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit(self.key)
        super().mouseReleaseEvent(event)


FLAG_ICONS = {
    "temporary": "truck",
    "prefinal": "report",
    "balance": "layers",
    "statement": "note",
    "mofasa": "shield",
    "affidavit": "check-plain",
    "delivery": "check",
    "financial": "money",
}


class FlagTile(Card):
    """کارت کوچک و قابل کلیک برای یکی از وضعیت‌های اسناد اختتام."""

    clicked = pyqtSignal(str)

    def __init__(self, key: str, title: str, colors: Dict[str, str], parent=None) -> None:
        super().__init__(colors, radius=14, padding=8, hoverable=True, parent=parent)
        self.key = key
        self.colors = colors
        self.title_text = title
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(60)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        top = QHBoxLayout()
        top.setSpacing(6)
        self.value_label = QLabel("—")
        self.value_label.setStyleSheet(
            f"font-size: 12.5pt; font-weight: 800; color: {colors['ink']};"
        )
        self.tile = IconTile(FLAG_ICONS.get(key, "check"), colors["brand"],
                             colors["brand_soft"], size=22, radius=7)
        top.addWidget(self.tile, 0, Qt.AlignTop)
        top.addWidget(self.value_label, 1)
        self.body.addLayout(top)

        self.title_label = QLabel(title)
        self.title_label.setWordWrap(True)
        self.title_label.setStyleSheet("font-size: 7.5pt; font-weight: 600;")
        set_text_aligned(self.title_label, title, "right")
        self.body.addWidget(self.title_label)
        self.body.addStretch(1)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)
        set_text_aligned(self.value_label, value, "right")

    def set_theme(self, colors: Dict[str, str]) -> None:
        super().set_theme(colors)
        self.colors = colors
        self.tile.set_colors(colors["brand"], colors["brand_soft"])
        self.value_label.setStyleSheet(
            f"font-size: 12.5pt; font-weight: 800; color: {colors['ink']};"
        )
        self.title_label.setStyleSheet(
            f"font-size: 7.5pt; font-weight: 600; color: {colors['ink_2']};"
        )

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit(self.key)
        super().mouseReleaseEvent(event)


class ChartCard(Card):
    """کارت نمودار با سرتیتر، توضیح و ناحیهٔ محتوا."""

    def __init__(self, title: str, subtitle: str, colors: Dict[str, str],
                 min_height: int = 260, parent=None) -> None:
        super().__init__(colors, radius=16, padding=13, parent=parent)
        self.setMinimumHeight(min_height)
        header = QHBoxLayout()
        header.setSpacing(6)
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("SectionTitle")
        set_text_aligned(self.title_label, title, "right")
        self.subtitle_label = QLabel(subtitle)
        self.subtitle_label.setObjectName("Muted")
        set_text_aligned(self.subtitle_label, subtitle, "right")
        title_box.addWidget(self.title_label)
        title_box.addWidget(self.subtitle_label)
        header.addLayout(title_box)
        header.addStretch(1)
        self.actions = QHBoxLayout()
        self.actions.setSpacing(6)
        header.addLayout(self.actions)
        self.body.addLayout(header)

        self.content = QWidget()
        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(0, 1, 0, 0)
        self.content_layout.setSpacing(5)
        self.body.addWidget(self.content, 1)
        self.set_theme_colors(colors)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.set_theme(colors)

    def add_action(self, widget: QWidget) -> None:
        self.actions.addWidget(widget)


class LegendDot(QWidget):
    """نقطهٔ رنگی کوچک برای راهنماها."""

    def __init__(self, color: str, size: int = 9, parent=None) -> None:
        super().__init__(parent)
        self.color = QColor(color)
        self.setFixedSize(size, size)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        painter.setBrush(self.color)
        painter.drawRoundedRect(QRectF(self.rect()), 2.5, 2.5)


class NoticeBar(QFrame):
    """نوار اطلاع‌رسانی ظریف (هشدار یا راهنما)."""

    actionClicked = pyqtSignal()
    secondaryClicked = pyqtSignal()

    def __init__(self, colors: Dict[str, str], tone: str = "warning", parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.tone = tone
        self.setObjectName("NoticeBar")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(10)
        self.icon = IconWidget("alert" if tone == "warning" else "info",
                               colors["warning"] if tone == "warning" else colors["info"], 18)
        layout.addWidget(self.icon)
        self.label = QLabel("")
        self.label.setWordWrap(True)
        self.label.setObjectName("Subtle")
        layout.addWidget(self.label, 1)
        self.secondary_button = QPushButton("")
        self.secondary_button.setObjectName("Ghost")
        self.secondary_button.setCursor(Qt.PointingHandCursor)
        self.secondary_button.clicked.connect(self.secondaryClicked.emit)
        self.secondary_button.setVisible(False)
        layout.addWidget(self.secondary_button)
        self.action_button = QPushButton("")
        self.action_button.setObjectName("Ghost")
        self.action_button.setCursor(Qt.PointingHandCursor)
        self.action_button.clicked.connect(self.actionClicked.emit)
        self.action_button.setVisible(False)
        layout.addWidget(self.action_button)
        self.close_button = IconButton("", "close", colors["ink_3"])
        self.close_button.setFixedSize(28, 28)
        self.close_button.setObjectName("Ghost")
        self.close_button.clicked.connect(self.hide)
        layout.addWidget(self.close_button)
        self.set_theme(colors)

    def set_text(self, text: str, action_text: str = "") -> None:
        self.label.setText(text)
        self.action_button.setText(action_text)
        self.action_button.setVisible(bool(action_text))

    def set_secondary(self, text: str) -> None:
        """دکمهٔ دوم نوار (مثلاً «بارگذاری فایل اکسل»)."""
        self.secondary_button.setText(text)
        self.secondary_button.setVisible(bool(text))

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        fg = colors["warning"] if self.tone == "warning" else colors["info"]
        bg = colors["warning_soft"] if self.tone == "warning" else colors["info_soft"]
        self.setStyleSheet(
            f"QFrame#NoticeBar {{ background: {bg}; border-radius: 14px; }}"
            f" QLabel {{ color: {fg}; }}"
        )
        self.label.setStyleSheet(f"color: {fg};")
        self.icon.set_color(fg)
        self.close_button.set_icon_color(fg)


# ---------------------------------------------------------------------------
# ورودی‌های بومی (بدون وابستگی به تصاویر QSS)
# ---------------------------------------------------------------------------
class Combo(QPushButton):
    """کمبوی سفارشی: دکمه‌ای با منوی کشویی و فلش رسم‌شده.

    ظاهر کاملاً نقاشی می‌شود تا در حالت RTL و پوستهٔ تیره یکدست بماند، اما
    همان رابط QComboBox (addItem/currentData/findData/…) را ارائه می‌دهد.
    """

    currentIndexChanged = pyqtSignal(int)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.colors: Dict[str, str] = THEMES["light"]
        self._items: List[Tuple[str, Any]] = []
        self._index = -1
        self._hover = False
        self.setText("")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(36)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setAttribute(Qt.WA_Hover, True)
        self._menu = QMenu(self)
        self.clicked.connect(self._show_menu)

    # ---- تم ----
    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def set_arrow_color(self, color: str) -> None:  # سازگاری با کد قدیمی
        self.update()

    def enterEvent(self, event) -> None:  # noqa: N802
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._hover = False
        self.update()
        super().leaveEvent(event)

    # ---- API شبیه QComboBox ----
    def addItem(self, text: str, data: Any = None) -> None:  # noqa: N802
        self._items.append((str(text), data))
        if self._index < 0:
            self._index = 0
        self.update()

    def addItems(self, texts: Sequence[str]) -> None:  # noqa: N802
        for text in texts:
            self.addItem(text, None)

    def clear(self) -> None:
        self._items = []
        self._index = -1
        self.update()

    def count(self) -> int:
        return len(self._items)

    def itemText(self, index: int) -> str:  # noqa: N802
        if 0 <= index < len(self._items):
            return self._items[index][0]
        return ""

    def currentIndex(self) -> int:  # noqa: N802
        return self._index

    def setCurrentIndex(self, index: int) -> None:  # noqa: N802
        if not self._items:
            self._index = -1
            self.update()
            return
        index = max(0, min(len(self._items) - 1, int(index)))
        changed = index != self._index
        self._index = index
        self.update()
        if changed:
            self.currentIndexChanged.emit(index)

    def currentText(self) -> str:  # noqa: N802
        return self.itemText(self._index)

    def currentData(self) -> Any:  # noqa: N802
        if 0 <= self._index < len(self._items):
            return self._items[self._index][1]
        return None

    def findData(self, data: Any) -> int:  # noqa: N802
        for index, (_text, value) in enumerate(self._items):
            if value == data:
                return index
        return -1

    def findText(self, text: str) -> int:  # noqa: N802
        for index, (label, _value) in enumerate(self._items):
            if label == text:
                return index
        return -1

    def _show_menu(self) -> None:
        self._menu.clear()
        for index, (text, _data) in enumerate(self._items):
            action = self._menu.addAction(text)
            action.setCheckable(True)
            action.setChecked(index == self._index)
            action.triggered.connect(lambda _checked=False, i=index: self.setCurrentIndex(i))
        if self._items:
            point = self.mapToGlobal(self.rect().bottomLeft())
            self._menu.setMinimumWidth(self.width())
            self._menu.exec_(point)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        c = self.colors
        body = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        enabled = self.isEnabled()
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(c["surface_2"] if not enabled else c["surface"]))
        painter.drawRoundedRect(body, 11, 11)
        border = QColor(c["brand_2"]) if (self._hover or self.hasFocus()) and enabled else QColor(c["border"])
        painter.setPen(QPen(border, 1))
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(body, 11, 11)

        arrow_color = QColor(c["brand_deep"] if self._hover else c["ink_3"])
        if not enabled:
            arrow_color = QColor(c["ink_3"])
        painter.setPen(QPen(arrow_color, 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        cx = body.left() + 16
        cy = body.center().y()
        path = QPainterPath()
        path.moveTo(cx - 4.2, cy - 2.0)
        path.lineTo(cx, cy + 2.4)
        path.lineTo(cx + 4.2, cy - 2.0)
        painter.drawPath(path)

        font = app_font(9, QFont.DemiBold)
        painter.setFont(font)
        painter.setPen(QPen(QColor(c["ink"] if enabled else c["ink_3"])))
        text_rect = body.adjusted(30, 0, -12, 0)
        metrics = QFontMetrics(font)
        text = metrics.elidedText(self.currentText(), Qt.ElideRight, int(text_rect.width()))
        painter.drawText(text_rect, Qt.AlignRight | Qt.AlignVCenter | Qt.AlignAbsolute, text)


class CheckBox(QCheckBox):
    """چک‌باکس با نشانگر سفارشیِ هماهنگ با پوسته (بدون وابستگی به استایل بومی)."""

    BOX = 18

    def __init__(self, text: str = "", parent=None) -> None:
        super().__init__(text, parent)
        self.colors = THEMES["light"]
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(26)
        self.setAttribute(Qt.WA_Hover, True)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def sizeHint(self) -> QSize:  # noqa: N802
        width = self.fontMetrics().horizontalAdvance(self.text())
        return QSize(width + self.BOX + 34, 26)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        c = self.colors
        box = QRectF(self.width() - self.BOX - 2,
                     (self.height() - self.BOX) / 2, self.BOX, self.BOX)
        checked = self.isChecked()
        enabled = self.isEnabled()
        hovered = self.underMouse()
        if checked:
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(c["brand"] if enabled else c["ink_3"]))
            painter.drawRoundedRect(box, 5.5, 5.5)
            painter.setPen(QPen(QColor(c["on_brand"]), 2.0, Qt.SolidLine,
                                Qt.RoundCap, Qt.RoundJoin))
            painter.setBrush(Qt.NoBrush)
            path = QPainterPath()
            path.moveTo(box.left() + 4.7, box.center().y() + 0.5)
            path.lineTo(box.center().x() - 0.7, box.bottom() - 4.6)
            path.lineTo(box.right() - 4.3, box.top() + 4.8)
            painter.drawPath(path)
        else:
            border = c["brand_2"] if (hovered and enabled) else c["border"]
            painter.setPen(QPen(QColor(border), 1.4))
            painter.setBrush(QColor(c["surface"]))
            painter.drawRoundedRect(box, 5.5, 5.5)
        painter.setFont(self.font())
        painter.setPen(QPen(QColor(c["ink"] if enabled else c["ink_3"])))
        text_rect = QRectF(0, 0, max(0.0, box.left() - 8), self.height())
        painter.drawText(
            text_rect, Qt.AlignRight | Qt.AlignVCenter | Qt.AlignAbsolute,
            self.fontMetrics().elidedText(self.text(), Qt.ElideRight,
                                          int(text_rect.width())),
        )


class SpinBox(QSpinBox):
    """عددشمار با دکمه‌های بومی."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setButtonSymbols(QAbstractSpinBox.UpDownArrows)
        self.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.setMinimumHeight(34)


class NumberStepper(QWidget):
    """عددشمار سفارشی با دکمه‌های ‎−‎/‎+‎ و بدون وابستگی به استایل بومی."""

    valueChanged = pyqtSignal(int)

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self._min = 0
        self._max = 100
        self._value = 0
        self._suffix = ""
        self._special = ""
        self.setMinimumHeight(38)
        self.setMinimumWidth(92)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 4, 5, 4)
        layout.setSpacing(2)
        self.minus_button = QPushButton("−")
        self.minus_button.setObjectName("Stepper")
        self.minus_button.setCursor(Qt.PointingHandCursor)
        self.minus_button.clicked.connect(lambda: self.setValue(self._value - 1))
        layout.addWidget(self.minus_button)
        self.value_label = QLabel("")
        self.value_label.setAlignment(Qt.AlignCenter)
        self.value_label.setStyleSheet(
            "font-size: 10pt; font-weight: 700; background: transparent;"
        )
        layout.addWidget(self.value_label, 1)
        self.plus_button = QPushButton("+")
        self.plus_button.setObjectName("Stepper")
        self.plus_button.setCursor(Qt.PointingHandCursor)
        self.plus_button.clicked.connect(lambda: self.setValue(self._value + 1))
        layout.addWidget(self.plus_button)
        self._sync()

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def setRange(self, minimum: int, maximum: int) -> None:  # noqa: N802
        self._min, self._max = int(minimum), int(maximum)
        self.setValue(self._value)

    def setSuffix(self, suffix: str) -> None:  # noqa: N802
        self._suffix = suffix or ""
        self._sync()

    def setSpecialValueText(self, text: str) -> None:  # noqa: N802
        self._special = text or ""
        self._sync()

    def value(self) -> int:
        return self._value

    def setValue(self, value: int) -> None:  # noqa: N802
        clamped = max(self._min, min(self._max, int(value)))
        changed = clamped != self._value
        self._value = clamped
        self._sync()
        if changed:
            self.valueChanged.emit(clamped)

    def _sync(self) -> None:
        if self._special and self._value == self._min:
            text = self._special
        else:
            text = f"{fa(self._value)}{self._suffix}"
        self.value_label.setText(text)
        self.minus_button.setEnabled(self._value > self._min)
        self.plus_button.setEnabled(self._value < self._max)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        body = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(self.colors["surface"]))
        painter.drawRoundedRect(body, 10, 10)
        painter.setPen(QPen(QColor(self.colors["border"]), 1))
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(body, 10, 10)



# ==========================================================================
# صفحه‌های برنامه
# ==========================================================================
# ---------------------------------------------------------------------------
# صفحه‌های برنامه
# ---------------------------------------------------------------------------
PAGE_TITLES: Dict[str, str] = {
    "dashboard": "داشبورد",
    "contracts": "قراردادها",
    "reports": "گزارش مدیریتی",
    "pdfs": "کتابخانهٔ PDF",
    "settings": "تنظیمات",
    "help": "راهنما",
}

PAGE_SUBTITLES: Dict[str, str] = {
    "dashboard": "نمای کلی سبد قراردادها، پیشرفت اختتام و موارد نیازمند پیگیری",
    "contracts": "جستجو، فیلتر و بررسی جزئیات همهٔ قراردادها",
    "reports": "خلاصهٔ مدیریتی قابل چاپ و خروجی اکسل سه‌بخشی",
    "pdfs": "فایل‌های PDF پوشهٔ قراردادها با جستجوی سریع",
    "settings": "منبع داده، پوسته، قلم و بازخوانی خودکار",
    "help": "شروع سریع، میان‌برها و نگاشت امکانات نسخهٔ وب",
}


def clear_layout(layout) -> None:
    """خالی کردن چیدمان و حذف ویجت‌های آن."""
    if layout is None:
        return
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        if widget is not None:
            widget.setParent(None)
            widget.deleteLater()
        else:
            child = item.layout()
            if child is not None:
                clear_layout(child)


def drain_layout(layout) -> None:
    """برداشتن آیتم‌های چیدمان بدون حذف ویجت‌ها (برای چیدمان دوباره)."""
    if layout is None:
        return
    while layout.count():
        layout.takeAt(0)


def page_step_label(text: str, number: int, colors: Dict[str, str]) -> QWidget:
    """یک گام شماره‌دار برای بخش «شروع سریع»."""
    row = QWidget()
    layout = QHBoxLayout(row)
    layout.setContentsMargins(0, 3, 0, 3)
    layout.setSpacing(10)
    badge = QLabel(fa(number))
    badge.setAlignment(Qt.AlignCenter)
    badge.setFixedSize(22, 22)
    badge.setStyleSheet(
        f"background: {colors['brand_soft']}; color: {colors['brand_deep']};"
        f"border-radius: 11px; font-weight: 800; font-size: 8pt;"
    )
    layout.addWidget(badge, 0, Qt.AlignTop)
    label = QLabel(text)
    label.setObjectName("Subtle")
    label.setWordWrap(True)
    layout.addWidget(label, 1)
    return row


def ask_confirm(parent, title: str, text: str, ok_text: str,
                cancel_text: str = "لغو") -> bool:
    """پرسش بله/خیر با دکمه‌های فارسی."""
    box = QMessageBox(parent)
    box.setWindowTitle(title)
    box.setIcon(QMessageBox.Question)
    box.setText(text)
    ok_button = box.addButton(ok_text, QMessageBox.AcceptRole)
    box.addButton(cancel_text, QMessageBox.RejectRole)
    box.exec_()
    return box.clickedButton() is ok_button


def _card_title(title: str, subtitle: str = "") -> QWidget:
    box = QWidget()
    layout = QVBoxLayout(box)
    layout.setContentsMargins(0, 0, 0, 3)
    layout.setSpacing(0)
    main = QLabel(title)
    main.setObjectName("SectionTitle")
    set_text_aligned(main, title, "right")
    layout.addWidget(main)
    if subtitle:
        sub = QLabel(subtitle)
        sub.setObjectName("Muted")
        sub.setWordWrap(True)
        set_text_aligned(sub, subtitle, "right")
        layout.addWidget(sub)
    return box


def _muted(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("Muted")
    label.setWordWrap(True)
    return label


# ---------------------------------------------------------------------------
# داشبورد
# ---------------------------------------------------------------------------
class ElidedLabel(QLabel):
    """برچسب تک‌خطی که متن بلند را با «…» کوتاه می‌کند."""

    def __init__(self, text: str = "", parent=None) -> None:
        super().__init__(parent)
        self._full_text = text
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.setMinimumWidth(0)
        self.setText(text)

    def full_text(self) -> str:
        return self._full_text

    def setText(self, text: str) -> None:  # noqa: N802
        self._full_text = text
        self._refresh()

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self._refresh()

    def _refresh(self) -> None:
        width = max(12, self.width())
        elided = self.fontMetrics().elidedText(self._full_text, Qt.ElideRight, width)
        QLabel.setText(self, elided)


class AttentionRow(QWidget):
    """یک ردیف «نیازمند پیگیری» در داشبورد."""

    activated = pyqtSignal(object)

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.record: Optional[ContractRecord] = None
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(46)
        self._hover = False
        self.setAttribute(Qt.WA_Hover, True)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 3, 4, 3)
        layout.setSpacing(8)
        self.avatar = Avatar("", 28)
        layout.addWidget(self.avatar)
        text_box = QVBoxLayout()
        text_box.setSpacing(1)
        self.name_label = ElidedLabel("—")
        self.name_label.setStyleSheet("font-weight: 700; font-size: 8.5pt;")
        self.meta_label = ElidedLabel("")
        self.meta_label.setObjectName("Muted")
        self.meta_label.setStyleSheet("font-size: 7.5pt;")
        text_box.addWidget(self.name_label)
        text_box.addWidget(self.meta_label)
        layout.addLayout(text_box, 1)
        self.chips = QHBoxLayout()
        self.chips.setSpacing(4)
        self.chips.setAlignment(Qt.AlignVCenter)
        layout.addLayout(self.chips, 0)

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
        if self._hover:
            painter.setPen(Qt.NoPen)
            painter.setBrush(alpha(self.colors["brand"], 16))
            painter.drawRoundedRect(QRectF(self.rect()).adjusted(0, 1, 0, -1), 10, 10)
        painter.setPen(QPen(QColor(self.colors["border_soft"]), 1))
        painter.drawLine(14, self.height() - 1, self.width() - 14, self.height() - 1)

    def set_record(self, record: ContractRecord, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.record = record
        self.avatar.set_name(record.col2 or record.col3 or "?")
        self.name_label.setText(record.col2 or "—")
        self.meta_label.setText(bidi(f"{record.col3} • {_truncate(record.col4 or '—', 30)}"))
        clear_layout(self.chips)
        if record.is_expired:
            self._add_chip("منقضی", colors["danger"], colors["danger_soft"])
        if record.open_flags:
            self._add_chip(f"{fa(len(record.open_flags))} مورد باز",
                           colors["warning"], colors["warning_soft"])
        if record.pdf_file:
            self._add_chip("PDF", colors["info"], colors["info_soft"])
        if not (record.is_expired or record.open_flags or record.pdf_file):
            self._add_chip("بررسی‌شده", colors["success"], colors["success_soft"])

    def _add_chip(self, text: str, fg: str, bg: str) -> None:
        chip = QLabel(text)
        chip.setStyleSheet(
            f"background: {bg}; color: {fg}; border-radius: 8px;"
            f"padding: 1px 6px; font-size: 7pt; font-weight: 700;"
        )
        self.chips.addWidget(chip)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton and self.record is not None:
            self.activated.emit(self.record)
        super().mouseReleaseEvent(event)


class DashboardPage(QWidget):
    """صفحهٔ نخست: کارت قهرمان، کاشی‌های آمار، نمودارها و پیگیری‌ها."""

    filterRequested = pyqtSignal(str)
    recordActivated = pyqtSignal(object)
    navigateRequested = pyqtSignal(str)
    exportRequested = pyqtSignal()
    reloadRequested = pyqtSignal()
    importRequested = pyqtSignal()
    localRunRequested = pyqtSignal()

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.data: Optional[DashboardData] = None
        self.preview_model = None  # در پنجرهٔ اصلی به مدل جدول وصل می‌شود
        self._tile_layout_state: Optional[bool] = None
        self._chart_layout_state: Optional[str] = None
        self._flag_layout_state: Optional[bool] = None
        self._chart_raw: Dict[str, Dict[str, str]] = {}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(self.scroll)

        content = QWidget()
        self.scroll.setWidget(content)
        self.body = QVBoxLayout(content)
        self.body.setContentsMargins(20, 10, 20, 14)
        self.body.setSpacing(9)

        # --- نوار خطای اتصال ---------------------------------------------
        self.error_banner = NoticeBar(colors, "warning")
        self.error_banner.setVisible(False)
        self.error_banner.actionClicked.connect(self.reloadRequested.emit)
        self.error_banner.secondaryClicked.connect(self.localRunRequested.emit)
        self.body.addWidget(self.error_banner)

        # --- کارت قهرمان -------------------------------------------------
        self.hero = Card(colors, radius=20, padding=14, variant="hero")
        hero_box = QVBoxLayout()
        hero_box.setSpacing(8)

        hero_top = QHBoxLayout()
        hero_top.setSpacing(18)
        hero_text = QVBoxLayout()
        hero_text.setSpacing(2)
        self.hero_eyebrow = QLabel("سند قراردادها — نمای زندهٔ سبد")
        self.hero_eyebrow.setStyleSheet(
            f"color: {colors['hero_muted']}; font-size: 8pt; font-weight: 700;")
        self.hero_title = QLabel("پیشرفت اختتام قراردادها")
        self.hero_title.setStyleSheet(
            f"color: {colors['hero_ink']}; font-size: 16pt; font-weight: 800;")
        self.hero_meta = QLabel("—")
        self.hero_meta.setStyleSheet(f"color: {colors['hero_muted']}; font-size: 8pt;")
        self.hero_meta.setWordWrap(True)
        hero_text.addStretch(1)
        hero_text.addWidget(self.hero_eyebrow)
        hero_text.addWidget(self.hero_title)
        hero_text.addWidget(self.hero_meta)
        hero_text.addStretch(1)
        hero_top.addLayout(hero_text, 1)

        gauge_col = QVBoxLayout()
        gauge_col.setSpacing(0)
        self.gauge = RingGauge(colors)
        self.gauge.set_hero_mode(True)
        self.gauge.setFixedSize(104, 104)
        gauge_col.addStretch(1)
        gauge_col.addWidget(self.gauge, 0, Qt.AlignHCenter)
        gauge_col.addStretch(1)
        hero_top.addLayout(gauge_col, 0)
        hero_box.addLayout(hero_top)

        divider = QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet("background: rgba(255,255,255,0.18);")
        hero_box.addWidget(divider)

        hero_bottom = QHBoxLayout()
        hero_bottom.setSpacing(8)
        self.chips_host = QWidget()
        self.chips_host.setAttribute(Qt.WA_TranslucentBackground, True)
        chips_layout = QHBoxLayout(self.chips_host)
        chips_layout.setContentsMargins(0, 0, 0, 0)
        chips_layout.setSpacing(0)
        self.hero_chips = chips_layout
        self.hero_kpis = []
        for caption in ("قرارداد", "میانگین پیشرفت", "فایل PDF"):
            block = QWidget()
            block.setAttribute(Qt.WA_TranslucentBackground, True)
            block_layout = QVBoxLayout(block)
            block_layout.setContentsMargins(0, 0, 0, 0)
            block_layout.setSpacing(1)
            value = QLabel("—")
            value.setObjectName("HeroKpi")
            hint = QLabel(caption)
            hint.setObjectName("HeroKpiHint")
            block_layout.addWidget(value)
            block_layout.addWidget(hint)
            self.hero_kpis.append((value, hint))
            chips_layout.addWidget(block)
            line = QFrame()
            line.setFixedSize(1, 26)
            line.setStyleSheet("background: rgba(255,255,255,0.22);")
            chips_layout.addSpacing(11)
            chips_layout.addWidget(line)
            chips_layout.addSpacing(11)
        hero_bottom.addWidget(self.chips_host, 0, Qt.AlignVCenter)
        hero_bottom.addStretch(1)

        self.view_button = QPushButton("مشاهدهٔ قراردادها")
        self.view_button.setObjectName("OnHeroSolid")
        self.view_button.setCursor(Qt.PointingHandCursor)
        self.view_button.setMinimumHeight(32)
        self.view_button.clicked.connect(lambda: self.navigateRequested.emit("contracts"))
        hero_bottom.addWidget(self.view_button, 0, Qt.AlignVCenter)

        self.reload_button = QPushButton("بازخوانی")
        self.reload_button.setObjectName("OnHero")
        self.reload_button.setCursor(Qt.PointingHandCursor)
        self.reload_button.setMinimumHeight(32)
        self.reload_button.setIcon(make_icon("refresh", colors["hero_ink"], 15))
        self.reload_button.clicked.connect(self.reloadRequested.emit)
        hero_bottom.addWidget(self.reload_button, 0, Qt.AlignVCenter)

        self.report_button = QPushButton("گزارش مدیریتی")
        self.report_button.setObjectName("OnHero")
        self.report_button.setCursor(Qt.PointingHandCursor)
        self.report_button.setMinimumHeight(32)
        self.report_button.setIcon(make_icon("chart", colors["hero_ink"], 15))
        self.report_button.clicked.connect(lambda: self.navigateRequested.emit("reports"))
        hero_bottom.addWidget(self.report_button, 0, Qt.AlignVCenter)

        hero_box.addLayout(hero_bottom)
        self.hero.body.addLayout(hero_box)
        self.body.addWidget(self.hero)

        # --- نوار هشدار ---------------------------------------------------
        self.banner = NoticeBar(colors, "warning")
        self.banner.setVisible(False)
        self.banner.actionClicked.connect(lambda: self.filterRequested.emit("expired"))
        self.body.addWidget(self.banner)

        # --- کاشی‌های آمار ------------------------------------------------
        self.tiles: Dict[str, StatTile] = {}
        self.tiles_grid = QGridLayout()
        self.tiles_grid.setHorizontalSpacing(10)
        self.tiles_grid.setVerticalSpacing(10)
        specs = [
            ("all", "کل قراردادها", "documents", colors["brand"]),
            ("progress", "میانگین پیشرفت اختتام", "gauge", colors["brand_2"]),
            ("deductions", "دارای کسورات", "wallet", colors["accent"]),
            ("pdf", "دارای فایل PDF", "pdf", colors["info"]),
        ]
        for key, title, kind, accent in specs:
            tile = StatTile(key, title, kind, accent, colors)
            tile.clicked.connect(self.filterRequested.emit)
            self.tiles[key] = tile
        self.body.addWidget(SectionHeader("شاخص‌های کلیدی", "نمای کلی سبد قراردادها"))
        self.body.addLayout(self.tiles_grid)

        # --- کارت‌های وضعیت اسناد اختتام ----------------------------------
        self.status_card = ChartCard(
            "وضعیت اسناد اختتام",
            "برای دیدن قراردادهای هر وضعیت، روی کارت آن کلیک کنید",
            colors, min_height=180,
        )
        self.flag_grid = QGridLayout()
        self.flag_grid.setHorizontalSpacing(8)
        self.flag_grid.setVerticalSpacing(8)
        self.flag_tiles: Dict[str, FlagTile] = {}
        for flag_key, flag_label, _column in FLAG_DEFS:
            tile = FlagTile(flag_key, flag_label, colors, self.status_card)
            tile.clicked.connect(self.filterRequested.emit)
            self.flag_tiles[flag_key] = tile
        self.status_card.body.addLayout(self.flag_grid)
        self.body.addWidget(self.status_card)

        # --- نمودارها -----------------------------------------------------
        self.charts_grid = QGridLayout()
        self.charts_grid.setHorizontalSpacing(10)
        self.charts_grid.setVerticalSpacing(10)

        self.progress_card = ChartCard(
            "وضعیت پیشرفت کارکرد",
            "بر اساس ستون‌های کارکرد — برای فیلتر کلیک کنید",
            colors, min_height=206,
        )
        self.donut = DonutChart(colors, CHART_COLORS["progress"])
        self.donut.setMinimumSize(116, 116)
        self.donut.setFixedWidth(160)
        self.donut.sliceClicked.connect(self._on_progress_slice)
        self.legend = LegendList(colors)
        self.legend.setMinimumWidth(160)
        self.legend.setMaximumWidth(206)
        self.legend.itemClicked.connect(self._on_progress_slice)
        donut_row = QHBoxLayout()
        donut_row.setSpacing(10)
        donut_row.addStretch(1)
        donut_row.addWidget(self.legend, 0)
        donut_row.addWidget(self.donut, 0)
        donut_row.addStretch(1)
        self.progress_card.content_layout.addLayout(donut_row, 1)

        self.types_card = ChartCard(
            "پرکاربردترین انواع قرارداد",
            "همهٔ انواع قرارداد — کلیک روی هر میله جدول را فیلتر می‌کند",
            colors, min_height=212,
        )
        self.types_bar = BarChart(colors)
        self.types_bar.barClicked.connect(self._on_type_bar)
        self.types_card.content_layout.addWidget(self.types_bar, 1)

        self.attention_card = ChartCard(
            "نیازمند پیگیری",
            "قراردادهایی با مورد باز یا الحاقیهٔ منقضی",
            colors, min_height=210,
        )
        follow_button = QPushButton("مشاهدهٔ همه")
        follow_button.setObjectName("Link")
        follow_button.setCursor(Qt.PointingHandCursor)
        follow_button.clicked.connect(lambda: self.navigateRequested.emit("contracts"))
        self.attention_card.add_action(follow_button)
        self.attention_box = QVBoxLayout()
        self.attention_box.setSpacing(0)
        self.attention_card.content_layout.addLayout(self.attention_box)
        self.attention_empty = _muted("موردی برای پیگیری وجود ندارد.")
        self.attention_card.content_layout.addWidget(self.attention_empty)
        self.attention_card.content_layout.addStretch(1)

        self.type_card = ChartCard(
            "ترکیب نوع قرارداد", "سهم هر نوع از کل سبد", colors, min_height=200,
        )
        self.type_pie = DonutChart(colors, CHART_COLORS["type"])
        self.type_pie.setMinimumSize(104, 104)
        self.type_pie.setFixedWidth(140)
        self.type_pie.sliceClicked.connect(self._on_type_slice)
        self.type_legend = LegendList(colors)
        self.type_legend.setMinimumWidth(152)
        self.type_legend.setMaximumWidth(190)
        self.type_legend.max_rows = 6
        self.type_legend.itemClicked.connect(self._on_type_slice)
        type_row = QHBoxLayout()
        type_row.setSpacing(8)
        type_row.addStretch(1)
        type_row.addWidget(self.type_legend, 0)
        type_row.addWidget(self.type_pie, 0)
        type_row.addStretch(1)
        self.type_card.content_layout.addLayout(type_row, 1)

        self.annex_card = ChartCard(
            "وضعیت الحاقیه",
            "توزیع زمان الحاقیه — برای فیلتر جدول کلیک کنید",
            colors, min_height=186,
        )
        self._annex_derived = False
        self._annex_done_labels: List[str] = []
        self.annex_donut = DonutChart(colors, CHART_COLORS["annex"])
        self.annex_donut.setMinimumSize(110, 110)
        self.annex_donut.setFixedWidth(150)
        self.annex_donut.sliceClicked.connect(self._on_annex_click)
        self.annex_legend = LegendList(colors)
        self.annex_legend.setMinimumWidth(168)
        self.annex_legend.setMaximumWidth(214)
        self.annex_legend.max_rows = 8
        self.annex_legend.itemClicked.connect(self._on_annex_click)
        annex_row = QHBoxLayout()
        annex_row.setSpacing(10)
        annex_row.addStretch(1)
        annex_row.addWidget(self.annex_legend, 0)
        annex_row.addWidget(self.annex_donut, 0)
        annex_row.addStretch(1)
        self.annex_card.content_layout.addLayout(annex_row, 1)

        self.completion_card = ChartCard(
            "وضعیت پیشرفت اختتام",
            "توزیع پیشرفت اختتام — برای فیلتر جدول کلیک کنید",
            colors, min_height=186,
        )
        self.completion_donut = DonutChart(colors, CHART_COLORS["completion"])
        self.completion_donut.setMinimumSize(110, 110)
        self.completion_donut.setFixedWidth(150)
        self.completion_donut.sliceClicked.connect(
            lambda label: self._emit_chart_filter("completion", label))
        self.completion_legend = LegendList(colors)
        self.completion_legend.setMinimumWidth(164)
        self.completion_legend.setMaximumWidth(210)
        self.completion_legend.itemClicked.connect(
            lambda label: self._emit_chart_filter("completion", label))
        completion_row = QHBoxLayout()
        completion_row.setSpacing(10)
        completion_row.addStretch(1)
        completion_row.addWidget(self.completion_legend, 0)
        completion_row.addWidget(self.completion_donut, 0)
        completion_row.addStretch(1)
        self.completion_card.content_layout.addLayout(completion_row, 1)

        self.body.addWidget(SectionHeader(
            "تحلیل داده‌ها", "روی هر نمودار کلیک کنید تا جدول قراردادها فیلتر شود"))
        self.body.addLayout(self.charts_grid, 1)
        self.body.addStretch(1)
        self._place_cards(False)

    # ---- چیدمان واکنش‌گرا ----
    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        width = self.width()
        mode = "wide" if width >= 1220 else ("mid" if width >= 900 else "narrow")
        self._place_cards(mode)

    def _place_cards(self, mode: str) -> None:
        compact = mode != "wide"
        if self._tile_layout_state != compact:
            self._tile_layout_state = compact
            drain_layout(self.tiles_grid)
            keys = ["all", "progress", "deductions", "pdf"]
            for index, key in enumerate(keys):
                if compact:
                    self.tiles_grid.addWidget(self.tiles[key], index // 2, index % 2)
                else:
                    self.tiles_grid.addWidget(self.tiles[key], 0, index)
            for column in range(4):
                self.tiles_grid.setColumnStretch(column, 1 if (not compact or column < 2) else 0)

        if self._flag_layout_state != compact:
            self._flag_layout_state = compact
            drain_layout(self.flag_grid)
            columns = 2 if compact else 4
            for index, tile in enumerate(self.flag_tiles.values()):
                tile.setParent(self.status_card)
                self.flag_grid.addWidget(tile, index // columns, index % columns)
                tile.show()
            for column in range(4):
                self.flag_grid.setColumnStretch(column, 1 if column < columns else 0)

        if self._chart_layout_state != mode:
            self._chart_layout_state = mode
            drain_layout(self.charts_grid)
            cards = (self.progress_card, self.types_card, self.annex_card,
                     self.completion_card, self.attention_card, self.type_card)
            for card in cards:
                card.setParent(self)
            if mode == "narrow":
                columns = 1
            elif mode == "mid":
                columns = 2
            else:
                columns = 3
            order = cards
            if mode == "wide":
                order = (self.progress_card, self.types_card, self.attention_card,
                         self.annex_card, self.completion_card, self.type_card)
            for index, card in enumerate(order):
                self.charts_grid.addWidget(card, index // columns, index % columns)
            for column in range(3):
                self.charts_grid.setColumnStretch(column, 1 if column < columns else 0)
            for card in cards:
                card.show()

    # ---- داده ----
    def show_error(self, message: str) -> None:
        """نمایش/پاک‌کردن نوار خطای بارگذاری در بالای داشبورد."""
        text = (message or "").strip()
        if not text:
            self.error_banner.setVisible(False)
            return
        first = text.splitlines()[0]
        self.error_banner.set_text(
            bidi(f"خواندن داده از دیتابیس ممکن نشد — {first}"), "تلاش دوباره"
        )
        self.error_banner.set_secondary("اجرای فایل اکسل (بدون دیتابیس)")
        self.error_banner.setVisible(True)

    def set_data(self, data: DashboardData) -> None:
        self.data = data
        self.show_error("")
        self.gauge.set_value(data.progress_percent)
        self.hero_meta.setText(bidi(
            f"منبع: {os.path.basename(data.excel_path) or '—'}  •  شیت {data.sheet_name}"
        ))
        if getattr(data, "metric_note", ""):
            self.hero_meta.setToolTip(data.metric_note)
        pdf_count = sum(1 for record in data.records if record.pdf_file)
        for (value_label, _hint), text in zip(
            self.hero_kpis,
            (fa(data.total_contracts), f"{data.progress_percent_text}٪", fa(pdf_count)),
        ):
            value_label.setText(text)

        total = max(1, data.total_contracts)
        types_dist = data.dist("type")
        self.tiles["all"].set_values(
            fa(data.total_contracts), f"مستند از شیت {data.sheet_name}", 1.0,
            badge=f"{fa(len(types_dist.labels))} نوع",
        )
        self.tiles["progress"].set_values(
            f"{data.progress_percent_text}٪", "میانگین پیشرفت اختتام سند",
            data.progress_percent / 100.0,
        )
        self.tiles["deductions"].set_values(
            fa(data.contracts_with_deductions), f"از کل {fa(data.total_contracts)} قرارداد",
            data.contracts_with_deductions / total,
            badge=f"{fa(round(data.contracts_with_deductions / total * 100))}٪",
        )
        pdf_hint = "قرارداد دارای فایل پیوست PDF"
        pdf_tip = ""
        if pdf_count == 0:
            if not data.pdf_folder:
                pdf_hint = "پوشهٔ PDF تنظیم نشده"
                pdf_tip = "پوشهٔ PDF پیدا نشد؛ در تنظیمات، پوشهٔ فایل‌های PDF را انتخاب کنید."
            elif not data.pdf_total:
                pdf_hint = "پوشهٔ PDF خالی است"
                pdf_tip = f"در پوشهٔ \u202a{data.pdf_folder}\u202c هیچ فایل PDF‌ای نیست."
            else:
                pdf_hint = f"{fa(data.pdf_total)} فایل بدون تطبیق"
                pdf_tip = (
                    f"{fa(data.pdf_total)} فایل PDF در پوشه هست، ولی نام هیچ‌کدام با شمارهٔ "
                    "قرارداد (مثل MEHR-003-1402.pdf) هم‌خوان نیست."
                )
        self.tiles["pdf"].set_values(fa(pdf_count), pdf_hint, pdf_count / total)
        if not pdf_tip and data.pdf_folder:
            pdf_tip = f"پوشهٔ PDF: \u202a{data.pdf_folder}\u202c"
        self.tiles["pdf"].setToolTip(pdf_tip)
        if data.total_contracts == 0:
            for key in ("all", "progress", "deductions", "pdf"):
                self.tiles[key].set_values("—", "داده‌ای برای نمایش نیست", None, badge="")

        if data.records and data.progress_percent_text == "—":
            self.progress_card.subtitle_label.setText(
                "درصد پیشرفت در کتاب کار فعلی نیست — «کتاب کار اصلی (۰۲)» را مشخص کنید"
            )
        else:
            self.progress_card.subtitle_label.setText(
                "بر اساس ستون‌های کارکرد — برای فیلتر کلیک کنید"
            )
        progress = data.dist("progress")
        progress_labels = self._display_labels("progress", progress)
        self.donut.set_data(progress_labels, progress.values,
                            "پیشرفت کل", CHART_COLORS["progress"])
        self.legend.set_data(progress_labels, progress.values, CHART_COLORS["progress"])

        annex = data.dist("annex")
        self._annex_derived = False
        annex_labels = self._display_labels("annex", annex)
        self.annex_donut.set_data(annex_labels, annex.values,
                                  "الحاقیه", CHART_COLORS["annex"])
        self.annex_legend.set_data(annex_labels, annex.values, CHART_COLORS["annex"])

        completion = data.dist("completion")
        completion_labels = self._display_labels("completion", completion)
        self.completion_donut.set_data(completion_labels, completion.values,
                                       "اختتام", CHART_COLORS["completion"])
        self.completion_legend.set_data(completion_labels, completion.values,
                                        CHART_COLORS["completion"])

        flag_counts = data.flag_counts
        no_status_data = bool(data.records) and not any(flag_counts.values())
        subtitle_label = getattr(self.status_card, "subtitle_label", None)
        if subtitle_label is not None:
            subtitle_label.setText(
                "در فایل فعلی هیچ علامت وضعیتی (M) نیست — صفر یعنی سندی ثبت نشده"
                if no_status_data else
                "برای دیدن قراردادهای هر وضعیت، روی کارت آن کلیک کنید"
            )
        for flag_key, tile in self.flag_tiles.items():
            count = flag_counts.get(flag_key, 0)
            tile.set_value(fa(count) if data.total_contracts else "—")
        self.types_bar.set_data(data.top_types)

        self.type_pie.set_data([pretty_label(label) for label in types_dist.labels],
                               types_dist.values, "قرارداد", CHART_COLORS["type"])
        self.type_legend.set_data(types_dist.labels, types_dist.values, CHART_COLORS["type"])

        records = [record for record in data.records
                   if record.open_flags or record.is_expired]
        self._fill_attention(records[:6])
        self.attention_empty.setVisible(not records)
        self.banner.setVisible(data.expired_count > 0)
        if data.expired_count:
            self.banner.set_text(
                f"{fa(data.expired_count)} قرارداد با الحاقیهٔ منقضی‌شده نیازمند بررسی است.",
                "نمایش موارد",
            )

    def _fill_attention(self, records: Sequence[ContractRecord]) -> None:
        clear_layout(self.attention_box)
        for record in list(records)[:3]:
            row = AttentionRow(self.colors)
            row.set_record(record, self.colors)
            row.activated.connect(self.recordActivated.emit)
            self.attention_box.addWidget(row)

    # ---- تعامل ----
    def _display_labels(self, name: str, distribution: Distribution) -> List[str]:
        mapping: Dict[str, str] = {}
        labels: List[str] = []
        for raw in distribution.labels:
            text = pretty_label(raw)
            if name == "progress":
                text = text.replace(" درصد", "٪")
            mapping[text] = raw
            labels.append(text)
        self._chart_raw[name] = mapping
        return labels

    def _on_annex_click(self, label: str) -> None:
        """کلیک روی کارت الحاقیه؛ در حالت مشتق بر اساس تعداد اسناد فیلتر می‌کند."""
        if getattr(self, "_annex_derived", False):
            labels = list(getattr(self, "_annex_done_labels", []))
            if label in labels:
                self.filterRequested.emit(f"docs:{labels.index(label)}")
                return
        self._emit_chart_filter("annex", label)

    def _emit_chart_filter(self, name: str, label: str) -> None:
        raw = self._chart_raw.get(name, {}).get(label, label)
        self.filterRequested.emit(f"{name}:{raw}")

    def _on_progress_slice(self, label: str) -> None:
        self._emit_chart_filter("progress", label)

    def _on_type_slice(self, label: str) -> None:
        self.filterRequested.emit(f"type:{label}")

    def _on_type_bar(self, label: str) -> None:
        self.filterRequested.emit(f"type:{label}")

    # ---- تم ----
    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        for card in self.findChildren(Card):
            card.set_theme(colors)
            if isinstance(card, ChartCard):
                card.set_theme_colors(colors)
        self.gauge.set_theme(colors)
        self.donut.set_theme(colors, CHART_COLORS["progress"])
        self.type_pie.set_theme(colors, CHART_COLORS["type"])
        self.annex_donut.set_theme(colors, CHART_COLORS["annex"])
        self.completion_donut.set_theme(colors, CHART_COLORS["completion"])
        self.legend.set_theme(colors, CHART_COLORS["progress"])
        self.type_legend.set_theme(colors, CHART_COLORS["type"])
        self.annex_legend.set_theme(colors, CHART_COLORS["annex"])
        self.completion_legend.set_theme(colors, CHART_COLORS["completion"])
        for tile in self.flag_tiles.values():
            tile.set_theme(colors)
        self.types_bar.set_theme(colors)
        self.banner.set_theme(colors)
        self.error_banner.set_theme(colors)
        for tile in self.tiles.values():
            tile.set_theme(colors)
        self.hero_eyebrow.setStyleSheet(
            f"color: {colors['hero_muted']}; font-size: 8.5pt; font-weight: 700;"
        )
        self.hero_title.setStyleSheet(
            f"color: {colors['hero_ink']}; font-size: 21pt; font-weight: 800;"
        )
        self.hero_meta.setStyleSheet(f"color: {colors['hero_muted']}; font-size: 9pt;")
        if self.data is not None:
            self.set_data(self.data)
        self.scroll.viewport().update()


# ---------------------------------------------------------------------------
# قراردادها
# ---------------------------------------------------------------------------
class ContractsPage(QWidget):
    """جدول قراردادها با جستجو، فیلتر، مرتب‌سازی و انتخاب چندردیفی."""

    recordActivated = pyqtSignal(object)
    exportRequested = pyqtSignal()
    detailRequested = pyqtSignal(object)
    bookmarkRequested = pyqtSignal(object)

    def __init__(self, notes: NotesStore, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.notes = notes
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 22, 30, 24)
        layout.setSpacing(12)

        # ---- نوار ابزار ----
        toolbar = Card(colors, radius=16, padding=14)
        row = QHBoxLayout()
        row.setSpacing(10)
        self.search_edit = SearchBox("جستجوی نام، شماره، موضوع یا کسورات…")
        self.search_edit.textChanged.connect(self._on_search)
        row.addWidget(self.search_edit, 1)
        self.type_combo = Combo()
        self.type_combo.set_theme_colors(colors)
        self.type_combo.setFixedWidth(168)
        self.type_combo.currentIndexChanged.connect(lambda _i: self._on_filter_changed())
        row.addWidget(self.type_combo)
        self.status_combo = Combo()
        self.status_combo.set_theme_colors(colors)
        self.status_combo.setFixedWidth(186)
        self.status_combo.currentIndexChanged.connect(lambda _i: self._on_filter_changed())
        row.addWidget(self.status_combo)
        self.bookmark_check = CheckBox("فقط نشان‌شده‌ها")
        self.bookmark_check.set_theme_colors(colors)
        self.bookmark_check.toggled.connect(lambda _checked: self._on_filter_changed())
        row.addWidget(self.bookmark_check)
        toolbar.body.addLayout(row)

        pills = QHBoxLayout()
        pills.setSpacing(6)
        self.pills = PillTabs([
            ("all", "همه"),
            ("deductions", "دارای کسورات"),
            ("expired", "منقضی‌شده"),
            ("needs_follow", "نیازمند پیگیری"),
        ], colors)
        self.pills.changed.connect(lambda _key: self._on_filter_changed())
        pills.addWidget(self.pills, 1)
        self.footer_label = QLabel("—")
        self.footer_label.setObjectName("ChipBrand")
        pills.addWidget(self.footer_label)
        export_button = QPushButton("خروجی اکسل")
        export_button.setObjectName("Primary")
        export_button.setIcon(make_icon("download", colors["on_brand"], 15))
        export_button.setCursor(Qt.PointingHandCursor)
        export_button.clicked.connect(self.exportRequested.emit)
        pills.addWidget(export_button)
        toolbar.body.addLayout(pills)
        layout.addWidget(toolbar)

        # ---- جدول ----
        table_card = Card(colors, radius=16, padding=6)
        self.model = ContractTableModel(notes, self)
        self.delegate = ContractDelegate(self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setItemDelegate(self.delegate)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(48)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        header = self.table.horizontalHeader()
        header.setHighlightSections(False)
        header.setSectionsClickable(True)
        header.setMinimumHeight(42)
        header.sectionClicked.connect(self._on_header_clicked)
        for index, spec in enumerate(ContractTableModel.COLUMNS):
            self.table.setColumnWidth(index, spec.width)
            header.setSectionResizeMode(index, QHeaderView.Interactive)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setColumnHidden(11, True)
        self.table.doubleClicked.connect(self._on_double_click)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        table_card.body.addWidget(self.table, 1)
        layout.addWidget(table_card, 1)

        self._extra_filter = ""
        self._columns_menu = QMenu(self)
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(180)
        self._search_timer.timeout.connect(self._apply_view)

        self.populate_combos()
        self.model.apply_view("", "all", False)
        self._refresh_counts()

    # ---- آماده‌سازی ----
    def populate_combos(self) -> None:
        records = self.model.all_records()
        self.type_combo.blockSignals(True)
        self.type_combo.clear()
        self.type_combo.addItem("همهٔ انواع", "")
        for label in sorted({record.f_value for record in records if record.f_value}, key=normalize_per):
            self.type_combo.addItem(label, label)
        self.type_combo.setCurrentIndex(0)
        self.type_combo.blockSignals(False)

        self.status_combo.blockSignals(True)
        self.status_combo.clear()
        self.status_combo.addItem("همهٔ وضعیت‌ها", "")
        for label in sorted({record.as_status for record in records if record.as_status}, key=normalize_per):
            self.status_combo.addItem(label, label)
        self.status_combo.setCurrentIndex(0)
        self.status_combo.blockSignals(False)

    def set_data(self, data: DashboardData) -> None:
        self.model.set_records(data.records)
        self.populate_combos()
        self.pills.set_current("all")
        self.bookmark_check.blockSignals(True)
        self.bookmark_check.setChecked(False)
        self.bookmark_check.blockSignals(False)
        self._apply_view()
        QTimer.singleShot(150, self._fit_columns)

    MIN_WIDTHS = (38, 120, 96, 140, 88, 54, 130, 86, 96, 84, 38, 22)

    def _fit_columns(self) -> None:
        """عرض ستون‌ها را دقیقاً در عرض دید جدول توزیع می‌کند (بدون اسکرول افقی)."""
        available = self.table.viewport().width() - 14
        if available <= 40:
            return
        specs = ContractTableModel.COLUMNS
        base_total = sum(spec.width for spec in specs) or 1
        widths = [(0 if self.table.isColumnHidden(index) else
                   max(minimum, int(spec.width * available / base_total)))
                  for index, (spec, minimum) in enumerate(zip(specs, self.MIN_WIDTHS))]
        overflow = sum(widths) - available
        if overflow > 0:
            shrinkable = sum(max(0, width - 34) for width in widths) or 1
            widths = [width - int(overflow * max(0, width - 34) / shrinkable)
                      for width in widths]
        leftover = available - sum(widths)
        if leftover != 0:
            widths[3] = max(60, widths[3] + leftover)
        widths[11] = 0
        for index, width in enumerate(widths):
            self.table.setColumnWidth(index, width)

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        QTimer.singleShot(0, self._fit_columns)

    def showEvent(self, event) -> None:  # noqa: N802
        super().showEvent(event)
        QTimer.singleShot(0, self._fit_columns)
        QTimer.singleShot(150, self._fit_columns)

    # ---- فیلتر ----
    def _active_key(self) -> str:
        if self._extra_filter:
            return self._extra_filter
        type_value = self.type_combo.currentData() or ""
        status_value = self.status_combo.currentData() or ""
        if type_value:
            return f"type:{type_value}"
        if status_value:
            return f"status:{status_value}"
        if self.bookmark_check.isChecked():
            return "bookmarked"
        return self.pills.current() or "all"

    def _on_search(self, _text: str) -> None:
        self._search_timer.start()

    def _on_filter_changed(self) -> None:
        self._extra_filter = ""
        if self.sender() and hasattr(self.sender(), "currentData"):
            if self.sender() is self.type_combo and (self.type_combo.currentData() or ""):
                self.status_combo.blockSignals(True)
                self.status_combo.setCurrentIndex(0)
                self.status_combo.blockSignals(False)
            if self.sender() is self.status_combo and (self.status_combo.currentData() or ""):
                self.type_combo.blockSignals(True)
                self.type_combo.setCurrentIndex(0)
                self.type_combo.blockSignals(False)
        self._apply_view()

    def _apply_view(self) -> None:
        self.model.apply_view(self.search_edit.text().strip(), self._active_key(),
                              self.bookmark_check.isChecked())
        self.table.clearSelection()
        self.table.setCurrentIndex(QModelIndex())
        self._refresh_counts()

    def _filter_description(self) -> str:
        key = self._active_key()
        if key in ("", "all", None):
            return ""
        simple = {
            "deductions": "دارای کسورات اسناد اختتام",
            "expired": "الحاقیهٔ منقضی‌شده",
            "needs_follow": "نیازمند پیگیری",
            "bookmarked": "فقط نشان‌شده‌ها",
        }
        if key in simple:
            return simple[key]
        if ":" in key:
            field, _, value = key.partition(":")
            titles = {
                "type": "نوع قرارداد",
                "status": "وضعیت پکیج",
                "progress": "پیشرفت کارکرد",
                "completion": "پیشرفت اختتام",
                "annex": "الحاقیه",
            }
            if field in ("progress", "completion", "annex"):
                value = pretty_label(value)
            return f"{titles.get(field, field)}: {value}"
        return FLAG_LABELS.get(key, key)

    def _refresh_counts(self) -> None:
        visible = self.model.visible_count()
        total = self.model.source_count()
        extra = f" از {fa(total)}" if visible != total else ""
        text = f"{fa(visible)}{extra} قرارداد"
        description = self._filter_description()
        if description:
            text += f" — فیلتر: {description}"
        self.footer_label.setText(bidi(text))

    def set_filter(self, key: str) -> None:
        self._extra_filter = ""
        known_pills = ("all", "deductions", "expired", "needs_follow", "bookmarked")
        if key in known_pills:
            if key == "bookmarked":
                self.bookmark_check.blockSignals(True)
                self.bookmark_check.setChecked(True)
                self.bookmark_check.blockSignals(False)
                self.pills.set_current("all")
            else:
                self.bookmark_check.blockSignals(True)
                self.bookmark_check.setChecked(False)
                self.bookmark_check.blockSignals(False)
                self.pills.set_current(key)
        elif key.startswith(("type:", "status:")):
            field, _, value = key.partition(":")
            if field == "type":
                index = self.type_combo.findData(value)
                self.type_combo.blockSignals(True)
                self.type_combo.setCurrentIndex(max(0, index))
                self.type_combo.blockSignals(False)
            elif field == "status":
                index = self.status_combo.findData(value)
                self.status_combo.blockSignals(True)
                self.status_combo.setCurrentIndex(max(0, index))
                self.status_combo.blockSignals(False)
        else:
            for combo in (self.type_combo, self.status_combo):
                combo.blockSignals(True)
                combo.setCurrentIndex(0)
                combo.blockSignals(False)
            self.bookmark_check.blockSignals(True)
            self.bookmark_check.setChecked(False)
            self.bookmark_check.blockSignals(False)
            self.pills.set_current("all")
            self._extra_filter = key
        self._apply_view()

    def current_filter(self) -> str:
        return self._active_key()

    def set_search(self, text: str) -> None:
        if self.search_edit.text() != text:
            self.search_edit.blockSignals(True)
            self.search_edit.setText(text)
            self.search_edit.blockSignals(False)
        self._apply_view()

    def clear_filters(self) -> None:
        self.type_combo.blockSignals(True)
        self.status_combo.blockSignals(True)
        self.type_combo.setCurrentIndex(0)
        self.status_combo.setCurrentIndex(0)
        self.type_combo.blockSignals(False)
        self.status_combo.blockSignals(False)
        self.bookmark_check.blockSignals(True)
        self.bookmark_check.setChecked(False)
        self.bookmark_check.blockSignals(False)
        self.pills.set_current("all")
        self._extra_filter = ""
        self._apply_view()

    def refresh_bookmarks(self) -> None:
        self.model.refresh_bookmarks()
        self._refresh_counts()

    # ---- انتخاب و ردیف‌ها ----
    def visible_records(self) -> List[ContractRecord]:
        return self.model.visible_records()

    def selected_records(self) -> List[ContractRecord]:
        rows = sorted({index.row() for index in self.table.selectionModel().selectedRows()})
        records = [self.model.record_at(row) for row in rows]
        return [record for record in records if record is not None]

    def record_at(self, row: int) -> Optional[ContractRecord]:
        return self.model.record_at(row)

    def _on_double_click(self, index) -> None:
        record = self.model.record_at(index.row())
        if record is not None:
            self.recordActivated.emit(record)

    def _on_header_clicked(self, column: int) -> None:
        if column < 0 or column >= len(ContractTableModel.COLUMNS):
            return
        order = Qt.AscendingOrder
        if self.table.horizontalHeader().sortIndicatorSection() == column:
            current = self.table.horizontalHeader().sortIndicatorOrder()
            order = Qt.DescendingOrder if current == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            order = Qt.AscendingOrder
        self.table.horizontalHeader().setSortIndicator(column, order)
        self.model.set_sort(column, order)
        self._refresh_counts()

    def _show_context_menu(self, point) -> None:
        record = self.model.record_at(self.table.rowAt(point.y()))
        menu = QMenu(self)
        detail_action = menu.addAction("جزئیات قرارداد")
        detail_action.setEnabled(record is not None)
        bookmark_action = menu.addAction("نشان‌گذاری / برداشتن نشان")
        bookmark_action.setEnabled(record is not None)
        menu.addSeparator()
        copy_action = menu.addAction("کپی شمارهٔ قرارداد")
        copy_action.setEnabled(record is not None)
        export_action = menu.addAction("خروجی اکسل نمای فعلی")
        chosen = menu.exec_(self.table.viewport().mapToGlobal(point))
        if chosen is None:
            return
        if chosen is detail_action and record is not None:
            self.recordActivated.emit(record)
        elif chosen is bookmark_action and record is not None:
            self.bookmarkRequested.emit(record)
        elif chosen is copy_action and record is not None:
            QApplication.clipboard().setText(record.col3)
        elif chosen is export_action:
            self.exportRequested.emit()

    def _column_menu(self) -> QMenu:
        self._columns_menu.clear()
        for index, spec in enumerate(ContractTableModel.COLUMNS):
            action = self._columns_menu.addAction(spec.title)
            action.setCheckable(True)
            action.setChecked(not self.table.isColumnHidden(index))
            action.toggled.connect(lambda checked, i=index: self.table.setColumnHidden(i, not checked))
        return self._columns_menu

    # ---- تم ----
    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.model.set_theme_colors(colors)
        self.delegate.set_theme_colors(colors)
        self.type_combo.set_theme_colors(colors)
        self.status_combo.set_theme_colors(colors)
        self.bookmark_check.set_theme_colors(colors)
        self.search_edit.set_icon_color(colors["ink_3"])
        for card in self.findChildren(Card):
            card.set_theme(colors)



# ==========================================================================
# صفحه‌های گزارش، PDF، تنظیمات و راهنما
# ==========================================================================
# ---------------------------------------------------------------------------
# گزارش مدیریتی
# ---------------------------------------------------------------------------
def _cell_width(widths: Optional[Sequence[str]], index: int) -> str:
    if not widths or index >= len(widths):
        return ""
    return f' width="{widths[index]}"'


def _report_table(headers: Sequence[str], rows: Sequence[Sequence[Any]],
                  colors: Dict[str, str], widths: Optional[Sequence[str]] = None,
                  table_width: str = "100%") -> str:
    """جدول گزارش؛ ستون‌ها معکوس چیده می‌شوند تا در RTL راست‌به‌چپ درست دیده شوند."""
    is_dark = QColor(colors["surface"]).lightness() < 128
    head_bg = colors["surface_3"] if is_dark else colors["brand_soft"]
    head_fg = colors["brand_2"] if is_dark else colors["brand_deep"]
    order = list(range(len(headers)))[::-1]
    head_cells = "".join(
        f'<th bgcolor="{head_bg}" align="center"{_cell_width(widths, i)} '
        f'style="padding:7px; color:{head_fg};">'
        f'{html_module.escape(str(headers[i]))}</th>'
        for i in order
    )
    body = ""
    for row in rows:
        cells = "".join(
            f'<td align="right"{_cell_width(widths, i)} style="padding:6px 8px;'
            f' border-bottom:1px solid {colors["border_soft"]};">'
            f'{html_module.escape(str(row[i]))}</td>'
            for i in order if i < len(row)
        )
        body += f"<tr>{cells}</tr>"
    colgroup = ""
    if widths:
        colgroup = "<colgroup>" + "".join(
            f'<col width="{widths[i]}">' for i in order if i < len(widths)
        ) + "</colgroup>"
    return (
        f'<table dir="rtl" width="{table_width}" cellspacing="0" cellpadding="0" border="0" '
        f'style="border-collapse:collapse; margin-bottom:14px;">{colgroup}'
        f"<tr>{head_cells}</tr>{body}</table>"
    )


def _report_grid(items: Sequence[Sequence[str]], colors: Dict[str, str],
                 per_row: int = 3) -> str:
    """شبکهٔ آمار گزارش: چند شاخص در هر ردیف با چیدمان راست‌به‌چپ."""
    cell_bg = colors["surface_2"]
    rows_html = ""
    for start in range(0, len(items), per_row):
        chunk = list(items[start:start + per_row])[::-1]
        cells = "".join(
            f'<td width="{int(100 / per_row)}%" align="right" bgcolor="{cell_bg}" '
            f'style="padding:9px 12px;">'
            f'<div style="color:{colors["ink_3"]}; font-size:8.5pt;">{html_module.escape(str(label))}</div>'
            f'<div style="color:{colors["ink"]}; font-size:13pt; font-weight:bold;">{value}</div>'
            f"</td>"
            for label, value in chunk
        )
        rows_html += f'<tr><td width="6" bgcolor="{cell_bg}"></td>{cells}</tr>'
    return (
        f'<table dir="rtl" width="100%" cellspacing="3" cellpadding="0" border="0" '
        f'style="border-collapse:separate; margin-bottom:14px;">{rows_html}</table>'
    )


def build_report_html(data: DashboardData, colors: Dict[str, str]) -> str:
    """ساخت گزارش مدیریتی HTML با ظاهری آراسته و قابل چاپ."""
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
        distribution = data.status_dist() if name == "status" else data.dist(name)
        labels = "، ".join(
            f"{pretty_label(label)} ({fa(value)})" for label, value in distribution.as_list()
        ) or "—"
        dist_rows.append([CHART_TITLES.get(name, name), labels])

    callout_rows = "".join(
        f'<td width="25%" align="right" style="padding:10px 12px;">'
        f'<div style="color:{colors["ink_3"]}; font-size:9pt;">{title}</div>'
        f'<div style="color:{colors["ink"]}; font-size:16pt; font-weight:bold;">{value}</div>'
        f"</td>"
        for title, value in (
            ("منقضی‌شده", fa_num(data.expired_count)),
            ("نیازمند پیگیری", fa_num(len(follow_rows))),
            ("پیشرفت اختتام", f"{data.progress_percent_text}٪"),
            ("کل قراردادها", fa_num(data.total_contracts)),
        )
    )
    is_dark = QColor(colors["surface"]).lightness() < 128
    band = colors["hero_1"] if is_dark else colors["brand_deep"]
    band_soft = colors["hero_muted"]
    band_title = colors["hero_ink"] if is_dark else "#FFFFFF"
    return f"""
    <div dir="rtl" style="font-family: Tahoma, sans-serif; color: {colors['ink']};">
      <table dir="rtl" width="100%" cellspacing="0" cellpadding="0" border="0"
             style="border-collapse:collapse; margin-bottom:12px;">
        <tr>
          <td bgcolor="{band}" style="padding:18px 20px; border-radius:12px;">
            <div style="color:{band_soft}; font-size:9pt;">گزارش مدیریتی — پروژهٔ مهر خلیج فارس</div>
            <div style="color:{band_title}; font-size:17pt; font-weight:bold; padding-top:4px;">
              خلاصهٔ وضعیت سبد قراردادها
            </div>
            <div style="color:{band_soft}; font-size:9pt; padding-top:4px;">
              تاریخ تهیه: {today_jalali()} &nbsp;•&nbsp; منبع داده: {html_module.escape(os.path.basename(data.excel_path) or '—')}
              &nbsp;•&nbsp; منبع: {html_module.escape(data.source_note or 'دیتابیس مرکزی')}
              &nbsp;•&nbsp; آخرین بارگذاری: {data.source_mtime_text}
            </div>
          </td>
        </tr>
      </table>
      <table dir="rtl" width="100%" cellspacing="0" cellpadding="0" border="0"
             bgcolor="{colors['surface_2']}" style="border-collapse:separate; margin-bottom:14px;">
        <tr>{callout_rows}</tr>
      </table>
      <h3 style="color:{colors['brand_deep']}; margin:12px 0 6px 0;">شاخص‌های کلی</h3>
      {_report_grid(overview, colors, per_row=3)}
      <h3 style="color:{colors['brand_deep']}; margin:12px 0 6px 0;">قراردادهای نیازمند پیگیری</h3>
      {_report_table(
          ["نام شرکت", "شمارهٔ قرارداد", "موضوع", "منقضی", "موارد باز", "شرح موارد"],
          shown, colors, ["18%", "11%", "22%", "7%", "7%", "35%"])}
      <div style="color:{colors['ink_3']}; margin-bottom:12px;">مجموع قراردادهای نیازمند پیگیری:
        {fa_num(len(follow_rows))}
        {f" — فقط {fa(len(shown))} مورد اول نمایش داده شده است." if len(follow_rows) > len(shown) else ""}
      </div>
      <h3 style="color:{colors['brand_deep']}; margin:12px 0 6px 0;">توزیع‌ها</h3>
      {_report_table(["نمودار", "توزیع (تعداد)"], dist_rows, colors, ["25%", "75%"])}
      <div style="color:{colors['ink_3']}; font-size:8.5pt; margin-top:10px;">
        تهیه‌شده با {APP_NAME} — نسخهٔ {APP_VERSION}
      </div>
    </div>
    """


class ReportsPage(QWidget):
    """پیش‌نمایش گزارش مدیریتی با خروجی اکسل و چاپ."""

    exportExcelRequested = pyqtSignal()
    printRequested = pyqtSignal()
    refreshRequested = pyqtSignal()

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self._data: Optional[DashboardData] = None
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 22, 30, 24)
        layout.setSpacing(12)

        toolbar = Card(colors, radius=16, padding=14)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.summary_chip = QLabel("—")
        self.summary_chip.setObjectName("ChipBrand")
        row.addWidget(self.summary_chip)
        self.follow_chip = QLabel("—")
        self.follow_chip.setObjectName("Chip")
        row.addWidget(self.follow_chip)
        row.addStretch(1)
        refresh_btn = QPushButton("بازخوانی")
        refresh_btn.setObjectName("Ghost")
        refresh_btn.setIcon(make_icon("refresh", colors["ink_2"], 15))
        refresh_btn.clicked.connect(self.refreshRequested.emit)
        row.addWidget(refresh_btn)
        print_btn = QPushButton("چاپ / ذخیرهٔ PDF")
        print_btn.setIcon(make_icon("printer", colors["ink_2"], 15))
        print_btn.setCursor(Qt.PointingHandCursor)
        print_btn.clicked.connect(self.printRequested.emit)
        row.addWidget(print_btn)
        export_btn = QPushButton("خروجی اکسل گزارش")
        export_btn.setObjectName("Primary")
        export_btn.setIcon(make_icon("download", "#FFFFFF", 15))
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.clicked.connect(self.exportExcelRequested.emit)
        row.addWidget(export_btn)
        toolbar.body.addLayout(row)
        layout.addWidget(toolbar)

        browser_card = Card(colors, radius=16, padding=8)
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(False)
        self.browser.setFrameShape(QFrame.NoFrame)
        self.browser.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        configure_rtl_browser(self.browser)
        self.browser.setStyleSheet(
            f"QTextBrowser {{ background: {colors['paper']}; border: none; padding: 12px; }}"
        )
        browser_card.body.addWidget(self.browser, 1)
        layout.addWidget(browser_card, 1)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.browser.setStyleSheet(
            f"QTextBrowser {{ background: {colors['paper']}; border: none; padding: 12px; }}"
        )
        for card in self.findChildren(Card):
            card.set_theme(colors)
        if self._data is not None:
            self.update_data(self._data)

    def update_data(self, data: DashboardData) -> None:
        self._data = data
        self.browser.setHtml(build_report_html(data, self.colors))
        configure_rtl_browser(self.browser)
        self.summary_chip.setText(bidi(f"{data.progress_percent_text}٪ پیشرفت اختتام"))
        follow = sum(1 for record in data.records if record.open_flags or record.is_expired)
        self.follow_chip.setText(bidi(f"{fa(follow)} نیازمند پیگیری"))

    def set_data(self, data: DashboardData) -> None:
        self.update_data(data)

    def html(self) -> str:
        return self.browser.toHtml()


# ---------------------------------------------------------------------------
# کتابخانهٔ PDF
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
        layout.setContentsMargins(30, 22, 30, 24)
        layout.setSpacing(12)

        toolbar = Card(colors, radius=16, padding=14)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.search_edit = SearchBox("جستجوی نام فایل PDF…")
        self.search_edit.textChanged.connect(self._apply_filter)
        row.addWidget(self.search_edit, 1)
        self.count_chip = QLabel("—")
        self.count_chip.setObjectName("Chip")
        row.addWidget(self.count_chip)
        refresh_btn = QPushButton("بازخوانی فهرست")
        refresh_btn.setIcon(make_icon("refresh", colors["ink_2"], 15))
        refresh_btn.clicked.connect(self.refresh_list)
        row.addWidget(refresh_btn)
        folder_btn = QPushButton("انتخاب پوشه")
        folder_btn.setIcon(make_icon("folder", colors["ink_2"], 15))
        folder_btn.clicked.connect(self.folderRequested.emit)
        row.addWidget(folder_btn)
        toolbar.body.addLayout(row)
        self.subtitle = QLabel("پوشهٔ PDF تنظیم نشده است.")
        self.subtitle.setObjectName("Muted")
        self.subtitle.setWordWrap(True)
        toolbar.body.addWidget(self.subtitle)
        layout.addWidget(toolbar)

        list_card = Card(colors, radius=16, padding=8)
        self.list_widget = QListWidget()
        self.list_widget.setSpacing(0)
        self.list_widget.setUniformItemSizes(True)
        self.list_widget.setMouseTracking(True)
        self.list_widget.itemDoubleClicked.connect(self._open_item)
        self._apply_list_style()
        list_card.body.addWidget(self.list_widget, 1)
        layout.addWidget(list_card, 1)

        footer = QHBoxLayout()
        self.footer_label = QLabel("—")
        self.footer_label.setObjectName("Muted")
        footer.addWidget(self.footer_label, 1)
        open_btn = QPushButton("باز کردن فایل انتخاب‌شده")
        open_btn.setObjectName("Primary")
        open_btn.setIcon(make_icon("external", "#FFFFFF", 15))
        open_btn.setCursor(Qt.PointingHandCursor)
        open_btn.clicked.connect(self._open_selected)
        footer.addWidget(open_btn)
        layout.addLayout(footer)

    def _apply_list_style(self) -> None:
        c = self.colors
        self.list_widget.setStyleSheet(
            f"QListWidget {{ background: transparent; border: none; outline: none; }}"
            f"QListWidget::item {{ padding: 8px 10px; border-radius: 9px;"
            f" color: {c['ink']}; font-size: 9.5pt; font-weight: 600;"
            f" border-bottom: 1px solid {c['border_soft']}; }}"
            f"QListWidget::item:hover {{ background: {c['surface_2']}; }}"
            f"QListWidget::item:selected {{ background: {c['brand_soft']};"
            f" color: {c['brand_deep']}; }}"
        )

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.search_edit.set_icon_color(colors["ink_3"])
        self._apply_list_style()
        for card in self.findChildren(Card):
            card.set_theme(colors)

    def set_folder(self, folder: str) -> None:
        self.folder = folder or ""
        self.refresh_list()

    def refresh_list(self) -> None:
        self._files = []
        error = ""
        if self.folder and os.path.isdir(self.folder):
            try:
                self._files = scan_pdf_files(self.folder, depth=2, limit=5000)
            except Exception as exc:
                error = str(exc)
        if self.folder:
            subtitle = f"پوشه: \u202a{self.folder}\u202c — {fa(len(self._files))} فایل PDF (زیرپوشه‌ها هم)"
        else:
            subtitle = "پوشهٔ PDF پیدا نشد — با دکمهٔ «انتخاب پوشه» مسیر فایل‌های PDF را بدهید."
        if error:
            subtitle = f"خطا در خواندن پوشه: {error}"
        self.subtitle.setText(subtitle)
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
            item = QListWidgetItem(f"\u202a{name}\u202c")
            item.setData(Qt.UserRole, name)
            item.setIcon(make_icon("pdf", self.colors["danger"], 18))
            item.setSizeHint(QSize(0, 38))
            self.list_widget.addItem(item)
        if not shown:
            message = ("پوشه‌ای انتخاب نشده است." if not self.folder
                       else "فایل PDF‌ای در این پوشه و زیرپوشه‌ها پیدا نشد.")
            item = QListWidgetItem(message)
            item.setFlags(Qt.NoItemFlags)
            self.list_widget.addItem(item)
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
# تنظیمات
# ---------------------------------------------------------------------------
class SettingsPage(QWidget):
    """تنظیم اتصال دیتابیس مرکزی، پوشهٔ PDF، پوسته و بازخوانی خودکار."""

    saveRequested = pyqtSignal(dict)
    reloadRequested = pyqtSignal()
    openDataFolderRequested = pyqtSignal()
    testRequested = pyqtSignal()
    detectRequested = pyqtSignal()
    verifyRequested = pyqtSignal()
    importRequested = pyqtSignal(str)
    localRunRequested = pyqtSignal(str)
    systemUpdateRequested = pyqtSignal(str)
    diagnoseRequested = pyqtSignal(str)
    sourceWorkbookRequested = pyqtSignal()
    workbookPickRequested = pyqtSignal()

    def __init__(self, config: AppConfig, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.config = config
        self.colors = colors

        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 22, 30, 20)
        layout.setSpacing(12)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        layout.addWidget(scroll, 1)
        content = QWidget()
        scroll.setWidget(content)
        holder = QHBoxLayout(content)
        holder.setContentsMargins(0, 0, 0, 0)
        holder.setSpacing(14)

        right_column = QVBoxLayout()
        right_column.setSpacing(14)
        left_column = QVBoxLayout()
        left_column.setSpacing(14)

        # ---- اتصال دیتابیس ----
        source_card = Card(colors, radius=16, padding=18)
        source_card.body.addWidget(_card_title(
            "اتصال دیتابیس",
            "داده‌ها از دیتابیس مرکزی (MySQL) خوانده می‌شود؛ رمز اختیاری است",
        ))
        form = QGridLayout()
        form.setHorizontalSpacing(6)
        form.setVerticalSpacing(8)

        form.addWidget(QLabel("میزبان:"), 0, 0)
        self.host_edit = QLineEdit(str(config.get("db_host") or "127.0.0.1"))
        self.host_edit.setPlaceholderText("127.0.0.1")
        self.host_edit.setMinimumWidth(96)
        form.addWidget(self.host_edit, 0, 1)
        form.addWidget(QLabel("پورت:"), 0, 2)
        self.port_spin = NumberStepper(colors)
        self.port_spin.setRange(1, 65535)
        self.port_spin.setValue(int(config.get("db_port") or 3306))
        form.addWidget(self.port_spin, 0, 3)

        form.addWidget(QLabel("دیتابیس:"), 1, 0)
        self.db_edit = QLineEdit(str(config.get("db_name") or "master_data"))
        self.db_edit.setPlaceholderText("master_data")
        self.db_edit.setMinimumWidth(96)
        self.db_edit.setToolTip(
            "دیتابیسی که «مدیریت دیتابیس مرکزی» ساخته است (پیش‌فرض: master_data)."
        )
        form.addWidget(self.db_edit, 1, 1)
        form.addWidget(QLabel("کاربر:"), 1, 2)
        self.user_edit = QLineEdit(str(config.get("db_user") or "master_app"))
        self.user_edit.setPlaceholderText("master_app")
        self.user_edit.setMinimumWidth(96)
        form.addWidget(self.user_edit, 1, 3)

        form.addWidget(QLabel("رمز عبور:"), 2, 0)
        self.password_edit = QLineEdit(str(config.get_password() or ""))
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("رمز — اگر ندارید خالی بگذارید")
        self.password_edit.setMinimumWidth(96)
        form.addWidget(self.password_edit, 2, 1)
        self.remember_check = CheckBox("ذخیرهٔ رمز")
        self.remember_check.set_theme_colors(colors)
        self.remember_check.setChecked(bool(config.get("db_remember_password")))
        form.addWidget(self.remember_check, 2, 2, 1, 2)

        form.addWidget(QLabel("شناسهٔ کتاب کار:"), 3, 0)
        self.workbook_spin = NumberStepper(colors)
        self.workbook_spin.setRange(1, 999999)
        self.workbook_spin.setValue(int(config.get("db_workbook_id") or 1))
        self.workbook_spin.setToolTip(
            "شناسهٔ کتاب کار در دیتابیس؛ کتاب کار ۰۲ خودتان را با هر شناسه‌ای "
            "(مثلاً ۳) ذخیره کرده‌اید، همان را بگذارید. اگر شناسه درست نباشد، "
            "برنامه خودش بهترین کتاب کار واقعی را انتخاب می‌کند."
        )
        workbook_row = QHBoxLayout()
        workbook_row.setSpacing(6)
        workbook_row.addWidget(self.workbook_spin, 1)
        self.workbook_pick_button = QPushButton("کتاب کارها…")
        self.workbook_pick_button.setObjectName("Ghost")
        self.workbook_pick_button.setCursor(Qt.PointingHandCursor)
        self.workbook_pick_button.setToolTip(
            "فهرست کتاب کارهای ذخیره‌شده در دیتابیس را نشان می‌دهد تا شناسهٔ "
            "کتاب کار ۰۲ (مثلاً ۳) را انتخاب کنید."
        )
        self.workbook_pick_button.clicked.connect(self.workbookPickRequested.emit)
        workbook_row.addWidget(self.workbook_pick_button, 0)
        form.addLayout(workbook_row, 3, 1)
        form.addWidget(QLabel("نام شیت:"), 3, 2)
        self.sheet_edit = QLineEdit(str(config.get("sheet_name") or SHEET_NAME_DEFAULT))
        self.sheet_edit.setPlaceholderText(SHEET_NAME_DEFAULT)
        self.sheet_edit.setMinimumWidth(96)
        form.addWidget(self.sheet_edit, 3, 3)

        form.addWidget(QLabel("پوشهٔ PDF:"), 4, 0)
        self.pdf_edit = QLineEdit(str(config.get("pdf_folder") or ""))
        resolved_pdf = resolve_pdf_folder(config)
        self.pdf_edit.setPlaceholderText(resolved_pdf or "پوشه‌ای که فایل‌های PDF قراردادها در آن است")
        self.pdf_edit.setMinimumWidth(96)
        form.addWidget(self.pdf_edit, 4, 1, 1, 2)
        browse_pdf = QPushButton("انتخاب…")
        browse_pdf.setIcon(make_icon("folder", colors["ink_2"], 15))
        browse_pdf.setCursor(Qt.PointingHandCursor)
        browse_pdf.clicked.connect(self._browse_pdf)
        form.addWidget(browse_pdf, 4, 3)
        self.pdf_note = _muted("")
        form.addWidget(self.pdf_note, 5, 0, 1, 4)
        form.setColumnStretch(1, 1)
        form.setColumnStretch(3, 1)
        source_card.body.addLayout(form)

        actions_row = QHBoxLayout()
        actions_row.setSpacing(8)
        detect_btn = QPushButton("دریافت تنظیمات مدیر دیتابیس")
        detect_btn.setObjectName("Ghost")
        detect_btn.setIcon(make_icon("settings", colors["ink_2"], 15))
        detect_btn.setToolTip(
            "میزبان، پورت، نام دیتابیس، کاربر و رمز از تنظیمات امن «مدیریت دیتابیس مرکزی» خوانده می‌شود."
        )
        detect_btn.setCursor(Qt.PointingHandCursor)
        detect_btn.clicked.connect(self.detectRequested.emit)
        actions_row.addWidget(detect_btn)
        test_btn = QPushButton("آزمایش اتصال")
        test_btn.setObjectName("Ghost")
        test_btn.setIcon(make_icon("external", colors["ink_2"], 15))
        test_btn.setCursor(Qt.PointingHandCursor)
        test_btn.clicked.connect(self.testRequested.emit)
        actions_row.addWidget(test_btn)
        verify_btn = QPushButton("بررسی اعداد دیتابیس")
        verify_btn.setObjectName("Ghost")
        verify_btn.setIcon(make_icon("check", colors["ink_2"], 15))
        verify_btn.setCursor(Qt.PointingHandCursor)
        verify_btn.setToolTip(
            "تعداد قرارداد، کسورات و هشت کارت وضعیت را با SQL از جدول‌های دیتابیس "
            "می‌شمارد و با اعداد داشبورد مقایسه می‌کند."
        )
        verify_btn.clicked.connect(self.verifyRequested.emit)
        actions_row.addWidget(verify_btn)
        pdf_btn = QPushButton("پیدا کردن پوشهٔ PDF")
        pdf_btn.setObjectName("Ghost")
        pdf_btn.setIcon(make_icon("folder", colors["ink_2"], 15))
        pdf_btn.setToolTip("پوشهٔ PDF را کنار فایل اکسل، کنار برنامه و در مسیر شبکه جستجو می‌کند.")
        pdf_btn.setCursor(Qt.PointingHandCursor)
        pdf_btn.clicked.connect(self._auto_pdf_folder)
        actions_row.addWidget(pdf_btn)
        actions_row.addStretch(1)
        source_card.body.addLayout(actions_row)
        source_card.body.addWidget(_muted(
            "داده‌ها فقط خوانده می‌شوند و برنامه هیچ تغییری در دیتابیس یا فایل اصلی نمی‌دهد."
        ))
        right_column.addWidget(source_card)

        # ---- بارگذاری فایل اکسل ----
        import_card = Card(colors, radius=16, padding=18)
        import_card.body.addWidget(_card_title(
            "بارگذاری فایل ۰۱ (خروجی سامانه)",
            "همین یک دکمه: انتخاب فایل، به‌روزرسانی کتاب کار، ذخیرهٔ خودکار در دیتابیس",
        ))
        excel_row = QHBoxLayout()
        excel_row.setSpacing(8)
        self.excel_edit = QLineEdit(str(config.get("last_excel_path") or ""))
        self.excel_edit.setPlaceholderText("مسیر فایل \u202a.xlsx\u202c")
        self.excel_edit.setMinimumWidth(96)
        excel_row.addWidget(self.excel_edit, 1)
        browse_excel = QPushButton("انتخاب فایل…")
        browse_excel.setObjectName("Ghost")
        browse_excel.setIcon(make_icon("folder", colors["ink_2"], 15))
        browse_excel.setCursor(Qt.PointingHandCursor)
        browse_excel.clicked.connect(self._browse_excel)
        excel_row.addWidget(browse_excel)
        import_card.body.addLayout(excel_row)
        self.import_button = QPushButton("بارگذاری در دیتابیس")
        self.import_button.setObjectName("Primary")
        self.import_button.setIcon(make_icon("upload", colors["on_brand"], 15))
        self.import_button.setCursor(Qt.PointingHandCursor)
        self.import_button.clicked.connect(self._emit_import)
        self.local_button = QPushButton("اجرای محلی (بدون دیتابیس)")
        self.local_button.setObjectName("Soft")
        self.local_button.setIcon(make_icon("table", colors["ink_2"], 15))
        self.local_button.setCursor(Qt.PointingHandCursor)
        self.local_button.setToolTip(
            "فایل اکسل را مستقیم از دیسک می‌خواند؛ به MySQL نیازی نیست."
        )
        self.local_button.clicked.connect(self._emit_local)
        self.update_button = QPushButton("بارگذاری فایل ۰۱ و به‌روزرسانی")
        self.update_button.setObjectName("Primary")
        self.update_button.setIcon(make_icon("refresh", colors["on_brand"], 15))
        self.update_button.setCursor(Qt.PointingHandCursor)
        self.update_button.setToolTip(
            "همین یک دکمه کافی است: فایل خروجی سامانه (Contract-Info-Report) را "
            "انتخاب می‌کنید؛ برنامه ردیف‌های «اجرا - مهر» را جدا می‌کند، ستون‌های "
            "L تا AJ را روی کتاب کار می‌نویسد، فایل را با اکسل بازمحاسبه می‌کند و "
            "خودکار در دیتابیس ذخیره می‌کند و داشبورد را بازخوانی می‌کند."
        )
        self.update_button.clicked.connect(self._emit_system_update)
        import_row = QHBoxLayout()
        import_row.setSpacing(8)
        self.diagnose_button = QPushButton("تشخیص فایل ۰۱")
        self.diagnose_button.setObjectName("Ghost")
        self.diagnose_button.setCursor(Qt.PointingHandCursor)
        self.diagnose_button.setToolTip(
            "گزارش می‌دهد برنامه کدام شیت، کدام ستون قرارداد و چند علامت وضعیت (M) "
            "در فایل پیدا کرده — برای وقتی آمار صفر می‌ماند."
        )
        self.diagnose_button.clicked.connect(self._emit_diagnose)
        self.source_book_button = QPushButton("کتاب کار اصلی (۰۲)")
        self.source_book_button.setObjectName("Ghost")
        self.source_book_button.setCursor(Qt.PointingHandCursor)
        self.source_book_button.setToolTip(
            "فایل اصلی خودتان (همان اکسلی که شیت colorcode دارد) را یک بار مشخص کنید؛ "
            "به‌روزرسانی‌های بعدی روی همان انجام می‌شود و اعداد داشبورد مثل قبل می‌مانند."
        )
        self.source_book_button.clicked.connect(self.sourceWorkbookRequested.emit)
        import_row.addWidget(self.update_button, 0)
        import_row.addStretch(1)
        import_card.body.addLayout(import_row)
        import_card.body.addWidget(_muted(
            "همین یک دکمه کافی است: فایل ۰۱ (خروجی سامانه) را انتخاب کنید. برنامه "
            "خودش فیلتر «اجرا - مهر» را می‌زند، ستون‌های L تا AJ را روی کتاب کار "
            "می‌نویسد، با اکسل بازمحاسبه می‌کند، در دیتابیس ذخیره می‌کند و داشبورد "
            "را بازخوانی می‌کند — بدون فایل دوم و بدون مرحلهٔ دستی."
        ))
        import_card.body.addWidget(_muted(
            "دیتابیس و جدول‌ها خودکار ساخته می‌شوند. اگر MySQL در دسترس نبود، "
            "همان نسخهٔ به‌روز محلی خوانده می‌شود تا داشبورد خالی نماند."
        ))
        right_column.addWidget(import_card)

        # ---- فایل‌های برنامه ----
        storage_card = Card(colors, radius=16, padding=18)
        storage_card.body.addWidget(_card_title("فایل‌های برنامه", "تنظیمات، یادداشت‌ها و نسخه"))
        self.config_path_label = QLabel(f"تنظیمات: \u202a{config.path}\u202c")
        self.config_path_label.setObjectName("Muted")
        self.config_path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.notes_path_label = QLabel(
            f"یادداشت‌ها: \u202a{data_dir() / NotesStore.FILENAME}\u202c"
        )
        self.notes_path_label.setObjectName("Muted")
        self.notes_path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        storage_card.body.addWidget(self.config_path_label)
        storage_card.body.addWidget(self.notes_path_label)
        self.workbook_path_label = QLabel(
            f"کتاب کار داخلی: \u202a{managed_workbook_path()}\u202c"
        )
        self.workbook_path_label.setObjectName("Muted")
        self.workbook_path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.workbook_path_label.setToolTip(
            "کتابی که برنامه از فایل ۰۱ سامانه می‌سازد و به‌روز می‌کند؛ برای اجرا "
            "بدون دیتابیس همین فایل استفاده می‌شود."
        )
        storage_card.body.addWidget(self.workbook_path_label)
        self.version_label = QLabel(f"نسخهٔ برنامه: {fa(APP_VERSION)}")
        self.version_label.setObjectName("Muted")
        self.loaded_label = QLabel("آخرین بارگذاری: —")
        self.loaded_label.setObjectName("Muted")
        storage_card.body.addWidget(self.version_label)
        storage_card.body.addWidget(self.loaded_label)
        open_folder = QPushButton("باز کردن پوشهٔ برنامه")
        open_folder.setIcon(make_icon("folder", colors["ink_2"], 15))
        open_folder.setCursor(Qt.PointingHandCursor)
        open_folder.clicked.connect(self.openDataFolderRequested.emit)
        storage_card.body.addWidget(open_folder, 0, Qt.AlignRight)
        right_column.addWidget(storage_card)

        # ---- وضعیت داده ----
        self.summary_card = Card(colors, radius=16, padding=18)
        self.summary_card.body.addWidget(_card_title("وضعیت داده", "خلاصهٔ آخرین بارگذاری"))
        summary_grid = QGridLayout()
        summary_grid.setHorizontalSpacing(8)
        summary_grid.setVerticalSpacing(8)
        self._summary_values: Dict[str, QLabel] = {}
        for index, (key, title) in enumerate((
            ("contracts", "کل قراردادها"),
            ("deductions", "دارای کسورات"),
            ("pdfs", "دارای فایل PDF"),
            ("follow", "نیازمند پیگیری"),
        )):
            box = QFrame()
            box.setObjectName("MiniStat")
            box.setStyleSheet(
                f"QFrame#MiniStat {{ background: {colors['surface_2']};"
                f" border-radius: 10px; }}"
            )
            box_layout = QVBoxLayout(box)
            box_layout.setContentsMargins(8, 6, 8, 6)
            box_layout.setSpacing(0)
            caption = QLabel(title)
            caption.setObjectName("Muted")
            caption.setStyleSheet("font-size: 7pt;")
            value = QLabel("—")
            value.setStyleSheet(
                f"color: {colors['ink']}; font-size: 13pt; font-weight: 800;"
            )
            box_layout.addWidget(caption)
            box_layout.addWidget(value)
            self._summary_values[key] = value
            summary_grid.addWidget(box, 0, index)
        for column in range(4):
            summary_grid.setColumnStretch(column, 1)
        self.summary_card.body.addLayout(summary_grid)
        right_column.addWidget(self.summary_card)
        right_column.addStretch(1)

        # ---- نمایش ----
        display_card = Card(colors, radius=16, padding=18)
        display_card.body.addWidget(_card_title("نمایش", "پوسته، اندازهٔ قلم و بازخوانی خودکار"))
        form2 = QGridLayout()
        form2.setHorizontalSpacing(8)
        form2.setVerticalSpacing(10)
        form2.addWidget(QLabel("پوسته:"), 0, 0)
        self.theme_combo = Combo()
        self.theme_combo.set_theme_colors(colors)
        self.theme_combo.addItem("روشن", "light")
        self.theme_combo.addItem("تیره", "dark")
        current_theme = str(config.get("theme") or "light")
        self.theme_combo.setCurrentIndex(max(0, self.theme_combo.findData(current_theme)))
        form2.addWidget(self.theme_combo, 0, 1)

        form2.addWidget(QLabel("اندازهٔ قلم:"), 1, 0)
        self.font_spin = NumberStepper(colors)
        self.font_spin.setRange(8, 16)
        self.font_spin.setSuffix(" pt")
        self.font_spin.setValue(int(config.get("font_size") or 10))
        form2.addWidget(self.font_spin, 1, 1)

        form2.addWidget(QLabel("بازخوانی خودکار:"), 2, 0)
        self.auto_spin = NumberStepper(colors)
        self.auto_spin.setRange(0, 120)
        self.auto_spin.setSuffix(" دقیقه")
        self.auto_spin.setSpecialValueText("غیرفعال")
        self.auto_spin.setValue(int(config.get("auto_refresh_minutes") or 5))
        form2.addWidget(self.auto_spin, 2, 1)
        form2.setColumnStretch(1, 1)
        display_card.body.addLayout(form2)

        display_card.body.addSpacing(2)

        left_column.addWidget(display_card)

        # ---- نکته‌ها ----
        tips_card = Card(colors, radius=16, padding=18)
        tips_card.body.addWidget(_card_title("نکته‌های کاربردی", "برای کار سریع‌تر با برنامه"))
        for kind, title, text in (
            ("refresh", "بازخوانی خودکار",
             "هر بار که کتاب کار تازه‌ای در دیتابیس بارگذاری شود، برنامه خودکار به‌روز می‌شود."),
            ("clock", "میان‌برها",
             "Ctrl+R بازخوانی، Ctrl+E خروجی اکسل، Ctrl+D تغییر پوسته و Ctrl+F جستجو."),
            ("shield", "اتصال امن",
             "میزبان، دیتابیس و کاربر از «مدیریت دیتابیس مرکزی» خوانده می‌شود؛ داده‌ها فقط خواندنی‌اند."),
        ):
            row = QHBoxLayout()
            row.setSpacing(10)
            row.addWidget(IconTile(kind, colors["brand"], colors["brand_soft"], size=34, radius=10))
            box = QVBoxLayout()
            box.setSpacing(0)
            heading = QLabel(title)
            heading.setStyleSheet("font-weight: 700;")
            body = QLabel(text)
            body.setObjectName("Muted")
            body.setWordWrap(True)
            box.addWidget(heading)
            box.addWidget(body)
            row.addLayout(box, 1)
            tips_card.body.addLayout(row)
        left_column.addWidget(tips_card)
        left_column.addStretch(1)

        holder.addLayout(right_column, 3)
        holder.addLayout(left_column, 2)

        save_bar = QHBoxLayout()
        save_bar.setSpacing(8)
        self.status_label = QLabel("")
        self.status_label.setObjectName("Muted")
        save_bar.addWidget(self.status_label, 1)
        self.save_button = QPushButton("ذخیره و اعمال")
        self.save_button.setObjectName("Primary")
        self.save_button.setIcon(make_icon("check", colors["on_brand"], 15))
        self.save_button.setCursor(Qt.PointingHandCursor)
        self.save_button.clicked.connect(self._emit_save)
        save_bar.addWidget(self.save_button)
        self.reload_button = QPushButton("بارگذاری مجدد داده‌ها")
        self.reload_button.setObjectName("Ghost")
        self.reload_button.setIcon(make_icon("refresh", colors["ink_2"], 15))
        self.reload_button.clicked.connect(self.reloadRequested.emit)
        save_bar.addWidget(self.reload_button)
        layout.addLayout(save_bar)
        self.refresh_pdf_note()

    def _browse_excel(self) -> None:
        start = self.excel_edit.text().strip() or str(app_dir())
        path, _selected = QFileDialog.getOpenFileName(
            self, "انتخاب فایل اکسل", start,
            "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
        )
        if path:
            self.excel_edit.setText(path)
            self.set_status("فایل انتخاب شد — «بارگذاری در دیتابیس» را بزنید.")

    def _emit_import(self) -> None:
        path = self.excel_edit.text().strip()
        if not path:
            self._browse_excel()
            path = self.excel_edit.text().strip()
        if path:
            self.set_status("در حال بارگذاری فایل اکسل در دیتابیس…")
            self.importRequested.emit(path)

    def _emit_local(self) -> None:
        path = self.excel_edit.text().strip()
        if not path:
            self._browse_excel()
            path = self.excel_edit.text().strip()
        if path:
            self.set_status("در حال خواندن فایل اکسل بدون دیتابیس…")
            self.localRunRequested.emit(path)

    def _emit_diagnose(self) -> None:
        path = self.excel_edit.text().strip()
        if not path:
            self._browse_excel()
            path = self.excel_edit.text().strip()
        if path:
            self.diagnoseRequested.emit(path)

    def _emit_system_update(self) -> None:
        path = self.excel_edit.text().strip()
        if not path:
            self._browse_excel()
            path = self.excel_edit.text().strip()
        if path:
            self.set_status("در حال خواندن فایل سامانه و تطبیق با کتاب کار…")
            self.systemUpdateRequested.emit(path)

    def set_importing(self, active: bool, path: str = "") -> None:
        """قفل‌کردن دکمهٔ بارگذاری هنگام اجرای عملیات."""
        if path:
            self.excel_edit.setText(path)
        self.import_button.setEnabled(not active)
        self.import_button.setText(
            "در حال بارگذاری…" if active else "بارگذاری در دیتابیس"
        )

    def refresh_pdf_note(self) -> None:
        folder = str(self.pdf_edit.text() or "").strip() or resolve_pdf_folder(self.config)
        if not folder:
            self.pdf_note.setText("پوشهٔ PDF پیدا نشد؛ با «انتخاب…» مسیر آن را بدهید.")
            return
        if not os.path.isdir(folder):
            self.pdf_note.setText(f"این مسیر وجود ندارد: {folder}")
            return
        total = len(scan_pdf_files(folder, depth=2, limit=5000))
        self.pdf_note.setText(
            f"\u202a{folder}\u202c — {fa(total)} فایل PDF شناسایی شد"
        )

    def _auto_pdf_folder(self) -> None:
        folder = resolve_pdf_folder(self.config)
        if not folder:
            self.pdf_note.setText("پوشهٔ PDF پیدا نشد؛ با «انتخاب…» مسیر آن را بدهید.")
            return
        self.pdf_edit.setText(folder)
        self.refresh_pdf_note()
        self.set_status("پوشهٔ PDF انتخاب شد — «ذخیره و اعمال» را بزنید.")

    def _browse_pdf(self) -> None:
        start = self.pdf_edit.text() or str(app_dir())
        folder = QFileDialog.getExistingDirectory(self, "انتخاب پوشهٔ PDF", start)
        if folder:
            self.pdf_edit.setText(folder)
            self.refresh_pdf_note()
            self.set_status("پوشهٔ PDF انتخاب شد — «ذخیره و اعمال» را بزنید.")

    def connection_values(self) -> Dict[str, Any]:
        """مقادیر فرم اتصال برای آزمایش/ذخیره."""
        return {
            "host": self.host_edit.text().strip() or "127.0.0.1",
            "port": int(self.port_spin.value()),
            "database": self.db_edit.text().strip() or "master_data",
            "user": self.user_edit.text().strip() or "master_app",
            "password": self.password_edit.text(),
            "workbook_id": int(self.workbook_spin.value()),
            "sheet_name": self.sheet_edit.text().strip() or SHEET_NAME_DEFAULT,
        }

    def apply_connection(self, values: Dict[str, Any]) -> None:
        """پر کردن فرم از تنظیمات خوانده‌شده (مدیریت دیتابیس مرکزی)."""
        if values.get("host"):
            self.host_edit.setText(str(values["host"]))
        if values.get("port"):
            self.port_spin.setValue(int(values["port"]))
        if values.get("database"):
            self.db_edit.setText(str(values["database"]))
        if values.get("user"):
            self.user_edit.setText(str(values["user"]))
        if "password" in values:
            self.password_edit.setText(str(values.get("password") or ""))
        self.remember_check.setChecked(bool(self.config.get("db_remember_password")))

    def _emit_save(self) -> None:
        values = self.connection_values()
        values.update({
            "db_host": values.pop("host"),
            "db_port": values.pop("port"),
            "db_name": values.pop("database"),
            "db_user": values.pop("user"),
            "db_password": values.pop("password"),
            "db_workbook_id": values.pop("workbook_id"),
            "db_remember_password": self.remember_check.isChecked(),
            "sheet_name": values.pop("sheet_name"),
            "pdf_folder": self.pdf_edit.text().strip(),
            "theme": self.theme_combo.currentData(),
            "font_size": self.font_spin.value(),
            "auto_refresh_minutes": self.auto_spin.value(),
        })
        self.saveRequested.emit(values)

    def update_data(self, data: DashboardData) -> None:
        """به‌روزرسانی کارت «وضعیت داده» پس از هر بارگذاری."""
        if not hasattr(self, "_summary_values"):
            return
        pdf_count = sum(1 for record in data.records if record.pdf_file)
        follow = sum(1 for record in data.records
                     if record.open_flags or record.is_expired)
        values = {
            "contracts": fa(data.total_contracts),
            "deductions": fa(data.contracts_with_deductions),
            "pdfs": fa(pdf_count),
            "follow": fa(follow),
        }
        for key, label in self._summary_values.items():
            label.setText(values.get(key, "—"))

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)

    def set_loaded(self, text: str, count: int = 0) -> None:
        suffix = f"  •  {fa(count)} قرارداد" if count else ""
        self.loaded_label.setText(f"آخرین بارگذاری: {text}{suffix}")

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.theme_combo.set_theme_colors(colors)
        for card in self.findChildren(Card):
            card.set_theme(colors)
        for stepper in self.findChildren(NumberStepper):
            stepper.set_theme(colors)
        for tile in self.findChildren(IconTile):
            tile.set_colors(colors["brand"], colors["brand_soft"])

    def load_values(self, config: AppConfig) -> None:
        self.host_edit.setText(str(config.get("db_host") or "127.0.0.1"))
        self.port_spin.setValue(int(config.get("db_port") or 3306))
        self.db_edit.setText(str(config.get("db_name") or "master_data"))
        self.user_edit.setText(str(config.get("db_user") or "master_app"))
        if not self.password_edit.text():
            self.password_edit.setText(str(config.get_password() or ""))
        self.remember_check.setChecked(bool(config.get("db_remember_password")))
        self.workbook_spin.setValue(int(config.get("db_workbook_id") or 1))
        self.sheet_edit.setText(str(config.get("sheet_name") or SHEET_NAME_DEFAULT))
        self.pdf_edit.setText(str(config.get("pdf_folder") or ""))
        self.refresh_pdf_note()
        theme_index = self.theme_combo.findData(str(config.get("theme") or "light"))
        self.theme_combo.setCurrentIndex(max(0, theme_index))
        self.font_spin.setValue(int(config.get("font_size") or 10))
        self.auto_spin.setValue(int(config.get("auto_refresh_minutes") or 0))


# ---------------------------------------------------------------------------
# راهنما
# ---------------------------------------------------------------------------
class HelpPage(QWidget):
    """راهنمای شروع سریع، میان‌برها و نگاشت امکانات نسخهٔ وب."""

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(self.scroll)
        content = QWidget()
        self.scroll.setWidget(content)
        self.layout = QVBoxLayout(content)
        self.layout.setContentsMargins(30, 22, 30, 26)
        self.layout.setSpacing(14)

        row = QHBoxLayout()
        row.setSpacing(14)

        start_card = Card(colors, radius=16, padding=18)
        start_card.body.addWidget(_card_title("شروع سریع", "چهار گام تا اولین گزارش"))
        steps = [
            "در صفحهٔ تنظیمات، میزبان/دیتابیس/کاربر MySQL را وارد کنید؛ اگر کاربر رمز ندارد "
            "فیلد رمز را خالی بگذارید (رمز اجباری نیست).",
            "در همان صفحه، فایل اکسل را با «بارگذاری فایل اکسل» انتخاب و در دیتابیس ذخیره کنید؛ "
            "دیتابیس و جدول‌ها در صورت نبود خودکار ساخته می‌شوند.",
            "پوشهٔ PDF را انتخاب کنید (برنامه خودش هم کنار فایل اکسل و کنار برنامه "
            "می‌گردد)؛ نام هر PDF باید شمارهٔ قرارداد باشد، مثل MEHR-003-1402.pdf — "
            "پسوندهایی مثل «نسخهٔ نهایی» یا «(1)» هم مشکلی ایجاد نمی‌کند. سپس Ctrl+R را بزنید.",
            "روی هر ردیف جدول دو کلیک کنید تا جزئیات، یادداشت و پیگیری باز شود.",
            "هر بار فایل ۰۱ سامانه را گرفتید، همان فایل را با «به‌روزرسانی از فایل سامانه "
            "(01)» بدهید؛ برنامه ردیف‌های «اجرا - مهر» ستون Sub Project را جدا می‌کند و "
            "ستون‌های L تا AJ را در شیت «خروجی کندو» (شیت سامانه) در یک نسخهٔ تازه می‌نویسد "
            "— فایل اصلی شما بازنویسی نمی‌شود.",
        ]
        for number, text in enumerate(steps, start=1):
            start_card.body.addWidget(page_step_label(text, number, colors))
        start_card.body.addStretch(1)
        row.addWidget(start_card, 1)

        shortcuts_card = Card(colors, radius=16, padding=18)
        shortcuts_card.body.addWidget(_card_title("میان‌برها", "دسترسی سریع از هر صفحه"))
        self._key_chips: List[QLabel] = []
        for keys, description in (
            ("Ctrl + R", "بازخوانی از دیتابیس"),
            ("Ctrl + F", "جستجو در جدول قراردادها"),
            ("Ctrl + E", "خروجی اکسل جدول"),
            ("Ctrl + D", "تغییر پوستهٔ روشن/تیره"),
            ("F1", "باز شدن راهنما"),
            ("Esc", "بستن پنل جزئیات"),
        ):
            key_chip = QLabel(keys)
            key_chip.setStyleSheet(
                f"background: {colors['surface_3']}; color: {colors['ink_2']};"
                f"border-radius: 7px; padding: 4px 9px; font-weight: 700; font-size: 8pt;"
            )
            key_chip.setFixedWidth(96)
            key_chip.setAlignment(Qt.AlignCenter)
            self._key_chips.append(key_chip)
            shortcut_row = QWidget()
            shortcut_layout = QHBoxLayout(shortcut_row)
            shortcut_layout.setContentsMargins(0, 2, 0, 2)
            shortcut_layout.setSpacing(10)
            shortcut_layout.addWidget(key_chip)
            label = QLabel(description)
            label.setObjectName("Subtle")
            shortcut_layout.addWidget(label, 1)
            shortcuts_card.body.addWidget(shortcut_row)
        shortcuts_card.body.addStretch(1)
        row.addWidget(shortcuts_card, 1)
        self.layout.addLayout(row)

        mapping_card = Card(colors, radius=16, padding=18)
        mapping_card.body.addWidget(_card_title(
            "نگاشت امکانات نسخهٔ وب به دسکتاپ",
            "همان محتوا و رفتار، این بار به‌صورت یک برنامهٔ مستقل",
        ))
        rows = [
            ("صفحهٔ اصلی و کارت‌های آماری", "داشبورد با کاشی‌های قابل کلیک"),
            ("فیلتر با کلیک روی کارت‌ها", "فیلتر جدول با کلیک روی کاشی، برش دونات یا میله"),
            ("نمودارهای Chart.js", "نمودارهای بومی Qt با تعامل و راهنما"),
            ("جدول قراردادها", "جدول با جستجو، مرتب‌سازی، مخفی‌کردن ستون و انتخاب چندردیفی"),
            ("پنجرهٔ جزئیات و یادداشت", "کشوی جزئیات در سمت راست پنجره"),
            ("خروجی اکسل مرورگر", "خروجی اکسل با انتخاب ستون و دامنه"),
            ("گزارش مدیریتی", "پیش‌نمایش چاپی و خروجی سه‌شیتی اکسل"),
            ("اتصال فایل PDF", "ستون فایل و باز کردن با برنامهٔ پیش‌فرض"),
            ("چت Ollama و ورود کاربران", "حذف شده؛ برنامهٔ تک‌کاربره و مستقل است"),
        ]
        for index, (web, desktop) in enumerate(rows):
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(2, 5, 2, 5)
            row_layout.setSpacing(10)
            draw = IconWidget("check" if index < 8 else "close",
                              colors["success"] if index < 8 else colors["danger"], 15)
            row_layout.addWidget(draw, 0, Qt.AlignTop)
            text = QLabel(
                f"<span style='color:{colors['ink']}; font-weight:600;'>"
                f"{html_module.escape(desktop)}</span>"
                f"<span style='color:{colors['ink_3']};'> &nbsp;←&nbsp; "
                f"{html_module.escape(web)}</span>"
            )
            text.setWordWrap(True)
            text.setAlignment(Qt.AlignRight | Qt.AlignVCenter | Qt.AlignAbsolute)
            row_layout.addWidget(text, 1)
            mapping_card.body.addWidget(row_widget)
            if index < len(rows) - 1:
                line = QFrame()
                line.setFixedHeight(1)
                line.setStyleSheet(f"background: {colors['border_soft']};")
                mapping_card.body.addWidget(line)
        mapping_card.body.addWidget(_card_title("یادداشت"))
        note = _muted(
            f"تنظیمات در mehr_qt_config.json و یادداشت‌ها در mehr_notes.json کنار برنامه ذخیره می‌شوند. "
            f"پیش‌نیاز اجرا: pip install PyQt5 openpyxl mysql-connector-python — نسخهٔ {APP_VERSION}"
        )
        mapping_card.body.addWidget(note)
        self.layout.addWidget(mapping_card)
        self.layout.addStretch(1)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        for card in self.findChildren(Card):
            card.set_theme(colors)
        for chip in self._key_chips:
            chip.setStyleSheet(
                f"background: {colors['surface_3']}; color: {colors['ink_2']};"
                f"border-radius: 7px; padding: 4px 9px; font-weight: 700; font-size: 8pt;"
            )
        self.scroll.viewport().update()



# ==========================================================================
# پوسته، پنجرهٔ اصلی و اجرای برنامه
# ==========================================================================
# ---------------------------------------------------------------------------
# پوستهٔ اصلی: ریل ناوبری، نوار بالا، کشوی جزئیات و پنجرهٔ برنامه
# ---------------------------------------------------------------------------
try:  # پشتیبانی چاپ اختیاری است
    from PyQt5.QtPrintSupport import QPrintDialog, QPrinter  # noqa: E402
except Exception:  # pragma: no cover
    QPrintDialog = None  # type: ignore
    QPrinter = None  # type: ignore


NAV_ITEMS: List[Tuple[str, str, str]] = [
    ("dashboard", "داشبورد", "dashboard"),
    ("contracts", "قراردادها", "table"),
    ("reports", "گزارش", "report"),
    ("pdfs", "PDF", "pdf"),
    ("settings", "تنظیمات", "settings"),
    ("help", "راهنما", "help"),
]


def install_crash_logging() -> None:
    """نصب قلاب‌های ثبت خطا تا برنامه بی‌صدا بسته نشود."""
    def _hook(exc_type, exc, tb):
        detail = "".join(traceback.format_exception(exc_type, exc, tb))
        log_event("Unhandled exception:\n" + detail)
        sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = _hook
    try:
        threading.excepthook = lambda args: log_event(
            "Unhandled thread exception:\n"
            + "".join(traceback.format_exception(args.exc_type, args.exc_value,
                                                 args.exc_traceback))
        )
    except Exception:
        pass
    try:
        handle = open(log_path(), "a", encoding="utf-8")
        faulthandler.enable(file=handle, all_threads=True)
    except Exception:
        pass
    log_event("قلاب‌های ثبت خطا فعال شد")


def open_path(target: str) -> None:
    """باز کردن پوشه یا فایل در برنامهٔ پیش‌فرض سیستم."""
    path = Path(str(target or "")).expanduser()
    if not path.exists():
        path = path.parent
    QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


def guarded(name: str):
    """تزئین‌کنندهٔ اسلات‌های Qt: خطا را ثبت می‌کند و برنامه را باز نگه می‌دارد."""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(self, *args, **kwargs):
            try:
                return func(self, *args, **kwargs)
            except Exception:
                detail = traceback.format_exc()
                log_event(f"{name} failed:\n{detail}")
                try:
                    QMessageBox.critical(
                        self, "خطا در برنامه",
                        "یک خطای پیش‌بینی‌نشده رخ داد؛ برنامه بسته نشد.\n"
                        f"جزئیات در این فایل ذخیره شد:\n{log_path()}",
                    )
                except Exception:
                    pass
                return None
        return wrapper
    return decorator


class LogoMark(QWidget):
    """نشان برنامه: مربع گرد با گرادیان برند و آیکون لایه‌ها."""

    def __init__(self, colors: Dict[str, str], size: int = 44, parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.setFixedSize(size, size)

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        body = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        gradient = QLinearGradient(body.topLeft(), body.bottomRight())
        gradient.setColorAt(0.0, QColor("#8E7BFF"))
        gradient.setColorAt(0.55, QColor("#6A5BF2"))
        gradient.setColorAt(1.0, QColor("#3F35AC"))
        painter.setPen(Qt.NoPen)
        painter.setBrush(gradient)
        painter.drawRoundedRect(body, body.width() * 0.32, body.width() * 0.32)
        painter.setPen(QPen(QColor(255, 255, 255, 46), 1))
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(body.adjusted(0.5, 0.5, -0.5, -0.5),
                                body.width() * 0.32, body.width() * 0.32)
        painter.setPen(QPen(QColor(255, 255, 255, 235)))
        font = app_font(max(11, int(body.width() * 0.42)), QFont.Black)
        painter.setFont(font)
        painter.drawText(body.adjusted(0, -body.height() * 0.04, 0, 0),
                         Qt.AlignCenter, "م")


class RailItem(QWidget):
    """یک دکمهٔ ناوبری در ریل کنار با آیکون و برچسب."""

    clicked = pyqtSignal(str)
    HEIGHT = 60

    def __init__(self, key: str, label: str, kind: str,
                 colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.key = key
        self.kind = kind
        self.label = label
        self.colors = colors
        self.active = False
        self._hover = False
        self.badge = 0
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(self.HEIGHT)
        self.setAttribute(Qt.WA_Hover, True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 8, 0, 6)
        layout.setSpacing(4)
        self.icon = IconWidget(kind, colors["rail_muted"], 20, 1.7)
        layout.addWidget(self.icon, 0, Qt.AlignHCenter)
        self.label_widget = QLabel(label)
        self.label_widget.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label_widget)
        self._sync()

    def set_active(self, active: bool) -> None:
        self.active = active
        self._sync()

    def set_badge(self, count: int) -> None:
        self.badge = max(0, int(count))
        self.update()

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self._sync()

    def _sync(self) -> None:
        color = self.colors["rail_ink"] if (self.active or self._hover) else self.colors["rail_muted"]
        self.icon.set_color(color)
        self.label_widget.setStyleSheet(
            f"color: {color}; font-size: 7.5pt; background: transparent;"
            f"font-weight: {800 if self.active else 500};"
        )
        self.update()

    def enterEvent(self, event) -> None:  # noqa: N802
        self._hover = True
        self._sync()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._hover = False
        self._sync()
        super().leaveEvent(event)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        body = QRectF(self.rect()).adjusted(9, 3, -9, -3)
        if self.active:
            gradient = QLinearGradient(body.topLeft(), body.bottomLeft())
            gradient.setColorAt(0.0, QColor(255, 255, 255, 26))
            gradient.setColorAt(1.0, QColor(255, 255, 255, 8))
            painter.setPen(QPen(QColor(255, 255, 255, 24), 1))
            painter.setBrush(gradient)
            painter.drawRoundedRect(body, 15, 15)
            # نشانگر لبهٔ داخلی + هالهٔ برند
            indicator = QRectF(body.left() + 1.5, body.center().y() - 11, 3, 22)
            painter.setPen(Qt.NoPen)
            glow = QColor(self.colors["brand_2"])
            glow.setAlpha(70)
            painter.setBrush(glow)
            painter.drawRoundedRect(indicator.adjusted(-2.5, -2, 2.5, 2), 4, 4)
            painter.setBrush(QColor(self.colors["brand_2"]))
            painter.drawRoundedRect(indicator, 1.5, 1.5)
        elif self._hover:
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(255, 255, 255, 12))
            painter.drawRoundedRect(body, 15, 15)

        if self.badge > 0:
            text = fa(str(self.badge))
            painter.setFont(app_font(7, QFont.Bold))
            metrics = painter.fontMetrics()
            width = max(16.0, metrics.width(text) + 9)
            pill = QRectF(body.right() - width - 2, body.top() + 1, width, 16)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(self.colors["accent"]))
            painter.drawRoundedRect(pill, 8, 8)
            painter.setPen(QPen(QColor("#231A05")))
            painter.drawText(pill, Qt.AlignCenter, text)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.key)
        super().mouseReleaseEvent(event)


class Rail(QFrame):
    """ریل ناوبری عمودی با نشان برند، ناوبری گروه‌بندی‌شده و پابرگ."""

    pageRequested = pyqtSignal(str)
    themeRequested = pyqtSignal()

    WIDTH = 94

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.setObjectName("Rail")
        self.setFixedWidth(self.WIDTH)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 18, 0, 14)
        layout.setSpacing(0)

        self.logo = LogoMark(colors, 44)
        layout.addWidget(self.logo, 0, Qt.AlignHCenter)
        title = QLabel("مهر")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(
            f"color: {colors['rail_ink']}; font-size: 10pt; font-weight: 800;"
            " background: transparent;"
        )
        layout.addSpacing(7)
        layout.addWidget(title)
        caption = QLabel("سامانهٔ قرارداد")
        caption.setAlignment(Qt.AlignCenter)
        caption.setStyleSheet(
            f"color: {colors['rail_muted']}; font-size: 6.5pt; background: transparent;"
        )
        layout.addSpacing(1)
        layout.addWidget(caption)
        self.title_label = title
        self.caption_label = caption

        layout.addSpacing(16)
        divider = QFrame()
        divider.setObjectName("RailDivider")
        divider.setFixedHeight(1)
        layout.addWidget(divider)
        layout.addSpacing(12)

        group = QLabel("پیمایش")
        group.setAlignment(Qt.AlignCenter)
        group.setStyleSheet(
            f"color: {colors['rail_muted']}; font-size: 6.5pt; font-weight: 700;"
            " letter-spacing: 1px; background: transparent;"
        )
        layout.addWidget(group)
        self.group_label = group
        layout.addSpacing(6)

        self.items: Dict[str, RailItem] = {}
        for key, label, kind in NAV_ITEMS:
            item = RailItem(key, label, kind, colors)
            item.clicked.connect(self.pageRequested.emit)
            layout.addWidget(item)
            self.items[key] = item

        layout.addStretch(1)

        self.theme_button = IconButton("", "moon", colors["rail_muted"])
        self.theme_button.setObjectName("RailUtility")
        self.theme_button.setFixedSize(38, 38)
        self.theme_button.setCursor(Qt.PointingHandCursor)
        self.theme_button.setToolTip("تغییر پوسته (Ctrl+D)")
        self.theme_button.clicked.connect(self.themeRequested.emit)
        layout.addWidget(self.theme_button, 0, Qt.AlignHCenter)
        layout.addSpacing(6)
        version = QLabel(f"نسخهٔ {fa(APP_VERSION)}")
        version.setAlignment(Qt.AlignCenter)
        version.setStyleSheet(
            f"color: {colors['rail_muted']}; font-size: 6.5pt; background: transparent;"
        )
        layout.addWidget(version)
        self.version_label = version

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        gradient = QLinearGradient(0, 0, 0, self.height())
        gradient.setColorAt(0.0, QColor(self.colors["rail_2"]))
        gradient.setColorAt(0.55, QColor(self.colors["rail"]))
        gradient.setColorAt(1.0, QColor(self.colors["rail"]))
        painter.fillRect(self.rect(), gradient)
        glow = QRadialGradient(QPointF(self.width() * 0.5, 70), self.width() * 2.2)
        brand = QColor(self.colors["brand"])
        brand.setAlpha(46)
        glow.setColorAt(0.0, brand)
        glow.setColorAt(1.0, QColor(0, 0, 0, 0))
        painter.fillRect(self.rect(), glow)
        painter.setPen(QPen(QColor(255, 255, 255, 16), 1))
        painter.drawLine(0, 0, 0, self.height())

    def set_active(self, key: str) -> None:
        for item_key, item in self.items.items():
            item.set_active(item_key == key)

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.logo.set_theme(colors)
        self.theme_button.set_icon_color(colors["rail_muted"])
        self.title_label.setStyleSheet(
            f"color: {colors['rail_ink']}; font-size: 10pt; font-weight: 800;"
            " background: transparent;"
        )
        self.caption_label.setStyleSheet(
            f"color: {colors['rail_muted']}; font-size: 6.5pt; background: transparent;"
        )
        self.group_label.setStyleSheet(
            f"color: {colors['rail_muted']}; font-size: 6.5pt; font-weight: 700;"
            " letter-spacing: 1px; background: transparent;"
        )
        self.version_label.setStyleSheet(
            f"color: {colors['rail_muted']}; font-size: 6.5pt; background: transparent;"
        )
        for item in self.items.values():
            item.set_theme(colors)
        self.update()

    def set_dark(self, dark: bool) -> None:
        self.theme_button.set_icon_kind("sun" if dark else "moon",
                                        self.colors["rail_muted"])


class TopBar(QFrame):
    """نوار بالای پنجره: عنوان صفحه، جستجو و دکمه‌های اقدام."""

    searchChanged = pyqtSignal(str)
    reloadRequested = pyqtSignal()
    themeToggled = pyqtSignal()
    menuRequested = pyqtSignal()

    HEIGHT = 96

    def __init__(self, colors: Dict[str, str], parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.setFixedHeight(self.HEIGHT)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(30, 15, 30, 12)
        layout.setSpacing(14)

        title_box = QVBoxLayout()
        title_box.setSpacing(3)
        title_box.setContentsMargins(0, 0, 0, 0)
        self.eyebrow = QLabel("پروژهٔ مهر خلیج فارس")
        self.eyebrow.setObjectName("Eyebrow")
        self.eyebrow.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.title = QLabel(PAGE_TITLES["dashboard"])
        self.title.setObjectName("PageTitle")
        self.title.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.subtitle = QLabel(PAGE_SUBTITLES["dashboard"])
        self.subtitle.setObjectName("PageSubtitle")
        self.subtitle.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        title_box.addWidget(self.eyebrow)
        title_box.addWidget(self.title)
        title_box.addWidget(self.subtitle)
        title_widget = QWidget()
        title_widget.setLayout(title_box)
        title_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(title_widget, 1)

        self.search_edit = SearchBox("جستجوی قرارداد، شرکت، موضوع…")
        self.search_edit.setMinimumWidth(300)
        self.search_edit.setMaximumWidth(430)
        self.search_edit.textChanged.connect(self.searchChanged.emit)
        layout.addWidget(self.search_edit, 0)

        self.status_pill = QLabel("در حال اتصال…")
        self.status_pill.setObjectName("StatusPill")
        self.status_pill.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_pill, 0)

        self.date_chip = QLabel(today_jalali(with_time=False))
        self.date_chip.setObjectName("Chip")
        layout.addWidget(self.date_chip, 0)

        self.reload_button = IconButton("", "refresh", colors["ink_2"])
        self.reload_button.setObjectName("Ghost")
        self.reload_button.setFixedSize(38, 38)
        self.reload_button.setToolTip("بازخوانی از دیتابیس (Ctrl+R)")
        self.reload_button.setCursor(Qt.PointingHandCursor)
        self.reload_button.clicked.connect(self.reloadRequested.emit)
        layout.addWidget(self.reload_button)

        self.theme_button = IconButton("", "moon", colors["ink_2"])
        self.theme_button.setObjectName("Ghost")
        self.theme_button.setFixedSize(38, 38)
        self.theme_button.setToolTip("تغییر پوسته (Ctrl+D)")
        self.theme_button.setCursor(Qt.PointingHandCursor)
        self.theme_button.clicked.connect(self.themeToggled.emit)
        layout.addWidget(self.theme_button)

        self.menu_button = IconButton("", "dots", colors["ink_2"])
        self.menu_button.setObjectName("Ghost")
        self.menu_button.setFixedSize(38, 38)
        self.menu_button.setToolTip("منوی برنامه")
        self.menu_button.setCursor(Qt.PointingHandCursor)
        self.menu_button.clicked.connect(self.menuRequested.emit)
        layout.addWidget(self.menu_button)

    def set_page(self, key: str) -> None:
        set_text_aligned(self.title, PAGE_TITLES.get(key, ""), "right")
        set_text_aligned(self.subtitle, PAGE_SUBTITLES.get(key, ""), "right")

    def set_theme(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self.status_pill.style().unpolish(self.status_pill)
        self.status_pill.style().polish(self.status_pill)
        self.reload_button.set_icon_color(colors["ink_2"])
        self.menu_button.set_icon_color(colors["ink_2"])
        self.theme_button.set_icon_color(colors["ink_2"])
        self.search_edit.set_icon_color(colors["ink_3"])
        self.date_chip.setStyleSheet("")
        self.update()

    def set_dark(self, dark: bool) -> None:
        self.theme_button.set_icon_kind("sun" if dark else "moon", self.colors["ink_2"])

    def set_connection(self, online: bool, text: str = "") -> None:
        """نشان وضعیت اتصال به دیتابیس در نوار بالا."""
        label = text or ("دیتابیس متصل" if online else "بدون اتصال")
        self.status_pill.setText(label)
        self.status_pill.setObjectName("StatusPill" if online else "StatusPillOff")
        self.status_pill.style().unpolish(self.status_pill)
        self.status_pill.style().polish(self.status_pill)
        self.status_pill.setToolTip(
            "اتصال به دیتابیس مرکزی برقرار است" if online
            else "اتصال به دیتابیس برقرار نیست — تنظیمات را بررسی کنید"
        )

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(self.colors["surface"]))
        painter.setPen(QPen(QColor(self.colors["border_soft"]), 1))
        painter.drawLine(0, self.height() - 1, self.width(), self.height() - 1)


class DetailDrawer(QWidget):
    """کشوی جزئیات قرارداد با مشخصات، وضعیت اسناد، فایل PDF و یادداشت‌ها."""

    closed = pyqtSignal()
    notesChanged = pyqtSignal(str)
    pdfRequested = pyqtSignal(str)
    revealRequested = pyqtSignal(str)
    bookmarkToggled = pyqtSignal(str)

    WIDTH = 440

    def __init__(self, colors: Dict[str, str], notes: NotesStore, parent=None) -> None:
        super().__init__(parent)
        self.colors = colors
        self.notes = notes
        self.record: Optional[ContractRecord] = None
        self._open = False
        self.setObjectName("Drawer")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(self.WIDTH)
        self.hide()
        self._apply_shell()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ---- سرصفحه ----
        header = QWidget()
        header.setObjectName("DrawerHead")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(18, 16, 16, 12)
        header_layout.setSpacing(10)
        self.avatar = Avatar("", 40)
        header_layout.addWidget(self.avatar)
        text_box = QVBoxLayout()
        text_box.setSpacing(1)
        self.name_label = QLabel("—")
        self.name_label.setStyleSheet("font-size: 11.5pt; font-weight: 800;")
        self.name_label.setWordWrap(True)
        self.meta_label = QLabel("—")
        self.meta_label.setObjectName("Muted")
        text_box.addWidget(self.name_label)
        text_box.addWidget(self.meta_label)
        header_layout.addLayout(text_box, 1)
        self.bookmark_button = IconButton("", "bookmark", colors["ink_3"])
        self.bookmark_button.setObjectName("Ghost")
        self.bookmark_button.setFixedSize(32, 32)
        self.bookmark_button.setToolTip("نشان‌گذاری")
        self.bookmark_button.setCursor(Qt.PointingHandCursor)
        self.bookmark_button.clicked.connect(self._toggle_bookmark)
        header_layout.addWidget(self.bookmark_button, 0, Qt.AlignTop)
        self.close_button = IconButton("", "close", colors["ink_3"])
        self.close_button.setObjectName("Ghost")
        self.close_button.setFixedSize(32, 32)
        self.close_button.setCursor(Qt.PointingHandCursor)
        self.close_button.clicked.connect(self.close_drawer)
        header_layout.addWidget(self.close_button, 0, Qt.AlignTop)
        outer.addWidget(header)

        divider = QFrame()
        divider.setFixedHeight(1)
        divider.setObjectName("DrawerLine")
        outer.addWidget(divider)

        # ---- بدنهٔ قابل اسکرول ----
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(self.scroll, 1)
        content = QWidget()
        content.setObjectName("DrawerBody")
        content.setLayoutDirection(Qt.RightToLeft)
        self.scroll.setWidget(content)
        body = QVBoxLayout(content)
        body.setContentsMargins(16, 12, 16, 18)
        body.setSpacing(10)

        self.chips = QHBoxLayout()
        self.chips.setSpacing(5)
        body.addLayout(self.chips)

        body.addWidget(self._section_title("مشخصات قرارداد"))
        self.info_box = QVBoxLayout()
        self.info_box.setSpacing(0)
        body.addLayout(self.info_box)

        body.addWidget(self._section_title("وضعیت اسناد اختتام"))
        self.flags_box = QVBoxLayout()
        self.flags_box.setSpacing(0)
        body.addLayout(self.flags_box)

        body.addWidget(self._section_title("فایل قرارداد"))
        pdf_card = QFrame()
        pdf_card.setObjectName("MiniCard")
        pdf_row = QHBoxLayout(pdf_card)
        pdf_row.setContentsMargins(10, 8, 10, 8)
        pdf_row.setSpacing(8)
        self.pdf_label = QLabel("فایلی متصل نیست.")
        self.pdf_label.setWordWrap(True)
        self.pdf_label.setObjectName("Muted")
        pdf_row.addWidget(self.pdf_label, 1)
        self.reveal_pdf_button = IconButton("", "folder", colors["ink_2"])
        self.reveal_pdf_button.setObjectName("Ghost")
        self.reveal_pdf_button.setFixedSize(30, 30)
        self.reveal_pdf_button.setToolTip("نمایش در پوشه")
        self.reveal_pdf_button.setCursor(Qt.PointingHandCursor)
        self.reveal_pdf_button.clicked.connect(self._reveal_pdf)
        pdf_row.addWidget(self.reveal_pdf_button)
        self.open_pdf_button = QPushButton("باز کردن")
        self.open_pdf_button.setObjectName("Ghost")
        self.open_pdf_button.setIcon(make_icon("external", colors["ink_2"], 14))
        self.open_pdf_button.setCursor(Qt.PointingHandCursor)
        self.open_pdf_button.clicked.connect(self._open_pdf)
        pdf_row.addWidget(self.open_pdf_button)
        body.addWidget(pdf_card)

        body.addWidget(self._section_title("یادداشت و پیگیری"))
        self.author_edit = QLineEdit()
        self.author_edit.setPlaceholderText("نام شما")
        self.follow_edit = QLineEdit()
        self.follow_edit.setPlaceholderText("تاریخ پیگیری (مثلاً ۱۴۰۵/۰۸/۱۲)")
        input_row = QHBoxLayout()
        input_row.setSpacing(8)
        input_row.addWidget(self.author_edit, 1)
        input_row.addWidget(self.follow_edit, 1)
        body.addLayout(input_row)
        self.note_edit = QPlainTextEdit()
        self.note_edit.setPlaceholderText("یادداشت این قرارداد…")
        self.note_edit.setFixedHeight(56)
        body.addWidget(self.note_edit)
        save_button = QPushButton("ثبت یادداشت")
        save_button.setObjectName("Primary")
        save_button.setIcon(make_icon("check", "#FFFFFF", 14))
        save_button.setCursor(Qt.PointingHandCursor)
        save_button.clicked.connect(self._save_note)
        body.addWidget(save_button)

        self.notes_box = QVBoxLayout()
        self.notes_box.setSpacing(6)
        body.addLayout(self.notes_box)
        self.notes_empty = QLabel("هنوز یادداشتی ثبت نشده است.")
        self.notes_empty.setObjectName("Muted")
        body.addWidget(self.notes_empty)
        body.addStretch(1)

        self._anim = QPropertyAnimation(self, b"pos", self)
        self._anim.setDuration(220)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self._anim.finished.connect(self._on_anim_finished)

    # ---- ظاهر ----
    def _apply_shell(self) -> None:
        c = self.colors
        self.setStyleSheet(
            f"QWidget#Drawer {{ background: {c['surface']};"
            f" border-left: 1px solid {c['border']}; }}"
            f"QWidget#DrawerHead {{ background: {c['surface']}; }}"
            f"QWidget#DrawerBody {{ background: {c['surface']}; }}"
            f"QFrame#DrawerLine {{ background: {c['border_soft']}; border: none; }}"
            f"QFrame#MiniCard {{ background: {c['surface_2']}; border-radius: 10px; }}"
        )

    def _section_title(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setStyleSheet(
            f"color: {self.colors['ink']}; font-size: 9.5pt; font-weight: 800;"
            f"padding-top: 6px;"
        )
        return label

    def _info_row(self, title: str, value: str, last: bool = False) -> QWidget:
        c = self.colors
        row = QFrame()
        row.setObjectName("InfoRow")
        border = "" if last else f"border-bottom: 1px solid {c['border_soft']};"
        row.setStyleSheet(f"QFrame#InfoRow {{ {border} }}")
        row.setLayoutDirection(Qt.RightToLeft)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 6, 0, 6)
        layout.setSpacing(10)
        key = QLabel(title)
        key.setObjectName("Muted")
        key.setFixedWidth(96)
        set_text_aligned(key, title, "right")
        key.setStyleSheet("font-size: 8pt;")
        layout.addWidget(key, 0, Qt.AlignTop)
        val = QLabel(value or "—")
        val.setWordWrap(True)
        val.setTextInteractionFlags(Qt.TextSelectableByMouse)
        val.setText(value or "—")
        val.setAlignment(Qt.AlignRight | Qt.AlignTop | Qt.AlignAbsolute)
        val.setMinimumWidth(0)
        val.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        val.setStyleSheet("font-size: 9pt; font-weight: 600;")
        layout.addWidget(val, 1)
        return row

    def _flag_row(self, label: str, state: bool, last: bool = False) -> QWidget:
        c = self.colors
        row = QFrame()
        row.setObjectName("FlagRow")
        border = "" if last else f"border-bottom: 1px solid {c['border_soft']};"
        row.setStyleSheet(f"QFrame#FlagRow {{ {border} }}")
        row.setLayoutDirection(Qt.RightToLeft)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 5, 0, 5)
        layout.setSpacing(8)
        chip = chip_label(
            "ثبت شده" if state else "نیازمند اقدام",
            c["success"] if state else c["warning"],
            c["success_soft"] if state else c["warning_soft"],
            radius=8, padding="2px 8px",
        )
        chip.setStyleSheet(chip.styleSheet() + "font-size: 7.6pt;")
        layout.addWidget(chip, 0)
        text = QLabel(label)
        set_text_aligned(text, label, "right")
        text.setWordWrap(True)
        text.setStyleSheet(
            f"color: {c['ink'] if state else c['ink_2']}; font-size: 9pt;"
        )
        layout.addWidget(text, 1)
        return row

    # ---- جای‌گیری ----
    def place(self) -> None:
        parent = self.parentWidget()
        if parent is None:
            return
        self.setGeometry(parent.width() - self.WIDTH, 0, self.WIDTH, parent.height())

    def open_drawer(self) -> None:
        self._open = True
        self.show()
        self.raise_()
        self.place()
        self._animate_to(self.parentWidget().width() - self.WIDTH)

    def close_drawer(self) -> None:
        if not self._open:
            self.hide()
            return
        self._open = False
        self._animate_to(self.parentWidget().width())

    def _animate_to(self, x: int) -> None:
        self._anim.stop()
        self._anim.setStartValue(self.pos())
        self._anim.setEndValue(QPoint(x, 0))
        self._anim.start()

    def _on_anim_finished(self) -> None:
        if not self._open:
            self.hide()
            self.closed.emit()

    def keyPressEvent(self, event) -> None:  # noqa: N802
        if event.key() == Qt.Key_Escape:
            self.close_drawer()
            return
        super().keyPressEvent(event)

    # ---- محتوا ----
    def set_record(self, record: ContractRecord) -> None:
        self.record = record
        self.avatar.set_name(record.col2 or "?")
        self.name_label.setText(record.col2 or "—")
        self.meta_label.setText(bidi(f"{record.col3} • ردیف {fa(record.excel_row)}"))
        clear_layout(self.chips)
        if record.is_expired:
            self.chips.addWidget(chip_label("الحاقیه منقضی", self.colors["danger"],
                                            self.colors["danger_soft"]))
        if record.f_value:
            self.chips.addWidget(chip_label(record.f_value, self.colors["brand_deep"],
                                            self.colors["brand_soft"]))
        if record.has_deductions:
            self.chips.addWidget(chip_label("دارای کسورات", self.colors["warning"],
                                            self.colors["warning_soft"]))
        self.chips.addStretch(1)
        self._sync_bookmark()
        self.pdf_label.setText(record.pdf_file or "فایلی متصل نیست.")
        self.open_pdf_button.setEnabled(bool(record.pdf_file))
        self.reveal_pdf_button.setEnabled(bool(record.pdf_file))

        clear_layout(self.info_box)
        rows = [
            ("شمارهٔ قرارداد", record.col3),
            ("موضوع", record.col4),
            ("نوع قرارداد", record.f_value),
            ("کسورات اسناد اختتام", record.deductions),
            ("پیشرفت کارکرد", pretty_label(record.progress_status)),
            ("پیشرفت اختتام", pretty_label(record.completion_status)),
            ("الحاقیه / زمان", record.annex_status),
            ("وضعیت پکیج آزادسازی", record.as_status),
            ("پیشرفت جدول", record.progress_table),
            ("اختتام جدول", record.completion_table),
            ("AW / AX / AY", f"{record.aw or '—'} / {record.ax or '—'} / {record.ay or '—'}"),
        ]
        for index, (title, value) in enumerate(rows):
            self.info_box.addWidget(self._info_row(title, str(value or "—"),
                                                   last=index == len(rows) - 1))

        clear_layout(self.flags_box)
        for index, (key, label, _column) in enumerate(FLAG_DEFS):
            self.flags_box.addWidget(self._flag_row(
                label, bool(record.flags.get(key)), last=index == len(FLAG_DEFS) - 1
            ))

        self._refresh_notes()
        self.author_edit.setText(str(self.notes.entry(record.key).get("last_author", "")))

    def _refresh_notes(self) -> None:
        clear_layout(self.notes_box)
        if self.record is None:
            self.notes_empty.setVisible(True)
            return
        items = self.notes.notes_for(self.record.key)
        self.notes_empty.setVisible(not items)
        for note in items:
            card = Card(self.colors, radius=10, padding=8, variant="flat")
            head = QLabel(note.get("author") or "بدون نام")
            head.setStyleSheet("font-weight: 700; font-size: 8.5pt;")
            card.body.addWidget(head)
            text = QLabel(note.get("text") or "")
            text.setWordWrap(True)
            text.setObjectName("Subtle")
            card.body.addWidget(text)
            if note.get("followup"):
                card.body.addWidget(_muted(f"پیگیری: {note['followup']}"))
            card.body.addWidget(_muted(note.get("created", "")[:10]))
            self.notes_box.addWidget(card)

    def _sync_bookmark(self) -> None:
        if self.record is None:
            return
        flagged = self.notes.is_bookmarked(self.record.key)
        self.bookmark_button.set_icon_kind("bookmark", self.colors["brand"] if flagged
                                           else self.colors["ink_3"])
        self.bookmark_button.setToolTip("برداشتن نشان" if flagged else "نشان‌گذاری")

    def _toggle_bookmark(self) -> None:
        if self.record is None:
            return
        self.notes.toggle_bookmark(self.record.key)
        self._sync_bookmark()
        self.bookmarkToggled.emit(self.record.key)

    def _save_note(self) -> None:
        if self.record is None:
            return
        text = self.note_edit.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "یادداشت", "متن یادداشت خالی است.")
            return
        author = self.author_edit.text().strip() or "کاربر"
        self.notes.add_note(self.record.key, author, text, self.follow_edit.text().strip())
        entry = self.notes.entry(self.record.key)
        entry["last_author"] = author
        self.notes.save()
        self.note_edit.clear()
        self.follow_edit.clear()
        self._refresh_notes()
        self.notesChanged.emit(self.record.key)

    def _open_pdf(self) -> None:
        if self.record is not None and self.record.pdf_file:
            self.pdfRequested.emit(self.record.pdf_file)

    def _reveal_pdf(self) -> None:
        if self.record is not None and self.record.pdf_file:
            self.revealRequested.emit(self.record.pdf_file)

    def set_theme_colors(self, colors: Dict[str, str]) -> None:
        self.colors = colors
        self._apply_shell()
        self.bookmark_button.set_icon_color(colors["ink_3"])
        self.close_button.set_icon_color(colors["ink_3"])
        self.reveal_pdf_button.set_icon_color(colors["ink_2"])
        self.open_pdf_button.setIcon(make_icon("external", colors["ink_2"], 14))
        for card in self.findChildren(Card):
            card.set_theme(colors)
        if self.record is not None:
            self.set_record(self.record)


class ExportDialog(QDialog):
    """انتخاب ستون‌ها و دامنهٔ خروجی اکسل."""

    def __init__(self, column_titles: Sequence[str], visible_count: int,
                 selected_count: int, total_count: int, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("خروجی اکسل")
        self.setMinimumWidth(600)
        self.setMinimumHeight(600)
        self.setLayoutDirection(Qt.RightToLeft)
        colors = THEMES["light"]
        app = QApplication.instance()
        if app is not None:
            colors = THEMES["dark"] if app.property("mehr_theme") == "dark" else THEMES["light"]

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(12)

        header = QHBoxLayout()
        header.setSpacing(12)
        header.addWidget(IconTile("download", colors["brand"], colors["brand_soft"], size=46))
        text_box = QVBoxLayout()
        text_box.setSpacing(1)
        title = QLabel("خروجی سفارشی اکسل")
        title.setObjectName("SectionTitle")
        subtitle = QLabel("دامنهٔ داده و ستون‌های خروجی را انتخاب کنید.")
        subtitle.setObjectName("Muted")
        text_box.addWidget(title)
        text_box.addWidget(subtitle)
        header.addLayout(text_box, 1)
        layout.addLayout(header)

        scope_label = QLabel("دامنهٔ داده")
        scope_label.setObjectName("CardTitle")
        layout.addWidget(scope_label)
        self.scope_tabs = PillTabs([
            ("visible", f"ردیف‌های فیلترشده ({fa(visible_count)})"),
            ("selected", f"انتخاب‌شده ({fa(selected_count)})"),
            ("all", f"همه ({fa(total_count)})"),
        ], colors)
        if not selected_count:
            self.scope_tabs._buttons["selected"].setEnabled(False)
        self.scope_tabs.changed.connect(lambda _key: self._update_summary())
        layout.addWidget(self.scope_tabs)

        columns_label = QLabel("ستون‌ها")
        columns_label.setObjectName("CardTitle")
        layout.addWidget(columns_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        holder = QFrame()
        grid = QGridLayout(holder)
        grid.setContentsMargins(12, 12, 12, 12)
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(8)
        self.checks: List[CheckBox] = []
        for index, name in enumerate(column_titles):
            check = CheckBox(name)
            check.setChecked(True)
            check.set_theme_colors(colors)
            check.stateChanged.connect(lambda _state: self._update_summary())
            self.checks.append(check)
            grid.addWidget(check, index // 2, index % 2)
        for column in range(2):
            grid.setColumnStretch(column, 1)
        scroll.setWidget(holder)
        layout.addWidget(scroll, 1)

        tools = QHBoxLayout()
        tools.setSpacing(6)
        select_all = QPushButton("انتخاب همه")
        select_all.setObjectName("Ghost")
        select_all.clicked.connect(lambda: [c.setChecked(True) for c in self.checks])
        tools.addWidget(select_all)
        select_none = QPushButton("هیچ‌کدام")
        select_none.setObjectName("Ghost")
        select_none.clicked.connect(lambda: [c.setChecked(False) for c in self.checks])
        tools.addWidget(select_none)
        tools.addStretch(1)
        self.summary = QLabel("")
        self.summary.setObjectName("Muted")
        tools.addWidget(self.summary)
        layout.addLayout(tools)

        footer = QHBoxLayout()
        footer.addStretch(1)
        cancel = QPushButton("انصراف")
        cancel.setObjectName("Ghost")
        cancel.clicked.connect(self.reject)
        footer.addWidget(cancel)
        save = QPushButton("ذخیرهٔ فایل")
        save.setObjectName("Primary")
        save.setIcon(make_icon("download", "#FFFFFF", 15))
        save.clicked.connect(self.accept)
        footer.addWidget(save)
        layout.addLayout(footer)
        self._update_summary()

    def _update_summary(self) -> None:
        count = len(self.selected_columns())
        self.summary.setText(bidi(f"{fa(count)} ستون انتخاب شده"))

    def selected_columns(self) -> List[int]:
        return [index for index, check in enumerate(self.checks) if check.isChecked()]

    def scope(self) -> str:
        return self.scope_tabs.current()

    def accept(self) -> None:
        if not self.selected_columns():
            QMessageBox.warning(self, "خروجی اکسل", "حداقل یک ستون انتخاب کنید.")
            return
        super().accept()


class MainWindow(QMainWindow):
    """پنجرهٔ اصلی برنامه با ریل ناوبری، نوار بالا و صفحه‌های پشته‌ای."""

    def __init__(self, config: AppConfig, notes: NotesStore) -> None:
        super().__init__()
        self.config = config
        self.notes = notes
        self.data: DashboardData = empty_dashboard()
        self.loader = None
        self._import_worker = None
        self._current_page = ""
        self._last_error = ""
        self.colors: Dict[str, str] = THEMES["light"]
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(make_app_icon())
        self.setMinimumSize(1120, 720)
        self._restore_geometry()
        self._build_ui()
        self._create_actions()
        theme = str(config.get("theme") or "light")
        self.apply_theme(theme, int(config.get("font_size") or 10))
        self._setup_auto_timer()
        QTimer.singleShot(240, self.initial_load)

    # ------------------------------------------------------------------
    # ساخت رابط
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("Central")
        self.setCentralWidget(central)
        self.central = central
        shell = QHBoxLayout(central)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        self.rail = Rail(self.colors)
        self.rail.pageRequested.connect(self.show_page)
        self.rail.themeRequested.connect(self.toggle_theme)
        shell.addWidget(self.rail)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)
        shell.addWidget(right, 1)

        self.topbar = TopBar(self.colors)
        self.topbar.reloadRequested.connect(self.load_data)
        self.topbar.themeToggled.connect(self.toggle_theme)
        self.topbar.menuRequested.connect(self._show_app_menu)
        self.topbar.searchChanged.connect(self._on_search_text_changed)
        right_layout.addWidget(self.topbar)

        self.stack = QStackedWidget()
        right_layout.addWidget(self.stack, 1)

        self.dashboard = DashboardPage(self.colors)
        self.contracts = ContractsPage(self.notes, self.colors)
        self.reports = ReportsPage(self.colors)
        self.pdfs = PdfLibraryPage(self.colors)
        self.settings = SettingsPage(self.config, self.colors)
        self.help = HelpPage(self.colors)
        for page in (self.dashboard, self.contracts, self.reports,
                     self.pdfs, self.settings, self.help):
            self.stack.addWidget(page)

        self.search_edit = self.topbar.search_edit
        self.status_source = QLabel("منبع: —")
        self.status_counts = QLabel("—")
        self.status_loaded = QLabel("—")
        self.statusBar().addWidget(self.status_source, 1)
        self.statusBar().addPermanentWidget(self.status_counts)
        self.statusBar().addPermanentWidget(self.status_loaded)

        self.detail = DetailDrawer(self.colors, self.notes, central)
        self.detail.closed.connect(self._on_drawer_closed)
        self.detail.notesChanged.connect(self._on_notes_changed)
        self.detail.pdfRequested.connect(self.open_pdf)
        self.detail.revealRequested.connect(self.reveal_pdf)
        self.detail.bookmarkToggled.connect(self._on_bookmark_changed)
        self.detail.setEnabled(False)

        self.toast = Toast(central)

        # ---- اتصال صفحه‌ها ----
        self.dashboard.filterRequested.connect(self._on_filter_requested)
        self.dashboard.recordActivated.connect(self.show_record_detail)
        self.dashboard.navigateRequested.connect(self.show_page)
        self.dashboard.reloadRequested.connect(self.load_data)
        self.dashboard.preview_model = self.contracts.model
        self.contracts.recordActivated.connect(self.show_record_detail)
        self.contracts.bookmarkRequested.connect(
            lambda record: self.toggle_bookmark(record.key)
        )
        self.contracts.exportRequested.connect(self.export_table)
        self.reports.exportExcelRequested.connect(self.export_management)
        self.reports.printRequested.connect(self.print_report)
        self.reports.refreshRequested.connect(self.load_data)
        self.pdfs.openRequested.connect(self.open_pdf)
        self.pdfs.folderRequested.connect(self.pick_pdf_folder)
        self.settings.saveRequested.connect(self.save_settings)
        self.settings.reloadRequested.connect(self.load_data)
        self.settings.openDataFolderRequested.connect(self._open_data_folder)
        self.settings.testRequested.connect(self.test_database)
        self.settings.verifyRequested.connect(self.verify_database_numbers)
        self.settings.detectRequested.connect(self.detect_database_settings)
        self.settings.importRequested.connect(self.import_excel)
        self.settings.systemUpdateRequested.connect(self.start_system_update)
        self.settings.diagnoseRequested.connect(self.show_diagnosis)
        self.settings.sourceWorkbookRequested.connect(self.pick_source_workbook)
        self.settings.workbookPickRequested.connect(self.pick_database_workbook)
        self.settings.localRunRequested.connect(self.open_local_excel)
        self.dashboard.importRequested.connect(self.import_excel_dialog)
        self.dashboard.localRunRequested.connect(self.open_local_excel_dialog)

        self.show_page("dashboard")
        self.detail.place()

    def _create_actions(self) -> None:
        self.shortcuts = []
        for sequence, slot in (
            ("Ctrl+R", self.load_data),
            ("Ctrl+F", self.focus_search),
            ("Ctrl+E", self.export_table),
            ("Ctrl+D", self.toggle_theme),
            ("Ctrl+I", self._toggle_detail_dock),
            ("F1", lambda: self.show_page("help")),
            ("Esc", self._on_escape),
        ):
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.setContext(Qt.ApplicationShortcut)
            shortcut.activated.connect(slot)
            self.shortcuts.append(shortcut)

        menu = self.menuBar()
        menu.setVisible(False)

    def _show_app_menu(self) -> None:
        menu = QMenu(self)
        for key in ("dashboard", "contracts", "reports", "pdfs", "settings", "help"):
            action = menu.addAction(PAGE_TITLES[key])
            action.triggered.connect(lambda _checked=False, k=key: self.show_page(k))
        menu.addSeparator()
        for label, slot in (
            ("بازخوانی از دیتابیس", self.load_data),
            ("تغییر پوسته", self.toggle_theme),
            ("تنظیمات اتصال", lambda: self.show_page("settings")),
            ("دربارهٔ برنامه", self.show_about),
            ("خروج", self.close),
        ):
            action = menu.addAction(label)
            action.triggered.connect(slot)
        menu.exec_(self.topbar.menu_button.mapToGlobal(
            self.topbar.menu_button.rect().bottomLeft()))

    # ------------------------------------------------------------------
    # ناوبری
    # ------------------------------------------------------------------
    def show_page(self, key: str) -> None:
        pages = {
            "dashboard": self.dashboard,
            "contracts": self.contracts,
            "reports": self.reports,
            "pdfs": self.pdfs,
            "settings": self.settings,
            "help": self.help,
        }
        page = pages.get(key)
        if page is None:
            return
        self.rail.set_active(key)
        self.topbar.set_page(key)
        self.stack.setCurrentWidget(page)
        self._current_page = key
        if key == "contracts":
            QTimer.singleShot(120, self.contracts._fit_columns)
        if key == "reports":
            self.reports.update_data(self.data)

    def focus_search(self) -> None:
        self.search_edit.setFocus()
        self.search_edit.selectAll()

    def _on_escape(self) -> None:
        if self.detail.isVisible():
            self.detail.close_drawer()
            return
        if self._current_page == "contracts":
            self.show_page("dashboard")

    def _toggle_detail_dock(self) -> None:
        if self.detail.isVisible():
            self.detail.close_drawer()
        elif self.contracts.table.currentIndex().isValid():
            self.show_record_detail(
                self.contracts.record_at(self.contracts.table.currentIndex().row())
            )

    def _on_drawer_closed(self) -> None:
        try:
            self.contracts.table.setFocus()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # بارگذاری داده
    # ------------------------------------------------------------------
    @guarded("initial_load")
    def initial_load(self) -> None:
        log_event("بارگذاری اولیه آغاز شد")
        self._startup_fallback_used = False
        if not str(self.config.get("db_host") or "").strip():
            # بدون میزبان، دیتابیس محلی روی همین سیستم امتحان می‌شود؛ اگر نبود
            # به‌طور خودکار فایل محلی بارگذاری می‌شود (برنامه بسته نمی‌شود).
            self.config.set("db_host", "127.0.0.1", save=False)
        if str(self.config.get("data_source") or "database") == "local":
            self.load_data(silent=True)
            return
        if previous_attempt_crashed():
            log_event(
                "اجرای قبلی وسط اتصال به دیتابیس بسته شد؛ اتصال خودکار غیرفعال شد"
            )
            self.show_page("settings")
            self.settings.set_status(
                "اجرای قبلی هنگام اتصال به دیتابیس بسته شد — «بارگذاری فایل اکسل» "
                "یا «اجرای محلی (بدون دیتابیس)» را بزنید."
            )
            self.dashboard.show_error(
                "اجرای قبلی هنگام اتصال به دیتابیس بسته شد؛ از «بارگذاری فایل اکسل» "
                "یا «اجرای محلی» استفاده کنید."
            )
            self.topbar.set_connection(False)
            self._notify("اتصال خودکار غیرفعال شد — راهنمای تنظیمات", "alert")
            return
        self.load_data(silent=True)

    @guarded("load_data")
    def load_data(self, *_args, silent: bool = False) -> None:
        if self._loader_running():
            return
        local_source = str(self.config.get("data_source") or "database") == "local"
        if not local_source and not str(self.config.get("db_host") or "").strip():
            self.show_page("settings")
            return
        settings = self.config.db_settings()
        pdf_folder = resolve_pdf_folder(self.config)
        if pdf_folder and not str(self.config.get("pdf_folder") or "").strip():
            self.config.set("pdf_folder", pdf_folder, save=False)
            log_event(f"پوشهٔ PDF خودکار شناسایی شد: {pdf_folder}")
        local_path = ""
        if str(self.config.get("data_source") or "database") == "local":
            candidate = str(self.config.get("local_excel_path") or "")
            if (not candidate or not Path(candidate).exists()) and managed_workbook_path().exists():
                candidate = str(managed_workbook_path())
                self.config.set("local_excel_path", candidate, save=False)
            if candidate and Path(candidate).exists():
                local_path = candidate
                settings["sheet_name"] = str(
                    self.config.get("sheet_name") or SHEET_NAME_DEFAULT
                )
            else:
                self.config.set("data_source", "database", save=False)
        self._set_loading(True)
        loader = DataLoader(settings, pdf_folder, self, local_path=local_path)
        loader.loaded.connect(lambda data, silent=silent: self._on_data_loaded(data, silent))
        loader.failed.connect(lambda message, silent=silent: self._on_load_failed(message, silent))
        loader.finished.connect(self._on_loader_finished)
        self.loader = loader
        loader.start()

    def _loader_running(self) -> bool:
        loader = self.loader
        if loader is None:
            return False
        try:
            return bool(loader.isRunning())
        except RuntimeError:
            self.loader = None
            return False

    @guarded("_set_loading")
    def _set_loading(self, loading: bool) -> None:
        self.topbar.reload_button.setEnabled(not loading)
        if loading:
            self.statusBar().showMessage("در حال خواندن داده از دیتابیس مرکزی…", 0)

    @guarded("_on_loader_finished")
    def _on_loader_finished(self) -> None:
        self._set_loading(False)
        self.topbar.reload_button.setEnabled(True)
        self.loader = None

    def _watch_contract_count(self, data: DashboardData) -> None:
        """اگر تعداد قراردادهای یک منبع ناگهان کم شود، کاربر را خبر می‌کند."""
        if getattr(data, "damaged", False) and not getattr(data, "repaired", 0):
            return
        source = str(data.source_note or "")
        previous = int(self.config.get("last_loaded_count") or 0)
        previous_source = str(self.config.get("last_loaded_source") or "")
        current = int(data.total_contracts or 0)
        self.config.set("last_loaded_count", current, save=False)
        self.config.set("last_loaded_source", source, save=False)
        self.config.save()
        if previous < 100 or current <= 0 or previous_source != source:
            return
        if current >= previous * 0.7:
            return
        message = (
            f"تعداد قراردادها از {fa(previous)} به {fa(current)} کاهش پیدا کرد.\n\n"
            "یعنی فایلی که خوانده شد با بار قبل فرق دارد. اگر خودتان فایل را عوض "
            "نکرده‌اید، نسخهٔ سالم کتاب کار را دوباره بارگذاری کنید."
        )
        log_event(
            f"افت تعداد قراردادها: {previous} → {current} | منبع: {source}"
        )
        self._notify("تعداد قراردادهای خوانده‌شده کم شد", "alert")
        self.statusBar().showMessage(
            f"توجه: {fa(previous)} → {fa(current)} قرارداد", 15000
        )
        box = QMessageBox(self)
        box.setWindowTitle("تعداد قراردادها کم شد")
        box.setIcon(QMessageBox.Warning)
        box.setText("فایل خوانده‌شده با بار قبل تفاوت دارد.")
        box.setInformativeText(message)
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    @guarded("_on_data_loaded")
    def _on_data_loaded(self, data: DashboardData, silent: bool = False) -> None:
        log_event(f"داده در رابط کاربری — {data.total_contracts} قرارداد")
        self._adopt_last_connection()
        if str(self.config.get("data_source") or "database") == "local":
            self.topbar.set_connection(True, "فایل محلی")
            self.topbar.status_pill.setToolTip(
                "داده‌ها از فایل اکسل محلی خوانده می‌شود (بدون دیتابیس)"
            )
        else:
            self.topbar.set_connection(True, "دیتابیس")
            self.topbar.status_pill.setToolTip(
                "داده‌ها از جدول‌های دیتابیس MySQL (mehr_contracts) خوانده می‌شود"
            )
        self.data = data
        choice = dict(LAST_WORKBOOK_CHOICE or {})
        if choice.get("switched") and choice.get("workbook_id"):
            new_id = int(choice["workbook_id"])
            if int(self.config.get("db_workbook_id") or 0) != new_id:
                self.config.set("db_workbook_id", new_id, save=False)
                self.config.save()
            try:
                self.settings.workbook_spin.setValue(new_id)
            except Exception:
                pass
            self.settings.set_status(
                f"کتاب کار کامل‌تر پیدا شد — شناسهٔ {fa(new_id)} استفاده می‌شود"
            )
            self.statusBar().showMessage(
                "شناسهٔ کتاب کار به‌طور خودکار به "
                f"{fa(new_id)} تغییر کرد (کتاب کار کامل‌تر دیتابیس)", 12000,
            )
            log_event(
                f"شناسهٔ کتاب کار در تنظیمات به {new_id} تغییر کرد (انتخاب خودکار)"
            )
        self._watch_contract_count(data)
        # هر بخش جدا به‌روزرسانی می‌شود؛ خطا در یک بخش نباید بقیهٔ داشبورد را
        # صفر/خالی بگذارد (همین باگ باعث می‌شد همهٔ کارت‌ها صفر بمانند).
        for label, action in (
            ("dashboard", lambda: self.dashboard.set_data(data)),
            ("contracts", lambda: self.contracts.set_data(data)),
            ("reports", lambda: self.reports.update_data(data)),
            ("pdfs", lambda: self.pdfs.set_folder(data.pdf_folder)),
            ("settings_values", lambda: self.settings.load_values(self.config)),
            ("settings_loaded",
             lambda: self.settings.set_loaded(data.excel_path, data.total_contracts)),
            ("settings_data", lambda: self.settings.update_data(data)),
        ):
            try:
                action()
            except Exception:
                log_event(f"به‌روزرسانی بخش «{label}» با خطا متوقف شد:\n{traceback.format_exc()}")
        source_text = data.source_note or "دیتابیس مرکزی"
        prefix = "" if str(source_text).startswith("فایل محلی") else "متصل — "
        self.settings.set_status(prefix + source_text)
        self.detail.setEnabled(bool(data.records))
        source = data.source_note or "دیتابیس مرکزی"
        self.status_source.setText(bidi(f"منبع: \u202a{data.excel_path}\u202c • {source}"))
        self.status_counts.setText(bidi(f"{fa(data.total_contracts)} قرارداد"))
        self.status_loaded.setText(f"بارگذاری: {data.loaded_at_text}")
        self.statusBar().showMessage(
            f"داده‌ها با موفقیت بارگذاری شد — {fa(data.total_contracts)} قرارداد"
            + (f" | {len(data.warnings)} هشدار" if data.warnings else ""),
            6000,
        )
        if not silent:
            self._notify(f"{fa(data.total_contracts)} قرارداد بارگذاری شد", "check")
        self._setup_auto_timer()
        system_only = self._looks_like_system_only(data) if data.records else False
        damaged = bool(getattr(data, "damaged", False))
        if damaged:
            self.statusBar().showMessage(
                "کتاب کار ذخیره‌شده ناقص است — اعداد داشبورد معتبر نیست", 10000
            )
            self._notify("کتاب کار ذخیره‌شده ناقص است (فرمول‌ها بدون مقدار)", "alert")
        if not silent and data.warnings and not system_only and not damaged:
            QMessageBox.warning(self, "هشدار داده‌ها", "\n".join(data.warnings))
        if damaged:
            QTimer.singleShot(50, lambda: self._ask_for_damaged_workbook(data))
        elif system_only:
            QTimer.singleShot(50, lambda d=data: self._maybe_restore_real_workbook(d))

    # ------------------------------------------------------------------
    # بازیابی خودکار وقتی دیتابیس کتاب کارِ ساخته‌شده از فایل ۰۱ را دارد
    # ------------------------------------------------------------------
    def _looks_like_system_only(self, data: DashboardData) -> bool:
        """آیا دادهٔ فعلی همان فایل ۰۱ است (بدون وضعیت/درصد)؟"""
        if not data.records:
            return False
        if data.progress_percent_text != "—":
            return False
        if data.source_kind == "generated":
            # کتاب کار ساخته‌شده از فایل ۰۱ هیچ‌وقت سلول درصد (ردیف ۲ ستون ۳۷) را ندارد.
            return True
        # ردیف‌های ذخیره‌شده با نسخهٔ قبل ممکن است نوع منبع را نداشته باشند؛
        # خود فایل ذخیره‌شده بررسی می‌شود تا «ساخته‌شده از ۰۱» بودنش معلوم شود.
        try:
            info = fetch_workbook_blob(self.config.db_settings())
        except Exception:
            return False
        if inspect_workbook_bytes(info.get("data") or b"").get("generated"):
            return True
        if any(data.flag_counts.values()):
            return False
        return False

    def _maybe_restore_real_workbook(self, data: DashboardData) -> None:
        """اگر کتاب کار اصلی روی سیستم پیدا شد، جای نسخهٔ ساخته‌شده می‌نشیند."""
        if getattr(self, "_restore_done", False):
            return
        if not self._looks_like_system_only(data):
            return
        self._restore_done = True
        real = find_real_workbook(
            str(self.config.get("last_excel_path") or ""),
            extra_paths=(
                self.config.get("source_workbook_path"),
                self.config.get("local_excel_path"),
                self.settings.excel_edit.text().strip(),
            ),
        )
        if not real:
            log_event("کتاب کار ساخته‌شده از فایل ۰۱ است و کتاب کار اصلی پیدا نشد")
            if not bool(self.config.get("allow_system_only")):
                self._ask_for_real_workbook()
            return
        log_event(f"کتاب کار اصلی پیدا شد و جای نسخهٔ ساخته‌شده می‌نشیند: {real}")
        settings = self.config.db_settings()
        sheet = str(self.config.get("sheet_name") or SHEET_NAME_DEFAULT)
        local_mode = str(self.config.get("data_source") or "database") == "local"
        restored = False
        blob = b""
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                blob = Path(real).expanduser().read_bytes()
                last_source = str(self.config.get("last_excel_path") or "")
                if last_source and Path(last_source).exists() \
                        and not has_colorcode_sheet(last_source):
                    # آخرین فایل ۰۱ کاربر هم دوباره روی کتاب کار اصلی اعمال می‌شود
                    try:
                        plan = plan_system_update(last_source, blob, sheet_name=sheet)
                        blob, applied = apply_system_update(plan, blob)
                        log_event(
                            f"آخرین فایل ۰۱ روی کتاب کار اصلی اعمال شد: {applied} سلول"
                        )
                    except Exception as exc:
                        log_event(f"اعمال آخرین فایل ۰۱ روی پایه ناموفق بود: {exc}")
                with tempfile.TemporaryDirectory(prefix="mehr-recover-") as temp:
                    temp_path = Path(temp) / Path(real).name
                    if temp_path.suffix.lower() not in {".xlsx", ".xlsm"}:
                        temp_path = temp_path.with_suffix(".xlsx")
                    temp_path.write_bytes(blob)
                    import_workbook_to_database(settings, str(temp_path), sheet_name=sheet)
                restored = True
            finally:
                QApplication.restoreOverrideCursor()
        except Exception as exc:
            log_event(f"جایگزینی کتاب کار اصلی در دیتابیس ناموفق بود: {exc}")
        if (local_mode or not restored) and blob:
            # نسخهٔ ترمیم‌شده (کتاب کار اصلی + آخرین فایل ۰۱) به‌صورت محلی هم ذخیره می‌شود
            try:
                managed_workbook_path().write_bytes(blob)
                self.config.set("data_source", "local", save=False)
                self.config.set("local_excel_path", str(managed_workbook_path()), save=False)
            except Exception as exc:
                log_event(f"ذخیرهٔ کتاب کار ترمیم‌شده ممکن نشد: {exc}")
        self.config.set("source_workbook_path", real, save=False)
        self.config.set("allow_system_only", False, save=False)
        self.config.save()
        self.settings.set_importing(False, real)
        self._notify(
            "کتاب کار اصلی پیدا شد؛ جای کتاب کار ساخته‌شده از فایل ۰۱ نشست", "check"
        )
        self.load_data(silent=True)

    @guarded("_ask_for_damaged_workbook")
    def _ask_for_damaged_workbook(self, data: DashboardData) -> None:
        """کتاب کار ذخیره‌شده فرمول‌هایی بدون مقدار دارد؛ فایل اصلی خواسته می‌شود."""
        if getattr(self, "_damaged_prompted", False):
            return
        self._damaged_prompted = True
        box = QMessageBox(self)
        box.setWindowTitle(DAMAGED_WORKBOOK_TITLE)
        box.setIcon(QMessageBox.Warning)
        box.setText("کتاب کار ذخیره‌شده در دیتابیس ناقص است و اعداد آن معتبر نیست.")
        box.setInformativeText(
            getattr(data, "damaged_reason", "") or damaged_workbook_message({})
        )
        load_btn = box.addButton("بارگذاری کتاب کار اصلی", QMessageBox.AcceptRole)
        box.addButton("بعداً", QMessageBox.RejectRole)
        box.exec_()
        if box.clickedButton() is load_btn:
            self.import_excel_dialog()

    @guarded("_ask_for_real_workbook")
    def _ask_for_real_workbook(self) -> None:
        box = QMessageBox(self)
        box.setWindowTitle("کتاب کار اصلی لازم است")
        box.setIcon(QMessageBox.Warning)
        box.setText(
            "کتاب کار ذخیره‌شده همان خروجی سامانه (۰۱) است و ستون‌های وضعیت و درصد "
            "پیشرفت را ندارد؛ به همین دلیل کارت‌ها صفر شده‌اند."
        )
        box.setInformativeText(
            "فایل ۰۱ فقط ستون‌های کارکرد را دارد. یک بار فایل اصلی خودتان را بدهید "
            "(همان فایلی که شیت colorcode دارد، مثل \u202aF.O.C(Mehr).xlsx\u202c)؛ "
            "از آن به بعد هر فایل ۰۱ روی همان اعمال می‌شود و اعداد دقیقاً مثل قبل "
            "برمی‌گردند."
        )
        pick_btn = box.addButton("انتخاب کتاب کار اصلی", QMessageBox.AcceptRole)
        box.addButton("بعداً", QMessageBox.RejectRole)
        box.exec_()
        if box.clickedButton() is pick_btn:
            last_source = str(self.config.get("last_excel_path") or "")
            if last_source and Path(last_source).exists() \
                    and not has_colorcode_sheet(last_source):
                # بعد از انتخاب پایه، همان فایل ۰۱ دوباره روی آن اعمال می‌شود
                self._pending_system_source = last_source
            self.pick_source_workbook()

    @guarded("_on_load_failed")
    def _on_load_failed(self, message: str, silent: bool = False) -> None:
        self._last_error = message
        log_event("بارگذاری ناموفق: " + " / ".join((message or "").splitlines()[:3]))
        self.topbar.set_connection(False)
        first = message.splitlines()[0] if message else "خطا در بارگذاری"
        self.statusBar().showMessage("بارگذاری داده‌ها ناموفق بود — " + first, 8000)
        self._notify("بارگذاری داده‌ها ناموفق بود", "alert")
        if not self.data.records:
            empty = empty_dashboard(first)
            self.dashboard.set_data(empty)
            self.dashboard.show_error(first)
            self.reports.update_data(empty)
        if silent and not self.data.records and not getattr(
            self, "_startup_fallback_used", False
        ):
            # اجرای اول: دیتابیس در دسترس نبود؛ بی‌سر‌و‌صدا با فایل محلی اجرا می‌شود.
            fallback = ""
            for candidate in (
                self.config.get("source_workbook_path"),
                self.config.get("local_excel_path"),
            ):
                text = str(candidate or "").strip()
                if text and Path(text).exists() and has_colorcode_sheet(text):
                    fallback = text
                    break
            if not fallback:
                fallback = find_seed_workbook(str(self.config.get("last_excel_path") or ""))
            if fallback:
                self._startup_fallback_used = True
                log_event(f"دیتابیس در دسترس نبود؛ اجرای خودکار با فایل محلی: {fallback}")
                self.config.set("data_source", "local", save=False)
                self.config.set("local_excel_path", fallback, save=False)
                self.config.save()
                self.settings.set_importing(False, fallback)
                self.settings.set_status(
                    "دیتابیس در دسترس نبود — داده‌ها از فایل محلی خوانده شد: "
                    + Path(fallback).name
                )
                self._notify("اجرای خودکار با فایل محلی (دیتابیس در دسترس نبود)", "table")
                self.load_data(silent=True)
                return
        if silent:
            if not self.data.records:
                self.settings.set_status(
                    "اتصال برقرار نشد — تنظیمات را بررسی کنید یا فایل اکسل را بارگذاری کنید."
                )
                if not getattr(self, "_auto_settings_shown", False):
                    self._auto_settings_shown = True
                    self.show_page("settings")
                    self._notify("راهنمای اتصال در صفحهٔ تنظیمات است", "info")
            return
        box = QMessageBox(self)
        box.setWindowTitle("خطا در خواندن داده‌ها")
        box.setIcon(QMessageBox.Critical)
        box.setText("خواندن داده از دیتابیس مرکزی ممکن نشد.")
        box.setInformativeText(message + f"\n\nفایل گزارش:\n{log_path()}")
        settings_button = box.addButton("تنظیمات اتصال", QMessageBox.AcceptRole)
        retry_button = box.addButton("تلاش دوباره", QMessageBox.ActionRole)
        import_button = box.addButton("بارگذاری فایل اکسل", QMessageBox.ActionRole)
        local_button = box.addButton("اجرای فایل محلی", QMessageBox.ActionRole)
        log_button = box.addButton("باز کردن فایل گزارش", QMessageBox.HelpRole)
        box.addButton("بستن", QMessageBox.RejectRole)
        box.exec_()
        clicked = box.clickedButton()
        if clicked is settings_button:
            self.show_page("settings")
        elif clicked is retry_button:
            self.load_data()
        elif clicked is import_button:
            self.import_excel_dialog()
        elif clicked is local_button:
            self.open_local_excel_dialog()
        elif clicked is log_button:
            open_path(str(log_path()))

    @guarded("_setup_auto_timer")
    def _setup_auto_timer(self) -> None:
        if not hasattr(self, "auto_timer"):
            self.auto_timer = QTimer(self)
            self.auto_timer.timeout.connect(self._auto_refresh)
        minutes = int(self.config.get("auto_refresh_minutes") or 0)
        self.auto_timer.stop()
        if minutes > 0:
            self.auto_timer.start(minutes * 60 * 1000)

    def _auto_refresh(self) -> None:
        if self._loader_running():
            return
        try:
            updated = fetch_workbook_updated_at(self.config.db_settings())
        except Exception:
            return
        if updated is None:
            return
        if self.data.source_mtime is None or updated > self.data.source_mtime:
            self.statusBar().showMessage("نسخهٔ تازه‌تری از کتاب کار در دیتابیس ثبت شده؛ بازخوانی…", 4000)
            self.load_data(silent=True)

    # ------------------------------------------------------------------
    # اتصال دیتابیس
    # ------------------------------------------------------------------
    def test_database(self) -> None:
        values = self.settings.connection_values()
        settings = dict(self.config.db_settings())
        settings.update(values)
        settings["workbook_id"] = int(values.get("workbook_id") or 1)
        counts = None
        try:
            counts = verify_database_numbers(settings)
        except Exception:
            counts = None
        try:
            info = fetch_workbook_blob(settings)
        except ExcelDataError as exc:
            if counts and int(counts.get("total") or 0) > 0:
                message = (
                    f"اتصال موفق — {fa(int(counts['total']))} ردیف در جدول mehr_contracts"
                )
                self.settings.set_status(message)
                self._notify("اتصال به دیتابیس برقرار است", "check")
                return
            self.settings.set_status("اتصال ناموفق — " + str(exc).splitlines()[0])
            QMessageBox.warning(self, "آزمایش اتصال", str(exc))
            return
        size_kb = max(1, len(info["data"]) // 1024)
        if counts and int(counts.get("total") or 0) > 0:
            message = (
                f"اتصال موفق — {info['filename']} ({fa(size_kb)} کیلوبایت) • "
                f"{fa(int(counts['total']))} ردیف در mehr_contracts"
            )
        else:
            message = (
                f"اتصال موفق — {info['filename']} ({fa(size_kb)} کیلوبایت) • "
                "جدول mehr_contracts خالی است؛ یک بار «بارگذاری در دیتابیس» را اجرا کنید"
            )
        self.settings.set_status(message)
        self._notify("اتصال به دیتابیس برقرار است", "check")

    @guarded("verify_database_numbers")
    def verify_database_numbers(self) -> None:
        """اعداد جدول‌های دیتابیس را با اعداد داشبورد مقایسه می‌کند."""
        values = self.settings.connection_values()
        settings = dict(self.config.db_settings())
        settings.update(values)
        settings["workbook_id"] = int(values.get("workbook_id") or 1)
        try:
            sql_counts = verify_database_numbers(settings)
        except Exception as exc:
            QMessageBox.warning(
                self, "بررسی اعداد دیتابیس",
                "خواندن جدول‌های دیتابیس ممکن نشد.\n"
                "اگر کتاب کار را تازه بارگذاری کرده‌اید، یک بار «بارگذاری در دیتابیس» "
                "را تکرار کنید تا جدول‌های mehr_contracts ساخته شوند.\n\n"
                f"{exc}",
            )
            return
        data = self.data
        labels = [("total", "تعداد قرارداد"), ("deductions", "کسورات اسناد اختتام")]
        labels += [(key, label) for key, label, _column in FLAG_DEFS]
        lines: List[str] = []
        mismatched = 0
        flag_counts = data.flag_counts if data.records else {}
        for key, label in labels:
            db_value = int(sql_counts.get(key, 0) or 0)
            if key == "total":
                app_value = data.total_contracts
            elif key == "deductions":
                app_value = data.contracts_with_deductions
            else:
                app_value = int(flag_counts.get(key, 0) or 0)
            same = db_value == app_value
            if not same:
                mismatched += 1
            lines.append(
                ("✅ " if same else "⚠️ ")
                + f"{label}: دیتابیس {fa(db_value)} / داشبورد {fa(app_value)}"
            )
        box = QMessageBox(self)
        box.setWindowTitle("بررسی اعداد دیتابیس")
        box.setIcon(QMessageBox.Information if mismatched == 0 else QMessageBox.Warning)
        box.setText(
            "اعداد دیتابیس با داشبورد یکسان است."
            if mismatched == 0 else
            f"{fa(mismatched)} مورد اختلاف پیدا شد."
        )
        box.setInformativeText("\n".join(lines))
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()
        log_event("بررسی اعداد دیتابیس: " + " | ".join(lines))

    def detect_database_settings(self) -> None:
        found = read_manager_settings()
        if not found:
            self.settings.set_status(
                "تنظیمات «مدیر دیتابیس مرکزی» پیدا نشد؛ مقادیر را دستی وارد کنید."
            )
            return
        password = read_manager_password()
        self.settings.apply_connection({**found, "password": password})
        if password:
            self.settings.set_status("تنظیمات و رمز از مدیر دیتابیس خوانده شد.")
        else:
            self.settings.set_status(
                "میزبان/دیتابیس/کاربر از مدیر دیتابیس خوانده شد — رمز را وارد کنید."
            )

    # ------------------------------------------------------------------
    # بارگذاری فایل اکسل در دیتابیس (شروع از ابتدا)
    # ------------------------------------------------------------------
    def open_local_excel_dialog(self) -> None:
        """انتخاب فایل اکسل و اجرای برنامه با همان فایل، بدون دیتابیس."""
        start = str(
            self.config.get("local_excel_path")
            or self.settings.excel_edit.text().strip()
            or self.config.get("last_excel_path")
            or app_dir()
        )
        path, _selected = QFileDialog.getOpenFileName(
            self, "انتخاب فایل اکسل برای اجرای محلی", start,
            "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
        )
        if path:
            self.open_local_excel(path)

    @guarded("open_local_excel")
    def open_local_excel(self, path: str) -> None:
        """خواندن مستقیم فایل اکسل از دیسک (بدون دیتابیس)."""
        candidate = Path(str(path or "")).expanduser()
        if candidate.exists() and not has_colorcode_sheet(str(candidate)):
            # کاربر فایل ۰۱ سامانه را داده است؛ خود برنامه کتاب کار داخلی را می‌سازد.
            self._system_update_note = (
                f"«{candidate.name}» شیت colorcode ندارد، پس به‌عنوان فایل ۰۱ سامانه "
                "خوانده می‌شود و کتاب کار داخلی برنامه از آن ساخته/به‌روزرسانی می‌شود."
            )
            self.start_local_system_update(str(candidate))
            return
        if not candidate.exists():
            box = QMessageBox(self)
            box.setWindowTitle("فایل اکسل")
            box.setIcon(QMessageBox.Warning)
            box.setText("فایل اکسل پیدا نشد.")
            box.setInformativeText(str(candidate))
            box.addButton("باشه", QMessageBox.AcceptRole)
            box.exec_()
            return
        self.config.set("data_source", "local", save=False)
        self.config.set("local_excel_path", str(candidate), save=False)
        self.config.save()
        self.settings.set_importing(False, str(candidate))
        self.settings.set_status(f"حالت محلی — {candidate.name}")
        self.statusBar().showMessage(
            "اجرای محلی: داده‌ها از فایل اکسل خوانده می‌شود (بدون دیتابیس).", 6000
        )
        self.load_data()

    @guarded("pick_database_workbook")
    def pick_database_workbook(self) -> None:
        """فهرست کتاب کارهای دیتابیس را نشان می‌دهد و شناسه انتخاب‌شده را ذخیره می‌کند."""
        values = self.settings.connection_values()
        settings = {
            "host": values["host"], "port": values["port"],
            "database": values["database"], "user": values["user"],
            "password": values["password"], "workbook_id": values["workbook_id"],
        }
        try:
            items = list_database_workbooks(settings)
        except Exception as exc:
            self._notify("خواندن فهرست کتاب کارها ممکن نشد", "alert")
            QMessageBox.warning(
                self, "کتاب کارهای دیتابیس",
                "اتصال به دیتابیس برقرار نشد یا جدول‌ها ساخته نشده‌اند:\n\n"
                + str(exc)[:600],
            )
            return
        if not items:
            QMessageBox.information(
                self, "کتاب کارهای دیتابیس",
                "هیچ کتاب کاری در دیتابیس ذخیره نشده است.",
            )
            return
        labels = []
        for item in items:
            labels.append(
                "شناسهٔ {wid} — {name} — {rows} ردیف{when}".format(
                    wid=fa(item["workbook_id"]), name=item["filename"],
                    rows=fa(item["rows"]),
                    when=(f" — {item['imported_at']}" if item["imported_at"] else ""),
                )
            )
        current = int(self.settings.workbook_spin.value() or 1)
        index = next(
            (i for i, item in enumerate(items) if item["workbook_id"] == current), 0
        )
        choice, ok = QInputDialog.getItem(
            self, "کتاب کارهای دیتابیس",
            "کتاب کاری که اعداد داشبورد و پایهٔ به‌روزرسانی از آن خوانده می‌شود:\n"
            "(کتاب کار ۰۲ خودتان را با شناسه‌اش انتخاب کنید)",
            labels, index, False,
        )
        if not ok or not choice:
            return
        picked = items[labels.index(choice)]
        self.settings.workbook_spin.setValue(int(picked["workbook_id"]))
        self.config.set("db_workbook_id", int(picked["workbook_id"]), save=False)
        self.config.save()
        self.settings.set_status(
            f"کتاب کار انتخاب شد — شناسهٔ {fa(picked['workbook_id'])} "
            f"«{picked['filename']}»"
        )
        self._notify("کتاب کار انتخاب شد؛ داده‌ها بازخوانی می‌شود", "database")
        self.load_data()

    def import_excel_dialog(self) -> None:
        """انتخاب فایل اکسل و بارگذاری آن در دیتابیس."""
        self.show_page("settings")
        start = str(
            self.settings.excel_edit.text().strip()
            or self.config.get("last_excel_path")
            or app_dir()
        )
        path, _selected = QFileDialog.getOpenFileName(
            self, "انتخاب فایل اکسل", start,
            "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
        )
        if path:
            self.import_excel(path)

    # ------------------------------------------------------------------
    # به‌روزرسانی از فایل سامانه (فایل ۰۱)
    # ------------------------------------------------------------------
    def start_system_update(self, path: str) -> None:
        """فایل ۰۱ را می‌خواند و یک «نسخهٔ به‌روز» از کتاب کار می‌سازد (سالم).

        کتاب کار اصلی هرگز بازنویسی نمی‌شود: فقط مقادیر L..AJ در شیتی که خروجی
        سامانه در آن پیست می‌شود (مثل «خروجی کندو») در یک فایل تازه نوشته می‌شود
        و فرمول‌ها دست‌نخورده می‌مانند.
        """
        if self._system_update_running():
            self._notify("یک به‌روزرسانی در حال‌اجراست؛ کمی صبر کنید", "clock")
            return
        if not str(path or "").strip():
            self.pick_system_update_file()
            return
        if not str(self.config.get("db_host") or "").strip():
            self.start_local_system_update(path)
            return
        self.show_page("settings")
        self._pending_system_source = path
        self.settings.set_status("در حال خواندن فایل سامانه و ساخت نسخهٔ به‌روز…")
        settings = dict(self.config.db_settings())
        settings.update({
            "source_workbook_path": self.config.get("source_workbook_path"),
            "local_excel_path": self.config.get("local_excel_path"),
            "last_excel_path": self.config.get("last_excel_path"),
            "extra_excel_path": self.settings.excel_edit.text().strip(),
        })
        self._start_system_copy(settings, path)

    def _start_system_copy(self, settings: Dict[str, Any], source_path: str,
                           forced_sheet: str = "") -> None:
        """اجرای کارگر ساخت نسخهٔ به‌روز (با امکان انتخاب دستی شیت مقصد)."""
        if not forced_sheet:
            forced_sheet = str(self.config.get("system_update_sheet") or "")
        worker = SystemCopyWorker(settings, source_path, self, forced_sheet=forced_sheet)
        worker.done.connect(self._on_system_copy_done)
        worker.failed.connect(self._on_system_update_failed)
        worker.sheet_needed.connect(self._on_system_sheet_needed)
        worker.finished.connect(lambda: self._set_system_update_busy(False))
        self._system_copy_request = (dict(settings), str(source_path))
        self._system_update_worker = worker
        self._system_update_busy = True
        worker.start()

    @guarded("_on_system_sheet_needed")
    def _on_system_sheet_needed(self, info: Dict[str, Any]) -> None:
        """شیت مقصد را از کاربر می‌پرسد؛ چون تشخیص خودکار مطمئن نبود."""
        sheets = [str(name) for name in (info.get("sheets") or [])]
        if not sheets:
            self._on_system_update_failed(
                "در کتاب کار هیچ شیتی پیدا نشد که بتوان خروجی سامانه را در آن نوشت."
            )
            return
        details = info.get("report") or []
        guess = ""
        for item in details:
            if item.get("hits"):
                guess = str(item.get("sheet") or "")
                break
        if not guess:
            guess = sheets[0]
        hints = []
        for item in details:
            hints.append(
                "{sheet} — هم‌خوانی {hits}".format(
                    sheet=item.get("sheet"), hits=fa(int(item.get("hits") or 0))
                )
            )
        self.settings.set_status("شیت مقصد پیدا نشد؛ انتخاب دستی لازم است")
        message = (
            "در کتاب کار «{name}» شیتی پیدا نشد که با خروجی سامانه هم‌خوانی روشن داشته باشد.\n\n"
            "شیت مقصد را انتخاب کنید (همان شیتی که ستون‌های L تا AJ فایل ۰۱ در آن "
            "پیست می‌شود):\n\n{hints}"
        ).format(name=info.get("base_filename") or "کتاب کار",
                 hints="\n".join(hints[:12]))
        choice, ok = QInputDialog.getItem(
            self, "انتخاب شیت مقصد", message, sheets,
            max(0, sheets.index(guess)) if guess in sheets else 0, False,
        )
        if not ok or not choice:
            self._notify("به‌روزرسانی لغو شد", "info")
            self.settings.set_status("به‌روزرسانی لغو شد — شیتی انتخاب نشد")
            return
        settings, source_path = getattr(self, "_system_copy_request", ({}, ""))
        if not settings:
            self._on_system_update_failed("درخواست به‌روزرسانی گم شد؛ دوباره تلاش کنید.")
            return
        self.config.set("system_update_sheet", str(choice), save=False)
        self.settings.set_status(f"ساخت نسخهٔ به‌روز روی شیت «{choice}»…")
        self._start_system_copy(settings, source_path, forced_sheet=str(choice))

    # ------------------------------------------------------------------
    # به‌روزرسانی محلی از فایل سامانه (بدون دیتابیس)
    # ------------------------------------------------------------------
    def start_local_system_update(self, source_path: str) -> None:
        """فقط فایل ۰۱: کتاب کار داخلی برنامه خودش ساخته یا به‌روزرسانی می‌شود.

        دیگر فایل ۰۲ (کتاب کار دارای شیت colorcode) از کاربر پرسیده نمی‌شود؛
        برنامه همان کار کپی/چسباندن دستی را روی کتاب کار داخلی خودش انجام می‌دهد.
        """
        if self._system_update_running():
            self._notify("یک به‌روزرسانی در حال‌اجراست؛ کمی صبر کنید", "clock")
            return
        if not str(source_path or "").strip():
            self.pick_system_update_file()
            return
        self.show_page("settings")
        self._pending_system_source = source_path
        base_paths = (
            self.config.get("source_workbook_path"),
            self.config.get("local_excel_path"),
            self.config.get("last_excel_path"),
            self.settings.excel_edit.text().strip(),
        )
        real = find_real_workbook(source_path, extra_paths=base_paths)
        if real:
            # کتاب کار اصلی موجود است؛ برنامه آن را بازنویسی نمی‌کند و فقط یک
            # نسخهٔ به‌روز از روی آن می‌سازد (فرمول‌ها دست‌نخورده می‌مانند).
            settings = dict(self.config.db_settings())
            settings.update({
                "source_workbook_path": real,
                "local_excel_path": real,
                "last_excel_path": str(self.config.get("last_excel_path") or ""),
                "extra_excel_path": self.settings.excel_edit.text().strip(),
            })
            self.settings.set_status("در حال ساخت نسخهٔ به‌روز از کتاب کار اصلی…")
            self._start_system_copy(settings, source_path)
            return
        seed = find_seed_workbook(source_path, extra_paths=base_paths)
        self.settings.set_status("در حال ساخت/به‌روزرسانی کتاب کار داخلی از فایل ۰۱…")
        worker = ManagedUpdateWorker(
            source_path, str(self.config.get("sheet_name") or SHEET_NAME_DEFAULT), self,
            seed_path=seed,
        )
        worker.done.connect(self._on_local_update_done)
        worker.failed.connect(self._on_managed_update_failed)
        worker.finished.connect(lambda: self._set_system_update_busy(False))
        self._system_update_worker = worker
        self._system_update_busy = True
        worker.start()

    @guarded("_on_managed_update_failed")
    def _on_managed_update_failed(self, message: str) -> None:
        self.settings.set_status("ساخت کتاب کار داخلی از فایل ۰۱ انجام نشد.")
        self._set_system_update_busy(False)
        self._pending_system_source = ""
        log_event("به‌روزرسانی کتاب کار داخلی ناموفق: "
                  + " / ".join((message or "").splitlines()[:3]))
        box = QMessageBox(self)
        box.setWindowTitle("به‌روزرسانی از فایل ۰۱")
        box.setIcon(QMessageBox.Warning)
        box.setText("خواندن فایل ۰۱ و به‌روزرسانی کتاب کار داخلی ممکن نشد.")
        box.setInformativeText(
            f"{message}\n\nفایل گزارش:\n{log_path()}\n\n"
            "اگر آمار صفر است، دکمهٔ «تشخیص فایل ۰۱» در تنظیمات ساختار فایل را نشان می‌دهد."
        )
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    def _on_local_update_done(self, result: Dict[str, Any]) -> None:
        plan = result.get("plan")
        created = bool(result.get("created"))
        output = str(result.get("output_path") or "")
        source = str(result.get("source_path") or getattr(plan, "source_path", "") or "")
        no_base = bool(result.get("no_source_workbook"))
        if source:
            self.config.set("last_excel_path", source, save=False)
        if output and not no_base:
            self.config.set("local_excel_path", output, save=False)
            self.config.set("data_source", "local", save=False)
            self.config.save()
        self._pending_system_source = ""
        if no_base and created and not (
            getattr(self, "_allow_system_only", False)
            or bool(self.config.get("allow_system_only"))
        ):
            # فایل ۰۱ به‌تنهایی ستون‌های وضعیت/درصد را ندارد؛ اجازه نمی‌دهیم
            # داشبورد با اعداد صفر بالا بیاید و از کاربر کتاب کار اصلی را می‌خواهیم.
            self._set_system_update_busy(False)
            self.settings.set_status(
                "کتاب کار اصلی پیدا نشد — فایل ۰۱ به‌تنهایی وضعیت/درصد ندارد"
            )
            log_event("فایل ۰۱ بدون کتاب کار اصلی اعمال شد؛ منتظر انتخاب کتاب کار اصلی")
            self._pending_system_source = source
            self._ask_for_real_workbook()
            return
        self._allow_system_only = False
        self.settings.set_status(
            ("کتاب کار داخلی ساخته شد" if created else "کتاب کار داخلی به‌روزرسانی شد")
            + f" — {fa(result.get('updated_cells') or 0)} مقدار، "
            + f"{fa(result.get('matched_rows') or 0)} قرارداد"
        )
        self._notify(
            "کتاب کار داخلی ساخته شد" if created else "کتاب کار داخلی به‌روزرسانی شد",
            "check",
        )
        if output:
            self.load_data()
        lines: List[str] = []
        seeded_from = str(result.get("seeded_from") or "")
        if result.get("no_source_workbook"):
            lines.append(
                "کتاب کار اصلی شما (همان فایلی که شیت colorcode دارد و اعداد داشبورد از "
                "آن می‌آمد) روی این سیستم پیدا نشد؛ بنابراین کتاب کار از خود فایل ۰۱ ساخته "
                "شد و ستون‌های وضعیت/پیشرفت در آن نیست."
            )
            lines.append(
                "برای اینکه اعداد دقیقاً مثل قبل شود: دکمهٔ «کتاب کار اصلی (۰۲)» را بزنید و "
                "یک بار همان فایل را بدهید (مثلاً \u202aF.O.C(Mehr).xlsx\u202c یا فایل ۰۲). "
                "از آن به بعد به‌روزرسانی‌ها روی همان انجام می‌شود."
            )
        if created:
            lines.append(
                "برای همین فایل ۰۱، کتاب کار داخلی برنامه ساخته شد: فیلتر «اجرا - مهر» "
                "و انتقال ستون‌های L تا AJ انجام شد — بدون نیاز به فایل ۰۲."
            )
            if seeded_from:
                lines.append(
                    "کتاب کار قبلی شما («" + Path(seeded_from).name + "») پایه گرفته شد تا "
                    "وضعیت‌های ثبت‌شده و ستون‌های دیگرش حفظ شوند."
                )
        else:
            lines.append(
                f"{fa(result.get('matched_rows') or 0)} قرارداد «اجرا - مهر» روی کتاب کار "
                f"داخلی نوشته شد ({fa(result.get('updated_cells') or 0)} مقدار)."
            )
            if plan is not None:
                lines.append(f"روش تطبیق: {plan.method}")
        summary = str(result.get("summary") or "")
        if summary:
            lines.append("")
            lines.append(summary)
        marks = result.get("status_marks")
        if created and marks is not None:
            if marks:
                lines.append("")
                lines.append(
                    f"علامت‌های وضعیت (M/✓/…) پیدا و منتقل شد: {fa(int(marks))} — "
                    "کارت‌ها و نمودارهای وضعیت پر شدند."
                )
            else:
                lines.append("")
                lines.append(
                    "در ستون‌های وضعیت فایل ۰۱ هیچ علامت M/✓ پیدا نشد؛ کارت‌های وضعیت "
                    "خالی می‌مانند. اگر وضعیت‌ها در فایل ۰۱ هستند (ستون‌های دیگر) بگویید "
                    "تا همان ستون‌ها خوانده شود؛ اگر وضعیت‌ها را دستی در شیت colorcode "
                    "می‌زنید، یک بار همان فایل ذخیره‌شده را بدهید تا برنامه پایه بگیرد."
                )
        if output:
            lines.append("")
            lines.append(f"کتاب کار داخلی: {output}")
            lines.append("دفعهٔ بعد فقط فایل ۰۱ سامانه را بدهید؛ همین فایل به‌روز می‌شود.")
        if result.get("backup_path"):
            lines.append(f"نسخهٔ پشتیبان: {result['backup_path']}")
        box = QMessageBox(self)
        box.setWindowTitle("کتاب کار داخلی ساخته شد" if created else "کتاب کار داخلی به‌روزرسانی شد")
        box.setIcon(QMessageBox.Information)
        box.setText(
            "داده‌ها آماده است؛ برنامه همین حالا با کتاب کار داخلی اجرا شد."
            if created else "کتاب کار داخلی برنامه با فایل ۰۱ به‌روزرسانی شد."
        )
        box.setInformativeText("\n".join(lines))
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    @guarded("pick_source_workbook")
    def pick_source_workbook(self) -> None:
        """کتاب کار اصلی کاربر (فایل ۰۲ / colorcode) را ثبت می‌کند."""
        start = str(
            self.config.get("source_workbook_path")
            or self.config.get("local_excel_path")
            or self.config.get("last_excel_path")
            or app_dir()
        )
        path, _selected = QFileDialog.getOpenFileName(
            self, "انتخاب کتاب کار اصلی (دارای شیت colorcode)", start,
            "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
        )
        if not path:
            return
        if not has_colorcode_sheet(path):
            QMessageBox.warning(
                self, "کتاب کار اصلی",
                "این فایل شیت «colorcode» ندارد.\n"
                "فایل اصلی خودتان (مثل \u202aF.O.C(Mehr).xlsx\u202c) یا فایل ۰۲ را انتخاب کنید.",
            )
            return
        if is_generated_workbook(path):
            QMessageBox.warning(
                self, "کتاب کار اصلی",
                "این فایل همان کتاب کار ساخته‌شده توسط برنامه از فایل ۰۱ است و "
                "ستون‌های وضعیت/درصد را ندارد.\n"
                "فایل اصلی خودتان (مثل \u202aF.O.C(Mehr).xlsx\u202c) را انتخاب کنید.",
            )
            return
        self.config.set("source_workbook_path", path, save=False)
        self.config.set("local_excel_path", path, save=False)
        try:
            data = read_local_workbook(
                path, sheet_name=str(self.config.get("sheet_name") or SHEET_NAME_DEFAULT)
            )
            summary = (
                f"{fa(data.total_contracts)} قرارداد — پیشرفت "
                f"{data.progress_percent_text}٪"
            )
        except Exception as exc:
            summary = f"خواندن فایل ممکن نشد: {exc}"
        db_mode = bool(str(self.config.get("db_host") or "").strip())
        stored_in_db = False
        extra = ""
        if db_mode:
            try:
                QApplication.setOverrideCursor(Qt.WaitCursor)
                try:
                    imported = import_workbook_to_database(
                        self.config.db_settings(), path,
                        sheet_name=str(self.config.get("sheet_name") or SHEET_NAME_DEFAULT),
                    )
                finally:
                    QApplication.restoreOverrideCursor()
                stored_in_db = True
                extra = (
                    f"\nدر دیتابیس ذخیره شد: {fa(int(imported.get('data_rows') or 0))} "
                    "ردیف در mehr_contracts"
                )
                log_event(f"کتاب کار اصلی در دیتابیس ذخیره شد: {path}")
            except Exception as exc:
                log_event(f"ذخیرهٔ کتاب کار اصلی در دیتابیس ناموفق بود: {exc}")
                extra = f"\nذخیره در دیتابیس ممکن نشد ({str(exc).splitlines()[0]})."
        if stored_in_db:
            self.config.set("data_source", "database", save=False)
        else:
            # اگر ذخیره در دیتابیس ممکن نشد، همان فایل به‌صورت محلی اجرا می‌شود؛
            # داشبورد نباید به‌خاطر در دسترس نبودن دیتابیس خالی بماند.
            self.config.set("data_source", "local", save=False)
        self.config.save()
        self._restore_done = True
        self.config.set("allow_system_only", False)
        self._notify("کتاب کار اصلی ثبت شد", "check")
        self.settings.set_status(f"کتاب کار اصلی: {Path(path).name}")
        box = QMessageBox(self)
        box.setWindowTitle("کتاب کار اصلی ثبت شد")
        box.setIcon(QMessageBox.Information)
        box.setText("از این به بعد به‌روزرسانی فایل ۰۱ روی همین کتاب کار انجام می‌شود.")
        box.setInformativeText(f"{Path(path).name}\n{summary}{extra}")
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()
        pending = str(getattr(self, "_pending_system_source", "") or "")
        self._pending_system_source = ""
        self.load_data()
        if pending:
            QTimer.singleShot(150, lambda p=pending: self.start_system_update(p))

    @guarded("show_diagnosis")
    def show_diagnosis(self, path: str) -> None:
        """گزارش ساختار فایل ۰۱ را نشان می‌دهد (برای عیب‌یابی آمار صفر)."""
        if not path or not Path(path).exists():
            start = str(self.config.get("last_excel_path") or app_dir())
            path, _selected = QFileDialog.getOpenFileName(
                self, "انتخاب فایل ۰۱ برای تشخیص", start,
                "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
            )
            if not path:
                return
        report = diagnose_system_file(
            path, str(self.config.get("sheet_name") or SHEET_NAME_DEFAULT)
        )
        report_path = data_dir() / "گزارش-تشخیص.txt"
        try:
            report_path.write_text(report, encoding="utf-8")
        except Exception:
            report_path = data_dir()
        box = QMessageBox(self)
        box.setWindowTitle("تشخیص فایل ۰۱")
        box.setIcon(QMessageBox.Information)
        box.setText("نتیجهٔ بررسی ساختار فایل:")
        box.setInformativeText(report[:1400])
        copy_button = box.addButton("کپی گزارش", QMessageBox.ActionRole)
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()
        if box.clickedButton() is copy_button:
            QApplication.clipboard().setText(report)
            self._notify("گزارش در کلیپ‌بورد کپی شد", "check")
        self.statusBar().showMessage(f"گزارش تشخیص ذخیره شد: {report_path}", 8000)

    def pick_system_update_file(self) -> None:
        start = str(
            self.config.get("last_excel_path")
            or self.settings.excel_edit.text().strip()
            or app_dir()
        )
        path, _selected = QFileDialog.getOpenFileName(
            self, "انتخاب فایل سامانه (فایل ۰۱)", start,
            "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
        )
        if path:
            self.start_system_update(path)

    def _system_update_running(self) -> bool:
        worker = getattr(self, "_system_update_worker", None)
        if worker is None:
            return False
        try:
            return bool(worker.isRunning())
        except RuntimeError:
            self._system_update_worker = None
            return False

    def _set_system_update_busy(self, busy: bool) -> None:
        self._system_update_busy = busy
        try:
            self.settings.update_button.setEnabled(not busy)
        except Exception:
            pass

    @guarded("_on_system_plan_ready")
    def _on_system_plan_ready(self, plan: SystemUpdatePlan, blob: bytes,
                              filename: str) -> None:
        self._pending_system_update = {"plan": plan, "blob": blob, "filename": filename}
        box = QMessageBox(self)
        box.setWindowTitle("تأیید به‌روزرسانی از فایل سامانه")
        box.setIcon(QMessageBox.Question)
        box.setText(
            "دیتابیس با همین فایل ۰۱ به‌روزرسانی می‌شود — همان کاری که دستی انجام "
            "می‌دادید، خودکار: فیلتر ستون Sub Project («اجرا - مهر»)، سپس نوشتن "
            "ستون‌های L تا AJ در شیت colorcode و بارگذاری دوبارهٔ همان شیت در دیتابیس."
        )
        note = str(getattr(self, "_system_update_note", "") or "")
        self._system_update_note = ""
        preview = preview_system_update(plan)
        box.setInformativeText((note + "\n\n" + preview) if note else preview)
        apply_btn = box.addButton("اعمال و به‌روزرسانی", QMessageBox.AcceptRole)
        box.addButton("لغو", QMessageBox.RejectRole)
        box.exec_()
        if box.clickedButton() is not apply_btn:
            self.settings.set_status("به‌روزرسانی لغو شد.")
            self._set_system_update_busy(False)
            return
        self.settings.set_status("در حال نوشتن مقادیر روی کتاب کار و ذخیره در دیتابیس…")
        worker = SystemUpdateApplyWorker(self.config.db_settings(), plan, blob, filename, self)
        worker.done.connect(self._on_system_update_done)
        worker.failed.connect(self._on_system_update_failed)
        worker.finished.connect(lambda: self._set_system_update_busy(False))
        self._system_update_worker = worker
        self._system_update_busy = True
        worker.start()

    @guarded("_on_system_copy_done")
    def _on_system_copy_done(self, result: Dict[str, Any]) -> None:
        plan: SystemUpdatePlan = result["plan"]
        self._pending_system_source = ""
        output = str(result.get("output_path") or "")
        imported = result.get("imported") or {}
        if imported and not imported.get("error"):
            self._on_auto_update_done(result, imported)
            return
        if result.get("recalculated") and result.get("healthy") and output:
            self._on_local_copy_ready(result, imported)
            return
        self.settings.set_status(
            f"نسخهٔ به‌روز ساخته شد — {fa(int(result.get('written_cells') or 0))} سلول"
        )
        self._notify("نسخهٔ به‌روز کتاب کار ساخته شد", "check")
        written = int(result.get("written_cells") or 0)
        lines = [
            f"فایل سامانه: {Path(plan.source_path).name}",
            f"ردیف‌های «اجرا - مهر»: {fa(plan.filtered_rows)} — "
            f"تطبیق‌شده: {fa(plan.matched_rows)} ({plan.method})",
            f"سلول‌های نوشته‌شده: {fa(written)}",
            f"شیت مقصد: «{result.get('sheet') or plan.sheet_name}» "
            "(همان شیتی که خروجی سامانه در آن پیست می‌شود)",
        ]
        if not written:
            lines += [
                "",
                "هیچ سلولی تغییر نکرد؛ یعنی مقادیر فایل سامانه با کتاب کار یکی "
                "بوده‌اند. برای همین نیازی به باز کردن فایل در اکسل و بارگذاری "
                "دوباره نیست — داشبورد همین اعداد را دارد.",
            ]
        if result.get("formula_cells"):
            lines.append(
                f"سلول‌های فرمولی دست‌نخورده: {fa(int(result['formula_cells']))}"
            )
        if result.get("cleaned_marks"):
            lines.append(
                "ترمیم وضعیت‌ها: "
                f"{fa(int(result['cleaned_marks']))} سلول که عدد ۰/۱ داشتند به "
                "علامت M/خالی تبدیل شد"
            )
        lines += [
            "",
            "کتاب کار اصلی شما بازنویسی نشد؛ این یک فایل تازه است.",
            "برای اینکه اعداد داشبورد هم به‌روز شوند: این فایل را در اکسل باز کنید "
            "(فرمول‌ها خودکار محاسبه می‌شوند)، Ctrl+S بزنید، و بعد از برنامه "
            "«بارگذاری فایل اکسل» را بزنید و همین فایل را انتخاب کنید.",
        ]
        if output:
            lines += ["", f"فایل تازه: {output}"]
        box = QMessageBox(self)
        box.setWindowTitle("نسخهٔ به‌روز کتاب کار ساخته شد")
        box.setIcon(QMessageBox.Information)
        box.setText("فایل ۰۱ روی یک نسخهٔ تازه از کتاب کار اعمال شد.")
        box.setInformativeText("\n".join(lines))
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()
        if output and os.environ.get("QT_QPA_PLATFORM", "") not in {"offscreen", "minimal"}:
            try:
                QDesktopServices.openUrl(QUrl.fromLocalFile(output))
            except Exception:
                pass

    def _on_local_copy_ready(self, result: Dict[str, Any],
                             imported: Dict[str, Any]) -> None:
        """نسخهٔ به‌روز با اکسل محاسبه شد؛ بدون دیتابیس هم داشبورد به‌روز می‌شود."""
        plan: SystemUpdatePlan = result["plan"]
        output = str(result.get("output_path") or "")
        self.config.set("data_source", "local", save=False)
        self.config.set("local_excel_path", output, save=False)
        self.config.set("last_excel_path", output, save=False)
        self.config.save()
        self.settings.excel_edit.setText(output)
        self.settings.set_status(
            f"به‌روزرسانی انجام شد — {fa(int(result.get('written_cells') or 0))} سلول"
        )
        self._notify("کتاب کار به‌روز شد؛ داشبورد بازخوانی می‌شود", "check")
        self.load_data()
        lines = [
            f"فایل سامانه: {Path(plan.source_path).name}",
            f"ردیف‌های «اجرا - مهر»: {fa(plan.filtered_rows)} — "
            f"تطبیق‌شده: {fa(plan.matched_rows)} ({plan.method})",
            f"سلول‌های نوشته‌شده: {fa(int(result.get('written_cells') or 0))}",
            "",
            "فایل با خودِ اکسل بازمحاسبه و ذخیره شد و داشبورد از همان نسخه "
            "خوانده می‌شود؛ کتاب کار اصلی شما دست‌نخورده است.",
        ]
        if imported.get("error"):
            lines.append("")
            lines.append(
                "دیتابیس در دسترس نبود، برای همین نسخهٔ محلی استفاده شد:\n"
                + str(imported["error"]).splitlines()[0][:200]
            )
        box = QMessageBox(self)
        box.setWindowTitle("به‌روزرسانی انجام شد")
        box.setIcon(QMessageBox.Information)
        box.setText("فایل ۰۱ روی کتاب کار اعمال شد.")
        box.setInformativeText("\n".join(lines))
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    def _on_auto_update_done(self, result: Dict[str, Any],
                             imported: Dict[str, Any]) -> None:
        """دیتابیس خودکار به‌روز شد؛ همین‌جا داشبورد هم بازخوانی می‌شود."""
        plan: SystemUpdatePlan = result["plan"]
        self.config.set("data_source", "database", save=False)
        self.config.set(
            "last_excel_path", str(result.get("output_path") or ""), save=False
        )
        self.config.save()
        rows = int(imported.get("data_rows") or 0)
        cells = int(result.get("written_cells") or 0)
        new_id = int(imported.get("workbook_id") or 0)
        base_id = int(result.get("base_workbook_id") or 0)
        if new_id:
            self.config.set("db_workbook_id", new_id, save=False)
            self.config.save()
            self.settings.workbook_spin.setValue(new_id)
        self.settings.set_status(
            f"به‌روزرسانی انجام شد — {fa(rows)} ردیف در دیتابیس، {fa(cells)} سلول"
            + (f" (شناسهٔ {fa(new_id)})" if new_id else "")
        )
        self._notify("دیتابیس و داشبورد به‌روز شدند", "check")
        self.statusBar().showMessage(
            f"به‌روزرسانی خودکار انجام شد — {fa(plan.matched_rows)} ردیف تطبیق، "
            f"{fa(cells)} سلول", 12000,
        )
        self.load_data()
        lines = [
            f"فایل سامانه: {Path(plan.source_path).name}",
            f"ردیف‌های «اجرا - مهر»: {fa(plan.filtered_rows)} — "
            f"تطبیق‌شده: {fa(plan.matched_rows)} ({plan.method})",
            f"سلول‌های نوشته‌شده: {fa(cells)}",
        ]
        if base_id:
            lines.append(
                f"کتاب کار پایه: شناسهٔ {fa(base_id)} «{result.get('base_filename') or '—'}»"
            )
        lines.append(f"ردیف‌های ذخیره‌شده در دیتابیس: {fa(rows)}")
        if new_id:
            lines.append(
                f"نسخهٔ به‌روز با شناسهٔ {fa(new_id)} در دیتابیس ذخیره شد"
                + (f"؛ کتاب کار پایه (شناسهٔ {fa(base_id)}) دست‌نخورده ماند"
                   if base_id else "")
            )
        lines += [
            "",
            "فایل با خودِ اکسل بازمحاسبه و ذخیره شد، بعد در دیتابیس بارگذاری شد؛ "
            "داشبورد در حال بازخوانی است.",
        ]
        box = QMessageBox(self)
        box.setWindowTitle("به‌روزرسانی کامل شد")
        box.setIcon(QMessageBox.Information)
        box.setText("فایل ۰۱ روی کتاب کار اعمال و در دیتابیس ذخیره شد.")
        box.setInformativeText("\n".join(lines))
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    @guarded("_on_system_update_done")
    def _on_system_update_done(self, result: Dict[str, Any]) -> None:
        plan: SystemUpdatePlan = result["plan"]
        self.config.set("last_excel_path", plan.source_path, save=False)
        self.settings.excel_edit.setText(plan.source_path)
        self.settings.set_status(
            f"به‌روزرسانی انجام شد — {fa(result['updated_cells'])} سلول در "
            f"{fa(plan.matched_rows)} ردیف"
        )
        self._notify("کتاب کار دیتابیس به‌روزرسانی شد", "check")
        self.load_data()
        lines = [
            f"{fa(plan.matched_rows)} ردیف از «اجرا - مهر» روی شیت «{plan.sheet_name}» اعمال شد.",
            f"سلول‌های نوشته‌شده: {fa(result['updated_cells'])}",
            f"تطبیق: {plan.method}",
        ]
        if plan.unmatched:
            lines.append(
                f"بدون تطبیق: {fa(len(plan.unmatched))} ردیف — نمونه: "
                + "، ".join(plan.unmatched[:4])
            )
        if result.get("backup_path"):
            lines.append(f"نسخهٔ پشتیبان: {result['backup_path']}")
        box = QMessageBox(self)
        box.setWindowTitle("به‌روزرسانی انجام شد")
        box.setIcon(QMessageBox.Information)
        box.setText("کتاب کار دیتابیس با فایل سامانه به‌روزرسانی شد.")
        box.setInformativeText("\n".join(lines))
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    @guarded("_on_system_update_failed")
    def _on_system_update_failed(self, message: str) -> None:
        self.settings.set_status("به‌روزرسانی از فایل سامانه انجام نشد.")
        self._set_system_update_busy(False)
        lowered = (message or "").lower()
        database_problem = any(token in lowered for token in (
            "اتصال به دیتابیس", "اتصال به سرور", "1045", "access denied",
            "نام کاربری یا رمز", "2003", "can't connect",
            "در دیتابیس پیدا نشد", "ذخیره نشده است", "excel_original_files",
        ))
        source = str(getattr(self, "_pending_system_source", "") or "")
        if "کتاب کار اصلی" in (message or ""):
            self._pending_system_source = source
            box = QMessageBox(self)
            box.setWindowTitle("کتاب کار اصلی لازم است")
            box.setIcon(QMessageBox.Warning)
            box.setText(
                "فایل ۰۱ سامانه ستون‌های وضعیت و درصد پیشرفت را ندارد. برای اینکه "
                "کارت‌ها و درصد صفر نشوند، اول باید کتاب کار اصلی شما (فایل دارای "
                "شیت colorcode) مشخص باشد."
            )
            box.setInformativeText(
                "یک بار دکمهٔ «کتاب کار اصلی (۰۲)» را بزنید و فایل اصلی خودتان "
                "(مثل \u202aF.O.C(Mehr).xlsx\u202c که شیت colorcode دارد) را انتخاب کنید. "
                "بعد از آن، همین فایل ۰۱ روی همان اعمال می‌شود و اعداد داشبورد "
                "مثل قبل برمی‌گردند."
            )
            pick_btn = box.addButton("انتخاب کتاب کار اصلی", QMessageBox.AcceptRole)
            only_btn = box.addButton(
                "ادامه با فایل ۰۱ (بدون وضعیت/درصد)", QMessageBox.DestructiveRole
            )
            box.addButton("بعداً", QMessageBox.RejectRole)
            box.exec_()
            if box.clickedButton() is pick_btn:
                self.pick_source_workbook()
            elif box.clickedButton() is only_btn and source:
                self._allow_system_only = True
                self.config.set("allow_system_only", True)
                self.start_local_system_update(source)
            return
        if database_problem and source:
            if ask_confirm(
                self, "دیتابیس در دسترس نیست",
                "برای نوشتن در دیتابیس مشکلی هست (اتصال یا کتاب کار ذخیره‌شده).\n\n"
                "می‌خواهید همین فایل ۰۱ را داخل خودِ برنامه اعمال کنم؟ برنامه کتاب کار "
                "داخلی خودش را می‌سازد/به‌روز می‌کند و داشبورد با همان داده اجرا می‌شود "
                "— بدون دیتابیس و بدون فایل ۰۲.",
                "بله، داخل برنامه انجام بده",
            ):
                self.start_local_system_update(source)
                return
        log_event("به‌روزرسانی از فایل سامانه ناموفق: "
                  + " / ".join((message or "").splitlines()[:3]))
        box = QMessageBox(self)
        box.setWindowTitle("به‌روزرسانی از فایل سامانه")
        box.setIcon(QMessageBox.Warning)
        box.setText("به‌روزرسانی انجام نشد.")
        box.setInformativeText(f"{message}\n\nفایل گزارش:\n{log_path()}")
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    @guarded("import_excel")
    def import_excel(self, path: str) -> None:
        if self._import_worker_running():
            self._notify("یک بارگذاری در حال اجراست؛ کمی صبر کنید", "clock")
            return
        if not str(path or "").strip():
            self.import_excel_dialog()
            return
        if Path(path).exists() and not has_colorcode_sheet(path):
            sheets = workbook_sheet_names(path)
            self._system_update_note = (
                "این فایل شیت colorcode ندارد، پس به‌عنوان «فایل ۰۱ سامانه» خوانده "
                f"می‌شود (شیت‌ها: {'، '.join(sheets) or '—'})."
            )
            self.start_system_update(path)
            return
        if not str(self.config.get("db_host") or "").strip():
            # دیتابیسی تنظیم نشده است؛ همان فایل دارای colorcode محلی اجرا می‌شود.
            self._notify("دیتابیس تنظیم نشده؛ فایل به‌صورت محلی اجرا می‌شود", "table")
            self.open_local_excel(path)
            return
        rows_in_file = count_workbook_contracts(path, str(self.config.get("sheet_name") or ""))
        current = int(getattr(self.data, "total_contracts", 0) or 0)
        if rows_in_file and current >= 100 and rows_in_file < current * 0.6:
            answer = QMessageBox.question(
                self, "تعداد قراردادها کم است",
                f"فایل انتخاب‌شده فقط {fa(rows_in_file)} قرارداد دارد، در حالی که "
                f"کتاب کار فعلی {fa(current)} قرارداد دارد.\n\n"
                "اگر این فایل را بارگذاری کنید، داشبورد با همین تعداد کم نمایش داده "
                "می‌شود. مطمئنید همین فایل درست است؟",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                self.settings.set_status("بارگذاری لغو شد — تعداد قراردادها کمتر بود")
                return
        self.show_page("settings")
        values = self.settings.connection_values()
        settings = {
            "host": values["host"],
            "port": values["port"],
            "database": values["database"],
            "user": values["user"],
            "password": values["password"],
            "workbook_id": values["workbook_id"],
            "sheet_name": values["sheet_name"],
        }
        self.settings.set_importing(True, path)
        self.settings.set_status(
            "در حال بارگذاری فایل در دیتابیس… (اگر کاربر MySQL رمز ندارد، فیلد رمز خالی باشد)"
        )
        worker = ImportWorker(settings, path, values["sheet_name"], self)
        worker.succeeded.connect(
            lambda info, values=values, path=path: self._on_import_succeeded(info, values, path)
        )
        worker.failed.connect(self._on_import_failed)
        worker.finished.connect(lambda: self.settings.set_importing(False))
        self._import_worker = worker
        worker.start()

    def _adopt_last_connection(self) -> None:
        """اگر ورود جایگزین (مثلاً root بدون رمز) جواب داد، همان را ذخیره می‌کند."""
        info = dict(LAST_CONNECTION or {})
        if not info:
            return
        user = str(info.get("user") or "")
        password = str(info.get("password") or "")
        if not user:
            return
        changed = (
            user != str(self.config.get("db_user") or "")
            or password != str(self.config.get_password() or "")
        )
        if not changed:
            return
        self.config.set("db_user", user, save=False)
        self.config.set_password(password, self.config.get("db_remember_password"))
        self.config.save()
        self.settings.apply_connection({"user": user, "password": password})
        log_event(f"ورود جایگزین ذخیره شد: کاربر={user}")
        self.settings.set_status(
            "اتصال با کاربر «" + user + "» برقرار شد"
            + (" (بدون رمز)" if not password else "")
        )
        self._notify(
            "اتصال برقرار شد — کاربر " + user + (" (بدون رمز)" if not password else ""),
            "check",
        )

    def _import_worker_running(self) -> bool:
        worker = getattr(self, "_import_worker", None)
        if worker is None:
            return False
        try:
            return bool(worker.isRunning())
        except RuntimeError:
            self._import_worker = None
            return False

    def _on_import_succeeded(self, info: Dict[str, Any],
                             values: Dict[str, Any], path: str) -> None:
        mapping = {
            "db_host": values.get("host"),
            "db_port": values.get("port"),
            "db_name": values.get("database"),
            "db_user": values.get("user"),
            "db_workbook_id": values.get("workbook_id"),
            "sheet_name": values.get("sheet_name"),
        }
        for key, value in mapping.items():
            if value not in (None, ""):
                self.config.set(key, value, save=False)
        self.config.set_password(
            str(values.get("password") or ""), self.config.get("db_remember_password")
        )
        self.config.set("last_excel_path", path, save=False)
        self.config.set("data_source", "database", save=False)
        self.config.save()
        self.settings.apply_connection(values)
        self._adopt_last_connection()
        rows_text = (
            f"، {fa(int(info['data_rows']))} ردیف در mehr_contracts"
            if info.get("data_rows") else ""
        )
        self.settings.set_status(
            f"فایل اکسل بارگذاری شد — {fa(info.get('size', 0))} بایت، "
            f"{fa(len(info.get('sheets') or []))} شیت{rows_text}"
        )
        self._notify("فایل اکسل در دیتابیس ذخیره شد؛ جدول mehr_contracts به‌روز است", "check")
        self.statusBar().showMessage("بارگذاری اکسل انجام شد؛ داده‌ها در حال خواندن…", 6000)
        self.load_data()

    @guarded("_on_import_failed")
    def _on_import_failed(self, message: str) -> None:
        self.settings.set_status("بارگذاری فایل اکسل ناموفق بود.")
        log_event("بارگذاری اکسل ناموفق: "
                  + " / ".join((message or "").splitlines()[:3]))
        self._notify("بارگذاری فایل اکسل ناموفق بود", "alert")
        box = QMessageBox(self)
        box.setWindowTitle("بارگذاری فایل اکسل")
        box.setIcon(QMessageBox.Critical)
        box.setText("ذخیرهٔ فایل اکسل در دیتابیس ممکن نشد.")
        box.setInformativeText(f"{message}\n\nفایل گزارش:\n{log_path()}")
        box.addButton("باشه", QMessageBox.AcceptRole)
        box.exec_()

    def pick_pdf_folder(self) -> None:
        start = str(self.settings.pdf_edit.text() or app_dir())
        folder = QFileDialog.getExistingDirectory(self, "انتخاب پوشهٔ PDF", start)
        if folder:
            self.settings.pdf_edit.setText(folder)
            self.config.set("pdf_folder", folder)
            self.load_data()

    def save_settings(self, values: Dict[str, Any]) -> None:
        password = str(values.pop("db_password", "") or "")
        remember = bool(values.pop("db_remember_password", True))
        for key, value in values.items():
            self.config.set(key, value, save=False)
        self.config.set("db_remember_password", remember, save=False)
        self.config.set_password(password, remember)
        self.config.save()
        theme = str(values.get("theme") or "light")
        font_size = int(values.get("font_size") or 10)
        self.apply_theme(theme, font_size, persist=False)
        self.settings.set_status(f"ذخیره شد — {jalali_datetime(datetime.now())}")
        self._setup_auto_timer()
        self.statusBar().showMessage("تنظیمات ذخیره شد.", 4000)
        self._notify("تنظیمات ذخیره شد", "check")
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
        if text.strip() and self._current_page != "contracts":
            self.show_page("contracts")

    def _on_filter_requested(self, key: str) -> None:
        self.contracts.set_filter(key)
        search = self.topbar.search_edit.text().strip()
        if search:
            self.contracts.set_search(search)
        self.show_page("contracts")

    def apply_view(self, filter_key: Optional[str] = None) -> None:
        search = self.topbar.search_edit.text().strip()
        key = filter_key if filter_key is not None else self.contracts.current_filter()
        self.contracts.set_search(search)
        self.contracts.model.apply_view(search, key, key == "bookmarked")
        self.contracts._refresh_counts()

    def on_card_clicked(self, filter_key: str) -> None:
        mapping = {"all": "all", "progress": "all", "deductions": "deductions", "pdf": "pdf"}
        key = mapping.get(filter_key or "all", "all")
        self.contracts.set_filter(key)
        self.apply_view(filter_key=key)
        self.show_page("contracts")
        if filter_key == "progress":
            self._notify("برای دیدن جزئیات پیشرفت، از فیلترهای جدول استفاده کنید", "info")

    def on_slice_clicked(self, field: str, label: str) -> None:
        key = f"{field}:{label}"
        self.contracts.set_filter(key)
        self.show_page("contracts")

    # ------------------------------------------------------------------
    # جزئیات، یادداشت و PDF
    # ------------------------------------------------------------------
    def show_record_detail(self, record: Optional[ContractRecord]) -> None:
        if record is None:
            return
        self.detail.record = None
        self.detail.set_record(record)
        self.detail.open_drawer()

    def _on_notes_changed(self, key: str) -> None:
        self.contracts.refresh_bookmarks()
        self.config.set("last_author", self.detail.author_edit.text().strip(), save=False)
        self.config.save()
        self.statusBar().showMessage("یادداشت ذخیره شد.", 3000)
        self._notify("یادداشت ذخیره شد", "check")

    def _on_bookmark_changed(self, key: str) -> None:
        self.contracts.refresh_bookmarks()
        flagged = self.notes.is_bookmarked(key)
        self.statusBar().showMessage(
            "قرارداد نشان‌گذاری شد." if flagged else "نشان قرارداد برداشته شد.", 3000
        )

    def toggle_bookmark(self, key: str) -> None:
        if not key:
            return
        flagged = self.notes.toggle_bookmark(key)
        self.contracts.refresh_bookmarks()
        record = self.data.find(key)
        if record and self.detail.record is not None and self.detail.record.key == key:
            self.detail.set_record(record)
        self.statusBar().showMessage(
            "قرارداد نشان‌گذاری شد." if flagged else "نشان قرارداد برداشته شد.", 3000
        )

    def _pdf_path(self, filename: str) -> str:
        folder = self.data.pdf_folder or resolve_pdf_folder(self.config)
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
        visible = self.contracts.visible_records()
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
        if QPrinter is None or QPrintDialog is None:
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

    def _open_data_folder(self) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(data_dir())))

    # ------------------------------------------------------------------
    # پوسته
    # ------------------------------------------------------------------
    def apply_theme(self, theme: str, font_size: int = 10, persist: bool = True) -> None:
        app = QApplication.instance()
        self.colors = apply_theme(app, theme, font_size)
        app.setProperty("mehr_theme", theme)
        if persist:
            self.config.set("theme", theme, save=False)
            self.config.set("font_size", font_size, save=False)
            self.config.save()
        for widget in (self.dashboard, self.contracts, self.reports,
                       self.pdfs, self.settings, self.help):
            if hasattr(widget, "set_theme_colors"):
                widget.set_theme_colors(self.colors)
        self.rail.set_theme(self.colors)
        self.rail.set_dark(theme == "dark")
        self.topbar.set_theme(self.colors)
        self.topbar.set_dark(theme == "dark")
        self.detail.set_theme_colors(self.colors)
        for combo in self.findChildren(Combo):
            combo.set_theme_colors(self.colors)
        for check in self.findChildren(CheckBox):
            check.set_theme_colors(self.colors)
        self.toast.colors = self.colors
        self.statusBar().setStyleSheet(
            f"QStatusBar {{ background: {self.colors['surface']};"
            f" border-top: 1px solid {self.colors['border_soft']}; }}"
        )
        self.central.setStyleSheet(f"QWidget#Central {{ background: {self.colors['bg']}; }}")

    def toggle_theme(self) -> None:
        current = str(self.config.get("theme") or "light")
        new_theme = "dark" if current == "light" else "light"
        self.apply_theme(new_theme, int(self.config.get("font_size") or 10))
        self.config.set("theme", new_theme)
        self.settings.load_values(self.config)
        self._notify("پوستهٔ روشن" if new_theme == "light" else "پوستهٔ تیره", "sun")

    def show_about(self) -> None:
        QMessageBox.about(
            self, "دربارهٔ برنامه",
            f"<div dir='rtl' style='font-family: Tahoma;'>"
            f"<h3>{APP_NAME}</h3>"
            f"<p>نسخهٔ {APP_VERSION} — برنامهٔ دسکتاپ تک‌فایلی بر پایهٔ PyQt5</p>"
            f"<p>داده‌ها مستقیماً از دیتابیس مرکزی MySQL (کتاب کار شناسهٔ ۱، شیت colorcode) "
            f"خوانده می‌شوند و به Flask، مرورگر، فایل HTML یا فایل اکسل محلی نیازی نیست.</p>"
            f"<p style='color:#66756f;'>مسیر تنظیمات: {self.config.path}</p>"
            f"</div>",
        )

    @guarded("_notify")
    def _notify(self, text: str, icon: str = "check") -> None:
        try:
            self.toast.show_message(text, icon, self.colors)
        except Exception:
            pass

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

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        if hasattr(self, "detail"):
            self.detail.place()

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
if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # pragma: no cover
        import traceback as _tb
        _tb.print_exc()
        log_file = write_error_log(_tb.format_exc())
        try:
            app = QApplication.instance() or QApplication(sys.argv[:1])
            message = f"برنامه با خطا متوقف شد:\n{exc}"
            if log_file:
                message += f"\n\nجزئیات در این فایل ذخیره شد:\n{log_file}"
            QMessageBox.critical(None, "خطای اجرای برنامه", message)
        except Exception:
            pass
        sys.exit(1)

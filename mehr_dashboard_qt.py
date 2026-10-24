#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""داشبورد قراردادهای مهر خلیج فارس — نسخهٔ ۴٫۰٫۰ (بازنویسی کامل از صفر)

یک فایل، دو چیز:
  ۱) داشبوردِ قراردادها از روی کتاب کار `colorcode`؛
  ۲) به‌روزرسانی خودکار همان کتاب کار از خروجی سامانه (فایل ۰۱).

همهٔ داده‌ها در ویندوز داخل `D:\\DATANEW` نگهداری می‌شوند:
    D:\\DATANEW\\database\\mehr.db      دیتابیس داخلی (SQLite)
    D:\\DATANEW\\workbooks\\            نسخهٔ کتاب کارهای ذخیره‌شده
    D:\\DATANEW\\output\\               چاپ نسخهٔ به‌روزِ ساخته‌شده
    D:\\DATANEW\\exports\\              خروجی اکسل و گزارش‌ها
    D:\\DATANEW\\backup\\               پشتیبان دیتابیس
    D:\\DATANEW\\logs\\                 گزارش برنامه و خطاها

اجرا:
    python mehr_dashboard_qt.py                 پنجرهٔ برنامه
    python mehr_dashboard_qt.py --version       نسخه
    python mehr_dashboard_qt.py --check         بررسی داده‌ها بدون پنجره
    python mehr_dashboard_qt.py --import-xlsx PATH   بارگذاری کتاب کار
    python mehr_dashboard_qt.py --update-from PATH   به‌روزرسانی از فایل ۰۱
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import traceback
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

warnings.filterwarnings("ignore", message=".*marked as a date.*")

APP_NAME = "داشبورد قراردادهای مهر خلیج فارس"
APP_VERSION = "4.0.0"
DATA_FOLDER_NAME = "DATANEW"
DEFAULT_SHEET = "colorcode"
EXPORT_SHEET_TITLE = "خروجی کندو"

try:  # openpyxl تنها کتابخانهٔ اجباری برای اکسل است
    import openpyxl
except Exception:  # pragma: no cover
    openpyxl = None

try:
    from PyQt5.QtCore import Qt, QThread, QTimer, pyqtSignal
    from PyQt5.QtGui import QBrush, QColor, QFont, QPainter, QPen
    from PyQt5.QtWidgets import (
        QAbstractItemView, QApplication, QComboBox, QDialog, QFileDialog, QFrame,
        QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox,
        QPlainTextEdit, QProgressBar, QPushButton, QScrollArea, QSizePolicy,
        QSpinBox, QSplitter, QStackedWidget, QTableWidget, QTableWidgetItem,
        QVBoxLayout, QWidget,
    )
    QT_READY = True
    QT_ERROR = ""
except Exception as exc:  # pragma: no cover - روی سیستمی که PyQt5 ندارد
    QT_READY = False
    QT_ERROR = str(exc)

try:
    from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
except Exception:  # pragma: no cover
    QPrintDialog = None
    QPrinter = None


# ===========================================================================
# ۱) مسیرها — همه‌چیز در D:\DATANEW
# ===========================================================================
_LOG_LINES: List[str] = []
_DATA_ROOT_CACHE: Optional[Path] = None


def app_dir() -> Path:
    """پوشهٔ برنامه."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_data_root(windows: Optional[bool] = None) -> Path:
    """پوشهٔ دادهٔ پیش‌فرض: `D:\\DATANEW` (روی ویندوز)."""
    is_windows = (os.name == "nt") if windows is None else bool(windows)
    if is_windows:
        return Path("D:/") / DATA_FOLDER_NAME
    # روی سیستم‌های غیرِ ویندوز (آزمون/لینوکس) همان نام در خانهٔ کاربر.
    return Path.home() / DATA_FOLDER_NAME


def data_root() -> Path:
    """ریشهٔ داده‌ها؛ قابل تغییر با متغیر محیطی MEHR_DATA_DIR."""
    global _DATA_ROOT_CACHE
    if _DATA_ROOT_CACHE is not None:
        return _DATA_ROOT_CACHE
    override = str(os.environ.get("MEHR_DATA_DIR") or "").strip()
    if override:
        target = Path(override).expanduser()
    else:
        target = default_data_root()
    try:
        target.mkdir(parents=True, exist_ok=True)
        probe = target / ".write-test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except Exception as exc:
        fallback = app_dir() / DATA_FOLDER_NAME
        fallback.mkdir(parents=True, exist_ok=True)
        _DATA_ROOT_CACHE = fallback
        log_line(
            f"پوشهٔ «{target}» قابل نوشتن نبود ({exc})؛ پوشهٔ «{fallback}» استفاده می‌شود."
        )
        return fallback
    _DATA_ROOT_CACHE = target
    return target


def data_subdir(name: str) -> Path:
    """زیرپوشه‌های استاندارد داده را می‌سازد."""
    folder = data_root() / name
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def database_path() -> Path:
    return data_subdir("database") / "mehr.db"


def logs_dir() -> Path:
    return data_subdir("logs")


def log_line(message: str) -> None:
    """یک خط در حافظه و در فایل `logs/mehr.log` ثبت می‌کند."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    text = f"[{stamp}] {message}"
    _LOG_LINES.append(text)
    if len(_LOG_LINES) > 2000:
        del _LOG_LINES[:500]
    try:
        folder = logs_dir()
        with open(folder / "mehr.log", "a", encoding="utf-8") as handle:
            handle.write(text + "\n")
    except Exception:
        pass


def log_text() -> str:
    return "\n".join(_LOG_LINES)


def write_error_log(detail: str) -> Optional[Path]:
    """خطاهای پیش‌بینی‌نشده را جدا ذخیره می‌کند."""
    try:
        path = logs_dir() / "errors.log"
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n")
            handle.write(detail.rstrip() + "\n")
        return path
    except Exception:
        return None


# ===========================================================================
# ۲) ابزارهای متن و عدد
# ===========================================================================
_PERSIAN_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def fa(value: Any) -> str:
    """عدد/متن را با رقم‌های فارسی برمی‌گرداند."""
    if isinstance(value, bool):
        return "بله" if value else "خیر"
    if isinstance(value, int):
        return f"{value:,}".translate(_PERSIAN_DIGITS)
    if isinstance(value, float):
        text = f"{value:,.2f}".rstrip("0").rstrip(".")
        return text.translate(_PERSIAN_DIGITS)
    return str(value if value is not None else "").translate(_PERSIAN_DIGITS)


def txt(value: Any) -> str:
    """متن تمیز یک سلول."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def num(value: Any) -> Optional[float]:
    """عدد را از سلول می‌خواند (درصد متنی و رقم فارسی هم پذیرفته می‌شود)."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().translate(
        str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    )
    if not text:
        return None
    percent = text.endswith("%") or text.endswith("٪")
    text = text.replace("%", "").replace("٪", "").replace(",", "").replace("٬", "")
    text = re.sub(r"[^\d.\-+eE]", "", text)
    if not text or text in {"-", ".", "+"}:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return number / 100.0 if percent and abs(number) > 1.5 else number


def norm(value: Any) -> str:
    """یکدست‌سازی فارسی/عربی و فاصله‌ها برای مقایسه."""
    text = str(value or "")
    text = (text.replace("ي", "ی").replace("ك", "ک").replace("\u200c", " ")
                .replace("\u200f", "").replace("\u200e", ""))
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def key_of(value: Any) -> str:
    """کلید مقایسهٔ شمارهٔ قرارداد/نام فایل (حروف و رقم فقط)."""
    return re.sub(r"[^0-9a-z\u0600-\u06ff]", "", norm(value))


SUB_PROJECT_WORDS = {"اجرا-مهر", "اجرامهر", "اجرا مهر"}


def is_mehr_subproject(value: Any) -> bool:
    """آیا ستون Sub Project همان «اجرا - مهر» است؟"""
    text = re.sub(r"\s+", "", norm(value))
    compact = text.replace("-", "").replace("_", "").replace("–", "").replace("—", "")
    return compact == "اجرامهر"


STATUS_MARKS = {"m", "م", "✓", "✔"}


def is_mark(value: Any) -> bool:
    """آیا سلول علامت «ثبت‌شده» است؟ (فقط M/م/تیک)"""
    if value is None or isinstance(value, (bool, int, float)):
        return False
    return norm(value) in STATUS_MARKS


def blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


# ===========================================================================
# ۳) ماتریس الزام مدارک («جدول شماره ۷» کتاب کار)
# ===========================================================================
DOC_NAMES: Tuple[str, ...] = (
    "صورتجلسه تحویل موقت",
    "صورت وضعیت ماقبل قطعی",
    "صورت جلسه برچیدن کارگاه",
    "صورتجلسه فصل مشترک",
    "گواهی رفع پانچ",
    "بالانس متریال",
    "مدارک مهندسی نهایی شده - Final Doc",
    "لایحه تاخیرات",
    "ادعاها، اضافه یا کسر کارها",
    "تسویه حساب کارگاهی",
    "صورت وضعیت قطعی",
    "صورتجلسه کارگروه اختتام - کمیسیون پایان قرارداد",
    "مفاصا حساب بیمه",
    "اقرار نامه ترک ادعاها",
    "آزاد سازی سپرده حسن انجام کار - بخش اول",
    "آزاد سازی تضامین",
    "صورت جلسه اختتام - اتمام دوره گارانتی",
    "صورتجلسه تحویل قطعی",
    "آزاد سازی سپرده حسن انجام کار - بخش دوم",
    "تسویه حساب مالی",
    "گواهی پایان کار",
    "گواهی رضایت کاری",
    "صورت جلسه اختتام - تحویل قطعی",
)

MATRIX: Dict[str, Tuple[str, ...]] = {
    "EPC": ("M", "M", "M", "O", "M", "M", "M", "M", "M", "M", "M", "M", "M", "M",
            "M", "M", "M", "M", "M", "M", "O", "O", "M"),
    "PC": ("M", "M", "M", "O", "M", "M", "M", "M", "M", "M", "M", "M", "M", "M",
           "M", "M", "M", "M", "M", "M", "O", "O", "M"),
    "C": ("M", "M", "M", "O", "M", "M", "M", "M", "M", "M", "M", "M", "M", "M",
          "M", "M", "M", "M", "M", "M", "O", "O", "M"),
    "PR": ("N/A", "N/A", "N/A", "N/A", "N/A", "M", "M", "M", "M", "N/A", "M",
           "M", "M", "M", "N/A", "M", "M", "N/A", "N/A", "M", "O", "O", "N/A"),
    "MTR": ("N/A", "N/A", "N/A", "N/A", "N/A", "M", "M", "M", "M", "N/A", "M",
            "M", "M", "M", "N/A", "M", "M", "N/A", "N/A", "M", "O", "O", "N/A"),
    "GNR": ("N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "O", "N/A",
            "M", "O", "M", "M", "N/A", "M", "O", "N/A", "N/A", "M", "N/A",
            "N/A", "O"),
    "E": ("O", "N/A", "N/A", "O", "N/A", "N/A", "N/A", "M", "M", "M", "M", "M",
          "M", "M", "M", "M", "O", "O", "O", "M", "O", "O", "O"),
    "LOG": ("N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "O", "M", "M", "M",
            "M", "M", "M", "M", "M", "O", "O", "O", "M", "O", "O", "O"),
    "TRS": ("N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "O", "M", "M", "M",
            "M", "M", "M", "M", "M", "O", "O", "O", "M", "O", "O", "O"),
    "RNT": ("N/A", "N/A", "O", "O", "N/A", "N/A", "N/A", "O", "O", "M", "M", "M",
            "M", "M", "M", "M", "N/A", "N/A", "O", "M", "O", "O", "O"),
    "MTN": ("N/A", "N/A", "O", "O", "N/A", "O", "N/A", "O", "O", "M", "M", "M",
            "M", "M", "M", "M", "O", "O", "O", "M", "O", "O", "O"),
}

TYPE_NAMES = {
    "EPC": ("EPC", "طراحی، تامین کالا و اجرا"),
    "PC": ("PC", "تامین کالا و اجرا"),
    "C": ("C", "اجرایی"),
    "PR": ("PR", "تأمین کالا"),
    "MTR": ("MTR", "تأمین متریال و مصالح"),
    "GNR": ("GNR", "خرید اقلام عمومی"),
    "E": ("E", "خدمات مشاوره و مهندسی"),
    "LOG": ("LOG", "خدمات پشتیبانی"),
    "TRS": ("TRS", "حمل و نقل"),
    "RNT": ("RNT", "اجاره"),
    "MTN": ("MTN", "نگهداری و تعمیرات"),
}

REQ_CODES = {"M", "O", "N", "A", "N/A", "NA", "-"}


def req_code(value: Any) -> str:
    text = re.sub(r"\s+", "", txt(value)).upper()
    if text in {"M", "O"}:
        return text
    if text in {"N/A", "NA", "N", "A", "-", ""}:
        return "N/A"
    return text


def type_key(value: Any) -> str:
    """نوع قرارداد را به کلید ماتریس تبدیل می‌کند."""
    text = txt(value)
    if not text:
        return ""
    upper = re.sub(r"\s+", " ", text).upper()
    if upper in MATRIX:
        return upper
    for code, names in TYPE_NAMES.items():
        for name in names:
            if norm(name) == norm(text):
                return code
    first = re.split(r"[\s/\-–—,،(]", upper, maxsplit=1)[0].strip()
    if first in MATRIX:
        return first
    return ""


def requirements_for(type_value: Any,
                     matrix: Optional[Dict[str, Sequence[str]]] = None) -> List[str]:
    key = type_key(type_value)
    table = matrix or MATRIX
    codes = table.get(key) if key else None
    if not codes:
        return []
    return [req_code(item) for item in codes]


def combos_for(docs: Sequence[Any], requirements: Sequence[str]) -> List[str]:
    """ستون‌های BK..CG: «1-M» یعنی ثبت‌شدهٔ الزامی، «0-M» یعنی الزامیِ ثبت‌نشده."""
    out: List[str] = []
    for index in range(len(DOC_NAMES)):
        code = req_code(requirements[index]) if index < len(requirements) else ""
        number = num(docs[index]) if index < len(docs) else None
        have = "1" if (number is not None and abs(number) > 0) else "0"
        out.append(f"{have}-{code}" if code else "")
    return out


def doc_state(count_value: Any, requirement: Any, combo: Any,
              percent: Optional[float]) -> str:
    """همان فرمول شیت colorcode: M فقط وقتی مدرک لازم است، ثبت نشده و پیشرفت > ۰."""
    number = num(count_value)
    if number is not None and abs(number) > 0:
        return "OK"
    if percent is None or percent <= 0:
        return "OK"
    text = txt(combo).strip()
    if re.match(r"^0\s*-\s*M$", text, re.IGNORECASE):
        return "M"
    if re.match(r"^[01]\s*-", text):
        return "OK"
    return "M" if req_code(requirement) == "M" else "OK"


def read_matrix_from_workbook(workbook) -> Dict[str, List[str]]:
    """ماتریس الزام را از شیت «جدول شماره ۷» خود کتاب کار می‌خواند."""
    if workbook is None:
        return {}
    for name in workbook.sheetnames:
        title = str(name)
        if is_dashboard_sheet(title) or norm(title) in norm(EXPORT_SHEET_TITLE):
            continue
        try:
            sheet = workbook[name]
            height = min(int(sheet.max_row or 0), 80)
            width = min(int(sheet.max_column or 0), 40)
            if height < 5 or width < 10:
                continue
            grid = [list(row) for row in sheet.iter_rows(
                min_row=1, max_row=height, max_col=width, values_only=True)]
        except Exception:
            continue
        found: Dict[str, List[str]] = {}
        for row in grid:
            for start in range(len(row)):
                if req_code(row[start]) not in REQ_CODES:
                    continue
                codes: List[str] = []
                for value in row[start:start + 30]:
                    if req_code(value) not in REQ_CODES:
                        break
                    codes.append(req_code(value))
                if len(codes) < 15:
                    continue
                for offset in range(1, 4):
                    column = start - offset
                    if column < 0:
                        break
                    code = type_key(row[column])
                    if code and code not in found:
                        found[code] = codes[:23]
                        break
        if len(found) >= 3:
            log_line(f"ماتریس الزام از شیت «{title}» خوانده شد ({fa(len(found))} نوع).")
            return found
    return {}


def is_dashboard_sheet(name: Any) -> bool:
    text = norm(name)
    return "colorcode" in text or "color code" in text or "کالرکد" in text


# ===========================================================================
# ۴) مدل داده و خواندن کتاب کار
# ===========================================================================
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


@dataclass
class Contract:
    """یک قرارداد، همان‌طور که در شیت colorcode هست."""

    excel_row: int
    contract_no: str = ""
    contract_id: str = ""
    company: str = ""
    subject: str = ""
    contract_type: str = ""
    employer: str = ""
    percent: Optional[float] = None
    status: str = ""
    problem: str = ""                                       # شرح وضعیت/کسورات
    docs: List[str] = field(default_factory=list)          # ۲۳ وضعیت OK/M
    counts: List[Any] = field(default_factory=list)        # ۲۳ شمارش خام
    flags: Dict[str, bool] = field(default_factory=dict)
    pdf_file: str = ""

    @property
    def key(self) -> str:
        return key_of(self.contract_no) or key_of(self.contract_id)

    @property
    def is_mark(self) -> bool:
        return bool(self.docs) and any(value == "M" for value in self.docs)

    @property
    def open_docs(self) -> List[str]:
        """نام مدرک‌هایی که «M» دارند."""
        return [DOC_NAMES[i] if i < len(DOC_NAMES) else f"مدرک {i + 1}"
                for i, value in enumerate(self.docs) if value == "M"]

    @property
    def has_problem(self) -> bool:
        """آیا برای این قرارداد شرح مشکل/کسورات ثبت شده است؟"""
        return bool(txt(self.problem).strip())


@dataclass
class Dashboard:
    """همهٔ چیزی که داشبورد و جدول لازم دارند."""

    contracts: List[Contract] = field(default_factory=list)
    source_name: str = ""
    sheet_name: str = DEFAULT_SHEET
    progress: Optional[float] = None          # ۰..۱
    percent: Optional[float] = None           # ۰..۱۰۰
    loaded_at: Optional[datetime] = None
    workbook_id: int = 0
    note: str = ""
    doc_titles: List[str] = field(default_factory=list)

    @property
    def total(self) -> int:
        return len(self.contracts)

    @property
    def with_deductions(self) -> int:
        return sum(1 for item in self.contracts if item.has_problem)

    @property
    def incomplete(self) -> int:
        """قراردادهایی که دست‌کم یک مدرک باز (M) دارند."""
        return sum(1 for item in self.contracts if item.is_mark)

    @property
    def open_documents(self) -> int:
        return sum(1 for item in self.contracts for value in item.docs if value == "M")

    def flag_counts(self) -> Dict[str, int]:
        return {key: sum(1 for item in self.contracts if item.flags.get(key))
                for key, _label, _column in FLAG_DEFS}

    def type_counts(self) -> List[Tuple[str, int]]:
        counter: Dict[str, int] = {}
        for item in self.contracts:
            label = item.contract_type or "نامشخص"
            counter[label] = counter.get(label, 0) + 1
        return sorted(counter.items(), key=lambda pair: (-pair[1], pair[0]))

    def status_counts(self) -> List[Tuple[str, int]]:
        counter: Dict[str, int] = {}
        for item in self.contracts:
            label = item.status or "نامشخص"
            counter[label] = counter.get(label, 0) + 1
        return sorted(counter.items(), key=lambda pair: (-pair[1], pair[0]))

    def percent_buckets(self) -> List[Tuple[str, int]]:
        buckets = [("۰٪", 0), ("۱ تا ۵۰٪", 1), ("۵۱ تا ۹۹٪", 51), ("۱۰۰٪", 100)]
        counts = {"۰٪": 0, "۱ تا ۵۰٪": 0, "۵۱ تا ۹۹٪": 0, "۱۰۰٪": 0}
        for item in self.contracts:
            value = (item.percent or 0.0) * 100
            if value <= 0:
                counts["۰٪"] += 1
            elif value < 50:
                counts["۱ تا ۵۰٪"] += 1
            elif value < 99.5:
                counts["۵۱ تا ۹۹٪"] += 1
            else:
                counts["۱۰۰٪"] += 1
        return [(label, counts[label]) for label, _ in buckets]

    def pending_docs(self) -> List[Tuple[str, int]]:
        """برای هر مدرک، چند قرارداد آن را ثبت نکرده است."""
        names = self.doc_titles or list(DOC_NAMES)
        totals = [0] * len(names)
        for item in self.contracts:
            for index, value in enumerate(item.docs):
                if value == "M" and index < len(totals):
                    totals[index] += 1
        pairs = [(names[i], totals[i]) for i in range(len(names))]
        return sorted(pairs, key=lambda pair: -pair[1])


def _percent_from_rows(rows: Sequence[Sequence[Any]]) -> Optional[float]:
    """درصد کل را مثل فرمول خودِ کتاب کار حساب می‌کند: میانگین ردیف‌های >۰."""
    values: List[float] = []
    for row in rows[3:]:
        number = num(row[36]) if len(row) > 36 else None
        if number is not None and number > 0:
            values.append(number)
    if not values:
        return None
    return sum(values) / len(values)


def open_rows(path_or_blob: Any, sheet_hint: str = "") -> Tuple[List[str], str, List[List[Any]]]:
    """شیت داده را باز می‌کند و ردیف‌ها را برمی‌گرداند (بدون فرمول)."""
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست. دستور: pip install openpyxl")
    source = path_or_blob
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(bytes(source))
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=True)
    try:
        names = [str(name) for name in workbook.sheetnames]
        target = ""
        if sheet_hint:
            for name in names:
                if norm(name) == norm(sheet_hint):
                    target = name
                    break
        if not target:
            for name in names:
                if is_dashboard_sheet(name):
                    target = name
                    break
        if not target:
            target = names[0] if names else ""
        if not target:
            raise RuntimeError("کتاب کار هیچ شیتی ندارد.")
        rows = [list(row) for row in workbook[target].iter_rows(values_only=True)]
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    return names, target, rows


def parse_workbook(path_or_blob: Any, sheet_hint: str = "",
                   source_name: str = "", workbook_id: int = 0) -> Dashboard:
    """کتاب کار را به یک داشبورد کامل تبدیل می‌کند."""
    names, sheet, rows = open_rows(path_or_blob, sheet_hint)
    if len(rows) < 4:
        raise ValueError("ساختار شیت غیرمنتظره است (کمتر از ۴ ردیف).")
    width = max((len(row) for row in rows), default=0)
    for row in rows:
        if len(row) < width:
            row.extend([None] * (width - len(row)))

    progress = num(rows[1][36]) if len(rows[1]) > 36 else None
    if progress is None:
        progress = _percent_from_rows(rows)
    header = rows[2] if len(rows) > 2 else []
    doc_columns: List[int] = []
    doc_titles: List[str] = []
    for offset in range(len(DOC_NAMES)):
        column = 11 + offset
        title = txt(header[column]) if column < len(header) else ""
        doc_columns.append(column)
        doc_titles.append(title or DOC_NAMES[offset])

    contracts: List[Contract] = []
    for index in range(3, len(rows)):
        row = rows[index]
        if all(blank(row[column]) for column in range(min(35, len(row)))):
            continue
        docs = [txt(row[column]).upper() if column < len(row) else ""
                for column in doc_columns]
        counts = [row[column] if column < len(row) else None for column in doc_columns]
        flags = {}
        for key, _label, column in FLAG_DEFS:
            flags[key] = is_mark(row[column] if column < len(row) else None)
        contracts.append(Contract(
            excel_row=index + 1,
            contract_no=txt(row[4]) if len(row) > 4 else "",
            contract_id=txt(row[2]) if len(row) > 2 else "",
            company=txt(row[3]) if len(row) > 3 else "",
            subject=txt(row[6]) if len(row) > 6 else "",
            contract_type=txt(row[5]) if len(row) > 5 else "",
            employer=txt(row[7]) if len(row) > 7 else "",
            percent=num(row[36]) if len(row) > 36 else None,
            status=txt(row[35]) if len(row) > 35 else "",
            problem=txt(row[34]) if len(row) > 34 else "",
            docs=docs,
            counts=counts,
            flags=flags,
        ))
    return Dashboard(
        contracts=contracts,
        source_name=source_name or (Path(str(path_or_blob)).name
                                    if not isinstance(path_or_blob, (bytes, bytearray))
                                    else "کتاب کار"),
        sheet_name=sheet,
        progress=progress,
        percent=(round(progress * 100, 2) if progress is not None else None),
        loaded_at=datetime.now(),
        workbook_id=workbook_id,
        doc_titles=doc_titles,
        note=f"شیت‌های کتاب کار: {'، '.join(names)}",
    )


def workbook_mark_count(path_or_blob: Any, sheet_hint: str = "") -> Dict[str, Any]:
    """فقط می‌شمارد: چند علامت M و درصد چقدر است (برای کنترل فایل خروجی)."""
    try:
        names, sheet, rows = open_rows(path_or_blob, sheet_hint)
    except Exception as exc:
        return {"ok": False, "error": str(exc), "marks": 0, "percent": None}
    marks = 0
    for row in rows[3:]:
        for column in range(11, min(34, len(row))):
            if is_mark(row[column]):
                marks += 1
    progress = num(rows[1][36]) if len(rows) > 1 and len(rows[1]) > 36 else None
    if progress is None:
        progress = _percent_from_rows(rows)
    return {
        "ok": True, "marks": marks, "sheet": sheet,
        "percent": (round(progress * 100, 2) if progress is not None else None),
        "rows": max(0, len(rows) - 3), "sheets": names,
    }


def workbook_profile(path_or_blob: Any) -> Dict[str, Any]:
    """کتاب کار برای پایهٔ به‌روزرسانی مناسب است؟ (colorcode + شیت خروجی)"""
    try:
        names, sheet, rows = open_rows(path_or_blob)
    except Exception:
        return {"ok": False}
    has_export = False
    export_name = ""
    for name in names:
        if norm(name) == norm(sheet):
            continue
        if norm(EXPORT_SHEET_TITLE) in norm(name) or "kandoo" in norm(name):
            has_export = True
            export_name = name
            break
    return {"ok": True, "sheets": names, "dashboard": sheet,
            "has_export": has_export, "export_sheet": export_name,
            "marks": sum(1 for row in rows[3:] for column in range(11, min(34, len(row)))
                         if is_mark(row[column]))}


# ===========================================================================
# ۵) خواندن «فایل ۰۱» (خروجی سامانه)
# ===========================================================================
SOURCE_HINTS = {
    "no": ("contract no", "contract no.", "شماره قرارداد", "شماره قرارداد / الحاقیه"),
    "id": ("contract id", "شناسه قرارداد", "شناسه قرارداد (کندو)"),
    "company": ("subcontractor", "sub contractor", "نام شرکت", "پیمانکار", "شرکت"),
    "subject": ("contract subject", "موضوع قرارداد", "موضوع"),
    "employer": ("clinet", "client", "کارفرما"),
    "sub": ("sub project", "subproject", "زیرپروژه", "زیر پروژه"),
    "type": ("excution type", "execution type", "نوع قرارداد", "نوع پیمان"),
}


def _find_column(header: Sequence[Any], hints: Sequence[str]) -> int:
    """ستونی که سرستونش با یکی از نام‌ها هم‌خوان است."""
    cells = [norm(value) for value in header]
    for hint in hints:
        needle = norm(hint)
        if not needle:
            continue
        for index, value in enumerate(cells):
            if value and (value == needle or (len(needle) > 6 and needle in value)):
                return index
    return -1


@dataclass
class SourceRecord:
    row: int
    contract_no: str
    contract_id: str
    company: str
    subject: str
    employer: str
    contract_type: str
    values: List[Any]        # ۲۵ مقدار L..AJ
    docs: List[Any]          # ۲۳ شمارش مدرک
    percent: Optional[float]

    @property
    def key(self) -> str:
        return key_of(self.contract_no) or key_of(self.contract_id)


@dataclass
class SourceFile:
    path: str
    sheet: str
    header_row: int
    headers: List[str]              # ۲۵ نام L..AJ
    doc_names: List[str]            # ۲۳ نام مدرک
    records: List[SourceRecord]
    total_rows: int
    filtered_rows: int
    columns: Dict[str, int]
    sheets: List[str] = field(default_factory=list)
    all_records: List[SourceRecord] = field(default_factory=list)

    @property
    def name(self) -> str:
        return Path(self.path).name


def read_source(path: str) -> SourceFile:
    """فایل ۰۱ سامانه را می‌خواند، «اجرا - مهر» را جدا می‌کند."""
    file_path = Path(str(path)).expanduser()
    if not file_path.exists():
        raise ValueError(f"فایل پیدا نشد:\n{file_path}")
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست.")
    workbook = openpyxl.load_workbook(str(file_path), data_only=True, read_only=True)
    try:
        sheets = [str(name) for name in workbook.sheetnames]
        candidates: List[Tuple[Tuple[int, int, int], str, int]] = []
        for name in sheets:
            head = [list(row) for row in workbook[name].iter_rows(
                max_row=12, max_col=90, values_only=True)]
            hint_hits, header_index = -1, -1
            for index, row in enumerate(head):
                hits = sum(1 for key in SOURCE_HINTS.values()
                           if _find_column(row, key) >= 0)
                if hits > hint_hits:
                    hint_hits, header_index = hits, index
            if hint_hits < 3 or header_index < 0:
                continue
            mehr_hits = 0
            for index, row in enumerate(workbook[name].iter_rows(
                    max_row=min(400, max(60, header_index + 300)), max_col=60,
                    values_only=True)):
                if index <= header_index:
                    continue
                if any(is_mehr_subproject(value) for value in list(row)[:25]):
                    mehr_hits += 1
            candidates.append(((1 if mehr_hits else 0, mehr_hits, hint_hits),
                               name, header_index))
        if not candidates:
            raise ValueError(
                "در این فایل هیچ شیتِ خروجی سامانه پیدا نشد (ستون Sub Project و "
                "Contract No لازم است)."
            )
        candidates.sort(key=lambda item: item[0], reverse=True)
        _score, best_name, best_header = candidates[0]
        rows = [list(row) for row in workbook[best_name].iter_rows(
            max_col=90, values_only=True)]
    finally:
        try:
            workbook.close()
        except Exception:
            pass

    header = rows[best_header] if 0 <= best_header < len(rows) else []
    columns = {key: _find_column(header, hints) for key, hints in SOURCE_HINTS.items()}
    sub_col = columns.get("sub", -1)
    if sub_col < 0:
        for row in rows[best_header + 1:best_header + 30]:
            for index, value in enumerate(row[:20]):
                if is_mehr_subproject(value):
                    sub_col = index
                    break
            if sub_col >= 0:
                break
    if sub_col < 0:
        sub_col = 5
    width = max((len(row) for row in rows), default=0)
    if width < 36:
        raise ValueError(
            f"این فایل فقط {fa(width)} ستون دارد؛ ستون‌های L تا AJ لازم است."
        )

    headers = [txt(header[index]) if index < len(header) else ""
               for index in range(11, 36)]
    doc_names = [headers[index] or DOC_NAMES[index] for index in range(23)]

    def cell(row: Sequence[Any], column: int) -> Any:
        if column < 0 or column >= len(row):
            return None
        return row[column]

    records: List[SourceRecord] = []
    all_records: List[SourceRecord] = []
    total = 0
    for index in range(best_header + 1, len(rows)):
        row = rows[index]
        if all(blank(value) for value in row[:36]):
            continue
        total += 1
        values = [cell(row, 11 + offset) for offset in range(25)]
        record = SourceRecord(
            row=index + 1,
            contract_no=txt(cell(row, columns.get("no", 3))),
            contract_id=txt(cell(row, columns.get("id", 0))),
            company=txt(cell(row, columns.get("company", 2))),
            subject=txt(cell(row, columns.get("subject", 4))),
            employer=txt(cell(row, columns.get("employer", 1))),
            contract_type=txt(cell(row, columns.get("type", 8))),
            values=values,
            docs=values[:23],
            percent=num(values[24]),
        )
        if key_of(record.contract_no) or key_of(record.contract_id):
            all_records.append(record)
        if is_mehr_subproject(cell(row, sub_col)):
            records.append(record)
    if not records:
        raise ValueError(
            "در این فایل هیچ ردیفی با «اجرا - مهر» در ستون Sub Project پیدا نشد.\n"
            f"شیت خوانده‌شده: «{best_name}» — ستون Sub Project: "
            f"{columns.get('sub', -1) + 1}"
        )
    log_line(
        f"فایل ۰۱ خوانده شد: {file_path.name} — شیت «{best_name}»، {fa(total)} ردیف، "
        f"{fa(len(records))} ردیف «اجرا - مهر»."
    )
    return SourceFile(path=str(file_path), sheet=best_name, header_row=best_header + 1,
                      headers=headers, doc_names=doc_names, records=records,
                      total_rows=total, filtered_rows=len(records),
                      columns=columns, sheets=sheets, all_records=all_records)


# ===========================================================================
# ۶) موتور به‌روزرسانی: «خروجی کندو ← فایل ۰۱» و «colorcode ← خروجی کندو»
# ===========================================================================
def _sheet_rows(sheet, max_rows: int = 20000, max_cols: int = 90) -> List[List[Any]]:
    out: List[List[Any]] = []
    for index, row in enumerate(sheet.iter_rows(max_col=max_cols, values_only=True), 1):
        out.append(list(row))
        if index >= max_rows:
            break
    return out


def _header_match_count(header: Sequence[Any], names: Sequence[str]) -> int:
    cells = [norm(value) for value in header]
    hits = 0
    for name in names:
        needle = norm(name)
        if needle and any(needle == cell for cell in cells):
            hits += 1
    return hits


def _pick_export_sheet(workbook, dashboard_name: str) -> str:
    """شیت «خروجی کندو» را با سرستون‌های ۲۳ مدرک پیدا می‌کند."""
    best, best_score = "", 0
    for name in workbook.sheetnames:
        if norm(name) == norm(dashboard_name):
            continue
        try:
            head = _sheet_rows(workbook[name], max_rows=12, max_cols=90)
        except Exception:
            continue
        if not head:
            continue
        score = max(_header_match_count(row, DOC_NAMES) for row in head[:12])
        if norm(EXPORT_SHEET_TITLE) in norm(name):
            score += 20
        has_contract = any(
            _find_column(row, SOURCE_HINTS["no"]) >= 0
            or _find_column(row, SOURCE_HINTS["id"]) >= 0
            for row in head[:12]
        )
        if not has_contract:
            score -= 12
        if score > best_score:
            best, best_score = str(name), score
    return best if best_score >= 8 else ""


def _doc_start(header: Sequence[Any]) -> int:
    """ستون اولین مدرک (L = ۱۱)؛ اگر سرستون‌ها جابه‌جا بودند از نام پیدا می‌شود."""
    cells = [norm(value) for value in header]
    for index, value in enumerate(cells):
        if value and value == norm(DOC_NAMES[0]):
            if index >= 6:
                return index
    return 11


def _map_rows(records: Sequence[Contract], source: SourceFile) -> Tuple[Dict[int, SourceRecord], str]:
    """ردیف‌های colorcode را به ردیف‌های فایل ۰۱ وصل می‌کند."""
    by_no: Dict[str, int] = {}
    by_id: Dict[str, int] = {}
    by_company: Dict[str, List[int]] = {}
    order_rows = [item.excel_row for item in records]
    for index, item in enumerate(records):
        if key_of(item.contract_no):
            by_no.setdefault(key_of(item.contract_no), index)
        if key_of(item.contract_id):
            by_id.setdefault(key_of(item.contract_id), index)
        if key_of(item.company):
            by_company.setdefault(key_of(item.company), []).append(index)

    def ratio(getter) -> float:
        hits = 0
        for item in source.records:
            index = getter(item)
            if index is not None:
                hits += 1
        return hits / max(1, len(source.records))

    no_hits = ratio(lambda item: by_no.get(key_of(item.contract_no), by_id.get(key_of(item.contract_id))))
    id_hits = ratio(lambda item: by_id.get(key_of(item.contract_id)))
    company_hits = ratio(lambda item: (by_company.get(key_of(item.company)) or [None])[0]
                         if len(by_company.get(key_of(item.company)) or []) == 1 else None)

    mapping: Dict[int, SourceRecord] = {}
    method = ""
    if no_hits >= 0.9:
        method = "شمارهٔ قرارداد"
        for item in source.records:
            index = by_no.get(key_of(item.contract_no), by_id.get(key_of(item.contract_id)))
            if index is not None:
                mapping[records[index].excel_row] = item
    elif id_hits >= 0.9:
        method = "شناسهٔ قرارداد"
        for item in source.records:
            index = by_id.get(key_of(item.contract_id))
            if index is not None:
                mapping[records[index].excel_row] = item
    elif company_hits >= 0.9:
        method = "نام شرکت"
        for item in source.records:
            found = by_company.get(key_of(item.company)) or []
            if len(found) == 1:
                mapping[records[found[0]].excel_row] = item
    if len(mapping) < max(3, len(records) // 2) and len(source.records) == len(order_rows):
        method = "ترتیب ردیف‌ها"
        mapping = {order_rows[index]: item
                   for index, item in enumerate(source.records) if index < len(order_rows)}
    if not mapping:
        raise ValueError(
            "هیچ ردیفی از فایل ۰۱ با کتاب کار تطبیق پیدا نکرد.\n"
            f"ردیف‌های «اجرا - مهر»: {fa(len(source.records))} — نمونه: "
            + "، ".join(item.contract_no or item.contract_id
                        for item in source.records[:5])
        )
    return mapping, method


def align_requirements(codes: Sequence[str],
                       names: Sequence[str]) -> List[str]:
    """ترتیب کدهای الزام را با ترتیب ستون‌های یک شیت هم‌راست می‌کند."""
    if len(codes) < len(DOC_NAMES):
        return list(codes)
    out: List[str] = []
    for index, name in enumerate(names):
        needle = norm(name)
        found = -1
        for position, canonical in enumerate(DOC_NAMES):
            if needle and norm(canonical) == needle:
                found = position
                break
        out.append(codes[found] if found >= 0 else codes[index])
    return out


def _requirements_from_matrix(record: SourceRecord,
                              matrix: Dict[str, Sequence[str]]) -> List[str]:
    codes = requirements_for(record.contract_type, matrix)
    if len(codes) < len(DOC_NAMES):
        codes = [req_code(item) for item in record.values[26:49]]
        if len(codes) < len(DOC_NAMES):
            codes = []
    return codes


def _write_export_rows(workbook, export_name: str, records: Sequence[Contract],
                       source: SourceFile, mapping: Dict[int, SourceRecord],
                       matrix: Dict[str, Sequence[str]], info: Dict[str, Any]) -> Dict[str, List[str]]:
    """دادهٔ فایل ۰۱ و محاسبهٔ همان شیت را در «خروجی کندو» می‌نویسد."""
    sheet = workbook[export_name]
    rows = _sheet_rows(sheet, max_rows=12, max_cols=90)
    header_row = -1
    for index, row in enumerate(rows):
        if _header_match_count(row, DOC_NAMES) >= 8:
            header_row = index
            break
    if header_row < 0:
        header_row = 3
    header = rows[header_row] if header_row < len(rows) else []
    doc_start = _doc_start(header)
    # نگاشت دادهٔ L..AJ بر اساس نام سرستون
    source_headers = [norm(item) for item in source.headers]
    mapping_columns: List[int] = []
    for offset in range(25):
        column = doc_start + offset
        if offset < len(source_headers) and source_headers[offset]:
            for index, value in enumerate(header):
                if norm(value) and norm(value) == source_headers[offset]:
                    column = index
                    break
        mapping_columns.append(column)
    # نگاشت ردیف شیت خروجی از روی شمارهٔ قرارداد
    grid = _sheet_rows(sheet, max_rows=20000, max_cols=90)
    no_col = _find_column(header, SOURCE_HINTS["no"])
    if no_col < 0:
        no_col = doc_start - 8
    id_col = _find_column(header, SOURCE_HINTS["id"])
    type_col = _find_column(header, SOURCE_HINTS["type"])
    if type_col < 0:
        type_col = doc_start - 2
    row_by_no: Dict[str, int] = {}
    row_by_id: Dict[str, int] = {}
    for index in range(header_row + 1, len(grid)):
        value = grid[index][no_col] if no_col < len(grid[index]) else None
        if key_of(value):
            row_by_no.setdefault(key_of(value), index + 1)
        value = grid[index][id_col] if 0 <= id_col < len(grid[index]) else None
        if key_of(value):
            row_by_id.setdefault(key_of(value), index + 1)
    first_export_row = header_row + 2
    requirements_by_row: Dict[int, List[str]] = {}
    written = 0
    calculated = 0
    for item in records:
        record = mapping.get(item.excel_row)
        if record is None:
            continue
        row = row_by_no.get(key_of(record.contract_no), row_by_id.get(key_of(record.contract_id), 0))
        if not row:
            row = item.excel_row + (first_export_row - 4)
        for offset in range(25):
            column = mapping_columns[offset] + 1
            value = record.values[offset]
            if blank(value) or column <= 0:
                continue
            cell = sheet.cell(row=row, column=column)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue
            if txt(cell.value) != txt(value):
                cell.value = value
                written += 1
        codes = align_requirements(_requirements_from_matrix(record, matrix),
                                   [txt(value) for value in header[doc_start:doc_start + 23]])
        if len(codes) < len(DOC_NAMES):
            continue
        requirements_by_row[row] = codes
        combos = combos_for(record.docs, codes)
        for offset in range(len(DOC_NAMES)):
            code_cell = sheet.cell(row=row, column=doc_start + 26 + offset)
            if blank(code_cell.value):
                code_cell.value = codes[offset]
                calculated += 1
            combo_cell = sheet.cell(row=row, column=doc_start + 51 + offset)
            if blank(combo_cell.value):
                combo_cell.value = combos[offset]
                calculated += 1
        count_cell = sheet.cell(row=row, column=doc_start + 49)
        if blank(count_cell.value):
            count_cell.value = sum(1 for code in codes if code == "M")
            calculated += 1
        missing_cell = sheet.cell(row=row, column=doc_start + 74)
        if blank(missing_cell.value):
            missing_cell.value = sum(1 for combo in combos if combo == "1-M")
            calculated += 1
    info["export_sheet"] = export_name
    info["export_written"] = written
    info["export_calculated"] = calculated
    info["export_mapping"] = ("نام سرستون" if doc_start != 11 else "ترتیب ستون‌ها")
    return requirements_by_row


def _write_colorcode(workbook, dashboard_name: str, records: Sequence[Contract],
                     source: SourceFile, mapping: Dict[int, SourceRecord],
                     requirements_by_row: Dict[int, List[str]],
                     matrix: Dict[str, Sequence[str]],
                     info: Dict[str, Any],
                     cached: Optional[Sequence[Sequence[Any]]] = None) -> None:
    """نتیجهٔ محاسبه را در شیت colorcode می‌نشاند.

    ردیف‌هایی که در فایل ۰۱ نیستند دست‌نخورده می‌مانند: همان مقداری که داشبورد
    تا امروز نشان می‌داد (از نسخهٔ ذخیره‌شدهٔ همان کتاب کار) دوباره نوشته می‌شود
    تا با ذخیرهٔ مجدد، علامت‌ها و درصدشان گم نشوند.
    """
    sheet = workbook[dashboard_name]
    header = [cell.value for cell in sheet[3]]
    doc_columns: List[int] = []
    for offset in range(len(DOC_NAMES)):
        column = 11 + offset
        needle = norm(DOC_NAMES[offset])
        found = -1
        for index, value in enumerate(header):
            if norm(value) and norm(value) == needle and index >= 6:
                found = index
                break
        if found < 0:
            for index, value in enumerate(header):
                if norm(value) and needle and (needle in norm(value) or norm(value) in needle):
                    found = index
                    break
        if found >= 0:
            column = found
        doc_columns.append(column)
    percent_col = -1
    for index, value in enumerate(header):
        text = norm(value)
        if "درصد" in text and "پیشرفت" in text:
            percent_col = index
            break
    if percent_col < 0:
        percent_col = 36
    source_names = [norm(item) for item in source.doc_names]
    # نگاشت هر ستون colorcode به ستون متناظر در فایل ۰۱
    source_index: List[int] = []
    for position, column in enumerate(doc_columns):
        name = norm(txt(header[column])) if column < len(header) else ""
        match = -1
        for index, value in enumerate(source_names):
            if name and value == name:
                match = index
                break
        if match < 0:
            match = position
        source_index.append(match)

    cached_rows = cached or []

    def cached_value(row: int, column: int) -> Any:
        index = row - 1
        if 0 <= index < len(cached_rows):
            line = cached_rows[index]
            if 0 <= column < len(line):
                return line[column]
        return None

    changed = 0
    kept = 0
    pending: List[Dict[str, Any]] = []
    percents: List[float] = []
    matched_rows = 0

    def put(row: int, column: int, value: Any) -> bool:
        nonlocal changed, kept
        if value is None or value == "":
            return False
        cell = sheet.cell(row=row, column=column + 1)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            pending.append({"row": row, "column": column + 1, "value": value})
            kept += 1
            return True
        if txt(cell.value) == txt(value):
            return False
        cell.value = value
        changed += 1
        return True

    for item in records:
        record = mapping.get(item.excel_row)
        if record is not None:
            matched_rows += 1
            codes = requirements_by_row.get(item.excel_row, [])
            if len(codes) < len(DOC_NAMES):
                codes = _requirements_from_matrix(record, matrix)
            combos = combos_for(record.docs, codes) if len(codes) >= len(DOC_NAMES) else []
            percent = record.percent
            for position, column in enumerate(doc_columns):
                offset = source_index[position]
                count_value = record.docs[offset] if offset < len(record.docs) else None
                code = codes[offset] if offset < len(codes) else ""
                combo = combos[offset] if offset < len(combos) else ""
                put(item.excel_row, column, doc_state(count_value, code, combo, percent))
        else:
            percent = num(cached_value(item.excel_row, percent_col))
            for column in doc_columns:
                state = txt(cached_value(item.excel_row, column)).upper()
                if state in ("OK", "M"):
                    put(item.excel_row, column, state)
        if percent is not None:
            put(item.excel_row, percent_col, percent)
            if percent > 0:
                percents.append(percent)
    if percents:
        total = sum(percents) / len(percents)
        put(2, percent_col, total)
        info["progress_total"] = total
    info["dashboard_sheet"] = dashboard_name
    info["dashboard_marks"] = changed
    info["dashboard_kept"] = kept
    info["pending_values"] = pending
    info["progress_rows"] = matched_rows
    info["dashboard_unmatched"] = len(records) - matched_rows


def _apply_pending(path: Path, pending: Sequence[Dict[str, Any]]) -> int:
    """مقدارهای محاسبه‌شده را جای فرمول‌هایی می‌نویسد که بازمحاسبه نشده‌اند."""
    if not pending or openpyxl is None:
        return 0
    workbook = openpyxl.load_workbook(str(path), data_only=False)
    written = 0
    try:
        sheet = None
        for name in workbook.sheetnames:
            if is_dashboard_sheet(name):
                sheet = workbook[name]
                break
        if sheet is None:
            return 0
        for item in pending:
            try:
                sheet.cell(row=int(item["row"]), column=int(item["column"]),
                           value=item.get("value"))
                written += 1
            except Exception:
                continue
        if written:
            try:
                workbook.calculation.fullCalcOnLoad = True
            except Exception:
                pass
            workbook.save(str(path))
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    return written


def recalculate_with_excel(path: Path) -> bool:
    """فایل را با خودِ اکسل باز، محاسبه و ذخیره می‌کند (فقط ویندوز)."""
    if os.name != "nt" or not Path(path).exists():
        return False
    quoted = str(path).replace("'", "''")
    script = (
        "$ErrorActionPreference = 'Stop'\n"
        "$excel = New-Object -ComObject Excel.Application\n"
        "$excel.Visible = $false\n"
        "$excel.DisplayAlerts = $false\n"
        f"$wb = $excel.Workbooks.Open('{quoted}', 0, $false)\n"
        "$excel.CalculateFullRebuild()\n"
        "$wb.Save()\n"
        "$wb.Close($false)\n"
        "$excel.Quit()\n"
    )
    script_path = Path(tempfile.gettempdir()) / "mehr_recalc.ps1"
    try:
        script_path.write_text(script, encoding="utf-8-sig")
        done = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy",
             "Bypass", "-File", str(script_path)],
            capture_output=True, timeout=600,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if done.returncode != 0:
            log_line("بازمحاسبهٔ اکسل انجام نشد: "
                     + (done.stderr or b"").decode("utf-8", "replace")[-300:])
            return False
        log_line(f"فایل با اکسل بازمحاسبه شد: {Path(path).name}")
        return True
    except Exception as exc:
        log_line(f"بازمحاسبهٔ اکسل ممکن نشد: {exc}")
        return False
    finally:
        try:
            script_path.unlink()
        except Exception:
            pass


def prepare_update(base_blob: bytes, source: SourceFile,
                   out_path: Path) -> Dict[str, Any]:
    """نسخهٔ به‌روز را می‌سازد و کنترل می‌کند که اعداد واقعی داشته باشد."""
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست.")
    workbook = openpyxl.load_workbook(io.BytesIO(base_blob), data_only=False)
    info: Dict[str, Any] = {"output_path": str(out_path), "notes": []}
    try:
        dashboard_name = ""
        for name in workbook.sheetnames:
            if is_dashboard_sheet(name):
                dashboard_name = str(name)
                break
        if not dashboard_name:
            raise ValueError("در کتاب کار پایه شیت colorcode پیدا نشد.")
        export_name = _pick_export_sheet(workbook, dashboard_name)
        if not export_name:
            raise ValueError(
                "در کتاب کار پایه شیت «خروجی کندو» پیدا نشد؛ دادهٔ فایل ۰۱ جایی "
                "برای نشستن ندارد."
            )
        base_dashboard = parse_workbook(base_blob, dashboard_name,
                                        source_name=Path(str(out_path)).name)
        _cached_names, _cached_sheet, cached_rows = open_rows(base_blob, dashboard_name)
        mapping, method = _map_rows(base_dashboard.contracts, source)
        info["method"] = method
        info["matched_rows"] = len(mapping)
        workbook_matrix = read_matrix_from_workbook(workbook)
        matrix = workbook_matrix or dict(MATRIX)
        info["matrix_source"] = "کتاب کار" if workbook_matrix else "پیش‌فرض"
        _write_export_rows(
            workbook, export_name, base_dashboard.contracts, source, mapping, matrix, info
        )
        # ردیف‌های export با ردیف‌های colorcode یکی نیستند؛ کلید مشترک ردیف colorcode
        per_colorcode_row: Dict[int, List[str]] = {}
        for item in base_dashboard.contracts:
            record = mapping.get(item.excel_row)
            if record is None:
                continue
            codes = align_requirements(_requirements_from_matrix(record, matrix),
                                       list(source.doc_names))
            per_colorcode_row[item.excel_row] = codes
        _write_colorcode(workbook, dashboard_name, base_dashboard.contracts,
                         source, mapping, per_colorcode_row, matrix, info,
                         cached_rows)
        try:
            workbook.calculation.fullCalcOnLoad = True
        except Exception:
            pass
        out_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(str(out_path))
    finally:
        try:
            workbook.close()
        except Exception:
            pass

    recalculated = False
    pending = list(info.get("pending_values") or [])
    if pending and os.name == "nt":
        recalculated = recalculate_with_excel(out_path)
        if not recalculated:
            replaced = _apply_pending(out_path, pending)
            info["replaced_formulas"] = replaced
    elif pending:
        replaced = _apply_pending(out_path, pending)
        info["replaced_formulas"] = replaced
    info["recalculated"] = recalculated
    check = workbook_mark_count(str(out_path))
    info["verified"] = bool(check.get("ok") and check.get("marks", 0) > 0)
    info["verified_marks"] = int(check.get("marks") or 0)
    info["verified_percent"] = check.get("percent")
    info["verified_rows"] = int(check.get("rows") or 0)
    info["written_cells"] = (int(info.get("export_written") or 0)
                             + int(info.get("dashboard_marks") or 0)
                             + int(info.get("dashboard_kept") or 0))
    if not info["verified"]:
        info["notes"].append(
            "فایل ساخته‌شده هیچ علامت وضعیت ندارد؛ برای همین جای دادهٔ فعلی نمی‌نشیند."
        )
    log_line(
        "نسخهٔ به‌روز ساخته شد: {name} — {rows} تطبیق ({method})، {marks} علامت، "
        "درصد {percent}".format(
            name=Path(str(out_path)).name, rows=fa(info.get("matched_rows") or 0),
            method=info.get("method") or "—", marks=fa(info["verified_marks"]),
            percent=("—" if info["verified_percent"] is None else fa(info["verified_percent"])),
        )
    )
    return info


# ===========================================================================
# ۷) ساخت کتاب کار پایه از فایل ۰۱ (وقتی هیچ کتاب کاری نیست)
# ===========================================================================
def build_base_workbook(source: SourceFile) -> bytes:
    """از فایل ۰۱ یک کتاب کار کامل (colorcode + خروجی کندو) می‌سازد."""
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست.")
    book = openpyxl.Workbook()
    try:
        book.properties.title = "کتاب کار داخلی برنامه (ساخته‌شده از فایل ۰۱)"
        book.properties.keywords = "mehr-generated"
    except Exception:
        pass
    export = book.active
    export.title = EXPORT_SHEET_TITLE
    export.sheet_view.rightToLeft = True
    export_headers = ["Contract ID", "Clinet", "SubContractor", "Contract No",
                      "Contract Subject", "Sub Project", "Contract Date",
                      "Contract Part", "Excution Type", "Contract Categor",
                      "Contract Status"]
    for column, title in enumerate(export_headers, start=1):
        export.cell(row=4, column=column, value=title)
    for offset, title in enumerate(DOC_NAMES):
        export.cell(row=4, column=12 + offset, value=title)
    export.cell(row=4, column=35, value="جمع تعداد مدارک")
    export.cell(row=4, column=36, value="درصد پیشرفت مدارک اختتام")
    for offset, title in enumerate(DOC_NAMES):
        export.cell(row=4, column=38 + offset, value=title)
    export.cell(row=4, column=61, value="جمع کلی تعداد مدارک مورد نیاز")
    for offset, title in enumerate(DOC_NAMES):
        export.cell(row=4, column=63 + offset, value=title)
    export.cell(row=4, column=86, value="جمع کلی تعداد مدارک لازم به ثبت")

    dashboard = book.create_sheet(DEFAULT_SHEET)
    dashboard.sheet_view.rightToLeft = True
    titles = ["ردیف", "کد قرارداد (مالی)", "شناسه قرارداد (کندو)", "نام شرکت",
              "شماره قرارداد", "نوع قرارداد", "موضوع قرارداد", "مبلغ قرارداد",
              "واحد پولی", "مبلغ آخرین صورت وضعیت", "مبلغ ضمانت‌نامه"]
    for column, title in enumerate(titles, start=1):
        dashboard.cell(row=3, column=column, value=title)
    for offset, title in enumerate(DOC_NAMES):
        dashboard.cell(row=3, column=12 + offset, value=title)
    dashboard.cell(row=3, column=35, value="شرح وضعیت قرارداد")
    dashboard.cell(row=3, column=36, value="وضعیت قرارداد در سامانه کندو")
    dashboard.cell(row=3, column=37, value="درصد پیشرفت جدید")

    percents: List[float] = []
    for index, record in enumerate(source.all_records or source.records):
        row = 4 + index
        codes = requirements_for(record.contract_type, MATRIX)
        if len(codes) < len(DOC_NAMES):
            codes = [req_code(value) for value in record.values[26:49]]
        combos = combos_for(record.docs, codes) if len(codes) >= len(DOC_NAMES) else [""] * 23
        states = [doc_state(record.docs[offset], codes[offset] if offset < len(codes) else "",
                            combos[offset] if offset < len(combos) else "", record.percent)
                  for offset in range(len(DOC_NAMES))]

        # شیت خروجی
        export.cell(row=5 + index, column=1, value=record.contract_id)
        export.cell(row=5 + index, column=2, value=record.employer)
        export.cell(row=5 + index, column=3, value=record.company)
        export.cell(row=5 + index, column=4, value=record.contract_no)
        export.cell(row=5 + index, column=5, value=record.subject)
        export.cell(row=5 + index, column=6, value="اجرا - مهر")
        export.cell(row=5 + index, column=9, value=record.contract_type)
        for offset in range(23):
            export.cell(row=5 + index, column=12 + offset, value=record.docs[offset])
            if len(codes) >= len(DOC_NAMES):
                export.cell(row=5 + index, column=38 + offset, value=codes[offset])
                export.cell(row=5 + index, column=63 + offset, value=combos[offset])
        export.cell(row=5 + index, column=35,
                    value=sum(1 for value in record.docs if (num(value) or 0) > 0))
        export.cell(row=5 + index, column=36, value=record.percent)
        if len(codes) >= len(DOC_NAMES):
            export.cell(row=5 + index, column=61,
                        value=sum(1 for code in codes if code == "M"))
            export.cell(row=5 + index, column=86,
                        value=sum(1 for combo in combos if combo == "1-M"))

        # شیت colorcode
        dashboard.cell(row=row, column=1, value=index + 1)
        dashboard.cell(row=row, column=3, value=record.contract_id)
        dashboard.cell(row=row, column=4, value=record.company)
        dashboard.cell(row=row, column=5, value=record.contract_no)
        dashboard.cell(row=row, column=6, value=record.contract_type)
        dashboard.cell(row=row, column=7, value=record.subject)
        dashboard.cell(row=row, column=36, value="اجرا - مهر")
        for offset, state in enumerate(states):
            dashboard.cell(row=row, column=12 + offset, value=state)
        dashboard.cell(row=row, column=37, value=record.percent)
        if record.percent is not None and record.percent > 0:
            percents.append(record.percent)
    if percents:
        dashboard.cell(row=2, column=37, value=sum(percents) / len(percents))
    buffer = io.BytesIO()
    book.save(buffer)
    log_line("کتاب کار پایه از فایل ۰۱ ساخته شد: {} قرارداد.".format(
        fa(len(source.all_records or source.records))))
    return buffer.getvalue()


# ===========================================================================
# ۸) دیتابیس داخلی (SQLite) — همه‌چیز در D:\DATANEW
# ===========================================================================
SCHEMA = """
CREATE TABLE IF NOT EXISTS workbooks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    sheet TEXT NOT NULL,
    sha256 TEXT NOT NULL,
    size INTEGER NOT NULL,
    rows INTEGER NOT NULL,
    kind TEXT NOT NULL DEFAULT 'excel',
    imported_at TEXT NOT NULL,
    blob BLOB NOT NULL
);
CREATE TABLE IF NOT EXISTS contracts (
    workbook_id INTEGER NOT NULL,
    excel_row INTEGER NOT NULL,
    contract_no TEXT DEFAULT '',
    contract_id TEXT DEFAULT '',
    company TEXT DEFAULT '',
    subject TEXT DEFAULT '',
    contract_type TEXT DEFAULT '',
    employer TEXT DEFAULT '',
    percent REAL,
    status TEXT DEFAULT '',
    docs TEXT DEFAULT '[]',
    counts TEXT DEFAULT '[]',
    flags INTEGER DEFAULT 0,
    problem TEXT DEFAULT '',
    PRIMARY KEY (workbook_id, excel_row)
);
CREATE TABLE IF NOT EXISTS notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contract_key TEXT NOT NULL,
    author TEXT DEFAULT '',
    text TEXT DEFAULT '',
    created TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS bookmarks (
    contract_key TEXT PRIMARY KEY,
    created TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
CREATE INDEX IF NOT EXISTS ix_contracts_no ON contracts(contract_no);
CREATE INDEX IF NOT EXISTS ix_contracts_company ON contracts(company);
CREATE INDEX IF NOT EXISTS ix_notes_key ON notes(contract_key);
"""


def connect_db() -> sqlite3.Connection:
    """اتصال به دیتابیس داخلی؛ جدول‌ها بار اول ساخته می‌شوند."""
    path = database_path()
    connection = sqlite3.connect(str(path), timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.executescript(SCHEMA)
    return connection


def setting_get(key: str, default: str = "") -> str:
    try:
        with connect_db() as connection:
            row = connection.execute("SELECT value FROM settings WHERE key=?",
                                     (key,)).fetchone()
        return str(row["value"]) if row else default
    except Exception:
        return default


def setting_set(key: str, value: Any) -> None:
    try:
        with connect_db() as connection:
            connection.execute(
                "INSERT INTO settings(key, value) VALUES(?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, str(value)),
            )
            connection.commit()
    except Exception as exc:
        log_line(f"ذخیرهٔ تنظیم «{key}» ممکن نشد: {exc}")


def _flags_mask(flags: Dict[str, bool]) -> int:
    mask = 0
    for index, (key, _label, _column) in enumerate(FLAG_DEFS):
        if flags.get(key):
            mask |= 1 << index
    return mask


def _mask_flags(mask: int) -> Dict[str, bool]:
    return {key: bool(int(mask or 0) & (1 << index))
            for index, (key, _label, _column) in enumerate(FLAG_DEFS)}


def safe_filename(name: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*]', "-", str(name or "workbook.xlsx")).strip()
    return clean[:120] or "workbook.xlsx"


def save_workbook(blob: bytes, dashboard: Dashboard, kind: str = "excel") -> int:
    """کتاب کار را در دیتابیس ذخیره می‌کند و شناسهٔ آن را برمی‌گرداند."""
    digest = hashlib.sha256(blob).hexdigest()
    filename = safe_filename(dashboard.source_name or "workbook.xlsx")
    with connect_db() as connection:
        cursor = connection.execute(
            "INSERT INTO workbooks(filename, sheet, sha256, size, rows, kind, "
            "imported_at, blob) VALUES(?,?,?,?,?,?,?,?)",
            (filename, dashboard.sheet_name, digest, len(blob), dashboard.total,
             kind, datetime.now().isoformat(timespec="seconds"), sqlite3.Binary(blob)),
        )
        workbook_id = int(cursor.lastrowid)
        rows = []
        for item in dashboard.contracts:
            rows.append((
                workbook_id, item.excel_row, item.contract_no, item.contract_id,
                item.company, item.subject, item.contract_type, item.employer,
                item.percent, item.status,
                json.dumps(item.docs, ensure_ascii=False),
                json.dumps([txt(value) for value in item.counts], ensure_ascii=False),
                _flags_mask(item.flags),
                item.problem,
            ))
        connection.executemany(
            "INSERT OR REPLACE INTO contracts(workbook_id, excel_row, contract_no, "
            "contract_id, company, subject, contract_type, employer, percent, status, "
            "docs, counts, flags, problem) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            rows,
        )
        connection.commit()
    try:
        folder = data_subdir("workbooks")
        (folder / f"{workbook_id:04d}-{filename}").write_bytes(blob)
    except Exception as exc:
        log_line(f"نوشتن نسخهٔ کتاب کار ممکن نشد: {exc}")
    setting_set("current_workbook", workbook_id)
    log_line(f"کتاب کار در دیتابیس ذخیره شد: شناسهٔ {fa(workbook_id)} — "
             f"{fa(dashboard.total)} قرارداد ({kind})")
    return workbook_id


def load_workbook_blob(workbook_id: int) -> bytes:
    with connect_db() as connection:
        row = connection.execute("SELECT blob FROM workbooks WHERE id=?",
                                 (int(workbook_id),)).fetchone()
    if not row:
        raise ValueError(f"کتاب کار با شناسهٔ {fa(workbook_id)} در دیتابیس نیست.")
    return bytes(row["blob"])


def load_dashboard(workbook_id: int = 0) -> Optional[Dashboard]:
    """داشبورد را از ردیف‌های ذخیره‌شده می‌سازد."""
    with connect_db() as connection:
        if not workbook_id:
            row = connection.execute(
                "SELECT id FROM workbooks ORDER BY id DESC LIMIT 1").fetchone()
            workbook_id = int(row["id"]) if row else 0
        if not workbook_id:
            return None
        meta = connection.execute(
            "SELECT * FROM workbooks WHERE id=?", (workbook_id,)).fetchone()
        if not meta:
            return None
        rows = connection.execute(
            "SELECT * FROM contracts WHERE workbook_id=? ORDER BY excel_row",
            (workbook_id,)).fetchall()
    if not rows:
        return None
    contracts: List[Contract] = []
    percents: List[float] = []
    for row in rows:
        try:
            docs = json.loads(row["docs"] or "[]")
        except Exception:
            docs = []
        try:
            counts = json.loads(row["counts"] or "[]")
        except Exception:
            counts = []
        percent = row["percent"]
        contracts.append(Contract(
            excel_row=int(row["excel_row"]),
            contract_no=str(row["contract_no"] or ""),
            contract_id=str(row["contract_id"] or ""),
            company=str(row["company"] or ""),
            subject=str(row["subject"] or ""),
            contract_type=str(row["contract_type"] or ""),
            employer=str(row["employer"] or ""),
            percent=(float(percent) if percent is not None else None),
            status=str(row["status"] or ""),
            problem=str(row["problem"] or ""),
            docs=[str(value or "") for value in docs],
            counts=list(counts),
            flags=_mask_flags(int(row["flags"] or 0)),
        ))
        if percent:
            percents.append(float(percent))
    progress = (sum(percents) / len(percents)) if percents else None
    return Dashboard(
        contracts=contracts,
        source_name=str(meta["filename"]),
        sheet_name=str(meta["sheet"]),
        progress=progress,
        percent=(round(progress * 100, 2) if progress is not None else None),
        loaded_at=datetime.now(),
        workbook_id=int(meta["id"]),
        note=f"دیتابیس داخلی — شناسهٔ {fa(int(meta['id']))}",
        doc_titles=list(DOC_NAMES),
    )


def list_workbooks() -> List[Dict[str, Any]]:
    try:
        with connect_db() as connection:
            rows = connection.execute(
                "SELECT id, filename, sheet, rows, size, kind, imported_at "
                "FROM workbooks ORDER BY id DESC").fetchall()
    except Exception:
        return []
    return [dict(row) for row in rows]


def database_summary() -> Dict[str, Any]:
    summary: Dict[str, Any] = {"workbooks": 0, "contracts": 0, "notes": 0,
                               "size": 0, "path": str(database_path()),
                               "current": 0}
    try:
        with connect_db() as connection:
            summary["workbooks"] = int(connection.execute(
                "SELECT COUNT(*) FROM workbooks").fetchone()[0])
            summary["contracts"] = int(connection.execute(
                "SELECT COUNT(*) FROM contracts").fetchone()[0])
            summary["notes"] = int(connection.execute(
                "SELECT COUNT(*) FROM notes").fetchone()[0])
    except Exception:
        pass
    try:
        summary["size"] = database_path().stat().st_size
    except Exception:
        pass
    try:
        summary["current"] = int(setting_get("current_workbook", "0") or 0)
    except Exception:
        pass
    return summary


def reset_database(keep_backup: bool = True) -> Path:
    """دیتابیس داخلی را از صفر می‌سازد (روی فایل‌های اکسل کاری ندارد)."""
    path = database_path()
    backup_note = ""
    if keep_backup and path.exists():
        folder = data_subdir("backup")
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        target = folder / f"mehr-{stamp}.db"
        try:
            shutil.copy2(path, target)
            backup_note = f" (پشتیبان: {target.name})"
        except Exception as exc:
            log_line(f"پشتیبان دیتابیس ساخته نشد: {exc}")
    for suffix in ("", "-wal", "-shm"):
        candidate = Path(str(path) + suffix)
        try:
            if candidate.exists():
                candidate.unlink()
        except Exception as exc:
            log_line(f"حذف {candidate.name} ممکن نشد: {exc}")
    with connect_db():
        pass
    log_line("دیتابیس داخلی از صفر ساخته شد" + backup_note)
    return path


def add_note(contract_key: str, author: str, text: str) -> None:
    if not str(text or "").strip():
        return
    with connect_db() as connection:
        connection.execute(
            "INSERT INTO notes(contract_key, author, text, created) VALUES(?,?,?,?)",
            (str(contract_key), str(author or ""), str(text).strip(),
             datetime.now().isoformat(timespec="seconds")),
        )
        connection.commit()


def notes_for(contract_key: str) -> List[Dict[str, Any]]:
    with connect_db() as connection:
        rows = connection.execute(
            "SELECT * FROM notes WHERE contract_key=? ORDER BY id DESC",
            (str(contract_key),)).fetchall()
    return [dict(row) for row in rows]


def toggle_bookmark(contract_key: str) -> bool:
    with connect_db() as connection:
        row = connection.execute("SELECT 1 FROM bookmarks WHERE contract_key=?",
                                 (str(contract_key),)).fetchone()
        if row:
            connection.execute("DELETE FROM bookmarks WHERE contract_key=?",
                               (str(contract_key),))
            state = False
        else:
            connection.execute(
                "INSERT INTO bookmarks(contract_key, created) VALUES(?,?)",
                (str(contract_key), datetime.now().isoformat(timespec="seconds")))
            state = True
        connection.commit()
    return state


def bookmarked_keys() -> set:
    try:
        with connect_db() as connection:
            rows = connection.execute("SELECT contract_key FROM bookmarks").fetchall()
        return {str(row[0]) for row in rows}
    except Exception:
        return set()


# ===========================================================================
# ۹) فایل‌های PDF قراردادها
# ===========================================================================
def pdf_index(folder: str, depth: int = 2, limit: int = 20000) -> Dict[str, str]:
    """فهرست PDF‌ها را بر اساس کلید قرارداد می‌سازد."""
    index: Dict[str, str] = {}
    base = Path(str(folder or "")).expanduser()
    if not base.exists() or not base.is_dir():
        return index
    seen = 0
    for path in base.rglob("*.pdf"):
        seen += 1
        if seen > limit:
            break
        stem = path.stem
        candidates = {key_of(stem)}
        for separator in ("__", "_", "-", " "):
            if separator in stem:
                candidates.add(key_of(stem.split(separator)[0]))
        for part in re.split(r"[_\-\s]+", stem):
            if len(part) >= 4:
                candidates.add(key_of(part))
        for candidate in candidates:
            if candidate and candidate not in index:
                index[candidate] = str(path)
    return index


def find_pdf(record: Contract, index: Dict[str, str]) -> str:
    for candidate in (record.contract_no, record.contract_id):
        key = key_of(candidate)
        if key and key in index:
            return index[key]
    return ""


# ===========================================================================
# ۱۰) خروجی اکسل و گزارش
# ===========================================================================
CONTRACT_COLUMNS = [
    ("ردیف", 6), ("شمارهٔ قرارداد", 22), ("شناسهٔ قرارداد", 14), ("نام شرکت", 32),
    ("نوع قرارداد", 12), ("موضوع قرارداد", 46), ("کارفرما", 26),
    ("درصد پیشرفت", 12), ("وضعیت", 22), ("اسناد باز", 40), ("شرح مشکل / کسورات", 42),
]


def export_contracts(path: str, contracts: Sequence[Contract],
                     pdfs: Optional[Dict[str, str]] = None) -> int:
    """خروجی اکسل قراردادها."""
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست.")
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "قراردادها"
    sheet.sheet_view.rightToLeft = True
    from openpyxl.styles import Alignment, Font, PatternFill
    header_fill = PatternFill("solid", fgColor="5A4FE0")
    for column, (title, width) in enumerate(CONTRACT_COLUMNS, start=1):
        cell = sheet.cell(row=1, column=column, value=title)
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF", name="Tahoma", size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.column_dimensions[cell.column_letter].width = width
    for index, item in enumerate(contracts, start=2):
        values = [
            index - 1, item.contract_no, item.contract_id, item.company,
            item.contract_type, item.subject, item.employer,
            (round((item.percent or 0) * 100, 2) if item.percent is not None else None),
            item.status, "، ".join(item.open_docs), item.problem,
        ]
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=index, column=column, value=value)
            cell.font = Font(name="Tahoma", size=9)
            cell.alignment = Alignment(vertical="center", wrap_text=column in (6, 10, 11))
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:K{max(1, len(contracts) + 1)}"
    book.save(str(path))
    log_line(f"خروجی اکسل قراردادها: {fa(len(contracts))} ردیف → {Path(path).name}")
    return len(contracts)


def export_report(path: str, dashboard: Dashboard) -> int:
    """گزارش مدیریتی سه‌شیتی."""
    if openpyxl is None:
        raise RuntimeError("کتابخانهٔ openpyxl نصب نیست.")
    from openpyxl.styles import Alignment, Font, PatternFill
    book = openpyxl.Workbook()
    summary = book.active
    summary.title = "خلاصه"
    summary.sheet_view.rightToLeft = True
    title_fill = PatternFill("solid", fgColor="5A4FE0")
    summary["A1"] = "گزارش مدیریتی قراردادهای مهر خلیج فارس"
    summary["A1"].font = Font(bold=True, size=13, name="Tahoma")
    summary.merge_cells("A1:D1")
    rows = [
        ("تعداد قرارداد", dashboard.total),
        ("درصد پیشرفت کل", dashboard.percent),
        ("قراردادهای ناقص (دارای مدرک باز)", dashboard.incomplete),
        ("قراردادهای دارای شرح مشکل", dashboard.with_deductions),
        ("اسناد باز (M)", dashboard.open_documents),
        ("کتاب کار", dashboard.source_name),
        ("زمان گزارش", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    for index, (label, value) in enumerate(rows, start=3):
        summary.cell(row=index, column=1, value=label).font = Font(bold=True, name="Tahoma")
        cell = summary.cell(row=index, column=2, value=value)
        cell.font = Font(name="Tahoma")
    summary.column_dimensions["A"].width = 30
    summary.column_dimensions["B"].width = 40

    flags = book.create_sheet("وضعیت‌ها")
    flags.sheet_view.rightToLeft = True
    for column, title in enumerate(("وضعیت", "تعداد"), start=1):
        cell = flags.cell(row=1, column=column, value=title)
        cell.fill = title_fill
        cell.font = Font(bold=True, color="FFFFFF", name="Tahoma")
    counts = dashboard.flag_counts()
    for index, (key, label, _column) in enumerate(FLAG_DEFS, start=2):
        flags.cell(row=index, column=1, value=label).font = Font(name="Tahoma")
        flags.cell(row=index, column=2, value=counts.get(key, 0)).font = Font(name="Tahoma")
    start = len(FLAG_DEFS) + 3
    flags.cell(row=start, column=1, value="نوع قرارداد").font = Font(bold=True, name="Tahoma")
    flags.cell(row=start, column=2, value="تعداد").font = Font(bold=True, name="Tahoma")
    for index, (label, count) in enumerate(dashboard.type_counts(), start=start + 1):
        flags.cell(row=index, column=1, value=label).font = Font(name="Tahoma")
        flags.cell(row=index, column=2, value=count).font = Font(name="Tahoma")
    flags.column_dimensions["A"].width = 34
    flags.column_dimensions["B"].width = 12

    pending = book.create_sheet("اسناد باز")
    pending.sheet_view.rightToLeft = True
    for column, title in enumerate(("مدرک", "تعداد قرارداد"), start=1):
        cell = pending.cell(row=1, column=column, value=title)
        cell.fill = title_fill
        cell.font = Font(bold=True, color="FFFFFF", name="Tahoma")
    for index, (label, count) in enumerate(dashboard.pending_docs(), start=2):
        pending.cell(row=index, column=1, value=label).font = Font(name="Tahoma")
        pending.cell(row=index, column=2, value=count).font = Font(name="Tahoma")
    pending.column_dimensions["A"].width = 52
    pending.column_dimensions["B"].width = 14

    contracts = book.create_sheet("قراردادها")
    contracts.sheet_view.rightToLeft = True
    for column, (title, width) in enumerate(CONTRACT_COLUMNS, start=1):
        cell = contracts.cell(row=1, column=column, value=title)
        cell.fill = title_fill
        cell.font = Font(bold=True, color="FFFFFF", name="Tahoma", size=9)
        contracts.column_dimensions[cell.column_letter].width = width
    for index, item in enumerate(dashboard.contracts, start=2):
        values = [index - 1, item.contract_no, item.contract_id, item.company,
                  item.contract_type, item.subject, item.employer,
                  (round((item.percent or 0) * 100, 2) if item.percent is not None else None),
                  item.status, "، ".join(item.open_docs)]
        for column, value in enumerate(values, start=1):
            cell = contracts.cell(row=index, column=column, value=value)
            cell.font = Font(name="Tahoma", size=9)
            cell.alignment = Alignment(wrap_text=column in (6, 10), vertical="center")
    contracts.freeze_panes = "A2"
    book.save(str(path))
    log_line(f"گزارش مدیریتی ساخته شد: {Path(path).name}")
    return dashboard.total


# ===========================================================================
# ۱۱) رابط کاربری (PyQt5)
# ===========================================================================
if QT_READY:

    THEMES = {
        "light": {
            "bg": "#F1F2F9", "surface": "#FFFFFF", "surface2": "#F7F7FC",
            "ink": "#171B26", "ink2": "#4E5566", "ink3": "#858B9C",
            "brand": "#5A4FE0", "brand_soft": "#ECEAFF", "border": "#E2E4EF",
            "ok": "#0E9F6E", "warn": "#C98A00", "danger": "#D64545",
            "shadow": "#D8DAE8",
        },
        "dark": {
            "bg": "#12141C", "surface": "#1B1E29", "surface2": "#222634",
            "ink": "#F2F3F8", "ink2": "#B9BFD0", "ink3": "#848CA0",
            "brand": "#8C83FF", "brand_soft": "#2A2947", "border": "#2C3040",
            "ok": "#3DD598", "warn": "#F2C14E", "danger": "#FF7B7B",
            "shadow": "#0B0D13",
        },
    }

    def make_font(size: int = 10, bold: bool = False) -> QFont:
        font = QFont("Tahoma")
        font.setPointSize(size)
        font.setBold(bold)
        return font

    class Card(QFrame):
        """کارت سفید گرد با عنوان."""

        def __init__(self, title: str = "", subtitle: str = "",
                     colors: Optional[Dict[str, str]] = None, parent=None) -> None:
            super().__init__(parent)
            self.colors = colors or THEMES["light"]
            self.setObjectName("Card")
            self.body = QVBoxLayout(self)
            self.body.setContentsMargins(16, 14, 16, 14)
            self.body.setSpacing(8)
            if title:
                row = QHBoxLayout()
                row.setSpacing(6)
                self.title_label = QLabel(title)
                self.title_label.setFont(make_font(10, True))
                row.addWidget(self.title_label)
                row.addStretch(1)
                self.subtitle_label = QLabel(subtitle)
                self.subtitle_label.setFont(make_font(8))
                row.addWidget(self.subtitle_label)
                self.body.addLayout(row)
            self.apply_colors(self.colors)

        def apply_colors(self, colors: Dict[str, str]) -> None:
            self.colors = colors
            self.setStyleSheet(
                f"QFrame#Card {{ background: {colors['surface']};"
                f" border: 1px solid {colors['border']}; border-radius: 14px; }}"
                f" QLabel {{ color: {colors['ink']}; background: transparent; }}"
            )
            if hasattr(self, "subtitle_label"):
                self.subtitle_label.setStyleSheet(f"color: {colors['ink3']};")

    class StatCard(Card):
        """کارت آمار: عنوان، عدد بزرگ، توضیح کوچک."""

        def __init__(self, label: str, colors: Dict[str, str], accent: str = "",
                     parent=None) -> None:
            super().__init__("", "", colors, parent)
            self.body.setSpacing(2)
            self.caption = QLabel(label)
            self.caption.setFont(make_font(8))
            self.value = QLabel("—")
            self.value.setFont(make_font(17, True))
            self.hint = QLabel("")
            self.hint.setFont(make_font(7))
            self.body.addWidget(self.caption)
            self.body.addWidget(self.value)
            self.body.addWidget(self.hint)
            self.accent = accent or colors["brand"]
            self.apply_colors(colors)

        def apply_colors(self, colors: Dict[str, str]) -> None:
            super().apply_colors(colors)
            self.accent = getattr(self, "accent", colors["brand"]) or colors["brand"]
            if not hasattr(self, "caption"):
                return
            self.caption.setStyleSheet(f"color: {colors['ink3']};")
            self.value.setStyleSheet(f"color: {self.accent};")
            self.hint.setStyleSheet(f"color: {colors['ink3']};")

        def set_value(self, value: str, hint: str = "") -> None:
            self.value.setText(value)
            self.hint.setText(hint)

    class BarList(QWidget):
        """نمودار میله‌ای افقی ساده با برچسب فارسی."""

        def __init__(self, colors: Dict[str, str], parent=None) -> None:
            super().__init__(parent)
            self.colors = colors
            self.items: List[Tuple[str, int]] = []
            self.setMinimumHeight(120)
            self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        def set_items(self, items: Sequence[Tuple[str, int]]) -> None:
            self.items = list(items)
            self.update()

        def apply_colors(self, colors: Dict[str, str]) -> None:
            self.colors = colors
            self.update()

        def paintEvent(self, event) -> None:  # noqa: N802
            painter = QPainter(self)
            painter.setRenderHint(QPainter.Antialiasing, True)
            colors = self.colors
            if not self.items:
                painter.setPen(QPen(QColor(colors["ink3"])))
                painter.drawText(self.rect(), Qt.AlignCenter, "داده‌ای نیست")
                return
            top = 6
            row_height = max(22, int((self.height() - 12) / max(1, len(self.items))))
            label_width = 118
            maximum = max(1, max(value for _label, value in self.items))
            for index, (label, value) in enumerate(self.items):
                y = top + index * row_height
                painter.setPen(QPen(QColor(colors["ink2"])))
                painter.setFont(make_font(8))
                painter.drawText(6, y, label_width, row_height - 6,
                                 int(Qt.AlignRight | Qt.AlignVCenter), label[:38])
                bar_left = label_width + 10
                bar_width = max(2, self.width() - bar_left - 56)
                painter.setPen(Qt.NoPen)
                painter.setBrush(QBrush(QColor(colors["surface2"])))
                painter.drawRoundedRect(bar_left, y + 4, bar_width, row_height - 12, 5, 5)
                filled = int(bar_width * (value / maximum))
                painter.setBrush(QBrush(QColor(colors["brand"])))
                painter.drawRoundedRect(bar_left, y + 4, max(4, filled),
                                        row_height - 12, 5, 5)
                painter.setPen(QPen(QColor(colors["ink2"])))
                painter.drawText(bar_left + bar_width + 6, y, 46, row_height - 6,
                                 int(Qt.AlignLeft | Qt.AlignVCenter), fa(value))

    class DashboardPage(QWidget):
        """صفحهٔ داشبورد: درصد کل، کارت‌ها و توزیع‌ها."""

        def __init__(self, colors: Dict[str, str], parent=None) -> None:
            super().__init__(parent)
            self.colors = colors
            outer = QVBoxLayout(self)
            outer.setContentsMargins(0, 0, 0, 0)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            outer.addWidget(scroll)
            content = QWidget()
            scroll.setWidget(content)
            layout = QVBoxLayout(content)
            layout.setContentsMargins(18, 16, 18, 18)
            layout.setSpacing(12)

            self.hero = Card("", "", colors)
            hero_layout = QHBoxLayout()
            hero_layout.setSpacing(18)
            left = QVBoxLayout()
            self.hero_title = QLabel("درصد پیشرفت کل قراردادها")
            self.hero_title.setFont(make_font(10))
            self.hero_value = QLabel("—")
            self.hero_value.setFont(make_font(34, True))
            self.hero_note = QLabel("")
            self.hero_note.setFont(make_font(8))
            left.addWidget(self.hero_title)
            left.addWidget(self.hero_value)
            left.addWidget(self.hero_note)
            hero_layout.addLayout(left, 1)
            right = QVBoxLayout()
            self.progress = QProgressBar()
            self.progress.setRange(0, 1000)
            self.progress.setTextVisible(False)
            self.progress.setFixedHeight(14)
            self.hero_count = QLabel("")
            self.hero_count.setFont(make_font(9, True))
            self.hero_source = QLabel("")
            self.hero_source.setFont(make_font(8))
            right.addWidget(self.hero_count)
            right.addWidget(self.progress)
            right.addWidget(self.hero_source)
            hero_layout.addLayout(right, 1)
            self.hero.body.addLayout(hero_layout)
            layout.addWidget(self.hero)

            kpi_row = QHBoxLayout()
            kpi_row.setSpacing(10)
            self.kpi_total = StatCard("کل قراردادها", colors)
            self.kpi_deductions = StatCard("قراردادهای ناقص", colors, colors["warn"])
            self.kpi_open = StatCard("اسناد باز (M)", colors, colors["danger"])
            self.kpi_percent = StatCard("درصد پیشرفت", colors, colors["ok"])
            for card in (self.kpi_total, self.kpi_deductions, self.kpi_open, self.kpi_percent):
                kpi_row.addWidget(card, 1)
            layout.addLayout(kpi_row)

            status_card = Card("وضعیت اسناد", "تعداد قراردادهای علامت‌دار", colors)
            grid = QGridLayout()
            grid.setSpacing(8)
            self.flag_cards: Dict[str, StatCard] = {}
            for index, (key, label, _column) in enumerate(FLAG_DEFS):
                card = StatCard(label, colors)
                self.flag_cards[key] = card
                grid.addWidget(card, index // 4, index % 4)
            status_card.body.addLayout(grid)
            layout.addWidget(status_card)

            charts = QHBoxLayout()
            charts.setSpacing(10)
            self.type_card = Card("توزیع نوع قرارداد", "", colors)
            self.type_chart = BarList(colors)
            self.type_card.body.addWidget(self.type_chart)
            charts.addWidget(self.type_card, 1)
            self.pending_card = Card("اسناد باز (بیشترین)", "", colors)
            self.pending_chart = BarList(colors)
            self.pending_card.body.addWidget(self.pending_chart)
            charts.addWidget(self.pending_card, 1)
            layout.addLayout(charts, 1)

            self.type_status = Card("وضعیت قرارداد در سامانه", "", colors)
            self.status_chart = BarList(colors)
            self.type_status.body.addWidget(self.status_chart)
            layout.addWidget(self.type_status)
            layout.addStretch(1)
            self.apply_colors(colors)

        def apply_colors(self, colors: Dict[str, str]) -> None:
            self.colors = colors
            for widget in (self.hero, self.kpi_total, self.kpi_deductions, self.kpi_open,
                           self.kpi_percent, self.type_card, self.pending_card,
                           self.type_status):
                widget.apply_colors(colors)
            for card in self.flag_cards.values():
                card.apply_colors(colors)
            self.type_chart.apply_colors(colors)
            self.status_chart.apply_colors(colors)
            self.pending_chart.apply_colors(colors)
            self.hero_title.setStyleSheet(f"color: {colors['ink2']};")
            self.hero_value.setStyleSheet(f"color: {colors['brand']};")
            self.hero_note.setStyleSheet(f"color: {colors['ink3']};")
            self.hero_count.setStyleSheet(f"color: {colors['ink']};")
            self.hero_source.setStyleSheet(f"color: {colors['ink3']};")
            self.progress.setStyleSheet(
                f"QProgressBar {{ background: {colors['surface2']}; border-radius: 7px; }}"
                f" QProgressBar::chunk {{ background: {colors['brand']}; border-radius: 7px; }}"
            )

        def show_data(self, data: Optional[Dashboard]) -> None:
            if data is None or not data.contracts:
                self.hero_value.setText("—")
                self.hero_count.setText("هنوز داده‌ای بارگذاری نشده است")
                self.hero_source.setText("از دکمهٔ «بارگذاری کتاب کار» یا «به‌روزرسانی از فایل ۰۱» شروع کنید")
                self.hero_note.setText("")
                self.progress.setValue(0)
                for card in self.flag_cards.values():
                    card.set_value("—")
                for card in (self.kpi_total, self.kpi_deductions, self.kpi_open, self.kpi_percent):
                    card.set_value("—")
                self.type_chart.set_items([])
                self.pending_chart.set_items([])
                self.status_chart.set_items([])
                return
            percent = data.percent or 0.0
            self.hero_value.setText(f"{fa(round(percent, 2))}٪")
            self.progress.setValue(int(max(0, min(100, percent)) * 10))
            self.hero_count.setText(f"{fa(data.total)} قرارداد")
            self.hero_source.setText(
                f"{data.source_name} — {data.note}" if data.note else data.source_name
            )
            self.hero_note.setText(
                f"{fa(data.open_documents)} مدرک باز، "
                f"{fa(data.incomplete)} قرارداد ناقص"
            )
            self.kpi_total.set_value(fa(data.total), "قرارداد در کتاب کار")
            self.kpi_deductions.set_value(fa(data.incomplete),
                                          "حداقل یک مدرک باز (M)")
            self.kpi_open.set_value(fa(data.open_documents), "مدرک لازمِ ثبت‌نشده")
            self.kpi_percent.set_value(f"{fa(round(percent, 2))}٪", "میانگین ستون AK")
            counts = data.flag_counts()
            for key, _label, _column in FLAG_DEFS:
                self.flag_cards[key].set_value(fa(counts.get(key, 0)), "قرارداد")
            self.type_chart.set_items(data.type_counts()[:8])
            self.pending_chart.set_items(
                [(label[:30], count) for label, count in data.pending_docs()[:8]])
            self.status_chart.set_items(data.status_counts()[:8])

    class ContractsPage(QWidget):
        """فهرست قراردادها با جستجو، فیلتر و جزئیات."""

        def __init__(self, colors: Dict[str, str], notes_store=None, parent=None) -> None:
            super().__init__(parent)
            self.colors = colors
            self.data: Optional[Dashboard] = None
            self.pdfs: Dict[str, str] = {}
            self.bookmarks = bookmarked_keys()
            self.visible: List[Contract] = []
            self.current: Optional[Contract] = None
            layout = QVBoxLayout(self)
            layout.setContentsMargins(18, 16, 18, 16)
            layout.setSpacing(10)

            toolbar = QHBoxLayout()
            toolbar.setSpacing(8)
            self.search = QLineEdit()
            self.search.setPlaceholderText("جستجو در شمارهٔ قرارداد، شرکت، موضوع…")
            self.search.setClearButtonEnabled(True)
            self.search.textChanged.connect(lambda _text: self.apply_filter())
            toolbar.addWidget(self.search, 2)
            self.filter = QComboBox()
            self.filter.addItem("همهٔ قراردادها", "all")
            self.filter.addItem("دارای شرح مشکل", "deductions")
            self.filter.addItem("دارای مدرک باز (M)", "open")
            self.filter.addItem("نشان‌شده‌ها", "bookmarked")
            for key, label, _column in FLAG_DEFS:
                self.filter.addItem(label, key)
            self.filter.currentIndexChanged.connect(lambda _index: self.apply_filter())
            toolbar.addWidget(self.filter, 1)
            self.count_label = QLabel("")
            toolbar.addWidget(self.count_label)
            layout.addLayout(toolbar)

            splitter = QSplitter(Qt.Horizontal)
            self.table = QTableWidget(0, 7)
            self.table.setHorizontalHeaderLabels(
                ["ردیف", "شمارهٔ قرارداد", "نام شرکت", "نوع", "درصد", "وضعیت", "باز"]
            )
            self.table.verticalHeader().setVisible(False)
            self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.table.setSelectionMode(QAbstractItemView.SingleSelection)
            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.table.setAlternatingRowColors(True)
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.Stretch)
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
            self.table.itemSelectionChanged.connect(self.show_selected)
            splitter.addWidget(self.table)

            self.detail = Card("جزئیات قرارداد", "", colors)
            self.detail_title = QLabel("یک قرارداد را انتخاب کنید")
            self.detail_title.setFont(make_font(10, True))
            self.detail_title.setWordWrap(True)
            self.detail_meta = QLabel("")
            self.detail_meta.setFont(make_font(8))
            self.detail_meta.setWordWrap(True)
            self.detail_docs = QLabel("")
            self.detail_docs.setFont(make_font(8))
            self.detail_docs.setWordWrap(True)
            self.notes_label = QLabel("یادداشت‌ها")
            self.notes_label.setFont(make_font(9, True))
            self.notes_box = QPlainTextEdit()
            self.notes_box.setReadOnly(True)
            self.notes_box.setMaximumHeight(130)
            note_row = QHBoxLayout()
            self.note_input = QLineEdit()
            self.note_input.setPlaceholderText("یادداشت تازه…")
            self.note_add = QPushButton("افزودن")
            self.note_add.clicked.connect(self.add_note)
            note_row.addWidget(self.note_input, 1)
            note_row.addWidget(self.note_add)
            action_row = QHBoxLayout()
            self.bookmark_button = QPushButton("نشان کردن")
            self.bookmark_button.clicked.connect(self.toggle_bookmark)
            self.pdf_button = QPushButton("باز کردن PDF")
            self.pdf_button.clicked.connect(self.open_pdf)
            action_row.addWidget(self.bookmark_button)
            action_row.addWidget(self.pdf_button)
            action_row.addStretch(1)
            self.detail.body.addWidget(self.detail_title)
            self.detail.body.addWidget(self.detail_meta)
            self.detail.body.addWidget(self.detail_docs)
            self.detail.body.addWidget(self.notes_label)
            self.detail.body.addWidget(self.notes_box)
            self.detail.body.addLayout(note_row)
            self.detail.body.addLayout(action_row)
            self.detail.body.addStretch(1)
            splitter.addWidget(self.detail)
            splitter.setStretchFactor(0, 3)
            splitter.setStretchFactor(1, 2)
            layout.addWidget(splitter, 1)
            self.apply_colors(colors)

        def apply_colors(self, colors: Dict[str, str]) -> None:
            self.colors = colors
            self.detail.apply_colors(colors)
            self.setStyleSheet(
                f"QLineEdit, QComboBox, QPlainTextEdit {{ background: {colors['surface']};"
                f" color: {colors['ink']}; border: 1px solid {colors['border']};"
                f" border-radius: 8px; padding: 5px 8px; }}"
                f" QTableWidget {{ background: {colors['surface']}; color: {colors['ink']};"
                f" border: 1px solid {colors['border']}; border-radius: 10px;"
                f" gridline-color: {colors['border']}; alternate-background-color:"
                f" {colors['surface2']}; }}"
                f" QHeaderView::section {{ background: {colors['brand_soft']};"
                f" color: {colors['ink']}; border: none; padding: 6px; font-weight: bold; }}"
                f" QPushButton {{ background: {colors['surface2']}; color: {colors['ink']};"
                f" border: 1px solid {colors['border']}; border-radius: 8px; padding: 6px 12px; }}"
                f" QPushButton:hover {{ background: {colors['brand_soft']}; }}"
            )
            self.detail_meta.setStyleSheet(f"color: {colors['ink2']};")
            self.detail_docs.setStyleSheet(f"color: {colors['ink2']};")
            self.count_label.setStyleSheet(f"color: {colors['ink3']};")

        def set_data(self, data: Optional[Dashboard], pdf_folder: str = "") -> None:
            self.data = data
            self.bookmarks = bookmarked_keys()
            self.pdfs = pdf_index(pdf_folder) if pdf_folder else {}
            self.apply_filter()

        def apply_filter(self) -> None:
            text = norm(self.search.text())
            choice = str(self.filter.currentData() or "all")
            records = list(self.data.contracts) if self.data else []
            if choice == "deductions":
                records = [item for item in records if item.has_problem]
            elif choice == "open":
                records = [item for item in records if item.is_mark]
            elif choice == "bookmarked":
                records = [item for item in records if item.key in self.bookmarks]
            elif choice != "all":
                records = [item for item in records if item.flags.get(choice)]
            if text:
                records = [item for item in records
                           if text in norm(item.contract_no) or text in norm(item.company)
                           or text in norm(item.subject) or text in norm(item.contract_id)
                           or text in norm(item.contract_type)]
            self.visible = records
            self.table.setRowCount(len(records))
            for row, item in enumerate(records):
                values = [
                    fa(row + 1), item.contract_no or item.contract_id, item.company,
                    item.contract_type,
                    (f"{fa(round((item.percent or 0) * 100, 2))}٪"
                     if item.percent is not None else "—"),
                    item.status, fa(len(item.open_docs)),
                ]
                for column, value in enumerate(values):
                    cell = QTableWidgetItem(str(value))
                    cell.setFont(make_font(8))
                    if column == 6 and item.is_mark:
                        cell.setForeground(QBrush(QColor(self.colors["danger"])))
                    self.table.setItem(row, column, cell)
            self.count_label.setText(f"{fa(len(records))} از {fa(len(self.data.contracts) if self.data else 0)}")

        def show_selected(self) -> None:
            rows = self.table.selectionModel().selectedRows() if self.table.selectionModel() else []
            if not rows:
                return
            index = rows[0].row()
            if 0 <= index < len(self.visible):
                self.current = self.visible[index]
                self.render_detail(self.current)

        def render_detail(self, item: Contract) -> None:
            self.detail_title.setText(
                f"{item.company or '—'} — {item.contract_no or item.contract_id or '—'}"
            )
            lines = [
                f"شناسهٔ کندو: {item.contract_id or '—'}    نوع: {item.contract_type or '—'}",
                "درصد پیشرفت: "
                + (f"{fa(round((item.percent or 0) * 100, 2))}٪" if item.percent is not None else "—")
                + f"    وضعیت: {item.status or '—'}",
                f"کارفرما: {item.employer or '—'}",
                f"موضوع: {item.subject or '—'}",
            ]
            if item.has_problem:
                lines.append(f"شرح مشکل / کسورات: {item.problem[:400]}")
            self.detail_meta.setText("\n".join(lines))
            if item.docs:
                chips = []
                for index, state in enumerate(item.docs):
                    name = DOC_NAMES[index] if index < len(DOC_NAMES) else f"مدرک {index + 1}"
                    chips.append(f"{'⚠️ ' if state == 'M' else '✅ '}{name}")
                self.detail_docs.setText("\n".join(chips))
            else:
                self.detail_docs.setText("")
            notes = notes_for(item.key)
            self.notes_box.setPlainText("\n".join(
                f"[{note.get('created', '')}] {note.get('author', '')}: {note.get('text', '')}"
                for note in notes) or "یادداشتی ثبت نشده است.")
            marked = item.key in self.bookmarks
            self.bookmark_button.setText("حذف نشان" if marked else "نشان کردن")

        def add_note(self) -> None:
            if not self.current:
                return
            text = self.note_input.text().strip()
            if not text:
                return
            add_note(self.current.key, setting_get("author", "کاربر"), text)
            self.note_input.clear()
            self.render_detail(self.current)

        def toggle_bookmark(self) -> None:
            if not self.current:
                return
            toggle_bookmark(self.current.key)
            self.bookmarks = bookmarked_keys()
            self.render_detail(self.current)

        def open_pdf(self) -> None:
            if not self.current:
                return
            path = find_pdf(self.current, self.pdfs)
            if not path:
                QMessageBox.information(self, "PDF", "برای این قرارداد فایل PDF پیدا نشد.")
                return
            from PyQt5.QtGui import QDesktopServices
            from PyQt5.QtCore import QUrl
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    class ReportPage(QWidget):
        """گزارش خلاصه + خروجی‌ها."""

        def __init__(self, colors: Dict[str, str], parent=None) -> None:
            super().__init__(parent)
            self.colors = colors
            self.data: Optional[Dashboard] = None
            layout = QVBoxLayout(self)
            layout.setContentsMargins(18, 16, 18, 16)
            layout.setSpacing(10)
            self.summary = Card("گزارش مدیریتی", "خلاصهٔ وضعیت قراردادها", colors)
            self.summary_text = QLabel("داده‌ای بارگذاری نشده است.")
            self.summary_text.setFont(make_font(9))
            self.summary_text.setWordWrap(True)
            self.summary.body.addWidget(self.summary_text)
            self.actions = Card("خروجی‌ها", "", colors)
            row = QHBoxLayout()
            self.export_button = QPushButton("خروجی اکسل قراردادها")
            self.report_button = QPushButton("گزارش مدیریتی (چهار شیت)")
            self.print_button = QPushButton("چاپ / ذخیرهٔ PDF")
            row.addWidget(self.export_button)
            row.addWidget(self.report_button)
            row.addWidget(self.print_button)
            row.addStretch(1)
            self.actions.body.addLayout(row)
            self.note = QLabel("فایل‌ها در پوشهٔ exports داخل D:\\DATANEW ذخیره می‌شوند.")
            self.note.setFont(make_font(8))
            self.actions.body.addWidget(self.note)
            layout.addWidget(self.summary)
            layout.addWidget(self.actions)
            layout.addStretch(1)
            self.apply_colors(colors)

        def apply_colors(self, colors: Dict[str, str]) -> None:
            self.colors = colors
            self.summary.apply_colors(colors)
            self.actions.apply_colors(colors)
            self.setStyleSheet(
                f"QPushButton {{ background: {colors['surface2']}; color: {colors['ink']};"
                f" border: 1px solid {colors['border']}; border-radius: 9px;"
                f" padding: 8px 14px; }} QPushButton:hover {{ background: {colors['brand_soft']}; }}"
            )
            self.note.setStyleSheet(f"color: {colors['ink3']};")

        def show_data(self, data: Optional[Dashboard]) -> None:
            self.data = data
            if data is None or not data.contracts:
                self.summary_text.setText("داده‌ای بارگذاری نشده است.")
                return
            counts = data.flag_counts()
            lines = [
                f"کتاب کار: {data.source_name} — {fa(data.total)} قرارداد",
                f"درصد پیشرفت کل: {fa(round(data.percent or 0, 2))}٪",
                f"قراردادهای ناقص (دارای مدرک باز): {fa(data.incomplete)}",
                f"اسناد باز (M): {fa(data.open_documents)}",
                "",
                "وضعیت‌ها: " + "، ".join(
                    f"{label}={fa(counts.get(key, 0))}"
                    for key, label, _column in FLAG_DEFS),
                "",
                "نوع قرارداد: " + "، ".join(f"{label}={fa(count)}"
                                            for label, count in data.type_counts()[:6]),
            ]
            self.summary_text.setText("\n".join(lines))

    class SettingsPage(QWidget):
        """تنظیمات، دیتابیس و پوشهٔ داده."""

        def __init__(self, colors: Dict[str, str], parent=None) -> None:
            super().__init__(parent)
            self.colors = colors
            outer = QVBoxLayout(self)
            outer.setContentsMargins(0, 0, 0, 0)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            outer.addWidget(scroll)
            content = QWidget()
            scroll.setWidget(content)
            layout = QVBoxLayout(content)
            layout.setContentsMargins(18, 16, 18, 18)
            layout.setSpacing(10)

            self.path_card = Card("محل داده‌ها", "همه‌چیز در D:\\DATANEW", colors)
            self.path_label = QLabel("")
            self.path_label.setFont(make_font(9))
            self.path_label.setWordWrap(True)
            self.path_card.body.addWidget(self.path_label)
            path_row = QHBoxLayout()
            self.open_folder = QPushButton("باز کردن پوشه")
            self.reset_button = QPushButton("ساخت دیتابیس از صفر")
            path_row.addWidget(self.open_folder)
            path_row.addWidget(self.reset_button)
            path_row.addStretch(1)
            self.path_card.body.addLayout(path_row)

            self.db_card = Card("دیتابیس داخلی", "SQLite — بدون نیاز به MySQL", colors)
            self.db_label = QLabel("")
            self.db_label.setFont(make_font(9))
            self.db_label.setWordWrap(True)
            self.db_card.body.addWidget(self.db_label)
            self.workbooks_table = QTableWidget(0, 4)
            self.workbooks_table.setHorizontalHeaderLabels(
                ["شناسه", "نام کتاب کار", "قرارداد", "زمان"])
            self.workbooks_table.verticalHeader().setVisible(False)
            self.workbooks_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.workbooks_table.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.workbooks_table.setMaximumHeight(180)
            head = self.workbooks_table.horizontalHeader()
            head.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            head.setSectionResizeMode(1, QHeaderView.Stretch)
            head.setSectionResizeMode(2, QHeaderView.ResizeToContents)
            head.setSectionResizeMode(3, QHeaderView.ResizeToContents)
            self.db_card.body.addWidget(self.workbooks_table)
            row = QHBoxLayout()
            self.import_button = QPushButton("بارگذاری کتاب کار…")
            self.load_selected = QPushButton("نمایش کتاب کار انتخاب‌شده")
            row.addWidget(self.import_button)
            row.addWidget(self.load_selected)
            row.addStretch(1)
            self.db_card.body.addLayout(row)

            self.pref_card = Card("تنظیمات نمایش", "", colors)
            pref = QGridLayout()
            pref.setHorizontalSpacing(10)
            pref.setVerticalSpacing(8)
            pref.addWidget(QLabel("پوشهٔ PDF:"), 0, 0)
            self.pdf_edit = QLineEdit()
            pref.addWidget(self.pdf_edit, 0, 1)
            pick = QPushButton("انتخاب…")
            pref.addWidget(pick, 0, 2)
            pref.addWidget(QLabel("پوسته:"), 1, 0)
            self.theme_combo = QComboBox()
            self.theme_combo.addItem("روشن", "light")
            self.theme_combo.addItem("تیره", "dark")
            pref.addWidget(self.theme_combo, 1, 1)
            pref.addWidget(QLabel("اندازهٔ قلم:"), 2, 0)
            self.font_spin = QSpinBox()
            self.font_spin.setRange(8, 16)
            pref.addWidget(self.font_spin, 2, 1)
            pref.setColumnStretch(1, 1)
            self.pref_card.body.addLayout(pref)
            self.save_button = QPushButton("ذخیرهٔ تنظیمات")
            self.pref_card.body.addWidget(self.save_button, 0, Qt.AlignLeft)

            self.log_card = Card("گزارش برنامه", "", colors)
            self.log_box = QPlainTextEdit()
            self.log_box.setReadOnly(True)
            self.log_box.setMaximumHeight(180)
            self.log_card.body.addWidget(self.log_box)
            refresh = QPushButton("تازه‌سازی گزارش")
            self.log_card.body.addWidget(refresh, 0, Qt.AlignLeft)

            layout.addWidget(self.path_card)
            layout.addWidget(self.db_card)
            layout.addWidget(self.pref_card)
            layout.addWidget(self.log_card)
            layout.addStretch(1)
            self.apply_colors(colors)

        def apply_colors(self, colors: Dict[str, str]) -> None:
            self.colors = colors
            for card in (self.path_card, self.db_card, self.pref_card, self.log_card):
                card.apply_colors(colors)
            self.setStyleSheet(
                f"QLineEdit, QComboBox, QSpinBox, QPlainTextEdit {{"
                f" background: {colors['surface']}; color: {colors['ink']};"
                f" border: 1px solid {colors['border']}; border-radius: 8px; padding: 5px 8px; }}"
                f" QPushButton {{ background: {colors['surface2']}; color: {colors['ink']};"
                f" border: 1px solid {colors['border']}; border-radius: 8px;"
                f" padding: 6px 12px; }} QPushButton:hover {{ background: {colors['brand_soft']}; }}"
                f" QTableWidget {{ background: {colors['surface']}; color: {colors['ink']};"
                f" border: 1px solid {colors['border']}; border-radius: 8px; }}"
                f" QHeaderView::section {{ background: {colors['brand_soft']};"
                f" color: {colors['ink']}; border: none; padding: 5px; }}"
            )
            self.db_label.setStyleSheet(f"color: {colors['ink2']};")
            self.path_label.setStyleSheet(f"color: {colors['ink']};")

        def refresh(self) -> None:
            root = data_root()
            self.path_label.setText(
                f"پوشهٔ داده: {root}\n"
                f"دیتابیس: {database_path()}\n"
                f"کتاب کارها: {root / 'workbooks'}\n"
                f"خروجی‌ها: {root / 'exports'}"
            )
            summary = database_summary()
            self.db_label.setText(
                f"{fa(summary['workbooks'])} کتاب کار، {fa(summary['contracts'])} قرارداد، "
                f"{fa(summary['notes'])} یادداشت — حجم {fa(round(summary['size'] / 1024, 1))} کیلوبایت"
            )
            rows = list_workbooks()
            self.workbooks_table.setRowCount(len(rows))
            for index, item in enumerate(rows):
                values = [fa(item.get("id")), item.get("filename"),
                          fa(item.get("rows")), str(item.get("imported_at") or "")[:16]]
                for column, value in enumerate(values):
                    self.workbooks_table.setItem(index, column, QTableWidgetItem(str(value)))
            self.log_box.setPlainText(log_text()[-4000:])

    class MainWindow(QWidget):
        """پنجرهٔ اصلی برنامه."""

        def __init__(self, colors: Optional[Dict[str, str]] = None) -> None:
            super().__init__()
            self.colors = colors or THEMES["light"]
            self.data: Optional[Dashboard] = None
            self.worker: Optional[QThread] = None
            self.theme_name = setting_get("theme", "light") or "light"
            self.colors = THEMES.get(self.theme_name, THEMES["light"])
            self.font_size = int(setting_get("font_size", "10") or 10)
            self.pdf_folder = setting_get("pdf_folder", "") or ""
            self.setWindowTitle(f"{APP_NAME} — نسخهٔ {APP_VERSION}")
            self.resize(1440, 900)
            self._build()
            self.apply_theme(self.theme_name, self.font_size)
            self.settings_page.refresh()
            QTimer.singleShot(120, self.load_initial)

        # ---------- ساخت رابط ----------
        def _build(self) -> None:
            root = QHBoxLayout(self)
            root.setContentsMargins(0, 0, 0, 0)
            root.setSpacing(0)
            self.rail = QFrame()
            self.rail.setObjectName("Rail")
            self.rail.setFixedWidth(210)
            rail = QVBoxLayout(self.rail)
            rail.setContentsMargins(14, 18, 14, 18)
            rail.setSpacing(6)
            self.brand = QLabel("داشبورد مهر\nخلیج فارس")
            self.brand.setFont(make_font(13, True))
            self.brand.setAlignment(Qt.AlignCenter)
            rail.addWidget(self.brand)
            self.version_label = QLabel(f"نسخهٔ {APP_VERSION}")
            self.version_label.setFont(make_font(7))
            self.version_label.setAlignment(Qt.AlignCenter)
            rail.addWidget(self.version_label)
            rail.addSpacing(14)
            self.nav_buttons: Dict[str, QPushButton] = {}
            for key, label in (("dashboard", "داشبورد"), ("contracts", "قراردادها"),
                               ("report", "گزارش"), ("settings", "تنظیمات")):
                button = QPushButton(label)
                button.setCheckable(True)
                button.setCursor(Qt.PointingHandCursor)
                button.clicked.connect(lambda _checked, name=key: self.show_page(name))
                rail.addWidget(button)
                self.nav_buttons[key] = button
            rail.addStretch(1)
            self.data_label = QLabel("")
            self.data_label.setFont(make_font(7))
            self.data_label.setWordWrap(True)
            self.data_label.setAlignment(Qt.AlignCenter)
            rail.addWidget(self.data_label)

            center = QWidget()
            layout = QVBoxLayout(center)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(0)
            topbar = QFrame()
            topbar.setObjectName("Topbar")
            topbar.setFixedHeight(66)
            bar = QHBoxLayout(topbar)
            bar.setContentsMargins(18, 10, 18, 10)
            self.page_title = QLabel("داشبورد")
            self.page_title.setFont(make_font(12, True))
            bar.addWidget(self.page_title)
            bar.addStretch(1)
            self.update_button = QPushButton("به‌روزرسانی از فایل ۰۱")
            self.update_button.setObjectName("Primary")
            self.update_button.setCursor(Qt.PointingHandCursor)
            self.update_button.clicked.connect(self.start_update)
            bar.addWidget(self.update_button)
            self.reload_button = QPushButton("بازخوانی")
            self.reload_button.clicked.connect(self.load_initial)
            bar.addWidget(self.reload_button)
            self.theme_button = QPushButton("تیره")
            self.theme_button.clicked.connect(self.toggle_theme)
            bar.addWidget(self.theme_button)
            layout.addWidget(topbar)

            self.stack = QStackedWidget()
            self.dashboard_page = DashboardPage(self.colors)
            self.contracts_page = ContractsPage(self.colors)
            self.report_page = ReportPage(self.colors)
            self.settings_page = SettingsPage(self.colors)
            for page in (self.dashboard_page, self.contracts_page,
                         self.report_page, self.settings_page):
                self.stack.addWidget(page)
            layout.addWidget(self.stack, 1)
            root.addWidget(center, 1)
            root.addWidget(self.rail)
            self.show_page("dashboard")

            self.settings_page.import_button.clicked.connect(self.import_workbook)
            self.settings_page.load_selected.clicked.connect(self.load_selected_workbook)
            self.settings_page.reset_button.clicked.connect(self.reset_database)
            self.settings_page.open_folder.clicked.connect(self.open_data_folder)
            self.settings_page.save_button.clicked.connect(self.save_settings)
            self.settings_page.pdf_edit.setText(self.pdf_folder)
            pick = None
            for child in self.settings_page.findChildren(QPushButton):
                if child.text() == "انتخاب…":
                    pick = child
            if pick is not None:
                pick.clicked.connect(self.pick_pdf_folder)
            for child in self.settings_page.findChildren(QPushButton):
                if child.text() == "تازه‌سازی گزارش":
                    child.clicked.connect(self.settings_page.refresh)
            self.settings_page.theme_combo.setCurrentIndex(
                1 if self.theme_name == "dark" else 0)
            self.settings_page.font_spin.setValue(self.font_size)

            self.report_page.export_button.clicked.connect(self.export_contracts)
            self.report_page.report_button.clicked.connect(self.export_report)
            self.report_page.print_button.clicked.connect(self.print_report)

        def apply_theme(self, name: str, font_size: int = 10) -> None:
            self.theme_name = "dark" if name == "dark" else "light"
            self.colors = THEMES[self.theme_name]
            self.font_size = int(font_size)
            app = QApplication.instance()
            if app is not None:
                base = make_font(self.font_size)
                app.setFont(base)
                app.setLayoutDirection(Qt.RightToLeft)
            colors = self.colors
            self.setStyleSheet(
                f"QWidget {{ color: {colors['ink']}; font-family: Tahoma;"
                f" font-size: {self.font_size}pt; }}"
                f" QWidget#Rail {{ background: {colors['surface']};"
                f" border-left: 1px solid {colors['border']}; }}"
                f" QFrame#Topbar {{ background: {colors['surface']};"
                f" border-bottom: 1px solid {colors['border']}; }}"
                f" QPushButton#Primary {{ background: {colors['brand']}; color: white;"
                f" border: none; border-radius: 9px; padding: 8px 16px; font-weight: bold; }}"
                f" QPushButton#Primary:hover {{ background: {colors['brand']}; }}"
                f" QMainWindow, QWidget {{ background: {colors['bg']}; }}"
            )
            self.brand.setStyleSheet(f"color: {colors['brand']};")
            self.version_label.setStyleSheet(f"color: {colors['ink3']};")
            self.data_label.setStyleSheet(f"color: {colors['ink3']};")
            for key, button in self.nav_buttons.items():
                button.setStyleSheet(
                    f"QPushButton {{ text-align: right; padding: 9px 12px;"
                    f" border: none; border-radius: 9px; background: transparent;"
                    f" color: {colors['ink2']}; }}"
                    f" QPushButton:hover {{ background: {colors['surface2']}; }}"
                    f" QPushButton:checked {{ background: {colors['brand_soft']};"
                    f" color: {colors['brand']}; font-weight: bold; }}"
                )
            self.theme_button.setText("روشن" if self.theme_name == "dark" else "تیره")
            for page in (self.dashboard_page, self.contracts_page, self.report_page):
                page.apply_colors(colors)
            self.settings_page.apply_colors(colors)
            setting_set("theme", self.theme_name)
            setting_set("font_size", self.font_size)

        def toggle_theme(self) -> None:
            self.apply_theme("light" if self.theme_name == "dark" else "dark",
                             self.font_size)

        def show_page(self, name: str) -> None:
            pages = {"dashboard": 0, "contracts": 1, "report": 2, "settings": 3}
            titles = {"dashboard": "داشبورد", "contracts": "قراردادها",
                      "report": "گزارش مدیریتی", "settings": "تنظیمات"}
            self.stack.setCurrentIndex(pages.get(name, 0))
            self.page_title.setText(titles.get(name, "داشبورد"))
            for key, button in self.nav_buttons.items():
                button.setChecked(key == name)
            if name == "settings":
                self.settings_page.refresh()
            if name == "report":
                self.report_page.show_data(self.data)
            if name == "contracts":
                self.contracts_page.set_data(self.data, self.pdf_folder)

        # ---------- داده ----------
        def apply_data(self, data: Optional[Dashboard]) -> None:
            self.data = data
            self.dashboard_page.show_data(data)
            self.contracts_page.set_data(data, self.pdf_folder)
            self.report_page.show_data(data)
            if data is None:
                self.data_label.setText("داده‌ای در دیتابیس نیست")
                return
            self.data_label.setText(
                f"{fa(data.total)} قرارداد\n{fa(round(data.percent or 0, 2))}٪ پیشرفت"
            )

        def load_initial(self) -> None:
            workbook_id = int(setting_get("current_workbook", "0") or 0)
            self._start_worker(lambda: load_dashboard(workbook_id), self._on_loaded,
                               "در حال خواندن داده‌ها…")

        def _on_loaded(self, data: Optional[Dashboard]) -> None:
            if data is None:
                log_line("هیچ کتاب کاری در دیتابیس نیست؛ منتظر بارگذاری می‌مانیم.")
            self.apply_data(data)

        def _start_worker(self, job, on_done, message: str = "") -> None:
            if self.worker is not None and self.worker.isRunning():
                QMessageBox.information(self, "در حال اجرا", "یک عملیات دیگر در حال اجراست.")
                return
            if message:
                self.page_title.setText(message)

            class JobThread(QThread):
                done = pyqtSignal(object)
                failed = pyqtSignal(str)

                def __init__(self, function, parent=None):
                    super().__init__(parent)
                    self.function = function

                def run(self):
                    try:
                        result = self.function()
                    except Exception as exc:
                        write_error_log(traceback.format_exc())
                        self.failed.emit(str(exc))
                        return
                    self.done.emit(result)

            thread = JobThread(job, self)
            thread.done.connect(lambda result: self._worker_done(result, on_done))
            thread.failed.connect(self._worker_failed)
            self.worker = thread
            self._set_busy(True)
            thread.start()

        def _worker_done(self, result: Any, callback) -> None:
            self._set_busy(False)
            callback(result)

        def _worker_failed(self, message: str) -> None:
            self._set_busy(False)
            QMessageBox.critical(self, "خطا", message)
            self.show_page("dashboard")

        def _set_busy(self, busy: bool) -> None:
            self.update_button.setEnabled(not busy)
            self.reload_button.setEnabled(not busy)
            self.settings_page.import_button.setEnabled(not busy)
            self.update_button.setText("در حال اجرا…" if busy else "به‌روزرسانی از فایل ۰۱")
            if not busy:
                self.page_title.setText(
                    {"dashboard": "داشبورد", "contracts": "قراردادها",
                     "report": "گزارش مدیریتی", "settings": "تنظیمات"}.get(
                        self._current_page(), "داشبورد"))

        def _current_page(self) -> str:
            index = self.stack.currentIndex()
            return {0: "dashboard", 1: "contracts", 2: "report", 3: "settings"}.get(index, "dashboard")

        # ---------- عملیات ----------
        def import_workbook(self, path: str = "") -> None:
            if not path:
                path, _filter = QFileDialog.getOpenFileName(
                    self, "انتخاب کتاب کار (دارای شیت colorcode)", str(data_root()),
                    "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
                )
            if not path:
                return
            self._start_worker(lambda: self._import_job(path), self._on_imported,
                               "در حال خواندن کتاب کار…")

        def _import_job(self, path: str) -> Dict[str, Any]:
            blob = Path(path).read_bytes()
            dashboard = parse_workbook(blob, "", Path(path).name)
            if not dashboard.contracts:
                raise ValueError("در این کتاب کار هیچ قراردادی پیدا نشد.")
            profile = workbook_profile(blob)
            kind = "generated" if not profile.get("has_export") else "excel"
            workbook_id = save_workbook(blob, dashboard, kind)
            dashboard.workbook_id = workbook_id
            dashboard.note = f"دیتابیس داخلی — شناسهٔ {fa(workbook_id)}"
            return {"dashboard": dashboard, "id": workbook_id}

        def _on_imported(self, result: Dict[str, Any]) -> None:
            dashboard = result["dashboard"]
            self.apply_data(dashboard)
            self.settings_page.refresh()
            self.show_page("dashboard")
            QMessageBox.information(
                self, "بارگذاری شد",
                f"کتاب کار «{dashboard.source_name}» با {fa(dashboard.total)} قرارداد "
                f"ذخیره شد (شناسهٔ {fa(result['id'])})."
            )

        def load_selected_workbook(self) -> None:
            rows = self.settings_page.workbooks_table.selectionModel()
            selected = rows.selectedRows() if rows else []
            if not selected:
                QMessageBox.information(self, "انتخاب", "یک کتاب کار را از فهرست انتخاب کنید.")
                return
            cell = self.settings_page.workbooks_table.item(selected[0].row(), 0)
            workbook_id = int(cell.text().translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789"))
                              or 0) if cell else 0
            if not workbook_id:
                return
            setting_set("current_workbook", workbook_id)
            self._start_worker(lambda: load_dashboard(workbook_id), self._on_loaded)

        def start_update(self) -> None:
            path, _filter = QFileDialog.getOpenFileName(
                self, "انتخاب فایل ۰۱ (خروجی سامانه)",
                setting_get("last_source", str(data_root())),
                "فایل اکسل (*.xlsx *.xlsm);;همهٔ فایل‌ها (*)",
            )
            if not path:
                return
            setting_set("last_source", str(Path(path).parent))
            self._start_worker(lambda: self._update_job(path), self._on_updated,
                               "در حال ساخت نسخهٔ به‌روز…")

        def _update_job(self, path: str) -> Dict[str, Any]:
            source = read_source(path)
            base_blob = b""
            base_name = ""
            current = int(setting_get("current_workbook", "0") or 0)
            candidates: List[Tuple[str, bytes]] = []
            if current:
                try:
                    blob = load_workbook_blob(current)
                    candidates.append((f"db-{current}", blob))
                except Exception as exc:
                    log_line(f"خواندن کتاب کار فعلی ممکن نشد: {exc}")
            folder = Path(path).parent
            for candidate in sorted(folder.glob("*.xlsx")):
                if candidate.name == Path(path).name:
                    continue
                try:
                    blob = candidate.read_bytes()
                except Exception:
                    continue
                profile = workbook_profile(blob)
                if profile.get("has_export"):
                    candidates.append((candidate.name, blob))
            for name, blob in candidates[:6]:
                try:
                    info = prepare_update(blob, source, Path(tempfile.gettempdir()) / "probe.xlsx")
                except Exception as exc:
                    log_line(f"کتاب کار «{name}» برای پایه مناسب نبود: {exc}")
                    continue
                if info.get("verified"):
                    base_blob, base_name = blob, name
                    break
            if not base_blob:
                base_blob = build_base_workbook(source)
                base_name = f"{Path(path).stem}.xlsx"
            stamp = datetime.now().strftime("%Y%m%d-%H%M")
            out_name = f"{Path(base_name).stem if base_name.startswith('db-') else Path(base_name).stem}-به‌روز-{stamp}.xlsx"
            out_path = data_subdir("output") / safe_filename(out_name)
            info = prepare_update(base_blob, source, out_path)
            if not info.get("verified"):
                raise ValueError(
                    "نسخهٔ به‌روز ساخته شد اما هیچ علامت وضعیت یا درصدی نداشت؛ "
                    "برای همین ذخیره نشد.\n"
                    f"فایل: {out_path}"
                )
            dashboard = parse_workbook(str(out_path), "", out_path.name)
            dashboard.note = f"از فایل ۰۱ ساخته شد — {fa(dashboard.total)} قرارداد"
            workbook_id = save_workbook(out_path.read_bytes(), dashboard, "generated")
            dashboard.workbook_id = workbook_id
            info["workbook_id"] = workbook_id
            info["output_path"] = str(out_path)
            info["dashboard"] = dashboard
            return info

        def _on_updated(self, info: Dict[str, Any]) -> None:
            dashboard = info.get("dashboard")
            if dashboard is not None:
                self.apply_data(dashboard)
            self.settings_page.refresh()
            self.show_page("dashboard")
            lines = [
                f"ردیف‌های «اجرا - مهر»: {fa(info.get('matched_rows') or 0)} "
                f"({info.get('method') or '—'})",
                f"علامت‌های وضعیت: {fa(info.get('verified_marks') or 0)}",
                f"درصد کل: {fa(info.get('verified_percent') or 0)}٪",
                f"کتاب کار جدید با شناسهٔ {fa(info.get('workbook_id') or 0)} ذخیره شد.",
                f"فایل: {info.get('output_path')}",
            ]
            QMessageBox.information(self, "به‌روزرسانی انجام شد", "\n".join(lines))

        def reset_database(self) -> None:
            answer = QMessageBox.question(
                self, "ساخت دیتابیس از صفر",
                "همهٔ کتاب کارها و یادداشت‌های دیتابیس داخلی پاک و از نو ساخته شود؟\n\n"
                "فایل‌های اکسل روی دیسک پاک نمی‌شوند. یک نسخهٔ پشتیبان هم در "
                "پوشهٔ backup ساخته می‌شود.",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            reset_database()
            setting_set("current_workbook", 0)
            self.apply_data(None)
            self.settings_page.refresh()
            QMessageBox.information(
                self, "آماده است",
                "دیتابیس از صفر ساخته شد.\n"
                "برای شروع، «بارگذاری کتاب کار…» یا «به‌روزرسانی از فایل ۰۱» را بزنید."
            )

        def pick_pdf_folder(self) -> None:
            folder = QFileDialog.getExistingDirectory(
                self, "انتخاب پوشهٔ PDF", self.pdf_folder or str(data_root()))
            if folder:
                self.pdf_folder = folder
                self.settings_page.pdf_edit.setText(folder)
                setting_set("pdf_folder", folder)
                self.contracts_page.set_data(self.data, self.pdf_folder)

        def save_settings(self) -> None:
            self.pdf_folder = self.settings_page.pdf_edit.text().strip()
            setting_set("pdf_folder", self.pdf_folder)
            theme = str(self.settings_page.theme_combo.currentData() or "light")
            size = int(self.settings_page.font_spin.value())
            self.apply_theme(theme, size)
            self.contracts_page.set_data(self.data, self.pdf_folder)
            QMessageBox.information(self, "ذخیره شد", "تنظیمات ذخیره و اعمال شد.")

        def open_data_folder(self) -> None:
            from PyQt5.QtCore import QUrl
            from PyQt5.QtGui import QDesktopServices
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(data_root())))

        def export_contracts(self) -> None:
            if not self.data or not self.data.contracts:
                QMessageBox.information(self, "خروجی", "داده‌ای برای خروجی نیست.")
                return
            folder = data_subdir("exports")
            name = f"قراردادها-{datetime.now():%Y%m%d-%H%M}.xlsx"
            path, _filter = QFileDialog.getSaveFileName(
                self, "ذخیرهٔ خروجی اکسل", str(folder / name), "فایل اکسل (*.xlsx)")
            if not path:
                return
            try:
                count = export_contracts(path, self.data.contracts)
            except Exception as exc:
                QMessageBox.critical(self, "خطا", str(exc))
                return
            QMessageBox.information(self, "انجام شد", f"{fa(count)} ردیف ذخیره شد:\n{path}")

        def export_report(self) -> None:
            if not self.data or not self.data.contracts:
                QMessageBox.information(self, "گزارش", "داده‌ای برای گزارش نیست.")
                return
            folder = data_subdir("exports")
            name = f"گزارش-مدیریتی-{datetime.now():%Y%m%d-%H%M}.xlsx"
            path, _filter = QFileDialog.getSaveFileName(
                self, "ذخیرهٔ گزارش مدیریتی", str(folder / name), "فایل اکسل (*.xlsx)")
            if not path:
                return
            try:
                export_report(path, self.data)
            except Exception as exc:
                QMessageBox.critical(self, "خطا", str(exc))
                return
            QMessageBox.information(self, "انجام شد", f"گزارش ذخیره شد:\n{path}")

        def print_report(self) -> None:
            if not self.data or not self.data.contracts:
                QMessageBox.information(self, "چاپ", "داده‌ای برای گزارش نیست.")
                return
            if QPrinter is None or QPrintDialog is None:
                QMessageBox.warning(self, "چاپ", "پشتیبانی چاپ در این سیستم نیست.")
                return
            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            dialog = QPrintDialog(printer, self)
            if dialog.exec_() != QDialog.Accepted:
                return
            from PyQt5.QtGui import QTextDocument
            counts = self.data.flag_counts()
            rows = "".join(
                f"<tr><td>{label}</td><td>{fa(counts.get(key, 0))}</td></tr>"
                for key, label, _column in FLAG_DEFS)
            html = (
                f"<div dir='rtl' style='font-family: Tahoma;'>"
                f"<h2>گزارش مدیریتی قراردادها</h2>"
                f"<p>کتاب کار: {self.data.source_name}</p>"
                f"<p>تعداد قرارداد: {fa(self.data.total)} — درصد پیشرفت: "
                f"{fa(round(self.data.percent or 0, 2))}٪</p>"
                f"<table border='1' cellspacing='0' cellpadding='6'>{rows}</table>"
                f"</div>"
            )
            document = QTextDocument()
            document.setHtml(html)
            document.print_(printer)

    def run_gui() -> int:
        """پنجرهٔ برنامه را باز می‌کند."""
        app = QApplication(sys.argv[:1])
        app.setApplicationName(APP_NAME)
        app.setLayoutDirection(Qt.RightToLeft)
        window = MainWindow()
        window.show()
        return app.exec_()


# ===========================================================================
# ۱۲) خط فرمان
# ===========================================================================
def cli_check() -> int:
    summary = database_summary()
    print(f"پوشهٔ داده: {data_root()}")
    print(f"دیتابیس: {summary['path']} ({fa(summary['workbooks'])} کتاب کار)")
    data = load_dashboard()
    if data is None:
        print("هیچ کتاب کاری در دیتابیس نیست.")
        return 1
    print(f"کتاب کار: {data.source_name} — {data.total} قرارداد")
    print(f"درصد پیشرفت: {data.percent}")
    print("وضعیت‌ها: " + "، ".join(
        f"{label}={count}" for (key, label, _c), count in
        zip(FLAG_DEFS, data.flag_counts().values())))
    return 0


def cli_import(path: str) -> int:
    blob = Path(path).read_bytes()
    dashboard = parse_workbook(blob, "", Path(path).name)
    workbook_id = save_workbook(blob, dashboard, "excel")
    print(f"ذخیره شد: شناسهٔ {workbook_id} — {dashboard.total} قرارداد — "
          f"درصد {dashboard.percent}")
    return 0


def cli_update(path: str) -> int:
    source = read_source(path)
    current = int(setting_get("current_workbook", "0") or 0)
    base_blob = load_workbook_blob(current) if current else b""
    if not base_blob:
        base_blob = build_base_workbook(source)
    out_path = data_subdir("output") / f"به-روز-{datetime.now():%Y%m%d-%H%M}.xlsx"
    info = prepare_update(base_blob, source, out_path)
    print(f"تطبیق: {info.get('matched_rows')} ({info.get('method')}) — "
          f"علامت: {info.get('verified_marks')} — درصد: {info.get('verified_percent')}")
    if info.get("verified"):
        dashboard = parse_workbook(str(out_path), "", out_path.name)
        workbook_id = save_workbook(out_path.read_bytes(), dashboard, "generated")
        print(f"ذخیره شد با شناسهٔ {workbook_id}: {out_path}")
        return 0
    print("فایل خروجی تأیید نشد.")
    return 2


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=f"{APP_NAME} — نسخهٔ {APP_VERSION}")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {APP_VERSION}")
    parser.add_argument("--check", action="store_true", help="بررسی داده‌ها بدون پنجره")
    parser.add_argument("--data-dir", help="مسیر داده‌ها (پیش‌فرض D:\\DATANEW)")
    parser.add_argument("--import-xlsx", metavar="PATH", help="بارگذاری کتاب کار")
    parser.add_argument("--update-from", metavar="PATH", help="به‌روزرسانی از فایل ۰۱")
    args = parser.parse_args(argv)
    if args.data_dir:
        os.environ["MEHR_DATA_DIR"] = os.path.abspath(args.data_dir)
    log_line(f"شروع برنامه — نسخهٔ {APP_VERSION}")
    log_line(f"پوشهٔ داده: {data_root()}")
    if args.check:
        return cli_check()
    if args.import_xlsx:
        return cli_import(args.import_xlsx)
    if args.update_from:
        return cli_update(args.update_from)
    if not QT_READY:
        print("خطا: PyQt5 نصب نیست.\n" + QT_ERROR)
        print("دستور نصب: pip install PyQt5 openpyxl")
        return 1
    return run_gui()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # pragma: no cover
        traceback.print_exc()
        path = write_error_log(traceback.format_exc())
        print(f"خطای اجرای برنامه: {exc}")
        if path:
            print(f"جزئیات: {path}")
        sys.exit(1)

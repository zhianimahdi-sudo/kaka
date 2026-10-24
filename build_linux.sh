#!/usr/bin/env bash
# ساخت فایل اجرایی تک‌فایلی لینوکس (نسخهٔ ۴ — SQLite داخلی، بدون MySQL)
set -e
python3 -m pip install --upgrade pip
python3 -m pip install PyQt5 openpyxl pyinstaller
rm -rf build dist

python3 -m PyInstaller \
  --noconfirm --onefile --windowed \
  --name MehrDashboard \
  --hidden-import openpyxl \
  --hidden-import openpyxl.styles \
  --hidden-import openpyxl.utils \
  --hidden-import openpyxl.worksheet \
  --hidden-import openpyxl.cell \
  mehr_dashboard_qt.py
echo "ساخته شد: dist/MehrDashboard"

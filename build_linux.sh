#!/usr/bin/env bash
# ساخت فایل اجرایی تک‌فایلی لینوکس (بدون نیاز به نصب پایتون روی سیستم مقصد)
set -e
python3 -m pip install --upgrade pip
python3 -m pip install PyQt5 openpyxl mysql-connector-python pyinstaller
rm -rf build dist
EXTRA=()
if [ -f "F.O.C(Mehr).xlsx" ]; then
  EXTRA+=(--add-data "F.O.C(Mehr).xlsx:.")
fi

python3 -m PyInstaller \
  --noconfirm --onefile --windowed \
  --name MehrDashboard \
  "${EXTRA[@]}" \
  --hidden-import openpyxl \
  --hidden-import openpyxl.styles \
  --hidden-import openpyxl.utils --hidden-import mysql.connector --collect-submodules mysql.connector \
  mehr_dashboard_qt.py
echo "ساخته شد: dist/MehrDashboard"

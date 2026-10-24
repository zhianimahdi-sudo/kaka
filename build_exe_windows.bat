@echo off
REM =====================================================================
REM  ساخت فایل اجرایی تک‌فایلی ویندوز (MehrDashboard.exe) — نسخهٔ ۴
REM  پیش‌نیاز: Python 3.9+ نصب شده و در PATH باشد
REM  داده‌ها هنگام اجرا در D:\DATANEW ساخته می‌شوند.
REM =====================================================================
chcp 65001 >nul
setlocal

echo [1/3] نصب پیش‌نیازها ...
python -m pip install --upgrade pip >nul
python -m pip install PyQt5 openpyxl pyinstaller || goto :error

echo [2/3] پاک‌سازی خروجی قبلی ...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/3] ساخت فایل تک‌فایلی ...
python -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name MehrDashboard ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  --hidden-import openpyxl.worksheet ^
  --hidden-import openpyxl.cell ^
  mehr_dashboard_qt.py || goto :error

echo.
echo =====================================================
echo  ساخته شد: dist\MehrDashboard.exe
echo  داده‌ها در D:\DATANEW ساخته می شوند؛ برای شروع دکمهٔ
echo  «به‌روزرسانی از فایل ۰۱» را بزنید.
echo =====================================================
goto :end

:error
echo.
echo خطا در ساخت. متن خطای بالا را بررسی کنید.
exit /b 1

:end
endlocal
pause

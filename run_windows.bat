@echo off
REM =====================================================================
REM  اجرای داشبورد مهر روی ویندوز (نصب خودکار پیش‌نیازها در اولین اجرا)
REM  نیاز: Python 3.9 یا بالاتر که در PATH باشد (python.org/downloads)
REM =====================================================================
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ==========================================================
echo   داشبورد قراردادهای مهر خلیج فارس
echo ==========================================================
echo.

python --version >nul 2>&1
if errorlevel 1 goto :nopython

python -c "import PyQt5, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [1/2] نصب پیش‌نیازها (فقط بار اول؛ چند دقیقه طول می‌کشد)...
    python -m pip install --upgrade pip >nul
    python -m pip install -r requirements.txt || goto :pipfail
    echo       نصب انجام شد.
) else (
    echo [1/2] پیش‌نیازها از قبل نصب هستند.
)

echo [2/2] اجرای برنامه...
echo.
echo   اگر دیتابیس آماده نیست، در صفحهٔ «تنظیمات» فایل اکسل را
echo   «بارگذاری در دیتابیس» کنید یا «اجرای محلی (بدون دیتابیس)» را بزنید.
echo.
python mehr_dashboard_qt.py
if errorlevel 1 goto :runfail
goto :end

:nopython
echo.
echo  [خطا] Python نصب نیست یا در PATH نیست.
echo  از https://www.python.org/downloads/ نسخهٔ 3.9 یا بالاتر را نصب کنید
echo  و هنگام نصب گزینهٔ "Add Python to PATH" را فعال کنید.
echo.
pause
exit /b 1

:pipfail
echo.
echo  [خطا] نصب پیش‌نیازها انجام نشد. اتصال اینترنت یا دسترسی pip را بررسی کنید.
echo  دستور دستی:  python -m pip install PyQt5 openpyxl mysql-connector-python
echo.
pause
exit /b 1

:runfail
echo.
echo  [خطا] برنامه با خطا بسته شد.
echo  فایل گزارش «mehr_error.log» در همان پوشهٔ این فایل ساخته می‌شود؛ آن را بفرستید.
echo.
pause
exit /b 1

:end
endlocal

@echo off
REM =====================================================================
REM  اجرای داشبورد مهر روی ویندوز (نصب خودکار پیش‌نیازها در اولین اجرا)
REM  نیاز: Python 3.9 یا بالاتر که در PATH باشد (python.org/downloads)
REM  داده‌ها: D:\DATANEW   —   بدون MySQL، بدون رمز، بدون سرویس
REM =====================================================================
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ==========================================================
echo   داشبورد قراردادهای مهر خلیج فارس — نسخهٔ ۴٫۰٫۰
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
echo   داده‌ها در پوشهٔ D:\DATANEW ساخته می‌شود.
echo   برای شروع: دکمهٔ «به‌روزرسانی از فایل ۰۱» یا «بارگذاری کتاب کار…».
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
echo  دستور دستی:  python -m pip install PyQt5 openpyxl
echo.
pause
exit /b 1

:runfail
echo.
echo  [خطا] برنامه با خطا بسته شد.
echo  فایل گزارش: D:\DATANEW\logs\errors.log  (یا mehr_error.log کنار برنامه)
echo  آن فایل را بفرستید.
echo.
pause
exit /b 1

:end
endlocal

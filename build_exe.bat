@echo off
echo ========================================================
echo   Membangun Executable: BPS Email Dispatcher Portable
echo ========================================================
echo.

REM Memastikan dependensi terpasang
python -m pip install -r requirements.txt

echo.
echo Mengompilasi dengan PyInstaller...
python -m PyInstaller --noconsole --onefile --name "BPS_Email_Dispatcher" --clean app.py

echo.
if exist "dist\BPS_Email_Dispatcher.exe" (
    echo ========================================================
    echo   BERHASIL! Executable siap di:
    echo   dist\BPS_Email_Dispatcher.exe
    echo ========================================================
) else (
    echo [ERROR] Kompilasi gagal. Periksa pesan error di atas.
)

pause

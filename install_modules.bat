@echo off
cd /d "%~dp0"
echo ========================================
echo Installation des modules Python
echo ========================================
echo.
python scripts\install_modules.py
echo.
pause

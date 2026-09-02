@echo off
cd /d "%~dp0"

REM Vérifie si Python est dans le PATH
python --version > nul 2>&1
if errorlevel 1 (
    echo Python n'est pas trouve dans le PATH
    echo Tentative d'utiliser Python depuis le chemin par defaut...
    set "PYTHON_CMD=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python39\python.exe"
    if not exist "%PYTHON_CMD%" (
        echo ERREUR: Python n'a pas ete trouve.
        echo Veuillez installer Python 3.9 ou superieur.
        pause
        exit /b 1
    )
) else (
    set "PYTHON_CMD=python"
)

REM Mode d'emploi : ouvrir_mode_emploi.bat ou MODE_EMPLOI.html

REM Lance l'application
"%PYTHON_CMD%" Interface.py

if errorlevel 1 (
    echo.
    echo Une erreur s'est produite lors du lancement de l'application.
    echo Verifiez que Python est bien installe et que tous les modules sont presents.
    pause
    exit /b 1
) 
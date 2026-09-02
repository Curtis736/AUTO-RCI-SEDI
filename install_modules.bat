@echo off
echo ========================================
echo Installation des modules Python
echo ========================================
echo.

REM Vérifier si Python est installé
python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR: Python n'est pas installe ou n'est pas dans le PATH
    echo Veuillez installer Python depuis https://www.python.org/
    pause
    exit /b 1
)

echo Python detecte:
python --version
echo.

REM Mettre à jour pip
echo Mise a jour de pip...
python -m pip install --upgrade pip
echo.

REM Installer tous les modules depuis requirements.txt
echo Installation des modules depuis requirements.txt...
python -m pip install -r requirements.txt
echo.

echo ========================================
echo Installation Poppler (conversion PDF)
echo ========================================
python -c "import poppler_setup; p=poppler_setup.ensure_poppler_installed(); print('Poppler OK:', p if p else 'echec - Internet requis')"
echo.

REM Installer pytest (optionnel pour les tests)
echo ========================================
echo Installation des modules optionnels (tests)...
echo ========================================
python -m pip install pytest
echo.

REM Vérifier l'installation
echo ========================================
echo Verification de l'installation...
echo ========================================
python -m pip list | findstr /i "pywin32 Pillow numpy python-docx docx2pdf pdf2image pypdf pytest"
echo.

echo ========================================
echo Installation terminee!
echo ========================================
pause

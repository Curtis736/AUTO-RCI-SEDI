@echo off
cd /d "%~dp0\.."
echo ========================================
echo    GENERATION DE LA DOCUMENTATION PDF
echo ========================================
echo.
pause

python scripts\create_documentation_pdf.py

echo.
echo ========================================
echo    FIN DE LA GENERATION
echo ========================================
echo.
echo Si la generation a reussi, vous trouverez
echo le fichier PDF dans le dossier courant.
echo.
pause 
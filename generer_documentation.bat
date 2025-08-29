@echo off
echo ========================================
echo    GENERATION DE LA DOCUMENTATION PDF
echo ========================================
echo.
echo Ce script va generer un PDF de la documentation
echo complete de l'application AUTO RCI.
echo.
pause

python create_documentation_pdf.py

echo.
echo ========================================
echo    FIN DE LA GENERATION
echo ========================================
echo.
echo Si la generation a reussi, vous trouverez
echo le fichier PDF dans le dossier courant.
echo.
pause 
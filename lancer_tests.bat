@echo off
REM Script batch pour lancer les tests unitaires AUTO_RCI
REM Ce script fonctionne sur Windows

echo ========================================
echo TESTS UNITAIRES AUTO_RCI
echo ========================================
echo.

REM Vérifier que Python est installé
python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR: Python n'est pas installé ou pas dans le PATH
    echo Veuillez installer Python et réessayer
    pause
    exit /b 1
)

echo Python détecté, lancement des tests...
echo.

REM Exécuter les tests
python scripts\run_tests.py

REM Afficher le résultat
if errorlevel 1 (
    echo.
    echo ========================================
    echo CERTAINS TESTS ONT ÉCHOUÉ
    echo ========================================
    pause
    exit /b 1
) else (
    echo.
    echo ========================================
    echo TOUS LES TESTS SONT PASSÉS!
    echo ========================================
    pause
    exit /b 0
)



















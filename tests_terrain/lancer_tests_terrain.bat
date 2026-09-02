@echo off
cd /d "%~dp0\.."
title AUTO_RCI - Tests terrain (LT reel)

echo ========================================
echo   TESTS TERRAIN - LT REEL
echo   test-unitaire/
echo   Doc : tests_terrain\DOCUMENTATION_TESTS.html
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR: Python introuvable dans le PATH
    pause
    exit /b 1
)

python tests_terrain\run_tests_terrain.py
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE%==0 (
    echo ========================================
    echo   TOUS LES TESTS SONT PASSES
    echo ========================================
) else (
    echo ========================================
    echo   CERTAINS TESTS ONT ECHOUE
    echo ========================================
)

pause
exit /b %EXIT_CODE%

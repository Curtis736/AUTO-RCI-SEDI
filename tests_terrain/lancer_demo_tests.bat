@echo off
cd /d "%~dp0\.."
title AUTO_RCI - Demo tests (succes puis echecs)

echo ========================================
echo   DEMO TESTS
echo   Etape 1 : tests qui reussissent
echo   Etape 2 : tests qui echouent (volontaire)
echo   Doc : tests_terrain\DOCUMENTATION_TESTS.html
echo ========================================
echo.

python tests_terrain\run_demo_tests.py
set EXIT_CODE=%ERRORLEVEL%

echo.
pause
exit /b %EXIT_CODE%

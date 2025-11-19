@echo off
:: Setting UTF-8 encoding
chcp 65001 >nul 2>&1
title Strength Calculation Generator

:menu
cls
echo ============================================================
echo     STRENGTH CALCULATION GENERATOR
echo ============================================================
echo.
echo   1. Test run (single product)
echo   2. Full processing (all products)
echo   3. Open results folder
echo   4. Open log file
echo   5. Open configuration
echo   6. Exit
echo.
set /p choice="Choose action (1-6): "

if "%choice%"=="1" goto test
if "%choice%"=="2" goto full
if "%choice%"=="3" goto results
if "%choice%"=="4" goto log
if "%choice%"=="5" goto config
if "%choice%"=="6" goto end

echo.
echo Invalid choice!
timeout /t 2 >nul
goto menu

:test
cls
echo ============================================================
echo TEST RUN
echo ============================================================
echo.
python test_single.py
echo.
pause
goto menu

:full
cls
echo ============================================================
echo FULL PROCESSING
echo ============================================================
echo.
set /p confirm="Are you sure? This may take 10-15 minutes (Y/N): "
if /i not "%confirm%"=="Y" goto menu
echo.
python main.py
echo.
pause
goto menu

:results
start "" "D:\Найденные_папки\Расчеты на прочность"
goto menu

:log
if exist "D:\Найденные_папки\log_РП.txt" (
    notepad "D:\Найденные_папки\log_РП.txt"
) else (
    echo Log file not found
    timeout /t 2 >nul
)
goto menu

:config
notepad "config\config.json"
goto menu

:end
exit

@echo off
:: Setting UTF-8 encoding
chcp 65001 >nul 2>&1

echo ========================================
echo Deploying Strength Calculation Generator
echo ========================================
echo.

:: Check Python installation
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo Install Python 3.8+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Python found
python --version
echo.

:: Install dependencies
echo Installing dependencies...
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo [OK] Dependencies installed
echo.

:: Check configuration
if not exist "config\config.json" (
    echo WARNING: config.json not found!
    echo Please create config\config.json file
    pause
    exit /b 1
)

echo [OK] Configuration found
echo.

echo ========================================
echo Deployment completed successfully!
echo ========================================
echo.
echo To test: python test_single.py
echo To run:  python main.py
echo.
pause

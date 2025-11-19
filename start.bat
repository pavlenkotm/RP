@echo off
echo ============================================================
echo FULL PROCESSING - All Products
echo ============================================================
echo.
echo This may take 10-15 minutes for 300 products
echo.
set /p confirm="Continue? (Y/N): "
if /i not "%confirm%"=="Y" exit

echo.
python main.py
echo.
pause

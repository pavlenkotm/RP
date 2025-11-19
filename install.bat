@echo off
echo Installing dependencies...
echo.

python -m pip install --upgrade pip
python -m pip install pandas==2.1.4
python -m pip install openpyxl==3.1.2
python -m pip install python-docx==1.1.0
python -m pip install pywin32==306
python -m pip install pdfplumber==0.10.3
python -m pip install Pillow==10.1.0
python -m pip install python-dateutil==2.8.2

echo.
echo Installation completed!
echo.
pause

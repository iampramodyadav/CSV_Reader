
REM pyinstaller --onefile --noconsole --icon=csv.ico csv_read.py
REM pyinstaller --onefile --windowed --icon=csv.ico csv_read.py
REM pyinstaller csv_read.spec
@echo off
REM Activate virtual environment
call "venv\temp1\Scripts\activate.bat"
REM pip install pyinstaller
REM pip install openpyxl
REM Change directory to your project folder
cd /d C:\Users\pramod\shared\tool\01table_data_editor

REM Run PyInstaller with the spec file
pyinstaller csv_read.spec

echo.
echo Build finished! Check the "dist" folder for your exe.
pause

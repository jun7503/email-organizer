@echo off
echo ====================================
echo Email Organizer - Build Script
echo ====================================
echo.

echo Checking Python installation...
python --version
if errorlevel 1 (
    echo ERROR: Python is not installed!
    echo Please install Python 3.10 or higher from python.org
    pause
    exit /b 1
)

echo.
echo Installing required packages...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building executable...
pyinstaller --onefile --windowed --name EmailOrganizer email_sorter.py

echo.
echo ====================================
echo Build complete!
echo ====================================
echo.
echo Your executable is located at:
echo dist\EmailOrganizer.exe
echo.
pause

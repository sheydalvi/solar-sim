@echo off
cd /d "%~dp0"

set ENV_NAME=repgen

echo Deleting old virtual environment (if exists)...
rmdir /s /q %ENV_NAME%

echo Creating virtual environment: %ENV_NAME%
python -m venv %ENV_NAME%

if errorlevel 1 (
    echo Failed to create virtual environment. Exiting...
    pause
    exit /b 1
)

echo Activating virtual environment...
call %ENV_NAME%\Scripts\activate.bat

echo Upgrading pip and essentials...
python -m pip install --upgrade pip setuptools wheel

echo Installing requirements from requirements.txt...
if not exist requirements.txt (
    echo ERROR: requirements.txt not found!
    pause
    exit /b 1
)

pip install -r requirements.txt

echo.
echo Environment "%ENV_NAME%" is ready and activated!
pause

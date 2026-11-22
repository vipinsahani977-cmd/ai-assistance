@echo off
title AI Voice Assistant Launcher

echo ======================================================
echo            AI VOICE ASSISTANT LAUNCHER
echo ======================================================

cd /d "%~dp0"

:: 1. Check for Virtual Environment
if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment folder .venv was not found!
    echo Please create it using: python -m venv .venv
    echo And install requirements: .venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

:: 2. Activate Virtual Environment
call .venv\Scripts\activate.bat

:: 3. Check Python and Key Packages
echo [*] Checking dependencies...
python -c "import customtkinter, speech_recognition, pyttsx3, google.genai, dotenv" >nul 2>&1
if errorlevel 1 (
    echo [!] Missing dependencies detected. Installing from requirements.txt...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies. Please check your internet connection.
        pause
        exit /b 1
    )
)

:: 4. Launch Application
echo [*] Starting AI Voice Assistant...
python main.py %*

if errorlevel 1 (
    echo.
    echo [ERROR] Assistant exited with error code %errorlevel%.
    pause
)

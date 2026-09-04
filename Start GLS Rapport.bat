@echo off
setlocal
cd /d "%~dp0"
title GLS Rapportgenerator

echo Starter GLS Rapportgenerator...
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    py app.py
    goto end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python app.py
    goto end
)

echo.
echo FEJL: Python blev ikke fundet.
echo Installer Python fra Microsoft Store eller python.org.
echo.

:end
pause

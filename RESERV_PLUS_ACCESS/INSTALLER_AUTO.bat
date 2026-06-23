@echo off
title RESERV+ Installation
color 0B

echo.
echo   RESERV+ - Demarrage de l'installation automatique...
echo.

cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALLER_AUTO.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   Le script a echoue. Consultez README_IMPORTATION_ACCESS.md
    echo.
    pause
)

@echo off
:: =============================================================================
:: RESERV+ — Lanceur d'installation
:: Double-cliquez sur ce fichier pour installer RESERV+ automatiquement
:: =============================================================================
title RESERV+ Installation
color 0B

echo.
echo   RESERV+ - Demarrage de l'installation automatique...
echo.

:: Aller dans le dossier du script (important si lancé depuis un autre endroit)
cd /d "%~dp0"

:: Lancer PowerShell avec bypass de la politique d'exécution
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALLER_AUTO.ps1"

:: Si PowerShell échoue (non installé, bloqué par GPO), afficher aide manuelle
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   [ATTENTION] Le script automatique a echoue.
    echo   Suivez le guide manuel : README_IMPORTATION_ACCESS.md
    echo.
    pause
)

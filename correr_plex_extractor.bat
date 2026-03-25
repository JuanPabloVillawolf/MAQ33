@echo off
title PLEX Cost Extractor
color 0A
echo.
echo  ======================================
echo    PLEX Cost Extractor - RIEDON
echo  ======================================
echo.

cd /d "%~dp0"

echo  Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python no esta instalado o no esta en el PATH.
    pause
    exit /b
)

echo  Iniciando extraccion de costos...
echo.
python plex_cost_extractor.py

echo.
echo  ======================================
echo  Proceso finalizado. Revisa el archivo:
echo  seer_export_con_costos.xlsx
echo  ======================================
echo.
pause

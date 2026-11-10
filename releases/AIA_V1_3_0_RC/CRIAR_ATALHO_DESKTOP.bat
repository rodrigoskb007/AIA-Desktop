@echo off
setlocal
cd /d "%~dp0"
title Criar atalho - AIA V1.3.0 RC
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0CRIAR_ATALHO_DESKTOP.ps1"
echo.
pause
endlocal

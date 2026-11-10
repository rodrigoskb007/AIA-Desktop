@echo off
setlocal
cd /d "%~dp0"
title Validacao - AIA V1.3.0 RC

where python >nul 2>&1
if %errorlevel%==0 (
    python "VALIDAR_AIA.py"
    goto :end
)
where py >nul 2>&1
if %errorlevel%==0 (
    py -3 "VALIDAR_AIA.py"
    goto :end
)

echo [ERRO] Python 3 nao encontrado.
:end
echo.
pause
endlocal

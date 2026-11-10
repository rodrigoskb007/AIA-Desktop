@echo off
setlocal
cd /d "%~dp0"
title Instalar dependencias - AIA V1.3.0 RC

echo ================================================
echo AIA V1.3.0 RC - INSTALACAO DE DEPENDENCIAS
echo ================================================
echo.

where python >nul 2>&1
if %errorlevel%==0 (
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    goto :check
)

where py >nul 2>&1
if %errorlevel%==0 (
    py -3 -m pip install --upgrade pip
    py -3 -m pip install -r requirements.txt
    goto :check
)

echo [ERRO] Python 3 nao foi encontrado.
echo Instale o Python 3 para Windows com Tcl/Tk e tente novamente.
pause
exit /b 1

:check
echo.
echo Executando validacao da instalacao...
call "%~dp0VALIDAR_AIA.bat"
endlocal

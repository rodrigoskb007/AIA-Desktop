@echo off
setlocal
cd /d "%~dp0"
title AIA V1.3.0 RC

where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" /b pythonw "AIA_V1_3_0_RC.py"
    goto :end
)

where pyw >nul 2>&1
if %errorlevel%==0 (
    start "" /b pyw -3 "AIA_V1_3_0_RC.py"
    goto :end
)

echo.
echo [ERRO] Python 3 nao foi encontrado neste computador.
echo Instale o Python 3 para Windows e execute INSTALAR_DEPENDENCIAS_AIA.bat.
echo.
pause
:end
endlocal

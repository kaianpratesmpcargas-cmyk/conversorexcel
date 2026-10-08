@echo off
setlocal
cd /d "%~dp0"
title MP CARGAS - Gerar EXE

echo ================================================
echo   MP CARGAS - GERADOR DO EXE
echo ================================================
echo.

set PY=python
%PY% --version >nul 2>&1
if errorlevel 1 (
    set PY=%LOCALAPPDATA%\Programs\Python\Python314\python.exe
)

if not exist "%PY%" (
    echo ERRO: Python nao foi encontrado.
    echo Instale o Python ou ajuste o caminho no arquivo GERAR_EXE.bat.
    pause
    exit /b 1
)

echo Python encontrado:
"%PY%" --version

echo.
echo Instalando dependencias...
"%PY%" -m pip install -r requirements.txt
if errorlevel 1 goto erro

"%PY%" -m pip install pyinstaller
if errorlevel 1 goto erro

echo.
echo Gerando EXE...
"%PY%" -m PyInstaller --clean --noconfirm MP_CARGAS_Conversor_SSW.spec
if errorlevel 1 goto erro

echo.
echo ================================================
echo   EXE GERADO COM SUCESSO
echo   dist\MP_CARGAS_Conversor_SSW.exe
echo ================================================
echo.
pause
exit /b 0

:erro
echo.
echo ================================================
echo   ERRO AO GERAR O EXE
echo ================================================
echo.
pause
exit /b 1

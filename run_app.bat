@echo off
title Email Feedback App - Inicializando...

REM Verifica se o Python está instalado
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Python nao esta instalado ou nao esta no PATH.
    echo Instale o Python 3.12+ em: https://www.python.org/downloads/
    pause
    exit /b
)

echo.
echo Verificando e instalando dependencias...
python -m pip install --upgrade pip
python -m pip install pandas openpyxl

echo.
echo Iniciando o sistema...
python main.py

echo.
pause

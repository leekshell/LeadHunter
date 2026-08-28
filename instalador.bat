@echo off
:: Define codificação UTF-8 para exibir acentos corretamente
chcp 65001 > nul
title LeadHunter - Instalador de Dependencias

echo ============================================================
echo           LEADHUNTER - INSTALADOR AUTOMATICO
echo ============================================================
echo.

:: 1. VERIFICA SE ESTA RODANDO COMO ADMINISTRADOR
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERRO] Este script precisa ser executado como ADMINISTRADOR!
    echo.
    echo Por favor, clique com o botão direito no arquivo .bat
    echo e selecione "Executar como administrador".
    echo.
    pause
    exit /b
)

echo [OK] Permissoes de Administrador confirmadas.
echo.

:: 2. VERIFICA SE O PYTHON ESTA INSTALADO
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERRO] Python nao foi encontrado no sistema!
    echo Baixe e instale o Python em https://www.python.org/
    echo Lembre-se de marcar a opcao "Add Python to PATH" na instalacao.
    echo.
    pause
    exit /b
)

echo [1/3] Atualizando o PIP...
python -m pip install --upgrade pip

echo.
echo [2/3] Instalando bibliotecas Python (Requirements)...
pip install customtkinter playwright httpx beautifulsoup4 openpyxl pillow psutil pyinstaller

echo.
echo [3/3] Baixando e instalando o navegador Chromium para o Playwright...
playwright install chromium

echo.
echo ============================================================
echo      INSTALACAO CONCLUIDA COM SUCESSO! TUDO PRONTO.
echo ============================================================
echo.
pause
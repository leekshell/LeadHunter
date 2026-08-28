@echo off
:: LeadHunter - Instalador automático (Windows)
:: Cria um ambiente virtual, instala dependencias e o Chromium do Playwright.
chcp 65001 > nul
title LeadHunter - Instalador de Dependencias

echo ============================================================
echo            LEADHUNTER - INSTALADOR AUTOMATICO
echo ============================================================
echo.

:: 1. VERIFICA O PYTHON
where python >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERRO] Python nao foi encontrado no sistema!
    echo Baixe e instale o Python em https://www.python.org/
    echo Marque a opcao "Add Python to PATH" na instalacao.
    pause
    exit /b 1
)

for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [OK] Python %PYVER% encontrado.
echo.

:: 2. CRIA O AMBIENTE VIRTUAL (evita conflitos com bibliotecas globais)
set "VENV=.venv"
if not exist "%VENV%\Scripts\python.exe" (
    echo [1/4] Criando ambiente virtual...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERRO] Nao foi possivel criar o ambiente virtual.
        pause
        exit /b 1
    )
) else (
    echo [1/4] Ambiente virtual ja existe. Pulando criacao.
)

call "%VENV%\Scripts\activate.bat"

echo.
echo [2/4] Atualizando o PIP...
python -m pip install --upgrade pip

echo.
echo [3/4] Instalando bibliotecas Python (requirements.txt)...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERRO] Falha ao instalar as dependencias.
    echo Confirme a conexao com a internet e tente novamente.
    pause
    exit /b 1
)

echo.
echo [4/4] Baixando e instalando o navegador Chromium para o Playwright...
python -m playwright install chromium

echo.
echo ============================================================
echo      INSTALACAO CONCLUIDA COM SUCESSO! TUDO PRONTO.
echo ============================================================
echo.
echo Iniciando o LeadHunter...
echo.
python App.py

pause

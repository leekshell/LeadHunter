@echo off
:: LeadHunter - Gerar executável (.exe) com PyInstaller
chcp 65001 > nul
title LeadHunter - Compilacao do Executavel

echo ============================================================
echo         LEADHUNTER - GERAR EXECUTAVEL (.EXE)
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [ERRO] Execute o instalador.bat primeiro para criar o ambiente virtual.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

echo [1/3] Garantindo dependencias...
pip install -r requirements.txt

echo.
echo [2/3] Garantindo navegador Chromium...
python -m playwright install chromium

echo.
echo [3/3] Gerando o executavel...
pyinstaller --noconfirm --onefile --windowed --name LeadHunter App.py

if errorlevel 1 (
    echo [ERRO] Falha ao gerar o executavel.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo      EXECUTAVEL GERADO EM: dist\LeadHunter.exe
echo ============================================================
echo.
pause

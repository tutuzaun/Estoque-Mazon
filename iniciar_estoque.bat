@echo off
chcp 65001 >nul
title Restaurante Mazon - Sistema de Gestao de Estoque

:: Garante que o script rode sempre a partir da pasta do projeto
cd /d "%~dp0"

echo ======================================================================
echo   RESTAURANTE MAZON - GESTAO DE ESTOQUE E INTELIGENCIA FINANCEIRA
echo ======================================================================
echo.

:: 1. Verifica se o Python esta disponivel no sistema
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado no PATH do Windows.
    echo Instale o Python 3.11+ em https://www.python.org/downloads/
    echo Certifique-se de marcar a opcao "Add python.exe to PATH".
    pause
    exit /b 1
)

:: 2. Cria o ambiente virtual (.venv) automaticamente caso nao exista
if not exist ".venv\Scripts\activate.bat" (
    echo [1/3] Criando ambiente virtual isolado (.venv)...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERRO] Falha ao criar o ambiente virtual .venv.
        pause
        exit /b 1
    )
)

:: 3. Ativa o ambiente virtual
call ".venv\Scripts\activate.bat"

:: 4. Instala ou atualiza os pacotes do requirements.txt se necessario
if not exist ".venv\.deps_v2_ok" (
    echo [2/3] Instalando/atualizando dependencias do projeto (requirements.txt)...
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if %errorlevel% equ 0 (
        echo ok > ".venv\.deps_v2_ok"
    ) else (
        echo [ERRO] Falha ao instalar as dependencias. Verifique sua conexao.
        pause
        exit /b 1
    )
) else (
    echo [2/3] Dependencias ja verificadas no ambiente .venv.
)

:: 5. Inicia a aplicacao Streamlit
echo [3/3] Iniciando servidor Streamlit (Restaurante Mazon)...
echo.
streamlit run app.py --browser.gatherUsageStats=false

pause

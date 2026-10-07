@echo off
cd /d "%~dp0backend"

REM Na primeira execucao (ou se a venv for apagada) cria o ambiente e instala tudo.
if not exist "venv\Scripts\activate.bat" (
    echo.
    echo === Primeira execucao: criando ambiente virtual e instalando dependencias ===
    echo === Isso leva ~30 segundos. So acontece uma vez.                          ===
    echo.
    python -m venv venv
    call venv\Scripts\activate.bat
    python -m pip install --upgrade pip
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate.bat
)

start "" http://127.0.0.1:8000
uvicorn app.main:app --port 8000

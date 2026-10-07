@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\python.exe" (
    echo venv nao encontrada. Rode iniciar.bat uma vez primeiro.
    pause
    exit /b 1
)
echo Rodando os testes do coletor e do parser...
echo.
venv\Scripts\python.exe testes.py
echo.
pause

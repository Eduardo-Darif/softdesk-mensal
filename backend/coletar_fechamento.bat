@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\python.exe" (
    echo [%date% %time%] ERRO: venv nao encontrada. Rode iniciar.bat uma vez.>> data\coletor.log
    exit /b 1
)
echo [%date% %time%] --- Fechamento (mes anterior) --->> data\coletor.log
venv\Scripts\python.exe coletar_chamados.py anterior >> data\coletor.log 2>&1
echo [%date% %time%] Fim (codigo %errorlevel%)>> data\coletor.log

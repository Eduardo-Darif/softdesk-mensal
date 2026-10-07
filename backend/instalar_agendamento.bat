@echo off
setlocal
set "PASTA=%~dp0"
echo Criando as tarefas agendadas do Softdesk Mensal (rodam as 10:00)...
echo.
schtasks /Create /TN "SoftdeskMensal - Diario" /TR "\"%PASTA%coletar_diario.bat\"" /SC DAILY /ST 10:00 /RU "%USERNAME%" /IT /F
schtasks /Create /TN "SoftdeskMensal - Fechamento" /TR "\"%PASTA%coletar_fechamento.bat\"" /SC MONTHLY /D 1 /ST 10:00 /RU "%USERNAME%" /IT /F
echo.
echo === Tarefas criadas ===
schtasks /Query /TN "SoftdeskMensal - Diario"
schtasks /Query /TN "SoftdeskMensal - Fechamento"
echo.
echo Pronto^! Diaria todo dia 10:00; Fechamento no dia 1 de cada mes.
pause

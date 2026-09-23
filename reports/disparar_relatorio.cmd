@echo off
REM Dispara o workflow do relatorio no GitHub (usa o gh ja autenticado neste PC).
REM Chamado pela tarefa agendada do Windows "FinancasCasal - Relatorio Telegram" (08:07 e 20:07).
REM force=false: se o relatorio ja foi enviado ha menos de 5h por outro gatilho, o script pula.
set LOG=%LOCALAPPDATA%\fincasal_relatorio.log
echo [%date% %time%] disparando... >> "%LOG%"
cd /d "%~dp0"
"C:\Program Files\GitHub CLI\gh.exe" workflow run relatorio.yml -f force=false >> "%LOG%" 2>&1
echo [%date% %time%] exit=%ERRORLEVEL% >> "%LOG%"

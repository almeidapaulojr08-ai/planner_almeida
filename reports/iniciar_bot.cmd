@echo off
REM Mantem o Salem (telegram_bot.js) rodando. Se cair, reinicia em 15 s.
REM Chamado pela tarefa agendada "FinancasCasal - Salem Telegram" (ao fazer logon), via iniciar_bot.vbs (janela oculta).
set LOG=%LOCALAPPDATA%\fincasal_bot.log
cd /d "%~dp0"
:loop
echo [%date% %time%] iniciando bot >> "%LOG%"
"C:\Program Files\nodejs\node.exe" telegram_bot.js >> "%LOG%" 2>&1
echo [%date% %time%] bot saiu (exit=%ERRORLEVEL%), reiniciando em 15s >> "%LOG%"
timeout /t 15 /nobreak > nul
goto loop

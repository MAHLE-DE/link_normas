@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
py -3 "..\02_BANCO_ATUAL\gerar_pendencias.py"
if errorlevel 1 goto erro
echo Planilha criada em 02_BANCO_ATUAL\pendencias_codigos_sem_cadeia.xlsx.
echo Guarde uma copia preenchida antes de gerar novamente.
pause
exit /b 0
:erro
echo Falha ao gerar pendencias. Confira o resultado da base completa e as fontes.
pause
exit /b 1

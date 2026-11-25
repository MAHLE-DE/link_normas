@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
py -3 "..\02_BANCO_ATUAL\aplicar_pendencias.py"
if errorlevel 1 goto erro
echo Aplicacao concluida. Confira os avisos; depois rode 01 e 03.
pause
exit /b 0
:erro
echo Falha ao aplicar pendencias. Confira os campos e destinos indicados no erro.
pause
exit /b 1

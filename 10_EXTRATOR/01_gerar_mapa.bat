@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

echo Gerando o mapa de codigos a partir das fontes do repositorio...
py -3 "..\02_BANCO_ATUAL\gerar_mapa_normas.py"
if errorlevel 1 goto erro
echo Mapa e planilha de revisao gerados em 02_BANCO_ATUAL.
pause
exit /b 0

:erro
echo Falha na geracao do mapa. Leia o erro acima e confira as fontes e dependencias.
pause
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

if not exist "..\05_DOCUMENTS\" echo 05_DOCUMENTS ausente: usando somente as fontes do repositorio.
echo Gerando mapa local em mapa_base, sem sobrescrever o mapa versionado...
py -3 "..\02_BANCO_ATUAL\gerar_mapa_normas.py" --saida "..\02_BANCO_ATUAL\mapa_base"
if errorlevel 1 goto erro
echo Mapa e planilha de revisao gerados em 02_BANCO_ATUAL\mapa_base.
pause
exit /b 0

:erro
echo Falha na geracao do mapa. Leia o erro acima e confira as fontes e dependencias.
pause
exit /b 1

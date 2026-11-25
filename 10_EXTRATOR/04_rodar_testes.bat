@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

echo Executando os testes do mapa e do extrator...
py -3 -m pytest "..\02_BANCO_ATUAL" "." -q
if errorlevel 1 goto erro
echo Todos os testes passaram.
pause
exit /b 0

:erro
echo Ha testes com falha. Confira os detalhes acima antes de usar os resultados.
pause
exit /b 1

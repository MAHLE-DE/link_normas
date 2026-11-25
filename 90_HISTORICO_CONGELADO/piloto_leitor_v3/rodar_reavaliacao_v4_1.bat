@echo off
setlocal
cd /d "%~dp0"
py -3 06_avaliar_piloto_v4_1.py
if errorlevel 1 goto erro
echo.
echo Reavaliacao concluida. Envie avaliacao_piloto_v4_1.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Mantenha todos os arquivos na pasta original do piloto V3.
pause
exit /b 1

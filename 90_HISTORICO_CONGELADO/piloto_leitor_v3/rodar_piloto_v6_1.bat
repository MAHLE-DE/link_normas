@echo off
setlocal
cd /d "%~dp0"
echo Executando extrator V6.1 com progresso por documento...
py -3 05_executar_piloto_v6_1.py
if errorlevel 1 goto erro
echo Avaliando V6.1...
py -3 06_avaliar_piloto_v6_1.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v6_1.xlsx e avaliacao_piloto_v6_1.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha na execucao. Leia a mensagem acima.
pause
exit /b 1

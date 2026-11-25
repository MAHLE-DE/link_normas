@echo off
setlocal
cd /d "%~dp0"
echo Selecionando 40 documentos oficiais para o segundo piloto...
py -3 selecionar_amostra_validacao_40_corrigido.py
if errorlevel 1 goto erro
echo.
echo Concluido. Resultados em 04_NOVO_TESTE_40.
pause
exit /b 0
:erro
echo.
echo Falha. Leia a mensagem acima.
pause
exit /b 1

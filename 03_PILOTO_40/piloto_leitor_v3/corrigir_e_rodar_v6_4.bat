@echo off
setlocal
cd /d "%~dp0"
echo Corrigindo V6.4...
py -3 corrigir_v6_4.py
if errorlevel 1 goto erro
py -3 -m py_compile norm_reference_extractor_v6_4.py
if errorlevel 1 goto erro
echo Executando V6.4 corrigida...
py -3 05_executar_piloto_v6_4.py
if errorlevel 1 goto erro
echo Avaliando V6.4...
py -3 06_avaliar_piloto_v6_4.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v6_4.xlsx e avaliacao_piloto_v6_4.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Leia a mensagem acima.
pause
exit /b 1

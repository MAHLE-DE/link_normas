@echo off
setlocal
cd /d "%~dp0"
echo Aplicando patch V6.4 sobre a V6.2...
py -3 aplicar_patch_v6_4.py
if errorlevel 1 goto erro
echo Validando scripts...
py -3 -m py_compile norm_reference_extractor_v6_4.py 05_executar_piloto_v6_4.py 06_avaliar_piloto_v6_4.py
if errorlevel 1 goto erro
echo Executando extrator V6.4...
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
echo Falha. Confirme que o patch foi extraido na pasta que contem os scripts da V6.2.
pause
exit /b 1

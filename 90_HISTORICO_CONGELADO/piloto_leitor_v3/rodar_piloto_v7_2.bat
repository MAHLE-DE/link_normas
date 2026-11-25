@echo off
setlocal
cd /d "%~dp0"
echo Aplicando V7.2 sobre a V7.0...
py -3 aplicar_patch_v7_2.py
if errorlevel 1 goto erro
py -3 -m py_compile norm_reference_extractor_v7_2.py 05_executar_piloto_v7_2.py 07_avaliar_piloto_v7_2.py
if errorlevel 1 goto erro
echo Executando extrator V7.2...
py -3 05_executar_piloto_v7_2.py
if errorlevel 1 goto erro
echo Avaliando com cadeias vigentes...
py -3 07_avaliar_piloto_v7_2.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v7_2.xlsx e avaliacao_piloto_v7_2.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Extraia o pacote na pasta que contem a V7.0.
pause
exit /b 1

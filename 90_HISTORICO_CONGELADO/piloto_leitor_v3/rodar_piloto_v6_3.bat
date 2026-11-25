@echo off
setlocal
cd /d "%~dp0"
echo Aplicando patch V6.3...
py -3 aplicar_patch_v6_3.py
if errorlevel 1 goto erro
echo Validando os scripts...
py -3 -m py_compile norm_reference_extractor_v6_3.py 05_executar_piloto_v6_3.py 06_avaliar_piloto_v6_3.py
if errorlevel 1 goto erro
echo Executando extrator V6.3...
py -3 05_executar_piloto_v6_3.py
if errorlevel 1 goto erro
echo Avaliando V6.3...
py -3 06_avaliar_piloto_v6_3.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v6_3.xlsx e avaliacao_piloto_v6_3.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Confirme que este patch foi extraido na pasta da V6.2.
pause
exit /b 1

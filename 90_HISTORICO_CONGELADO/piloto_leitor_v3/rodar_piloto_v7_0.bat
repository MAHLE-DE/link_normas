@echo off
setlocal
cd /d "%~dp0"
echo Aplicando V7.0 sobre a V6.4 corrigida...
py -3 aplicar_patch_v7_0.py
if errorlevel 1 goto erro
py -3 -m py_compile norm_reference_extractor_v7_0.py 05_executar_piloto_v7_0.py 07_avaliar_piloto_v7_0.py
if errorlevel 1 goto erro
echo Executando extrator V7.0...
py -3 05_executar_piloto_v7_0.py
if errorlevel 1 goto erro
echo Avaliando contra o ground truth corrigido...
py -3 07_avaliar_piloto_v7_0.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v7_0.xlsx e avaliacao_piloto_v7_0.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Confirme que este pacote foi extraido na pasta da V6.4 corrigida.
pause
exit /b 1

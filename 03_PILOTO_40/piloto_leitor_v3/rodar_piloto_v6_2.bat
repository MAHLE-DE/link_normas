@echo off
setlocal
cd /d "%~dp0"
echo Aplicando patch V6.2...
py -3 aplicar_patch_v6_2.py
if errorlevel 1 goto erro
echo Executando extrator V6.2...
py -3 05_executar_piloto_v6_2.py
if errorlevel 1 goto erro
echo Avaliando V6.2...
py -3 06_avaliar_piloto_v6_2.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v6_2.xlsx e avaliacao_piloto_v6_2.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Verifique se o patch foi extraido na mesma pasta da V6.1.
pause
exit /b 1

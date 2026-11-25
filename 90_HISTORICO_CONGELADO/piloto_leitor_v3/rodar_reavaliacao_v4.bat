@echo off
setlocal
cd /d "%~dp0"
py -3 06_avaliar_piloto_v4.py
if errorlevel 1 goto erro
echo.
echo Reavaliacao concluida. Envie avaliacao_piloto_v4.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Confirme que esta pasta contem resultado_piloto.xlsx, base_mestra_cadeias_consolidada.xlsx e equivalences_pilot.json.
pause
exit /b 1

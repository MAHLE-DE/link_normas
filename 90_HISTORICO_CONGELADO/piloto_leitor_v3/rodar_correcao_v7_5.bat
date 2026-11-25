@echo off
setlocal
cd /d "%~dp0"
py -3 corrigir_ground_truth_e_grafo_v7_5.py
if errorlevel 1 goto erro
py -3 11_reavaliar_v7_4_com_correcao_v7_5.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie avaliacao_v7_5.xlsx.
pause
exit /b 0
:erro
echo Falha. Mantenha resultado_piloto_v7_4.xlsx na mesma pasta.
pause
exit /b 1

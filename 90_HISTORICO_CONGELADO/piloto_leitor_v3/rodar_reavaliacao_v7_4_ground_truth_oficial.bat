@echo off
setlocal
cd /d "%~dp0"
py -3 10_reavaliar_v7_4_ground_truth_oficial.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie avaliacao_piloto_v7_4_ground_truth_oficial.xlsx.
pause
exit /b 0
:erro
echo Falha. Mantenha nesta pasta resultado_piloto_v7_4.xlsx e os arquivos deste pacote.
pause
exit /b 1

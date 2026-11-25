@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

echo Validando as 40 normas contra o gabarito...
py -3 extrator_normas.py --entrada "..\03_PILOTO_40\corpus\normas" --manifesto "..\03_PILOTO_40\corpus\corpus_40.json" --saida "saida\resultado_40.xlsx" --gabarito "..\03_PILOTO_40\corpus\gabarito_40_v1.json"
if errorlevel 1 goto erro
echo Validacao concluida. Confira as metricas acima e a aba Avaliacao em saida\resultado_40.xlsx.
pause
exit /b 0

:erro
echo Falha na validacao. Leia o erro acima e confira o corpus, o mapa e as dependencias.
pause
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
py -3 extrator_normas.py --entrada "..\TESTE40_V3" --manifesto "..\TESTE40_V3\corpus\corpus_40_v3.json" --gabarito "..\TESTE40_V3\corpus\gabarito_40_v3.json" --saida "saida\resultado_v3.xlsx"
if errorlevel 1 goto erro
echo Validacao da terceira amostra concluida. Confira as metricas e a aba Avaliacao.
pause
exit /b 0
:erro
echo Falha na validacao da terceira amostra. Confira a mensagem acima.
pause
exit /b 1

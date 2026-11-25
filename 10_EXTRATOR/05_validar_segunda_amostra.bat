@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
py -3 extrator_normas.py --entrada "..\04_NOVO_TESTE_40\normas_v2" --manifesto "..\04_NOVO_TESTE_40\corpus\corpus_40_v2.json" --gabarito "..\04_NOVO_TESTE_40\corpus\gabarito_40_v2.json" --saida "saida\resultado_v2.xlsx"
if errorlevel 1 goto erro
echo Validacao da segunda amostra concluida. Confira as metricas e a aba Avaliacao.
pause
exit /b 0
:erro
echo Falha na validacao da segunda amostra. Confira a mensagem acima.
pause
exit /b 1

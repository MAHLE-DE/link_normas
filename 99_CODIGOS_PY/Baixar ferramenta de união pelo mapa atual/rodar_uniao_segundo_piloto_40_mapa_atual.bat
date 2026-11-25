@echo off
setlocal
cd /d "%~dp0"
echo Lendo caminhos atuais de estrutura_link_normas.json...
py -3 03_unir_amostra_validacao_40_mapa_atual.py
if errorlevel 1 (
  echo.
  echo Falha. Verifique a mensagem acima. A ferramenta usa somente os caminhos encontrados no mapa atual.
  pause
  exit /b 1
)
echo.
echo Concluido. As saidas foram gravadas na mesma pasta da amostra_validacao_40.json indicada pelo mapa.
pause

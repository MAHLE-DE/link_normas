@echo off
setlocal
cd /d "%~dp0"
echo V7.3: executando a V7.2 existente...
if exist rodar_piloto_v7_2.bat (
  call rodar_piloto_v7_2.bat
  if errorlevel 1 goto erro
) else if not exist resultado_piloto_v7_2.xlsx (
  echo ERRO: nao encontrei rodar_piloto_v7_2.bat nem resultado_piloto_v7_2.xlsx.
  goto erro
)
echo V7.3: aplicando grafo completo, historico, supressao de familias e recuperacao contextual...
py -3 08_posprocessar_v7_3.py
if errorlevel 1 goto erro
echo V7.3: avaliando...
py -3 09_avaliar_piloto_v7_3.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v7_3.xlsx e avaliacao_piloto_v7_3.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Extraia este pacote na mesma pasta da V7.2.
pause
exit /b 1

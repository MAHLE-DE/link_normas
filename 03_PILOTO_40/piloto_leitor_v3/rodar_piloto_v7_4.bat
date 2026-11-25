@echo off
setlocal
cd /d "%~dp0"
echo Aplicando V7.4 sobre a V7.3...
py -3 aplicar_patch_v7_4.py
if errorlevel 1 goto erro
py -3 -m py_compile 08_posprocessar_v7_4.py 09_avaliar_piloto_v7_4.py
if errorlevel 1 goto erro
if not exist resultado_piloto_v7_2.xlsx (
  echo Executando a V7.2 para gerar o resultado-base...
  call rodar_piloto_v7_2.bat
  if errorlevel 1 goto erro
)
echo Pos-processando V7.4...
py -3 08_posprocessar_v7_4.py
if errorlevel 1 goto erro
echo Avaliando V7.4...
py -3 09_avaliar_piloto_v7_4.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v7_4.xlsx e avaliacao_piloto_v7_4.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha. Confirme que o pacote foi extraido na mesma pasta da V7.3 e V7.2.
pause
exit /b 1

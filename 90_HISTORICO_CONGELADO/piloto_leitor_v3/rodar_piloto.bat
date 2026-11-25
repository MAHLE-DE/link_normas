@echo off
setlocal
cd /d "%~dp0"

py -3 -c "import fitz, docx, openpyxl, pypdf" >nul 2>&1
if errorlevel 1 (
  echo Bibliotecas ausentes. Iniciando instalacao automatica...
  py -3 -m pip install --upgrade pip
  if errorlevel 1 goto erro
  py -3 -m pip install PyMuPDF python-docx openpyxl pypdf
  if errorlevel 1 goto erro
)

echo Executando leitor no piloto de 40 normas...
py -3 05_executar_piloto.py
if errorlevel 1 goto erro

echo Avaliando resultado...
py -3 06_avaliar_piloto.py
if errorlevel 1 goto erro

echo.
echo Piloto concluido. Envie resultado_piloto.xlsx e avaliacao_piloto.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha na execucao. Leia a mensagem acima.
pause
exit /b 1

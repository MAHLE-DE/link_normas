@echo off
setlocal
cd /d "%~dp0"

py -3 -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Instalando openpyxl...
  py -3 -m pip install openpyxl
  if errorlevel 1 goto erro
)

echo Comparando cadastro SharePoint com 05_DOCUMENTS...
py -3 comparar_sharepoint_com_05_documents.py
if errorlevel 1 goto erro

echo.
echo Concluido. Os resultados foram salvos em:
echo C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\02_BANCO_ATUAL
pause
exit /b 0

:erro
echo.
echo Falha na comparacao. Leia a mensagem acima.
pause
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
py -3 -c "import fitz, docx, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Instalando dependencias...
  py -3 -m pip install PyMuPDF python-docx openpyxl
  if errorlevel 1 goto erro
)
echo Executando extrator V6...
py -3 05_executar_piloto_v6.py
if errorlevel 1 goto erro
echo Avaliando V6...
py -3 06_avaliar_piloto_v6.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie resultado_piloto_v6.xlsx e avaliacao_piloto_v6.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha na execucao. Leia a mensagem acima.
pause
exit /b 1

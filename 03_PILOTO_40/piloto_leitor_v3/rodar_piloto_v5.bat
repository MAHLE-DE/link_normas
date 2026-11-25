@echo off
setlocal
cd /d "%~dp0"

for %%F in (piloto_40_normas.json equivalences_pilot.json base_mestra_cadeias_consolidada.xlsx) do (
  if not exist "%%F" (
    echo Arquivo necessario ausente: %%F
    echo Extraia este pacote dentro da pasta original do piloto V3.
    pause
    exit /b 1
  )
)

py -3 -c "import fitz, docx, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Instalando dependencias...
  py -3 -m pip install PyMuPDF python-docx openpyxl
  if errorlevel 1 goto erro
)

echo Executando extrator V5...
py -3 05_executar_piloto_v5.py
if errorlevel 1 goto erro

echo Avaliando V5...
py -3 06_avaliar_piloto_v5.py
if errorlevel 1 goto erro

echo.
echo Concluido. Envie resultado_piloto_v5.xlsx e avaliacao_piloto_v5.xlsx.
pause
exit /b 0
:erro
echo.
echo Falha na execucao. Leia a mensagem acima.
pause
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
py -3 -c "import openpyxl, pymupdf, docx" >nul 2>&1
if errorlevel 1 (
 echo Instalando dependencias...
 py -3 -m pip install openpyxl pymupdf python-docx
 if errorlevel 1 goto erro
)
echo Consolidando cadeias oficiais, legadas e pendencias...
py -3 consolidar_familias_cadeias_legadas.py
if errorlevel 1 goto erro
echo.
echo Concluido. Resultados em 02_BANCO_ATUAL:
echo - familia_cadeias_completa_v7_5.xlsx
echo - grafo_cadeias_v7_5_completo.json
echo - pendencias_cadeias_v7_5.json
pause
exit /b 0
:erro
echo Falha. Confirme cadastro_canonico_v3_SEM_URL_SHAREPOINT.json e 05_DOCUMENTS.
pause
exit /b 1

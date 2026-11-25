@echo off
setlocal
cd /d "%~dp0"
py -3 -c "import pypdf" >nul 2>&1
if errorlevel 1 (
 echo Instalando pypdf...
 py -3 -m pip install pypdf
 if errorlevel 1 goto erro
)
echo Unindo PDFs exclusivamente de Valid-Documents...
py -3 08_unir_40_normas_valid_documents.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie o PDF e o manifesto JSON gerados.
pause
exit /b 0
:erro
echo.
echo Falha. Leia a lista de arquivos pendentes exibida acima.
pause
exit /b 1

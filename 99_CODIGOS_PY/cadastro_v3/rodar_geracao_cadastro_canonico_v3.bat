@echo off
setlocal
cd /d "%~dp0"

py -3 -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Instalando openpyxl...
  py -3 -m pip install openpyxl
  if errorlevel 1 goto erro
)

echo Gerando as duas versoes do Cadastro Canonico V3...
py -3 gerar_cadastro_canonico_v3.py
if errorlevel 1 goto erro

echo.
echo Concluido. Arquivos gerados em 02_BANCO_ATUAL:
echo - cadastro_canonico_v3_COM_URL_SHAREPOINT.json
echo - cadastro_canonico_v3_SEM_URL_SHAREPOINT.json
echo - diagnostico_cadastro_canonico_v3.xlsx
pause
exit /b 0

:erro
echo.
echo Falha. Leia a mensagem acima.
pause
exit /b 1

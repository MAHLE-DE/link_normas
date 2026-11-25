@echo off
setlocal
cd /d "%~dp0"
echo [1/2] Compilando banco de cadastros...
py -3 01_compilar_banco_cadastros.py
if errorlevel 1 goto :erro

echo.
echo [2/2] Mapeando validstandards...
py -3 02_mapear_validstandards.py
if errorlevel 1 goto :erro

echo.
echo Concluido. Arquivos gerados em:
echo C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\banco_cadastros_compilado.json
echo C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\inventario_validstandards.json
pause
exit /b 0
:erro
echo.
echo O processo falhou. Revise a mensagem acima.
pause
exit /b 1

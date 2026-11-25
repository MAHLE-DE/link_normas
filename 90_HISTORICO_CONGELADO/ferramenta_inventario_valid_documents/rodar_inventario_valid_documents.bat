@echo off
setlocal
cd /d "%~dp0"
py -3 07_mapear_valid_documents.py
if errorlevel 1 goto erro
echo.
echo Concluido. Envie o arquivo:
echo C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas\inventario_valid_documents.json
pause
exit /b 0
:erro
echo.
echo Falha. Leia a mensagem acima.
pause
exit /b 1

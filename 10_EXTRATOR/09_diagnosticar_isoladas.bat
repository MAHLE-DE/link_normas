@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

if not exist "..\05_DOCUMENTS\" goto sem_documentos
echo Diagnosticando as normas sem nenhuma ligacao (somente leitura)...
py -3 diagnosticar_isoladas.py
if errorlevel 1 goto erro
echo.
echo Pronto. Faca push de 02_BANCO_ATUAL\diagnostico_isoladas.xlsx
pause
exit /b 0

:sem_documentos
echo A pasta 05_DOCUMENTS nao foi encontrada na raiz do repositorio.
pause
exit /b 1

:erro
echo Falha no diagnostico. Leia o erro acima.
pause
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

echo Instalando dependencias para o Python selecionado por py -3...
py -3 -m pip install PyMuPDF python-docx openpyxl pytest
if errorlevel 1 goto erro
echo Dependencias instaladas. Execute 01_gerar_mapa.bat.
pause
exit /b 0

:erro
echo Falha na instalacao. Confira o Python 3.10 ou mais recente e a conexao com a internet.
pause
exit /b 1

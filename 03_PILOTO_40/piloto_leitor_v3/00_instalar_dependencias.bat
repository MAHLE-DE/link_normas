@echo off
setlocal
cd /d "%~dp0"
echo Atualizando o instalador de pacotes Python...
py -3 -m pip install --upgrade pip
if errorlevel 1 goto erro

echo.
echo Instalando bibliotecas necessarias...
py -3 -m pip install PyMuPDF python-docx openpyxl pypdf
if errorlevel 1 goto erro

echo.
echo Testando as bibliotecas...
py -3 -c "import fitz, docx, openpyxl, pypdf; print('Dependencias instaladas corretamente.')"
if errorlevel 1 goto erro

echo.
echo Instalacao concluida.
pause
exit /b 0
:erro
echo.
echo A instalacao falhou. Copie a mensagem acima para analise.
pause
exit /b 1

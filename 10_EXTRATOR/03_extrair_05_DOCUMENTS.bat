@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

if not exist "..\05_DOCUMENTS\" goto sem_documentos
echo Extraindo PDFs e DOCX de 05_DOCUMENTS e suas subpastas...
py -3 extrator_normas.py --entrada "..\05_DOCUMENTS" --recursivo --saida "saida\resultado_05_DOCUMENTS.xlsx"
if errorlevel 1 goto erro
echo Extracao concluida. Confira Correlacoes, Descartes e Ignorados em saida\resultado_05_DOCUMENTS.xlsx.
pause
exit /b 0

:sem_documentos
echo A pasta 05_DOCUMENTS nao foi encontrada na raiz do repositorio.
echo Copie as normas para essa pasta e execute este arquivo novamente.
pause
exit /b 1

:erro
echo Falha na extracao. Leia o erro acima e confira o mapa e as dependencias.
pause
exit /b 1

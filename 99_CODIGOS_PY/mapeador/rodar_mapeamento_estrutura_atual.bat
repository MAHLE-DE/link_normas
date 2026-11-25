@echo off
setlocal
cd /d "%~dp0"

echo Mapeando a estrutura atual de link_normas...
py -3 mapear_estrutura_atual.py
if errorlevel 1 goto erro

echo.
echo Concluido.
echo O arquivo estrutura_atual_link_normas.json foi criado em:
echo C:\Users\M0242198\cauaferrari\ESTUDOS\link_normas
pause
exit /b 0

:erro
echo.
echo Falha no mapeamento. Leia a mensagem acima.
pause
exit /b 1

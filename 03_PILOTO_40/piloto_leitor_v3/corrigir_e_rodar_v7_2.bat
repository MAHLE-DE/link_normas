@echo off
setlocal
cd /d "%~dp0"
echo Instalando hotfix V7.2...
copy /Y aplicar_patch_v7_2_hotfix.py aplicar_patch_v7_2.py >nul
if errorlevel 1 goto erro
call rodar_piloto_v7_2.bat
exit /b %errorlevel%
:erro
echo Falha ao instalar o hotfix.
pause
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
echo Lendo o mapa em 100_MAPA e unindo a amostra atual...
py -3 unir_40_normas_por_mapa_v2.py
if errorlevel 1 (
  echo.
  echo Falha. Consulte a mensagem acima.
  pause
  exit /b 1
)
echo.
echo Concluido.
pause

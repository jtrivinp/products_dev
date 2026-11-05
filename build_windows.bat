@echo off
setlocal
cd /d "%~dp0"
echo === Build BPS Compensador ===
python --version
if errorlevel 1 (
  echo ERROR: Python no esta instalado o no esta disponible en PATH.
  echo Instale Python 3.12+ y active la opcion "Add Python to PATH".
  pause
  exit /b 1
)
python -m PyInstaller --version
if errorlevel 1 (
  echo ERROR: PyInstaller no esta instalado.
  echo Ejecute: python -m pip install -r requirements.txt
  pause
  exit /b 1
)
python -m PyInstaller --noconfirm --clean --windowed --name BPS_Compensador --paths src src\bps_compensador\main.py
if errorlevel 1 (
  echo ERROR: El empaquetado no ha terminado correctamente.
  pause
  exit /b 1
)
echo.
echo Build terminado: dist\BPS_Compensador\BPS_Compensador.exe
echo Mantenga config, input, output y logs junto a la carpeta del ejecutable.
pause

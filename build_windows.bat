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
set "CONFIG_BACKUP=%TEMP%\BPS_Compensador_config_backup"
if exist "%CONFIG_BACKUP%" rmdir /S /Q "%CONFIG_BACKUP%"
if exist "dist\BPS_Compensador\config" xcopy /E /I /Y "dist\BPS_Compensador\config" "%CONFIG_BACKUP%" >nul
python -m PyInstaller --noconfirm --clean --windowed --name BPS_Compensador --paths src src\bps_compensador\main.py
if errorlevel 1 (
  echo ERROR: El empaquetado no ha terminado correctamente.
  pause
  exit /b 1
)
if exist "%CONFIG_BACKUP%\config.ini" (
  xcopy /E /I /Y "%CONFIG_BACKUP%" "dist\BPS_Compensador\config" >nul
) else (
  if not exist "dist\BPS_Compensador\config" mkdir "dist\BPS_Compensador\config"
  copy /Y "config\config.ini" "dist\BPS_Compensador\config\config.ini" >nul
  if exist "config\auxiliares" xcopy /E /I /Y "config\auxiliares" "dist\BPS_Compensador\config\auxiliares" >nul
)
if exist "%CONFIG_BACKUP%" rmdir /S /Q "%CONFIG_BACKUP%"
if not exist "dist\BPS_Compensador\config\auxiliares" mkdir "dist\BPS_Compensador\config\auxiliares"
for %%F in ("config\auxiliares\*.xlsx") do if not exist "dist\BPS_Compensador\config\auxiliares\%%~nxF" copy /Y "%%F" "dist\BPS_Compensador\config\auxiliares\%%~nxF" >nul
for %%D in (input output logs) do if not exist "dist\BPS_Compensador\%%D" mkdir "dist\BPS_Compensador\%%D"
echo.
echo Build terminado: dist\BPS_Compensador\BPS_Compensador.exe
echo Mantenga config, input, output y logs junto a la carpeta del ejecutable.
pause

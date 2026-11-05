@echo off
setlocal
python -m PyInstaller --noconfirm --clean --windowed --name BPS_Compensador --paths src src\bps_compensador\main.py
echo Copie config, input, output y logs junto a dist\BPS_Compensador.

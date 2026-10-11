@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Dang tao hinh cho bai bao...
python tao_hinh.py %*
echo.
pause

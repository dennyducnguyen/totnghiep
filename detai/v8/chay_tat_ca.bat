@echo off
REM Chay lai toan bo bai bao v8 (Windows). Can: Python + cac goi trong requirements.txt, Node.js.
REM Lan dau: pip install -r requirements.txt   va   cd scripts ^&^& npm install ^&^& cd ..
chcp 65001 >nul
cd /d "%~dp0"
python scripts\01_phan_tich.py || goto :loi
python scripts\02_ve_hinh.py || goto :loi
python scripts\03_bang_va_so_lieu.py || goto :loi
node scripts\04_dung_word.js || goto :loi
python scripts\05_kiem_tra_trich_dan.py || goto :loi
echo.
echo XONG: Bai-bao-EBI-63-tinh-2017-2024-v8.docx
pause
exit /b 0
:loi
echo CO LOI - xem thong bao phia tren.
pause
exit /b 1

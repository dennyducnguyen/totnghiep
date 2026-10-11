#!/bin/sh
# Chạy lại toàn bộ bài báo v8 (macOS/Linux). Lần đầu: pip install -r requirements.txt && (cd scripts && npm install)
set -e
cd "$(dirname "$0")"
python3 scripts/01_phan_tich.py > /dev/null
python3 scripts/02_ve_hinh.py
python3 scripts/03_bang_va_so_lieu.py
node scripts/04_dung_word.js
python3 scripts/05_kiem_tra_trich_dan.py | tail -2

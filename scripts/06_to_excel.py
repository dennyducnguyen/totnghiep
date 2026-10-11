# -*- coding: utf-8 -*-
"""
Chuyen toan bo CSV trong data/ sang Excel, giu nguyen cau truc thu muc:
  data/2017/*.csv        -> data/excel/2017/*.xlsx
  data/tonghop/*.csv     -> data/excel/tonghop/*.xlsx
  (gom ca cac file _meta.csv)

Noi dung giu nguyen; gia tri so ghi thanh KIEU SO that (khong phai text) de
loc/tinh duoc ngay trong Excel; ky hieu thieu ('..', '-') giu nguyen dang chu.
Dinh dang nhe: font Arial, dam + to nen dong tieu de, co dinh dong dau,
gian rong cot dau.

Cach dung:  python scripts/06_to_excel.py
"""
import csv
import glob
import os
import re

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'data', 'csv')
DST = os.path.join(ROOT, 'data', 'excel')

SUBDIRS = ['2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024',
           'tonghop']
NUM_RE = re.compile(r'-?\d+(?:\.\d+)?$')

FONT = Font(name='Arial', size=10)
FONT_HEAD = Font(name='Arial', size=10, bold=True)
FILL_HEAD = PatternFill('solid', start_color='DDEBF7')


def cell_value(text, is_header):
    """Chuoi so -> so that; con lai giu nguyen chu."""
    if not is_header and text and NUM_RE.fullmatch(text):
        f = float(text)
        return int(f) if f.is_integer() and '.' not in text else f
    return text


def convert(csv_path, xlsx_path):
    rows = list(csv.reader(open(csv_path, encoding='utf-8-sig')))
    wb = Workbook()
    ws = wb.active
    ws.title = 'Data'
    for i, row in enumerate(rows, 1):
        head = (i == 1)
        for j, val in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=cell_value(val, head))
            c.font = FONT_HEAD if head else FONT
            if head:
                c.fill = FILL_HEAD
    ws.freeze_panes = 'A2'
    if rows:
        widths = [0] * len(rows[0])
        for row in rows[:80]:
            for j, val in enumerate(row):
                if j < len(widths):
                    widths[j] = max(widths[j], len(val))
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = min(w + 2, 45)
    wb.save(xlsx_path)


def main():
    total = 0
    for sub in SUBDIRS:
        src_dir = os.path.join(SRC, sub)
        dst_dir = os.path.join(DST, sub)
        if not os.path.isdir(src_dir):
            continue
        os.makedirs(dst_dir, exist_ok=True)
        for f in sorted(glob.glob(os.path.join(src_dir, '*.csv'))):
            name = os.path.basename(f)[:-4] + '.xlsx'
            convert(f, os.path.join(dst_dir, name))
            total += 1
        print(f'{sub}: xong')
    print(f'Da chuyen {total} file CSV -> {DST}')


if __name__ == '__main__':
    main()

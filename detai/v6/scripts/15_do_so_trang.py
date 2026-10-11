# -*- coding: utf-8 -*-
"""15 - Dò số trang của từng mục/bảng/hình/phương trình trong bản in PDF của bài báo.
Đầu vào : ket-qua/Bai-bao-ban-in-PDF-de-tra-so-trang.pdf (xuất từ Word)
Đầu ra  : ket-qua/so_trang_bai_bao.json  (script 13 dùng để ghi "tr. X" trong file hướng dẫn)
Chạy    : python scripts/15_do_so_trang.py
"""
import json, os
import fitz

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = os.path.join(ROOT, 'ket-qua', 'Bai-bao-ban-in-PDF-de-tra-so-trang.pdf')
OUT = os.path.join(ROOT, 'ket-qua', 'so_trang_bai_bao.json')

MUC = ['1. Giới thiệu', '2. Cơ sở lý thuyết', '2.1.', '2.2.', '2.3.', '2.4.', '2.5.', '3. Dữ liệu và phương pháp',
       '3.1.', '3.2.', '3.3.', '4. Kết quả', '4.1.', '4.2.', '4.3.', '4.4.', '4.5.', '4.6.', '5. Thảo luận',
       '6. Kết luận', 'Tài liệu tham khảo']
BANG = [f'Bảng {k}.' for k in ['1', '2', '3', '4', '5', '5b', '6', '7', '8', '8b', '9']]
HINH = [f'Hình {k}.' for k in range(1, 9)]
PT = {'(1)': 'T = Σk sk Tk', '(2)': '(1/T) ln(EBI', '(3)': 'ln(EBIi,t) = β0', '(4)': 'ȲL − ȲN'}


def main():
    d = fitz.open(PDF)
    found = {}
    for i, pg in enumerate(d):
        t = pg.get_text()
        lines = [l.strip() for l in t.split('\n')]
        for k in MUC + BANG + HINH:
            if k not in found and any(l.startswith(k) for l in lines):
                found[k] = i + 1
        for k, v in PT.items():
            if k not in found and v in t.replace('\n', ' '):
                found[k] = i + 1
    thieu = [k for k in MUC + BANG + HINH + list(PT) if k not in found]
    json.dump(found, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'Đã ghi {OUT} | {d.page_count} trang | {len(found)} mục' + (f' | THIẾU: {thieu}' if thieu else ''))


if __name__ == '__main__':
    main()

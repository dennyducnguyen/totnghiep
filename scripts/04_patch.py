# -*- coding: utf-8 -*-
"""
Giai doan 3c-1: Va cac loi do watermark/text an trong PDF goc (khong sua duoc
bang parser). MOI GIA TRI DEU DA DOI CHIEU VISION voi anh trang PDF goc
(2026-07-21, render zoom 2x):

1. NGTK2024 tr.581 (so trang trai): watermark "TONG CUC THONG KE" chen ky tu
   vao nhan "Tuyen Quang" thanh "CBTuaycoe KBna Qnn uga ng". Gia tri dong
   (358 373 335 320 295) khop anh = dong TUYEN QUANG. Cao Bang, Bac Kan
   khong duoc in trong PDF (khong co trang trai) - khong phai loi.
2. NGTK2024 tr.666 (rung trong moi): nhan "Ca Mau" bi chen thanh
   "BCac M Laieuu". Gia tri (3.7 4.2 4.1 4.2 5.3) khop anh = dong CA MAU.
   Bac Lieu khong duoc in trong PDF - khong phai loi.
3. NGTK2024 tr.1107 (chay no): dong text AN '4 905' (trung gia tri Ca Mau
   2023) lan vao text layer -> xoa dong rac. Dong Ca Mau that (7, 2226) da co.
4. NGTK2023 tr.1103 (chay no): dong rac an '4 905' tuong tu -> xoa.
   Them dong "Thai Binh" trong (PDF in nguyen dong trang - khong co so lieu).

Chay SAU 01_extract.py (neu trich lai thi phai chay lai script nay):
  python scripts/04_patch.py
"""
import csv
import glob
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load(pattern):
    f = [x for x in glob.glob(os.path.join(ROOT, pattern)) if '_b2' not in x][0]
    return f, list(csv.reader(open(f, encoding='utf-8-sig')))


def save(f, rows):
    with open(f, 'w', newline='', encoding='utf-8-sig') as fh:
        csv.writer(fh).writerows(rows)


def main():
    n = 0

    # 1. b0581 2024: dong roi -> Tuyen Quang
    f, rows = load(r'data/csv/2024/b0581_so-trang-trai-phan-theo-dia-phuong.csv')
    for r in rows:
        if r[1] == 'KHAC' and r[2:] == ['358', '373', '335', '320', '295']:
            r[0], r[1] = 'Tuyên Quang', 'TINH'
            n += 1
    save(f, rows)

    # 2. b0665 2024: dong roi -> Ca Mau
    f, rows = load(r'data/csv/2024/b0665_dien-tich-rung-trong-moi-tap-trung*.csv')
    for r in rows:
        if r[1] == 'KHAC' and r[2:] == ['3.7', '4.2', '4.1', '4.2', '5.3']:
            r[0], r[1] = 'Cà Mau', 'TINH'
            n += 1
    save(f, rows)

    # 3+4. chay no 2024/2023: xoa dong rac text an ['','KHAC','4','','','905']
    for pat in (r'data/csv/2024/b1106_mot-so-chi-tieu-ve-vu-chay-no*.csv',
                r'data/csv/2023/b1102_mot-so-chi-tieu-ve-vu-chay-no*.csv'):
        f, rows = load(pat)
        before = len(rows)
        rows = [r for r in rows
                if not (r[0] == '' and r[1] == 'KHAC'
                        and r[2:] == ['4', '', '', '905'])]
        n += before - len(rows)
        # 2023: them dong Thai Binh trong (PDF in dong trang), sau Hung Yen
        if '2023' in f and not any(r[0] == 'Thái Bình' for r in rows):
            ncols = len(rows[0]) - 2
            for i, r in enumerate(rows):
                if r[0] == 'Hưng Yên':
                    rows.insert(i + 1, ['Thái Bình', 'TINH'] + [''] * ncols)
                    n += 1
                    break
        save(f, rows)

    print(f'Da va {n} dong (ky vong 5: 2 doi nhan + 2 xoa rac + 1 them Thai Binh).')


if __name__ == '__main__':
    main()

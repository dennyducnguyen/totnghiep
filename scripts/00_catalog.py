# -*- coding: utf-8 -*-
"""
Giai doan 0: Lap danh muc (catalog) tat ca cac bang trong 8 file Nien giam Thong ke.

Nhan dien dua tren font (da khao sat NGTK2017):
  - Tieu de Viet:  Arial,Bold  size ~12
  - Tieu de Anh:   Arial,Italic size ~12
  - "(Tiep theo)": Arial regular size ~12 (hoac tieu de Anh bat dau "(Cont.)")
  - Header chuong: Bold ~10 o dong dau trang
  - Than bang:     size ~8.3

Output: docs/catalog.csv (utf-8-sig), moi dong = 1 bang (gom ca cac trang "Tiep theo").
"""
import csv
import glob
import os
import re

import fitz

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF_DIR = os.path.join(ROOT, 'niengiam')
OUT_CSV = os.path.join(ROOT, 'docs', 'catalog.csv')

TITLE_SIZE_MIN = 11.0   # tieu de bang
TITLE_SIZE_MAX = 16.0   # tren nguong nay coi la trang bia/chia chuong
BODY_SIZE_MAX = 10.0    # span nho hon coi la noi dung bang
MIN_NUM_SPANS = 10      # trang bang that su phai co du o so lieu
                        # (loai trang "GIAI THICH THUAT NGU" chi co chu)

# Chuan hoa ten chuong: (chuoi nhan dien dau tieu, ten chinh thuc)
CHAPTER_MAP = [
    ('Đơn vị Hành chính', 'Đơn vị hành chính, Đất đai và Khí hậu'),
    ('Dân số', 'Dân số và Lao động'),
    ('TKQG', 'Tài khoản quốc gia, Ngân sách nhà nước, Ngân hàng và Bảo hiểm'),
    ('Công nghiệp', 'Công nghiệp, Đầu tư và Xây dựng'),
    ('DN, HTX', 'Doanh nghiệp, Hợp tác xã và Cơ sở kinh tế cá thể phi nông nghiệp'),
    ('Nông, Lâm', 'Nông, Lâm nghiệp và Thủy sản'),
    ('Thương mại', 'Thương mại và Du lịch'),
    ('Chỉ số giá', 'Chỉ số giá'),
    ('Vận tải', 'Vận tải và Bưu chính, Viễn thông'),
    ('Khoa học', 'Khoa học, Công nghệ và Giáo dục'),
    ('Gi¸o dôc', 'Giáo dục'),
    ('Giáo dục', 'Giáo dục'),
    ('Y tÕ', 'Y tế, Văn hóa, Thể thao và Mức sống dân cư'),
    ('Y tế', 'Y tế, Văn hóa, Thể thao và Mức sống dân cư'),
    ('Số liệu', 'Số liệu thống kê nước ngoài'),
]


def canon_chapter(raw: str):
    for key, name in CHAPTER_MAP:
        if raw.startswith(key):
            return name
    return None  # khong nhan dien duoc -> ke thua tu bang truoc


def norm(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip()


def parse_page(page):
    """Tra ve dict mo ta trang: chuong, tieu de VN/EN, co phai trang tiep theo."""
    d = page.get_text('dict')
    spans = []
    for block in d['blocks']:
        for line in block.get('lines', []):
            for sp in line['spans']:
                t = sp['text'].strip()
                if t:
                    spans.append(sp)

    # o so lieu: span co chua chu so, kich thuoc nho (than bang)
    num_count = sum(1 for sp in spans
                    if sp['size'] < BODY_SIZE_MAX
                    and re.search(r'\d', sp['text'])
                    and re.fullmatch(r'[-\d.,%‰…() ]+', sp['text'].strip()))

    title_vn_parts, title_en_parts, cont = [], [], False
    chapter = ''
    for sp in spans:
        size, font, text = sp['size'], sp['font'], sp['text'].strip()
        is_title_size = TITLE_SIZE_MIN <= size <= TITLE_SIZE_MAX
        if is_title_size:
            if text.startswith('(Tiếp theo)') or text == '(Tiếp theo)':
                cont = True
                text = text.replace('(Tiếp theo)', '').strip()
                if not text:
                    continue
            if 'Bold' in font:
                title_vn_parts.append(text)
            elif 'Italic' in font:
                title_en_parts.append(text)
        elif not chapter and 'Bold' in font and 9.0 <= size <= 10.5:
            chapter = text

    title_en = norm(' '.join(title_en_parts))
    if title_en.startswith('(Cont.)'):
        cont = True
        title_en = norm(title_en.replace('(Cont.)', ''))

    return {
        'chuong': norm(chapter),
        'tieu_de_vn': norm(' '.join(title_vn_parts)),
        'tieu_de_en': title_en,
        'tiep_theo': cont,
        'la_trang_bang': num_count >= MIN_NUM_SPANS,
    }


def catalog_file(path):
    """Quet 1 file PDF, tra ve danh sach bang."""
    fname = os.path.basename(path)
    year = re.search(r'(\d{4})', fname).group(1)
    doc = fitz.open(path)
    tables = []
    skipped_cont = 0  # trang "Tiep theo" khong co bang goc dung truoc (bat thuong)
    for i, page in enumerate(doc):
        info = parse_page(page)
        if not info['la_trang_bang'] or not info['tieu_de_vn']:
            continue
        # loai trang giai thich thuat ngu / phuong phap tinh (khong phai bang)
        up = info['tieu_de_vn'].upper()
        if up.startswith('GIẢI THÍCH') or up.startswith('EXPLANATION'):
            continue
        if info['tiep_theo']:
            if tables and tables[-1]['trang_cuoi'] >= i - 3:
                tables[-1]['trang_cuoi'] = i + 1
            else:
                skipped_cont += 1
            continue
        tables.append({
            'nam': year,
            'file': fname,
            'trang_dau': i + 1,
            'trang_cuoi': i + 1,
            'chuong': info['chuong'],
            'tieu_de_vn': info['tieu_de_vn'],
            'tieu_de_en': info['tieu_de_en'],
        })
    doc.close()

    # chuan hoa ten chuong; khong nhan dien duoc thi ke thua tu bang lien truoc
    prev = ''
    for t in tables:
        c = canon_chapter(t['chuong'])
        t['chuong'] = c if c else prev
        prev = t['chuong']
    return tables, skipped_cont


def main():
    all_rows = []
    for path in sorted(glob.glob(os.path.join(PDF_DIR, '*.pdf'))):
        tables, skipped = catalog_file(path)
        n_prov = 0
        for t in tables:
            vn = t['tieu_de_vn'].lower()
            en = t['tieu_de_en'].lower()
            t['dia_phuong'] = 1 if ('theo địa phương' in vn or 'by province' in en) else 0
            t['trang_thai'] = 'chua_lam'
            n_prov += t['dia_phuong']
        all_rows.extend(tables)
        print(f"{os.path.basename(path)}: {len(tables)} bang, trong do {n_prov} bang theo dia phuong"
              + (f" | CANH BAO: {skipped} trang 'Tiep theo' mo coi" if skipped else ''))

    os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
    cols = ['nam', 'file', 'trang_dau', 'trang_cuoi', 'chuong',
            'tieu_de_vn', 'tieu_de_en', 'dia_phuong', 'trang_thai']
    with open(OUT_CSV, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(all_rows)
    print(f"\nDa ghi {len(all_rows)} dong vao {OUT_CSV}")


if __name__ == '__main__':
    main()

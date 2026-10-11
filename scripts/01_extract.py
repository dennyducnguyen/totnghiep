# -*- coding: utf-8 -*-
"""
Giai doan 1: Trich xuat cac bang "phan theo dia phuong" tu Nien giam Thong ke.

Phuong phap (da khao sat): bang khong ke du duong -> khong dung extract_tables.
Dung page.extract_words(extra_attrs=['size']) lay tung tu kem toa do + co chu:
  - co chu < 10  -> noi dung bang (tieu de/header chuong co chu lon, tu loai)
  - gom dong theo toa do 'top'; gom cot theo canh phai (x1) cua o so (can phai)
  - nhan dong (ten tinh/vung) nam ben trai cot so dau tien
  - dong nhan tieng Viet khong so + dong so ke tiep mang nhan EN -> lay nhan VN

Cach dung:
  python scripts/01_extract.py 2017          # trich cac bang dia_phuong=1 nam 2017
  python scripts/01_extract.py 2017 101      # chi trich bang bat dau o trang 101 (debug)

Output: data/<nam>/bXXXX_<slug>.csv ; data/<nam>/_meta.csv ; cap nhat trang_thai catalog.
"""
import csv
import os
import re
import sys

import pdfplumber

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vn_names import classify_label, slugify

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CATALOG = os.path.join(ROOT, 'docs', 'catalog.csv')

BODY_SIZE_MAX = 10.0     # chu than bang ~8.3; tieu de 12, header chuong 10 -> loai
ROW_TOL = 5.0            # gom dong: chenh lech top toi da (dong cach nhau ~13pt;
                         # co o bi dat lech ~3pt do loi dan trang cua PDF)
COL_TOL = 6.0            # gom cot: chenh lech x1 toi da
NUM_RE = re.compile(r'-?\d[\d.]*(?:,\d+)?\*?')   # 1.234,5  10,2  2010  5,8*
MISSING_TOKENS = {'..', '...', '-', '…'}


def to_number_str(tok: str) -> str:
    """'1.234,5' -> '1234.5' (kieu quoc te). Giu nguyen token thieu ('..')."""
    t = tok.strip().rstrip('*')
    if t in MISSING_TOKENS or not t:
        return t
    return t.replace('.', '').replace(',', '.')


def cluster(values, tol):
    """Gom 1 chieu: tra ve danh sach tam cum (trung binh), sap tang dan."""
    centers = []
    for v in sorted(values):
        if centers and v - centers[-1][-1] <= tol:
            centers[-1].append(v)
        else:
            centers.append([v])
    return [sum(c) / len(c) for c in centers]


def group_rows(words):
    """Gom tu thanh dong: gom cum toa do top roi gan tu vao cum gan nhat.
    (Khong noi day chuyen theo tu truoc do - tranh 1 o lech vai pt keo
    ca dong dinh vao dong tren.)"""
    centers = cluster([w['top'] for w in words], ROW_TOL)
    buckets = {c: [] for c in centers}
    for w in words:
        c = min(centers, key=lambda c: abs(c - w['top']))
        buckets[c].append(w)
    return [sorted(buckets[c], key=lambda w: w['x0']) for c in sorted(buckets)]


def parse_table_page(page):
    """Phan tich 1 trang bang -> (headers, data_rows, unit).
    headers: list nhan cot. data_rows: list (label, [gia tri theo cot])."""
    words = [w for w in page.extract_words(extra_attrs=['size'])
             if w['size'] < BODY_SIZE_MAX]
    if not words:
        return None, [], ''

    rows = group_rows(words)

    # cat bo phan chu thich cuoi trang ((*)..., Nguon:...) - so lieu trong
    # chu thich lam ban cum cot. Chi cat khi o NUA DUOI trang va sat le trai
    # (dau (*) superscript cua tieu de cung size nho nhung nam dau trang/ben phai)
    half_h = page.height / 2
    cut = len(rows)
    for i, r in enumerate(rows):
        first = r[0]
        if first['top'] > half_h and first['x0'] < 200 and \
                (first['text'].startswith('(*') or first['text'].startswith('Nguồn')):
            cut = i
            break
    rows = rows[:cut]

    # dong du lieu dau tien: nhan dien theo NHAN (CA NUOC / vung / tinh) thay
    # vi toa do, vi header cot co the chua so ("Duoi 0,5 ty dong") va keo dai
    # sang trai. Nhan = cac tu truoc o so dau tien cua dong.
    def row_label_prefix(r):
        parts = []
        for w in r:
            if NUM_RE.fullmatch(w['text']) or w['text'] in MISSING_TOKENS:
                break
            parts.append(w['text'])
        return ' '.join(parts)

    def row_has_val(r):
        return any(NUM_RE.fullmatch(w['text']) or w['text'] in MISSING_TOKENS
                   for w in r)

    def is_value_only_row(r):
        # dong chi co gia tri, khong co chu o vung nhan ben trai
        return row_has_val(r) and not any(w['x0'] < 200 for w in r)

    first_data_idx = None
    for i, r in enumerate(rows):
        # nhan phai nhan dien duoc (header cot cung co the mang chu "Tong so"
        # viet thuong nhung se khong duoc phan loai)
        if classify_label(row_label_prefix(r))[0] == 'KHAC':
            continue
        if row_has_val(r):
            first_data_idx = i
            break
        # dong CHI CO NHAN (vd "CA NUOC" / "WHOLE COUNTRY" tach dong voi gia
        # tri - gap o NGTK2018): chap nhan neu 1-2 dong sau la dong thuan gia tri
        if any(is_value_only_row(rows[j]) for j in (i + 1, i + 2)
               if j < len(rows)):
            first_data_idx = i
            break
    if first_data_idx is None:
        return None, [], ''

    # cum cot: canh phai (x1) cua cac o so TRONG VUNG DU LIEU; giu cum du day
    # de loai so le con sot (nguong thich nghi voi so dong cua trang)
    data_zone = rows[first_data_idx:]
    num_words = [w for r in data_zone for w in r if NUM_RE.fullmatch(w['text'])]
    if not num_words:
        return None, [], ''
    raw_centers = cluster([w['x1'] for w in num_words], COL_TOL)

    def nearest(centers, x):
        return min(centers, key=lambda c: abs(c - x))

    counts = {c: 0 for c in raw_centers}
    for w in num_words:
        counts[nearest(raw_centers, w['x1'])] += 1
    min_count = min(5, max(2, len(data_zone) // 3))
    col_x = [c for c in raw_centers if counts[c] >= min_count]
    if not col_x:
        return None, [], ''
    core_nums = [w for w in num_words if abs(nearest(col_x, w['x1']) - w['x1']) <= COL_TOL]
    strong_left = min(w['x0'] for w in core_nums) - 4.0

    # Ranh gioi nhan: dua vao KHOANG TRONG lon giua khoi ten (trai) va khoi so.
    # (Neu chi dua strong_left, cot so THUA nhu "Tong so" trong bang giuong benh
    # se bi rot va gia tri dinh vao ten tinh -> "Bac Ninh 4518".)
    # Tim khe ngang >=25pt dau tien, chi xet vung ben trai cot so manh nhat.
    x0s = sorted({round(w['x0']) for r in rows[first_data_idx:] for w in r
                  if w['x0'] < strong_left})
    label_boundary = strong_left
    for a, b in zip(x0s, x0s[1:]):
        if b - a >= 25:
            label_boundary = (a + b) / 2
            break

    # phuc hoi cot THUA dan dau (vd "Tong so" chi vai tinh co gia tri): cum so
    # nam giua ranh gioi nhan va cot manh nhat, co >=2 gia tri
    leading = [c for c in raw_centers
               if label_boundary < c < min(col_x) and counts[c] >= 2]
    if leading:
        col_x = sorted(set(col_x) | set(leading))

    def col_of(w):
        best, bd = None, 1e9
        for j, cx in enumerate(col_x):
            d = abs(w['x1'] - cx)
            if d < bd:
                best, bd = j, d
        return best

    # header: cac dong truoc dong du lieu dau tien
    header_rows = rows[:first_data_idx]
    # dong don vi tinh: dong header dau tien, KHONG chua nam (2010...) va
    # (ngan, hoac co dau '/' hay ' - ' — vd "Tre em duoi mot tuoi tu vong/1000
    # tre em sinh song", "Trieu dong - Mill. dongs")
    # (co the 2 dong: Viet + Anh -> gom lien tiep)
    unit_parts = []
    while header_rows:
        joined = ' '.join(w['text'] for w in header_rows[0]).strip()
        has_year = re.search(r'\b(19|20)\d{2}\b', joined)
        # khong nuot nhan cot ("So bo"/"Prel."/"Uoc tinh" nam tren dong nam)
        if joined in ('Sơ bộ', 'Prel.', 'Sơ bộ Prel.', 'Ước tính', 'Est.',
                      'Ước tính Est.'):
            break
        if not has_year and (len(joined) <= 40 or '/' in joined
                             or ' - ' in joined):
            unit_parts.append(joined)
            header_rows = header_rows[1:]
        else:
            break
    unit = ' | '.join(unit_parts)
    # tu header thuong can giua cot -> gan theo TAM tu so voi dai cot
    # (dai cot j = (canh phai cot j-1, canh phai cot j])
    def header_col(w):
        c = (w['x0'] + w['x1']) / 2
        for j, cx in enumerate(col_x):
            if c <= cx + 2:
                return j
        return len(col_x) - 1

    header_parts = [[] for _ in col_x]
    for r in header_rows:
        for w in r:
            if w['x1'] > label_boundary - 20:  # bo tu thuoc cot nhan ("Tinh, TP")
                header_parts[header_col(w)].append(w['text'])
    headers = [' '.join(p).strip() for p in header_parts]

    # du lieu
    data_rows = []
    pending_label = None
    for r in rows[first_data_idx:]:
        label_words = [w for w in r if w['x0'] < label_boundary]
        value_words = [w for w in r if w['x0'] >= label_boundary]
        label = ' '.join(w['text'] for w in label_words).strip()
        # o gia tri: so hoac ky hieu thieu ('..', '-')
        nums = [w for w in value_words
                if NUM_RE.fullmatch(w['text']) or w['text'] in MISSING_TOKENS]

        if label and not nums:
            # dong chi co nhan (ten vung tieng Viet dong tren, nhan dai xuong dong)
            pending_label = f'{pending_label} {label}' if pending_label else label
            continue
        if not nums:
            continue
        vals = [''] * len(col_x)
        for w in nums:
            vals[col_of(w)] = to_number_str(w['text'])
        # chon nhan: thu (1) nhan dong nay, (2) nhan VN dong tren (pending),
        # (3) ghep pending + nhan dong nay (vung song ngu xuong dong nhieu lan)
        candidates = [label, pending_label,
                      f'{pending_label} {label}'.strip() if pending_label else None]
        use_label = None
        for c in candidates:
            if c and classify_label(c)[0] != 'KHAC':
                use_label = c
                break
        if use_label is None:
            use_label = pending_label if pending_label else label
        pending_label = None
        # bo dong rac: nhan khong nhan dien duoc va toan o trong/thieu
        if classify_label(use_label)[0] == 'KHAC' and \
                all(v in ('', '..', '...', '-', '…') for v in vals):
            continue
        data_rows.append((use_label, vals))
    return headers, data_rows, unit


def extract_table(pdf, trang_dau, trang_cuoi):
    """Trich 1 bang (co the nhieu trang). Tra ve (headers, rows, unit, warnings).

    Bang "trai ngang": danh sach tinh lap lai tu dau (CA NUOC xuat hien lan 2)
    cho nhom cot tiep theo -> tach thanh cac phan doc roi ghep ngang theo nhan.
    """
    warns, unit = [], ''
    segments = []  # moi phan: {'headers': [...], 'rows': [(label, vals)]}
    for p in range(trang_dau - 1, trang_cuoi):
        headers, data_rows, u = parse_table_page(pdf.pages[p])
        if headers is None:
            warns.append(f'trang {p+1}: khong doc duoc')
            continue
        if not unit:
            unit = u
        starts_ca_nuoc = data_rows and classify_label(data_rows[0][0])[0] == 'CA_NUOC'
        if not segments or (starts_ca_nuoc and any(
                classify_label(l)[0] == 'CA_NUOC' for l, _ in segments[-1]['rows'])):
            segments.append({'headers': headers, 'rows': []})
        elif len(headers) != len(segments[-1]['headers']):
            warns.append(f'trang {p+1}: so cot {len(headers)} khac trang dau phan '
                         f'{len(segments[-1]["headers"])}')
        segments[-1]['rows'].extend(data_rows)

    if not segments:
        return None, [], unit, warns, []
    if len(segments) == 1:
        rows_out = segments[0]['rows']
        # bang xep chong doc tren cung 1 trang: dong CA NUOC xuat hien lan 2 ->
        # tach thanh phan phu (bang thu 2 cung trang, chua co tieu de rieng)
        ca_idx = [i for i, (l, _) in enumerate(rows_out)
                  if classify_label(l)[0] == 'CA_NUOC']
        if len(ca_idx) >= 2:
            cut = ca_idx[1]
            warns.append(f'2 bang xep chong cung trang: tach o dong {cut} '
                         f'-> file phu _b2')
            extra = [(segments[0]['headers'], rows_out[cut:], unit)]
            return segments[0]['headers'], rows_out[:cut], unit, warns, extra
        return segments[0]['headers'], rows_out, unit, warns, []

    # ghep ngang: khoa = (nhan chuan, lan xuat hien thu n trong phan)
    warns.append(f'bang trai ngang: ghep {len(segments)} phan cot theo ten tinh')
    all_headers, all_rows = [], []
    key_order, row_map = [], {}
    for si, seg in enumerate(segments):
        w = len(seg['headers'])
        seen = {}
        for label, vals in seg['rows']:
            _, ten = classify_label(label)
            seen[ten] = seen.get(ten, 0) + 1
            key = (ten, seen[ten])
            if key not in row_map:
                row_map[key] = {'label': label, 'segs': {}}
                key_order.append(key)
            row_map[key]['segs'][si] = vals
        all_headers.extend(h if si == 0 else f'{h} ({si+1})' for h in seg['headers'])

    seg_widths = [len(s['headers']) for s in segments]
    for key in key_order:
        vals = []
        for si, wd in enumerate(seg_widths):
            vals.extend(row_map[key]['segs'].get(si, [''] * wd))
        all_rows.append((row_map[key]['label'], vals))
    return all_headers, all_rows, unit, warns, []


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else '2017'
    only_page = int(sys.argv[2]) if len(sys.argv) > 2 else None

    rows = list(csv.DictReader(open(CATALOG, encoding='utf-8-sig')))
    targets = [r for r in rows if r['nam'] == year and r['dia_phuong'] == '1']
    if only_page:
        targets = [r for r in targets if int(r['trang_dau']) == only_page]
    if not targets:
        print('Khong co bang nao khop dieu kien.')
        return

    pdf_path = os.path.join(ROOT, 'niengiam', targets[0]['file'])
    out_dir = os.path.join(ROOT, 'data', 'csv', year)
    os.makedirs(out_dir, exist_ok=True)

    meta_rows, n_ok, n_err = [], 0, 0
    with pdfplumber.open(pdf_path) as pdf:
        for t in targets:
            p1, p2 = int(t['trang_dau']), int(t['trang_cuoi'])
            name = f"b{p1:04d}_{slugify(t['tieu_de_vn'])}"
            try:
                headers, data, unit, warns, extra = extract_table(pdf, p1, p2)
                if not headers or not data:
                    raise ValueError('khong tim thay du lieu')

                def write_csv(fp, hdrs, rows_data):
                    nt = 0
                    out = []
                    for label, vals in rows_data:
                        loai, ten = classify_label(label)
                        nt += (loai == 'TINH')
                        out.append([ten, loai] + vals)
                    with open(fp, 'w', newline='', encoding='utf-8-sig') as f:
                        w = csv.writer(f)
                        w.writerow(['chi_muc', 'loai'] + hdrs)
                        w.writerows(out)
                    return nt

                n_tinh = write_csv(os.path.join(out_dir, name + '.csv'),
                                   headers, data)
                # bang phu (xep chong cung trang) -> file _b2, _b3...
                for k, (h2, r2, u2) in enumerate(extra, start=2):
                    write_csv(os.path.join(out_dir, f'{name}_b{k}.csv'), h2, r2)
                status = 'da_trich'
                n_ok += 1
                note = '; '.join(warns)
            except Exception as e:
                status, note, n_tinh, unit = 'loi', str(e), 0, ''
                n_err += 1
            t['trang_thai'] = status
            meta_rows.append({
                'file_csv': name + '.csv', 'trang_dau': p1, 'trang_cuoi': p2,
                'tieu_de_vn': t['tieu_de_vn'], 'don_vi': unit, 'so_tinh': n_tinh,
                'trang_thai': status, 'ghi_chu': note,
            })
            if status == 'loi' or note:
                print(f"  [{status}] tr.{p1}: {t['tieu_de_vn'][:60]} | {note}")

    with open(os.path.join(out_dir, '_meta.csv'), 'w', newline='',
              encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(meta_rows[0].keys()))
        w.writeheader()
        w.writerows(meta_rows)

    if not only_page:  # ghi nguoc trang_thai vao catalog
        with open(CATALOG, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    print(f"\nXong {year}: {n_ok} bang OK, {n_err} bang loi. Output: {out_dir}")


if __name__ == '__main__':
    main()

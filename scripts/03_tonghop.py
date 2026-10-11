# -*- coding: utf-8 -*-
"""
Giai doan 3a: Noi chuoi cac chi tieu trung ten qua cac nam -> data/tonghop/.

Quy uoc da chot voi nguoi dung (2026-07-21):
  - Dinh dang WIDE: moi dong 1 tinh, moi cot 1 nam (bang kieu cot-nam, type A).
    Bang kieu cot-danh-muc (type B, vd giuong benh Tong so/Cong lap): chong cac
    nam len nhau, them cot 'nam' dau tien.
  - Trung nam giua cac nien giam: lay NIEN GIAM MOI NHAT co gia tri khong thieu
    (so da dieu chinh); thieu thi lui ve nien giam cu hon.
  - Chi noi tu dong cac chi tieu TRUNG TEN (sau chuan hoa bo so trang/nam).
    Manh dut do doi ten -> de buoc 3b (bang anh xa duyet tay).

Ten file: <slug>_tu-<nam min>-<nam max>.csv
Output:   data/tonghop/*.csv + data/tonghop/_meta.csv
"""
import csv
import glob
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vn_names import PROVINCES_63, REGIONS_6

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'data', 'csv', 'tonghop')
YEARS = ['2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024']
MISSING = ('', '..', '...', '-', '…')

# thu tu dong chuan: CA NUOC -> vung -> cac tinh cua vung
_REG_SIZES = [11, 14, 14, 5, 6, 13]
ROW_ORDER = ['CẢ NƯỚC']
_i = 0
for reg, n in zip(REGIONS_6, _REG_SIZES):
    ROW_ORDER.append(reg)
    ROW_ORDER.extend(PROVINCES_63[_i:_i + n])
    _i += n
ROW_RANK = {name: i for i, name in enumerate(ROW_ORDER)}


# Anh xa doi ten bang giua cac nien giam (nguoi dung duyet 2026-07-21,
# chi tiet + kiem chung gia tri: docs/anh-xa-doi-ten.md).
# Khoa = slug cu (2017-2021), gia tri = slug moi (2022+, ten hien hanh cua TCTK).
RENAME_MAP = {
    'ty-le-tang-tu-nhien-dan-so-phan-theo-dia-phuong':
        'ty-le-tang-dan-so-tu-nhien-phan-theo-dia-phuong',
    'dien-tich-cay-luong-thuc-co-hat-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-cay-luong-thuc-co-hat-phan-theo-dia-phuong',
    'dien-tich-lua-ca-nam-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-lua-ca-nam-phan-theo-dia-phuong',
    'nang-suat-lua-ca-nam-phan-theo-dia-phuong':
        'nang-suat-gieo-trong-lua-ca-nam-phan-theo-dia-phuong',
    'dien-tich-lua-dong-xuan-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-lua-dong-xuan-phan-theo-dia-phuong',
    'nang-suat-lua-dong-xuan-phan-theo-dia-phuong':
        'nang-suat-gieo-trong-lua-dong-xuan-phan-theo-dia-phuong',
    'dien-tich-lua-he-thu-va-thu-dong-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-lua-he-thu-va-thu-dong-phan-theo-dia-phuong',
    'nang-suat-lua-he-thu-va-thu-dong-phan-theo-dia-phuong':
        'nang-suat-gieo-trong-lua-he-thu-va-thu-dong-phan-theo-dia-phuong',
    'dien-tich-lua-mua-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-lua-mua-phan-theo-dia-phuong',
    'nang-suat-lua-mua-phan-theo-dia-phuong':
        'nang-suat-gieo-trong-lua-mua-phan-theo-dia-phuong',
    'dien-tich-ngo-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-ngo-phan-theo-dia-phuong',
    'nang-suat-ngo-phan-theo-dia-phuong':
        'nang-suat-gieo-trong-ngo-phan-theo-dia-phuong',
    'dien-tich-khoai-lang-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-khoai-lang-phan-theo-dia-phuong',
    'dien-tich-san-phan-theo-dia-phuong':
        'dien-tich-gieo-trong-san-phan-theo-dia-phuong',
}


def slug_norm(base):
    s = re.sub(r'^b\d+_', '', base).replace('.csv', '')
    s = re.sub(r'(19|20)\d{2}', 'X', s)
    # bo phan NGAY THANG sau cac moc thoi gian nhung GIU phan duoi phan biet
    # ("...tai-thoi-diem-31-12-X-phan-theo-quy-mo-von" -> giu "phan-theo-quy-mo-von";
    # cat het den cuoi se lam 2 bang khac nhau va cham cung slug -> mat du lieu)
    s = re.sub(r'-(tinh-den|tai-thoi-diem)(?:-[0-9X.]+)+', r'-\1', s)
    s = re.sub(r'-nam-hoc(?:-[0-9X.]+)+', '-nam-hoc', s)
    s = re.sub(r'-nam-X(?=-|$)', '', s)
    s = s.strip('-')
    return RENAME_MAP.get(s, s)


def out_name(slug, y1, y2):
    s = re.sub(r'-nam-X(?=-|$)', '', slug)
    s = re.sub(r'-X(?=-|$)', '', s).strip('-')
    if y1 == y2:
        return f'{s}_nam-{y1}.csv'
    return f'{s}_tu-{y1}-{y2}.csv'


def year_cols_of(header):
    """Cot nao la cot NAM don: header chua DUNG MOT nam duy nhat.
    Cot ty le "Nam 2018 so voi nam 2017 (%)" chua 2 nam -> KHONG phai cot nam."""
    out = {}
    for j, h in enumerate(header[2:], 2):
        m = set(re.findall(r'\b((?:19|20)\d{2})\b', h))
        if len(m) == 1:
            out[j] = m.pop()
    return out


def has_dup_year_cols(header):
    """Bang co >=2 cot cung tro ve 1 nam = bang nhieu KHOI KHAI NIEM
    (vd ho ngheo 2 chuan, HIV moi/luy ke) -> KHONG duoc noi kieu A."""
    ys = list(year_cols_of(header).values())
    return len(ys) != len(set(ys))


def load_files():
    """slug -> {nam nien giam: (path, header, rows)}"""
    idx = defaultdict(dict)
    for y in YEARS:
        for f in glob.glob(os.path.join(ROOT, 'data', 'csv', y, 'b*.csv')):
            base = os.path.basename(f)
            rows = list(csv.reader(open(f, encoding='utf-8-sig')))
            slug = slug_norm(base)
            if base.endswith('_b2.csv'):
                # bang thu 2 xep chong cung trang (hien chi co: chi so phat
                # trien san luong che bup 2021, gia tri dang chi so ~100)
                slug = 'chi-so-phat-trien-' + slug.replace('-b2', '').replace('_b2', '')
            # va cham slug trong CUNG nam (2 bang khac nhau bi trung ten do
            # ten file cat o 60 ky tu) -> danh hau to thu tu; thu tu bang
            # trong nien giam on dinh giua cac nam nen chuoi van noi dung
            base_slug, n = slug, 2
            while y in idx[slug]:
                slug = f'{base_slug}--{n}'
                n += 1
            if slug != base_slug:
                print(f'  (va cham ten: {y}/{base} -> {slug})')
            idx[slug][y] = (base, rows[0], rows[1:])
    return idx


def load_meta():
    """(nam, file_csv) -> dict meta"""
    m = {}
    for y in YEARS:
        p = os.path.join(ROOT, 'data', 'csv', y, '_meta.csv')
        for r in csv.DictReader(open(p, encoding='utf-8-sig')):
            m[(y, r['file_csv'])] = r
    return m


def merge_type_a(files):
    """Wide: dong = chi_muc, cot = nam du lieu. files: {nam_ng: (base, header, rows)}
    Tra ve (data_years_sorted, {label: {nam: gia_tri}}, {label: loai})."""
    cell = {}     # (label, nam du lieu) -> (nam nien giam, gia tri)
    loai_of = {}
    for y in sorted(files):          # nien giam cu truoc, moi sau
        base, header, rows = files[y]
        ycols = year_cols_of(header)
        for r in rows:
            label, loai = r[0], r[1]
            loai_of.setdefault(label, loai)
            for j, dy in ycols.items():
                if j >= len(r):
                    continue
                v = r[j]
                key = (label, dy)
                if v not in MISSING:
                    cell[key] = v            # moi nhat co gia tri -> thang
                elif key not in cell:
                    cell[key] = v            # giu '..' neu chua co gi tot hon
    data_years = sorted({dy for (_, dy) in cell})
    table = defaultdict(dict)
    for (label, dy), v in cell.items():
        table[label][dy] = v
    return data_years, table, loai_of


def merge_type_b(files):
    """Chong nam: moi nien giam 1 khoi, them cot 'nam'. Header lay nam moi nhat.
    Chi noi khi so cot bang nhau."""
    ncols = {y: len(files[y][1]) for y in files}
    if len(set(ncols.values())) != 1:
        return None
    newest = max(files)
    header = files[newest][1]
    out_rows = []
    for y in sorted(files):
        for r in files[y][2]:
            out_rows.append([y] + r)
    return ['nam'] + header, out_rows


def sort_labels(labels):
    return sorted(labels, key=lambda l: (ROW_RANK.get(l, 999), l))


def main():
    os.makedirs(OUT, exist_ok=True)
    idx = load_files()
    metas = load_meta()

    meta_out = []
    n_a = n_b = n_skip = 0
    singles, b_fail = [], []

    for slug in sorted(idx):
        files = idx[slug]
        if len(files) < 2:
            singles.append(slug)   # van xu ly nhu binh thuong, chi de thong ke

        newest = max(files)
        base_new = files[newest][0]
        src_meta = metas.get((newest, base_new), {})
        tieu_de = src_meta.get('tieu_de_vn', '')
        don_vi = src_meta.get('don_vi', '')
        nguon = '; '.join(f'{y}:{files[y][0]}' for y in sorted(files))

        # type A neu MOI nam deu co >=2 cot nam don VA khong co bang nao chua
        # 2 cot cung nam (2 khoi khai niem -> phai giu nguyen cot, kieu B)
        dup_concept = any(has_dup_year_cols(files[y][1]) for y in files)
        is_a = (not dup_concept
                and all(len(year_cols_of(files[y][1])) >= 2 for y in files))

        if is_a:
            data_years, table, loai_of = merge_type_a(files)
            fname = out_name(slug, data_years[0], data_years[-1])
            with open(os.path.join(OUT, fname), 'w', newline='',
                      encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(['chi_muc', 'loai'] + data_years)
                for label in sort_labels(table):
                    w.writerow([label, loai_of.get(label, 'KHAC')]
                               + [table[label].get(dy, '') for dy in data_years])
            n_a += 1
            meta_out.append({
                'file_csv': fname, 'kieu': 'A_cot_nam', 'tieu_de_vn': tieu_de,
                'don_vi': don_vi, 'nam_dau': data_years[0],
                'nam_cuoi': data_years[-1],
                'cac_nam_co': ' '.join(data_years),
                'so_nien_giam': len(files), 'nguon': nguon,
                'ghi_chu': 'chi co trong 1 nien giam' if len(files) == 1 else '',
            })
        else:
            merged = merge_type_b(files)
            y1, y2 = min(files), max(files)
            if merged is None:
                # cau truc cot doi giua cac nam. CHIA NHOM trong chuoi:
                #  (a) cac file du dieu kien cot-nam (>=2 cot nam don, khong
                #      trung nam) -> noi wide voi nhau
                #  (b) phan con lai: gom cac nam co HEADER GIONG HET -> chong
                #      khoi; khac header -> tach rieng tung nam
                b_fail.append((slug, {y: len(files[y][1]) for y in files}))
                a_years = {y for y in files
                           if len(year_cols_of(files[y][1])) >= 2
                           and not has_dup_year_cols(files[y][1])}
                sub_a = {y: files[y] for y in a_years}
                rest = {y: files[y] for y in files if y not in a_years}

                if len(sub_a) >= 1:
                    data_years, table, loai_of = merge_type_a(sub_a)
                    fname = out_name(slug, data_years[0], data_years[-1])
                    with open(os.path.join(OUT, fname), 'w', newline='',
                              encoding='utf-8-sig') as f:
                        w = csv.writer(f)
                        w.writerow(['chi_muc', 'loai'] + data_years)
                        for label in sort_labels(table):
                            w.writerow([label, loai_of.get(label, 'KHAC')]
                                       + [table[label].get(dy, '')
                                          for dy in data_years])
                    newest_a = max(sub_a)
                    m_a = metas.get((newest_a, sub_a[newest_a][0]), {})
                    meta_out.append({
                        'file_csv': fname, 'kieu': 'A_cot_nam',
                        'tieu_de_vn': m_a.get('tieu_de_vn', tieu_de),
                        'don_vi': m_a.get('don_vi', ''),
                        'nam_dau': data_years[0], 'nam_cuoi': data_years[-1],
                        'cac_nam_co': ' '.join(data_years),
                        'so_nien_giam': len(sub_a),
                        'nguon': '; '.join(f'{y}:{sub_a[y][0]}'
                                           for y in sorted(sub_a)),
                        'ghi_chu': 'phan cot-nam cua chuoi doi cau truc',
                    })
                    n_a += 1

                # gom phan con lai: cung so cot + cung TUI TU cua toan header
                # (tu ngu giong het nhung co the roi sang cot ke do ngat dong
                # "trang/trai"; bo dau "thuỷ/thủy", bo nam in trong header)
                def hkey(header):
                    from vn_names import norm_key
                    joined = re.sub(r'\b(19|20)\d{2}\b', 'X', ' '.join(header))
                    bag = sorted(norm_key(w) for w in joined.split() if w.strip())
                    return f'{len(header)}|' + '|'.join(bag)

                groups = defaultdict(dict)
                for y in rest:
                    groups[hkey(rest[y][1])][y] = rest[y]
                for g in groups.values():
                    gy1, gy2 = min(g), max(g)
                    if len(g) >= 2:
                        header_g, rows_g = merge_type_b(g)
                        # header khoi chong lay tu nien giam moi nhat - xoa nam
                        # cu the ke phia sau ("luy ke den 31/12/2024" ap cho ca
                        # khoi 2018 la sai; cot 'nam' da cho biet nam)
                        header_g = [re.sub(r'\b(19|20)\d{2}\b', '…', h)
                                    for h in header_g]
                        fname = out_name(slug, gy1, gy2)
                        with open(os.path.join(OUT, fname), 'w', newline='',
                                  encoding='utf-8-sig') as f:
                            w = csv.writer(f)
                            w.writerow(header_g)
                            w.writerows(rows_g)
                        kieu = 'B_chong_nam'
                        note = 'khoi cung cau truc cua chuoi doi cau truc'
                        gap = [str(v) for v in range(int(gy1), int(gy2) + 1)
                               if str(v) not in g]
                        if gap:
                            note += f'; KHONG co nam {" ".join(gap)}'
                        nguon_g = '; '.join(f'{y}:{g[y][0]}' for y in sorted(g))
                        ng = len(g)
                    else:
                        y = gy1
                        base_y, header_y, rows_y = g[y]
                        fname = out_name(slug, y, y)
                        with open(os.path.join(OUT, fname), 'w', newline='',
                                  encoding='utf-8-sig') as f:
                            w = csv.writer(f)
                            w.writerow(header_y)
                            w.writerows(rows_y)
                        kieu, note = 'C_tach_nam', ('cau truc rieng, khong khop '
                                                    'khoi nao - xu tay neu can')
                        nguon_g, ng = f'{y}:{g[y][0]}', 1
                    m_g = metas.get((max(g), g[max(g)][0]), {})
                    meta_out.append({
                        'file_csv': fname, 'kieu': kieu,
                        'tieu_de_vn': m_g.get('tieu_de_vn', tieu_de),
                        'don_vi': m_g.get('don_vi', ''), 'nam_dau': gy1,
                        'nam_cuoi': gy2, 'cac_nam_co': ' '.join(sorted(g)),
                        'so_nien_giam': ng, 'nguon': nguon_g, 'ghi_chu': note,
                    })
                n_skip += 1
                continue
            header, rows = merged
            fname = out_name(slug, y1, y2)
            with open(os.path.join(OUT, fname), 'w', newline='',
                      encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(header)
                w.writerows(rows)
            n_b += 1
            meta_out.append({
                'file_csv': fname, 'kieu': 'B_chong_nam', 'tieu_de_vn': tieu_de,
                'don_vi': don_vi, 'nam_dau': y1, 'nam_cuoi': y2,
                'cac_nam_co': ' '.join(sorted(files)),
                'so_nien_giam': len(files), 'nguon': nguon,
                'ghi_chu': ('bang goc co nhieu khoi khai niem cung nam '
                            '- giu nguyen cot, moi khoi = 1 nien giam'
                            if dup_concept else
                            ('chi co trong 1 nien giam' if len(files) == 1
                             else 'moi khoi = so lieu cua 1 nien giam')),
            })

    # --- xu ly rieng TY LE HO NGHEO: tach 2 chuan ngheo ---
    # (a) chuoi A tu 2019+ toan cot nam = theo CHUAN NGHEO DA CHIEU -> doi ten
    for m in meta_out:
        if (m['kieu'] == 'A_cot_nam'
                and m['file_csv'].startswith('ty-le-ho-ngheo-phan-theo-dia-phuong_tu-')):
            new = m['file_csv'].replace(
                'ty-le-ho-ngheo-phan-theo-dia-phuong',
                'ty-le-ho-ngheo-theo-chuan-ngheo-da-chieu')
            os.replace(os.path.join(OUT, m['file_csv']), os.path.join(OUT, new))
            m['file_csv'] = new
            m['ghi_chu'] = ('chuoi theo CHUAN NGHEO DA CHIEU (ap dung tu 2016); '
                            'chuan cu xem file ...theo-chuan-chinh-phu')
    # (b) chuan CHINH PHU cu (2010-2016): lay 4 cot dau bang ho ngheo NGTK2017
    #     (header: "Theo chuan ngheo cua Chinh phu": 2010, 2014, 2015, 2016)
    src = glob.glob(os.path.join(ROOT, 'data', 'csv', '2017', 'b0837_ty-le-ho-ngheo*.csv'))
    if src:
        rows_hn = list(csv.reader(open(src[0], encoding='utf-8-sig')))
        fname = 'ty-le-ho-ngheo-theo-chuan-chinh-phu_tu-2010-2016.csv'
        with open(os.path.join(OUT, fname), 'w', newline='',
                  encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['chi_muc', 'loai', '2010', '2014', '2015', '2016'])
            for r in rows_hn[1:]:
                w.writerow(r[:6])
        meta_out.append({
            'file_csv': fname, 'kieu': 'A_cot_nam',
            'tieu_de_vn': 'Tỷ lệ hộ nghèo phân theo địa phương '
                          '(theo chuẩn nghèo của Chính phủ giai đoạn cũ)',
            'don_vi': '%', 'nam_dau': '2010', 'nam_cuoi': '2016',
            'cac_nam_co': '2010 2014 2015 2016', 'so_nien_giam': 1,
            'nguon': '2017:' + os.path.basename(src[0]),
            'ghi_chu': '4 cot dau (khoi "Theo chuan ngheo cua Chinh phu") '
                       'cua bang ho ngheo NGTK2017',
        })

    with open(os.path.join(OUT, '_meta.csv'), 'w', newline='',
              encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(meta_out[0].keys()))
        w.writeheader()
        w.writerows(meta_out)

    n_files = len(glob.glob(os.path.join(OUT, '*.csv'))) - 1  # tru _meta
    print(f'Da noi: {n_a} chuoi kieu A (wide tinh x nam), {n_b} kieu B '
          f'(chong nam), {n_skip} chuoi doi cau truc -> tach rieng tung nam.')
    print(f'Trong do chi tieu chi co 1 nien giam: {len(singles)}')
    print(f'Tong file CSV trong tonghop (khong tinh _meta): {n_files} | '
          f'dong meta: {len(meta_out)}')
    if b_fail:
        print('\nChuoi doi cau truc (da tach tung nam, xu tay neu can noi):')
        for slug, d in b_fail:
            print(f'  {slug[:55]}: {d}')


if __name__ == '__main__':
    main()

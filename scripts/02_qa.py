# -*- coding: utf-8 -*-
"""
Giai doan 2: Kiem dinh (QA) du lieu da trich xuat.

Kiem tra tung file CSV trong data/<nam>/:
  1. Co dong CA NUOC; du 6 vung; so tinh (63 hoac it hon co ly do)
  2. Khong trung lap tinh
  3. Ty le o hop le (so / ky hieu thieu '..' '-')
  4. Tinh cong don: tong 6 vung ~ CA NUOC (chi ap dung chi tieu cong duoc;
     chi tieu ty le/chi so/binh quan se khong khop -> ghi nhan 'khong cong')

Cach dung:  python scripts/02_qa.py 2017
Output:     docs/qa-<nam>.md ; cap nhat trang_thai='da_kiem' trong catalog
            cho cac bang dat.
"""
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vn_names import PROVINCES_63, REGIONS_6

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CATALOG = os.path.join(ROOT, 'docs', 'catalog.csv')
MISSING = {'..', '...', '-', '…', ''}


def is_num(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def qa_file(path):
    rows = list(csv.reader(open(path, encoding='utf-8-sig')))
    header, body = rows[0], rows[1:]
    ncols = len(header) - 2

    tinh = [r for r in body if r[1] == 'TINH']
    vung = [r for r in body if r[1] == 'VUNG']
    ca_nuoc = [r for r in body if r[1] == 'CA_NUOC']
    khac = [r for r in body if r[1] == 'KHAC']

    dups = {r[0] for r in tinh if [x[0] for x in tinh].count(r[0]) > 1}

    # dong nhan bi roi do watermark (chu xoay chen vao ten tinh): nhan KHAC
    # co gia tri so nhung khong khop danh muc hop le va co dang ky tu lon xon
    LEGIT_KHAC = ('nec', 'không', 'khác', 'dầu khí', 'phân theo', 'hạng',
                  'trực thuộc', 'quân', 'viện', 'tòa', 'quốc phòng', 'cấp trung',
                  'người', 'trẻ em', 'nạn nhân', 'north central', 'of which',
                  'trong đó', 'chia ra', 'đông', 'tây', 'nam', 'bắc')
    garbled = []
    for r in khac:
        low = r[0].lower()
        has_val = any(c not in MISSING for c in r[2:])
        if has_val and not any(t in low for t in LEGIT_KHAC):
            garbled.append(r[0])

    cells = [c for r in body for c in r[2:]]
    n_bad = sum(1 for c in cells if c not in MISSING and not is_num(c))
    ty_le_hop_le = 1 - n_bad / len(cells) if cells else 0

    # tinh cong don theo tung cot: tong 6 vung so voi CA NUOC
    additivity = None
    if ca_nuoc and len(vung) == 6:
        diffs = []
        for j in range(2, 2 + ncols):
            try:
                total = float(ca_nuoc[0][j])
                s = sum(float(v[j]) for v in vung if v[j] not in MISSING)
                if total:
                    diffs.append(abs(s - total) / abs(total))
            except (ValueError, IndexError):
                continue
        if diffs:
            additivity = sorted(diffs)[len(diffs) // 2]  # trung vi

    issues = []
    if not ca_nuoc:
        # mot so bang PDF goc khong co dong ca nuoc (ty suat di cu thuan,
        # chi so gia khong gian SCOLI...) -> canh bao neu du lieu tinh van day
        if len(tinh) >= 60:
            issues.append('PDF goc khong co dong CA NUOC (dac thu chi tieu)')
        else:
            issues.append('thieu dong CA NUOC')
    if len(vung) != 6:
        issues.append(f'chi co {len(vung)}/6 vung')
    if dups:
        issues.append(f'tinh trung lap: {", ".join(sorted(dups))}')
    if ty_le_hop_le < 0.98:
        issues.append(f'chi {ty_le_hop_le:.0%} o hop le')
    if garbled:
        issues.append(f'dong bi roi (watermark?), can nhap tay: {garbled}')

    return {
        'n_tinh': len(tinh), 'n_vung': len(vung), 'n_ca_nuoc': len(ca_nuoc),
        'n_khac': len(khac), 'ncols': ncols, 'ty_le_hop_le': ty_le_hop_le,
        'additivity': additivity, 'issues': issues, 'garbled': garbled,
    }


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else '2017'
    data_dir = os.path.join(ROOT, 'data', 'csv', year)
    meta = list(csv.DictReader(open(os.path.join(data_dir, '_meta.csv'),
                                    encoding='utf-8-sig')))

    results = []
    for m in meta:
        if m['trang_thai'] != 'da_trich':
            results.append((m, None))
            continue
        r = qa_file(os.path.join(data_dir, m['file_csv']))
        results.append((m, r))

    n_pass = n_warn = n_fail = 0
    lines_fail, lines_warn, lines_add = [], [], []
    for m, r in results:
        if r is None:
            n_fail += 1
            lines_fail.append(f"- tr.{m['trang_dau']}: {m['tieu_de_vn']} — "
                              f"LOI TRICH XUAT: {m['ghi_chu']}")
            continue
        label = f"tr.{m['trang_dau']}-{m['trang_cuoi']} `{m['file_csv']}`"
        if r['issues'] and ('thieu dong CA NUOC' in r['issues'][0]
                            or any('trung lap' in i or 'hop le' in i
                                   for i in r['issues'])):
            n_fail += 1
            lines_fail.append(f"- {label}: {'; '.join(r['issues'])}")
        elif r['issues'] or r['n_tinh'] < 63:
            n_warn += 1
            why = r['issues'] + ([f"{r['n_tinh']}/63 tinh (PDF goc khong liet ke "
                                  f"du - dac thu chi tieu)"] if r['n_tinh'] < 63 else [])
            lines_warn.append(f"- {label}: {'; '.join(why)}")
        else:
            n_pass += 1
        if r['additivity'] is not None:
            tag = 'KHOP' if r['additivity'] < 0.01 else 'khong cong duoc (ty le/chi so?)'
            lines_add.append((r['additivity'], f"- {label}: lech trung vi "
                             f"{r['additivity']:.2%} -> {tag}"))

    n_khop = sum(1 for d, _ in lines_add if d < 0.01)

    out = os.path.join(ROOT, 'docs', f'qa-{year}.md')
    with open(out, 'w', encoding='utf-8') as f:
        f.write(f"# Bao cao QA du lieu NGTK{year}\n\n")
        f.write(f"Tong so bang: {len(results)} | Dat: {n_pass} | "
                f"Canh bao (hop le nghiep vu): {n_warn} | Loi: {n_fail}\n\n")
        f.write(f"## Kiem tra tinh cong don (tong 6 vung vs CA NUOC)\n\n")
        f.write(f"{n_khop}/{len(lines_add)} bang khop <1% (bang khong khop "
                f"thuong la chi tieu ty le/chi so/binh quan - dung nhu ky vong).\n\n")
        bad_add = [l for d, l in sorted(lines_add, key=lambda x: -x[0]) if d >= 0.01]
        if bad_add:
            f.write("Cac bang KHONG khop (kiem tra thu cong neu chi tieu dang le "
                    "phai cong duoc):\n\n" + '\n'.join(bad_add[:40]) + '\n\n')
        gar = [(m, r) for (m, r) in results if r and r.get('garbled')]
        if gar:
            f.write("## Bang co dong bi roi do watermark — CAN NHAP TAY o GD3\n\n")
            for m, r in gar:
                f.write(f"- tr.{m['trang_dau']} `{m['file_csv']}`: {r['garbled']}\n")
            f.write("\n")
        if lines_fail:
            f.write("## LOI (can xu ly)\n\n" + '\n'.join(lines_fail) + '\n\n')
        if lines_warn:
            f.write("## Canh bao (da xem xet, chap nhan duoc)\n\n"
                    + '\n'.join(lines_warn) + '\n\n')

    # cap nhat catalog: bang dat -> da_kiem
    cat = list(csv.DictReader(open(CATALOG, encoding='utf-8-sig')))
    passed_pages = {m['trang_dau'] for (m, r) in results
                    if r and not any('trung lap' in i or 'hop le' in i
                                     or i == 'thieu dong CA NUOC'
                                     for i in r['issues'])}
    for c in cat:
        if c['nam'] == year and c['dia_phuong'] == '1' \
                and c['trang_dau'] in passed_pages:
            c['trang_thai'] = 'da_kiem'
    with open(CATALOG, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(cat[0].keys()))
        w.writeheader()
        w.writerows(cat)

    print(f"QA {year}: dat={n_pass}, canh bao={n_warn}, loi={n_fail}. "
          f"Bao cao: {out}")


if __name__ == '__main__':
    main()

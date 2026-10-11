# -*- coding: utf-8 -*-
"""10 - Doi chieu cac doan trich trong bang voi file PDF da tai ve.

Doc cot 'Doan noi dung goc can su dung', tach cac manh dat trong ngoac kep,
tim trong PDF (co lop text) -> ghi so trang tim thay.
Ket qua: ket-qua/doi_chieu_pdf.json
"""
import json, os, re, unicodedata
import fitz, openpyxl

V2 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(V2, 'tai-lieu-doi-chieu')
MAN = os.path.join(V2, 'ket-qua', 'manifest_tai_lieu.json')

def chuan(s):
    s = unicodedata.normalize('NFKC', s or '')
    s = s.replace('\u2019', "'").replace('\u2018', "'").replace('\u201c', '"').replace('\u201d', '"')
    s = s.replace('\u2013', '-').replace('\u2014', '-').replace('\u2212', '-').replace('\xa0', ' ')
    return re.sub(r'\s+', ' ', s).strip().lower()

def manh_trich(txt):
    """Lay cac doan trong ngoac kep, dai >= 25 ky tu."""
    t = (txt or '').replace('\u201c', '"').replace('\u201d', '"')
    out = [m.strip() for m in re.findall(r'"([^"]{25,400})"', t)]
    if not out:  # khong co ngoac kep -> lay cau dai nhat
        cau = [c.strip() for c in re.split(r'(?<=[.;])\s+', t) if len(c.strip()) >= 40]
        out = cau[:3]
    return out[:5]

def rut(s):
    """Chi giu chu va so (bo dau cau, xuong dong, gach noi) de so khop ben vung."""
    return re.sub(r'[^a-z0-9]+', '', chuan(s))

def tim(pages_rut, frag):
    """Tra ve (so_trang, ty_le_khop). Thu cac tien to ngan dan."""
    f = rut(frag)
    if len(f) < 20:
        return 0, 0.0
    for cat in (len(f), 160, 110, 80, 55, 40):
        probe = f[:cat]
        if len(probe) < 40:
            break
        for i, t in enumerate(pages_rut):
            if probe in t:
                return i + 1, round(len(probe) / len(f), 2)
    return 0, 0.0

def main():
    wb = openpyxl.load_workbook(os.path.join(V2, 'Bang-doi-chieu-trich-dan.xlsx'))
    ws = wb['Doi_chieu_trich_dan']
    rows = list(ws.iter_rows(min_row=4, values_only=True))
    hdr = rows[0]
    man = {r['stt']: r for r in json.load(open(MAN, encoding='utf-8'))}
    kq = {}
    for r in rows[1:]:
        if not r[0]:
            continue
        stt = int(r[0]); rec = man.get(stt, {})
        p = os.path.join(OUT, rec.get('file', ''))
        item = {'stt': stt, 'file': rec.get('file'), 'co_file': False,
                'so_trang': 0, 'quet': False, 'tim_thay': [], 'khong_thay': []}
        if os.path.exists(p) and os.path.getsize(p) > 20000:
            item['co_file'] = True
            d = fitz.open(p)
            pages = [rut(pg.get_text()) for pg in d]
            item['so_trang'] = d.page_count
            item['quet'] = sum(len(t) for t in pages) < 1500
            if not item['quet']:
                for frag in manh_trich(dict(zip(hdr, r))['Đoạn nội dung gốc cần sử dụng']):
                    tr, ty = tim(pages, frag)
                    (item['tim_thay'] if tr else item['khong_thay']).append(
                        {'doan': frag[:160], 'trang_pdf': tr, 'khop': ty})
        kq[stt] = item
        print(f"{stt:>3} file={'x' if item['co_file'] else '-'} quet={'x' if item['quet'] else '-'} "
              f"thay={len(item['tim_thay'])} khong={len(item['khong_thay'])}")
    json.dump(kq, open(os.path.join(V2, 'ket-qua', 'doi_chieu_pdf.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)

if __name__ == '__main__':
    main()

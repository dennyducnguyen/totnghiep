# -*- coding: utf-8 -*-
"""08 - Thu thap tai lieu goc ve thu muc tai-lieu-doi-chieu/.

Buoc 1: doc Bang-doi-chieu-trich-dan.xlsx -> danh sach 51 tai lieu.
Buoc 2: tra OpenAlex (theo DOI hoac tieu de) de lay ban open access.
Buoc 3: tai PDF ve, ghi manifest JSON.
"""
import json, os, re, sys, time, unicodedata
import openpyxl, requests

V2 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(V2, 'tai-lieu-doi-chieu')
MANIFEST = os.path.join(V2, 'ket-qua', 'manifest_tai_lieu.json')
os.makedirs(OUT, exist_ok=True)

UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                    '(KHTML, like Gecko) Chrome/124.0 Safari/537.36'}

def khong_dau(s):
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s.replace('\u0111', 'd').replace('\u0110', 'D')

def slug(s, n=70):
    s = khong_dau(s)
    s = re.sub(r'[^A-Za-z0-9]+', '-', s).strip('-')
    return s[:n].strip('-')

def doc_bang():
    wb = openpyxl.load_workbook(os.path.join(V2, 'Bang-doi-chieu-trich-dan.xlsx'))
    ws = wb['Doi_chieu_trich_dan']
    rows = list(ws.iter_rows(min_row=4, values_only=True))
    hdr, out = rows[0], []
    for r in rows[1:]:
        if not r[0]:
            continue
        d = dict(zip(hdr, r))
        apa = d['Tài liệu tham khảo APA đầy đủ'] or ''
        m = re.search(r'(10\.\d{4,9}/[^\s,;]+)', apa)
        out.append({
            'stt': int(d['STT']),
            'tac_gia': (d['Tác giả'] or '').strip(),
            'nam': str(d['Năm'] or '').strip(),
            'ten': (d['Tên bài báo / tài liệu'] or '').strip(),
            'apa': apa.strip(),
            'doi': m.group(1).rstrip('.') if m else None,
        })
    return out

def tieu_de_goc(ten):
    """Cat phan tap chi/nha xuat ban de con lai tieu de bai."""
    t = re.split(r'\.\s+[A-ZÀ-Ỹ][^.]*?,?\s*\d*\(?', ten)[0]
    return re.split(r'\s+https?://', t)[0].strip(' .')

def openalex(rec):
    base = 'https://api.openalex.org/works'
    try:
        if rec['doi']:
            r = requests.get(f"{base}/doi:{rec['doi']}", headers=UA, timeout=30)
            if r.status_code == 200:
                return r.json()
        q = tieu_de_goc(rec['ten'])
        r = requests.get(base, params={'filter': f'title.search:{q}', 'per-page': 3},
                         headers=UA, timeout=30)
        if r.status_code == 200:
            res = r.json().get('results') or []
            for w in res:
                if str(w.get('publication_year') or '') == rec['nam'] or len(res) == 1:
                    return w
            return res[0] if res else None
    except Exception as e:
        print('  ! openalex loi:', e)
    return None

def tai(url, dest, referer=None):
    h = dict(UA)
    if referer:
        h['Referer'] = referer
    try:
        r = requests.get(url, headers=h, timeout=90, allow_redirects=True)
    except Exception as e:
        return False, str(e)
    if r.status_code != 200:
        return False, f'HTTP {r.status_code}'
    ct = (r.headers.get('Content-Type') or '').lower()
    body = r.content
    if not (body[:5] == b'%PDF-' or 'pdf' in ct):
        return False, f'khong phai PDF ({ct[:40]}, {len(body)} bytes)'
    if len(body) < 20000:
        return False, f'PDF qua nho ({len(body)} bytes)'
    with open(dest, 'wb') as f:
        f.write(body)
    return True, f'{len(body)//1024} KB'

def main():
    recs = doc_bang()
    man = []
    for rec in recs:
        ho = rec['tac_gia'].split(',')[0].split('&')[0].strip().rstrip('.')
        ten_file = f"{rec['stt']:02d}_{slug(ho, 28)}_{rec['nam']}_{slug(tieu_de_goc(rec['ten']), 60)}.pdf"
        dest = os.path.join(OUT, ten_file)
        rec['file'] = ten_file
        if os.path.exists(dest) and os.path.getsize(dest) > 20000:
            rec['trang_thai'] = 'da co'
            man.append(rec); print(f"{rec['stt']:02d} da co"); continue
        w = openalex(rec)
        cand = []
        if w:
            rec['oa_id'] = w.get('id'); rec['oa_title'] = w.get('title')
            rec['doi'] = rec['doi'] or (w.get('doi') or '').replace('https://doi.org/', '') or None
            for loc in ([w.get('best_oa_location')] + (w.get('locations') or [])):
                if not loc: continue
                for k in ('pdf_url', 'landing_page_url'):
                    u = loc.get(k)
                    if u and u not in cand and (k == 'pdf_url' or u.endswith('.pdf')):
                        cand.append(u)
        rec['ung_vien'] = cand
        ok, note = False, 'khong tim thay ban open access'
        for u in cand[:6]:
            ok, note = tai(u, dest)
            print(f"  {'OK ' if ok else '-- '}{u[:95]} :: {note}")
            if ok:
                rec['nguon_tai'] = u
                break
        rec['trang_thai'] = 'da tai' if ok else 'chua tai'
        rec['ghi_chu_tai'] = note
        man.append(rec)
        print(f"{rec['stt']:02d} {rec['trang_thai']}: {ten_file[:70]}")
        time.sleep(0.4)
    os.makedirs(os.path.dirname(MANIFEST), exist_ok=True)
    json.dump(man, open(MANIFEST, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    n = sum(1 for r in man if r['trang_thai'] in ('da tai', 'da co'))
    print(f'\n== {n}/{len(man)} tai lieu co file ==')

if __name__ == '__main__':
    main()

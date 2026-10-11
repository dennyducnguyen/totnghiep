# -*- coding: utf-8 -*-
"""09 - Tach trang PDF quet thanh anh PNG de doc bang vision.

Dung: python scripts/09_pdf_sang_anh.py <ten-file-hoac-STT> <trang-bat-dau> <trang-ket-thuc> [dpi]
Trang tinh tu 1. Anh ghi vao <scratch>/img/<STT>_p<NN>.png
"""
import glob, json, os, sys
import fitz

V2 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.environ.get('IMG_DIR', os.path.join(V2, 'ket-qua', 'anh-quet'))

def tim(arg):
    if arg.isdigit():
        g = glob.glob(os.path.join(V2, 'tai-lieu-doi-chieu', f'{int(arg):02d}_*.pdf'))
        return g[0] if g else None
    return arg

def main():
    src = tim(sys.argv[1]); a = int(sys.argv[2]); b = int(sys.argv[3])
    dpi = int(sys.argv[4]) if len(sys.argv) > 4 else 150
    os.makedirs(OUT, exist_ok=True)
    d = fitz.open(src)
    stem = os.path.basename(src)[:2]
    for i in range(a - 1, min(b, d.page_count)):
        pix = d[i].get_pixmap(dpi=dpi)
        p = os.path.join(OUT, f'{stem}_p{i+1:03d}.png')
        pix.save(p)
        print(p, f'{os.path.getsize(p)//1024} KB')

if __name__ == '__main__':
    main()

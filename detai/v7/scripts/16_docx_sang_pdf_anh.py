# -*- coding: utf-8 -*-
"""Chuyển một file .docx sang PDF bằng MS Word (COM) rồi tách từng trang thành ảnh PNG để kiểm tra bố cục.
Chạy: python scripts/16_docx_sang_pdf_anh.py <file.docx> [dpi]
Đầu ra: ket-qua-v7/kiem-tra/<tên>.pdf và <tên>_pNNN.png. Máy không có LibreOffice/pandoc nên dùng Word COM."""
import os, sys, time
import fitz
import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src = os.path.abspath(sys.argv[1]); dpi = int(sys.argv[2]) if len(sys.argv) > 2 else 80
out_dir = os.path.join(ROOT, "ket-qua-v7", "kiem-tra"); os.makedirs(out_dir, exist_ok=True)
base = os.path.splitext(os.path.basename(src))[0]
pdf = os.path.join(out_dir, base + ".pdf")

word = win32com.client.DispatchEx("Word.Application"); word.Visible = False
try:
    doc = word.Documents.Open(src, ReadOnly=True)
    doc.ExportAsFixedFormat(pdf, 17)   # 17 = wdExportFormatPDF
    doc.Close(False)
finally:
    word.Quit()
time.sleep(0.5)
d = fitz.open(pdf); print("PDF:", pdf, "| số trang:", len(d))
for i, p in enumerate(d):
    pix = p.get_pixmap(dpi=dpi); pix.save(os.path.join(out_dir, f"{base}_p{i + 1:03d}.png"))
print("Đã tách", len(d), "trang ảnh vào", out_dir)

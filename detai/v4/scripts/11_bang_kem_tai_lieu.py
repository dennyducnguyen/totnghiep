# -*- coding: utf-8 -*-
"""11 - Dung Bang-doi-chieu-trich-dan-kem-tai-lieu.xlsx: bang doi chieu + tham chieu file goc.

Dau vao:
  Bang-doi-chieu-trich-dan.xlsx        (do script 06 tao ra, 51 tai lieu)
  ket-qua/manifest_tai_lieu.json       (ket qua tai ve - script 08/round2)
  ket-qua/doi_chieu_pdf.json           (doi chieu doan trich voi PDF - script 10)
  ket-qua/vision_pdf_quet.json         (noi dung doc bang vision cho PDF quet)
Dau ra:
  Bang-doi-chieu-trich-dan-kem-tai-lieu.xlsx  (3 sheet: doi chieu / kho tai lieu / chu thich)

Moi dong tro toi file tuong ung trong tai-lieu-doi-chieu/ bang hyperlink tuong doi.
"""
import json, os
import fitz, openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

V2 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
THUMUC = 'tai-lieu-doi-chieu'
NGUON = os.path.join(V2, 'Bang-doi-chieu-trich-dan.xlsx')
DICH = os.path.join(V2, 'Bang-doi-chieu-trich-dan-kem-tai-lieu.xlsx')

# Ly do khong tai duoc + lien ket tra cuu chinh thuc (16 tai lieu)
KHONG_TAI = {
    5:  ('Bài tạp chí có tường phí (Taylor & Francis), không có bản open access',
         'https://doi.org/10.1080/09654310701748009'),
    9:  ('Bài open access nhưng cả ScienceDirect lẫn kho Politecnico di Milano đều chặn tải tự động '
         '(Cloudflare) – mở link bằng trình duyệt để tải thủ công',
         'https://doi.org/10.1016/j.pirs.2024.100118'),
    11: ('Sách (Oxford University Press) – không có bản PDF hợp pháp để tải',
         'https://global.oup.com/academic/product/measuring-inequality-9780199594047'),
    12: ('Bài tạp chí có tường phí; bản AISeL yêu cầu phiên trình duyệt nên không tải tự động được',
         'https://aisel.aisnet.org/misq/vol13/iss3/6/'),
    21: ('Bài tạp chí trên JSTOR, có tường phí',
         'https://doi.org/10.2307/1928447'),
    24: ('Sách (Cambridge University Press) – không có bản PDF hợp pháp để tải',
         'https://doi.org/10.1017/CBO9781139164887'),
    30: ('Sách (Free Press, tái bản lần 5) – không có bản PDF hợp pháp để tải',
         'https://archive.org/details/diffusionofinnov0000roge'),
    33: ('Sách (North-Holland, 1967) – không có bản PDF hợp pháp để tải',
         'https://archive.org/details/economicsinforma0000thei'),
    34: ('', ''),  # da tai duoc, giu cho de tranh KeyError
    35: ('Sách (Lexington Books, 1990) – không có bản PDF hợp pháp để tải',
         'https://archive.org/details/processesoftechn0000torn'),
    40: ('Sách (Polity Press, 2020) – không có bản PDF hợp pháp để tải',
         'https://www.wiley.com/en-us/The+Digital+Divide-p-9781509534456'),
    41: ('Bài tạp chí có tường phí (Elsevier), không có bản lưu trữ mở',
         'https://doi.org/10.1016/j.telpol.2010.12.013'),
    42: ('Không tìm thấy bản công bố trực tuyến của báo cáo này trên ciem.org.vn hay nguồn khác – '
         'cần xin trực tiếp từ CIEM hoặc kiểm tra lại trích dẫn',
         'http://ciem.org.vn/'),
    43: ('Bài tạp chí có tường phí (Emerald), không có bản open access',
         'https://doi.org/10.1108/CR-03-2025-0105'),
    44: ('Bài tạp chí có tường phí (Taylor & Francis), OpenAlex xác nhận closed access',
         'https://doi.org/10.1080/14631377.2025.2511511'),
    47: ('Sách (MIT Press, 2003) – không có bản PDF hợp pháp để tải',
         'https://mitpress.mit.edu/9780262731737/technology-and-social-inclusion/'),
    48: ('Bài tạp chí trên University of Chicago Press / JSTOR, có tường phí',
         'https://doi.org/10.1086/450136'),
}

# Ghi chu rieng cho vai file tai duoc nhung khac ban in
GHI_CHU_FILE = {
    16: 'Bản tải về là Báo cáo Chỉ số TMĐT Việt Nam (EBI) 2023 – đại diện cho chuỗi báo cáo 2017–2025 '
        'được trích dẫn. Các năm khác VECOM không cho tải trực tiếp (trang dùng trình xem nội bộ).',
    32: 'Bản tải về là working paper (Economics Working Paper 117, UPF & Yale, 6/1995) – cùng nội dung '
        'với bài The Economic Journal 106(437), 1019–1036 (1996) nhưng đánh số trang riêng.',
    37: 'Bản tải về là bản preprint của chương sách (nội dung như bản Emerald 2015).',
    50: 'Bản tải về lấy từ Europe PMC (PMC9758626), có trang bìa COVID-19 resource centre của Elsevier.',
}

VIEN = Border(*[Side(style='thin', color='D0D0D0')] * 4)
DAM = Font(bold=True)
NEN_TIEU_DE = PatternFill('solid', fgColor='1F4E79')
CHU_TIEU_DE = Font(bold=True, color='FFFFFF', size=10)
XANH = PatternFill('solid', fgColor='D9EAD3')   # co file, khop nguyen van
VANG = PatternFill('solid', fgColor='FFF2CC')   # co file, doan trong bang la tom tat
CAM = PatternFill('solid', fgColor='FCE5CD')    # PDF quet - doc bang vision
XAM = PatternFill('solid', fgColor='EFEFEF')    # chua tai duoc


def doc_json(ten):
    return json.load(open(os.path.join(V2, 'ket-qua', ten), encoding='utf-8'))


def main():
    wb0 = openpyxl.load_workbook(NGUON)
    ws0 = wb0['Doi_chieu_trich_dan']
    tieu_de_bang = ws0['A1'].value
    mo_ta = ws0['A2'].value
    rows = list(ws0.iter_rows(min_row=4, values_only=True))
    hdr, du_lieu = list(rows[0]), [r for r in rows[1:] if r[0]]

    man = {r['stt']: r for r in doc_json('manifest_tai_lieu.json')}
    dc = {int(k): v for k, v in doc_json('doi_chieu_pdf.json').items()}
    vis = {int(k): v for k, v in doc_json('vision_pdf_quet.json').items()}

    cot_moi = ['File tài liệu gốc (thư mục tai-lieu-doi-chieu)', 'Tình trạng tài liệu gốc',
               'Số trang', 'Trang trong file gốc chứa đoạn trích',
               'Kết quả đối chiếu với file gốc',
               'Nội dung gốc đọc bằng vision (PDF quét, không có lớp text)',
               'Nguồn tải / liên kết tra cứu']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Doi_chieu_trich_dan'
    ws['A1'] = tieu_de_bang
    ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = (mo_ta or '') + ('\n[v2 – 07/09/2026] Bổ sung 7 cột cuối: mỗi dòng trỏ tới file tài liệu gốc '
                                'đã tải về trong thư mục tai-lieu-doi-chieu/ (bấm vào tên file để mở). '
                                'Với PDF quét (không trích được text), nội dung gốc được đọc bằng vision '
                                'và chép vào cột “Nội dung gốc đọc bằng vision”.')
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[2].height = 58

    tat_ca = hdr + cot_moi
    for j, h in enumerate(tat_ca, 1):
        c = ws.cell(row=4, column=j, value=h)
        c.fill, c.font = NEN_TIEU_DE, CHU_TIEU_DE
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        c.border = VIEN
    ws.row_dimensions[4].height = 46

    n_file = n_quet = n_thieu = 0
    for i, r in enumerate(du_lieu):
        dong = 5 + i
        stt = int(r[0])
        rec, d = man.get(stt, {}), dc.get(stt, {})
        ten_file = rec.get('file', '')
        co_file = d.get('co_file', False)
        quet = d.get('quet', False)
        v = vis.get(stt)

        for j, val in enumerate(r, 1):
            c = ws.cell(row=dong, column=j, value=val)
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.border = VIEN

        if co_file:
            n_file += 1
            tinh_trang = 'PDF quét – đã đọc bằng vision' if quet else 'PDF có lớp text'
            so_trang = d.get('so_trang', '')
            thay = d.get('tim_thay', [])
            khong = d.get('khong_thay', [])
            trang = ', '.join(str(t['trang_pdf']) for t in thay) or ''
            if quet:
                n_quet += 1
                trang = f"tr. PDF {v['trang_pdf']} (tr. in {v['trang_in']})" if v else ''
                ket_qua = v['ket_luan'] if v else 'PDF quét – cần đọc bằng vision'
                nen = CAM
            elif thay and not khong:
                ket_qua = f'Đã tìm thấy nguyên văn toàn bộ {len(thay)} đoạn trích trong file gốc.'
                nen = XANH
            elif thay:
                ket_qua = (f'Tìm thấy nguyên văn {len(thay)}/{len(thay) + len(khong)} đoạn. '
                           f'Các đoạn còn lại là tóm tắt/diễn giải của tác giả, không có trong file gốc '
                           f'dưới dạng nguyên văn: ' + ' | '.join(k['doan'][:80] for k in khong))
                nen = VANG
            else:
                ket_qua = ('Không tìm thấy nguyên văn trong file gốc – đoạn trong cột "Đoạn nội dung gốc" '
                           'là tóm tắt/diễn giải của tác giả từ phiếu đọc, không phải trích nguyên văn. '
                           'Đã có file gốc để đối chiếu thủ công.')
                nen = VANG
            nguon = rec.get('nguon_tai', '') or 'có sẵn trong bai-bao-mau/'
            noi_dung_vision = v['noi_dung'] if v else ''
        else:
            n_thieu += 1
            ly_do, link = KHONG_TAI.get(stt, ('Chưa tải được', ''))
            ten_file = ''
            tinh_trang = 'CHƯA TẢI ĐƯỢC'
            so_trang = ''
            trang = ''
            ket_qua = ly_do
            noi_dung_vision = ''
            nguon = link
            nen = XAM

        if stt in GHI_CHU_FILE:
            ket_qua = (ket_qua + '\n[Ghi chú file] ' + GHI_CHU_FILE[stt]).strip()

        gt = [ten_file or '(chưa có)', tinh_trang, so_trang, trang, ket_qua, noi_dung_vision, nguon]
        for k, val in enumerate(gt):
            c = ws.cell(row=dong, column=len(hdr) + 1 + k, value=val)
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.border, c.fill = VIEN, nen
        if co_file:
            c = ws.cell(row=dong, column=len(hdr) + 1)
            c.hyperlink = f'{THUMUC}/{ten_file}'
            c.font = Font(color='0563C1', underline='single')

    rong = {1: 6, 2: 42, 3: 26, 4: 7, 5: 34, 6: 40, 7: 46, 8: 46, 9: 26, 10: 24, 11: 44, 12: 30,
            13: 40, 14: 20, 15: 8, 16: 18, 17: 52, 18: 60, 19: 34}
    for j in range(1, len(tat_ca) + 1):
        ws.column_dimensions[get_column_letter(j)].width = rong.get(j, 24)
    ws.freeze_panes = 'C5'
    ws.auto_filter.ref = f'A4:{get_column_letter(len(tat_ca))}{4 + len(du_lieu)}'

    # --- Sheet 2: kho tai lieu ---
    w2 = wb.create_sheet('Kho_tai_lieu')
    h2 = ['STT', 'Tác giả', 'Năm', 'Tên file trong tai-lieu-doi-chieu', 'Dung lượng (KB)',
          'Số trang', 'Có lớp text?', 'URL đã tải về']
    w2['A1'] = ('KHO TÀI LIỆU GỐC – thư mục tai-lieu-doi-chieu/ '
                '(bấm vào tên file để mở; đường dẫn tương đối so với thư mục v2)')
    w2['A1'].font = Font(bold=True, size=12)
    for j, h in enumerate(h2, 1):
        c = w2.cell(row=3, column=j, value=h)
        c.fill, c.font = NEN_TIEU_DE, CHU_TIEU_DE
        c.alignment = Alignment(wrap_text=True, horizontal='center')
    dong = 4
    for stt in sorted(man):
        rec, d = man[stt], dc.get(stt, {})
        p = os.path.join(V2, THUMUC, rec['file'])
        if not os.path.exists(p):
            continue
        kb = os.path.getsize(p) // 1024
        w2.cell(row=dong, column=1, value=stt)
        w2.cell(row=dong, column=2, value=rec['tac_gia'])
        w2.cell(row=dong, column=3, value=rec['nam'])
        c = w2.cell(row=dong, column=4, value=rec['file'])
        c.hyperlink = f"{THUMUC}/{rec['file']}"
        c.font = Font(color='0563C1', underline='single')
        w2.cell(row=dong, column=5, value=kb)
        w2.cell(row=dong, column=6, value=d.get('so_trang', ''))
        w2.cell(row=dong, column=7, value='Không (PDF quét)' if d.get('quet') else 'Có')
        w2.cell(row=dong, column=8, value=rec.get('nguon_tai', 'có sẵn trong bai-bao-mau/'))
        dong += 1
    for j, w in enumerate([6, 30, 7, 72, 14, 9, 16, 70], 1):
        w2.column_dimensions[get_column_letter(j)].width = w
    w2.freeze_panes = 'A4'

    # --- Sheet 3: chu thich ---
    w3 = wb.create_sheet('Chu_thich')
    ghi = [
        'BẢNG ĐỐI CHIẾU TRÍCH DẪN – PHIÊN BẢN v4 (07/09/2026)',
        '',
        f'Tổng số tài liệu trong danh mục: {len(du_lieu)}',
        f'Đã tải được file gốc: {n_file}  (trong đó {n_quet} file là PDF quét, không có lớp text)',
        f'Chưa tải được: {n_thieu}  (sách có bản quyền, bài có tường phí, hoặc trang nguồn chặn tải tự động)',
        '',
        'Ý nghĩa màu ở nhóm 7 cột cuối:',
        '  Xanh lá  : đã có file gốc và tìm thấy NGUYÊN VĂN toàn bộ đoạn trích trong file.',
        '  Vàng     : đã có file gốc nhưng đoạn ở cột "Đoạn nội dung gốc cần sử dụng" là tóm tắt/diễn giải '
        'của tác giả (từ phiếu đọc), không khớp nguyên văn – cần mở file để lấy câu gốc và số trang.',
        '  Cam      : file gốc là PDF quét (không trích được text). Nội dung đã được tách trang thành ảnh '
        '(lưu ở tai-lieu-doi-chieu/anh-quet/) và đọc bằng vision; kết quả nằm ở cột '
        '"Nội dung gốc đọc bằng vision".',
        '  Xám      : chưa tải được file gốc; cột "Kết quả đối chiếu" ghi lý do, cột "Nguồn tải" ghi liên kết tra cứu.',
        '',
        'Cách đối chiếu tự động (scripts/10_doi_chieu_pdf.py): tách các đoạn trong ngoặc kép ở cột '
        '"Đoạn nội dung gốc cần sử dụng", chuẩn hoá (bỏ dấu câu, xuống dòng, gạch nối, phân biệt hoa thường) '
        'rồi tìm trong từng trang PDF. Nếu khớp, ghi số trang PDF (không phải số trang in của tạp chí).',
        '',
        'Bốn phát hiện cần xử lý trong bài viết:',
        '  1. STT 32 (Sala-i-Martin): số trang 1326–1329 ghi trong bảng cũ không thuộc bài Economic Journal '
        '106(437), 1019–1036. Nội dung là đúng (đã xác nhận nguyên văn trong bản working paper).',
        '  2. STT 7 (QĐ 1568/QĐ-BCT): câu trích trong bài là diễn giải, không nguyên văn – xem cột '
        '"Kết quả đối chiếu" để lấy đúng câu.',
        '  3. STT 34 (QĐ 645/QĐ-TTg): mục tiêu tổng quát không có cụm "kênh phân phối quan trọng" '
        'và "liên kết vùng" như bài đang trích.',
        '  4. STT 2 (Anselin 1995 – LISA): nội dung trong phiếu đọc lẫn phần "convergence clubs" '
        '(vốn thuộc Phillips & Sul 2007). Cần tách lại nội dung của hai tài liệu này.',
        '',
        'Quy trình tái tạo:',
        '  python scripts/08_thu_thap_tai_lieu.py     # tải tài liệu về tai-lieu-doi-chieu/',
        '  python scripts/09_pdf_sang_anh.py <STT> <trang đầu> <trang cuối> [dpi]   # tách PDF quét thành ảnh',
        '  python scripts/10_doi_chieu_pdf.py         # đối chiếu đoạn trích với PDF',
        '  python scripts/11_bang_doi_chieu_v2.py     # dựng file Bang-doi-chieu-trich-dan-kem-tai-lieu.xlsx này',
    ]
    for i, t in enumerate(ghi, 1):
        c = w3.cell(row=i, column=1, value=t)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        if i == 1:
            c.font = Font(bold=True, size=12)
    w3.column_dimensions['A'].width = 135

    wb.save(DICH)
    print(f'Da ghi {DICH}')
    print(f'  {n_file} tai lieu co file goc ({n_quet} PDF quet), {n_thieu} chua tai duoc')


if __name__ == '__main__':
    main()

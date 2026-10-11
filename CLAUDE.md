# CLAUDE.md

> File hướng dẫn DUY NHẤT của dự án. `AGENTS.md` chỉ trỏ về đây — mọi cập nhật ghi vào file này, không ghi vào AGENTS.md.

## Trọng tâm: đề tài tốt nghiệp (`detai/`)

Mục tiêu cuối cùng của repo là **đề tài/bài báo tốt nghiệp** về **khoảng cách số / thương mại điện tử giữa 63 tỉnh thành Việt Nam, 2017–2024**, đo bằng chỉ số EBI (VECOM) kết hợp biến từ Niên giám Thống kê. Trên máy người dùng thư mục này nằm tại `D:\ai\thongke\detai` (trong repo: `detai/`). Phần trích xuất Niên giám (`niengiam/`, `data/`, `scripts/`) là **hạ tầng dữ liệu phục vụ đề tài**, đã xong.

- **Tác giả:** Đào Thị Hồng Vân. Repository: `https://github.com/dennyducnguyen/totnghiep` (nhánh chính `main`).
- **Bản hiện hành: `detai/v7/`** (14/09/2026, chỉnh sửa sau phản biện "tinh gọn phương pháp"). Bài chính: `v7/Bai-bao-EBI-63-tinh-2017-2024-v7.docx`; báo cáo chẩn đoán: `v7/BAO-CAO-CHINH-SUA-V7.docx`. **Đọc `detai/v7/GHI-CHU-V7.txt` trước khi làm gì với đề tài.**
- **Quy ước phiên bản:** mỗi vòng sửa lớn = thư mục `detai/vN/` là bản sao đầy đủ, tự chứa của bản trước + file mới có hậu tố `-vN`/`_vN`; **không ghi đè file của bản trước** (bộ dựng docx v7 từ chối ghi đè). Có ghi chú `GHI-CHU-VN.txt` riêng. Các thư mục `v2`–`v6` và file ở gốc `detai/` là lịch sử — chỉ tra cứu, không sửa.
- **Đầu vào của đề tài:** `v7/du-lieu-dau-vao/EBI_NGTK_Panel_2017_2024.xlsx` (panel EBI + biến NGTK), `doi_chieu_citations.xlsx`, `hdnc18062026.docx` (hướng dẫn nghiên cứu ban đầu); EBI gốc của VECOM: `Dulieu_EBI_2017_2025.xlsx` (gốc repo). Biến bổ sung lấy từ `data/csv/tonghop/`.
- **Thiết kế v7:** RQ1 mức độ (Gini chính, Theil-T, SD ln EBI; Gini từng trụ cột); RQ2 phân rã Theil-T 6 vùng; RQ3 FE tỉnh + năm, SE cụm tỉnh (Hausman χ²(8) = 72,4). Đã LOẠI: beta có điều kiện, Oaxaca–Blinder, biến giả vùng + Wald, VIF, hệ số chuẩn hóa, H1–H5 cũ.
- **Các điểm phương pháp bắt buộc tôn trọng:**
  - **Điểm gãy EBI 2021:** kỳ 2021 là kỳ đầu tiên chỉ còn 3 trụ cột + thang mới → không so sánh mức tuyệt đối qua 2020/2021; phân tích tách giai đoạn. Kỳ 2025 đổi trọng số 0,1/0,6/0,3 → không nối vào chuỗi.
  - **(S1)** từ 2021 EBI chứa chỉ tiêu DN/dân số, thu nhập, tên miền/dân số (VECOM 2023 tr. 49–53) → "mật độ doanh nghiệp" trùng một phần định nghĩa với biến phụ thuộc; không viết nó là "yếu tố quyết định".
  - **(S2)** hội tụ beta mất ý nghĩa khi loại Hà Nội & TP.HCM → không kết luận "tỉnh thấp đuổi kịp".
  - Dữ liệu thiếu: 445 quan sát, 54–58 tỉnh/năm, 53 tỉnh đủ 8 năm (thiếu không ngẫu nhiên). Ngôn ngữ **liên hệ thống kê**, không nhân quả.
- **Trích dẫn:** chỉ dùng tài liệu đã kiểm chứng có file gốc trong `v7/tai-lieu-doi-chieu/` (hoặc kinh điển đã duyệt: Rogers 2003, Theil 1967). Không bịa số liệu, không thêm tài liệu chưa xác minh DOI/bản gốc. Mỗi lần sửa trích dẫn phải cập nhật bảng đối chiếu (`06/10/11_*_v7.py`).
- **Chạy lại v7** (trong `detai/v7/`): `python scripts/01_phan_tich_v7.py` → `python hinh-bai-bao/tao_hinh_v7.py` → `python scripts/02_bang_docx_v7.py` → `node scripts/03_build_docx_v7.js` (thêm `bao_cao_chinh_sua_v7.js` để dựng báo cáo) → `python scripts/16_docx_sang_pdf_anh.py <docx>` (cần MS Word) → `06_bang_trich_dan_v7.py` → `10_doi_chieu_pdf_v7.py` → `11_bang_kem_tai_lieu_v7.py`. Cần Python (pandas, statsmodels, linearmodels, scipy, openpyxl, pymupdf, matplotlib, pywin32) và Node.js (gói `docx`).
- **Kỹ thuật docx:** công thức là OMML tự sinh, nhúng bằng `ImportedXmlComponent.fromXmlString(xml).root[0]`; phương trình hiển thị đặt trong bảng 3 cột không viền. Số trong bảng/bài theo kiểu Việt (dấu phẩy thập phân).
- **Việc còn treo của đề tài:** các file phụ trợ v6 (`MO-TA-CONG-THUC-TINH.docx`, `HUONG-DAN-TINH-TUNG-BUOC.*`, `SO-TINH-TOAN-DAY-DU-BAI-BAO.xlsx`) chưa viết lại theo v7. Ở gốc repo có `EBI_NGTK_Panel_2017_2024{,_edited,_rev1,_rev2}.xlsx` và `Bao_cao_doi_chieu_EBI_NGTK.docx` chưa được ghi chú nguồn gốc — hỏi người dùng trước khi dùng thay cho bản trong `v7/du-lieu-dau-vao/`.

## Vai trò

Trợ lý nghiên cứu cho đề tài trên: (1) phân tích, viết, chỉnh sửa bài báo và kiểm chứng trích dẫn trong `detai/`; (2) khi cần biến mới, trích/làm sạch số liệu thống kê theo tỉnh (chủ yếu từ PDF Niên giám Thống kê) thành CSV/Excel có cấu trúc.

## Hạ tầng dữ liệu Niên giám (đã xong)

- **Phạm vi thu thập:** không chỉ chỉ tiêu CNTT — mọi dữ liệu theo tỉnh (dân số, lao động, GRDP, thu nhập, giáo dục, y tế, hạ tầng, doanh nghiệp, đầu tư, thương mại, viễn thông…) làm biến giải thích/kiểm soát.
- **Đơn vị hành chính:** dùng **63 tỉnh thành TRƯỚC sáp nhập** (trước sắp xếp 2025). Nguồn mới dùng đơn vị sau sáp nhập thì quy đổi/ghi chú rõ, không trộn hai hệ thống.
- **Nguồn:** `niengiam/NGTK2017.pdf` → `NGTK2024.pdf` (Niên giám năm X chứa số đến năm X ước tính và các năm trước).

### Cấu trúc thư mục

- `niengiam/` — PDF nguồn (không sửa, không xóa).
- `data/csv/<năm>/` — dữ liệu trích theo từng niên giám; `data/csv/tonghop/` — chuỗi đã nối (MỘT CỬA); `data/excel/` — bản .xlsx gương cấu trúc csv.
- `scripts/` — script trích xuất/xử lý Niên giám. `docs/` — catalog, tiến độ, QA, ánh xạ đổi tên.
- `detai/` — đề tài tốt nghiệp (xem trên).

### Quy ước dữ liệu

- **CSV** UTF-8 (BOM nếu mở bằng Excel), long/tidy khi có thể: `tinh, ma_tinh, nam, chi_tieu, gia_tri, don_vi, nguon`.
- **Tên tỉnh** chuẩn theo danh mục 63 tỉnh (kèm mã TCTK), một cách viết duy nhất ("TP. Hồ Chí Minh", "Bà Rịa - Vũng Tàu").
- **Số:** dấu chấm thập phân trong CSV; ghi rõ đơn vị. PDF dùng dấu phẩy thập phân + dấu chấm hàng nghìn — chuyển đổi cẩn thận.
- **Nguồn gốc:** mỗi bảng ghi file PDF, số trang, tên bảng. **Thiếu:** để trống/`NA`, không điền 0. Số "sơ bộ"/"ước tính" có cột ghi chú nếu phân biệt được.
- CSV chứa dữ liệu; Markdown chỉ chứa ghi chú (catalog, tiến độ, QA), không chứa bảng dữ liệu.

### Pipeline (PDF không đi qua context — script đọc PDF ghi thẳng ra CSV)

- **GĐ0 Catalog:** `scripts/00_catalog.py` → `docs/catalog.csv` (3.550 bảng, 1.206 bảng "phân theo địa phương"), nhận dạng bằng chữ ký font (tiêu đề VN Bold 12, EN Italic 12, thân ~8.3).
- **GĐ1 Trích xuất:** `scripts/01_extract.py <năm>` → `data/csv/<năm>/` + `_meta.csv`: 1.207 CSV, 380.480 ô. **Đọc mục "Bài học" (#1–18) trong `docs/tien-do.md` trước khi sửa parser hoặc thêm năm.**
- **GĐ2 Kiểm định:** `scripts/02_qa.py <năm>` → `docs/qa-<năm>.md`; 0 lỗi mọi năm, đã đối chiếu vision chọn mẫu.
- **GĐ3 Nối chuỗi + vá:** `scripts/04_patch.py` (vá 5 dòng watermark/text ẩn), `scripts/03_tonghop.py` → `data/csv/tonghop/` (228 file + `_meta`, phủ 1.207/1.207). Trùng năm lấy niên giám mới nhất; đổi tên ghép theo `RENAME_MAP` (`docs/anh-xa-doi-ten.md`).
- **Thứ tự khi trích lại từ đầu:** `01_extract.py <năm>` → `04_patch.py` → `02_qa.py <năm>` → `03_tonghop.py` → `06_to_excel.py`.
- Tùy chọn còn lại: đặt tên cột chuỗi kiểu B, trích bảng cấp quốc gia (dia_phuong=0).

### Lưu ý kỹ thuật

- Cả 8 PDF có text layer đầy đủ → trích trực tiếp, không cần OCR. Font TCVN3 lỗi chỉ ở bìa/tiêu đề chạy.
- `extract_tables()` của pdfplumber thất bại (bảng không kẻ đủ). Cách đúng: `page.extract_words()` → gom dòng theo `top`, tách cột theo cụm `x0`.
- Từ NGTK2022 bố cục đổi → test riêng nhóm 2022–2024. NGTK2018 lỗi "ƣ" thay "ư" trong tiêu đề.
- Số cùng một năm có thể khác giữa các kỳ niên giám (sơ bộ → chính thức → điều chỉnh) → luôn lấy niên giám mới nhất.
- NGTK KHÔNG có PCI, logistics index, tỷ lệ Internet theo tỉnh–năm (chỉ 2024) → cần nguồn khác (niên giám tỉnh, Sách trắng CNTT-TT, DTI, PCI, PAPI…) hoặc ghi là hạn chế.

## Quy trình nối phiên (bắt buộc)

- Đầu phiên: đọc `docs/tien-do.md`; nếu làm đề tài, đọc thêm `GHI-CHU` của bản mới nhất trong `detai/` (hiện là `detai/v7/GHI-CHU-V7.txt`); nếu làm dữ liệu, đọc `docs/catalog.csv`.
- Cuối phiên: cập nhật `docs/tien-do.md` (và `GHI-CHU-VN.txt` nếu tạo bản đề tài mới). Phát hiện ảnh hưởng giải pháp → cập nhật ngay vào CLAUDE.md (KHÔNG vào AGENTS.md).
- Máy người dùng: Windows 11, PowerShell; gốc dự án `D:\ai\thongke\`.

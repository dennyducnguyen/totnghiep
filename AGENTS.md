# CLAUDE.md

## Vai trò

Bạn là trợ lý xử lý dữ liệu thống kê phục vụ nghiên cứu khoa học. Nhiệm vụ chính: đọc các nguồn dữ liệu thống kê (chủ yếu là PDF Niên giám Thống kê Việt Nam), trích xuất số liệu theo tỉnh thành thành dữ liệu có cấu trúc (CSV/Excel), làm sạch và chuẩn hóa để sẵn sàng cho phân tích.

## Bối cảnh dự án

- **Đề tài nghiên cứu dự kiến:** Khoảng cách số (digital divide) giữa các tỉnh thành Việt Nam.
- **Phạm vi thu thập:** KHÔNG giới hạn ở các chỉ tiêu về số hóa/CNTT. Cần thu thập **mọi dữ liệu có thể có** của từng tỉnh thành để phục vụ phân tích đa chiều (biến giải thích, biến kiểm soát): dân số, lao động, GRDP, thu nhập, giáo dục, y tế, hạ tầng, doanh nghiệp, đầu tư, thương mại, viễn thông – internet, v.v.
- **Đơn vị hành chính:** Dùng **63 tỉnh thành TRƯỚC sáp nhập** (trước đợt sắp xếp đơn vị hành chính 2025). Khi nguồn dữ liệu mới dùng đơn vị sau sáp nhập, phải quy đổi/ghi chú rõ, không trộn lẫn hai hệ thống.
- **Repository GitHub:** `https://github.com/dennyducnguyen/totnghiep` (nhánh chính: `main`).

## Dữ liệu hiện có

- `niengiam/NGTK2017.pdf` → `NGTK2024.pdf`: Niên giám Thống kê toàn quốc các năm 2017–2024 (Tổng cục Thống kê). Lưu ý: Niên giám năm X thường chứa số liệu đến năm X (ước tính) và các năm trước đó.

## Cấu trúc thư mục (quy ước)

- `niengiam/` — PDF nguồn (không sửa, không xóa).
- `data/csv/<năm>/` — dữ liệu đã trích theo từng niên giám; `data/csv/tonghop/` — chuỗi đã nối; `data/excel/` — bản .xlsx gương y cấu trúc csv (người dùng tổ chức lại 2026-07-21: mọi CSV nằm dưới `data/csv/`).
- `scripts/` — script Python trích xuất/xử lý.
- `docs/` — ghi chú về nguồn, phương pháp, định nghĩa chỉ tiêu.

## Quy ước dữ liệu

- **Định dạng đầu ra:** CSV, encoding UTF-8 (có BOM nếu cần mở bằng Excel), dạng long/tidy khi có thể: `tinh, ma_tinh, nam, chi_tieu, gia_tri, don_vi, nguon`.
- **Tên tỉnh:** chuẩn hóa theo danh mục 63 tỉnh thành (kèm mã tỉnh của TCTK nếu có). Thống nhất một cách viết (ví dụ "TP. Hồ Chí Minh", "Bà Rịa - Vũng Tàu"), tránh biến thể chính tả.
- **Số liệu:** dùng dấu chấm thập phân (kiểu quốc tế) trong CSV; ghi rõ đơn vị (nghìn người, tỷ đồng, %, ...). Số liệu trong PDF tiếng Việt dùng dấu phẩy thập phân và dấu chấm ngăn cách hàng nghìn — phải chuyển đổi cẩn thận.
- **Nguồn gốc:** mỗi bảng trích xuất phải ghi lại nguồn (file PDF, số trang, tên bảng) để kiểm chứng được.
- **Giá trị thiếu:** để trống hoặc ghi `NA`, không điền 0. Số liệu "sơ bộ"/"ước tính" cần cột ghi chú nếu phân biệt được.

## Kế hoạch thực hiện (chia giai đoạn theo kiểu bảng, KHÔNG theo số trang)

Nguyên tắc: PDF không đi qua context của Claude — script Python đọc PDF và ghi thẳng ra CSV. Token chỉ tốn cho việc viết/sửa parser và xem mẫu, nên chia phiên theo **số kiểu layout bảng cần xử lý**, không cắt file theo trang.

Vai trò định dạng file:
- **CSV** = nơi chứa dữ liệu (mỗi bảng một file, ghép bằng script). KHÔNG dùng Markdown chứa dữ liệu bảng.
- **Markdown** = ghi chú cho người và cho phiên sau: catalog, tiến độ, QA.

Các giai đoạn:
- **Giai đoạn 0 — Catalog (ĐÃ XONG 2026-07-20):** `scripts/00_catalog.py` quét 8 PDF bằng chữ ký font (tiêu đề VN = Bold 12, tiêu đề EN = Italic 12, thân bảng ~8.3) → `docs/catalog.csv`: 3.550 bảng, trong đó 1.206 bảng "phân theo địa phương". Xương sống của mọi phiên sau.
- **Giai đoạn 1 — Trích xuất (ĐÃ XONG cả 8 file, 2026-07-21):** `scripts/01_extract.py <năm>` → `data/csv/<năm>/*.csv` + `_meta.csv`. Tổng **1.207 CSV, 380.480 ô số liệu**. Đã xử lý: chú thích cuối trang, ô thiếu `..`, ô lệch dòng, header chứa số, TỔNG SỐ hoa/thường, nhãn song ngữ nhiều dòng, bảng trải ngang, bảng xếp chồng, cột thưa dẫn đầu, alias tên tỉnh. **Đọc mục "Bài học" (#1–18) trong docs/tien-do.md trước khi sửa parser hoặc thêm năm.**
- **Giai đoạn 2 — Kiểm định (ĐÃ XONG cả 8 file, 0 lỗi mọi năm):** `scripts/02_qa.py <năm>` → `docs/qa-<năm>.md`; kiểm 63 tỉnh/6 vùng/CẢ NƯỚC, tính cộng dồn, gắn cờ dòng bị watermark. Đã đối chiếu vision chọn mẫu mỗi năm.
- **Giai đoạn 3a+3b — Nối chuỗi (ĐÃ XONG 2026-07-21):** `scripts/03_tonghop.py` → `data/csv/tonghop/` = MỘT CỬA chứa toàn bộ dữ liệu: 261 file (149 wide tỉnh×năm `_tu-<y1>-<y2>` + 65 chồng-năm + 47 tách-năm `_nam-<y>`) + `_meta.csv`, phủ đủ 1.207/1.207 file nguồn. Quy tắc: trùng năm lấy niên giám mới nhất; 14 cặp đổi tên ghép theo `RENAME_MAP` (duyệt tại `docs/anh-xa-doi-ten.md`); bảng 2 khối khái niệm cùng năm (hộ nghèo, HIV) không nối kiểu wide; bảng trùng tên (cắt 60 ký tự) mang hậu tố `--2`/`--3`.
- **Giai đoạn 3c — Vá + ghép (ĐÃ XONG 2026-07-21):** `scripts/04_patch.py` vá 5 dòng lỗi watermark/text ẩn (giá trị đối chiếu vision, ghi trong docstring); 03_tonghop nâng cấp chia-nhóm-trong-chuỗi ghép xong 7 chuỗi đổi cấu trúc, hộ nghèo tách 2 chuẩn. Tonghop: 228 file + _meta, phủ đủ 1.207/1.207.
- **Thứ tự pipeline khi trích lại từ đầu:** `01_extract.py <năm>` → `04_patch.py` → `02_qa.py <năm>` → `03_tonghop.py` → `06_to_excel.py` (bản Excel trong `data/excel/`, cấu trúc gương với CSV).
- Việc tùy chọn còn lại: đặt tên cột chuỗi kiểu B khi cần dùng, xuất Excel, trích nhóm bảng cấp quốc gia (dia_phuong=0).

**Quy trình nối phiên (bắt buộc):** đầu phiên đọc `docs/tien-do.md` + `docs/catalog.csv` để biết đang ở đâu; cuối phiên cập nhật hai file đó. Phát hiện mới ảnh hưởng giải pháp thì cập nhật ngay vào CLAUDE.md.

## Lưu ý kỹ thuật

- **Đã kiểm chứng (2026-07-20):** cả 8 file NGTK 2017–2024 đều có text layer đầy đủ (không phải scan ảnh) → trích xuất trực tiếp bằng Python, KHÔNG cần chuyển sang ảnh + vision/OCR. Font TCVN3 lỗi chỉ xuất hiện ở bìa/tiêu đề chạy một vài trang, không ảnh hưởng dữ liệu bảng.
- **Phương pháp trích bảng đã kiểm chứng:** bảng NGTK không có đường kẻ đầy đủ nên `extract_tables()` mặc định của pdfplumber thất bại. Cách đúng: dùng `page.extract_words()` lấy từng từ kèm tọa độ (x0, top), gom dòng theo `top`, tách cột theo cụm `x0` (các cột số căn phải tại vị trí x ổn định trên toàn trang). Đã test OK trên bảng phân theo địa phương.
- PDF Niên giám có bảng phức tạp (nhiều dòng tiêu đề, gộp ô). Luôn đối chiếu xác suất vài giá trị với PDF gốc sau khi trích xuất.
- **Từ NGTK2022 bố cục đổi** (nhiều bảng hơn, chương đổi tên/tách): parser cần test riêng nhóm 2022–2024. NGTK2018 có lỗi font "ƣ" thay "ư" trong tiêu đề — chuẩn hóa khi trích.
- Chi tiết phát hiện từng giai đoạn: xem `docs/tien-do.md` (luôn đọc file này đầu phiên).
- Một số chỉ tiêu về CNTT/viễn thông theo tỉnh không có trong Niên giám toàn quốc — có thể cần thêm nguồn: Niên giám tỉnh, Sách trắng CNTT-TT (Bộ TT&TT), báo cáo chỉ số DTI, PCI, PAPI...
- Máy đang chạy Windows 11; shell chính là PowerShell.

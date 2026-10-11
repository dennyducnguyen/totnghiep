# Tiến độ dự án

> File này là cầu nối giữa các phiên làm việc. Đầu phiên: đọc file này + `catalog.csv`. Cuối phiên: cập nhật lại.

> **LƯU Ý ĐƯỜNG DẪN (từ 2026-07-21):** người dùng tổ chức lại `data/` thành 2 nhánh gương nhau: `data/csv/{2017..2024,tonghop}` và `data/excel/{...}`. Mọi đường dẫn `data/<năm>`, `data/tonghop` trong các ghi chú cũ bên dưới hiểu là `data/csv/<năm>`, `data/csv/tonghop`. Scripts đã cập nhật theo.

## Trạng thái tổng quan

| Giai đoạn | Mô tả | Trạng thái |
|---|---|---|
| 0 | Lập catalog bảng cho 8 file NGTK | ✅ XONG (2026-07-20) |
| 1 | Trích xuất bảng "phân theo địa phương" | ✅ XONG cả 8 file (2026-07-21) |
| 2 | Kiểm định (QA) | ✅ XONG cả 8 file — 0 lỗi mọi năm |
| 3a | Nối chuỗi liên tục → `data/tonghop/` | ✅ XONG (2026-07-21): 176 chuỗi |
| 3b | Ghép chuỗi đứt do đổi tên (14 cặp, user duyệt) | ✅ XONG — xem `docs/anh-xa-doi-ten.md` |
| 3c | Xử tay phần còn lại + xuất Excel | chưa (xem Việc treo) |

## Trạng thái theo file

| File | Bảng ĐP | File CSV | Ô số liệu | QA (đạt / cảnh báo / lỗi) |
|---|---|---|---|---|
| NGTK2017 | 134 | 134 | 43.026 | 108 / 26 / 0 |
| NGTK2018 | 138 | 138 | 43.624 | 109 / 29 / 0 |
| NGTK2019 | 138 | 138 | 43.736 | 107 / 31 / 0 |
| NGTK2020 | 138 | 138 | 44.145 | 106 / 32 / 0 |
| NGTK2021 | 138 | 139¹ | 44.055 | 108 / 30 / 0 |
| NGTK2022 | 169 | 169 | 53.201 | 127 / 42 / 0 |
| NGTK2023 | 170 | 170 | 53.576 | 126 / 44 / 0 |
| NGTK2024 | 181 | 181 | 55.117 | 144 / 37 / 0 |
| **Tổng** | **1206** | **1207** | **380.480** | **835 / 271 / 0** |

¹ 2021 có thêm file `_b2` (bảng chỉ số chè búp xếp chồng cùng trang).

"Cảnh báo" = hợp lệ về nghiệp vụ (bảng vốn không đủ 63 tỉnh, không có dòng CẢ NƯỚC, hoặc chỉ tiêu tỷ lệ không cộng dồn được). Xem chi tiết trong `docs/qa-<năm>.md`.

## Sản phẩm đã có

- `docs/catalog.csv` — danh mục 3.550 bảng (cột: nam, file, trang_dau, trang_cuoi, chuong, tieu_de_vn, tieu_de_en, dia_phuong, trang_thai). Mỗi dòng = 1 bảng, đã gộp các trang "(Tiếp theo)". Các bảng 2017 đã có trạng thái `da_trich`/`da_kiem`.
- `scripts/00_catalog.py` — sinh catalog (`python scripts/00_catalog.py`).
- `scripts/vn_names.py` — danh mục chuẩn 63 tỉnh / 6 vùng + hàm chuẩn hóa, phân loại nhãn (dùng chung).
- `scripts/01_extract.py` — parser trích bảng theo địa phương (`python scripts/01_extract.py 2017` hoặc thêm số trang để debug 1 bảng).
- `scripts/02_qa.py` — kiểm định (`python scripts/02_qa.py 2017`) → `docs/qa-<năm>.md`, cập nhật catalog.
- `data/2017/*.csv` — 134 bảng đã trích (cột: chi_muc, loai [CA_NUOC/VUNG/TINH/KHAC], các cột số liệu; số đã đổi sang dấu chấm thập phân; ô thiếu giữ nguyên `..`).
- `data/2017/_meta.csv` — siêu dữ liệu từng bảng (đơn vị tính, số tỉnh, trạng thái, ghi chú).
- `docs/qa-2017.md` — báo cáo kiểm định chi tiết.

## Phát hiện quan trọng trong Giai đoạn 0

1. **Nhận diện bảng bằng font rất đáng tin:** tiêu đề Việt = `Arial,Bold` size 12, tiêu đề Anh = `Arial,Italic` size 12, trang nối tiếp có "(Tiếp theo)"/"(Cont.)", header chương = Bold size 10, thân bảng size ~8.3. Chữ ký này ổn định trên CẢ 8 file 2017–2024 → parser Giai đoạn 1 có thể dựa vào đây.
2. **Trang "GIẢI THÍCH THUẬT NGỮ"** đầu mỗi chương dễ nhận nhầm là bảng — đã lọc bằng (a) yêu cầu ≥10 ô số liệu size nhỏ, (b) loại tiêu đề bắt đầu "GIẢI THÍCH"/"EXPLANATION".
3. **Từ NGTK2022 bố cục đổi:** thêm nhiều bảng (488–499 so với ~415), chương "Giáo dục" thành "Khoa học, Công nghệ và Giáo dục" (2023+), tách "Mức sống dân cư" nhiều bảng hơn. Số bảng theo địa phương tăng 138 → 169–181. Parser Giai đoạn 1 cần test riêng trên nhóm 2022–2024.
4. **3 dòng nhiễu còn lại trong catalog** (chấp nhận được, dia_phuong=0 nên không ảnh hưởng Giai đoạn 1): 2020 tr.35 và 2021 tr.31 (trang infographic tổng quan), 2022 tr.241 (trang bìa chương TKQG).
5. NGTK2018 có lỗi font chữ "ƣ" (ư kiểu cũ TCVN) trong một số tiêu đề — khi chuẩn hóa tên chỉ tiêu cần thay `ƣ→ư`, `ơ̛→ơ` (đã thấy "PHƢƠNG PHÁP").

## Bài học từ NGTK2017 (đọc kỹ trước khi chạy 2018–2024)

Các bẫy đã gặp và cách parser hiện tại xử lý (tất cả đã code trong `01_extract.py`/`vn_names.py`):

1. **Số liệu trong chú thích cuối trang** làm bẩn cụm cột → cắt các dòng bắt đầu `(*`/`Nguồn` ở nửa dưới trang, sát lề trái; cụm cột phải có đủ số ô (ngưỡng thích nghi).
2. **Ô thiếu `..` / `-` phải được coi là giá trị** — nếu không, dòng tỉnh toàn ô thiếu bị rớt và tên tỉnh đó dính sang dòng kế (sai nhãn dây chuyền).
3. **PDF có ô đặt lệch ~3pt** so với dòng (lỗi dàn trang gốc, vd Đà Nẵng tr.679) → gom dòng bằng gom cụm tọa độ top (ngưỡng 5pt), không nối dây chuyền.
4. **Header cột có thể chứa số** ("Dưới 0,5 tỷ đồng") → xác định dòng dữ liệu đầu tiên bằng nhận diện nhãn (CẢ NƯỚC/vùng/tỉnh + có số liệu), không dựa tọa độ.
5. **"TỔNG SỐ - TOTAL" viết HOA = dòng cả nước** (bảng FDI, du lịch, nhà ở...); "Tổng số" viết thường là header cột — phân biệt bằng chữ hoa.
6. **Nhãn song ngữ**: cùng dòng ("CẢ NƯỚC - WHOLE COUNTRY") hoặc tách 2–3 dòng (vùng dài) → phân loại bằng cắt ' - ' + so tiền tố; dòng EN mang số liệu, dòng VN chỉ có tên.
7. **Bảng trải ngang** (63 tỉnh lặp lại cho nhóm cột tiếp theo, vd Số DN theo quy mô tr.359, tr.378) → phát hiện CẢ NƯỚC xuất hiện lần 2, ghép ngang theo tên tỉnh.
8. **~22 bảng ít hơn 63 tỉnh là ĐÚNG với PDF gốc** (lúa hè thu ~30 tỉnh, tàu cá biển 29, sắn 40, đại học 45, FDI 59...). QA coi là cảnh báo, không phải lỗi.
9. **Vài bảng không có dòng CẢ NƯỚC** trong PDF (tỷ suất di cư thuần, chỉ số SCOLI) — hợp lệ.
10. **Kiểm tính cộng dồn**: bảng cộng được khớp <1%; bảng tỷ lệ/chỉ số/bình quân lệch ~500% (= tổng 6 vùng, đúng kỳ vọng); bảng doanh nghiệp lệch 2–18% có thể do phần "không phân bổ" trong số liệu gốc — không phải lỗi trích.
11. **Hạn chế còn lại (chấp nhận, xử ở Giai đoạn 3):** nhãn header cột của bảng nhiều tầng bị nối chuỗi lộn xộn (vd "chuẩn According Government's 2014") — GIÁ TRỊ đúng cột, chỉ tên cột xấu; khi hợp nhất sẽ đặt tên cột thủ công theo từng bảng.
12. PDF có in số hiệu bảng (font to: "18", "360"...) — chưa khai thác, có thể dùng làm ID ổn định giữa các năm.
13. **Dòng đơn vị tính có thể dài, chứa số và trải 2 dòng Việt+Anh** (vd "Trẻ em dưới một tuổi tử vong/1000 trẻ em sinh sống") → nhận diện: dòng header không chứa NĂM (19xx/20xx) và (ngắn ≤40 hoặc chứa `/` hoặc ` - `), gom liên tiếp, chặn từ khóa "Sơ bộ"/"Prel." (phát hiện nhờ vision check bảng b0107, 2026-07-21).

### Bài học bổ sung từ 2018–2024 (2026-07-21)

14. **Dòng CẢ NƯỚC tách rời giá trị** (NGTK2018 bảng đơn vị hành chính): tên "CẢ NƯỚC / WHOLE COUNTRY" nằm dòng riêng, giá trị ở dòng dưới → cho phép nhận dòng dữ liệu đầu tiên nếu 1–2 dòng sau là dòng thuần giá trị.
15. **2 bảng xếp chồng dọc trên cùng 1 trang** (NGTK2021 b0558 sản lượng chè + chỉ số chè): dòng CẢ NƯỚC xuất hiện lần 2 giữa trang → tách bảng thứ 2 ra file `<tên>_b2.csv`. Chỉ 1 ca/690 bảng đợt 1.
16. **Cột thưa dẫn đầu** (NGTK2022 b0983 giường bệnh, b0985 bác sĩ): cột "Tổng số" chỉ vài tỉnh có giá trị → rớt dưới ngưỡng cụm cột, làm ranh giới nhãn sai và giá trị dính vào tên tỉnh ("Bắc Ninh 4518"). **Sửa quan trọng:** ranh giới nhãn giờ xác định bằng KHOẢNG TRỐNG lớn giữa khối tên và khối số (khe ≥25pt), không dựa cột số mạnh nhất; thêm bước khôi phục cột thưa dẫn đầu (≥2 giá trị, nằm giữa nhãn và cột mạnh). Đã kiểm không hồi quy: số ô 2017–2021 giữ nguyên từng con số.
17. **Đổi tên tỉnh giữa các năm** (NGTK2024): "Thừa Thiên - Huế" viết gọn thành "Huế" → thêm alias trong `vn_names.py` (`_PROV_ALIASES`), vẫn quy về slot "Thừa Thiên - Huế" để nhất quán 63 tỉnh. Ảnh hưởng lớn: sửa xong đạt tăng 2→144 bảng ở 2024. **Lưu ý cho tương lai:** nếu bổ sung năm mới, kiểm tra tỉnh nào bị 62/63 hàng loạt = tên bị đổi.
18. **Watermark "TỔNG CỤC THỐNG KÊ" chèn cấp ký tự** (NGTK2024, 3 bảng: b0581, b0665, b1106): chữ watermark cùng size 9, upright=True, xen vào tên tỉnh làm rối ("CBTuắaycoê KBnạ..."). KHÔNG gỡ tự động an toàn được → QA gắn cờ mục riêng "cần nhập tay ở GĐ3". Mỗi bảng mất ~2–4 dòng tỉnh, phần còn lại vẫn đúng.

## Giai đoạn 3a+3b — Thư mục tổng hợp `data/tonghop/` (ĐÃ XONG 2026-07-21)

`data/tonghop/` = **MỘT CỬA chứa toàn bộ dữ liệu** (yêu cầu người dùng): 261 file CSV + `_meta.csv`, phủ đủ 1.207/1.207 file nguồn (đã đối soát tự động qua cột `nguon`). Sinh bởi `scripts/03_tonghop.py`:

- **149 chuỗi kiểu A** (`A_cot_nam`, wide: dòng = tỉnh, cột = năm dữ liệu; tên `<slug>_tu-<năm đầu>-<năm cuối>.csv`, chỉ 1 niên giám thì `_nam-<năm>.csv`). Quy tắc trùng năm: **niên giám mới nhất có giá trị không thiếu thắng**.
- **65 chuỗi kiểu B** (`B_chong_nam`: bảng cột-danh-mục, chồng các năm với cột `nam` đầu tiên; gồm cả bảng nhiều khối khái niệm cùng năm như HIV — giữ nguyên mọi cột).
- **47 file kiểu C** (`C_tach_nam`: 7 chuỗi đổi cấu trúc cột giữa các năm — tách riêng từng niên giám `_nam-<năm>.csv` để không trộn khái niệm; ghi chú trong meta).
- **14 cặp đổi tên đã ghép** qua `RENAME_MAP` (user duyệt — `docs/anh-xa-doi-ten.md`); chuỗi lấy tên mới của TCTK.
- Các bẫy đã xử trong 3a: (a) bảng có **2 cột cùng năm = 2 khối khái niệm** (hộ nghèo 2 chuẩn, HIV mới/lũy kế) → cấm nối kiểu A, ép kiểu B; (b) **va chạm slug** do tên file cắt 60 ký tự (các bảng "số DN đang hoạt động...", "thu nhập BQ...") → hậu tố thứ tự `--2`/`--3`, thứ tự bảng ổn định giữa các năm nên chuỗi vẫn nối đúng; (c) file `_b2` (chỉ số chè búp 2021) đưa vào với tên riêng `chi-so-phat-trien-...`.
- Dòng trong file wide xếp theo thứ tự chuẩn: CẢ NƯỚC → vùng → các tỉnh của vùng.

## Giai đoạn 3c — Vá watermark + ghép chuỗi đổi cấu trúc (ĐÃ XONG 2026-07-21)

**3c-1. Vá watermark (`scripts/04_patch.py`, chạy SAU 01_extract nếu trích lại):**
mọi giá trị vá đều đối chiếu vision với ảnh trang gốc, ghi rõ trong docstring script:
- 2024 tr.581: dòng rối = **Tuyên Quang** (Cao Bằng, Bắc Kạn không được in — không có trang trại, tổng vùng khớp 12 tỉnh).
- 2024 tr.666: dòng rối = **Cà Mau** (Bạc Liêu không được in).
- 2024 tr.1107 + 2023 tr.1103: xóa dòng text ẨN trùng lặp `4 905` (giá trị Cà Mau 2023 lọt vào text layer cả 2 năm, mắt thường không thấy).
- 2023: thêm dòng **Thái Bình trống** (PDF in nguyên dòng trắng — không có số liệu cháy nổ).
- QA sau vá: 2023 = 127 đạt/0 lỗi, 2024 = 145 đạt/0 lỗi.

**3c-2. Ghép 7 chuỗi đổi cấu trúc (nâng cấp `03_tonghop.py`):**
- Quy tắc cột-năm chặt hơn: header chứa ĐÚNG 1 năm (loại cột tỷ lệ "Năm 2018 so với 2017") → 2 chuỗi số DN đang hoạt động tự nối wide 2017–2024.
- **Chia nhóm trong chuỗi:** phần file đủ chuẩn cột-năm nối wide; phần còn lại gom khối theo cùng số cột + cùng TÚI TỪ header (chấp từ bị rơi sang cột kề, bỏ dấu, bỏ năm); không khớp thì tách năm. Header khối chồng xóa năm cụ thể (cột `nam` đã cho biết).
- Kết quả 7 chuỗi: DN×2 → wide 2017–2024; bác sĩ → wide 2018–2021 + file 2022 riêng; HIV nhiễm → khối 2018–2024 + 2017 riêng; trang trại lĩnh vực → 2 khối 2017–2020 và 2021–2024 (danh mục đổi); trường phổ thông → khối 2017–2023 + 2024 riêng (thêm cột); **hộ nghèo → tách 2 chuẩn**: `ty-le-ho-ngheo-theo-chuan-chinh-phu_tu-2010-2016` + `ty-le-ho-ngheo-theo-chuan-ngheo-da-chieu_tu-2016-2024`.
- Tonghop hiện tại: **228 file** (152 A + 66 B + 10 C/khác) + `_meta.csv`; đối soát vẫn đủ 1.207/1.207 nguồn.

## Bản Excel — `data/excel/` (ĐÃ XONG 2026-07-21)

`scripts/06_to_excel.py` chuyển **1.444 file CSV → 1.444 file .xlsx**, giữ nguyên cấu trúc thư mục (`excel/2017/...excel/tonghop/`, gồm cả `_meta`). Nội dung y nguyên; giá trị số ghi kiểu số thật (lọc/tính được ngay), ô thiếu `..` giữ dạng chữ; font Arial, dòng tiêu đề đậm + nền xanh, cố định dòng đầu, cột giãn tự động. Đã đối soát số lượng từng thư mục khớp 100%. Lưu ý: nếu chạy lại pipeline CSV thì chạy lại script này để đồng bộ bản Excel.

## Việc tiếp theo (tùy chọn)

- **Đặt lại tên cột** cho chuỗi kiểu B (header còn lộn xộn — hạn chế #11): làm khi cần dùng bảng cụ thể, tra `_meta.csv` + PDF gốc.
- Trích thêm ~2.340 bảng cấp quốc gia (dia_phuong=0) nếu nghiên cứu cần.

## Giao thức kiểm vision mỗi năm (bắt buộc, sau khi QA tự động đạt)

Vision KHÔNG dùng để trích dữ liệu (text layer đã chính xác) — chỉ dùng kiểm chọn mẫu,
vì có lớp lỗi QA tự động không thấy (vd đơn vị trộn vào tên cột ở b0107; giá trị nhầm cột
trong bảng tỷ lệ không kiểm cộng dồn được). Mỗi năm render 3–5 trang (fitz, zoom 2x) so với CSV:

1. Một bảng tỷ lệ/chỉ số (điểm mù của kiểm cộng dồn vùng).
2. Một bảng header nhiều tầng ("phân theo X và theo địa phương").
3. Một bảng nhiều trang "(Tiếp theo)".
4. Mọi bảng QA báo cảnh báo bất thường.

Riêng năm 2022 (bố cục mới đầu tiên): tăng lên 6–8 trang; 2023–2024 nếu 2022 sạch thì về mức thường.
Kết quả ghi vào `docs/qa-<năm>.md`.

## Nhật ký phiên

### 2026-07-20 — Phiên 1
- Tạo CLAUDE.md, khảo sát 8 PDF: đều có text layer, không cần OCR/vision.
- Chốt phương pháp trích bảng: `extract_words()` + gom cụm tọa độ (extract_tables mặc định thất bại vì bảng không kẻ đủ đường).
- ✅ Hoàn thành Giai đoạn 0: catalog 3.550 bảng / 8 file, trong đó 1.206 bảng phân theo địa phương.
- ✅ Giai đoạn 1+2 cho NGTK2017: trích 134/134 bảng → `data/2017/`, QA 0 lỗi. Đối chiếu vision 2 trang (79, 838) khớp 100%. Sửa 8 lỗi parser trong quá trình (xem mục Bài học).
- ⏸️ DỪNG theo yêu cầu: đợi người dùng ra lệnh mới chạy 2018–2024.

### 2026-07-21 — Phiên 2b: Vòng kiểm định tổng thể (sau khi xong 8 file)

Chạy 5 phép kiểm trên toàn bộ 1.207 CSV + 1 phép đối chiếu chéo giữa các niên giám:

1. **Cấu trúc file:** 0 dòng lệch số cột so với header.
2. **Nhãn rối/watermark (quét cả 8 năm):** ngoài 3 bảng 2024 đã biết, thêm 1 dòng ở 2023 `b1102` (cháy nổ) — do chính PDF in lặp giá trị cuối bảng, thành 1 dòng thừa không nhãn. Bỏ khi hợp nhất.
3. **Tỉnh trùng lặp:** 0 file.
4. **Sanity:** diện tích cả nước ~331,2–331,3 nghìn km² (đúng các lần điều chỉnh chính thức), dân số tăng đơn điệu 93,7 → 101,3 triệu (2017→2024).
5. **Định dạng ô:** 0 ô không hợp lệ trên 380.480 ô.
6. **Đối chiếu chéo giữa niên giám** (cùng chỉ tiêu, cột năm trùng, 4 dòng đại diện): 8.671 cặp giá trị — **95,6% khớp chính xác tuyệt đối**. 4,4% lệch đã xác minh là do **TCTK điều chỉnh số liệu giữa các kỳ** chứ không phải lỗi trích: (a) giá trị lệch đều tồn tại đúng trong PDF gốc từng năm (kiểm 6 ca bằng text gốc); (b) lệch dồn theo nguyên khối bảng — cụm dân số 2018→2019 trùng đợt điều chỉnh theo Tổng điều tra dân số 2019, cụm lâm nghiệp/chăn nuôi là điều chỉnh phương pháp ngành.

**Kết luận: dữ liệu đạt độ chính xác dùng được cho nghiên cứu.** Lưu ý cho phân tích: số liệu cùng một năm có thể khác nhau giữa các kỳ niên giám (sơ bộ → chính thức → điều chỉnh) — khi hợp nhất GĐ3 cần quy ước lấy số của niên giám MỚI NHẤT có chứa năm đó.

### 2026-07-21 — Phiên 2
- Vision kiểm b0107 (2017): phát hiện + sửa lỗi dòng đơn vị tính dài trộn vào tên cột (bài học #13).
- ✅ Giai đoạn 1+2 cho 2018–2024 (7 file còn lại). Kết quả: **1.207 CSV, 380.480 ô số liệu, 0 lỗi mọi năm.**
- Sửa 5 lỗi parser mới (bài học #14–18): dòng CẢ NƯỚC tách rời, bảng xếp chồng, cột thưa dẫn đầu, alias "Huế", watermark 2024.
- Đối chiếu vision: 2018 (b0038, b0116), 2022 (b0983, b0985), 2024 (b0103) — tất cả khớp.
- Kiểm không hồi quy: số ô 2017–2021 giữ nguyên từng con số sau khi sửa parser.
- Còn 3 bảng 2024 dính watermark cần nhập tay ở GĐ3.

## Việc treo / cần xử lý sau

- (nhỏ) 3 dòng nhiễu trong catalog (2020 tr.35, 2021 tr.31, 2022 tr.241) — dia_phuong=0, không ảnh hưởng.
- (GĐ3) Đặt lại tên cột cho các bảng header nhiều tầng khi hợp nhất.
- (GĐ3) Nhập tay 3 bảng watermark 2024: b0581, b0665, b1106 (xem qa-2024.md).
- (GĐ3) Bỏ dòng thừa không nhãn ở 2023 b1102 (PDF gốc in lặp giá trị cuối bảng).
- (GĐ3) Quy ước chọn số khi cùng năm xuất hiện ở nhiều niên giám: lấy niên giám mới nhất (số đã điều chỉnh).
- (GĐ3) Vài dòng KHAC là tên vùng bị PDF ngắt dòng lạ (vd "North Central and Central coastal") — gán lại nhãn khi hợp nhất, giá trị vẫn đúng.

## Phiên 2026-09-05 — Bài báo khoa học EBI (thư mục `detai/`)

- **Đầu vào:** `detai/EBI_NGTK_Panel_2017_2024.xlsx` (panel 63 tỉnh × 2017–2024, EBI của VECOM + biến NGTK), `detai/hdnc18062026.docx` (hướng dẫn: 3 RQ, H1–H5, Bảng 2–8), 3 bài mẫu trong `detai/bai-bao-mau/`.
- **Bổ sung từ kho `data/csv/tonghop/`:** hàng hóa vận chuyển đường bộ/người (logistics), siêu thị + TTTM/100.000 dân (ô trống NGTK = 0), DN đăng ký mới/1.000 dân. NGTK KHÔNG có PCI, logistics index, tỷ lệ Internet theo tỉnh–năm (chỉ 2024) → ghi hạn chế trong bài.
- **Điểm gãy EBI 2021:** VECOM đổi phương pháp (bỏ G2B, trọng số 0,2/0,4/0,4; trung bình 41,6 → 8,5). Xử lý: chỉ số tương đối (EBI/trung bình năm × 100), tách 2 giai đoạn cho hội tụ beta, hiệu ứng năm trong hồi quy.
- **Scripts:** `detai/scripts/01_phan_tich.py` (Bảng 1–8 + hình → `detai/ket-qua/`), `01b_hinh_khung.py`, `02_bang_docx.py` (→ `bang_docx.json`, số kiểu Việt), `noi_dung.js` (toàn văn bài), `03_build_docx.js` (docx-js; cần `NODE_PATH` trỏ tới node_modules có gói `docx`), `04_bo_sung_tai_lieu.py` (thêm 2 sheet vào `Nghien_Cuu_Tai_Lieu_EBI.xlsx`).
- **Sản phẩm:** `detai/Bai-bao-EBI-63-tinh-2017-2024.docx` (23 trang, ~9.500 từ thân bài, 8 bảng chính + 3 bảng phụ, 6 hình, 50 TLTK APA). Tác giả: Đào Thị Hồng Vân (chừa chỗ đơn vị/email).
- **Kiểm chứng tài liệu:** 5/6 bài "bổ sung 2020–2026" (E2–E6) sai hoặc không tồn tại (E6 DOI 404) → không dùng; 12 bài thay thế đã xác minh DOI, ghi ở sheet `Tai_Lieu_Thay_The_2026` và `Kiem_Chung_E1_E6`.
- **Kết quả chính:** Gini EBI 0,13–0,15 (2017–2020), 0,25–0,35 (2021–2024); phần trong-vùng chiếm ~2/3 Theil (bác H1 vế "giữa vùng chiếm ưu thế"); loại HN+HCM thì Theil giảm >50%; B2B+B2C đóng góp 70–80% Gini (H2); beta tuyệt đối −0,043 (2017–20) và −0,089 (2021–24), có điều kiện không ý nghĩa (H3 một phần); mật độ DN là yếu tố mạnh nhất (beta chuẩn hóa 0,48), hộ có máy tính không ý nghĩa (H4 một phần); Oaxaca–Blinder: ~79–80% khoảng cách lõi–ngoại vi do nguồn lực, riêng mật độ DN ~69% (H5).
- **Bổ sung 2026-09-05 (phiên 2):** gom phần tạo hình thành bộ tự chứa `detai/hinh-bai-bao/` = `tao_hinh.py` (6 hình, mỗi hình một hàm) + `chay_tao_hinh.bat` + `du_lieu/panel_63_tinh_2017_2024.csv` (bản sao `ket-qua/panel_phan_tich.csv`) + thư mục kết quả `hinh/` + `HUONG-DAN-TAO-HINH.docx` (sinh bởi `scripts/05_huong_dan_hinh.js`). `01_phan_tich.py` KHÔNG còn vẽ hình; `03_build_docx.js` lấy hình từ `hinh-bai-bao/hinh/` (tên file Hinh1…Hinh6 khớp số hình trong bài); biến môi trường `OUT_FILE` để ghi docx ra chỗ khác khi file gốc đang mở trong Word.
- **Phiên bản v2 (2026-09-05, phiên 3) — `detai/v2/`:** bản sao đầy đủ của `detai/` (giữ nguyên bản cũ) với: (a) 8 hình riêng, không ghép tiểu đồ (Hình 3 Theil mức, Hình 4 tỷ trọng giữa vùng, Hình 6/7 beta hai giai đoạn), mỗi hình có chú thích và một đoạn giải thích riêng trong bài; `v2/hinh-bai-bao/tao_hinh.py` + hướng dẫn cập nhật; (b) `v2/Bang-doi-chieu-trich-dan.xlsx` (sinh bởi `v2/scripts/06_bang_trich_dan.py`): 50 TLTK × [tên bài, tác giả, RQ gốc, mô hình gốc, đoạn gốc, câu trích trong bài (tự động từ noi_dung.js), dạng APA in-text, mục, APA đầy đủ, trạng thái kiểm chứng tô màu]; (c) `01_phan_tich.py` tự tìm thư mục `data/csv/tonghop` bằng cách đi ngược lên. `v2/scripts/node_modules` đã có gói docx → chạy thẳng `node scripts/03_build_docx.js` trong v2.
- **Bài báo v6 – bản "v2" công thức (2026-09-13):** `detai/v6/Bai-bao-EBI-63-tinh-2017-2024-v2.docx` (file cũ giữ nguyên). Mục 3.1 và 3.3 viết lại với 16 phương trình Word gốc (OMML) đánh số + danh sách "trong đó" giải thích ký hiệu; ký hiệu Theil-T đổi T → Th. Dựng bằng `v6/scripts/03_build_docx_v2.js` (+ `noi_dung_v2.js`, `cong_thuc_v2.js`, `cong_thuc_omml.js`, `ghi_chu_bang_v2.js`). Bài học kỹ thuật: API Math của docx-js thiếu gạch trên/mũ/chữ đứng → tự sinh chuỗi OMML, nhúng bằng `ImportedXmlComponent.fromXmlString(xml).root[0]` (thiếu `.root[0]` sinh thẻ `<undefined>`); phương trình hiển thị đặt trong bảng 3 cột không viền để số (n) căn phải. Chi tiết: `detai/v6/GHI-CHU-V6.txt` mục 9.
- **Phiên bản v7 (2026-09-14) — `detai/v7/`:** bản sao đầy đủ của v6 + chỉnh sửa theo phản biện "tinh gọn phương pháp" (yêu cầu 14/09/2026). Sản phẩm: `BAO-CAO-CHINH-SUA-V7.docx` (chẩn đoán v6, thiết kế mới, bảng đánh giá 17 phương pháp, kế hoạch chỉnh sửa, phân loại tài liệu, bảng kiểm 15 mục; 12 trang) và `Bai-bao-EBI-63-tinh-2017-2024-v7.docx` (25 trang, 34 TLTK, 7 bảng + 5 bảng phụ lục, 4 hình, 8 + 5 phương trình OMML). Hai phát hiện then chốt: (S1) từ khi đổi phương pháp, EBI chứa chỉ tiêu DN/dân số (B2B), thu nhập (B2C), tên miền/dân số (NNL&HT) theo Báo cáo EBI 2023 tr. 49–53 → biến "mật độ doanh nghiệp" của v6 trùng định nghĩa một phần với biến phụ thuộc; (S2) hội tụ beta của v6 mất ý nghĩa khi loại HN & TP.HCM. Thiết kế v7: RQ1 Gini + Theil + SD ln; RQ2 phân rã Theil 6 vùng; RQ3 FE tỉnh + năm (Hausman χ²=72,4), kiểm tra 2021–2024, hạng phần trăm, bỏ mật độ DN, OLS 2018–2020; loại beta có điều kiện, Oaxaca–Blinder, biến giả vùng + Wald, VIF, hệ số chuẩn hóa, H1–H5; mục riêng về điểm gãy 2021 và dữ liệu thiếu (445 quan sát, 54–58 tỉnh/năm, 53 tỉnh đủ 8 năm). Pipeline v7: `scripts/01_phan_tich_v7.py` → `hinh-bai-bao/tao_hinh_v7.py` → `02_bang_docx_v7.py` → `node 03_build_docx_v7.js [mô-đun]` → `16_docx_sang_pdf_anh.py` (Word COM) → `06/10/11_*_v7.py` (bảng đối chiếu trích dẫn v7). Chi tiết: `detai/v7/GHI-CHU-V7.txt`. Các file phụ trợ của v6 (mô tả công thức, hướng dẫn từng bước, sổ Excel toàn quy mô) chưa viết lại theo v7. Điểm gãy đã xác nhận bằng `Dulieu_EBI_2017_2025.xlsx`: kỳ báo cáo 2021 là kỳ đầu tiên 3 trụ cột + thang mới; kỳ 2025 đổi trọng số 0,1/0,6/0,3 (không dùng).

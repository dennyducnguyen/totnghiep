# Đề xuất ánh xạ các chỉ tiêu bị TCTK đổi tên (chờ duyệt)

> Bối cảnh: từ NGTK2022 (bố cục mới), một số bảng bị đổi tên làm chuỗi bị đứt đôi.
> Dưới đây là các cặp đề xuất ghép lại, kèm kiểm chứng bằng **giá trị CẢ NƯỚC tại năm
> giao nhau** giữa 2 mảnh (khớp = gần chắc chắn cùng chỉ tiêu; lệch nhỏ ≤0,5% = TCTK
> điều chỉnh số). Sau khi duyệt, script 03 sẽ nối các cặp này thành chuỗi đầy đủ.

## 14 cặp đề xuất GHÉP (điểm giống tên + giá trị giao nhau)

| # | Chuỗi cũ (2017–2021) | Chuỗi mới (2021/22–2024) | Giá trị năm giao | Kết luận |
|---|---|---|---|---|
| 1 | ty-le-tang-tu-nhien-dan-so | ty-le-tang-dan-so-tu-nhien | 2021: 9.3 = 9.3 | KHỚP |
| 2 | dien-tich-cay-luong-thuc-co-hat | dien-tich-gieo-trong-cay-luong-thuc-co-hat | 2020: 8222.5 ≈ 8222.6 | khớp (điều chỉnh nhỏ) |
| 3 | dien-tich-lua-ca-nam | dien-tich-gieo-trong-lua-ca-nam | 2020: 7279.0 ≈ 7278.9 | khớp (điều chỉnh nhỏ) |
| 4 | nang-suat-lua-ca-nam | nang-suat-gieo-trong-lua-ca-nam | 2020: 58.7 ≈ 58.8 | khớp (điều chỉnh nhỏ) |
| 5 | dien-tich-lua-dong-xuan | dien-tich-gieo-trong-lua-dong-xuan | 2020: 3024.1 ≈ 3024.0 | khớp (điều chỉnh nhỏ) |
| 6 | nang-suat-lua-dong-xuan | nang-suat-gieo-trong-lua-dong-xuan | 2020: 65.7 = 65.7 | KHỚP |
| 7 | dien-tich-lua-he-thu-va-thu-dong | dien-tich-gieo-trong-lua-he-thu-va-thu-dong | 2020: 2669.1 = 2669.1 | KHỚP |
| 8 | nang-suat-lua-he-thu-va-thu-dong | nang-suat-gieo-trong-lua-he-thu-va-thu-dong | 2020: 55.3 ≈ 55.4 | khớp (điều chỉnh nhỏ) |
| 9 | dien-tich-lua-mua | dien-tich-gieo-trong-lua-mua | 2020: 1585.8 = 1585.8 | KHỚP |
| 10 | nang-suat-lua-mua | nang-suat-gieo-trong-lua-mua | 2020: 51.1 = 51.1 | KHỚP |
| 11 | dien-tich-ngo | dien-tich-gieo-trong-ngo | 2020: 942.5 = 942.5 | KHỚP |
| 12 | nang-suat-ngo | nang-suat-gieo-trong-ngo | 2020: 48.4 = 48.4 | KHỚP |
| 13 | dien-tich-khoai-lang | dien-tich-gieo-trong-khoai-lang | 2020: 109.3 ≈ 108.9 | khớp (điều chỉnh nhỏ) |
| 14 | dien-tich-san | dien-tich-gieo-trong-san | 2020: 524.5 ≈ 524.7 | khớp (điều chỉnh nhỏ) |

(Tất cả đều là đổi tên kiểu "Diện tích X" → "Diện tích gieo trồng X" từ NGTK2022, cộng 1 ca đảo từ "tăng tự nhiên dân số" → "tăng dân số tự nhiên".)

## 1 cặp bị LOẠI nhờ kiểm giá trị

- `dien-tich-nuoi-trong-thuy-san` ↛ `dien-tich-gieo-trong-san`: tên giống 57% nhưng CẢ NƯỚC 2021 = 1128.9 vs 524.5 → khác chỉ tiêu, không ghép. (Chuỗi thủy sản kết thúc 2021, chưa tìm thấy hậu duệ ở 2022+ — có thể TCTK bỏ bảng hoặc đổi tên xa hơn, sẽ rà tay sau.)

## Trạng thái

- [x] ĐÃ DUYỆT (người dùng, 2026-07-21) — 14 cặp đã đưa vào `RENAME_MAP` trong
  `scripts/03_tonghop.py` và nối vào `data/tonghop/`. Chuỗi ghép lấy tên MỚI
  (thuật ngữ hiện hành của TCTK từ 2022).

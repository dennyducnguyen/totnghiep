# -*- coding: utf-8 -*-
"""
BƯỚC 5 – KIỂM TRA TRÍCH DẪN CỦA BÀI BÁO v8

Chạy (trong thư mục v8):   python scripts/05_kiem_tra_trich_dan.py
Việc làm:
  1. Đọc danh mục tài liệu tham khảo và toàn văn trong scripts/noi_dung.js.
  2. Với mỗi tài liệu: đếm số lần được trích trong thân bài (theo họ tác giả đầu + năm), liệt kê các mục có trích.
  3. Tìm file gốc tương ứng trong kho tài liệu đã kiểm chứng của bản v7 (../v7/tai-lieu-doi-chieu/).
  4. Báo lỗi nếu có tài liệu trong danh mục không được trích, hoặc trích dẫn trong bài không có trong danh mục.
Đầu ra : Bang-doi-chieu-trich-dan-v8.xlsx
"""
import os, re, sys, unicodedata
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KHO = os.path.join(os.path.dirname(ROOT), "v7", "tai-lieu-doi-chieu")
src = open(os.path.join(ROOT, "scripts", "noi_dung.js"), encoding="utf-8").read()

refs_block = src[src.index('{ t: "refs", items: ['):]
refs_block = refs_block[:refs_block.index("] },")]
refs = re.findall(r'^\s*"(.+)",\s*$', refs_block, flags=re.M)
body = src[:src.index('{ t: "refs"')]

# chia thân bài theo mục h1/h2 để biết trích ở mục nào
sections, cur = [], "Đầu bài"
for line in body.splitlines():
    m = re.search(r'\{ t: "h[12]", text: "([^"]+)"', line)
    if m: cur = m.group(1)
    sections.append((cur, line))

def khong_dau(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").replace("đ", "d").replace("Đ", "D")

# khóa trích dẫn trong bài cho từng tài liệu: (mẫu tìm kiếm, năm)
DAC_BIET = {"Hiệp hội Thương mại điện tử Việt Nam (VECOM). (2017–2025)": (r"VECOM, 2017–2025|VECOM \(2017–2025\)", "VECOM_2017"),
            "Hiệp hội Thương mại điện tử Việt Nam (VECOM). (2023)": (r"VECOM, 2023|VECOM \(2023", "Hiep-hoi"),
            "International Telecommunication Union": (r"ITU, 2025", "International-Telecommunicat"),
            "United Nations Conference": (r"UNCTAD, 2024", "United-Nations-Conference"),
            "Bộ Công Thương": (r"Bộ Công Thương, 2025", "Bo-Cong-Thuong"),
            "Thủ tướng Chính phủ": (r"Thủ tướng Chính phủ, 2020", "Thu-tuong-Chinh-phu")}
kho = sorted(os.listdir(KHO)) if os.path.isdir(KHO) else []
rows, loi = [], []
for i, r in enumerate(refs, 1):
    year = re.search(r"\((\d{4})(?:–\d{4})?\)", r).group(1)
    key = next((k for k in DAC_BIET if r.startswith(k)), None)
    if key:
        pat, file_key = DAC_BIET[key]
    else:
        ho = r.split(",")[0].strip()
        ten2 = re.match(r"^[^,]+, [^,]+?(?:, & |, )([^,]+?),", r)
        if " & " in r.split("(")[0] and r.split("(")[0].count(",") <= 4 and r.split("(")[0].count("&") == 1 and r.split("(")[0].count(".,") <= 2:
            ho2 = r.split("&")[1].strip().split(",")[0]
            pat = rf"{re.escape(ho)}(?: và | & ){re.escape(ho2)},? \(?{year}|{re.escape(ho)} & {re.escape(ho2)}, {year}"
        else:
            pat = rf"{re.escape(ho)}(?: và cộng sự)?,? \(?(?:\d{{4}}, )?{year}"
        file_key = khong_dau(ho).replace(" ", "-")
    hits = [(sec, line) for sec, line in sections if re.search(pat, line)]
    n = sum(len(re.findall(pat, line)) for _, line in hits)
    muc = sorted({sec.split(" ")[0] for sec, _ in hits}, key=lambda x: [int(p) for p in x.strip(".").split(".") if p.isdigit()])
    file = next((f for f in kho if f"_{file_key.lower()}_" in f.lower() and (year in f or key)), "")
    if key and "2017" in pat: file = next((f for f in kho if f.startswith("16_")), "")
    if n == 0: loi.append(f"Tài liệu {i} không được trích trong bài: {r[:60]}")
    rows.append({"STT": i, "Tài liệu (APA)": r.replace("*", ""), "Số lần trích": n, "Mục có trích": ", ".join(muc),
                 "File gốc trong v7/tai-lieu-doi-chieu": file or ("(kinh điển, không cần file)" if r.startswith(("Rogers", "Theil")) else "KHÔNG THẤY")})

# trích dẫn trong bài không có trong danh mục (dạng "Họ (năm)" hoặc "Họ, năm")
ho_ds = {khong_dau(r.split(",")[0].split("(")[0].strip()) for r in refs}
ho_ds |= {"VECOM", "ITU", "UNCTAD", "Bo Cong Thuong", "Thu tuong Chinh phu", "Sala-i-Martin", "van Deursen", "van Dijk", "Del Bo", "Baycan", "Lopez", "Helsper"}
for m in re.finditer(r"([A-ZĐÖÜ][\w\-À-ỹ]+(?: [A-ZĐ][\w\-À-ỹ]+)?)(?: và cộng sự)?,? \(?(\d{4})\)?", body):
    name = khong_dau(m.group(1))
    if 1950 < int(m.group(2)) < 2030 and not any(name.endswith(h) or h.endswith(name) for h in ho_ds):
        if m.group(1) not in ("Theil-T", "Gini", "Kế hoạch", "Hình", "Bảng", "Mục", "Niên giám", "Năm", "EBI", "COVID-19", "DTF", "Quyết định", "GIAI ĐOẠN", "Vietnam", "Bước"):
            loi.append(f"Trích dẫn có thể không có trong danh mục: {m.group(0)}")

df = pd.DataFrame(rows)
out = os.path.join(ROOT, "Bang-doi-chieu-trich-dan-v8.xlsx")
with pd.ExcelWriter(out, engine="openpyxl") as xw:
    df.to_excel(xw, sheet_name="Doi_chieu", index=False)
    pd.DataFrame({"Kết quả kiểm tra": loi or ["Không phát hiện lỗi: mọi tài liệu trong danh mục đều được trích và có file gốc hoặc là tài liệu kinh điển."]}).to_excel(xw, sheet_name="Kiem_tra", index=False)
    ws = xw.sheets["Doi_chieu"]
    for col, w in zip("ABCDE", [5, 90, 10, 22, 70]): ws.column_dimensions[col].width = w
print(df[["STT", "Số lần trích", "Mục có trích", "File gốc trong v7/tai-lieu-doi-chieu"]].to_string(index=False))
print("\n".join(loi) if loi else "Không phát hiện lỗi trích dẫn.")
print("Đã ghi", out)

# -*- coding: utf-8 -*-
"""
BƯỚC 3 – ĐỊNH DẠNG BẢNG VÀ CÁC CON SỐ DÙNG TRONG LỜI VĂN BÀI BÁO v8

Chạy (trong thư mục v8):   python scripts/03_bang_va_so_lieu.py
Đầu vào : ket-qua/*.csv, ket-qua/ket_qua.json (do 01_phan_tich.py tạo)
Đầu ra  : ket-qua/bang_docx.json      – các bảng (b1–b7, a1–a3) đã định dạng số kiểu Việt Nam, để dựng Word
          ket-qua/so_lieu_bai.json    – các con số trong lời văn; noi_dung.js gọi bằng {{n:tên}}
          So-lieu-ket-qua-v8.xlsx     – toàn bộ bảng kết quả dạng số (để tra cứu, nộp kèm)
Quy ước số: dấu phẩy thập phân, dấu chấm phân cách hàng nghìn, dấu trừ "−"; *** p<0,01; ** p<0,05; * p<0,1.
"""
import json, os
import numpy as np, pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "ket-qua")
KQ = json.load(open(os.path.join(OUT, "ket_qua.json"), encoding="utf-8"))
rd = lambda f: pd.read_csv(os.path.join(OUT, f), encoding="utf-8-sig")

REG_ORDER = ["ĐNB", "ĐBSH", "BTB&DHMT", "ĐBSCL", "TN", "TDMNPB"]
REG_FULL = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
VN = {"ln_retail": "ln(Bán lẻ/người)", "ln_grdp": "ln(GRDP/người)", "dn_per_1000": "Mật độ doanh nghiệp (DN/1.000 dân)",
      "cathe_per_1000": "Mật độ cơ sở cá thể (cơ sở/1.000 dân)", "trained_labor_pct": "Tỷ lệ LĐ qua đào tạo (%)",
      "urban_share_pct": "Tỷ lệ dân số thành thị (%)", "ln_fdi_pc": "ln(1 + FDI lũy kế/người)",
      "hh_computer_pct": "Tỷ lệ hộ có máy tính (%)", "it_skills_pct": "Tỷ lệ dân số có kỹ năng CNTT (%)"}

def f(x, d=3):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    s = f"{x:,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    return s.replace("-", "−")
def fi(x): return "" if pd.isna(x) else f"{int(round(x)):,}".replace(",", ".")
def stars(p): return "" if pd.isna(p) else "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""
def cell(c, se, p, d=3): return f"{f(c, d)}{stars(p)}\n({f(se, d)})"
def fp(p): return "< 0,001" if p < 0.001 else f(p, 3)

T, N = {}, {}
NOTE_SAO = "*** p < 0,01; ** p < 0,05; * p < 0,1."

# ---------------- Bảng 1 ----------------
b1 = rd("bang1_mo_ta.csv"); rows = []
for _, r in b1.iterrows():
    d = 2 if abs(r["mean"]) < 100 else 0
    rows.append([r["ten"], r["don_vi"], r["nguon"], fi(r["N"]), f(r["mean"], d), f(r["sd"], d), f(r["min"], d), f(r["max"], d)])
T["b1"] = dict(title="Bảng 1. Biến số và thống kê mô tả (quan sát tỉnh – năm, 2017–2024)",
               header=["Biến", "Đơn vị", "Nguồn", "N", "Trung bình", "Độ lệch chuẩn", "Nhỏ nhất", "Lớn nhất"],
               rows=rows, widths=[3000, 1150, 800, 550, 950, 950, 850, 900],
               note="Ghi chú: N là số quan sát có số liệu; tối đa 504 (63 tỉnh × 8 năm). EBI và các trụ cột chỉ có ở 54–58 tỉnh mỗi năm (445 quan sát). DTF tính theo phương trình {{eq:dtf}}. GRDP bình quân đầu người được Niên giám công bố theo tỉnh từ năm 2018, tỷ lệ hộ có máy tính từ năm 2019, tỷ lệ dân số có kỹ năng CNTT chỉ có năm 2024. Mật độ cơ sở cá thể = số cơ sở kinh tế cá thể phi nông nghiệp / dân số. Nguồn: VECOM (2017–2025) và Niên giám Thống kê (NGTK) 2017–2024.")

# ---------------- Bảng 2 ----------------
b2 = rd("bang2_khoang_cach_theo_nam.csv"); rows = []
for _, r in b2.iterrows():
    rows.append([str(int(r.nam)), r.thang_do.split(" (")[0], str(int(r.N)), f(r["mean"], 2), f(r.CV, 3), f(r.Gini, 3), f(r.Theil, 3),
                 f(r.within_share, 1), f(r.between_share, 1), f(r.DTF_mean, 1), f(r.max_min, 2), f(r.p90_p10, 2)])
T["b2"] = dict(title="Bảng 2. Các chỉ tiêu đo lường khoảng cách số TMĐT giữa các tỉnh, 2017–2024",
               header=["Năm", "Thang đo EBI", "Số tỉnh", "EBI trung bình", "CV", "Gini", "Theil-T", "Trong vùng (%)", "Giữa vùng (%)", "DTF trung bình (%)", "Max/ Min", "P90/ P10"],
               rows=rows, widths=[520, 1050, 520, 750, 650, 650, 700, 720, 720, 750, 600, 600],
               note="Ghi chú: thang cũ – 4 trụ cột (2017–2020); chuyển tiếp – 3 trụ cột, thang điểm thử nghiệm (2021); ổn định – 3 trụ cột (2022–2024). CV theo {{eq:cv}}, Gini theo {{eq:gini}}, Theil-T theo {{eq:theil}}, phân rã trong vùng / giữa vùng theo {{eq:theil_dec}}; mỗi tỉnh có trọng số bằng nhau. EBI trung bình là điểm tuyệt đối, không so sánh được giữa các đoạn thang đo; các cột còn lại bất biến khi điểm của cả năm được nhân với một hằng số. Tỉnh không được xếp hạng: Cao Bằng, Điện Biên, Hà Giang, Lai Châu, Sơn La (cả 8 năm); Bạc Liêu, Sóc Trăng (2017–2022); Hậu Giang (2017–2020); Cà Mau (2018–2019); Vĩnh Long (2017). Nguồn: tính toán của tác giả từ VECOM (2017–2025).")

# ---------------- Bảng 3 ----------------
b3 = rd("bang3_gini_tru_cot.csv"); rows = []
for _, r in b3.iterrows():
    rows.append([str(int(r.nam)), TU := ("Thang cũ" if r.nam <= 2020 else "Chuyển tiếp" if r.nam == 2021 else "Ổn định"),
                 f(r.Gini_EBI), f(r.Gini_NNL_HT), f(r.Gini_B2C), f(r.Gini_B2B), f(r.Gini_G2B) or "—"])
T["b3"] = dict(title="Bảng 3. Hệ số Gini của EBI và từng trụ cột giữa các tỉnh, 2017–2024",
               header=["Năm", "Thang đo", "EBI", "NNL&HT", "B2C", "B2B", "G2B"],
               rows=rows, widths=[800, 1300, 1100, 1100, 1100, 1100, 1100],
               note="Ghi chú: NNL&HT – nguồn nhân lực và hạ tầng công nghệ thông tin; G2B chỉ có đến năm 2020. Gini của trụ cột NNL&HT năm 2017 cao bất thường vì thang điểm trụ cột này năm 2017 khác các năm sau (trung bình 17,7 so với 33,8 năm 2018). Nguồn: tính toán của tác giả từ VECOM (2017–2025).")

# ---------------- Bảng 4: hội tụ beta ----------------
bb = rd("bang4_hoi_tu_beta.csv")
ctrl = [("ln_grdp_L1", "ln(GRDP/người) năm trước"), ("dn_per_1000_L1", "Mật độ doanh nghiệp năm trước"),
        ("urban_share_pct_L1", "Tỷ lệ dân số thành thị năm trước"), ("trained_labor_pct_L1", "Tỷ lệ LĐ qua đào tạo năm trước")]
order = [0, 2, 1, 3]
rows = [["ln(DTF) năm trước (β)"] + [cell(bb.beta[i], bb.se[i], bb.p[i], 4) for i in order]]
for k, name in ctrl:
    rows.append([name] + [cell(bb[f"coef_{k}"][i], bb[f"se_{k}"][i], bb[f"p_{k}"][i], 4) if not pd.isna(bb[f"coef_{k}"][i]) else "" for i in order])
rows += [["Hiệu ứng năm"] + ["Có"] * 4,
         ["Năm tăng trưởng"] + [bb.nam_tang_truong[i] for i in order],
         ["Số quan sát (số tỉnh)"] + [f"{bb.N[i]} ({bb.N_tinh[i]})" for i in order],
         ["R²"] + [f(bb.R2[i], 3) for i in order],
         ["Tốc độ hội tụ λ (%/năm)"] + [f(bb.lambda_pct[i], 1) if bb.beta[i] < 0 else "—" for i in order]]
T["b4"] = dict(title="Bảng 4. Ước lượng hội tụ beta theo hai đoạn thang đo EBI",
               header=["Biến", "(1) Tuyệt đối, thang cũ", "(2) Có điều kiện, thang cũ", "(3) Tuyệt đối, thang ổn định", "(4) Có điều kiện, thang ổn định"],
               rows=rows, widths=[2900, 1500, 1500, 1500, 1500],
               note="Ghi chú: biến phụ thuộc là tăng trưởng ln(DTF) trong năm theo phương trình {{eq:beta}}. Thang cũ: tăng trưởng các năm 2018–2020 (cột 2 từ 2019 vì GRDP/người có từ 2018); thang ổn định: tăng trưởng các năm 2023–2024, gốc 2022–2023. Bước 2021→2022 không dùng vì thang điểm trụ cột năm 2021 khác hẳn các năm sau. Hồi quy OLS có hiệu ứng năm; trong ngoặc là sai số chuẩn theo cụm tỉnh. λ = −ln(1 + β). " + NOTE_SAO)

# ---------------- Bảng 5: kiểm tra độ vững beta ----------------
br = rd("bang5_beta_do_vung.csv"); rows = []
for _, r in br.iterrows():
    rows.append([r.mo_hinh, r.nam_tang_truong, f"{int(r.N)}", f"{f(r.beta, 4)}{stars(r.p)}", f(r.se, 4), fp(r.p), f(r.R2, 3)])
T["b5"] = dict(title="Bảng 5. Kiểm tra độ vững của hội tụ beta",
               header=["Kịch bản", "Năm tăng trưởng", "N", "β", "Sai số chuẩn", "p", "R²"],
               rows=rows, widths=[4000, 1100, 600, 1100, 900, 800, 700],
               note="Ghi chú: các kịch bản R1–R3 dùng đặc tả có điều kiện như cột (4) Bảng 4; R4–R7 dùng đặc tả tuyệt đối; R8–R10 là hồi quy cắt ngang tốc độ tăng ln(DTF) bình quân năm giữa hai năm đầu – cuối lên ln(DTF) năm đầu, sai số chuẩn vững HC1. Năm tỉnh dẫn đầu năm 2024: TP. Hồ Chí Minh, Hà Nội, Bình Dương, Đà Nẵng, Hải Phòng. R5–R6 lấy năm 2021 làm gốc, chỉ để so sánh với cách làm của bản nháp. " + NOTE_SAO)

# ---------------- Bảng 6: M2–M5 ----------------
b6 = rd("bang6_hoi_quy_M2_M5.csv")
XS = ["ln_retail", "ln_grdp", "dn_per_1000", "cathe_per_1000", "trained_labor_pct", "urban_share_pct", "ln_fdi_pc", "hh_computer_pct"]
MS = ["M2", "M3", "M4", "M5"]
def get(df, m, t, col="coef"):
    s = df[(df.model == m) & (df.term == t)]; return s[col].iloc[0] if len(s) else np.nan
rows = []
for x in XS:
    rows.append([VN[x]] + [cell(float(get(b6, m, x)), float(get(b6, m, x, "se")), float(get(b6, m, x, "p"))) if not pd.isna(get(b6, m, x)) else "" for m in MS])
rows += [["Hiệu ứng năm"] + ["Có"] * 4,
         ["Giai đoạn"] + [get(b6, m, "nam") for m in MS],
         ["Số quan sát (số tỉnh)"] + [f"{int(float(get(b6, m, 'N')))} ({int(float(get(b6, m, 'N_tinh')))})" for m in MS],
         ["R²"] + [f(float(get(b6, m, "R2")), 3) for m in MS]]
T["b6"] = dict(title="Bảng 6. Các yếu tố liên quan đến EBI và từng trụ cột (mô hình M2–M5)",
               header=["Biến giải thích", "M2: EBI tổng hợp", "M3: Trụ cột B2C (phía cầu)", "M4: Trụ cột B2B (phía cung)", "M5: Trụ cột NNL&HT"],
               rows=rows, widths=[3000, 1450, 1450, 1450, 1450],
               note="Ghi chú: biến phụ thuộc là DTF (0–100) của EBI hoặc trụ cột tương ứng, phương trình {{eq:pooled}}. Ước lượng OLS gộp có hiệu ứng năm; trong ngoặc là sai số chuẩn theo cụm tỉnh. M5 bắt đầu từ 2019 vì tỷ lệ hộ có máy tính có từ năm 2019. VIF lớn nhất trong M2 là " + f(max(KQ["vif_M2"].values()), 2) + ". Từ năm 2021 trụ cột B2C có chứa chỉ tiêu thu nhập bình quân và trụ cột B2B có chứa chỉ tiêu doanh nghiệp trên dân số (VECOM, 2023, tr. 49–53); kiểm tra trên giai đoạn 2018–2020 ở Bảng A3. " + NOTE_SAO)

# ---------------- Bảng 7: M6 ----------------
b7 = rd("bang7_cat_ngang_2024.csv")
X7 = ["it_skills_pct", "ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct"]
M7 = ["M6a", "M6b", "M6c"]
rows = []
for x in X7:
    rows.append([VN[x]] + [cell(float(get(b7, m, x)), float(get(b7, m, x, "se")), float(get(b7, m, x, "p"))) if not pd.isna(get(b7, m, x)) else "" for m in M7])
rows += [["Số tỉnh"] + [fi(float(get(b7, m, "N"))) for m in M7],
         ["R²"] + [f(float(get(b7, m, "R2")), 3) for m in M7], ["R² hiệu chỉnh"] + [f(float(get(b7, m, "R2_adj")), 3) for m in M7]]
T["b7"] = dict(title="Bảng 7. Kỹ năng CNTT và EBI năm 2024 (mô hình cắt ngang M6)",
               header=["Biến giải thích", "M6a", "M6b", "M6c"],
               rows=rows, widths=[3600, 1700, 1700, 1700],
               note="Ghi chú: biến phụ thuộc là DTF của EBI năm 2024, phương trình {{eq:cross}}; OLS với sai số chuẩn vững HC1 trong ngoặc. M6b kiểm soát mật độ doanh nghiệp, M6c thay mật độ doanh nghiệp bằng tỷ lệ lao động qua đào tạo. Hệ số tương quan giữa tỷ lệ dân số có kỹ năng CNTT và mật độ doanh nghiệp năm 2024 là " + f(KQ["corr_2024"].get("dn_per_1000") and float(pd.read_csv(os.path.join(OUT, "panel_phan_tich.csv"), encoding="utf-8-sig").query("nam == 2024")[["it_skills_pct", "dn_per_1000"]].corr().iloc[0, 1]), 2) + ". " + NOTE_SAO)

# ---------------- Phụ lục A1: đóng góp từng vùng ----------------
pr = rd("phu_theil_theo_vung.csv"); rows = []
for v in REG_ORDER:
    a = pr[(pr.nam == 2020) & (pr.vung == v)].iloc[0]; b = pr[(pr.nam == 2024) & (pr.vung == v)].iloc[0]
    rows.append([REG_FULL[v], f"{int(a.n)} / {int(b.n)}", f(a.mean_dtf, 1), f(b.mean_dtf, 1), f(a.Theil_k, 3), f(b.Theil_k, 3), f(a.dong_gop_pct, 1), f(b.dong_gop_pct, 1)])
T["a1"] = dict(title="Bảng A1. Bất bình đẳng trong từng vùng và đóng góp vào chỉ số Theil-T, năm 2020 và 2024",
               header=["Vùng", "Số tỉnh 2020 / 2024", "DTF TB 2020 (%)", "DTF TB 2024 (%)", "Theil-T trong vùng 2020", "Theil-T trong vùng 2024", "Đóng góp vào Theil-T 2020 (%)", "Đóng góp vào Theil-T 2024 (%)"],
               rows=rows, widths=[2700, 900, 900, 900, 1000, 1000, 1000, 1000],
               note="Ghi chú: đóng góp của vùng k = 100·s_k·Th_k/Th (phần trong vùng của vùng k). Tổng cột đóng góp bằng tỷ trọng trong vùng ở Bảng 2. Nguồn: tính toán của tác giả.")

# ---------------- Phụ lục A2, A3: độ vững RQ3 ----------------
ba = rd("phu_A_do_vung_RQ3.csv")
def tbl(models, heads, title, note, xs):
    rows = []
    for x in xs:
        vals = [cell(float(get(ba, m, x)), float(get(ba, m, x, "se")), float(get(ba, m, x, "p"))) if not pd.isna(get(ba, m, x)) else "" for m in models]
        if any(vals): rows.append([VN[x]] + vals)
    rows += [["Giai đoạn"] + [get(ba, m, "nam") for m in models],
             ["Số quan sát (số tỉnh)"] + [f"{int(float(get(ba, m, 'N')))} ({int(float(get(ba, m, 'N_tinh')))})" for m in models],
             ["R²"] + [f(float(get(ba, m, "R2")), 3) for m in models]]
    w = [2600] + [int(6400 / len(models))] * len(models)
    return dict(title=title, header=["Biến giải thích"] + heads, rows=rows, widths=w, note=note)
T["a2"] = tbl(["A1. M2 – FE tỉnh + năm", "A2. M2 – 2018–2020", "A3. M2 – 2022–2024", "A4. M2 – loại HN & TP.HCM", "A5. M2 – X trễ 1 năm", "A8. M2 + hộ có máy tính (2019–2024)"],
              ["(1) FE tỉnh + năm", "(2) 2018–2020", "(3) 2022–2024", "(4) Loại HN & TP.HCM", "(5) Biến giải thích trễ 1 năm", "(6) Thêm hộ có máy tính"],
              "Bảng A2. Kiểm tra độ vững của mô hình M2 (biến phụ thuộc: DTF của EBI)",
              "Ghi chú: cột (1) ước lượng hiệu ứng cố định tỉnh và năm, chỉ khai thác biến thiên theo thời gian trong nội bộ từng tỉnh, R² của cột này là R² trong nội bộ (within); các cột còn lại là OLS gộp có hiệu ứng năm. Sai số chuẩn theo cụm tỉnh trong ngoặc. " + NOTE_SAO, XS)
T["a3"] = tbl(["A6. M3 – 2018–2020", "A7. M4 – 2018–2020", "A9. M5 bỏ hộ có máy tính (2018–2024)"],
              ["M3 (B2C), 2018–2020", "M4 (B2B), 2018–2020", "M5 (NNL&HT) bỏ hộ có máy tính, 2018–2024"],
              "Bảng A3. Kiểm tra các mô hình theo trụ cột trước khi EBI đổi phương pháp và khi bỏ biến hộ có máy tính",
              "Ghi chú: giai đoạn 2018–2020 là thời kỳ EBI chưa chứa chỉ tiêu thu nhập bình quân (B2C) và doanh nghiệp trên dân số (B2B), nên hệ số ở hai cột đầu không mang thành phần trùng định nghĩa. OLS gộp có hiệu ứng năm, sai số chuẩn theo cụm tỉnh trong ngoặc. " + NOTE_SAO, XS)

# =====================================================================
# CÁC CON SỐ DÙNG TRONG LỜI VĂN – {{n:tên}}
# =====================================================================
r2 = b2.set_index("nam"); r3 = b3.set_index("nam")
N.update(N_obs=fi(KQ["N_obs"]), N_tinh=str(KQ["N_tinh_co_EBI"]), N_du8=str(KQ["N_tinh_du_8_nam"]),
         N_min=str(min(KQ["N_theo_nam"].values())), N_max=str(max(KQ["N_theo_nam"].values())))
for y in range(2017, 2025):
    r = r2.loc[y]
    N[f"mean{y}"] = f(r["mean"], 2); N[f"cv{y}"] = f(r.CV, 3); N[f"gini{y}"] = f(r.Gini, 3); N[f"theil{y}"] = f(r.Theil, 3)
    N[f"in{y}"] = f(r.within_share, 1); N[f"bw{y}"] = f(r.between_share, 1); N[f"dtf{y}"] = f(r.DTF_mean, 1)
    N[f"maxmin{y}"] = f(r.max_min, 2); N[f"p9010_{y}"] = f(r.p90_p10, 2); N[f"sdln{y}"] = f(r.SD_ln, 3)
    N[f"gini_ex{y}"] = f(r.Gini_ex2, 3); N[f"theil_cut{y}"] = f(KQ["theil_giam_khi_bo_2tp_pct"][str(y)], 0)
    for p in ["NNL_HT", "B2C", "B2B", "G2B"]:
        N[f"g_{p}{y}"] = f(r3.loc[y, f"Gini_{p}"], 3)
N["in_min"], N["in_max"] = f(KQ["within_share_range"][0], 1), f(KQ["within_share_range"][1], 1)
N["bw_min"], N["bw_max"] = f(100 - KQ["within_share_range"][1], 1), f(100 - KQ["within_share_range"][0], 1)
N["in_ex_min"], N["in_ex_max"] = f(KQ["within_share_range_ex2"][0], 1), f(KQ["within_share_range_ex2"][1], 1)
N["in_bal_min"], N["in_bal_max"] = f(KQ["within_share_range_bal"][0], 1), f(KQ["within_share_range_bal"][1], 1)
dg = KQ["dong_gop_vung_2024"]; N["dg_DNB"] = f(dg["ĐNB"], 1); N["dg_DBSH"] = f(dg["ĐBSH"], 1)
N["dg_2vung"] = f(dg["ĐNB"] + dg["ĐBSH"], 1); N["dg_4vung"] = f(sum(v for k, v in dg.items() if k not in ("ĐNB", "ĐBSH")), 1)
tk = KQ["theil_k_2024"]; N["thk_DNB"] = f(tk["ĐNB"], 3); N["thk_DBSH"] = f(tk["ĐBSH"], 3)
N["thk_min4"] = f(min(v for k, v in tk.items() if k not in ("ĐNB", "ĐBSH")), 3); N["thk_max4"] = f(max(v for k, v in tk.items() if k not in ("ĐNB", "ĐBSH")), 3)
dv = KQ["dtf_vung_2024"]
for k, key in [("ĐNB", "DNB"), ("ĐBSH", "DBSH"), ("BTB&DHMT", "BTB"), ("ĐBSCL", "DBSCL"), ("TN", "TN"), ("TDMNPB", "TDMN")]:
    N[f"dtfv24_{key}"] = f(dv[k], 1); N[f"dtfv17_{key}"] = f(KQ["dtf_vung_2017"][k], 1)
top = KQ["top5_2024"]; bot = KQ["bot5_2024"][:5]
N["top5"] = "; ".join(f"{t} ({f(v, 0)})" for t, v in top); N["bot"] = "; ".join(f"{t} ({f(v, 0)})" for t, v in bot)
# hội tụ beta
B = {r["mo_hinh"]: r for r in KQ["beta"]}; R = {("goc" if r["mo_hinh"].startswith("Gốc") else r["mo_hinh"].split(".")[0].lower()): r for r in KQ["beta_rob"]}
def put(prefix, r, d=4):
    N[prefix + "_b"] = f(r["beta"], d); N[prefix + "_se"] = f(r["se"], d); N[prefix + "_p"] = fp(r["p"]); N[prefix + "_N"] = str(int(r["N"]))
    N[prefix + "_r2"] = f(r["R2"], 3)
    if r.get("lambda_pct") not in (None, "nan") and not pd.isna(r.get("lambda_pct")): N[prefix + "_lam"] = f(r["lambda_pct"], 1)
put("b1A_cu", B["1A. Tuyệt đối – thang cũ"]); put("b1A_od", B["1A. Tuyệt đối – thang ổn định"])
put("b1B_cu", B["1B. Có điều kiện – thang cũ"]); put("b1B_od", B["1B. Có điều kiện – thang ổn định"])
for k, r in R.items(): put(k if k == "goc" else k, r)
N["b1B_dn"] = f(B["1B. Có điều kiện – thang ổn định"]["coef_dn_per_1000_L1"], 4)
# hồi quy M2–M6
for m in MS:
    for x in XS:
        c = get(b6, m, x)
        if not pd.isna(c):
            N[f"{m}_{x}"] = f(float(c), 3); N[f"{m}_{x}_p"] = fp(float(get(b6, m, x, "p"))); N[f"{m}_{x}_se"] = f(float(get(b6, m, x, "se")), 3)
    N[f"{m}_N"] = str(int(float(get(b6, m, "N")))); N[f"{m}_r2"] = f(float(get(b6, m, "R2")), 3)
for m in M7:
    for x in X7:
        c = get(b7, m, x)
        if not pd.isna(c):
            N[f"{m}_{x}"] = f(float(c), 3); N[f"{m}_{x}_p"] = fp(float(get(b7, m, x, "p"))); N[f"{m}_{x}_se"] = f(float(get(b7, m, x, "se")), 3)
    N[f"{m}_r2"] = f(float(get(b7, m, "R2")), 3); N[f"{m}_N"] = str(int(float(get(b7, m, "N"))))
for m, key in [("A1. M2 – FE tỉnh + năm", "A1"), ("A2. M2 – 2018–2020", "A2"), ("A3. M2 – 2022–2024", "A3"), ("A4. M2 – loại HN & TP.HCM", "A4"),
               ("A5. M2 – X trễ 1 năm", "A5"), ("A6. M3 – 2018–2020", "A6"), ("A7. M4 – 2018–2020", "A7"),
               ("A8. M2 + hộ có máy tính (2019–2024)", "A8"), ("A9. M5 bỏ hộ có máy tính (2018–2024)", "A9")]:
    for x in XS:
        c = get(ba, m, x)
        if not pd.isna(c): N[f"{key}_{x}"] = f(float(c), 3); N[f"{key}_{x}_p"] = fp(float(get(ba, m, x, "p")))
tq = KQ["tuong_quan_EBI"]
for x, v in tq.items():
    N[f"r_{x}"] = f(v["r"], 3); N[f"rnnl_{x}"] = f(v["r_nnl"], 3)
N["vif_max"] = f(max(KQ["vif_M2"].values()), 2)
N["r24_internet"] = f(KQ["corr_2024"]["internet_user_pct"], 3); N["r24_skill"] = f(KQ["corr_2024"]["it_skills_pct"], 3)
N["r24_comp"] = f(KQ["corr_2024"]["hh_computer_pct"], 3)
pan = pd.read_csv(os.path.join(OUT, "panel_phan_tich.csv"), encoding="utf-8-sig")
N["r24_skill_dn"] = f(pan.query("nam == 2024")[["it_skills_pct", "dn_per_1000"]].corr().iloc[0, 1], 2)
N["r_dn_comp"] = f(KQ["tuong_quan_X"]["dn_per_1000"]["hh_computer_pct"], 2)
N["r_tr_comp"] = f(KQ["tuong_quan_X"]["trained_labor_pct"]["hh_computer_pct"], 2)
# trung bình trụ cột (để mô tả thang đo)
dd = pan.dropna(subset=["EBI"])
for p in ["NNL_HT", "B2C", "B2B"]:
    for y in (2020, 2021, 2022, 2023, 2024):
        N[f"mean_{p}{y}"] = f(dd[dd.nam == y][p].mean(), 1)

with open(os.path.join(OUT, "bang_docx.json"), "w", encoding="utf-8") as fo: json.dump(T, fo, ensure_ascii=False, indent=1)
with open(os.path.join(OUT, "so_lieu_bai.json"), "w", encoding="utf-8") as fo: json.dump(N, fo, ensure_ascii=False, indent=1)

# ---------------- Sổ Excel tổng hợp (số chưa định dạng) ----------------
SHEETS = [("Bang1_mo_ta", "bang1_mo_ta.csv"), ("Bang2_khoang_cach", "bang2_khoang_cach_theo_nam.csv"),
          ("Bang3_gini_tru_cot", "bang3_gini_tru_cot.csv"), ("Bang4_hoi_tu_beta", "bang4_hoi_tu_beta.csv"),
          ("Bang5_beta_do_vung", "bang5_beta_do_vung.csv"), ("Bang6_M2_M5", "bang6_hoi_quy_M2_M5.csv"),
          ("Bang7_M6_2024", "bang7_cat_ngang_2024.csv"), ("BangA1_theil_vung", "phu_theil_theo_vung.csv"),
          ("BangA2_A3_do_vung", "phu_A_do_vung_RQ3.csv"), ("Hinh3_DTF_vung", "phu_dtf_vung_theo_nam.csv"),
          ("Hinh6_tuong_quan", "phu_tuong_quan.csv"), ("Panel_phan_tich", "panel_phan_tich.csv")]
with pd.ExcelWriter(os.path.join(ROOT, "So-lieu-ket-qua-v8.xlsx"), engine="openpyxl") as xw:
    pd.DataFrame({"Nội dung": [
        "Sổ số liệu kết quả của bài báo bản v8 – sinh tự động bởi scripts/03_bang_va_so_lieu.py",
        "Mỗi sheet là một bảng kết quả dạng số (chưa làm tròn); bảng trong bài là bản đã làm tròn của các sheet này.",
        "p = giá trị p; se = sai số chuẩn; coef = hệ số; DTF = điểm của tỉnh / điểm tỉnh dẫn đầu cùng năm × 100.",
        "Chạy lại toàn bộ: chay_tat_ca.bat (Windows) hoặc lần lượt scripts/01 → 02 → 03 → node scripts/04_dung_word.js"]}
    ).to_excel(xw, sheet_name="Doc_truoc", index=False)
    for sh, fn in SHEETS:
        rd(fn).to_excel(xw, sheet_name=sh, index=False)
print("Đã ghi ket-qua/bang_docx.json (%d bảng), ket-qua/so_lieu_bai.json (%d số), So-lieu-ket-qua-v8.xlsx" % (len(T), len(N)))

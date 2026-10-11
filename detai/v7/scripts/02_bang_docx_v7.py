# -*- coding: utf-8 -*-
"""Định dạng các bảng kết quả v7 (ket-qua-v7/*.csv) thành JSON (ket-qua-v7/bang_docx_v7.json) để dựng file Word bài báo v7.
Bảng chính: b1–b7; bảng phụ lục: a1–a5. Dấu phẩy thập phân kiểu Việt Nam."""
import json, os
import numpy as np, pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "ket-qua-v7")
KQ = json.load(open(os.path.join(OUT, "ket_qua_v7.json"), encoding="utf-8"))
REG_ORDER = ["ĐNB", "ĐBSH", "BTB&DHMT", "ĐBSCL", "TN", "TDMNPB"]
REG_FULL = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
VN = {"ln_grdp": "ln(GRDP/người)", "dn_per_1000": "Mật độ doanh nghiệp (DN/1.000 dân)",
      "urban_share_pct": "Tỷ lệ dân số thành thị (%)", "trained_labor_pct": "Tỷ lệ LĐ qua đào tạo (%)",
      "ln_freight": "ln(Hàng hóa vận chuyển đường bộ/người)", "modern_retail_100k": "Siêu thị + TTTM/100.000 dân",
      "ln_fdi_pc": "ln(1 + FDI lũy kế/người)", "hh_computer_pct": "Tỷ lệ hộ có máy tính (%)",
      "internet_user_pct": "Tỷ lệ người sử dụng Internet (%)", "it_skills_pct": "Tỷ lệ người có kỹ năng CNTT (%)"}
X = ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct", "ln_freight", "modern_retail_100k", "ln_fdi_pc", "hh_computer_pct"]

def f(x, d=3):
    if x is None or (isinstance(x, float) and np.isnan(x)) or (isinstance(x, str) and x == ""): return ""
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",").replace("-", "−")
def fi(x): return "" if pd.isna(x) else f"{int(round(x)):,}".replace(",", ".")
def stars(p): return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""
def cell(coef, se, p, d=3): return f"{f(coef, d)}{stars(p)}\n({f(se, d)})"

T = {}
# ---------------- Bảng 1: mô tả biến ----------------
b1 = pd.read_csv(os.path.join(OUT, "bang1_mo_ta.csv"))
rows = []
for _, r in b1.iterrows():
    d = 2 if r["mean"] < 100 else 1 if r["mean"] < 1000 else 0
    rows.append([r["ten"], r["don_vi"], r["nguon"], str(int(r["N"])), f(r["mean"], d), f(r["sd"], d), f(r["min"], d), f(r["max"], d)])
T["b1"] = dict(title="Bảng 1. Biến số và thống kê mô tả (tỉnh – năm, 2017–2024)",
               header=["Biến", "Đơn vị", "Nguồn", "N", "Trung bình", "Độ lệch chuẩn", "Nhỏ nhất", "Lớn nhất"],
               rows=rows, widths=[2900, 1000, 800, 600, 1000, 1000, 900, 900],
               note="Ghi chú: N là số quan sát tỉnh – năm có số liệu; tối đa 504 (63 tỉnh × 8 năm). EBI và các trụ cột khuyết ở 5–9 tỉnh mỗi năm (Bảng 2). GRDP bình quân đầu người được Niên giám công bố theo tỉnh từ năm 2018, tỷ lệ hộ có máy tính từ năm 2019. Nguồn: tính toán của tác giả từ Báo cáo Chỉ số Thương mại điện tử Việt Nam (VECOM, 2017–2025) và Niên giám Thống kê (NGTK) 2017–2024.")

# ---------------- Bảng 2: độ phủ và khả năng so sánh ----------------
b2 = pd.read_csv(os.path.join(OUT, "bang2_do_phu.csv"))
rows = []
for _, r in b2.iterrows():
    rows.append([str(int(r.nam)), str(int(r.N)), f(r["mean"], 1), f(r["median"], 1), f(r["min"], 1), f(r["max"], 1),
                 f(r.ratio_max_median, 2), f(r.ratio_p90_p10, 2), f(r.spearman_prev, 3) if not pd.isna(r.spearman_prev) else "—"])
sp = KQ["spearman"]
T["b2"] = dict(title="Bảng 2. Độ phủ của EBI và các chỉ báo về khả năng so sánh theo năm",
               header=["Năm", "Số tỉnh có EBI", "Trung bình", "Trung vị", "Nhỏ nhất", "Lớn nhất", "Lớn nhất / trung vị", "P90 / P10", "Tương quan hạng Spearman với năm trước"],
               rows=rows, widths=[600, 800, 900, 900, 800, 800, 1000, 900, 1400],
               note=("Ghi chú: EBI theo năm công bố báo cáo. Các tỉnh không được xếp hạng: Cao Bằng, Hà Giang, Lai Châu, Điện Biên, Sơn La (cả 8 năm); Bạc Liêu, Hậu Giang, Sóc Trăng (2017–2022, riêng Hậu Giang đến 2020); Vĩnh Long (2017); Cà Mau (2018–2019). "
                     f"Tương quan hạng Spearman giữa hai đầu giai đoạn: 2017–2020 là {f(sp['2017-2020']['rho'], 3)}; 2021–2024 là {f(sp['2021-2024']['rho'], 3)}; qua điểm gãy 2020–2021 là {f(sp['2020-2021']['rho'], 3)}; toàn kỳ 2017–2024 là {f(sp['2017-2024']['rho'], 3)}. Nguồn: tính toán của tác giả từ VECOM (2017–2025)."))

# ---------------- Bảng 3: EBI tương đối theo vùng ----------------
b3 = pd.read_csv(os.path.join(OUT, "bang3_vung.csv")).set_index("vung6").loc[REG_ORDER]
rows = []
for r, x in b3.iterrows():
    rows.append([REG_FULL[r], str(int(x["n_tinh"])), f(x["2017–2020"], 1), f(x["2021–2024"], 1), f(x["EBI_rel"], 1), f(x["NNL_HT_rel"], 1), f(x["B2C_rel"], 1), f(x["B2B_rel"], 1)])
hi, lo = b3.loc["ĐNB"], b3.loc[["TDMNPB", "TN", "ĐBSCL"]].min()
rows.append(["Tỷ số vùng cao nhất / vùng thấp nhất", "", f(hi["2017–2020"] / lo["2017–2020"], 2), f(hi["2021–2024"] / lo["2021–2024"], 2),
             f(hi["EBI_rel"] / lo["EBI_rel"], 2), f(hi["NNL_HT_rel"] / lo["NNL_HT_rel"], 2), f(hi["B2C_rel"] / lo["B2C_rel"], 2), f(hi["B2B_rel"] / lo["B2B_rel"], 2)])
T["b3"] = dict(title="Bảng 3. EBI tương đối theo sáu vùng kinh tế – xã hội (trung bình các tỉnh có số liệu trong năm = 100)",
               header=["Vùng", "Số tỉnh có số liệu", "EBI 2017–2020", "EBI 2021–2024", "EBI 2024", "NNL&HT 2024", "B2C 2024", "B2B 2024"],
               rows=rows, widths=[2900, 800, 1100, 1100, 900, 1100, 900, 900],
               note="Ghi chú: EBI tương đối tính theo phương trình {{eq:ebi_rel}}; trung bình vùng là trung bình giản đơn của các tỉnh có số liệu trong vùng. Hai cột giai đoạn là trung bình của các năm trong giai đoạn. NNL&HT: trụ cột nguồn nhân lực và hạ tầng công nghệ thông tin. Nguồn: tính toán của tác giả.")

# ---------------- Bảng 4: Gini, Theil và phân rã ----------------
b4 = pd.read_csv(os.path.join(OUT, "bang4_gini_theil.csv"))
rows = []
for _, r in b4.iterrows():
    rows.append([str(int(r.nam)), str(int(r.N)), f(r.Gini), f(r.Theil), f(r.Theil_between), f(r.Theil_within), f(r.Theil_between_share * 100, 1), f(r.SD_ln),
                 f(r.Gini_ex2), f(r.Theil_ex2), f(r.between_share_ex2 * 100, 1)])
T["b4"] = dict(title="Bảng 4. Mức độ bất bình đẳng EBI giữa các tỉnh và phân rã Theil-T theo sáu vùng, 2017–2024",
               header=["Năm", "N", "Gini", "Theil-T (Th)", "Giữa các vùng (Th_B)", "Trong vùng (Th_W)", "Tỷ trọng giữa vùng (%)", "SD của ln(EBI)",
                       "Gini, loại HN & TP.HCM", "Th, loại HN & TP.HCM", "Tỷ trọng giữa vùng, loại HN & TP.HCM (%)"],
               rows=rows, widths=[550, 450, 700, 750, 850, 800, 900, 800, 900, 900, 1100],
               note="Ghi chú: mỗi tỉnh là một quan sát có trọng số bằng nhau. Gini theo phương trình {{eq:gini}}, Theil-T theo {{eq:theil}}, phân rã theo {{eq:theil_dec}}. SD của ln(EBI) theo {{eq:sdln}}. Các chỉ số bất biến với phép nhân thang đo nên so sánh được trong từng giai đoạn; bước nhảy năm 2021 phản ánh việc thay đổi phương pháp tính EBI (Mục 4.2), không được đọc như thay đổi thực của khoảng cách. Ba cột cuối tính lại sau khi loại Hà Nội và TP. Hồ Chí Minh. Nguồn: tính toán của tác giả.")

# ---------------- Bảng 5: đóng góp từng vùng ----------------
b5 = pd.read_csv(os.path.join(OUT, "bang5_theil_vung.csv"))
rows = []
for r in REG_ORDER:
    a = b5[(b5.nam == 2020) & (b5.vung == r)].iloc[0]; b = b5[(b5.nam == 2024) & (b5.vung == r)].iloc[0]
    rows.append([REG_FULL[r], str(int(a.n)), f(a.mean_rel, 1), f(a.Theil_k), f(a.ty_trong_trong_T, 1), str(int(b.n)), f(b.mean_rel, 1), f(b.Theil_k), f(b.ty_trong_trong_T, 1)])
tw20 = b5[b5.nam == 2020].ty_trong_trong_T.sum(); tw24 = b5[b5.nam == 2024].ty_trong_trong_T.sum()
rows.append(["Tổng phần trong vùng (Th_W / Th)", "", "", "", f(tw20, 1), "", "", "", f(tw24, 1)])
T["b5"] = dict(title="Bảng 5. Bất bình đẳng trong nội bộ từng vùng và đóng góp vào tổng Theil-T, năm 2020 và 2024",
               header=["Vùng", "N 2020", "EBI tương đối 2020", "Th_k 2020", "Đóng góp vào Th 2020 (%)", "N 2024", "EBI tương đối 2024", "Th_k 2024", "Đóng góp vào Th 2024 (%)"],
               rows=rows, widths=[2600, 550, 900, 750, 950, 550, 900, 750, 950],
               note="Ghi chú: Th_k là chỉ số Theil-T tính riêng trong vùng k; đóng góp của vùng k vào tổng Theil-T bằng s_k·Th_k / Th × 100 với s_k là tỷ trọng tổng điểm EBI của vùng (phương trình {{eq:theil_dec}}). Tổng đóng góp của sáu vùng bằng tỷ trọng phần trong vùng (Bảng 4). Nguồn: tính toán của tác giả.")

# ---------------- Bảng 6: Gini từng trụ cột ----------------
b6 = pd.read_csv(os.path.join(OUT, "bang6_gini_tru_cot.csv"))
rows = []
for _, r in b6.iterrows():
    rows.append([str(int(r.nam)), f(r.Gini_EBI), f(r.Gini_NNL_HT), f(r.Gini_B2C), f(r.Gini_B2B), f(r.get("Gini_G2B", np.nan)) if not pd.isna(r.get("Gini_G2B", np.nan)) else "—"])
T["b6"] = dict(title="Bảng 6. Hệ số Gini của EBI và của từng trụ cột, 2017–2024",
               header=["Năm", "EBI", "NNL&HT", "B2C", "B2B", "G2B"],
               rows=rows, widths=[900, 1400, 1400, 1400, 1400, 1400],
               note="Ghi chú: Gini tính theo phương trình {{eq:gini}} trên điểm của từng trụ cột. Trọng số của các trụ cột trong EBI (phục hồi từ dữ liệu bằng hồi quy không hệ số chặn, R² ≥ 0,9997): 0,3/0,3/0,3/0,1 cho NNL&HT/B2C/B2B/G2B năm 2017; 0,2/0,35/0,35/0,1 giai đoạn 2018–2020; 0,2/0,4/0,4 cho NNL&HT/B2C/B2B từ năm 2021 (G2B bị loại). Đóng góp của từng trụ cột vào Gini của EBI trình bày ở Bảng A2. Nguồn: tính toán của tác giả.")

# ---------------- Bảng 7: hồi quy FE ----------------
b7 = pd.read_csv(os.path.join(OUT, "bang7_fe.csv"))
def g7(m, t):
    q = b7[(b7.model == m) & (b7.term == t)]; return None if q.empty else q.iloc[0]
def bang_hoi_quy(models, key, title, header, note, lagged=None):
    rows = []
    for v in X:
        row = [VN[v]]
        for m in models:
            t = v + "_L1" if (lagged and m in lagged) else v
            r = g7(m, t); row.append("" if r is None else cell(r.coef, r.se, r.p, 4 if "hạng" in m else 3))
        rows.append(row)
    for stat, name in [("N", "Số quan sát"), ("N_tinh", "Số tỉnh"), ("R2_within", "R² (sau khi khử hiệu ứng)"), ("R2_overall", "R²")]:
        row = [name]; any_ = False
        for m in models:
            r = g7(m, stat)
            if r is None or (stat == "R2_overall" and not m.startswith("OLS")) or (stat == "R2_within" and m.startswith("OLS")): row.append("")
            else: row.append(fi(r.coef) if stat in ("N", "N_tinh") else f(r.coef, 3)); any_ = True
        if any_: rows.append(row)
    rows.append(["Hiệu ứng cố định tỉnh"] + ["Không" if m.startswith("OLS") else "Có" for m in models])
    rows.append(["Hiệu ứng cố định năm"] + ["Có"] * len(models))
    T[key] = dict(title=title, header=header, rows=rows, widths=[2600] + [round(6470 / len(models))] * len(models), note=note)

H = KQ["hausman"]
bang_hoi_quy(["FE1: 2019–2024", "FE2: 2018–2024 (không hộ máy tính)", "FE3: 2021–2024", "FE5: 2019–2024, Y = hạng phần trăm", "OLS gộp 2019–2024 (so sánh)"], "b7",
             "Bảng 7. Các đặc điểm kinh tế – xã hội có liên hệ với EBI: hồi quy dữ liệu bảng hiệu ứng cố định tỉnh và năm",
             ["Biến giải thích", "(1) FE, 2019–2024", "(2) FE, 2018–2024", "(3) FE, 2021–2024", "(4) FE, Y = hạng phần trăm, 2019–2024", "(5) OLS gộp, 2019–2024 (so sánh)"],
             f"Ghi chú: biến phụ thuộc là ln(EBI) ở các cột (1)–(3) và (5), là hạng phần trăm của tỉnh trong năm (0–100) ở cột (4). Mô hình theo phương trình {{{{eq:fe}}}}; sai số chuẩn cụm theo tỉnh trong ngoặc; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%. Cột (2) không có tỷ lệ hộ có máy tính (Niên giám công bố từ 2019) nên mẫu dài hơn. Cột (5) là OLS gộp với hiệu ứng năm, chỉ để so sánh: hệ số phản ánh chủ yếu biến thiên giữa các tỉnh. Kiểm định Hausman giữa hiệu ứng cố định và hiệu ứng ngẫu nhiên trên mẫu (1): χ²({H['df']}) = {f(H['stat'], 1)}, p < 0,001. R² (sau khi khử hiệu ứng) là R² của hồi quy trên số liệu đã trừ trung bình tỉnh và trung bình năm. Nguồn: tính toán của tác giả.")

# ---------------- Phụ lục A1: mẫu cân bằng và MLD ----------------
rows = []
for _, r in b4.iterrows():
    rows.append([str(int(r.nam)), str(int(r.N_bal)), f(r.Gini_bal), f(r.Theil_bal), f(r.between_share_bal * 100, 1), f(r.MLD), f(r.MLD_between_share * 100, 1), f(r.CV)])
T["a1"] = dict(title="Bảng A1. Kiểm tra độ vững của Bảng 4: mẫu cân bằng 53 tỉnh và chỉ số MLD",
               header=["Năm", "N (cân bằng)", "Gini (cân bằng)", "Theil-T (cân bằng)", "Tỷ trọng giữa vùng, Theil (cân bằng, %)", "MLD (mẫu đầy đủ)", "Tỷ trọng giữa vùng, MLD (%)", "Hệ số biến thiên (mẫu đầy đủ)"],
               rows=rows, widths=[600, 800, 900, 1000, 1300, 1000, 1200, 1100],
               note="Ghi chú: mẫu cân bằng gồm 53 tỉnh có EBI đủ cả 8 năm (loại Bạc Liêu, Cà Mau, Hậu Giang, Sóc Trăng, Vĩnh Long và 5 tỉnh không được xếp hạng). MLD là độ lệch logarit trung bình, GE(0), phân rã theo phương trình {{eq:mld_dec}} với trọng số là tỷ trọng số tỉnh; MLD nhạy hơn với đuôi dưới của phân phối. Nguồn: tính toán của tác giả.")

# ---------------- Phụ lục A2: phân rã Gini theo nguồn ----------------
rows = []
for _, r in b6.iterrows():
    g2b = r.get("share_G2B", np.nan)
    rows.append([str(int(r.nam)), f(r.Gini_EBI_taitao), f(r.share_NNL_HT * 100, 1), f(r.share_B2C * 100, 1), f(r.share_B2B * 100, 1), f(g2b * 100, 1) if not pd.isna(g2b) else "—"])
T["a2"] = dict(title="Bảng A2. Đóng góp của từng trụ cột vào Gini của EBI (phân rã Gini theo nguồn), 2017–2024",
               header=["Năm", "Gini EBI (tái tạo từ trụ cột)", "Đóng góp NNL&HT (%)", "Đóng góp B2C (%)", "Đóng góp B2B (%)", "Đóng góp G2B (%)"],
               rows=rows, widths=[800, 1500, 1400, 1300, 1300, 1300],
               note="Ghi chú: phân rã theo phương trình {{eq:gini_src}}: G = Σ S_k G_k R_k, trong đó S_k là tỷ trọng của trụ cột k trong EBI, G_k là Gini của trụ cột và R_k là tương quan Gini giữa trụ cột với thứ hạng EBI; đóng góp (%) = S_k G_k R_k / G × 100. Gini EBI ở đây tính trên EBI tái tạo từ các trụ cột theo trọng số nêu ở Bảng 6, nên có thể lệch nhẹ so với Bảng 4 (năm 2023: 0,258 so với 0,262) do VECOM làm tròn điểm trụ cột. Nguồn: tính toán của tác giả.")

# ---------------- Phụ lục A3: hội tụ beta tuyệt đối ----------------
pb = pd.read_csv(os.path.join(OUT, "phu_beta_tuyet_doi.csv"))
rows = []
for _, r in pb.iterrows():
    rows.append([r.period, cell(r.beta, r.se, r.p, 4), f(r.lam * 100, 1) if r.lam > 0 and r.p < 0.1 else "—", f(r.half_life, 1) if r.lam > 0 and r.p < 0.1 else "—", str(int(r.N)), f(r.R2, 3)])
T["a3"] = dict(title="Bảng A3. Hồi quy hội tụ beta tuyệt đối theo giai đoạn (phân tích bổ sung)",
               header=["Giai đoạn / mẫu", "β (SE)", "Tốc độ hội tụ λ (%/năm)", "Thời gian bán hội tụ (năm)", "N", "R²"],
               rows=rows, widths=[3000, 1400, 1300, 1300, 600, 800],
               note="Ghi chú: mô hình (1/T)·ln(EBI_iT/EBI_i0) = α + β·ln(EBI_i0) + ε_i, ước lượng bằng OLS với sai số chuẩn vững HC1; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%; λ = −ln(1 + βT)/T và thời gian bán hội tụ = ln2/λ chỉ báo cáo khi β âm có ý nghĩa ở mức 10%. Hai dòng cuối loại Hà Nội và TP. Hồ Chí Minh. Không hồi quy xuyên điểm gãy 2020–2021. Nguồn: tính toán của tác giả.")

# ---------------- Phụ lục A4: FE bổ sung ----------------
bang_hoi_quy(["FE4: 2019–2024, bỏ mật độ DN", "FE6: 2019–2024, loại năm 2021", "FE7: X trễ 1 năm (2020–2024)", "OLS gộp 2018–2020 (trước đổi phương pháp)"], "a4",
             "Bảng A4. Kiểm tra độ vững của Bảng 7",
             ["Biến giải thích", "(6) FE, bỏ mật độ DN, 2019–2024", "(7) FE, loại năm 2021", "(8) FE, biến giải thích trễ 1 năm, 2020–2024", "(9) OLS gộp 2018–2020, trước đổi phương pháp"],
             "Ghi chú: biến phụ thuộc ln(EBI); sai số chuẩn cụm theo tỉnh trong ngoặc; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%. Cột (8) dùng toàn bộ biến giải thích trễ một năm. Cột (9) là OLS gộp với hiệu ứng năm trên giai đoạn EBI chưa chứa chỉ tiêu doanh nghiệp/dân số (Mục 4.2); không có tỷ lệ hộ có máy tính vì chỉ có từ 2019. Nguồn: tính toán của tác giả.",
             lagged=["FE7: X trễ 1 năm (2020–2024)"])

# ---------------- Phụ lục A5: cắt ngang 2024 ----------------
pc = pd.read_csv(os.path.join(OUT, "phu_2024_cat_ngang.csv"))
mods = ["C1", "C2", "C3 (bỏ mật độ DN)"]
rows = []
for v in ["internet_user_pct", "it_skills_pct", "ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct"]:
    row = [VN[v]]
    for m in mods:
        q = pc[(pc.model == m) & (pc.term == v)]; row.append("" if q.empty else cell(q.iloc[0].coef, q.iloc[0].se, q.iloc[0].p))
    rows.append(row)
for stat, name in [("N", "Số quan sát"), ("R2_adj", "R² hiệu chỉnh")]:
    rows.append([name] + [fi(pc[(pc.model == m) & (pc.term == stat)].iloc[0].coef) if stat == "N" else f(pc[(pc.model == m) & (pc.term == stat)].iloc[0].coef, 3) for m in mods])
T["a5"] = dict(title="Bảng A5. Cắt ngang năm 2024 với các chỉ tiêu ở tầng kỹ năng và sử dụng (OLS, biến phụ thuộc ln(EBI 2024))",
               header=["Biến giải thích", "(1) Chỉ biến số", "(2) Biến số + kinh tế", "(3) Biến số + kinh tế, bỏ mật độ DN"],
               rows=rows, widths=[3200, 1900, 1900, 1900],
               note="Ghi chú: 58 tỉnh có EBI 2024; sai số chuẩn vững HC1 trong ngoặc; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%. Tỷ lệ người sử dụng Internet và tỷ lệ người có kỹ năng công nghệ thông tin lấy từ Niên giám Thống kê 2024 (chỉ công bố theo tỉnh từ năm 2024) nên không đưa được vào mô hình bảng. Kết quả chỉ mang tính tương quan cắt ngang một năm. Nguồn: tính toán của tác giả.")

json.dump(T, open(os.path.join(OUT, "bang_docx_v7.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("OK", list(T))
for k in ["b2", "b4", "b7"]:
    print(T[k]["title"]); [print(r) for r in T[k]["rows"]]

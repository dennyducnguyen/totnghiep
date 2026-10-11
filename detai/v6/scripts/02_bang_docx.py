# -*- coding: utf-8 -*-
"""Định dạng các bảng kết quả (ket-qua/*.csv) thành JSON để dựng file Word (scripts/03_build_docx.js)."""
import json, os
import numpy as np, pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "ket-qua")
KQ = json.load(open(os.path.join(OUT, "ket_qua.json"), encoding="utf-8"))
REG_ORDER = ["ĐNB", "ĐBSH", "BTB&DHMT", "ĐBSCL", "TN", "TDMNPB"]
REG_FULL = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
VN = {"ln_grdp": "ln(GRDP/người)", "dn_per_1000": "Mật độ doanh nghiệp (DN/1.000 dân)",
      "urban_share_pct": "Tỷ lệ dân số thành thị (%)", "trained_labor_pct": "Tỷ lệ LĐ qua đào tạo (%)",
      "ln_freight": "ln(Hàng hóa vận chuyển đường bộ/người)", "modern_retail_100k": "Siêu thị + TTTM/100.000 dân",
      "ln_fdi_pc": "ln(1 + FDI lũy kế/người)", "hh_computer_pct": "Tỷ lệ hộ có máy tính (%)",
      "internet_user_pct": "Tỷ lệ người sử dụng Internet (%)", "it_skills_pct": "Tỷ lệ người có kỹ năng CNTT (%)",
      "Intercept": "Hằng số", "const": "Hằng số"}

def f(x, d=3):
    if x is None or (isinstance(x, float) and np.isnan(x)): return ""
    return f"{x:,.{d}f}".replace(",", " ").replace(".", ",").replace("-", "−")   # kiểu Việt: dấu phẩy thập phân
def fi(x): return "" if pd.isna(x) else f"{int(round(x)):,}".replace(",", ".")
def stars(p):
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.1 else ""
def cell(coef, se, p, d=3):
    return f"{f(coef, d)}{stars(p)}\n({f(se, d)})"

T = {}
# ---------------- Bảng 1 ----------------
b1 = pd.read_csv(os.path.join(OUT, "bang1_mo_ta.csv"))
rows = []
for _, r in b1.iterrows():
    d = 2 if r["mean"] < 100 else 1 if r["mean"] < 1000 else 0
    rows.append([r["ten"], r["don_vi"], r["nguon"], str(int(r["N"])), f(r["mean"], d), f(r["sd"], d), f(r["min"], d), f(r["max"], d)])
T["b1"] = dict(title="Bảng 1. Mô tả biến và thống kê mô tả (63 tỉnh, 2017–2024)",
               header=["Biến", "Đơn vị", "Nguồn", "N", "Trung bình", "Độ lệch chuẩn", "Nhỏ nhất", "Lớn nhất"],
               rows=rows, widths=[2900, 1000, 800, 600, 1000, 1000, 900, 900],
               note="Nguồn: tính toán của tác giả từ Báo cáo Chỉ số Thương mại điện tử Việt Nam (VECOM, 2017–2024) và Niên giám Thống kê (NGTK) các năm 2017–2024. Số quan sát tối đa là 504 (63 tỉnh × 8 năm); EBI khuyết ở 5–9 tỉnh miền núi mỗi năm; GRDP, thu nhập, HDI công bố từ 2018; tỷ lệ hộ có máy tính công bố từ 2019; trụ cột G2B bị bỏ từ 2021.")

# ---------------- Bảng 2 ----------------
b2 = pd.read_csv(os.path.join(OUT, "bang2_vung.csv")).set_index("vung6").loc[REG_ORDER]
rows = []
for r, x in b2.iterrows():
    rows.append([REG_FULL[r], str(int(x["n_tinh"])), f(x["2017–2020"], 1), f(x["2021–2024"], 1), f(x["EBI_rel"], 1), f(x["NNL_HT_rel"], 1), f(x["B2C_rel"], 1), f(x["B2B_rel"], 1)])
hi, lo = b2.loc["ĐNB"], b2.loc[["TDMNPB", "TN", "ĐBSCL"]].min()
rows.append(["Tỷ số vùng cao nhất / vùng thấp nhất", "", f(hi["2017–2020"] / lo["2017–2020"], 2), f(hi["2021–2024"] / lo["2021–2024"], 2),
             f(hi["EBI_rel"] / lo["EBI_rel"], 2), f(hi["NNL_HT_rel"] / lo["NNL_HT_rel"], 2), f(hi["B2C_rel"] / lo["B2C_rel"], 2), f(hi["B2B_rel"] / lo["B2B_rel"], 2)])
T["b2"] = dict(title="Bảng 2. EBI tương đối theo 6 vùng kinh tế – xã hội (trung bình các tỉnh có số liệu trong năm = 100)",
               header=["Vùng", "Số tỉnh", "EBI 2017–2020", "EBI 2021–2024", "EBI 2024", "NNL&HT 2024", "B2C 2024", "B2B 2024"],
               rows=rows, widths=[2900, 700, 1100, 1100, 900, 1100, 900, 900],
               note="Ghi chú: EBI tương đối = EBI của tỉnh / trung bình giản đơn của các tỉnh có số liệu trong cùng năm × 100; cách chuẩn hóa này cho phép so sánh qua điểm gãy phương pháp năm 2021. Trung bình vùng là trung bình giản đơn của các tỉnh trong vùng. NNL&HT: trụ cột nguồn nhân lực và hạ tầng CNTT. Nguồn: tính toán của tác giả.")

# ---------------- Bảng 3 ----------------
b3 = pd.read_csv(os.path.join(OUT, "bang3_theil.csv")); ex = pd.read_csv(os.path.join(OUT, "phu_theil_bo_HN_HCM.csv"))
rows = []
for (_, r), (_, e) in zip(b3.iterrows(), ex.iterrows()):
    rows.append([str(int(r.nam)), str(int(r.N)), f(r.Gini), f(r.CV), f(r.Theil), f(r.Theil_between), f(r.Theil_within), f(r.Theil_between_share * 100, 1), f(r.MLD_between_share * 100, 1), f(e.Theil_ex2), f(e.between_share_ex2 * 100, 1)])
T["b3"] = dict(title="Bảng 3. Mức độ và cấu trúc bất bình đẳng EBI giữa các tỉnh: Gini, CV và phân rã Theil theo 6 vùng, 2017–2024",
               header=["Năm", "N", "Gini", "CV", "Theil-T", "Giữa các vùng", "Trong vùng", "Tỷ trọng giữa vùng (%)", "Tỷ trọng giữa vùng, MLD (%)", "Theil-T loại HN & TP.HCM", "Tỷ trọng giữa vùng khi loại HN & TP.HCM (%)"],
               rows=rows, widths=[600, 500, 700, 700, 800, 900, 800, 1000, 1000, 1000, 1200],
               note="Ghi chú: mỗi tỉnh là một quan sát có trọng số bằng nhau. Theil-T = Σ(y_i/ȳ)ln(y_i/ȳ)/N, phân rã theo Theil (1967): T = T_giữa + T_trong. MLD (mean log deviation) là chỉ số GE(0), nhạy hơn với đuôi dưới của phân phối. Các chỉ số này bất biến với thang đo nên so sánh được trong từng năm; mức nhảy năm 2021 phản ánh việc VECOM thay đổi phương pháp tính EBI (xem mục 3.1). Nguồn: tính toán của tác giả.")

# ---------------- Bảng 4 ----------------
b4 = pd.read_csv(os.path.join(OUT, "bang4_gini_tru_cot.csv"))
rows = []
for _, r in b4.iterrows():
    rows.append([str(int(r.nam)), f(r.Gini_EBI), f(r.Gini_NNL_HT), f(r.Gini_B2C), f(r.Gini_B2B), f(r.get("Gini_G2B", np.nan)),
                 f(r.share_NNL_HT * 100, 1), f(r.share_B2C * 100, 1), f(r.share_B2B * 100, 1), f(r.get("share_G2B", np.nan) * 100 if not pd.isna(r.get("share_G2B", np.nan)) else np.nan, 1)])
T["b4"] = dict(title="Bảng 4. Hệ số Gini của EBI và từng trụ cột; đóng góp của từng trụ cột vào Gini EBI (phân rã Gini theo nguồn), 2017–2024",
               header=["Năm", "Gini EBI", "Gini NNL&HT", "Gini B2C", "Gini B2B", "Gini G2B", "Đóng góp NNL&HT (%)", "Đóng góp B2C (%)", "Đóng góp B2B (%)", "Đóng góp G2B (%)"],
               rows=rows, widths=[600, 800, 900, 800, 800, 800, 1100, 1000, 1000, 1000],
               note="Ghi chú: EBI là tổng có trọng số của các trụ cột (trọng số phục hồi từ dữ liệu: 2017 là 0,3/0,3/0,3/0,1; 2018–2020 là 0,2/0,35/0,35/0,1 cho NNL&HT/B2C/B2B/G2B; từ 2021 là 0,2/0,4/0,4 cho NNL&HT/B2C/B2B). Theo phân rã Gini theo nguồn, G = Σ S_k G_k R_k, trong đó S_k là tỷ trọng của trụ cột k trong EBI, G_k là Gini của trụ cột và R_k là tương quan Gini giữa trụ cột với thứ hạng EBI; đóng góp (%) = S_k G_k R_k / G × 100. Cột “Gini EBI” của bảng này được tính trên EBI tái tạo từ các trụ cột theo trọng số nêu trên, nên có thể lệch chút ít so với cột Gini của Bảng 3 (tính trên EBI do VECOM công bố); chênh lệch chỉ thấy được ở năm 2023 (0,258 so với 0,262) và bắt nguồn từ việc VECOM làm tròn điểm trụ cột. Nguồn: tính toán của tác giả.")

# ---------------- Bảng 5 ----------------
b5 = pd.read_csv(os.path.join(OUT, "bang5b_beta.csv"))
lab = {0: "Tuyệt đối, 2017–2020 (thang cũ)", 1: "Tuyệt đối, 2021–2024 (thang mới)", 2: "Tuyệt đối, 2017–2024 (EBI tương đối)",
       5: "Có điều kiện, 2017–2020", 6: "Có điều kiện, 2021–2024", 7: "Có điều kiện, 2017–2024 (EBI tương đối)"}
rows = []
for i, name in lab.items():
    r = b5.iloc[i]
    rows.append([name, cell(r.beta, r.se, r.p, 4), f(r.lam * 100, 1) if r.lam > 0 else "—", f(r.half_life, 1) if r.lam > 0 else "—", str(int(r.N)), f(r.R2, 3)])
bp = KQ["beta_panel"]
rows.append(["Bảng động (tăng trưởng hằng năm EBI tương đối, hiệu ứng năm, loại bước nhảy 2020→2021)",
             cell(bp["beta"], bp["se"], bp["p"], 4), f(-np.log(1 + bp["beta"]) * 100, 1), f(np.log(2) / (-np.log(1 + bp["beta"])), 1), str(bp["N"]), ""])
T["b5"] = dict(title="Bảng 5. Kiểm định hội tụ beta của EBI giữa các tỉnh",
               header=["Mô hình", "β (SE)", "Tốc độ hội tụ λ (%/năm)", "Thời gian bán hội tụ (năm)", "N", "R²"],
               rows=rows, widths=[3600, 1300, 1300, 1300, 600, 800],
               note="Ghi chú: mô hình (1/T)·ln(EBI_iT/EBI_i0) = α + β·ln(EBI_i0) + γ′Z_i + ε_i; sai số chuẩn vững (HC1) trong ngoặc; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%. λ = −ln(1 + βT)/T; thời gian bán hội tụ = ln2/λ. Biến điều kiện Z (giá trị năm gốc; riêng GRDP/người lấy năm 2018 cho gốc 2017): ln(GRDP/người), mật độ doanh nghiệp, tỷ lệ dân số thành thị, tỷ lệ lao động qua đào tạo. Dòng cuối: hồi quy bảng g_it = α_t + β·ln(EBI tương đối_i,t−1) + ε_it, sai số chuẩn cụm theo tỉnh. Nguồn: tính toán của tác giả.")
# hệ số biến điều kiện
rows = []
for v in ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct"]:
    rows.append([VN[v]] + [f"{f(b5.iloc[i]['b_' + v], 4)}{stars(b5.iloc[i]['p_' + v])}" for i in [5, 6, 7]])
T["b5c"] = dict(title="Bảng 5b. Hệ số của các biến điều kiện trong mô hình hội tụ beta có điều kiện",
                header=["Biến điều kiện (năm gốc)", "2017–2020", "2021–2024", "2017–2024 (EBI tương đối)"],
                rows=rows, widths=[3600, 1800, 1800, 1800],
                note="Ghi chú: ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10% (sai số chuẩn HC1). Nguồn: tính toán của tác giả.")

# ---------------- Bảng 6 ----------------
b6 = pd.read_csv(os.path.join(OUT, "bang6_ols.csv")); vif = pd.read_csv(os.path.join(OUT, "bang6_vif.csv")).set_index("bien")["VIF"]
models = ["M1: 2018–2024", "M2: 2019–2024 (+máy tính)", "M3: X trễ 1 năm", "M4: 2021–2024", "M2 chuẩn hóa"]
X = ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct", "hh_computer_pct", "ln_freight", "modern_retail_100k", "ln_fdi_pc"]
def get(m, term):
    q = b6[(b6.model == m) & (b6.term == term)]
    return None if q.empty else q.iloc[0]
rows = []
for v in X:
    row = [VN[v]]
    for m in models:
        term = v + "_L1" if m.startswith("M3") else v
        r = get(m, term); row.append("" if r is None else cell(r.coef, r.se, r.p))
    row.append(f(vif[v], 2)); rows.append(row)
for stat, name in [("N", "Số quan sát"), ("R2", "R²"), ("R2_adj", "R² hiệu chỉnh")]:
    row = [name]
    for m in models:
        r = get(m, stat); row.append(fi(r.coef) if stat == "N" else f(r.coef, 3))
    row.append(""); rows.append(row)
rows.append(["Hiệu ứng cố định năm"] + ["Có"] * 5 + [""])
T["b6"] = dict(title="Bảng 6. Các yếu tố quyết định EBI: hồi quy OLS gộp, biến phụ thuộc ln(EBI)",
               header=["Biến giải thích", "(1) 2018–2024", "(2) 2019–2024", "(3) X trễ 1 năm", "(4) 2021–2024", "(5) Hệ số chuẩn hóa của (2)", "VIF của (2)"],
               rows=rows, widths=[2500, 1200, 1200, 1200, 1200, 1300, 700],
               note="Ghi chú: sai số chuẩn cụm theo tỉnh trong ngoặc; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%. Cột (3) dùng toàn bộ biến giải thích trễ một năm (t−1). Cột (5) là hệ số beta chuẩn hóa (mọi biến trừ trung bình, chia độ lệch chuẩn). Hiệu ứng cố định năm hấp thụ thay đổi phương pháp EBI năm 2021. VIF tính trên các biến giải thích của cột (2), không gồm biến giả năm. Nguồn: tính toán của tác giả.")

# ---------------- Bảng 6b: 2024 cắt ngang ----------------
bc = pd.read_csv(os.path.join(OUT, "bang6b_2024_cross.csv"))
mods = ["C1: chỉ biến số", "C2: + kinh tế", "C3: máy tính + kinh tế"]
rows = []
for v in ["internet_user_pct", "it_skills_pct", "hh_computer_pct", "ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct"]:
    row = [VN[v]]
    for m in mods:
        q = bc[(bc.model == m) & (bc.term == v)]; row.append("" if q.empty else cell(q.iloc[0].coef, q.iloc[0].se, q.iloc[0].p))
    rows.append(row)
for stat, name in [("N", "Số quan sát"), ("R2_adj", "R² hiệu chỉnh")]:
    rows.append([name] + [fi(bc[(bc.model == m) & (bc.term == stat)].iloc[0].coef) if stat == "N" else f(bc[(bc.model == m) & (bc.term == stat)].iloc[0].coef, 3) for m in mods])
T["b6b"] = dict(title="Bảng 9. Kiểm định bổ sung với các chỉ tiêu số chi tiết (chỉ có năm 2024): OLS cắt ngang, biến phụ thuộc ln(EBI 2024)",
                header=["Biến giải thích", "(1) Chỉ biến số", "(2) Biến số + kinh tế", "(3) Máy tính + kinh tế"],
                rows=rows, widths=[3200, 1900, 1900, 1900],
                note="Ghi chú: 58 tỉnh có EBI năm 2024; sai số chuẩn vững HC1 trong ngoặc; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%. Tỷ lệ người sử dụng Internet và tỷ lệ người có kỹ năng CNTT lấy từ NGTK 2024 (chỉ công bố theo tỉnh từ năm 2024). Nguồn: tính toán của tác giả.")

# ---------------- Bảng 7 ----------------
b7 = pd.read_csv(os.path.join(OUT, "bang7_vung.csv")); mods = ["M5: chỉ vùng", "M6: X + vùng", "M7: FE tỉnh"]
def g7(m, t):
    q = b7[(b7.model == m) & (b7.term == t)]; return None if q.empty else q.iloc[0]
rows = []
for r in ["ĐBSH", "BTB&DHMT", "ĐBSCL", "TN", "TDMNPB"]:
    row = [REG_FULL[r] + " (so với Đông Nam Bộ)"]
    for m in mods:
        x = g7(m, f"C(vung6)[T.{r}]"); row.append("" if x is None else cell(x.coef, x.se, x.p))
    rows.append(row)
for v in X:
    row = [VN[v]]
    for m in mods:
        x = g7(m, v); row.append("" if x is None else cell(x.coef, x.se, x.p))
    rows.append(row)
for stat, name in [("N", "Số quan sát"), ("R2", "R²"), ("R2_adj", "R² hiệu chỉnh")]:
    rows.append([name] + [fi(g7(m, stat).coef) if stat == "N" else f(g7(m, stat).coef, 3) for m in mods])
rows.append(["Hiệu ứng cố định năm", "Có", "Có", "Có"]); rows.append(["Hiệu ứng cố định tỉnh", "Không", "Không", "Có"])
w = KQ["wald_vung_M6"]
T["b7"] = dict(title="Bảng 7. Khoảng cách vùng trước và sau khi kiểm soát các yếu tố cấu trúc: biến giả vùng và hiệu ứng cố định tỉnh, 2019–2024, biến phụ thuộc ln(EBI)",
               header=["Biến", "(6) Chỉ biến giả vùng", "(7) Biến giải thích + biến giả vùng", "(8) Hiệu ứng cố định tỉnh"],
               rows=rows, widths=[3400, 1800, 1900, 1800],
               note=f"Ghi chú: sai số chuẩn cụm theo tỉnh trong ngoặc; ***, **, * tương ứng mức ý nghĩa 1%, 5%, 10%. Nhóm gốc là Đông Nam Bộ. Kiểm định Wald cho giả thuyết 5 biến giả vùng đồng thời bằng 0 trong cột (7): χ²(5) = {w['stat']:.2f}, p = {w['p']:.3f}. Cột (8) chỉ khai thác biến thiên theo thời gian trong nội bộ từng tỉnh. Nguồn: tính toán của tác giả.")

# ---------------- Bảng 8 ----------------
b8 = pd.read_csv(os.path.join(OUT, "bang8_oaxaca.csv"))
rows = []
for _, r in b8.iterrows():
    ref = "Hệ số gộp (Jann, 2008)" if r.ref == "pooled" else "Hệ số vùng lõi"
    rows.append([r["sample"].replace("M2", "Bảng 6, cột 2").replace("M1", "Bảng 6, cột 1").replace("M4", "Bảng 6, cột 4"), ref, f"{int(r.nA)} / {int(r.nB)}", f(r.yA), f(r.yB), f(r.gap),
                 f(r.endow) + ("\n(" + f(r.endow_se) + ")" if not pd.isna(r.endow_se) else ""), f(r.endow_share * 100, 1),
                 f(r.coef) + ("\n(" + f(r.coef_se) + ")" if not pd.isna(r.coef_se) else ""), f(r.coef_share * 100, 1)])
T["b8"] = dict(title="Bảng 8. Phân tách Oaxaca–Blinder khoảng cách ln(EBI) giữa vùng lõi (Đông Nam Bộ, Đồng bằng sông Hồng) và vùng ngoại vi (4 vùng còn lại)",
               header=["Mẫu / bộ biến", "Hệ số tham chiếu", "N lõi / ngoại vi", "ln EBI lõi", "ln EBI ngoại vi", "Khoảng cách", "Phần nguồn lực", "%", "Phần hệ số", "%"],
               rows=rows, widths=[1500, 1300, 900, 700, 800, 800, 1000, 600, 1000, 600],
               note="Ghi chú: phân tách hai thành phần theo phương trình (4) ở Mục 3.3, với β* là véc-tơ hệ số tham chiếu; mô hình gồm biến giả năm. Sai số chuẩn bootstrap (300 lần, lấy mẫu lại theo tỉnh) trong ngoặc cho mẫu chính. Nguồn: tính toán của tác giả.")
det = pd.read_csv(os.path.join(OUT, "bang8_oaxaca_chi_tiet.csv"), index_col=0)
rows = []
for v in X:
    if v in det.index:
        rows.append([VN[v], f(det.loc[v, "endow"]), f(det.loc[v, "endow_share"] * 100, 1), f(det.loc[v, "coef"])])
ysum = det.loc[[i for i in det.index if i.startswith("y_")]].sum()
rows.append(["Biến giả năm (gộp)", f(ysum.endow), f(ysum.endow / b8.iloc[0].gap * 100, 1), f(ysum.coef)])
rows.append(["Hằng số", "", "", f(det.loc["const", "coef"])])
rows.append(["Tổng", f(det.endow.sum()), f(det.endow.sum() / b8.iloc[0].gap * 100, 1), f(det.coef.sum())])
T["b8b"] = dict(title="Bảng 8b. Đóng góp chi tiết của từng biến vào phần nguồn lực (mẫu 2019–2024, hệ số tham chiếu gộp)",
                header=["Biến", "Phần nguồn lực", "% khoảng cách", "Phần hệ số"],
                rows=rows, widths=[3600, 1600, 1600, 1600],
                note="Ghi chú: phần hệ số theo từng biến phụ thuộc vào gốc đo của biến nên chỉ có ý nghĩa ở tổng số. Nguồn: tính toán của tác giả.")

json.dump(T, open(os.path.join(OUT, "bang_docx.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("OK", list(T))
for k in ["b3", "b5", "b8"]:
    print(T[k]["title"]); [print(r) for r in T[k]["rows"]]

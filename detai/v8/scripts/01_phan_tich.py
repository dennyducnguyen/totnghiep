# -*- coding: utf-8 -*-
"""
BƯỚC 1 – PHÂN TÍCH ĐỊNH LƯỢNG CHO BÀI BÁO BẢN v8
"Khoảng cách số trong thương mại điện tử cấp tỉnh tại Việt Nam giai đoạn 2017–2024:
 mức độ, cấu trúc, hội tụ và các yếu tố liên quan"

Chạy (trong thư mục v8):   python scripts/01_phan_tich.py
Đầu vào : du-lieu/EBI_NGTK_Panel_2017_2024.xlsx (sheet Panel_2017_2024 và Digital_2024_cross)
Đầu ra  : ket-qua/*.csv, ket-qua/ket_qua.json (các con số dùng trong bài), ket-qua/panel_phan_tich.csv

Thiết kế (theo hướng của bản nháp v8/phantichketqua.docx, đã sửa các điểm phương pháp):
  RQ1  Mức độ và cấu trúc khoảng cách
       - CV, Gini, Theil-T theo từng năm; DTF (khoảng cách tới tỉnh dẫn đầu) trung bình; Max/Min, P90/P10.
       - H1: phân rã Theil-T theo 6 vùng (giữa vùng / trong vùng).
       - H2: Gini từng trụ cột (NNL&HT, B2C, B2B, G2B).
  RQ2  Hội tụ
       - H3: hội tụ sigma = CV, Gini, SD ln(EBI) giảm trong giai đoạn thang đo ổn định 2022–2024.
       - H4: hội tụ beta dạng bảng năm: Δln DTF_it = α + β ln DTF_i,t−1 + λ_t + ε, SE cụm tỉnh.
             Mô hình chính: tăng trưởng 2023–2024 (gốc 2022–2023, cùng thang đo ổn định).
             KHÔNG lấy 2021 làm gốc vì thang điểm trụ cột năm 2021 khác hẳn các năm sau.
             So sánh: giai đoạn 2018–2020 (thang cũ). Kiểm tra vững: loại HN & TP.HCM, loại 5 tỉnh dẫn đầu,
             biến giả vùng, sai phân dài 2022→2024.
  RQ3  Các yếu tố liên quan
       - Ma trận tương quan Pearson.
       - M2–M5: OLS gộp, Y = DTF (0–100) của EBI hoặc trụ cột, hiệu ứng năm, SE cụm tỉnh.
       - M6: hồi quy cắt ngang 2024 với tỷ lệ người có kỹ năng CNTT.
       - Kiểm tra vững: hiệu ứng cố định tỉnh + năm; riêng 2018–2020 (trước khi EBI chứa chỉ tiêu DN/dân số và
         thu nhập); riêng 2022–2024; loại HN & TP.HCM; biến giải thích trễ 1 năm.
"""
import json, os, warnings
import numpy as np, pandas as pd
import statsmodels.api as sm
import statsmodels.formula.api as smf
from scipy import stats
from linearmodels.panel import PanelOLS

warnings.filterwarnings("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))          # thư mục v8/
DATA = os.path.join(ROOT, "du-lieu", "EBI_NGTK_Panel_2017_2024.xlsx")
OUT = os.path.join(ROOT, "ket-qua"); os.makedirs(OUT, exist_ok=True)
KQ = {}

REG_ORDER = ["ĐNB", "ĐBSH", "BTB&DHMT", "ĐBSCL", "TN", "TDMNPB"]
REG_FULL = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
HAI_TP = ["Hà Nội", "TP. Hồ Chí Minh"]
TU_CHU = {2017: "Thang cũ (4 trụ cột)", 2018: "Thang cũ (4 trụ cột)", 2019: "Thang cũ (4 trụ cột)",
          2020: "Thang cũ (4 trụ cột)", 2021: "Chuyển tiếp (3 trụ cột)", 2022: "Ổn định (3 trụ cột)",
          2023: "Ổn định (3 trụ cột)", 2024: "Ổn định (3 trụ cột)"}
PILLARS = ["NNL_HT", "B2C", "B2B", "G2B"]

def save(df, name, index=False):
    df.to_csv(os.path.join(OUT, name), index=index, encoding="utf-8-sig")

# =====================================================================
# 1. DỮ LIỆU VÀ BIẾN
# =====================================================================
df = pd.read_excel(DATA, sheet_name="Panel_2017_2024").drop(columns=["vung_ebi"], errors="ignore")
dig = pd.read_excel(DATA, sheet_name="Digital_2024_cross")
df = df.sort_values(["tinh", "nam"]).reset_index(drop=True)

# DTF (distance to frontier): điểm của tỉnh so với tỉnh dẫn đầu cùng năm (tỉnh dẫn đầu = 100).
# Chỉ số này bất biến khi cả năm bị nhân với một hằng số, nên so sánh được vị thế qua các thang đo.
for c in ["EBI"] + PILLARS:
    df[c + "_dtf"] = df[c] / df.groupby("nam")[c].transform("max") * 100
df["ln_dtf"] = np.log(df["EBI_dtf"])
df["ln_retail"] = np.log(df["retail_pc_tr"])
df["ln_grdp"] = np.log(df["grdp_pc_tr"])
df["cathe_per_1000"] = df["coso_cathe"] / df["pop_th"]           # cơ sở cá thể / 1.000 dân (pop_th: nghìn người)
df["ln_fdi_pc"] = np.log1p(df["fdi_pc_usd"])
X_ALL = ["ln_retail", "ln_grdp", "dn_per_1000", "cathe_per_1000", "trained_labor_pct", "urban_share_pct",
         "ln_fdi_pc", "hh_computer_pct"]
for v in X_ALL:
    df[v + "_L1"] = df.groupby("tinh")[v].shift(1)
df["lag_nam"] = df.groupby("tinh")["nam"].shift(1)
df["ln_dtf_L1"] = df.groupby("tinh")["ln_dtf"].shift(1)
df["g_dtf"] = df["ln_dtf"] - df["ln_dtf_L1"]
df.loc[df["lag_nam"] != df["nam"] - 1, ["g_dtf", "ln_dtf_L1"]] = np.nan
df = df.merge(dig[["tinh", "internet_user_pct", "hh_internet_pct", "social_media_pct", "it_skills_pct"]]
              .assign(nam=2024), on=["tinh", "nam"], how="left")
save(df, "panel_phan_tich.csv")
d = df.dropna(subset=["EBI"]).copy()                       # 445 quan sát tỉnh – năm có EBI
cnt = d.groupby("tinh").size()
KQ["N_obs"] = int(len(d)); KQ["N_tinh_co_EBI"] = int(d.tinh.nunique())
KQ["N_tinh_du_8_nam"] = int((cnt == 8).sum())
KQ["N_theo_nam"] = d.groupby("nam").size().astype(int).to_dict()
KQ["tinh_khong_bao_gio_co"] = sorted(set(df.tinh) - set(d.tinh))
KQ["tinh_thieu_theo_nam"] = {int(y): sorted(df[(df.nam == y) & df.EBI.isna()].tinh) for y in range(2017, 2025)}

# =====================================================================
# 2. HÀM ĐO BẤT BÌNH ĐẲNG
# =====================================================================
def gini(x):
    x = np.sort(np.asarray(x, float)); n = len(x)
    return (2 * np.sum(np.arange(1, n + 1) * x) / (n * x.sum())) - (n + 1) / n
def theil_T(x):
    x = np.asarray(x, float); m = x.mean(); return np.mean(x / m * np.log(x / m))
def theil_decomp(x, g):
    x = np.asarray(x, float); g = np.asarray(g); n = len(x); m = x.mean()
    T = theil_T(x); within = between = 0.0; parts = {}
    for k in np.unique(g):
        xk = x[g == k]; sk = len(xk) / n * xk.mean() / m
        within += sk * theil_T(xk); between += sk * np.log(xk.mean() / m)
        parts[k] = dict(n=len(xk), s_k=sk, Theil_k=theil_T(xk), mean_dtf=xk.mean())
    return T, between, within, parts
def cv(x):
    x = np.asarray(x, float); return x.std(ddof=1) / x.mean()

# =====================================================================
# 3. BẢNG 1 – THỐNG KÊ MÔ TẢ BIẾN
# =====================================================================
VARS = [("EBI", "Chỉ số TMĐT tổng hợp (EBI)", "điểm", "VECOM"),
        ("EBI_dtf", "Khoảng cách tới tỉnh dẫn đầu của EBI (DTF)", "%", "Tính toán"),
        ("NNL_HT", "Trụ cột Nguồn nhân lực & hạ tầng CNTT", "điểm", "VECOM"),
        ("B2C", "Trụ cột giao dịch B2C", "điểm", "VECOM"),
        ("B2B", "Trụ cột giao dịch B2B", "điểm", "VECOM"),
        ("G2B", "Trụ cột giao dịch G2B (2017–2020)", "điểm", "VECOM"),
        ("retail_pc_tr", "Bán lẻ hàng hóa & dịch vụ bình quân đầu người", "triệu đồng", "NGTK"),
        ("grdp_pc_tr", "GRDP bình quân đầu người (từ 2018)", "triệu đồng", "NGTK"),
        ("dn_per_1000", "Mật độ doanh nghiệp đang hoạt động", "DN/1.000 dân", "NGTK"),
        ("cathe_per_1000", "Mật độ cơ sở kinh tế cá thể phi nông nghiệp", "cơ sở/1.000 dân", "NGTK"),
        ("trained_labor_pct", "Tỷ lệ lao động 15+ đã qua đào tạo", "%", "NGTK"),
        ("urban_share_pct", "Tỷ lệ dân số thành thị", "%", "NGTK"),
        ("fdi_pc_usd", "Vốn FDI đăng ký lũy kế/người", "USD/người", "NGTK"),
        ("hh_computer_pct", "Tỷ lệ hộ có máy tính (từ 2019)", "%", "NGTK"),
        ("it_skills_pct", "Tỷ lệ dân số có kỹ năng CNTT (2024)", "%", "NGTK")]
rows = []
for v, name, unit, src in VARS:
    s = (dig["it_skills_pct"] if v == "it_skills_pct" else df[v]).dropna()
    rows.append(dict(bien=v, ten=name, don_vi=unit, nguon=src, N=len(s), mean=s.mean(), sd=s.std(),
                     min=s.min(), max=s.max()))
save(pd.DataFrame(rows), "bang1_mo_ta.csv")

# =====================================================================
# 4. BẢNG 2 – MỨC ĐỘ KHOẢNG CÁCH THEO NĂM + PHÂN RÃ THEIL (RQ1, H1)
# =====================================================================
rows, parts_rows = [], []
bal = d[d.tinh.isin(cnt[cnt == 8].index)]
for y, g in d.groupby("nam"):
    T, Tb, Tw, parts = theil_decomp(g.EBI, g.vung6)
    g2 = g[~g.tinh.isin(HAI_TP)]; T2, Tb2, Tw2, _ = theil_decomp(g2.EBI, g2.vung6)
    gb = bal[bal.nam == y]; T3, Tb3, Tw3, _ = theil_decomp(gb.EBI, gb.vung6)
    rows.append(dict(nam=y, thang_do=TU_CHU[y], N=len(g), mean=g.EBI.mean(), median=g.EBI.median(),
                     min=g.EBI.min(), max=g.EBI.max(), CV=cv(g.EBI), Gini=gini(g.EBI), Theil=T,
                     Theil_between=Tb, Theil_within=Tw, within_share=Tw / T * 100, between_share=Tb / T * 100,
                     SD_ln=np.log(g.EBI).std(ddof=1), DTF_mean=g.EBI_dtf.mean(),
                     max_min=g.EBI.max() / g.EBI.min(), p90_p10=g.EBI.quantile(.9) / g.EBI.quantile(.1),
                     Gini_ex2=gini(g2.EBI), Theil_ex2=T2, within_share_ex2=Tw2 / T2 * 100, CV_ex2=cv(g2.EBI),
                     N_bal=len(gb), Gini_bal=gini(gb.EBI), within_share_bal=Tw3 / T3 * 100, CV_bal=cv(gb.EBI)))
    for k, p in parts.items():
        parts_rows.append(dict(nam=y, vung=k, **p, dong_gop_pct=p["s_k"] * p["Theil_k"] / T * 100))
b2 = pd.DataFrame(rows); save(b2, "bang2_khoang_cach_theo_nam.csv")
pr = pd.DataFrame(parts_rows); save(pr, "phu_theil_theo_vung.csv")
KQ["bang2"] = b2.round(4).to_dict("records")
KQ["within_share_range"] = [float(b2.within_share.min()), float(b2.within_share.max())]
KQ["within_share_range_ex2"] = [float(b2.within_share_ex2.min()), float(b2.within_share_ex2.max())]
KQ["within_share_range_bal"] = [float(b2.within_share_bal.min()), float(b2.within_share_bal.max())]
KQ["theil_giam_khi_bo_2tp_pct"] = {int(r.nam): float((1 - r.Theil_ex2 / r.Theil) * 100) for r in b2.itertuples()}
KQ["dong_gop_vung_2024"] = pr[pr.nam == 2024].set_index("vung")["dong_gop_pct"].round(1).to_dict()
KQ["theil_k_2024"] = pr[pr.nam == 2024].set_index("vung")["Theil_k"].round(3).to_dict()
d24 = d[d.nam == 2024]
KQ["top5_2024"] = d24.nlargest(5, "EBI")[["tinh", "EBI"]].values.tolist()
KQ["bot5_2024"] = d24.nsmallest(6, "EBI")[["tinh", "EBI"]].values.tolist()
TOP5 = [t for t, _ in KQ["top5_2024"]]

# DTF trung bình theo vùng và năm (Hình 3)
reg = d.groupby(["nam", "vung6"])["EBI_dtf"].mean().unstack()[REG_ORDER]
save(reg.round(2), "phu_dtf_vung_theo_nam.csv", index=True)
KQ["dtf_vung_2024"] = reg.loc[2024].round(1).to_dict()
KQ["dtf_vung_2017"] = reg.loc[2017].round(1).to_dict()

# =====================================================================
# 5. BẢNG 3 – GINI TỪNG TRỤ CỘT (H2)
# =====================================================================
rows = []
for y, g in d.groupby("nam"):
    r = dict(nam=y, Gini_EBI=gini(g.EBI))
    for p in PILLARS:
        s = g[p].dropna()
        r["Gini_" + p] = gini(s) if len(s) else np.nan
        r["CV_" + p] = cv(s) if len(s) else np.nan
    rows.append(r)
b3 = pd.DataFrame(rows); save(b3, "bang3_gini_tru_cot.csv")
KQ["bang3"] = b3.round(4).to_dict("records")

# =====================================================================
# 6. HỘI TỤ SIGMA (H3) – kiểm định xu hướng đơn giản trong từng đoạn thang đo
# =====================================================================
sig = b2[["nam", "CV", "Gini", "Theil", "SD_ln"]].copy(); save(sig, "phu_sigma.csv")
KQ["sigma"] = {k: {int(r.nam): float(getattr(r, k)) for r in sig.itertuples()} for k in ["CV", "Gini", "Theil", "SD_ln"]}

# =====================================================================
# 7. HỘI TỤ BETA (H4)
# =====================================================================
def beta_panel(data, label, years, controls=None, region=False):
    dd = data[data.nam.isin(years)].dropna(subset=["g_dtf", "ln_dtf_L1"] + (controls or []))
    f = "g_dtf ~ ln_dtf_L1 + C(nam)" + ("".join(" + " + c for c in controls) if controls else "") + (" + C(vung6)" if region else "")
    m = smf.ols(f, data=dd).fit(cov_type="cluster", cov_kwds={"groups": dd["tinh"]})
    b = m.params["ln_dtf_L1"]
    out = dict(mo_hinh=label, nam_tang_truong=f"{min(years)}–{max(years)}", N=int(m.nobs), N_tinh=int(dd.tinh.nunique()),
               beta=b, se=m.bse["ln_dtf_L1"], t=m.tvalues["ln_dtf_L1"], p=m.pvalues["ln_dtf_L1"], R2=m.rsquared,
               lambda_pct=(-np.log(1 + b) * 100 if 1 + b > 0 else np.nan))
    for c in controls or []:
        out["coef_" + c] = m.params[c]; out["p_" + c] = m.pvalues[c]; out["se_" + c] = m.bse[c]
    return out

CTRL = ["ln_grdp_L1", "dn_per_1000_L1", "urban_share_pct_L1", "trained_labor_pct_L1"]
POST, PRE = [2023, 2024], [2018, 2019, 2020]
d_ex2 = d[~d.tinh.isin(HAI_TP)]; d_ex5 = d[~d.tinh.isin(TOP5)]
beta_rows = [
    beta_panel(d, "1A. Tuyệt đối – thang cũ", PRE),
    beta_panel(d, "1A. Tuyệt đối – thang ổn định", POST),
    beta_panel(d, "1B. Có điều kiện – thang cũ", [2019, 2020], CTRL),
    beta_panel(d, "1B. Có điều kiện – thang ổn định", POST, CTRL),
]
rob_rows = [
    beta_panel(d, "Gốc: 1B có điều kiện, 63 tỉnh có số liệu", POST, CTRL),
    beta_panel(d_ex2, "R1. Loại Hà Nội và TP. Hồ Chí Minh", POST, CTRL),
    beta_panel(d_ex5, "R2. Loại 5 tỉnh dẫn đầu năm 2024", POST, CTRL),
    beta_panel(d, "R3. Thêm biến giả 6 vùng", POST, CTRL, region=True),
    beta_panel(d_ex2, "R4. Tuyệt đối, loại Hà Nội và TP. Hồ Chí Minh", POST),
    beta_panel(d, "R5. Tuyệt đối, tính cả bước 2021→2022", [2022, 2023, 2024]),
    beta_panel(d_ex2, "R6. Như R5, loại Hà Nội và TP. Hồ Chí Minh", [2022, 2023, 2024]),
    beta_panel(d_ex2, "R7. Thang cũ, loại Hà Nội và TP. Hồ Chí Minh", PRE),
]
# sai phân dài 2022→2024 (cắt ngang)
def beta_long(data, y0, y1, label):
    a = data[data.nam == y0].set_index("tinh")["EBI_dtf"]; b = data[data.nam == y1].set_index("tinh")["EBI_dtf"]
    x = pd.concat([a.rename("y0"), b.rename("y1")], axis=1).dropna()
    x["g"] = np.log(x.y1 / x.y0) / (y1 - y0); x["l"] = np.log(x.y0)
    m = smf.ols("g ~ l", data=x).fit(cov_type="HC1")
    return dict(mo_hinh=label, nam_tang_truong=f"{y0}→{y1}", N=int(m.nobs), N_tinh=int(m.nobs), beta=m.params["l"],
                se=m.bse["l"], t=m.tvalues["l"], p=m.pvalues["l"], R2=m.rsquared, lambda_pct=np.nan)
rob_rows.append(beta_long(d, 2022, 2024, "R8. Sai phân dài 2022→2024 (cắt ngang)"))
rob_rows.append(beta_long(d_ex2, 2022, 2024, "R9. Như R8, loại Hà Nội và TP. Hồ Chí Minh"))
rob_rows.append(beta_long(d, 2017, 2020, "R10. Sai phân dài 2017→2020 (thang cũ)"))
bb = pd.DataFrame(beta_rows); save(bb, "bang4_hoi_tu_beta.csv")
br = pd.DataFrame(rob_rows); save(br, "bang5_beta_do_vung.csv")
KQ["beta"] = bb.round(4).to_dict("records"); KQ["beta_rob"] = br.round(4).to_dict("records")
# dữ liệu cho hình hội tụ beta (gộp 2 năm tăng trưởng 2023–2024)
sc = d[d.nam.isin(POST)].dropna(subset=["g_dtf", "ln_dtf_L1"])[["tinh", "nam", "vung6", "ln_dtf_L1", "g_dtf"]]
save(sc, "phu_hinh_beta.csv")

# =====================================================================
# 8. TƯƠNG QUAN (Hình 6) VÀ HỒI QUY M2–M5 (RQ3)
# =====================================================================
CORR_V = ["EBI_dtf", "NNL_HT_dtf", "B2C_dtf", "B2B_dtf", "ln_retail", "ln_grdp", "dn_per_1000", "cathe_per_1000",
          "trained_labor_pct", "urban_share_pct", "ln_fdi_pc", "hh_computer_pct"]
cc = d[CORR_V].corr(); save(cc.round(3), "phu_tuong_quan.csv", index=True)
pv = {}
for v in CORR_V[4:]:
    s = d[["EBI_dtf", v]].dropna(); r, p = stats.pearsonr(s.EBI_dtf, s[v]); pv[v] = dict(r=r, p=p, N=len(s))
    s = d[["NNL_HT_dtf", v]].dropna(); r2, p2 = stats.pearsonr(s.NNL_HT_dtf, s[v]); pv[v].update(r_nnl=r2, p_nnl=p2)
KQ["tuong_quan_EBI"] = pv

MODELS = {
    "M2": ("EBI_dtf", ["ln_retail", "ln_grdp", "dn_per_1000", "cathe_per_1000", "trained_labor_pct", "urban_share_pct", "ln_fdi_pc"], 2018),
    "M3": ("B2C_dtf", ["ln_retail", "ln_grdp", "urban_share_pct"], 2018),
    "M4": ("B2B_dtf", ["dn_per_1000", "cathe_per_1000", "trained_labor_pct", "ln_fdi_pc"], 2018),
    "M5": ("NNL_HT_dtf", ["trained_labor_pct", "urban_share_pct", "ln_grdp", "hh_computer_pct"], 2019),
}
def pooled(data, y, X, name, lag=False):
    Xu = [x + "_L1" for x in X] if lag else X
    dd = data.dropna(subset=Xu + [y])
    m = smf.ols(f"{y} ~ " + " + ".join(Xu) + " + C(nam)", data=dd).fit(cov_type="cluster", cov_kwds={"groups": dd["tinh"]})
    out = [dict(model=name, term=x, coef=float(m.params[xu]), se=float(m.bse[xu]), p=float(m.pvalues[xu])) for x, xu in zip(X, Xu)]
    out += [dict(model=name, term="N", coef=int(m.nobs)), dict(model=name, term="N_tinh", coef=int(dd.tinh.nunique())),
            dict(model=name, term="R2", coef=float(m.rsquared)), dict(model=name, term="nam", coef=f"{dd.nam.min()}–{dd.nam.max()}")]
    return out
def fe2(data, y, X, name):
    dd = data.dropna(subset=X + [y]).set_index(["tinh", "nam"])
    m = PanelOLS(dd[y], dd[X], entity_effects=True, time_effects=True, drop_absorbed=True).fit(cov_type="clustered", cluster_entity=True)
    out = [dict(model=name, term=x, coef=float(m.params[x]), se=float(m.std_errors[x]), p=float(m.pvalues[x])) for x in X if x in m.params]
    out += [dict(model=name, term="N", coef=int(m.nobs)), dict(model=name, term="N_tinh", coef=int(m.entity_info["total"])),
            dict(model=name, term="R2", coef=float(m.rsquared)),
            dict(model=name, term="nam", coef=f"{dd.index.get_level_values(1).min()}–{dd.index.get_level_values(1).max()}")]
    return out

rows = []
for name, (y, X, y0) in MODELS.items():
    rows += pooled(d[d.nam >= y0], y, X, name)
b6 = pd.DataFrame(rows); save(b6, "bang6_hoi_quy_M2_M5.csv")

y2, X2, _ = MODELS["M2"]; y3, X3, _ = MODELS["M3"]; y4, X4, _ = MODELS["M4"]
rows = []
rows += fe2(d[d.nam >= 2018], y2, X2, "A1. M2 – FE tỉnh + năm")
rows += pooled(d[d.nam.between(2018, 2020)], y2, X2, "A2. M2 – 2018–2020")
rows += pooled(d[d.nam >= 2022], y2, X2, "A3. M2 – 2022–2024")
rows += pooled(d[(d.nam >= 2018) & ~d.tinh.isin(HAI_TP)], y2, X2, "A4. M2 – loại HN & TP.HCM")
rows += pooled(d[d.nam >= 2019], y2, X2, "A5. M2 – X trễ 1 năm", lag=True)
rows += pooled(d[d.nam.between(2018, 2020)], y3, X3, "A6. M3 – 2018–2020")
rows += pooled(d[d.nam.between(2018, 2020)], y4, X4, "A7. M4 – 2018–2020")
rows += pooled(d[d.nam >= 2019], y2, X2 + ["hh_computer_pct"], "A8. M2 + hộ có máy tính (2019–2024)")
rows += pooled(d[d.nam >= 2018], "NNL_HT_dtf", ["trained_labor_pct", "urban_share_pct", "ln_grdp"], "A9. M5 bỏ hộ có máy tính (2018–2024)")
ba = pd.DataFrame(rows); save(ba, "phu_A_do_vung_RQ3.csv")

# =====================================================================
# 9. M6 – CẮT NGANG 2024: KỸ NĂNG CNTT (H5e)
# =====================================================================
c24 = d[d.nam == 2024].dropna(subset=["it_skills_pct"])
m6a = smf.ols("EBI_dtf ~ it_skills_pct", data=c24).fit(cov_type="HC1")
m6b = smf.ols("EBI_dtf ~ it_skills_pct + ln_grdp + dn_per_1000 + urban_share_pct", data=c24).fit(cov_type="HC1")
m6c = smf.ols("EBI_dtf ~ it_skills_pct + ln_grdp + urban_share_pct + trained_labor_pct", data=c24).fit(cov_type="HC1")
rows = []
for m, name in [(m6a, "M6a"), (m6b, "M6b"), (m6c, "M6c")]:
    for k in m.params.index:
        if k != "Intercept":
            rows.append(dict(model=name, term=k, coef=m.params[k], se=m.bse[k], p=m.pvalues[k]))
    rows += [dict(model=name, term="N", coef=int(m.nobs)), dict(model=name, term="R2", coef=m.rsquared),
             dict(model=name, term="R2_adj", coef=m.rsquared_adj)]
b7 = pd.DataFrame(rows); save(b7, "bang7_cat_ngang_2024.csv")
save(c24[["tinh", "vung6", "EBI", "EBI_dtf", "it_skills_pct", "internet_user_pct"]], "phu_hinh_ky_nang.csv")
KQ["corr_2024"] = c24[["EBI_dtf", "it_skills_pct", "internet_user_pct", "hh_internet_pct", "social_media_pct",
                       "hh_computer_pct", "dn_per_1000"]].corr()["EBI_dtf"].round(3).to_dict()

# hệ số phóng đại phương sai (VIF) của M2 – ghi chú dưới bảng
from statsmodels.stats.outliers_influence import variance_inflation_factor
dv = d[d.nam >= 2018].dropna(subset=X2); Xv = sm.add_constant(dv[X2])
KQ["vif_M2"] = {x: float(variance_inflation_factor(Xv.values, i + 1)) for i, x in enumerate(X2)}
KQ["tuong_quan_X"] = d[d.nam >= 2018][X_ALL].corr().round(3).to_dict()

# độ lệch chuẩn (giữa các quan sát) của biến X – để đọc độ lớn hệ số
KQ["sd_X"] = d[d.nam >= 2018][X_ALL].std().round(3).to_dict()

with open(os.path.join(OUT, "ket_qua.json"), "w", encoding="utf-8") as f:
    json.dump(KQ, f, ensure_ascii=False, indent=1, default=lambda o: float(o) if isinstance(o, (np.floating, np.integer)) else str(o))

pd.set_option("display.width", 220)
print(b2[["nam", "N", "mean", "CV", "Gini", "Theil", "within_share", "DTF_mean", "max_min", "p90_p10", "Gini_ex2", "within_share_ex2", "within_share_bal"]].round(3).to_string())
print(b3.round(3).to_string())
print(bb[["mo_hinh", "nam_tang_truong", "N", "N_tinh", "beta", "se", "p", "R2"]].round(4).to_string())
print(br[["mo_hinh", "nam_tang_truong", "N", "beta", "se", "p", "R2"]].round(4).to_string())
print(b6.pivot_table(index="term", columns="model", values="coef", aggfunc="first").to_string())
print(b6[b6.p.notna()].pivot(index="term", columns="model", values="p").round(3).to_string())
print(ba[ba.p.notna()].assign(s=lambda x: x.coef.round(3).astype(str) + " (" + x.p.round(3).astype(str) + ")").pivot(index="term", columns="model", values="s").to_string())
print(b7.round(4).to_string())
print({k: (round(v["r"], 3), round(v["p"], 4)) for k, v in pv.items()})
print("XONG –", OUT)

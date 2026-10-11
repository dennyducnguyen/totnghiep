# -*- coding: utf-8 -*-
"""
Phân tích định lượng cho bài báo: Khoảng cách thương mại điện tử (EBI) giữa 63 tỉnh Việt Nam 2017–2024.
Đầu vào : detai/EBI_NGTK_Panel_2017_2024.xlsx  + một số bảng bổ sung từ data/csv/tonghop/
Đầu ra  : detai/ket-qua/  (bang_*.csv, ket_qua.json, panel_phan_tich.csv); hình vẽ: xem detai/hinh-bai-bao/tao_hinh.py

RQ1 (H1–H2): Bảng 2 trung bình vùng; Bảng 3 phân rã Theil; Bảng 4 Gini trụ cột (Lerman–Yitzhaki)
RQ2 (H3)   : Bảng 5 sigma (CV) + beta tuyệt đối / có điều kiện
RQ3 (H4–H5): Bảng 6 OLS gộp + VIF; Bảng 7 biến giả vùng + FE tỉnh; Bảng 8 Oaxaca–Blinder
"""
import json, os, sys, warnings
import numpy as np, pandas as pd
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.stats.outliers_influence import variance_inflation_factor

warnings.filterwarnings("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))          # detai/
PROJ = ROOT
while not os.path.isdir(os.path.join(PROJ, "data", "csv", "tonghop")) and os.path.dirname(PROJ) != PROJ:
    PROJ = os.path.dirname(PROJ)                                              # đi ngược lên tới thư mục thongke/
OUT = os.path.join(ROOT, "ket-qua"); os.makedirs(OUT, exist_ok=True)
TONGHOP = os.path.join(PROJ, "data", "csv", "tonghop")

def tim_file(ten):
    """Tìm file dữ liệu gốc: du-lieu-dau-vao/ của phiên bản này -> thư mục phiên bản -> thư mục cha (detai/)."""
    for d in (os.path.join(ROOT, "du-lieu-dau-vao"), ROOT, os.path.dirname(ROOT)):
        p = os.path.join(d, ten)
        if os.path.exists(p): return p
    raise FileNotFoundError(ten)


REG_ORDER = ["ĐBSH", "TDMNPB", "BTB&DHMT", "TN", "ĐNB", "ĐBSCL"]
REG_FULL = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
KQ = {}   # số liệu chính để viết bài

# =====================================================================
# 1. DỮ LIỆU
# =====================================================================
df = pd.read_excel(tim_file("EBI_NGTK_Panel_2017_2024.xlsx"), sheet_name="Panel_2017_2024")
dig = pd.read_excel(tim_file("EBI_NGTK_Panel_2017_2024.xlsx"), sheet_name="Digital_2024_cross")
# v3: bỏ cột vung_ebi (7 nhãn vùng của VECOM); toàn bài chỉ dùng vung6 (6 vùng theo Niên giám, BTB&DHMT gộp)
df = df.drop(columns=["vung_ebi"], errors="ignore")

def load_tonghop(fname, newname, years=range(2017, 2025)):
    t = pd.read_csv(os.path.join(TONGHOP, fname), encoding="utf-8-sig")
    t = t[t["loai"] == "TINH"].rename(columns={"chi_muc": "tinh"})
    cols = [c for c in t.columns if c.isdigit() and int(c) in years]
    long = t.melt(id_vars="tinh", value_vars=cols, var_name="nam", value_name=newname)
    long["nam"] = long["nam"].astype(int)
    long[newname] = pd.to_numeric(long[newname].replace("..", np.nan), errors="coerce")
    return long

extra = [("khoi-luong-hang-hoa-van-chuyen-bang-duong-bo-phan-theo-dia-p_tu-2010-2024.csv", "freight_road_kt"),
         ("so-luong-sieu-thi-co-den-31-12-hang-nam-phan-theo-dia-phuong_tu-2010-2024.csv", "sieuthi"),
         ("so-luong-trung-tam-thuong-mai-co-den-31-12-hang-nam-phan-the_tu-2010-2024.csv", "tttm"),
         ("so-doanh-nghiep-dang-ky-thanh-lap-moi-phan-theo-dia-phuong_tu-2016-2024.csv", "dn_moi")]
for f, n in extra:
    df = df.merge(load_tonghop(f, n), on=["tinh", "nam"], how="left")
    miss = df[df[n].isna()]["tinh"].unique()
    if len(miss) > 8: print("CẢNH BÁO tên tỉnh không khớp cho", n, ":", miss[:10])

# biến dẫn xuất
df["freight_pc"] = df["freight_road_kt"] / df["pop_th"]                  # tấn/người (nghìn tấn / nghìn người)
df[["sieuthi", "tttm"]] = df[["sieuthi", "tttm"]].fillna(0)                   # ô trống trong NGTK = không có cơ sở
df["modern_retail_100k"] = (df["sieuthi"] + df["tttm"]) / df["pop_th"] * 100   # cơ sở / 100.000 dân
df["dn_moi_1000"] = df["dn_moi"] / df["pop_th"]                           # DN mới / 1.000 dân
df["ln_ebi"] = np.log(df["EBI"])
df["ln_grdp"] = np.log(df["grdp_pc_tr"])
df["ln_income"] = np.log(df["income_month_th"])
df["ln_retail_pc"] = np.log(df["retail_pc_tr"])
df["ln_freight"] = np.log(df["freight_pc"])
df["ln_fdi_pc"] = np.log1p(df["fdi_pc_usd"])
df["ln_dn"] = np.log(df["dn_per_1000"])
# EBI tương đối (trung bình các tỉnh có số liệu trong năm = 100) để so sánh xuyên điểm gãy 2021
df["EBI_rel"] = df["EBI"] / df.groupby("nam")["EBI"].transform("mean") * 100
for p in ["NNL_HT", "B2C", "B2B"]:
    df[p + "_rel"] = df[p] / df.groupby("nam")[p].transform("mean") * 100
df["core"] = df["vung6"].isin(["ĐNB", "ĐBSH"]).astype(int)   # vùng lõi (2 vùng động lực) vs ngoại vi
# biến trễ 1 năm
df = df.sort_values(["tinh", "nam"])
for v in ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct", "hh_computer_pct",
          "ln_freight", "modern_retail_100k", "ln_fdi_pc"]:
    df[v + "_L1"] = df.groupby("tinh")[v].shift(1)
df.to_csv(os.path.join(OUT, "panel_phan_tich.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 2. HÀM BẤT BÌNH ĐẲNG
# =====================================================================
def gini(x):
    x = np.sort(np.asarray(x, float)); n = len(x)
    return (2 * np.sum(np.arange(1, n + 1) * x) / (n * x.sum())) - (n + 1) / n

def theil_T(x):
    x = np.asarray(x, float); m = x.mean(); return np.mean(x / m * np.log(x / m))

def mld(x):
    x = np.asarray(x, float); m = x.mean(); return np.mean(np.log(m / x))

def theil_decomp(x, g):
    """Theil-T phân rã theo nhóm g (mỗi tỉnh trọng số bằng nhau)."""
    x = np.asarray(x, float); g = np.asarray(g); n = len(x); m = x.mean()
    T = theil_T(x); within = 0.0; between = 0.0
    for k in np.unique(g):
        xk = x[g == k]; nk = len(xk); mk = xk.mean(); sk = nk / n * mk / m
        within += sk * theil_T(xk); between += sk * np.log(mk / m)
    return T, between, within

def mld_decomp(x, g):
    x = np.asarray(x, float); g = np.asarray(g); n = len(x); m = x.mean()
    L = mld(x); within = 0.0; between = 0.0
    for k in np.unique(g):
        xk = x[g == k]; nk = len(xk); mk = xk.mean(); pk = nk / n
        within += pk * mld(xk); between += pk * np.log(m / mk)
    return L, between, within

def gini_source_decomp(comp, weights):
    """Lerman–Yitzhaki (1985): G = Σ S_k G_k R_k với Y = Σ w_k X_k."""
    comp = np.asarray(comp, float); Y = comp @ weights; n = len(Y)
    rankY = pd.Series(Y).rank().values / n
    G = gini(Y); rows = []
    for k in range(comp.shape[1]):
        Xk = comp[:, k] * weights[k]
        Sk = Xk.mean() / Y.mean(); Gk = gini(Xk)
        rankk = pd.Series(Xk).rank().values / n
        Rk = np.cov(Xk, rankY)[0, 1] / np.cov(Xk, rankk)[0, 1]
        rows.append(dict(S=Sk, G=gini(comp[:, k]), R=Rk, contrib=Sk * Gk * Rk))
    tot = sum(r["contrib"] for r in rows)
    for r in rows: r["share"] = r["contrib"] / tot
    return G, rows

WEIGHTS = {2017: dict(NNL_HT=.30, B2C=.30, B2B=.30, G2B=.10),
           2018: dict(NNL_HT=.20, B2C=.35, B2B=.35, G2B=.10), 2019: dict(NNL_HT=.20, B2C=.35, B2B=.35, G2B=.10),
           2020: dict(NNL_HT=.20, B2C=.35, B2B=.35, G2B=.10),
           2021: dict(NNL_HT=.20, B2C=.40, B2B=.40), 2022: dict(NNL_HT=.20, B2C=.40, B2B=.40),
           2023: dict(NNL_HT=.20, B2C=.40, B2B=.40), 2024: dict(NNL_HT=.20, B2C=.40, B2B=.40)}

# =====================================================================
# 3. BẢNG 1 – THỐNG KÊ MÔ TẢ
# =====================================================================
VARS = [("EBI", "Chỉ số TMĐT tổng hợp (EBI)", "điểm", "VECOM"),
        ("NNL_HT", "Trụ cột Nguồn nhân lực & Hạ tầng CNTT", "điểm", "VECOM"),
        ("B2C", "Trụ cột Giao dịch B2C", "điểm", "VECOM"),
        ("B2B", "Trụ cột Giao dịch B2B", "điểm", "VECOM"),
        ("G2B", "Trụ cột Giao dịch G2B (2017–2020)", "điểm", "VECOM"),
        ("grdp_pc_tr", "GRDP bình quân đầu người", "triệu đồng", "NGTK"),
        ("dn_per_1000", "Mật độ doanh nghiệp", "DN/1.000 dân", "NGTK"),
        ("urban_share_pct", "Tỷ lệ dân số thành thị", "%", "NGTK"),
        ("trained_labor_pct", "Tỷ lệ lao động 15+ đã qua đào tạo", "%", "NGTK"),
        ("hh_computer_pct", "Tỷ lệ hộ có máy tính (2019–2024)", "%", "NGTK"),
        ("freight_pc", "Khối lượng hàng hóa vận chuyển đường bộ/người", "tấn/người", "NGTK"),
        ("modern_retail_100k", "Siêu thị + trung tâm thương mại/100.000 dân", "cơ sở", "NGTK"),
        ("fdi_pc_usd", "Vốn FDI đăng ký lũy kế/người", "USD/người", "NGTK"),
        ("retail_pc_tr", "Tổng mức bán lẻ HH&DV/người", "triệu đồng", "NGTK"),
        ("income_month_th", "Thu nhập bình quân/người/tháng (2018–2024)", "nghìn đồng", "NGTK"),
        ("density", "Mật độ dân số", "người/km²", "NGTK")]
rows = []
for v, name, unit, src in VARS:
    s = df[v].dropna()
    rows.append(dict(bien=v, ten=name, don_vi=unit, nguon=src, N=len(s), mean=s.mean(), sd=s.std(),
                     min=s.min(), max=s.max()))
b1 = pd.DataFrame(rows); b1.to_csv(os.path.join(OUT, "bang1_mo_ta.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 4. BẢNG 2 – TRUNG BÌNH VÙNG (EBI tương đối, trung bình cả nước = 100)
# =====================================================================
d = df.dropna(subset=["EBI"]).copy()
d["gd"] = np.where(d["nam"] <= 2020, "2017–2020", "2021–2024")
b2a = d.groupby(["vung6", "gd"])["EBI_rel"].mean().unstack("gd")
b2b = d[d.nam == 2024].groupby("vung6")[["EBI_rel", "NNL_HT_rel", "B2C_rel", "B2B_rel"]].mean()
b2c = d.groupby("vung6").agg(n_tinh=("tinh", "nunique"))
b2 = pd.concat([b2c, b2a, b2b], axis=1).loc[REG_ORDER]
b2["ten"] = [REG_FULL[r] for r in b2.index]
# tỉnh dẫn đầu/cuối 2024
top = d[d.nam == 2024].nlargest(5, "EBI")[["tinh", "EBI"]]; bot = d[d.nam == 2024].nsmallest(5, "EBI")[["tinh", "EBI"]]
KQ["top5_2024"] = top.values.tolist(); KQ["bot5_2024"] = bot.values.tolist()
KQ["ratio_max_median"] = {int(y): float(g.EBI.max() / g.EBI.median()) for y, g in d.groupby("nam")}
KQ["ratio_p90_p10"] = {int(y): float(g.EBI.quantile(.9) / g.EBI.quantile(.1)) for y, g in d.groupby("nam")}
b2.to_csv(os.path.join(OUT, "bang2_vung.csv"), encoding="utf-8-sig")
# cho tra cứu: trung bình vùng mọi năm
d.groupby(["nam", "vung6"])["EBI_rel"].mean().unstack().loc[:, REG_ORDER].to_csv(
    os.path.join(OUT, "phu_vung_theo_nam.csv"), encoding="utf-8-sig")
KQ["so_tinh_thieu_EBI"] = {int(y): sorted(df[(df.nam == y) & df.EBI.isna()]["tinh"].tolist()) for y in range(2017, 2025)}

# =====================================================================
# 5. BẢNG 3 – GINI, THEIL, MLD và PHÂN RÃ THEO 6 VÙNG
# =====================================================================
rows = []
for y, g in d.groupby("nam"):
    T, Tb, Tw = theil_decomp(g.EBI, g.vung6); L, Lb, Lw = mld_decomp(g.EBI, g.vung6)
    rows.append(dict(nam=y, N=len(g), Gini=gini(g.EBI), Theil=T, Theil_between=Tb, Theil_within=Tw,
                     Theil_between_share=Tb / T, MLD=L, MLD_between=Lb, MLD_within=Lw, MLD_between_share=Lb / L,
                     CV=g.EBI.std() / g.EBI.mean(), SD_ln=np.log(g.EBI).std()))
b3 = pd.DataFrame(rows); b3.to_csv(os.path.join(OUT, "bang3_theil.csv"), index=False, encoding="utf-8-sig")
# phân rã within theo từng vùng (đóng góp) năm 2024 & 2020
rows = []
for y in [2020, 2024]:
    g = d[d.nam == y]; n = len(g); m = g.EBI.mean()
    for r in REG_ORDER:
        xk = g[g.vung6 == r].EBI; sk = len(xk) / n * xk.mean() / m
        rows.append(dict(nam=y, vung=r, n=len(xk), mean_rel=xk.mean() / m * 100, Theil_k=theil_T(xk),
                         dong_gop_within=sk * theil_T(xk)))
pd.DataFrame(rows).to_csv(os.path.join(OUT, "phu_theil_vung.csv"), index=False, encoding="utf-8-sig")
# kiểm tra vững: bỏ Hà Nội & TP.HCM
rows = []
for y, g in d.groupby("nam"):
    g2 = g[~g.tinh.isin(["Hà Nội", "TP. Hồ Chí Minh"])]
    T, Tb, Tw = theil_decomp(g2.EBI, g2.vung6)
    rows.append(dict(nam=y, Gini_ex2=gini(g2.EBI), Theil_ex2=T, between_share_ex2=Tb / T, CV_ex2=g2.EBI.std() / g2.EBI.mean()))
pd.DataFrame(rows).to_csv(os.path.join(OUT, "phu_theil_bo_HN_HCM.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 6. BẢNG 4 – GINI TRỤ CỘT + PHÂN RÃ LERMAN–YITZHAKI
# =====================================================================
rows = []
for y, g in d.groupby("nam"):
    w = WEIGHTS[y]; cols = list(w.keys()); gg = g.dropna(subset=cols)
    G, parts = gini_source_decomp(gg[cols].values, np.array([w[c] for c in cols]))
    r = dict(nam=y, N=len(gg), Gini_EBI=G)
    for c, p in zip(cols, parts):
        r[f"Gini_{c}"] = p["G"]; r[f"S_{c}"] = p["S"]; r[f"R_{c}"] = p["R"]; r[f"share_{c}"] = p["share"]
    rows.append(r)
b4 = pd.DataFrame(rows); b4.to_csv(os.path.join(OUT, "bang4_gini_tru_cot.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 7. BẢNG 5 – HỘI TỤ SIGMA VÀ BETA
# =====================================================================
# sigma: CV theo năm cho EBI và trụ cột (đã có CV EBI ở b3)
sig = d.groupby("nam").agg(CV_EBI=("EBI", lambda x: x.std() / x.mean()),
                           CV_NNL_HT=("NNL_HT", lambda x: x.std() / x.mean()),
                           CV_B2C=("B2C", lambda x: x.std() / x.mean()),
                           CV_B2B=("B2B", lambda x: x.std() / x.mean()),
                           SDln_EBI=("EBI", lambda x: np.log(x).std()))
sig.to_csv(os.path.join(OUT, "bang5a_sigma.csv"), encoding="utf-8-sig")

def beta_conv(dd, y0, y1, ycol="EBI", cond=None, cond_year=None):
    a = dd[dd.nam == y0].set_index("tinh")[[ycol]].rename(columns={ycol: "y0"})
    b = dd[dd.nam == y1].set_index("tinh")[[ycol]].rename(columns={ycol: "y1"})
    x = a.join(b, how="inner").dropna(); T = y1 - y0
    x["growth"] = np.log(x.y1 / x.y0) / T; x["ln_y0"] = np.log(x.y0)
    rhs = ["ln_y0"]
    if cond:
        cy = cond_year or y0
        z = dd[dd.nam == cy].set_index("tinh")[cond]
        x = x.join(z, how="inner").dropna(); rhs += cond
    X = sm.add_constant(x[rhs]); m = sm.OLS(x["growth"], X).fit(cov_type="HC1")
    b_ = m.params["ln_y0"]; lam = -np.log(1 + b_ * T) / T if (1 + b_ * T) > 0 else np.nan
    out = dict(period=f"{y0}–{y1}", N=int(m.nobs), beta=b_, se=m.bse["ln_y0"], p=m.pvalues["ln_y0"],
               lam=lam, half_life=(np.log(2) / lam if lam and lam > 0 else np.nan), R2=m.rsquared,
               const=m.params["const"])
    for c in (cond or []):
        out[f"b_{c}"] = m.params[c]; out[f"p_{c}"] = m.pvalues[c]
    return out, x, m

COND = ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct"]
rows = []
# tuyệt đối
for (y0, y1, col) in [(2017, 2020, "EBI"), (2021, 2024, "EBI"), (2017, 2024, "EBI_rel"), (2017, 2020, "EBI_rel"), (2021, 2024, "EBI_rel")]:
    o, _, _ = beta_conv(d, y0, y1, col); o["model"] = "Tuyệt đối"; o["ycol"] = col; rows.append(o)
# có điều kiện (ln_grdp chỉ có từ 2018 -> dùng giá trị 2018 cho gốc 2017)
for (y0, y1, col) in [(2017, 2020, "EBI"), (2021, 2024, "EBI"), (2017, 2024, "EBI_rel")]:
    o, _, _ = beta_conv(d, y0, y1, col, cond=COND, cond_year=(2018 if y0 == 2017 else y0))
    o["model"] = "Có điều kiện"; o["ycol"] = col; rows.append(o)
b5 = pd.DataFrame(rows); b5.to_csv(os.path.join(OUT, "bang5b_beta.csv"), index=False, encoding="utf-8-sig")
# dữ liệu cho hình beta
_, xb1, _ = beta_conv(d, 2017, 2020); xb1.to_csv(os.path.join(OUT, "phu_beta_2017_2020.csv"), encoding="utf-8-sig")
_, xb2, _ = beta_conv(d, 2021, 2024); xb2.to_csv(os.path.join(OUT, "phu_beta_2021_2024.csv"), encoding="utf-8-sig")
_, xb3, _ = beta_conv(d, 2017, 2024, "EBI_rel"); xb3.to_csv(os.path.join(OUT, "phu_beta_2017_2024_rel.csv"), encoding="utf-8-sig")
# beta dạng panel (tăng trưởng hằng năm, hiệu ứng năm) – kiểm tra vững
pn = d.sort_values(["tinh", "nam"]).copy()
pn["g"] = pn.groupby("tinh")["EBI_rel"].transform(lambda s: np.log(s).diff())
pn["ln_y0"] = pn.groupby("tinh")["EBI_rel"].shift(1).pipe(np.log)
pn = pn.dropna(subset=["g", "ln_y0"])
pn = pn[pn.nam != 2021]   # loại bước nhảy 2020->2021 do đổi phương pháp
mp = smf.ols("g ~ ln_y0 + C(nam)", data=pn).fit(cov_type="cluster", cov_kwds={"groups": pn["tinh"]})
KQ["beta_panel"] = dict(beta=float(mp.params["ln_y0"]), se=float(mp.bse["ln_y0"]), p=float(mp.pvalues["ln_y0"]), N=int(mp.nobs))

# =====================================================================
# 8. BẢNG 6 – OLS GỘP + VIF
# =====================================================================
def fit_ols(formula, data, cluster="tinh"):
    m = smf.ols(formula, data=data).fit(cov_type="cluster", cov_kwds={"groups": data[cluster]})
    return m

def tidy(m, name):
    rows = []
    for k in m.params.index:
        rows.append(dict(model=name, term=k, coef=m.params[k], se=m.bse[k], p=m.pvalues[k]))
    rows.append(dict(model=name, term="N", coef=m.nobs, se=np.nan, p=np.nan))
    rows.append(dict(model=name, term="R2", coef=m.rsquared, se=np.nan, p=np.nan))
    rows.append(dict(model=name, term="R2_adj", coef=m.rsquared_adj, se=np.nan, p=np.nan))
    return rows

X1 = ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct", "ln_freight", "modern_retail_100k", "ln_fdi_pc"]
X2 = X1 + ["hh_computer_pct"]
d6a = d.dropna(subset=X1 + ["ln_ebi"]).copy()                    # 2018–2024
d6b = d.dropna(subset=X2 + ["ln_ebi"]).copy()                    # 2019–2024
d6c = d.dropna(subset=[v + "_L1" for v in X2] + ["ln_ebi"]).copy()  # trễ 1 năm: 2020–2024
F1 = "ln_ebi ~ " + " + ".join(X1) + " + C(nam)"
F2 = "ln_ebi ~ " + " + ".join(X2) + " + C(nam)"
F3 = "ln_ebi ~ " + " + ".join(v + "_L1" for v in X2) + " + C(nam)"
m1 = fit_ols(F1, d6a); m2 = fit_ols(F2, d6b); m3 = fit_ols(F3, d6c)
# chuẩn hoá (beta chuẩn hoá) cho mô hình 2
d6b_z = d6b.copy()
for v in X2 + ["ln_ebi"]:
    d6b_z[v] = (d6b_z[v] - d6b_z[v].mean()) / d6b_z[v].std()
m2z = fit_ols(F2, d6b_z)
# mô hình chỉ giai đoạn 2021–2024 (thang mới)
d6d = d6b[d6b.nam >= 2021]; m4 = fit_ols(F2, d6d)
rows = tidy(m1, "M1: 2018–2024") + tidy(m2, "M2: 2019–2024 (+máy tính)") + tidy(m3, "M3: X trễ 1 năm") + \
       tidy(m4, "M4: 2021–2024") + tidy(m2z, "M2 chuẩn hóa")
b6 = pd.DataFrame(rows); b6.to_csv(os.path.join(OUT, "bang6_ols.csv"), index=False, encoding="utf-8-sig")
# VIF cho mô hình 2 (không tính biến giả năm)
Xv = sm.add_constant(d6b[X2]); vif = pd.DataFrame({"bien": X2, "VIF": [variance_inflation_factor(Xv.values, i + 1) for i in range(len(X2))]})
vif.to_csv(os.path.join(OUT, "bang6_vif.csv"), index=False, encoding="utf-8-sig")
# ma trận tương quan
d6b[["ln_ebi"] + X2].corr().to_csv(os.path.join(OUT, "phu_tuong_quan.csv"), encoding="utf-8-sig")

# 2024 cắt ngang với biến số chi tiết
d24 = d[d.nam == 2024].merge(dig[["tinh", "internet_user_pct", "hh_internet_pct", "mobile_user_pct", "social_media_pct", "it_skills_pct"]], on="tinh")
mc1 = smf.ols("ln_ebi ~ internet_user_pct + it_skills_pct + ln_grdp + dn_per_1000 + urban_share_pct", data=d24).fit(cov_type="HC1")
mc2 = smf.ols("ln_ebi ~ internet_user_pct + it_skills_pct", data=d24).fit(cov_type="HC1")
mc3 = smf.ols("ln_ebi ~ hh_computer_pct + ln_grdp + dn_per_1000 + urban_share_pct + trained_labor_pct", data=d24).fit(cov_type="HC1")
pd.DataFrame(tidy(mc2, "C1: chỉ biến số") + tidy(mc1, "C2: + kinh tế") + tidy(mc3, "C3: máy tính + kinh tế")).to_csv(
    os.path.join(OUT, "bang6b_2024_cross.csv"), index=False, encoding="utf-8-sig")
KQ["corr_2024"] = d24[["EBI", "internet_user_pct", "it_skills_pct", "hh_computer_pct", "hh_internet_pct"]].corr()["EBI"].round(3).to_dict()

# =====================================================================
# 9. BẢNG 7 – BIẾN GIẢ VÙNG + FE TỈNH
# =====================================================================
d6b["vung6"] = pd.Categorical(d6b["vung6"], categories=["ĐNB"] + [r for r in REG_ORDER if r != "ĐNB"])
F7a = "ln_ebi ~ C(vung6) + C(nam)"
F7b = F2 + " + C(vung6)"
m7a = fit_ols(F7a, d6b); m7b = fit_ols(F7b, d6b)
m7c = fit_ols(F2 + " + C(tinh)", d6b)   # FE tỉnh
rows = tidy(m7a, "M5: chỉ vùng") + tidy(m7b, "M6: X + vùng") + tidy(m7c, "M7: FE tỉnh")
b7 = pd.DataFrame(rows); b7 = b7[~b7.term.str.startswith("C(tinh)")]
b7.to_csv(os.path.join(OUT, "bang7_vung.csv"), index=False, encoding="utf-8-sig")
# kiểm định F: nhóm biến giả vùng trong M6
from statsmodels.formula.api import ols as _ols
r_wald = m7b.wald_test(" = ".join([f"C(vung6)[T.{r}]" for r in REG_ORDER if r != "ĐNB"]) + " = 0", scalar=True) if False else None
try:
    hyp = ", ".join([f"C(vung6)[T.{r}] = 0" for r in REG_ORDER if r != "ĐNB"])
    w = m7b.wald_test(hyp, scalar=True); KQ["wald_vung_M6"] = dict(stat=float(w.statistic), p=float(w.pvalue))
except Exception as e:
    KQ["wald_vung_M6"] = str(e)

# =====================================================================
# 10. BẢNG 8 – OAXACA–BLINDER (vùng lõi ĐNB+ĐBSH vs ngoại vi)
# =====================================================================
def oaxaca(data, y, X, group, ref="pooled", year_dummies=True):
    A = data[data[group] == 1]; B = data[data[group] == 0]
    def design(dd):
        Xm = dd[X].copy()
        if year_dummies:
            Xm = pd.concat([Xm, pd.get_dummies(dd["nam"], prefix="y", drop_first=True).astype(float)], axis=1)
        return sm.add_constant(Xm)
    XA, XB = design(A), design(B)
    XA, XB = XA.align(XB, join="outer", axis=1, fill_value=0.0)
    bA = sm.OLS(A[y], XA).fit().params; bB = sm.OLS(B[y], XB).fit().params
    if ref == "pooled":
        P = pd.concat([A, B]); XP = design(P); XP, _ = XP.align(XA, join="outer", axis=1, fill_value=0.0)
        XP["grp"] = P[group].values          # Jann (2008): thêm biến nhóm để tránh thiên lệch
        bR = sm.OLS(P[y], XP).fit().params.drop("grp")
    elif ref == "A": bR = bA
    else: bR = bB
    xa, xb = XA.mean(), XB.mean()
    gap = A[y].mean() - B[y].mean()
    endow = ((xa - xb) * bR)           # theo từng biến
    coef = (xa * (bA - bR) + xb * (bR - bB))
    return dict(gap=gap, endow=endow.sum(), coef=coef.sum(), endow_detail=endow, coef_detail=coef, nA=len(A), nB=len(B),
                yA=A[y].mean(), yB=B[y].mean())

def oaxaca_boot(data, y, X, group, ref, B=500, seed=1):
    rng = np.random.default_rng(seed); prov = data["tinh"].unique(); res = []
    for _ in range(B):
        samp = rng.choice(prov, size=len(prov), replace=True)
        boot = pd.concat([data[data.tinh == p] for p in samp])
        try:
            o = oaxaca(boot, y, X, group, ref); res.append((o["endow"], o["coef"]))
        except Exception: pass
    res = np.array(res); return res.std(axis=0)

rows = []
for name, dd, X in [("2019–2024 (M2)", d6b, X2), ("2018–2024 (M1)", d6a, X1), ("2021–2024 (M4)", d6d, X2)]:
    for ref in ["pooled", "A"]:
        o = oaxaca(dd, "ln_ebi", X, "core", ref)
        se = oaxaca_boot(dd, "ln_ebi", X, "core", ref, B=300) if (name == "2019–2024 (M2)") else (np.nan, np.nan)
        rows.append(dict(sample=name, ref=ref, nA=o["nA"], nB=o["nB"], yA=o["yA"], yB=o["yB"], gap=o["gap"],
                         endow=o["endow"], endow_se=se[0], coef=o["coef"], coef_se=se[1],
                         endow_share=o["endow"] / o["gap"], coef_share=o["coef"] / o["gap"]))
        if name == "2019–2024 (M2)" and ref == "pooled":
            det = pd.DataFrame({"endow": o["endow_detail"], "coef": o["coef_detail"]})
            det["endow_share"] = det["endow"] / o["gap"]; det.to_csv(os.path.join(OUT, "bang8_oaxaca_chi_tiet.csv"), encoding="utf-8-sig")
b8 = pd.DataFrame(rows); b8.to_csv(os.path.join(OUT, "bang8_oaxaca.csv"), index=False, encoding="utf-8-sig")
# khoảng cách thô giữa lõi và ngoại vi theo năm (EBI_rel)
KQ["core_vs_periphery_rel"] = d.groupby(["nam", "core"])["EBI_rel"].mean().unstack().round(1).to_dict()

# (Phần vẽ hình đã chuyển sang detai/hinh-bai-bao/tao_hinh.py — đọc ket-qua/panel_phan_tich.csv)

# =====================================================================
# 12. LƯU KẾT QUẢ CHÍNH
# =====================================================================
KQ["bang3"] = b3.round(4).to_dict("records"); KQ["bang5b"] = b5.round(4).to_dict("records")
KQ["bang8"] = b8.round(4).to_dict("records"); KQ["vif"] = vif.round(2).to_dict("records")
KQ["N_M1"], KQ["N_M2"], KQ["N_M3"] = int(m1.nobs), int(m2.nobs), int(m3.nobs)
with open(os.path.join(OUT, "ket_qua.json"), "w", encoding="utf-8") as f:
    json.dump(KQ, f, ensure_ascii=False, indent=1, default=float)
print("XONG. Kết quả tại", OUT)
print(b3.round(3).to_string()); print(b5.round(4).to_string()); print(vif.round(2).to_string())
print(b8.round(3).to_string())
print(m2.summary().tables[1])

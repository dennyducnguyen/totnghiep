# -*- coding: utf-8 -*-
"""
PHÂN TÍCH ĐỊNH LƯỢNG CHO BÀI BÁO BẢN v7 (thiết kế tinh gọn sau phản biện 14/09/2026)
"Khoảng cách thương mại điện tử giữa các tỉnh, thành phố Việt Nam giai đoạn 2017–2024:
 mức độ, cấu trúc và các yếu tố liên quan"

Đầu vào : du-lieu-dau-vao/EBI_NGTK_Panel_2017_2024.xlsx  + 4 bảng bổ sung từ data/csv/tonghop/ (đi ngược lên thư mục thongke/)
Đầu ra  : ket-qua-v7/  (bang*.csv, phu_*.csv, ket_qua_v7.json, panel_v7.csv);  bo-sung/ = kết quả CHỈ dùng cho báo cáo chẩn đoán
          (Oaxaca–Blinder, hội tụ beta có điều kiện) – KHÔNG đưa vào bài báo v7.

Thiết kế phương pháp v7 (mỗi câu hỏi một phương pháp):
  RQ1 mức độ & thay đổi   : Gini (chính) + Theil-T; xu hướng trong từng giai đoạn (hội tụ sigma = SD của ln EBI giảm);
                            kiểm tra vững: loại Hà Nội & TP.HCM, mẫu cân bằng 53 tỉnh, MLD (phụ lục).
  RQ2 giữa vùng / trong vùng : phân rã Theil-T theo 6 vùng; đóng góp từng vùng vào phần trong vùng.
  RQ3 yếu tố liên hệ      : hồi quy dữ liệu bảng hiệu ứng cố định tỉnh + năm, SE cụm tỉnh (chính);
                            kiểm tra vững: mẫu 2018–2024, 2021–2024, bỏ mật độ DN, biến phụ thuộc = hạng phần trăm,
                            X trễ 1 năm, OLS gộp 2018–2020 (trước khi EBI chứa chỉ tiêu DN/dân số), Hausman FE–RE.
  Phụ lục                 : Gini từng trụ cột + phân rã Gini theo nguồn; hội tụ beta tuyệt đối theo giai đoạn; cắt ngang 2024.
Kết quả của v6 (ket-qua/) giữ nguyên để đối chiếu; script này không ghi đè.
"""
import json, os, warnings
import numpy as np, pandas as pd
import statsmodels.api as sm
import statsmodels.formula.api as smf
from scipy import stats

warnings.filterwarnings("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))          # v7/
PROJ = ROOT
while not os.path.isdir(os.path.join(PROJ, "data", "csv", "tonghop")) and os.path.dirname(PROJ) != PROJ:
    PROJ = os.path.dirname(PROJ)
OUT = os.path.join(ROOT, "ket-qua-v7"); os.makedirs(OUT, exist_ok=True)
OUT_BS = os.path.join(OUT, "bo-sung"); os.makedirs(OUT_BS, exist_ok=True)
TONGHOP = os.path.join(PROJ, "data", "csv", "tonghop")

def tim_file(ten):
    for d in (os.path.join(ROOT, "du-lieu-dau-vao"), ROOT, os.path.dirname(ROOT)):
        p = os.path.join(d, ten)
        if os.path.exists(p): return p
    raise FileNotFoundError(ten)

REG_ORDER = ["ĐNB", "ĐBSH", "BTB&DHMT", "ĐBSCL", "TN", "TDMNPB"]
REG_FULL = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
KQ = {}

# =====================================================================
# 1. DỮ LIỆU (giống v6, bỏ các biến không dùng)
# =====================================================================
df = pd.read_excel(tim_file("EBI_NGTK_Panel_2017_2024.xlsx"), sheet_name="Panel_2017_2024")
dig = pd.read_excel(tim_file("EBI_NGTK_Panel_2017_2024.xlsx"), sheet_name="Digital_2024_cross")
df = df.drop(columns=["vung_ebi"], errors="ignore")

def load_tonghop(fname, newname, years=range(2017, 2025)):
    t = pd.read_csv(os.path.join(TONGHOP, fname), encoding="utf-8-sig")
    t = t[t["loai"] == "TINH"].rename(columns={"chi_muc": "tinh"})
    cols = [c for c in t.columns if c.isdigit() and int(c) in years]
    long = t.melt(id_vars="tinh", value_vars=cols, var_name="nam", value_name=newname)
    long["nam"] = long["nam"].astype(int)
    long[newname] = pd.to_numeric(long[newname].replace("..", np.nan), errors="coerce")
    return long

for f, n in [("khoi-luong-hang-hoa-van-chuyen-bang-duong-bo-phan-theo-dia-p_tu-2010-2024.csv", "freight_road_kt"),
             ("so-luong-sieu-thi-co-den-31-12-hang-nam-phan-theo-dia-phuong_tu-2010-2024.csv", "sieuthi"),
             ("so-luong-trung-tam-thuong-mai-co-den-31-12-hang-nam-phan-the_tu-2010-2024.csv", "tttm")]:
    df = df.merge(load_tonghop(f, n), on=["tinh", "nam"], how="left")

df["freight_pc"] = df["freight_road_kt"] / df["pop_th"]
df[["sieuthi", "tttm"]] = df[["sieuthi", "tttm"]].fillna(0)
df["modern_retail_100k"] = (df["sieuthi"] + df["tttm"]) / df["pop_th"] * 100
df["ln_ebi"] = np.log(df["EBI"])
df["ln_grdp"] = np.log(df["grdp_pc_tr"])
df["ln_freight"] = np.log(df["freight_pc"])
df["ln_fdi_pc"] = np.log1p(df["fdi_pc_usd"])
df["EBI_rel"] = df["EBI"] / df.groupby("nam")["EBI"].transform("mean") * 100
for p in ["NNL_HT", "B2C", "B2B"]:
    df[p + "_rel"] = df[p] / df.groupby("nam")[p].transform("mean") * 100
# hạng phần trăm trong năm (0–100): bất biến với mọi phép biến đổi đơn điệu của thang đo EBI
df["EBI_pct_rank"] = df.groupby("nam")["EBI"].rank(pct=True) * 100
df["core"] = df["vung6"].isin(["ĐNB", "ĐBSH"]).astype(int)
df = df.sort_values(["tinh", "nam"])
X1 = ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct", "ln_freight", "modern_retail_100k", "ln_fdi_pc"]
X2 = X1 + ["hh_computer_pct"]
for v in X2:
    df[v + "_L1"] = df.groupby("tinh")[v].shift(1)
df.to_csv(os.path.join(OUT, "panel_v7.csv"), index=False, encoding="utf-8-sig")
d = df.dropna(subset=["EBI"]).copy()

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
    x = np.asarray(x, float); g = np.asarray(g); n = len(x); m = x.mean()
    T = theil_T(x); within = between = 0.0
    for k in np.unique(g):
        xk = x[g == k]; sk = len(xk) / n * xk.mean() / m
        within += sk * theil_T(xk); between += sk * np.log(xk.mean() / m)
    return T, between, within
def mld_decomp(x, g):
    x = np.asarray(x, float); g = np.asarray(g); n = len(x); m = x.mean()
    L = mld(x); within = between = 0.0
    for k in np.unique(g):
        xk = x[g == k]; pk = len(xk) / n
        within += pk * mld(xk); between += pk * np.log(m / xk.mean())
    return L, between, within
def gini_source_decomp(comp, weights):
    comp = np.asarray(comp, float); Y = comp @ weights; n = len(Y)
    rankY = pd.Series(Y).rank().values / n; G = gini(Y); rows = []
    for k in range(comp.shape[1]):
        Xk = comp[:, k] * weights[k]; Sk = Xk.mean() / Y.mean(); Gk = gini(Xk)
        rankk = pd.Series(Xk).rank().values / n
        Rk = np.cov(Xk, rankY)[0, 1] / np.cov(Xk, rankk)[0, 1]
        rows.append(dict(S=Sk, G=gini(comp[:, k]), R=Rk, contrib=Sk * Gk * Rk))
    tot = sum(r["contrib"] for r in rows)
    for r in rows: r["share"] = r["contrib"] / tot
    return G, rows
WEIGHTS = {2017: dict(NNL_HT=.30, B2C=.30, B2B=.30, G2B=.10),
           **{y: dict(NNL_HT=.20, B2C=.35, B2B=.35, G2B=.10) for y in (2018, 2019, 2020)},
           **{y: dict(NNL_HT=.20, B2C=.40, B2B=.40) for y in (2021, 2022, 2023, 2024)}}

# =====================================================================
# 3. BẢNG 1 – MÔ TẢ BIẾN (chỉ các biến dùng trong bài v7)
# =====================================================================
VARS = [("EBI", "Chỉ số TMĐT tổng hợp (EBI)", "điểm", "VECOM"),
        ("NNL_HT", "Trụ cột nguồn nhân lực & hạ tầng CNTT", "điểm", "VECOM"),
        ("B2C", "Trụ cột giao dịch B2C", "điểm", "VECOM"),
        ("B2B", "Trụ cột giao dịch B2B", "điểm", "VECOM"),
        ("G2B", "Trụ cột giao dịch G2B (chỉ 2017–2020)", "điểm", "VECOM"),
        ("grdp_pc_tr", "GRDP bình quân đầu người (từ 2018)", "triệu đồng", "NGTK"),
        ("dn_per_1000", "Mật độ doanh nghiệp đang hoạt động", "DN/1.000 dân", "NGTK"),
        ("urban_share_pct", "Tỷ lệ dân số thành thị", "%", "NGTK"),
        ("trained_labor_pct", "Tỷ lệ lao động 15+ đã qua đào tạo", "%", "NGTK"),
        ("hh_computer_pct", "Tỷ lệ hộ có máy tính (từ 2019)", "%", "NGTK"),
        ("freight_pc", "Khối lượng hàng hóa vận chuyển đường bộ/người", "tấn/người", "NGTK"),
        ("modern_retail_100k", "Siêu thị + trung tâm thương mại/100.000 dân", "cơ sở", "NGTK"),
        ("fdi_pc_usd", "Vốn FDI đăng ký lũy kế còn hiệu lực/người", "USD/người", "NGTK")]
rows = []
for v, name, unit, src in VARS:
    s = df[v].dropna()
    rows.append(dict(bien=v, ten=name, don_vi=unit, nguon=src, N=len(s), mean=s.mean(), sd=s.std(), min=s.min(), max=s.max()))
pd.DataFrame(rows).to_csv(os.path.join(OUT, "bang1_mo_ta.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 4. BẢNG 2 – ĐỘ PHỦ VÀ KHẢ NĂNG SO SÁNH THEO NĂM
# =====================================================================
piv = d.pivot(index="tinh", columns="nam", values="EBI")
rows = []
for y in range(2017, 2025):
    g = d[d.nam == y]["EBI"]
    rho_prev = np.nan
    if y > 2017:
        s = piv[[y - 1, y]].dropna(); rho_prev = stats.spearmanr(s[y - 1], s[y]).correlation
    rows.append(dict(nam=y, N=len(g), mean=g.mean(), median=g.median(), min=g.min(), max=g.max(),
                     ratio_max_median=g.max() / g.median(), ratio_p90_p10=g.quantile(.9) / g.quantile(.1),
                     spearman_prev=rho_prev, thieu=", ".join(sorted(df[(df.nam == y) & df.EBI.isna()]["tinh"]))))
b2 = pd.DataFrame(rows); b2.to_csv(os.path.join(OUT, "bang2_do_phu.csv"), index=False, encoding="utf-8-sig")
KQ["so_tinh_thieu_EBI"] = {int(y): sorted(df[(df.nam == y) & df.EBI.isna()]["tinh"].tolist()) for y in range(2017, 2025)}
KQ["tinh_thieu_vung"] = df[df.EBI.isna()].groupby("vung6")["tinh"].apply(lambda s: sorted(set(s))).to_dict()
cnt = d.groupby("tinh").size(); KQ["tinh_du_8_nam"] = int((cnt == 8).sum())
KQ["tinh_khong_du_8_nam"] = sorted(cnt[cnt < 8].index.tolist())
KQ["N_tinh_nam"] = int(len(d))
sp = {}
for a, b in [(2017, 2020), (2020, 2021), (2021, 2024), (2017, 2024)]:
    s = piv[[a, b]].dropna(); sp[f"{a}-{b}"] = dict(N=len(s), rho=float(stats.spearmanr(s[a], s[b]).correlation))
KQ["spearman"] = sp
# trọng số trụ cột phục hồi bằng hồi quy không hệ số chặn (đối chiếu)
w_rows = []
for y, g in d.groupby("nam"):
    cols = list(WEIGHTS[y].keys()); gg = g.dropna(subset=cols)
    b = sm.OLS(gg["EBI"], gg[cols]).fit()
    w_rows.append(dict(nam=y, **{c: round(float(b.params[c]), 3) for c in cols}, R2=round(float(b.rsquared), 4)))
pd.DataFrame(w_rows).to_csv(os.path.join(OUT, "phu_trong_so_tru_cot.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 5. BẢNG 3 – EBI TƯƠNG ĐỐI THEO VÙNG (giữ như v6)
# =====================================================================
d["gd"] = np.where(d["nam"] <= 2020, "2017–2020", "2021–2024")
b3a = d.groupby(["vung6", "gd"])["EBI_rel"].mean().unstack("gd")
b3b = d[d.nam == 2024].groupby("vung6")[["EBI_rel", "NNL_HT_rel", "B2C_rel", "B2B_rel"]].mean()
b3c = d.groupby("vung6").agg(n_tinh=("tinh", "nunique"))
b3 = pd.concat([b3c, b3a, b3b], axis=1).loc[REG_ORDER]; b3["ten"] = [REG_FULL[r] for r in b3.index]
b3.to_csv(os.path.join(OUT, "bang3_vung.csv"), encoding="utf-8-sig")
d.groupby(["nam", "vung6"])["EBI_rel"].mean().unstack().loc[:, REG_ORDER].to_csv(os.path.join(OUT, "phu_vung_theo_nam.csv"), encoding="utf-8-sig")
top = d[d.nam == 2024].nlargest(5, "EBI")[["tinh", "EBI"]]; bot = d[d.nam == 2024].nsmallest(5, "EBI")[["tinh", "EBI"]]
KQ["top5_2024"] = top.values.tolist(); KQ["bot5_2024"] = bot.values.tolist()
KQ["core_vs_periphery_rel"] = d.groupby(["nam", "core"])["EBI_rel"].mean().unstack().round(1).to_dict()

# =====================================================================
# 6. BẢNG 4 – GINI, THEIL VÀ PHÂN RÃ THEO 6 VÙNG (+ loại HN & TP.HCM; + mẫu cân bằng; + MLD)
# =====================================================================
rows = []
bal = d[d.tinh.isin(cnt[cnt == 8].index)]
for y, g in d.groupby("nam"):
    T, Tb, Tw = theil_decomp(g.EBI, g.vung6); L, Lb, Lw = mld_decomp(g.EBI, g.vung6)
    g2 = g[~g.tinh.isin(["Hà Nội", "TP. Hồ Chí Minh"])]; T2, Tb2, Tw2 = theil_decomp(g2.EBI, g2.vung6)
    gb = bal[bal.nam == y]; T3, Tb3, Tw3 = theil_decomp(gb.EBI, gb.vung6)
    rows.append(dict(nam=y, N=len(g), Gini=gini(g.EBI), Theil=T, Theil_between=Tb, Theil_within=Tw, Theil_between_share=Tb / T,
                     SD_ln=np.log(g.EBI).std(), CV=g.EBI.std() / g.EBI.mean(),
                     MLD=L, MLD_between_share=Lb / L,
                     Gini_ex2=gini(g2.EBI), Theil_ex2=T2, between_share_ex2=Tb2 / T2, SD_ln_ex2=np.log(g2.EBI).std(),
                     N_bal=len(gb), Gini_bal=gini(gb.EBI), Theil_bal=T3, between_share_bal=Tb3 / T3))
b4 = pd.DataFrame(rows); b4.to_csv(os.path.join(OUT, "bang4_gini_theil.csv"), index=False, encoding="utf-8-sig")
KQ["bang4"] = b4.round(4).to_dict("records")
# đóng góp từng vùng vào phần trong vùng (2020 và 2024)
rows = []
for y in [2020, 2024]:
    g = d[d.nam == y]; n = len(g); m = g.EBI.mean()
    T, Tb, Tw = theil_decomp(g.EBI, g.vung6)
    for r in REG_ORDER:
        xk = g[g.vung6 == r].EBI; sk = len(xk) / n * xk.mean() / m
        rows.append(dict(nam=y, vung=r, n=len(xk), mean_rel=xk.mean() / m * 100, s_k=sk, Theil_k=theil_T(xk),
                         dong_gop_within=sk * theil_T(xk), ty_trong_trong_Tw=sk * theil_T(xk) / Tw * 100,
                         ty_trong_trong_T=sk * theil_T(xk) / T * 100))
pd.DataFrame(rows).to_csv(os.path.join(OUT, "bang5_theil_vung.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 7. BẢNG 6 – GINI TỪNG TRỤ CỘT (chính) + PHÂN RÃ GINI THEO NGUỒN (phụ lục)
# =====================================================================
rows = []
for y, g in d.groupby("nam"):
    w = WEIGHTS[y]; cols = list(w.keys()); gg = g.dropna(subset=cols)
    G, parts = gini_source_decomp(gg[cols].values, np.array([w[c] for c in cols]))
    r = dict(nam=y, N=len(gg), Gini_EBI_taitao=G, Gini_EBI=gini(gg.EBI))
    for c, p in zip(cols, parts):
        r[f"Gini_{c}"] = p["G"]; r[f"S_{c}"] = p["S"]; r[f"R_{c}"] = p["R"]; r[f"share_{c}"] = p["share"]
        r[f"CV_{c}"] = gg[c].std() / gg[c].mean()
    rows.append(r)
b6 = pd.DataFrame(rows); b6.to_csv(os.path.join(OUT, "bang6_gini_tru_cot.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 8. PHỤ LỤC – HỘI TỤ BETA TUYỆT ĐỐI THEO GIAI ĐOẠN (không dùng trong bài chính)
# =====================================================================
def beta_conv(dd, y0, y1, ycol="EBI"):
    a = dd[dd.nam == y0].set_index("tinh")[[ycol]].rename(columns={ycol: "y0"})
    b = dd[dd.nam == y1].set_index("tinh")[[ycol]].rename(columns={ycol: "y1"})
    x = a.join(b, how="inner").dropna(); T = y1 - y0
    x["growth"] = np.log(x.y1 / x.y0) / T; x["ln_y0"] = np.log(x.y0)
    m = sm.OLS(x["growth"], sm.add_constant(x[["ln_y0"]])).fit(cov_type="HC1")
    b_ = m.params["ln_y0"]; lam = -np.log(1 + b_ * T) / T if (1 + b_ * T) > 0 else np.nan
    return dict(period=f"{y0}–{y1}", N=int(m.nobs), beta=b_, se=m.bse["ln_y0"], p=m.pvalues["ln_y0"], lam=lam,
                half_life=(np.log(2) / lam if lam and lam > 0 else np.nan), R2=m.rsquared)
rows = [beta_conv(d, 2017, 2020), beta_conv(d, 2021, 2024), beta_conv(d, 2022, 2024)]
# thêm: loại HN & HCM
d_ex = d[~d.tinh.isin(["Hà Nội", "TP. Hồ Chí Minh"])]
for r in [beta_conv(d_ex, 2017, 2020), beta_conv(d_ex, 2021, 2024)]:
    r["period"] += " (loại HN & TP.HCM)"; rows.append(r)
pd.DataFrame(rows).to_csv(os.path.join(OUT, "phu_beta_tuyet_doi.csv"), index=False, encoding="utf-8-sig")

# =====================================================================
# 9. BẢNG 7 – HỒI QUY DỮ LIỆU BẢNG HIỆU ỨNG CỐ ĐỊNH TỈNH + NĂM (RQ3)
# =====================================================================
from linearmodels.panel import PanelOLS, RandomEffects

def fe(data, y, X, entity=True, time=True):
    dd = data.dropna(subset=X + [y]).set_index(["tinh", "nam"])
    m = PanelOLS(dd[y], sm.add_constant(dd[X]) if not entity else dd[X], entity_effects=entity, time_effects=time,
                 drop_absorbed=True).fit(cov_type="clustered", cluster_entity=True)
    return m

def tidy_fe(m, name, X):
    rows = []
    for k in X:
        if k in m.params.index:
            rows.append(dict(model=name, term=k, coef=float(m.params[k]), se=float(m.std_errors[k]), p=float(m.pvalues[k])))
    rows += [dict(model=name, term="N", coef=int(m.nobs), se=np.nan, p=np.nan),
             dict(model=name, term="N_tinh", coef=int(m.entity_info["total"]), se=np.nan, p=np.nan),
             dict(model=name, term="R2_within", coef=float(m.rsquared), se=np.nan, p=np.nan),   # R² của hồi quy sau khi khử hiệu ứng tỉnh và năm
             dict(model=name, term="R2_overall", coef=float(m.rsquared_overall), se=np.nan, p=np.nan)]
    return rows

d7 = d.copy()
d7["ln_ebi_pct"] = d7["EBI_pct_rank"]
specs = [
    ("FE1: 2019–2024", d7[d7.nam >= 2019], "ln_ebi", X2),
    ("FE2: 2018–2024 (không hộ máy tính)", d7[d7.nam >= 2018], "ln_ebi", X1),
    ("FE3: 2021–2024", d7[d7.nam >= 2021], "ln_ebi", X2),
    ("FE4: 2019–2024, bỏ mật độ DN", d7[d7.nam >= 2019], "ln_ebi", [v for v in X2 if v != "dn_per_1000"]),
    ("FE5: 2019–2024, Y = hạng phần trăm", d7[d7.nam >= 2019], "EBI_pct_rank", X2),
    ("FE6: 2019–2024, loại năm 2021", d7[(d7.nam >= 2019) & (d7.nam != 2021)], "ln_ebi", X2),
    ("FE7: X trễ 1 năm (2020–2024)", d7[d7.nam >= 2020], "ln_ebi", [v + "_L1" for v in X2]),
]
rows = []; MODELS = {}
for name, dd, y, X in specs:
    m = fe(dd, y, X); MODELS[name] = m; rows.append(tidy_fe(m, name, X))
rows = [r for rr in rows for r in rr]
# OLS gộp (so sánh) – hiệu ứng năm, SE cụm tỉnh
def pooled(data, y, X, name):
    dd = data.dropna(subset=X + [y])
    m = smf.ols(f"{y} ~ " + " + ".join(X) + " + C(nam)", data=dd).fit(cov_type="cluster", cov_kwds={"groups": dd["tinh"]})
    out = [dict(model=name, term=k, coef=float(m.params[k]), se=float(m.bse[k]), p=float(m.pvalues[k])) for k in X]
    out += [dict(model=name, term="N", coef=int(m.nobs), se=np.nan, p=np.nan),
            dict(model=name, term="N_tinh", coef=int(dd.tinh.nunique()), se=np.nan, p=np.nan),
            dict(model=name, term="R2_overall", coef=float(m.rsquared), se=np.nan, p=np.nan)]
    return out, m
o, m_pool = pooled(d7[d7.nam >= 2019], "ln_ebi", X2, "OLS gộp 2019–2024 (so sánh)"); rows += o
o, m_pre = pooled(d7[(d7.nam >= 2018) & (d7.nam <= 2020)], "ln_ebi", X1, "OLS gộp 2018–2020 (trước đổi phương pháp)"); rows += o
b7 = pd.DataFrame(rows); b7.to_csv(os.path.join(OUT, "bang7_fe.csv"), index=False, encoding="utf-8-sig")

# Hausman FE vs RE trên mẫu FE1
dd = d7[d7.nam >= 2019].dropna(subset=X2 + ["ln_ebi"]).set_index(["tinh", "nam"])
Xt = pd.concat([dd[X2], pd.get_dummies(dd.index.get_level_values("nam"), prefix="y", drop_first=True).set_index(dd.index).astype(float)], axis=1)
m_fe = PanelOLS(dd["ln_ebi"], Xt, entity_effects=True, drop_absorbed=True).fit()
m_re = RandomEffects(dd["ln_ebi"], sm.add_constant(Xt)).fit()
bdiff = m_fe.params[X2] - m_re.params[X2]
V = m_fe.cov.loc[X2, X2] - m_re.cov.loc[X2, X2]
H = float(bdiff.T @ np.linalg.pinv(V.values) @ bdiff); pH = float(1 - stats.chi2.cdf(H, len(X2)))
KQ["hausman"] = dict(stat=H, df=len(X2), p=pH)
# độ lệch chuẩn trong nội bộ tỉnh của các biến X (để đọc độ lớn hệ số)
dd1 = d7[d7.nam >= 2019].dropna(subset=X2 + ["ln_ebi"])
w = dd1[X2].sub(dd1.groupby("tinh")[X2].transform("mean"))
KQ["within_sd"] = w.std().round(4).to_dict(); KQ["overall_sd"] = dd1[X2].std().round(4).to_dict()
KQ["within_sd_ln_ebi"] = float((dd1["ln_ebi"] - dd1.groupby("tinh")["ln_ebi"].transform("mean")).std())
dd1[["ln_ebi"] + X2].corr().to_csv(os.path.join(OUT, "phu_tuong_quan.csv"), encoding="utf-8-sig")
# VIF (chỉ để ghi chú)
from statsmodels.stats.outliers_influence import variance_inflation_factor
Xv = sm.add_constant(dd1[X2]); KQ["vif_max"] = float(max(variance_inflation_factor(Xv.values, i + 1) for i in range(len(X2))))

# =====================================================================
# 10. PHỤ LỤC – CẮT NGANG 2024 VỚI CHỈ TIÊU KỸ NĂNG / SỬ DỤNG
# =====================================================================
d24 = d[d.nam == 2024].merge(dig[["tinh", "internet_user_pct", "hh_internet_pct", "it_skills_pct"]], on="tinh")
mc1 = smf.ols("ln_ebi ~ internet_user_pct + it_skills_pct", data=d24).fit(cov_type="HC1")
mc2 = smf.ols("ln_ebi ~ internet_user_pct + it_skills_pct + ln_grdp + dn_per_1000 + urban_share_pct", data=d24).fit(cov_type="HC1")
mc3 = smf.ols("ln_ebi ~ internet_user_pct + it_skills_pct + ln_grdp + urban_share_pct + trained_labor_pct", data=d24).fit(cov_type="HC1")
rows = []
for m, name in [(mc1, "C1"), (mc2, "C2"), (mc3, "C3 (bỏ mật độ DN)")]:
    for k in m.params.index:
        rows.append(dict(model=name, term=k, coef=m.params[k], se=m.bse[k], p=m.pvalues[k]))
    rows += [dict(model=name, term="N", coef=m.nobs), dict(model=name, term="R2_adj", coef=m.rsquared_adj)]
pd.DataFrame(rows).to_csv(os.path.join(OUT, "phu_2024_cat_ngang.csv"), index=False, encoding="utf-8-sig")
KQ["corr_2024"] = d24[["EBI", "internet_user_pct", "it_skills_pct", "hh_computer_pct", "hh_internet_pct", "dn_per_1000"]].corr()["EBI"].round(3).to_dict()

# =====================================================================
# 11. BỔ SUNG CHO BÁO CÁO CHẨN ĐOÁN (KHÔNG đưa vào bài): Oaxaca–Blinder có/không HN & TP.HCM; beta có điều kiện
# =====================================================================
def oaxaca(data, y, X, group, ref="pooled"):
    A = data[data[group] == 1]; B = data[data[group] == 0]
    def design(dd):
        Xm = pd.concat([dd[X], pd.get_dummies(dd["nam"], prefix="y", drop_first=True).astype(float)], axis=1)
        return sm.add_constant(Xm)
    XA, XB = design(A), design(B); XA, XB = XA.align(XB, join="outer", axis=1, fill_value=0.0)
    bA = sm.OLS(A[y], XA).fit().params; bB = sm.OLS(B[y], XB).fit().params
    if ref == "pooled":
        P = pd.concat([A, B]); XP = design(P); XP, _ = XP.align(XA, join="outer", axis=1, fill_value=0.0)
        XP["grp"] = P[group].values; bR = sm.OLS(P[y], XP).fit().params.drop("grp")
    else: bR = bA
    xa, xb = XA.mean(), XB.mean(); gap = A[y].mean() - B[y].mean()
    endow = (xa - xb) * bR; coef = xa * (bA - bR) + xb * (bR - bB)
    return dict(gap=gap, endow=endow.sum(), coef=coef.sum(), endow_share=endow.sum() / gap, nA=len(A), nB=len(B),
                endow_dn_share=float(endow.get("dn_per_1000", np.nan) / gap))
rows = []
d6b = d[d.nam >= 2019].dropna(subset=X2 + ["ln_ebi"])
for name, dd in [("2019–2024 đầy đủ", d6b), ("2019–2024 loại HN & TP.HCM", d6b[~d6b.tinh.isin(["Hà Nội", "TP. Hồ Chí Minh"])]),
                 ("2019–2020 (trước đổi PP)", d6b[d6b.nam <= 2020]), ("2021–2024", d6b[d6b.nam >= 2021])]:
    for ref in ["pooled", "A"]:
        o = oaxaca(dd, "ln_ebi", X2, "core", ref); o.update(sample=name, ref=ref); rows.append(o)
pd.DataFrame(rows).to_csv(os.path.join(OUT_BS, "oaxaca_kiem_tra.csv"), index=False, encoding="utf-8-sig")
KQ["oaxaca_bo_sung"] = pd.DataFrame(rows).round(3).to_dict("records")

# =====================================================================
# 12. LƯU
# =====================================================================
KQ["bang7_models"] = sorted(set(b7.model))
with open(os.path.join(OUT, "ket_qua_v7.json"), "w", encoding="utf-8") as f:
    json.dump(KQ, f, ensure_ascii=False, indent=1, default=float)
print("XONG. Kết quả tại", OUT)
print(b2.round(3).to_string())
print(b4[["nam", "N", "Gini", "Theil", "Theil_between_share", "SD_ln", "Gini_ex2", "between_share_ex2", "Gini_bal", "between_share_bal", "MLD_between_share"]].round(3).to_string())
print(b7[b7.term.isin(X2 + ["N", "R2_within"])].pivot(index="term", columns="model", values="coef").round(4).to_string())
print(b7[b7.term.isin(X2)].pivot(index="term", columns="model", values="p").round(3).to_string())
print("Hausman:", KQ["hausman"]); print("within SD:", KQ["within_sd"]); print("spearman:", KQ["spearman"])
print(pd.DataFrame(KQ["oaxaca_bo_sung"]).to_string())
print(pd.read_csv(os.path.join(OUT, "phu_beta_tuyet_doi.csv")).round(4).to_string())

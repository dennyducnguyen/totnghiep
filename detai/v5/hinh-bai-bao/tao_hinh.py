# -*- coding: utf-8 -*-
"""
TẠO 8 HÌNH CHO BÀI BÁO (phiên bản v2: mỗi hình một nội dung, không ghép tiểu đồ)
"Khoảng cách thương mại điện tử giữa các tỉnh, thành phố Việt Nam giai đoạn 2017–2024:
 mức độ, cấu trúc, hội tụ và các yếu tố quyết định"

CÁCH CHẠY (chọn một):
  1) Nhấp đúp file  chay_tao_hinh.bat
  2) Mở PowerShell/CMD tại thư mục này rồi gõ:  python tao_hinh.py
  3) Chạy riêng một hình:  python tao_hinh.py 3    (chỉ tạo Hình 3; có thể ghi nhiều số: 3 6 7)

ĐẦU VÀO : du_lieu/panel_63_tinh_2017_2024.csv  (bảng 63 tỉnh × 2017–2024, đã ghép EBI của VECOM
          với các chỉ tiêu Niên giám Thống kê; xem HUONG-DAN-TAO-HINH.docx, mục 5)
ĐẦU RA  : thư mục hinh/  →  Hinh1_khung_phan_tich.png, Hinh2_ebi_vung.png, Hinh3_theil_muc.png,
                             Hinh4_theil_ty_trong.png, Hinh5_sigma_cv.png, Hinh6_beta_2017_2020.png,
                             Hinh7_beta_2021_2024.png, Hinh8_oaxaca_dong_gop.png

Thư viện cần có: pandas, numpy, matplotlib, statsmodels  (cài: pip install pandas numpy matplotlib statsmodels)
"""
import os, sys
import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib
matplotlib.use("Agg")                      # vẽ ra file, không cần cửa sổ
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

# ----------------------------------------------------------------------------------------------
# 0. THIẾT LẬP CHUNG (có thể sửa ở đây: màu, font, độ phân giải)
# ----------------------------------------------------------------------------------------------
THU_MUC = os.path.dirname(os.path.abspath(__file__))
FILE_DU_LIEU = os.path.join(THU_MUC, "du_lieu", "panel_63_tinh_2017_2024.csv")
THU_MUC_HINH = os.path.join(THU_MUC, "hinh")
os.makedirs(THU_MUC_HINH, exist_ok=True)

DPI = 200                                   # độ phân giải file PNG
for ten_font in ["Arial", "Segoe UI", "DejaVu Sans"]:   # font có dấu tiếng Việt; lấy font đầu tiên tìm thấy
    if any(f.name == ten_font for f in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = ten_font
        break
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#c3c2b7", "axes.labelcolor": "#0b0b0b",
                     "xtick.color": "#52514e", "ytick.color": "#52514e", "grid.color": "#e1e0d9",
                     "figure.dpi": DPI, "savefig.dpi": DPI, "savefig.bbox": "tight"})

# Bảng màu 6 vùng (thứ tự cố định, không đổi giữa các hình) và kiểu nét (để in đen trắng vẫn phân biệt được)
MAU = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#4a3aa7"]
NET = ["-", "--", "-.", ":", (0, (5, 1)), (0, (3, 1, 1, 1))]
XAM = "#898781"; XAM_NHAT = "#c3c2b7"; MUC = "#0b0b0b"; MUC_NHAT = "#52514e"

VUNG = ["ĐBSH", "TDMNPB", "BTB&DHMT", "TN", "ĐNB", "ĐBSCL"]
TEN_VUNG = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
VUNG_LOI = ["ĐNB", "ĐBSH"]                  # vùng lõi dùng cho Hình 6 (Oaxaca–Blinder)
NAM_DOI_PHUONG_PHAP = 2020.5                # VECOM đổi cách tính EBI từ báo cáo 2021 → vạch dọc trên hình

# ----------------------------------------------------------------------------------------------
# 1. ĐỌC DỮ LIỆU
# ----------------------------------------------------------------------------------------------
def doc_du_lieu():
    df = pd.read_csv(FILE_DU_LIEU, encoding="utf-8-sig")
    d = df.dropna(subset=["EBI"]).copy()                   # chỉ giữ các tỉnh–năm có EBI (445 quan sát)
    if "EBI_rel" not in d.columns:                          # EBI tương đối: trung bình các tỉnh trong năm = 100
        d["EBI_rel"] = d["EBI"] / d.groupby("nam")["EBI"].transform("mean") * 100
    return d

# ----------------------------------------------------------------------------------------------
# 2. CÁC HÀM THỐNG KÊ DÙNG CHO HÌNH
# ----------------------------------------------------------------------------------------------
def theil_T(x):
    x = np.asarray(x, float); m = x.mean()
    return np.mean(x / m * np.log(x / m))

def mld(x):
    x = np.asarray(x, float); m = x.mean()
    return np.mean(np.log(m / x))

def phan_ra_theil(x, nhom):
    """Theil-T = phần trong vùng + phần giữa các vùng (mỗi tỉnh trọng số bằng nhau)."""
    x = np.asarray(x, float); nhom = np.asarray(nhom); n = len(x); m = x.mean()
    T = theil_T(x); trong = 0.0; giua = 0.0
    for k in np.unique(nhom):
        xk = x[nhom == k]; sk = len(xk) / n * xk.mean() / m
        trong += sk * theil_T(xk); giua += sk * np.log(xk.mean() / m)
    return T, giua, trong

def phan_ra_mld(x, nhom):
    x = np.asarray(x, float); nhom = np.asarray(nhom); n = len(x); m = x.mean()
    L = mld(x); trong = 0.0; giua = 0.0
    for k in np.unique(nhom):
        xk = x[nhom == k]; pk = len(xk) / n
        trong += pk * mld(xk); giua += pk * np.log(m / xk.mean())
    return L, giua, trong

def hoi_tu_beta(d, nam_goc, nam_cuoi):
    """Hồi quy (1/T)ln(EBI_T/EBI_0) = a + b ln(EBI_0); trả về bảng số liệu và (a, b, p)."""
    a = d[d.nam == nam_goc].set_index("tinh")["EBI"].rename("y0")
    b = d[d.nam == nam_cuoi].set_index("tinh")["EBI"].rename("y1")
    x = pd.concat([a, b], axis=1).dropna(); T = nam_cuoi - nam_goc
    x["tang_truong"] = np.log(x.y1 / x.y0) / T
    x["ln_y0"] = np.log(x.y0)
    m = sm.OLS(x["tang_truong"], sm.add_constant(x["ln_y0"])).fit(cov_type="HC1")
    return x, m.params["const"], m.params["ln_y0"], m.pvalues["ln_y0"]

def oaxaca_chi_tiet(d, bien_X, nhom="core"):
    """Phân tách Oaxaca–Blinder hai thành phần (hệ số tham chiếu gộp theo Jann, 2008);
    trả về phần nguồn lực theo từng biến, chia cho tổng khoảng cách (tỷ lệ %)."""
    dd = d.dropna(subset=bien_X + ["ln_ebi"]).copy()
    A = dd[dd[nhom] == 1]; B = dd[dd[nhom] == 0]
    def thiet_ke(g):
        X = pd.concat([g[bien_X], pd.get_dummies(g["nam"], prefix="y", drop_first=True).astype(float)], axis=1)
        return sm.add_constant(X)
    XA, XB = thiet_ke(A), thiet_ke(B)
    XA, XB = XA.align(XB, join="outer", axis=1, fill_value=0.0)
    P = pd.concat([A, B]); XP = thiet_ke(P); XP, _ = XP.align(XA, join="outer", axis=1, fill_value=0.0)
    XP["nhom"] = P[nhom].values
    b_goc = sm.OLS(P["ln_ebi"], XP).fit().params.drop("nhom")        # hệ số tham chiếu gộp
    khoang_cach = A["ln_ebi"].mean() - B["ln_ebi"].mean()
    nguon_luc = (XA.mean() - XB.mean()) * b_goc
    return (nguon_luc / khoang_cach * 100)[bien_X]

# ----------------------------------------------------------------------------------------------
# 3. SÁU HÌNH
# ----------------------------------------------------------------------------------------------
def hinh1_khung_phan_tich(d=None):
    """Hình 1 (Mục 2.1): sơ đồ khung phân tích – ba tầng khoảng cách số và các cơ chế giải thích."""
    fig, ax = plt.subplots(figsize=(7.2, 4.2)); ax.set_xlim(0, 100); ax.set_ylim(0, 60); ax.axis("off")
    def hop(x, y, w, h, tieu_de, noi_dung, fc="#eef4fc", ec="#2a78d6"):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.4,rounding_size=1.2", fc=fc, ec=ec, lw=1.2))
        ax.text(x + w / 2, y + h - 2.2, tieu_de, ha="center", va="top", fontsize=8.2, fontweight="bold", color=MUC)
        ax.text(x + w / 2, y + h / 2 - 2.3, noi_dung, ha="center", va="center", fontsize=7, color=MUC_NHAT, linespacing=1.35)
    def mui_ten(x0, y0, x1, y1):
        ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle="-|>", mutation_scale=11, color=MUC_NHAT, lw=1.1))
    hop(2, 40, 28, 16, "Tầng 1 – Tiếp cận vật chất", "Hạ tầng số, thiết bị\n(hộ có máy tính, Internet)")
    hop(36, 40, 28, 16, "Tầng 2 – Kỹ năng và sử dụng", "Nhân lực số, kỹ năng CNTT,\nlao động qua đào tạo")
    hop(70, 40, 28, 16, "Tầng 3 – Kết quả (EBI)", "Giao dịch B2C, B2B thực tế\ncủa doanh nghiệp và người dân", fc="#fdeee7", ec="#eb6834")
    mui_ten(30.4, 48, 35.6, 48); mui_ten(64.4, 48, 69.6, 48)
    hop(2, 6, 44, 22, "Cơ chế phát triển vùng",
        "Cực tăng trưởng → lan tỏa (Williamson)\nKhối lượng tới hạn & khuếch tán (Rogers)\nHiệu quả vs khuếch tán (Wang et al., 2021)", fc="#f4f4f2", ec=XAM)
    hop(54, 6, 44, 22, "Nguồn lực cấu trúc của tỉnh",
        "Quy mô kinh tế (GRDP/người), mật độ doanh nghiệp,\nđô thị hóa, vốn nhân lực, logistics, FDI,\nbán lẻ hiện đại, hạ tầng số", fc="#f4f4f2", ec=XAM)
    mui_ten(24, 28.4, 24, 39.6)
    ax.text(26, 34, "RQ1–RQ2: mức độ, cấu trúc,\nhội tụ σ/β của khoảng cách", ha="left", va="center", fontsize=6.5, color=MUC_NHAT, style="italic")
    mui_ten(76, 28.4, 84, 39.6)
    ax.text(74, 34, "RQ3: yếu tố quyết định,\nphân tách Oaxaca–Blinder", ha="right", va="center", fontsize=6.5, color=MUC_NHAT, style="italic")
    return fig, "Hinh1_khung_phan_tich.png"

def hinh2_ebi_vung(d):
    """Hình 2 (Mục 4.1): EBI tương đối theo 6 vùng, 2017–2024 (trung bình các tỉnh = 100)."""
    pv = d.groupby(["nam", "vung6"])["EBI_rel"].mean().unstack()[VUNG]
    fig, ax = plt.subplots(figsize=(6.5, 3.6))
    for i, v in enumerate(VUNG):
        ax.plot(pv.index, pv[v], color=MAU[i], ls=NET[i], lw=1.8, marker="o", ms=3.5, label=TEN_VUNG[v])
    ax.axhline(100, color=XAM, lw=0.8, ls=":")
    ax.axvline(NAM_DOI_PHUONG_PHAP, color=XAM_NHAT, lw=0.8)
    ax.text(NAM_DOI_PHUONG_PHAP - 0.05, ax.get_ylim()[0] + 2, "Đổi phương pháp EBI (2021) →", fontsize=7, color=MUC_NHAT, va="bottom", ha="right")
    ax.set_ylabel("EBI tương đối (trung bình các tỉnh = 100)"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5)
    ax.legend(fontsize=7, frameon=False, ncol=2, loc="upper left", bbox_to_anchor=(0, -0.18))
    return fig, "Hinh2_ebi_vung.png"

def _bang_theil(d):
    dong = []
    for nam, g in d.groupby("nam"):
        T, Tg, Tt = phan_ra_theil(g.EBI, g.vung6); L, Lg, Lt = phan_ra_mld(g.EBI, g.vung6)
        dong.append(dict(nam=nam, Theil=T, giua=Tg, trong=Tt, ty_trong_T=Tg / T * 100, ty_trong_MLD=Lg / L * 100))
    return pd.DataFrame(dong)

def hinh3_theil_muc(d):
    """Hình 3 (Mục 4.1): chỉ số Theil-T tổng của EBI và hai thành phần giữa các vùng / trong vùng theo năm."""
    b = _bang_theil(d)
    fig, ax = plt.subplots(figsize=(5.6, 3.3))
    ax.bar(b.nam - 0.2, b.giua, width=0.4, color=MAU[0], label="Giữa các vùng")
    ax.bar(b.nam + 0.2, b.trong, width=0.4, color="#9ec5f4", label="Trong vùng")
    ax.plot(b.nam, b.Theil, color=MUC, lw=1.5, marker="o", ms=3, label="Theil tổng")
    ax.axvline(NAM_DOI_PHUONG_PHAP, color=XAM_NHAT, lw=0.8)
    ax.set_ylabel("Chỉ số Theil-T"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5); ax.legend(fontsize=7, frameon=False)
    return fig, "Hinh3_theil_muc.png"

def hinh4_theil_ty_trong(d):
    """Hình 4 (Mục 4.1): tỷ trọng phần chênh lệch giữa các vùng trong tổng bất bình đẳng, theo Theil-T và MLD."""
    b = _bang_theil(d)
    fig, ax = plt.subplots(figsize=(5.6, 3.3))
    ax.plot(b.nam, b.ty_trong_T, color=MAU[0], lw=1.8, marker="o", ms=3.5, label="Theil-T")
    ax.plot(b.nam, b.ty_trong_MLD, color=MAU[1], lw=1.8, ls="--", marker="s", ms=3.5, label="MLD")
    ax.axhline(50, color=XAM, lw=0.8, ls=":"); ax.text(2017, 51.5, "50% (giữa vùng = trong vùng)", fontsize=6.5, color=MUC_NHAT)
    ax.axvline(NAM_DOI_PHUONG_PHAP, color=XAM_NHAT, lw=0.8)
    ax.set_ylim(0, 100); ax.set_ylabel("Tỷ trọng giữa các vùng (%)"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5); ax.legend(fontsize=7, frameon=False)
    return fig, "Hinh4_theil_ty_trong.png"

def hinh5_sigma(d):
    """Hình 5 (Mục 4.3): hội tụ sigma – hệ số biến thiên (CV) của EBI và ba trụ cột theo năm."""
    cv = lambda s: s.std() / s.mean()
    b = d.groupby("nam").agg(EBI=("EBI", cv), NNL_HT=("NNL_HT", cv), B2C=("B2C", cv), B2B=("B2B", cv))
    nhan = [("EBI", "EBI tổng hợp"), ("NNL_HT", "Nguồn nhân lực & hạ tầng"), ("B2C", "Giao dịch B2C"), ("B2B", "Giao dịch B2B")]
    fig, ax = plt.subplots(figsize=(6.2, 3.3))
    for i, (cot, ten) in enumerate(nhan):
        ax.plot(b.index, b[cot], color=MAU[i], ls=NET[i], lw=1.8, marker="o", ms=3.5, label=ten)
    ax.axvline(NAM_DOI_PHUONG_PHAP, color=XAM_NHAT, lw=0.8)
    ax.set_ylabel("Hệ số biến thiên (CV)"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5)
    ax.legend(fontsize=7, frameon=False, ncol=2)
    return fig, "Hinh5_sigma_cv.png"

def _ve_beta(d, y0, y1, tieu_de):
    x, a, b, p = hoi_tu_beta(d, y0, y1)
    fig, ax = plt.subplots(figsize=(5.4, 3.6))
    ax.scatter(x.ln_y0, x.tang_truong * 100, s=16, color=MAU[0], alpha=.8, edgecolor="white", lw=.5)
    xx = np.linspace(x.ln_y0.min(), x.ln_y0.max(), 50)
    ax.plot(xx, (a + b * xx) * 100, color=MAU[1], lw=1.6, label="Đường hồi quy OLS")
    gan_nhan = ["Hà Nội", "TP. Hồ Chí Minh", "Đà Nẵng", "Bình Dương"] + list(x.nlargest(2, "tang_truong").index) + list(x.nsmallest(1, "tang_truong").index)
    for t in gan_nhan:
        if t in x.index:
            lech = {"Hà Nội": (2, -10), "TP. Hồ Chí Minh": (3, 4), "Đà Nẵng": (-3, 4)}.get(t, (3, 2))
            ax.annotate(t, (x.loc[t, "ln_y0"], x.loc[t, "tang_truong"] * 100), fontsize=6.5, color=MUC_NHAT,
                        xytext=lech, textcoords="offset points", ha=("right" if t == "Đà Nẵng" else "left"))
    ax.set_xlabel("ln(EBI năm gốc)"); ax.set_ylabel("Tăng trưởng EBI bình quân năm (%)")
    ax.set_title(tieu_de, fontsize=9, loc="left"); ax.grid(lw=0.4)
    ax.text(0.98, 0.95, f"β = {b:.3f} (p = {p:.3f}), N = {len(x)}", transform=ax.transAxes, ha="right", va="top", fontsize=7.5)
    return fig

def hinh6_beta_2017_2020(d):
    """Hình 6 (Mục 4.3): hội tụ beta tuyệt đối giai đoạn 2017–2020 (thang đo cũ của EBI)."""
    return _ve_beta(d, 2017, 2020, "Giai đoạn 2017–2020 (thang đo cũ)"), "Hinh6_beta_2017_2020.png"

def hinh7_beta_2021_2024(d):
    """Hình 7 (Mục 4.3): hội tụ beta tuyệt đối giai đoạn 2021–2024 (thang đo mới của EBI)."""
    return _ve_beta(d, 2021, 2024, "Giai đoạn 2021–2024 (thang đo mới)"), "Hinh7_beta_2021_2024.png"

def hinh8_oaxaca(d):
    """Hình 8 (Mục 4.6): đóng góp của từng biến vào phần nguồn lực của khoảng cách ln(EBI)
    giữa vùng lõi (ĐNB + ĐBSH) và vùng ngoại vi, mẫu 2019–2024, hệ số tham chiếu gộp."""
    bien_X = ["ln_grdp", "dn_per_1000", "urban_share_pct", "trained_labor_pct", "ln_freight", "modern_retail_100k", "ln_fdi_pc", "hh_computer_pct"]
    nhan = {"ln_grdp": "ln GRDP/người", "dn_per_1000": "Mật độ DN", "urban_share_pct": "Đô thị hóa", "trained_labor_pct": "LĐ qua đào tạo",
            "ln_freight": "ln Vận tải HH/người", "modern_retail_100k": "Bán lẻ hiện đại", "ln_fdi_pc": "ln FDI/người", "hh_computer_pct": "Hộ có máy tính"}
    dd = d.copy(); dd["core"] = dd["vung6"].isin(VUNG_LOI).astype(int)
    dong_gop = oaxaca_chi_tiet(dd, bien_X)
    fig, ax = plt.subplots(figsize=(5.8, 3.0))
    ax.barh([nhan[k] for k in bien_X][::-1], dong_gop.values[::-1], color=MAU[0], height=0.6)
    ax.axvline(0, color=XAM, lw=0.8); ax.set_xlabel("Đóng góp vào khoảng cách ln(EBI) lõi – ngoại vi (%)"); ax.grid(axis="x", lw=0.4)
    return fig, "Hinh8_oaxaca_dong_gop.png"

DANH_SACH_HINH = {1: hinh1_khung_phan_tich, 2: hinh2_ebi_vung, 3: hinh3_theil_muc, 4: hinh4_theil_ty_trong,
                  5: hinh5_sigma, 6: hinh6_beta_2017_2020, 7: hinh7_beta_2021_2024, 8: hinh8_oaxaca}

# ----------------------------------------------------------------------------------------------
# 4. CHẠY
# ----------------------------------------------------------------------------------------------
if __name__ == "__main__":
    chon = [int(a) for a in sys.argv[1:]] or sorted(DANH_SACH_HINH)     # ví dụ: python tao_hinh.py 2 5
    d = doc_du_lieu()
    print(f"Đọc {FILE_DU_LIEU}: {len(d)} quan sát tỉnh–năm có EBI, {d.tinh.nunique()} tỉnh, {d.nam.min()}–{d.nam.max()}")
    for so in chon:
        fig, ten_file = DANH_SACH_HINH[so](d)
        duong_dan = os.path.join(THU_MUC_HINH, ten_file)
        fig.savefig(duong_dan); plt.close(fig)
        print(f"  Hình {so} → {duong_dan}")
    print("XONG. Mở thư mục 'hinh' để xem kết quả.")

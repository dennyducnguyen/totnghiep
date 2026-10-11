# -*- coding: utf-8 -*-
"""
BƯỚC 2 – VẼ 7 HÌNH CHO BÀI BÁO BẢN v8 (mỗi hình một nội dung, không ghép tiểu đồ)

Chạy (trong thư mục v8):   python scripts/02_ve_hinh.py          (vẽ cả 7 hình)
                           python scripts/02_ve_hinh.py 3 5      (chỉ vẽ Hình 3 và 5)
Đầu vào : ket-qua/*.csv do scripts/01_phan_tich.py tạo ra
Đầu ra  : hinh/Hinh1_khung_nghien_cuu.png ... hinh/Hinh7_ky_nang_cntt.png  (200 dpi)

  Hình 1  Khung nghiên cứu (3 câu hỏi, 5 nhóm giả thuyết)
  Hình 2  Khoảng cách tới tỉnh dẫn đầu (DTF) trung bình theo 6 vùng, 2017–2024    (RQ1)
  Hình 3  Phân rã Theil-T: tỷ trọng trong vùng / giữa vùng, 2017–2024            (H1)
  Hình 4  Hội tụ sigma: CV và Gini của EBI, 2017–2024                             (H3)
  Hình 5  Hội tụ beta trong giai đoạn thang đo ổn định                            (H4)
  Hình 6  Ma trận tương quan Pearson                                              (RQ3)
  Hình 7  Kỹ năng CNTT và EBI năm 2024                                            (H5e)
"""
import os, sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from matplotlib.ticker import FuncFormatter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KQ = os.path.join(ROOT, "ket-qua")
HINH = os.path.join(ROOT, "hinh"); os.makedirs(HINH, exist_ok=True)

DPI = 200
for ten_font in ["Arial", "Segoe UI", "DejaVu Sans"]:
    if any(f.name == ten_font for f in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = ten_font; break
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#c3c2b7", "axes.labelcolor": "#0b0b0b",
                     "xtick.color": "#52514e", "ytick.color": "#52514e", "grid.color": "#e1e0d9",
                     "figure.dpi": DPI, "savefig.dpi": DPI, "savefig.bbox": "tight"})
MAU = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#4a3aa7"]
NET = ["-", "--", "-.", ":", (0, (5, 1)), (0, (3, 1, 1, 1))]
XAM = "#898781"; XAM_NHAT = "#c3c2b7"; MUC = "#0b0b0b"; MUC_NHAT = "#52514e"
VUNG = ["ĐNB", "ĐBSH", "BTB&DHMT", "ĐBSCL", "TN", "TDMNPB"]
TEN_VUNG = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}

def vn(x, d=1):
    """Số kiểu Việt Nam: dấu phẩy thập phân."""
    return f"{x:.{d}f}".replace(".", ",").replace("-", "−")
def truc_vn(ax, d=1, truc="y"):
    f = FuncFormatter(lambda x, _: vn(x, d))
    (ax.yaxis if truc == "y" else ax.xaxis).set_major_formatter(f)
def to_vung_thang_do(ax):
    """Tô nền ba đoạn thang đo EBI: thang cũ 2017–2020, chuyển tiếp 2021, ổn định 2022–2024."""
    ax.axvspan(2020.5, 2021.5, color="#f4f4f2", zorder=0)
    ax.axvline(2020.5, color=XAM_NHAT, lw=0.8, zorder=1); ax.axvline(2021.5, color=XAM_NHAT, lw=0.8, zorder=1)
def nhan_thang_do(ax, y):
    for x, t in [(2018.75, "Thang cũ (4 trụ cột)"), (2021, "Chuyển\ntiếp"), (2023, "Thang ổn định (3 trụ cột)")]:
        ax.text(x, y, t, ha="center", va="top", fontsize=6.8, color=MUC_NHAT)

# ---------------------------------------------------------------------
def hinh1():
    fig, ax = plt.subplots(figsize=(7.6, 4.6)); ax.set_xlim(0, 100); ax.set_ylim(0, 70); ax.axis("off")
    def hop(x, y, w, h, tieu_de, noi_dung, fc="#eef4fc", ec="#2a78d6", fs=6.2):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.4,rounding_size=1.2", fc=fc, ec=ec, lw=1.2))
        ax.text(x + w / 2, y + h - 1.4, tieu_de, ha="center", va="top", fontsize=7.4, fontweight="bold", color=MUC)
        ax.text(x + w / 2, y + (h - 7) / 2, noi_dung, ha="center", va="center", fontsize=fs, color=MUC_NHAT, linespacing=1.35)
    def mui_ten(x0, y0, x1, y1, ls="-"):
        ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle="-|>", mutation_scale=11, color=MUC_NHAT, lw=1.1, ls=ls))
    # hàng trên: chuỗi phân tích
    hop(1, 44, 32, 24, "Điều kiện kinh tế – xã hội\ncủa tỉnh (Niên giám Thống kê)",
        "Cầu: bán lẻ/người, GRDP/người,\nđô thị hóa\nCung: mật độ DN, cơ sở cá thể, FDI\nNhân lực & hạ tầng: LĐ qua đào tạo,\nhộ có máy tính, kỹ năng CNTT",
        fc="#f4f4f2", ec=XAM)
    hop(40, 44, 22, 24, "Mức độ phát triển\nTMĐT (EBI – VECOM)",
        "EBI tổng hợp và\n3 trụ cột: NNL&HT,\nB2C, B2B (G2B đến 2020)\nDTF: vị thế so với\ntỉnh dẫn đầu", fc="#fdeee7", ec="#eb6834")
    hop(69, 44, 30, 24, "Khoảng cách số TMĐT\ngiữa các tỉnh",
        "Mức độ: CV, Gini, Theil-T\nCấu trúc: trong vùng / giữa vùng,\ntheo trụ cột\nĐộng thái: hội tụ sigma, beta")
    mui_ten(33.5, 56, 39.5, 56); mui_ten(62.5, 56, 68.5, 56)
    # hàng dưới: câu hỏi và giả thuyết
    hop(1, 13, 44, 21, "RQ3 – Các yếu tố liên quan (M2–M6)",
        "H5a: phía cầu → B2C;  H5b: phía cung → B2B\nH5c: lao động qua đào tạo → NNL&HT\nH5d: hộ có máy tính;  H5e: kỹ năng CNTT (2024)", fc="#ffffff", ec="#2a78d6")
    hop(51, 13, 23, 21, "RQ1 – Mức độ, cấu trúc",
        "H1: trong vùng chiếm\nphần lớn bất bình đẳng\nH2: B2B, NNL&HT phân\nhóa mạnh hơn B2C", fc="#ffffff", ec="#2a78d6")
    hop(77, 13, 22, 21, "RQ2 – Hội tụ",
        "H3: hội tụ sigma\nH4: hội tụ beta\n(tỉnh thấp tăng\nnhanh hơn)", fc="#ffffff", ec="#2a78d6")
    mui_ten(30, 34.4, 36.5, 54.5, ls="--"); mui_ten(62.5, 34.4, 78, 43.6, ls="--"); mui_ten(88, 34.4, 88, 43.6, ls="--")
    ax.text(50, 3.5, "Bối cảnh đo lường: EBI đổi phương pháp – thang cũ 2017–2020, chuyển tiếp 2021, thang ổn định 2022–2024\n"
            "→ dùng thước đo bất biến thang đo (CV, Gini, Theil-T, DTF)", ha="center", va="center", fontsize=6.4, color=MUC_NHAT,
            bbox=dict(boxstyle="round,pad=0.5", fc="#f4f4f2", ec=XAM_NHAT))
    return fig, "Hinh1_khung_nghien_cuu.png"

def hinh2():
    b = pd.read_csv(os.path.join(KQ, "bang2_khoang_cach_theo_nam.csv"))
    fig, ax = plt.subplots(figsize=(6.2, 3.4))
    ax.bar(b.nam, b.within_share, width=0.62, color=MAU[0], label="Trong nội bộ vùng")
    ax.bar(b.nam, b.between_share, width=0.62, bottom=b.within_share, color="#9ec5f4", label="Giữa các vùng")
    for _, r in b.iterrows():
        ax.text(r.nam, r.within_share / 2, vn(r.within_share), ha="center", va="center", fontsize=7, color="white", fontweight="bold")
        ax.text(r.nam, r.within_share + r.between_share / 2, vn(r.between_share), ha="center", va="center", fontsize=7, color=MUC)
    ax.set_ylim(0, 100); ax.set_ylabel("Tỷ trọng trong chỉ số Theil-T (%)"); ax.set_xlabel("Năm")
    ax.set_xticks(b.nam); ax.legend(fontsize=7, frameon=False, ncol=2, loc="upper left", bbox_to_anchor=(0, -0.15))
    return fig, "Hinh3_theil_phan_ra.png"

def hinh3():
    pv = pd.read_csv(os.path.join(KQ, "phu_dtf_vung_theo_nam.csv"), index_col=0)
    fig, ax = plt.subplots(figsize=(6.5, 3.7))
    to_vung_thang_do(ax)
    for i, v in enumerate(VUNG):
        ax.plot(pv.index, pv[v], color=MAU[i], ls=NET[i], lw=1.8, marker="o", ms=3.5, label=TEN_VUNG[v])
    ax.set_ylim(0, 68); nhan_thang_do(ax, 67.5)
    ax.set_ylabel("DTF trung bình của vùng (%)"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5)
    truc_vn(ax, 0)
    ax.legend(fontsize=7, frameon=False, ncol=2, loc="upper left", bbox_to_anchor=(0, -0.15))
    return fig, "Hinh2_dtf_vung.png"

def hinh4():
    b = pd.read_csv(os.path.join(KQ, "bang2_khoang_cach_theo_nam.csv"))
    fig, ax = plt.subplots(figsize=(6.3, 3.5))
    to_vung_thang_do(ax)
    ax.plot(b.nam, b.CV, color=MAU[0], lw=1.8, marker="o", ms=4, label="Hệ số biến thiên CV (trục trái)")
    ax.set_ylabel("Hệ số biến thiên (CV)"); ax.set_xlabel("Năm"); ax.set_ylim(0, 1.45); truc_vn(ax, 1)
    ax2 = ax.twinx(); ax2.spines["right"].set_visible(True)
    ax2.plot(b.nam, b.Gini, color=MAU[1], lw=1.8, ls="--", marker="s", ms=4, label="Hệ số Gini (trục phải)")
    ax2.set_ylabel("Hệ số Gini"); ax2.set_ylim(0, 0.42); truc_vn(ax2, 2)
    nhan_thang_do(ax, 1.43)
    h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, fontsize=7, frameon=False, ncol=2, loc="upper left", bbox_to_anchor=(0, -0.15))
    ax.set_xticks(b.nam)
    return fig, "Hinh4_hoi_tu_sigma.png"

def hinh5():
    s = pd.read_csv(os.path.join(KQ, "phu_hinh_beta.csv"))
    fig, ax = plt.subplots(figsize=(6.0, 3.9))
    for i, (nam, g) in enumerate(s.groupby("nam")):
        ax.scatter(g.ln_dtf_L1, g.g_dtf, s=16, color=MAU[i], alpha=0.75, label=f"Tăng trưởng {nam - 1}→{nam}", zorder=3)
    # đường hồi quy sau khi khử trung bình năm (tương đương hiệu ứng năm)
    sd = s.copy(); sd["g_c"] = sd.g_dtf - sd.groupby("nam").g_dtf.transform("mean")
    sd["x_c"] = sd.ln_dtf_L1 - sd.groupby("nam").ln_dtf_L1.transform("mean")
    b = (sd.x_c * sd.g_c).sum() / (sd.x_c ** 2).sum(); a = s.g_dtf.mean() - b * s.ln_dtf_L1.mean()
    xx = np.linspace(s.ln_dtf_L1.min(), s.ln_dtf_L1.max(), 50)
    ax.plot(xx, a + b * xx, color=MUC, lw=1.4, label=f"Đường hồi quy (β = {vn(b, 3)})")
    hai = s[s.tinh.isin(["Hà Nội", "TP. Hồ Chí Minh"])]
    ax.scatter(hai.ln_dtf_L1, hai.g_dtf, s=48, facecolors="none", edgecolors=MUC, lw=1.0, zorder=4,
               label="Hà Nội và TP. Hồ Chí Minh")
    ax.axhline(0, color=XAM, lw=0.7, ls=":")
    ax.set_xlabel("ln(DTF) năm trước"); ax.set_ylabel("Tăng trưởng ln(DTF) trong năm")
    truc_vn(ax, 1, "x"); truc_vn(ax, 1); ax.grid(lw=0.4)
    ax.legend(fontsize=7, frameon=False, loc="upper right")
    return fig, "Hinh5_hoi_tu_beta.png"

def hinh6():
    c = pd.read_csv(os.path.join(KQ, "phu_tuong_quan.csv"), index_col=0)
    nhan = {"EBI_dtf": "EBI (DTF)", "NNL_HT_dtf": "NNL&HT (DTF)", "B2C_dtf": "B2C (DTF)", "B2B_dtf": "B2B (DTF)",
            "ln_retail": "ln Bán lẻ/người", "ln_grdp": "ln GRDP/người", "dn_per_1000": "Mật độ DN",
            "cathe_per_1000": "Mật độ cơ sở cá thể", "trained_labor_pct": "LĐ qua đào tạo", "urban_share_pct": "Đô thị hóa",
            "ln_fdi_pc": "ln FDI/người", "hh_computer_pct": "Hộ có máy tính"}
    rows = list(c.index[:4]); cols = list(c.columns[4:])
    m = c.loc[rows, cols].values
    fig, ax = plt.subplots(figsize=(7.0, 2.9))
    im = ax.imshow(m, cmap="RdBu_r", vmin=-1, vmax=1, aspect="auto")
    ax.set_xticks(range(len(cols))); ax.set_xticklabels([nhan[x] for x in cols], rotation=35, ha="right", fontsize=7.5)
    ax.set_yticks(range(len(rows))); ax.set_yticklabels([nhan[x] for x in rows], fontsize=7.5)
    for i in range(m.shape[0]):
        for j in range(m.shape[1]):
            ax.text(j, i, vn(m[i, j], 2), ha="center", va="center", fontsize=7, color="white" if abs(m[i, j]) > 0.55 else MUC)
    for s in ax.spines.values(): s.set_visible(False)
    cb = fig.colorbar(im, ax=ax, fraction=0.03, pad=0.02, ticks=[-1, -0.5, 0, 0.5, 1]); cb.ax.tick_params(labelsize=7)
    cb.ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: vn(x, 1)))
    return fig, "Hinh6_tuong_quan.png"

def hinh7():
    s = pd.read_csv(os.path.join(KQ, "phu_hinh_ky_nang.csv"))
    fig, ax = plt.subplots(figsize=(6.0, 3.8))
    for i, v in enumerate(VUNG):
        g = s[s.vung6 == v]
        ax.scatter(g.it_skills_pct, g.EBI_dtf, s=18, color=MAU[i], alpha=0.85, label=TEN_VUNG[v], zorder=3)
    b, a = np.polyfit(s.it_skills_pct, s.EBI_dtf, 1); xx = np.linspace(s.it_skills_pct.min(), s.it_skills_pct.max(), 50)
    ax.plot(xx, a + b * xx, color=MUC, lw=1.3)
    for _, r in s.nlargest(4, "EBI_dtf").iterrows():
        ax.annotate(r.tinh, (r.it_skills_pct, r.EBI_dtf), fontsize=6.3, color=MUC_NHAT, xytext=(-4, 3), textcoords="offset points", ha="right")
    ax.set_xlabel("Tỷ lệ dân số có kỹ năng CNTT năm 2024 (%)"); ax.set_ylabel("DTF của EBI năm 2024 (%)")
    truc_vn(ax, 0, "x"); truc_vn(ax, 0); ax.grid(lw=0.4)
    ax.legend(fontsize=6.6, frameon=False, ncol=2, loc="upper left", bbox_to_anchor=(0, -0.15))
    return fig, "Hinh7_ky_nang_cntt.png"

DS = {1: hinh1, 2: hinh3, 3: hinh2, 4: hinh4, 5: hinh5, 6: hinh6, 7: hinh7}
if __name__ == "__main__":
    for so in [int(a) for a in sys.argv[1:]] or sorted(DS):
        fig, ten = DS[so](); fig.savefig(os.path.join(HINH, ten)); plt.close(fig); print("  Hình", so, "→ hinh/" + ten)
    print("XONG.")

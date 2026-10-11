# -*- coding: utf-8 -*-
"""
TẠO 4 HÌNH CHO BÀI BÁO BẢN v7 (mỗi hình một nội dung, không ghép tiểu đồ)
"Khoảng cách thương mại điện tử giữa các tỉnh, thành phố Việt Nam giai đoạn 2017–2024:
 mức độ, cấu trúc và các yếu tố liên quan"

CÁCH CHẠY: python tao_hinh_v7.py            (tạo cả 4 hình)
           python tao_hinh_v7.py 2 3        (chỉ tạo Hình 2 và 3)
ĐẦU VÀO : ../ket-qua-v7/panel_v7.csv  (do scripts/01_phan_tich_v7.py tạo ra)
ĐẦU RA  : hinh-v7/Hinh1_khung_phan_tich.png, Hinh2_ebi_vung.png, Hinh3_theil_phan_ra.png, Hinh4_gini_tru_cot.png
Các hình của bản v6 trong thư mục hinh/ giữ nguyên.
"""
import os, sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

THU_MUC = os.path.dirname(os.path.abspath(__file__))
FILE_DU_LIEU = os.path.join(os.path.dirname(THU_MUC), "ket-qua-v7", "panel_v7.csv")
THU_MUC_HINH = os.path.join(THU_MUC, "hinh-v7"); os.makedirs(THU_MUC_HINH, exist_ok=True)

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
VUNG = ["ĐBSH", "TDMNPB", "BTB&DHMT", "TN", "ĐNB", "ĐBSCL"]
TEN_VUNG = {"ĐBSH": "Đồng bằng sông Hồng", "TDMNPB": "Trung du và miền núi phía Bắc",
            "BTB&DHMT": "Bắc Trung Bộ và Duyên hải miền Trung", "TN": "Tây Nguyên",
            "ĐNB": "Đông Nam Bộ", "ĐBSCL": "Đồng bằng sông Cửu Long"}
NAM_DOI_PP = 2020.5

def doc_du_lieu():
    df = pd.read_csv(FILE_DU_LIEU, encoding="utf-8-sig")
    return df.dropna(subset=["EBI"]).copy()

def gini(x):
    x = np.sort(np.asarray(x, float)); n = len(x)
    return (2 * np.sum(np.arange(1, n + 1) * x) / (n * x.sum())) - (n + 1) / n
def theil_T(x):
    x = np.asarray(x, float); m = x.mean(); return np.mean(x / m * np.log(x / m))
def phan_ra_theil(x, nhom):
    x = np.asarray(x, float); nhom = np.asarray(nhom); n = len(x); m = x.mean()
    T = theil_T(x); trong = giua = 0.0
    for k in np.unique(nhom):
        xk = x[nhom == k]; sk = len(xk) / n * xk.mean() / m
        trong += sk * theil_T(xk); giua += sk * np.log(xk.mean() / m)
    return T, giua, trong

def hinh1_khung_phan_tich(d=None):
    """Hình 1 (Mục 3): khung phân tích – chuỗi nguồn lực → năng lực ứng dụng & hệ sinh thái → mức độ phát triển TMĐT → khoảng cách."""
    fig, ax = plt.subplots(figsize=(7.4, 4.0)); ax.set_xlim(0, 100); ax.set_ylim(0, 58); ax.axis("off")
    def hop(x, y, w, h, tieu_de, noi_dung, fc="#eef4fc", ec="#2a78d6"):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.4,rounding_size=1.2", fc=fc, ec=ec, lw=1.2))
        ax.text(x + w / 2, y + h - 2.0, tieu_de, ha="center", va="top", fontsize=7.8, fontweight="bold", color=MUC)
        ax.text(x + w / 2, y + h / 2 - 2.6, noi_dung, ha="center", va="center", fontsize=6.1, color=MUC_NHAT, linespacing=1.35)
    def mui_ten(x0, y0, x1, y1):
        ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle="-|>", mutation_scale=11, color=MUC_NHAT, lw=1.1))
    hop(1, 34, 22, 20, "Nguồn lực và điều kiện\nnền tảng (biến đầu vào)",
        "Quy mô kinh tế, mật độ DN,\nđô thị hóa, lao động qua đào tạo,\nhạ tầng số, logistics, FDI", fc="#f4f4f2", ec=XAM)
    hop(27, 34, 22, 20, "Năng lực ứng dụng và\nhệ sinh thái TMĐT (cơ chế)",
        "Khối lượng tới hạn, hiệu ứng\nmạng và học hỏi giữa DN,\nnền tảng – giao nhận – thanh toán", fc="#f4f4f2", ec=XAM)
    hop(53, 34, 22, 20, "Mức độ phát triển TMĐT\n(chỉ báo: EBI)",
        "Chỉ số tổng hợp: NNL&HT,\ngiao dịch B2C, giao dịch B2B\n(chưa tách riêng tầng kết quả)", fc="#fdeee7", ec="#eb6834")
    hop(79, 34, 20, 20, "Khoảng cách\ngiữa các tỉnh (kết quả)",
        "Gini, Theil-T;\ngiữa vùng / trong vùng;\nthay đổi theo giai đoạn")
    mui_ten(23.4, 44, 26.6, 44); mui_ten(49.4, 44, 52.6, 44); mui_ten(75.4, 44, 78.6, 44)
    hop(1, 4, 46, 20, "Biến kiểm soát và bối cảnh",
        "Hiệu ứng cố định tỉnh (địa lý, thể chế, lịch sử);\nhiệu ứng cố định năm (chu kỳ, COVID-19,\nđổi phương pháp tính EBI năm 2021)", fc="#ffffff", ec=XAM_NHAT)
    hop(53, 4, 46, 20, "Câu hỏi nghiên cứu",
        "RQ1: mức độ và thay đổi của khoảng cách\nRQ2: giữa vùng hay trong vùng\nRQ3: đặc điểm nào liên hệ với EBI (FE tỉnh + năm)", fc="#ffffff", ec=XAM_NHAT)
    mui_ten(24, 24.4, 40, 33.6); mui_ten(76, 24.4, 88, 33.6)
    return fig, "Hinh1_khung_phan_tich.png"

def hinh2_ebi_vung(d):
    """Hình 2 (Mục 5.2): EBI tương đối theo 6 vùng, 2017–2024 (trung bình các tỉnh có số liệu trong năm = 100)."""
    pv = d.groupby(["nam", "vung6"])["EBI_rel"].mean().unstack()[VUNG]
    fig, ax = plt.subplots(figsize=(6.5, 3.6))
    for i, v in enumerate(VUNG):
        ax.plot(pv.index, pv[v], color=MAU[i], ls=NET[i], lw=1.8, marker="o", ms=3.5, label=TEN_VUNG[v])
    ax.axhline(100, color=XAM, lw=0.8, ls=":")
    ax.axvline(NAM_DOI_PP, color=XAM_NHAT, lw=0.8)
    ax.text(NAM_DOI_PP - 0.05, ax.get_ylim()[0] + 2, "Đổi phương pháp tính EBI →", fontsize=7, color=MUC_NHAT, va="bottom", ha="right")
    ax.set_ylabel("EBI tương đối (trung bình các tỉnh = 100)"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5)
    ax.legend(fontsize=7, frameon=False, ncol=2, loc="upper left", bbox_to_anchor=(0, -0.18))
    return fig, "Hinh2_ebi_vung.png"

def hinh3_theil_phan_ra(d):
    """Hình 3 (Mục 5.3): chỉ số Theil-T của EBI và hai thành phần giữa các vùng / trong vùng, 2017–2024."""
    dong = []
    for nam, g in d.groupby("nam"):
        T, Tg, Tt = phan_ra_theil(g.EBI, g.vung6); dong.append(dict(nam=nam, Theil=T, giua=Tg, trong=Tt))
    b = pd.DataFrame(dong)
    fig, ax = plt.subplots(figsize=(5.8, 3.3))
    ax.bar(b.nam, b.giua, width=0.6, color=MAU[0], label="Giữa các vùng (Th_B)")
    ax.bar(b.nam, b.trong, width=0.6, bottom=b.giua, color="#9ec5f4", label="Trong nội bộ vùng (Th_W)")
    for _, r in b.iterrows():
        ax.text(r.nam, r.Theil + 0.006, f"{r.giua / r.Theil * 100:.0f}%", ha="center", va="bottom", fontsize=6.8, color=MUC_NHAT)
    ax.axvline(NAM_DOI_PP, color=XAM_NHAT, lw=0.8)
    ax.set_ylabel("Chỉ số Theil-T"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5)
    ax.legend(fontsize=7, frameon=False, loc="upper left")
    ax.text(0.99, 0.97, "Số trên cột: tỷ trọng giữa các vùng (%)", transform=ax.transAxes, ha="right", va="top", fontsize=6.8, color=MUC_NHAT)
    return fig, "Hinh3_theil_phan_ra.png"

def hinh4_gini_tru_cot(d):
    """Hình 4 (Mục 5.2): hệ số Gini của EBI và từng trụ cột theo năm."""
    rows = []
    for nam, g in d.groupby("nam"):
        r = dict(nam=nam, EBI=gini(g.EBI), NNL_HT=gini(g.NNL_HT), B2C=gini(g.B2C), B2B=gini(g.B2B))
        if nam <= 2020: r["G2B"] = gini(g.G2B.dropna())
        rows.append(r)
    b = pd.DataFrame(rows)
    nhan = [("EBI", "EBI tổng hợp"), ("NNL_HT", "Nguồn nhân lực & hạ tầng"), ("B2C", "Giao dịch B2C"), ("B2B", "Giao dịch B2B"), ("G2B", "Giao dịch G2B (đến 2020)")]
    fig, ax = plt.subplots(figsize=(6.2, 3.4))
    for i, (cot, ten) in enumerate(nhan):
        ax.plot(b.nam, b[cot], color=MAU[i], ls=NET[i], lw=1.8 if cot == "EBI" else 1.4, marker="o", ms=3.5, label=ten)
    ax.axvline(NAM_DOI_PP, color=XAM_NHAT, lw=0.8)
    ax.set_ylabel("Hệ số Gini"); ax.set_xlabel("Năm"); ax.grid(axis="y", lw=0.5); ax.set_ylim(0, 0.55)
    ax.legend(fontsize=7, frameon=False, ncol=2, loc="upper left")
    return fig, "Hinh4_gini_tru_cot.png"

DANH_SACH_HINH = {1: hinh1_khung_phan_tich, 2: hinh2_ebi_vung, 3: hinh3_theil_phan_ra, 4: hinh4_gini_tru_cot}

if __name__ == "__main__":
    chon = [int(a) for a in sys.argv[1:]] or sorted(DANH_SACH_HINH)
    d = doc_du_lieu()
    print(f"Đọc {FILE_DU_LIEU}: {len(d)} quan sát tỉnh–năm có EBI")
    for so in chon:
        fig, ten = DANH_SACH_HINH[so](d)
        fig.savefig(os.path.join(THU_MUC_HINH, ten)); plt.close(fig); print("  Hình", so, "→", ten)
    print("XONG.")

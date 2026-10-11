# -*- coding: utf-8 -*-
"""12 - Tính các VÍ DỤ MINH HOẠ QUY MÔ NHỎ (6 tỉnh, số thật) cho từng phương pháp trong bài,
   ghi kết quả từng bước ra ket-qua/vi_du_tung_buoc.json và sổ Excel có công thức sống
   HUONG-DAN-TINH-TUNG-BUOC.xlsx (mỗi sheet = một phương pháp, ô công thức Excel để người đọc bấm xem).

Cùng một bộ số này được script 13 (Node) dùng để dựng file Word HUONG-DAN-TINH-TUNG-BUOC.docx.
Chạy: python scripts/12_vi_du_tung_buoc.py
"""
import json, os
import numpy as np, pandas as pd
import statsmodels.api as sm
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_JSON = os.path.join(ROOT, 'ket-qua', 'vi_du_tung_buoc.json')
OUT_XLSX = os.path.join(ROOT, 'HUONG-DAN-TINH-TUNG-BUOC.xlsx')

TINH = ['Hà Nội', 'Bình Dương', 'Đà Nẵng', 'Cần Thơ', 'Lâm Đồng', 'Thái Nguyên']
VUNG_NGAN = {'ĐBSH': 'ĐB sông Hồng', 'ĐNB': 'Đông Nam Bộ', 'BTB&DHMT': 'BTB & DH miền Trung',
             'ĐBSCL': 'ĐB sông Cửu Long', 'TN': 'Tây Nguyên', 'TDMNPB': 'TD & MN phía Bắc'}
# Nhóm vùng thu gọn cho ví dụ phân rã (3 nhóm, mỗi nhóm 2 tỉnh) để phần "giữa vùng" có ý nghĩa
NHOM = {'Hà Nội': 'Lõi', 'Bình Dương': 'Lõi', 'Đà Nẵng': 'Miền Trung', 'Cần Thơ': 'ĐBSCL – TN – TDMN',
        'Lâm Đồng': 'ĐBSCL – TN – TDMN', 'Thái Nguyên': 'ĐBSCL – TN – TDMN'}
# Sửa: để ví dụ gọn, gom Đà Nẵng cùng nhóm với Cần Thơ -> 2 nhóm "Lõi" (2 tỉnh) và "Ngoại vi" (4 tỉnh)
NHOM = {t: ('Lõi' if t in ('Hà Nội', 'Bình Dương') else 'Ngoại vi') for t in TINH}

df = pd.read_csv(os.path.join(ROOT, 'ket-qua', 'panel_phan_tich.csv'))
VD = {'tinh': TINH, 'nhom': NHOM}


def r(x, n=4):
    return float(np.round(x, n))


# ------------------------------------------------------------------ 0. Bảng số liệu gốc của 6 tỉnh
d24 = df[(df.nam == 2024) & df.tinh.isin(TINH)].set_index('tinh').loc[TINH]
d21 = df[(df.nam == 2021) & df.tinh.isin(TINH)].set_index('tinh').loc[TINH]
goc = []
for t in TINH:
    goc.append(dict(tinh=t, vung=VUNG_NGAN[d24.loc[t, 'vung6']], nhom=NHOM[t],
                    EBI_2024=float(d24.loc[t, 'EBI']), NNL_HT=float(d24.loc[t, 'NNL_HT']), B2C=float(d24.loc[t, 'B2C']),
                    B2B=float(d24.loc[t, 'B2B']), EBI_2021=float(d21.loc[t, 'EBI']),
                    dn_per_1000=r(d24.loc[t, 'dn_per_1000'], 2), grdp_pc_tr=r(d24.loc[t, 'grdp_pc_tr'], 1)))
VD['bang_goc'] = goc

# ------------------------------------------------------------------ 1. EBI tương đối (dùng toàn bộ 58 tỉnh 2024)
tb_2024 = float(df[(df.nam == 2024)].EBI.mean())
VD['ebi_rel'] = dict(so_tinh_2024=int(df[(df.nam == 2024)].EBI.notna().sum()), trung_binh_2024=r(tb_2024, 3),
                     vi_du=[dict(tinh=t, EBI=float(d24.loc[t, 'EBI']), EBI_rel=r(d24.loc[t, 'EBI'] / tb_2024 * 100, 1)) for t in TINH])

# ------------------------------------------------------------------ 2. Gini (thứ hạng)
y = np.array([g['EBI_2024'] for g in goc]); n = len(y)
ys = np.sort(y); j = np.arange(1, n + 1)
gini_steps = [dict(j=int(jj), y=float(yy), j_y=float(jj * yy)) for jj, yy in zip(j, ys)]
S1 = float((j * ys).sum()); S = float(ys.sum())
G = 2 * S1 / (n * S) - (n + 1) / n
VD['gini'] = dict(n=n, sap_xep=gini_steps, tong_j_y=S1, tong_y=S, so_hang_1=r(2 * S1 / (n * S)), so_hang_2=r((n + 1) / n), G=r(G))
# kiểm tra bằng công thức chênh lệch cặp
pairs = np.abs(y[:, None] - y[None, :]).sum() / (2 * n * n * y.mean())
VD['gini']['G_kiem_tra_cach_2'] = r(pairs)

# ------------------------------------------------------------------ 3. CV, SD ln
mu = y.mean(); sd = y.std(ddof=1)
VD['cv'] = dict(mean=r(mu, 3), sd=r(sd, 3), CV=r(sd / mu), ln_y=[r(v, 3) for v in np.log(y)], SD_ln=r(np.log(y).std(ddof=1)),
                lech=[r(v - mu, 3) for v in y], binh_phuong=[r((v - mu) ** 2, 2) for v in y], tong_bp=r(((y - mu) ** 2).sum(), 2))

# ------------------------------------------------------------------ 4. Theil-T và phân rã 2 nhóm
ratio = y / mu; lnr = np.log(ratio); term = ratio * lnr
T = float(term.mean())
VD['theil'] = dict(mean=r(mu, 3), buoc=[dict(tinh=t, y=float(v), y_tren_mu=r(a, 4), ln=r(b, 4), tich=r(c, 4))
                                          for t, v, a, b, c in zip(TINH, y, ratio, lnr, term)], tong=r(term.sum()), T=r(T))
grp = np.array([NHOM[t] for t in TINH])
pr = []
within = between = 0.0
for k in ['Lõi', 'Ngoại vi']:
    yk = y[grp == k]; nk = len(yk); mk = yk.mean(); sk = nk / n * mk / mu
    Tk = float(np.mean(yk / mk * np.log(yk / mk)))
    within += sk * Tk; between += sk * np.log(mk / mu)
    pr.append(dict(nhom=k, n=nk, mean=r(mk, 3), s_k=r(sk), T_k=r(Tk), s_T=r(sk * Tk), ln_mu_ratio=r(np.log(mk / mu)), s_ln=r(sk * np.log(mk / mu))))
VD['theil']['phan_ra'] = dict(nhom=pr, within=r(within), between=r(between), tong=r(within + between),
                              ty_trong_giua=r(between / (within + between) * 100, 1))

# ------------------------------------------------------------------ 5. Phân rã Gini theo trụ cột (2024: 0,2/0,4/0,4)
w = np.array([.2, .4, .4]); comp = d24[['NNL_HT', 'B2C', 'B2B']].values.astype(float)
Y = comp @ w; rankY = pd.Series(Y).rank().values / n
def gini_fn(x):
    x = np.sort(np.asarray(x, float)); m = len(x)
    return 2 * np.sum(np.arange(1, m + 1) * x) / (m * x.sum()) - (m + 1) / m
rows = []
for k, name in enumerate(['NNL_HT', 'B2C', 'B2B']):
    Xk = comp[:, k] * w[k]; Sk = Xk.mean() / Y.mean(); Gk = gini_fn(Xk)
    rankk = pd.Series(Xk).rank().values / n
    Rk = np.cov(Xk, rankY)[0, 1] / np.cov(Xk, rankk)[0, 1]
    rows.append(dict(tru_cot=name, w=float(w[k]), wX=[r(v, 2) for v in Xk], S=r(Sk), G=r(Gk), R=r(Rk), tich=r(Sk * Gk * Rk)))
tot = sum(x['tich'] for x in rows)
for x in rows: x['dong_gop_pct'] = r(x['tich'] / tot * 100, 1)
VD['gini_nguon'] = dict(w=w.tolist(), Y_tai_tao=[r(v, 2) for v in Y], EBI_cong_bo=y.tolist(), rank_Y=[r(v, 3) for v in rankY],
                        G_Y=r(gini_fn(Y)), tru_cot=rows, tong_tich=r(tot))

# ------------------------------------------------------------------ 6. Hội tụ beta tuyệt đối 2021–2024 (6 tỉnh)
y0 = np.array([g['EBI_2021'] for g in goc]); y1 = y; Tn = 3
growth = np.log(y1 / y0) / Tn; lny0 = np.log(y0)
X = sm.add_constant(lny0); m = sm.OLS(growth, X).fit()
b = float(m.params[1]); a = float(m.params[0])
lam = -np.log(1 + b * Tn) / Tn; half = np.log(2) / lam
VD['beta'] = dict(T=Tn, buoc=[dict(tinh=t, y0=float(v0), y1=float(v1), ln_y0=r(np.log(v0)), ln_y1=r(np.log(v1)), g=r(gg))
                              for t, v0, v1, gg in zip(TINH, y0, y1, growth)],
                  mean_x=r(lny0.mean()), mean_g=r(growth.mean()),
                  Sxy=r(((lny0 - lny0.mean()) * (growth - growth.mean())).sum()), Sxx=r(((lny0 - lny0.mean()) ** 2).sum()),
                  beta=r(b), alpha=r(a), R2=r(m.rsquared, 3), lam=r(lam), lam_pct=r(lam * 100, 1), half_life=r(half, 1),
                  # ban day du trong bai
                  bai_bao=dict(beta=-0.0888, se=0.0180, N=56, lam_pct=10.3, half=6.7))

# ------------------------------------------------------------------ 7. Sigma: CV của 6 tỉnh 2021 vs 2024
VD['sigma'] = dict(CV_2021=r(y0.std(ddof=1) / y0.mean()), CV_2024=r(y1.std(ddof=1) / y1.mean()),
                   mean_2021=r(y0.mean(), 3), sd_2021=r(y0.std(ddof=1), 3), mean_2024=r(y1.mean(), 3), sd_2024=r(y1.std(ddof=1), 3))

# ------------------------------------------------------------------ 8. OLS đơn giản: ln EBI ~ mật độ DN (6 tỉnh 2024)
xdn = d24['dn_per_1000'].values.astype(float); lny = np.log(y)
Xo = sm.add_constant(xdn); mo = sm.OLS(lny, Xo).fit()   # SE cổ điển để khớp LINEST trong Excel; bài dùng SE cụm
b1 = float(mo.params[1]); b0 = float(mo.params[0])
VD['ols'] = dict(buoc=[dict(tinh=t, x=r(v, 2), ln_ebi=r(l, 4), du_bao=r(b0 + b1 * v, 4), phan_du=r(l - b0 - b1 * v, 4)) for t, v, l in zip(TINH, xdn, lny)],
                 mean_x=r(xdn.mean(), 3), mean_y=r(lny.mean(), 4),
                 Sxy=r(((xdn - xdn.mean()) * (lny - lny.mean())).sum()), Sxx=r(((xdn - xdn.mean()) ** 2).sum()),
                 b1=r(b1), b0=r(b0), R2=r(mo.rsquared, 3), se=r(float(mo.bse[1])), t=r(float(mo.tvalues[1]), 2), p=r(float(mo.pvalues[1])),
                 doc_he_so_pct=r((np.exp(b1) - 1) * 100, 1), bai_bao=dict(he_so=0.070, se=0.007, pct=7.3))
# VIF minh hoạ với 2 biến
x2 = d24[['dn_per_1000', 'ln_grdp']].values.astype(float)
r2 = sm.OLS(x2[:, 0], sm.add_constant(x2[:, 1])).fit().rsquared
VD['vif'] = dict(R2_dn_tren_grdp=r(r2, 3), VIF=r(1 / (1 - r2), 2), bai_bao_max=5.37)

# ------------------------------------------------------------------ 9. Oaxaca–Blinder 2 nhóm, 1 biến (mật độ DN)
A = [t for t in TINH if NHOM[t] == 'Lõi']; Bn = [t for t in TINH if NHOM[t] == 'Ngoại vi']
def fit(ts):
    xx = d24.loc[ts, 'dn_per_1000'].values.astype(float); yy = np.log(d24.loc[ts, 'EBI'].values.astype(float))
    mm = sm.OLS(yy, sm.add_constant(xx)).fit(); return float(mm.params[0]), float(mm.params[1]), xx.mean(), yy.mean()
aA, bA, xA, yA = fit(A); aB, bB, xB, yB = fit(Bn)
# hệ số tham chiếu = hệ số nhóm A (phiên bản đơn giản để minh hoạ; bài dùng hệ số gộp Jann 2008)
gap = yA - yB; endow = (xA - xB) * bA; coef = gap - endow
VD['oaxaca'] = dict(A=A, B=Bn, xA=r(xA, 3), xB=r(xB, 3), yA=r(yA, 4), yB=r(yB, 4), aA=r(aA), bA=r(bA), aB=r(aB), bB=r(bB),
                    gap=r(gap), endow=r(endow), coef=r(coef), endow_pct=r(endow / gap * 100, 1), coef_pct=r(coef / gap * 100, 1),
                    gap_pct_ebi=r((np.exp(gap) - 1) * 100, 1),
                    bai_bao=dict(gap=0.455, endow=0.360, endow_pct=79.1, coef=0.095, coef_pct=20.9))

json.dump(VD, open(OUT_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('Đã ghi', OUT_JSON)

# ================================================================== EXCEL CÓ CÔNG THỨC SỐNG
wb = Workbook()
DAM = Font(bold=True); TD = Font(bold=True, color='FFFFFF'); NEN = PatternFill('solid', fgColor='1F4E79')
VANG = PatternFill('solid', fgColor='FFF2CC'); XANH = PatternFill('solid', fgColor='E2EFDA'); XAM = PatternFill('solid', fgColor='F2F2F2')
VIEN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
WR = Alignment(wrap_text=True, vertical='top')


def sheet(title):
    ws = wb.create_sheet(title); ws.sheet_view.showGridLines = True; return ws


def tieu_de(ws, dong, text, size=13):
    c = ws.cell(row=dong, column=1, value=text); c.font = Font(bold=True, size=size); return dong + 1


def ghi(ws, dong, cot, val, **kw):
    c = ws.cell(row=dong, column=cot, value=val)
    if kw.get('bold'): c.font = DAM
    if kw.get('fill'): c.fill = kw['fill']
    if kw.get('wrap'): c.alignment = WR
    if kw.get('fmt'): c.number_format = kw['fmt']
    if kw.get('border', True): c.border = VIEN
    return c


def header(ws, dong, cols, start=1):
    for i, h in enumerate(cols):
        c = ws.cell(row=dong, column=start + i, value=h); c.font = TD; c.fill = NEN; c.alignment = Alignment(wrap_text=True, horizontal='center'); c.border = VIEN
    return dong + 1


def chu_thich(ws, dong, text):
    c = ws.cell(row=dong, column=1, value=text); c.alignment = WR; c.font = Font(italic=True, color='595959')
    ws.merge_cells(start_row=dong, start_column=1, end_row=dong, end_column=8); ws.row_dimensions[dong].height = 30
    return dong + 1


# ---------- Sheet 0: Hướng dẫn
ws = wb.active; ws.title = '0_Huong_dan'
rr = tieu_de(ws, 1, 'SỔ TÍNH MINH HOẠ TỪNG BƯỚC – 6 TỈNH, SỐ THẬT NĂM 2021 & 2024', 14)
for line in [
    'Mỗi sheet là một phương pháp trong bài báo. Ô màu VÀNG = số liệu đầu vào; ô màu XANH = công thức Excel (bấm vào ô để xem công thức); ô XÁM = kết quả cuối.',
    'Bộ số ví dụ: 6 tỉnh Hà Nội, Bình Dương (vùng lõi) và Đà Nẵng, Cần Thơ, Lâm Đồng, Thái Nguyên (vùng ngoại vi), lấy đúng số trong panel_phan_tich.csv.',
    'Kết quả trên 6 tỉnh chỉ để HIỂU CÁCH TÍNH; con số trong bài báo tính trên 54–58 tỉnh mỗi năm (xem cột "Trong bài báo" ở mỗi sheet).',
    'Đọc kèm file Word HUONG-DAN-TINH-TUNG-BUOC.docx: Word giải thích ý nghĩa và diễn giải; Excel cho thấy phép tính chạy thật.',
    'Sheet: 1_Du_lieu → 2_EBI_tuong_doi → 3_Gini → 4_CV_sigma → 5_Theil_phan_ra → 6_Gini_tru_cot → 7_Beta_hoi_tu → 8_OLS → 9_Oaxaca.',
]:
    ws.cell(row=rr, column=1, value=line).alignment = WR; ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8); ws.row_dimensions[rr].height = 32; rr += 1
ws.column_dimensions['A'].width = 120

# ---------- Sheet 1: dữ liệu
ws = sheet('1_Du_lieu'); rr = tieu_de(ws, 1, 'Bảng số liệu gốc của 6 tỉnh (trích panel_phan_tich.csv)')
rr = header(ws, rr, ['Tỉnh', 'Vùng (Niên giám)', 'Nhóm ví dụ', 'EBI 2024', 'NNL&HT 2024', 'B2C 2024', 'B2B 2024', 'EBI 2021', 'DN/1.000 dân 2024', 'GRDP/người 2024 (tr.đ)'])
R0 = rr
for g in goc:
    for i, v in enumerate([g['tinh'], g['vung'], g['nhom'], g['EBI_2024'], g['NNL_HT'], g['B2C'], g['B2B'], g['EBI_2021'], g['dn_per_1000'], g['grdp_pc_tr']]):
        ghi(ws, rr, i + 1, v, fill=VANG if i >= 3 else None)
    rr += 1
R1 = rr - 1
rr = chu_thich(ws, rr + 1, 'Các sheet sau tham chiếu trực tiếp vào vùng ô D–J của sheet này. Vùng: ĐBSH = Đồng bằng sông Hồng; ĐNB = Đông Nam Bộ; BTB&DHMT = Bắc Trung Bộ và Duyên hải miền Trung; TN = Tây Nguyên; TDMNPB = Trung du và miền núi phía Bắc; ĐBSCL = Đồng bằng sông Cửu Long.')
for col, wd in zip('ABCDEFGHIJ', [14, 22, 12, 10, 12, 10, 10, 10, 16, 20]): ws.column_dimensions[col].width = wd
DL = "'1_Du_lieu'!"

# ---------- Sheet 2: EBI tương đối
ws = sheet('2_EBI_tuong_doi'); rr = tieu_de(ws, 1, 'EBI tương đối = EBI của tỉnh / trung bình các tỉnh có số liệu trong năm × 100   (Bảng 2, Hình 2)')
ghi(ws, rr, 1, 'Trung bình EBI của 58 tỉnh năm 2024 (tính trên toàn bộ dữ liệu):', bold=True, border=False); ghi(ws, rr, 6, r(tb_2024, 3), fill=VANG, fmt='0.000'); rr += 1
ghi(ws, rr, 1, '(Trong ví dụ này trung bình lấy từ 58 tỉnh, không phải 6 tỉnh, để số EBI tương đối khớp với bài)', border=False); rr += 2
rr = header(ws, rr, ['Tỉnh', 'EBI 2024', 'EBI tương đối', 'Đọc kết quả'])
for i, g in enumerate(goc):
    ghi(ws, rr, 1, g['tinh']); ghi(ws, rr, 2, f'={DL}D{R0 + i}', fill=VANG, fmt='0.0')
    ghi(ws, rr, 3, f'=B{rr}/$F$2*100', fill=XANH, fmt='0.0')
    ghi(ws, rr, 4, f'=IF(C{rr}>=100,"cao hơn trung bình cả nước "&TEXT(C{rr}-100,"0")&" điểm","thấp hơn trung bình cả nước "&TEXT(100-C{rr},"0")&" điểm")', fill=XANH)
    rr += 1
for col, wd in zip('ABCD', [14, 10, 14, 48]): ws.column_dimensions[col].width = wd

# ---------- Sheet 3: Gini
ws = sheet('3_Gini'); rr = tieu_de(ws, 1, 'Hệ số Gini theo công thức thứ hạng   (Bảng 3, Bảng 4)')
rr = chu_thich(ws, rr, 'G = 2·Σ(j·y_j) / (N·Σy_j) − (N+1)/N, với y đã SẮP TĂNG DẦN và j = thứ hạng 1..N. G = 0: mọi tỉnh bằng nhau; G càng gần 1 càng tập trung.')
rr = header(ws, rr, ['Thứ hạng j', 'Tỉnh', 'EBI 2024 (sắp tăng dần)', 'j × y'])
G0 = rr
order = np.argsort(y)
for jj, idx in enumerate(order, 1):
    ghi(ws, rr, 1, jj, fill=VANG); ghi(ws, rr, 2, TINH[idx]); ghi(ws, rr, 3, f'=SMALL({DL}$D${R0}:$D${R1},A{rr})', fill=XANH, fmt='0.0')
    ghi(ws, rr, 4, f'=A{rr}*C{rr}', fill=XANH, fmt='0.0'); rr += 1
G1 = rr - 1
ghi(ws, rr, 2, 'Tổng', bold=True); ghi(ws, rr, 3, f'=SUM(C{G0}:C{G1})', fill=XANH, fmt='0.0'); ghi(ws, rr, 4, f'=SUM(D{G0}:D{G1})', fill=XANH, fmt='0.0'); rT = rr; rr += 2
ghi(ws, rr, 1, 'N =', bold=True); ghi(ws, rr, 2, f'=COUNT(C{G0}:C{G1})', fill=XANH); rN = rr; rr += 1
ghi(ws, rr, 1, 'Số hạng 1 = 2·Σ(j·y)/(N·Σy)', bold=True); ghi(ws, rr, 3, f'=2*D{rT}/(B{rN}*C{rT})', fill=XANH, fmt='0.0000'); rS1 = rr; rr += 1
ghi(ws, rr, 1, 'Số hạng 2 = (N+1)/N', bold=True); ghi(ws, rr, 3, f'=(B{rN}+1)/B{rN}', fill=XANH, fmt='0.0000'); rS2 = rr; rr += 1
ghi(ws, rr, 1, 'GINI = số hạng 1 − số hạng 2', bold=True); ghi(ws, rr, 3, f'=C{rS1}-C{rS2}', fill=XAM, fmt='0.0000'); rr += 2
rr = chu_thich(ws, rr, f'Trong bài báo: Gini EBI năm 2024 trên 58 tỉnh = 0,245 (Bảng 3). Ví dụ 6 tỉnh cho {r(G, 3)} vì mẫu nhỏ gồm cả Hà Nội (84) và Thái Nguyên (20).')
for col, wd in zip('ABCD', [30, 14, 24, 12]): ws.column_dimensions[col].width = wd

# ---------- Sheet 4: CV / sigma
ws = sheet('4_CV_sigma'); rr = tieu_de(ws, 1, 'Hệ số biến thiên CV = độ lệch chuẩn / trung bình → hội tụ sigma khi CV giảm theo thời gian   (Bảng 3, Hình 5)')
rr = header(ws, rr, ['Tỉnh', 'EBI 2021', 'EBI 2024', 'ln(EBI 2024)'])
C0 = rr
for i, g in enumerate(goc):
    ghi(ws, rr, 1, g['tinh']); ghi(ws, rr, 2, f'={DL}H{R0 + i}', fill=VANG, fmt='0.00'); ghi(ws, rr, 3, f'={DL}D{R0 + i}', fill=VANG, fmt='0.0'); ghi(ws, rr, 4, f'=LN(C{rr})', fill=XANH, fmt='0.000'); rr += 1
C1 = rr - 1; rr += 1
for lab, f2, f3, fmt in [('Trung bình', f'=AVERAGE(B{C0}:B{C1})', f'=AVERAGE(C{C0}:C{C1})', '0.000'),
                         ('Độ lệch chuẩn mẫu (STDEV.S)', f'=_xlfn.STDEV.S(B{C0}:B{C1})', f'=_xlfn.STDEV.S(C{C0}:C{C1})', '0.000')]:
    ghi(ws, rr, 1, lab, bold=True); ghi(ws, rr, 2, f2, fill=XANH, fmt=fmt); ghi(ws, rr, 3, f3, fill=XANH, fmt=fmt); rr += 1
ghi(ws, rr, 1, 'CV = SD / trung bình', bold=True); ghi(ws, rr, 2, f'=B{rr - 1}/B{rr - 2}', fill=XAM, fmt='0.000'); ghi(ws, rr, 3, f'=C{rr - 1}/C{rr - 2}', fill=XAM, fmt='0.000'); rCV = rr; rr += 1
ghi(ws, rr, 1, 'SD của ln(EBI) 2024', bold=True); ghi(ws, rr, 4, f'=_xlfn.STDEV.S(D{C0}:D{C1})', fill=XAM, fmt='0.000'); rr += 1
ghi(ws, rr, 1, 'Kết luận', bold=True); ghi(ws, rr, 2, f'=IF(C{rCV}<B{rCV},"CV giảm từ "&TEXT(B{rCV},"0.00")&" xuống "&TEXT(C{rCV},"0.00")&" → có hội tụ sigma","CV không giảm → không hội tụ sigma")', fill=XANH); ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=6); rr += 2
rr = chu_thich(ws, rr, 'Trong bài báo (58 tỉnh): CV EBI giảm từ 1,253 (2021) xuống 0,608 (2024); SD ln(EBI) từ 0,528 xuống 0,410 – Mục 4.3.')
for col, wd in zip('ABCD', [30, 12, 12, 14]): ws.column_dimensions[col].width = wd

# ---------- Sheet 5: Theil + phân rã
ws = sheet('5_Theil_phan_ra'); rr = tieu_de(ws, 1, 'Chỉ số Theil-T và phân rã thành phần GIỮA nhóm + TRONG nhóm   (Bảng 3, Hình 3, Hình 4, phương trình (1))')
rr = chu_thich(ws, rr, 'T = (1/N)·Σ (y_i/ȳ)·ln(y_i/ȳ).   Phân rã: T = Σ_k s_k·T_k (trong nhóm) + Σ_k s_k·ln(ȳ_k/ȳ) (giữa nhóm), với s_k = (N_k/N)·(ȳ_k/ȳ).')
rr = header(ws, rr, ['Tỉnh', 'Nhóm', 'y = EBI 2024', 'y/ȳ', 'ln(y/ȳ)', '(y/ȳ)·ln(y/ȳ)'])
T0 = rr
for i, g in enumerate(goc):
    ghi(ws, rr, 1, g['tinh']); ghi(ws, rr, 2, g['nhom'], fill=VANG); ghi(ws, rr, 3, f'={DL}D{R0 + i}', fill=VANG, fmt='0.0')
    ghi(ws, rr, 4, f'=C{rr}/$C${rr + 99}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 5, f'=LN(D{rr})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 6, f'=D{rr}*E{rr}', fill=XANH, fmt='0.0000'); rr += 1
T1 = rr - 1
# thay placeholder 99 bằng ô trung bình thật
rMean = rr
ghi(ws, rr, 1, 'ȳ (trung bình chung)', bold=True); ghi(ws, rr, 3, f'=AVERAGE(C{T0}:C{T1})', fill=XANH, fmt='0.000'); rr += 1
for row in range(T0, T1 + 1):
    ws.cell(row=row, column=4).value = f'=C{row}/$C${rMean}'
ghi(ws, rr, 1, 'N', bold=True); ghi(ws, rr, 3, f'=COUNT(C{T0}:C{T1})', fill=XANH); rN = rr; rr += 1
ghi(ws, rr, 1, 'Theil-T = (1/N)·Σ cột F', bold=True); ghi(ws, rr, 6, f'=SUM(F{T0}:F{T1})/C{rN}', fill=XAM, fmt='0.0000'); rTT = rr; rr += 2
rr = tieu_de(ws, rr, 'Phân rã theo 2 nhóm', 11)
rr = header(ws, rr, ['Nhóm k', 'N_k', 'ȳ_k', 's_k = (N_k/N)(ȳ_k/ȳ)', 'T_k (Theil trong nhóm)', 's_k·T_k', 'ln(ȳ_k/ȳ)', 's_k·ln(ȳ_k/ȳ)'])
P0 = rr
for k in ['Lõi', 'Ngoại vi']:
    ghi(ws, rr, 1, k, fill=VANG); ghi(ws, rr, 2, f'=COUNTIF($B${T0}:$B${T1},A{rr})', fill=XANH)
    ghi(ws, rr, 3, f'=AVERAGEIF($B${T0}:$B${T1},A{rr},$C${T0}:$C${T1})', fill=XANH, fmt='0.000')
    ghi(ws, rr, 4, f'=B{rr}/$C${rN}*C{rr}/$C${rMean}', fill=XANH, fmt='0.0000')
    ghi(ws, rr, 5, f'=SUMPRODUCT(($B${T0}:$B${T1}=A{rr})*($C${T0}:$C${T1}/C{rr})*LN($C${T0}:$C${T1}/C{rr}))/B{rr}', fill=XANH, fmt='0.0000')
    ghi(ws, rr, 6, f'=D{rr}*E{rr}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 7, f'=LN(C{rr}/$C${rMean})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 8, f'=D{rr}*G{rr}', fill=XANH, fmt='0.0000'); rr += 1
P1 = rr - 1
ghi(ws, rr, 1, 'Tổng', bold=True); ghi(ws, rr, 6, f'=SUM(F{P0}:F{P1})', fill=XAM, fmt='0.0000'); ghi(ws, rr, 8, f'=SUM(H{P0}:H{P1})', fill=XAM, fmt='0.0000'); rTot = rr; rr += 1
ghi(ws, rr, 1, 'Phần TRONG nhóm (within)', bold=True); ghi(ws, rr, 6, f'=F{rTot}', fill=XAM, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 'Phần GIỮA nhóm (between)', bold=True); ghi(ws, rr, 6, f'=H{rTot}', fill=XAM, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 'Kiểm tra: within + between = Theil-T?', bold=True); ghi(ws, rr, 6, f'=F{rTot}+H{rTot}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 7, f'=IF(ABS(F{rr}-F{rTT})<0.00001,"✓ bằng nhau","✗ sai")', fill=XANH); rr += 1
ghi(ws, rr, 1, 'Tỷ trọng giữa nhóm (%) = between / T × 100', bold=True); ghi(ws, rr, 6, f'=H{rTot}/F{rTT}*100', fill=XAM, fmt='0.0'); rr += 2
rr = chu_thich(ws, rr, 'Trong bài báo (58 tỉnh, 6 vùng, 2024): Theil-T = 0,128; giữa vùng = 0,047; trong vùng = 0,081; tỷ trọng giữa vùng = 36,9% (Bảng 3). Ví dụ 2 nhóm ở đây cho tỷ trọng giữa nhóm cao hơn vì hai tỉnh lõi tách hẳn khỏi bốn tỉnh còn lại.')
for col, wd in zip('ABCDEFGH', [34, 12, 14, 20, 22, 12, 12, 16]): ws.column_dimensions[col].width = wd

# ---------- Sheet 6: Gini theo trụ cột
ws = sheet('6_Gini_tru_cot'); rr = tieu_de(ws, 1, 'Phân rã Gini theo nguồn: G(EBI) = Σ_k S_k·G_k·R_k   (Bảng 4)')
rr = chu_thich(ws, rr, 'EBI = 0,2·NNL&HT + 0,4·B2C + 0,4·B2B (trọng số từ 2021). S_k = tỷ trọng thành phần k trong EBI trung bình; G_k = Gini của thành phần w_k·X_k; R_k = cov(w_k·X_k, hạng EBI)/cov(w_k·X_k, hạng của chính nó).')
rr = header(ws, rr, ['Tỉnh', 'NNL&HT', 'B2C', 'B2B', '0,2·NNL&HT', '0,4·B2C', '0,4·B2B', 'EBI tái tạo = tổng', 'EBI công bố', 'Hạng EBI /N', 'Hạng 0,2·NNL /N', 'Hạng 0,4·B2C /N', 'Hạng 0,4·B2B /N'])
S0 = rr
for i, g in enumerate(goc):
    ghi(ws, rr, 1, g['tinh']); ghi(ws, rr, 2, f'={DL}E{R0 + i}', fill=VANG); ghi(ws, rr, 3, f'={DL}F{R0 + i}', fill=VANG); ghi(ws, rr, 4, f'={DL}G{R0 + i}', fill=VANG)
    ghi(ws, rr, 5, f'=0.2*B{rr}', fill=XANH, fmt='0.0'); ghi(ws, rr, 6, f'=0.4*C{rr}', fill=XANH, fmt='0.0'); ghi(ws, rr, 7, f'=0.4*D{rr}', fill=XANH, fmt='0.0')
    ghi(ws, rr, 8, f'=E{rr}+F{rr}+G{rr}', fill=XANH, fmt='0.0'); ghi(ws, rr, 9, f'={DL}D{R0 + i}', fill=VANG, fmt='0.0')
    ghi(ws, rr, 10, f'=_xlfn.RANK.AVG(H{rr},$H${S0}:$H${S0 + 5},1)/6', fill=XANH, fmt='0.000')
    ghi(ws, rr, 11, f'=_xlfn.RANK.AVG(E{rr},$E${S0}:$E${S0 + 5},1)/6', fill=XANH, fmt='0.000')
    ghi(ws, rr, 12, f'=_xlfn.RANK.AVG(F{rr},$F${S0}:$F${S0 + 5},1)/6', fill=XANH, fmt='0.000')
    ghi(ws, rr, 13, f'=_xlfn.RANK.AVG(G{rr},$G${S0}:$G${S0 + 5},1)/6', fill=XANH, fmt='0.000'); rr += 1
S1r = rr - 1; rr += 1
rr = header(ws, rr, ['Trụ cột k', 'S_k = TB(w·X)/TB(EBI tái tạo)', 'G_k = Gini(w·X)', 'R_k = COVAR(w·X, hạng EBI)/COVAR(w·X, hạng chính nó)', 'S_k·G_k·R_k', 'Đóng góp (%)'])
Q0 = rr
gini_xl = lambda rng: f'(2*SUMPRODUCT(_xlfn.RANK.AVG({rng},{rng},1),{rng})/(6*SUM({rng}))-(6+1)/6)'
for name, col, rk in [('NNL&HT', 'E', 'K'), ('B2C', 'F', 'L'), ('B2B', 'G', 'M')]:
    rng = f'${col}${S0}:${col}${S1r}'
    ghi(ws, rr, 1, name, fill=VANG); ghi(ws, rr, 2, f'=AVERAGE({rng})/AVERAGE($H${S0}:$H${S1r})', fill=XANH, fmt='0.0000')
    ghi(ws, rr, 3, f'={gini_xl(rng)}', fill=XANH, fmt='0.0000')
    ghi(ws, rr, 4, f'=_xlfn.COVARIANCE.S({rng},$J${S0}:$J${S1r})/_xlfn.COVARIANCE.S({rng},${rk}${S0}:${rk}${S1r})', fill=XANH, fmt='0.0000')
    ghi(ws, rr, 5, f'=B{rr}*C{rr}*D{rr}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 6, f'=E{rr}/SUM($E${Q0}:$E${Q0 + 2})*100', fill=XAM, fmt='0.0'); rr += 1
ghi(ws, rr, 1, 'Tổng Σ S·G·R', bold=True); ghi(ws, rr, 5, f'=SUM(E{Q0}:E{Q0 + 2})', fill=XAM, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 'Gini của EBI tái tạo (để đối chiếu)', bold=True); ghi(ws, rr, 5, f'={gini_xl(f"$H${S0}:$H${S1r}")}', fill=XANH, fmt='0.0000'); rr += 2
rr = chu_thich(ws, rr, 'Trong bài báo (2024, 58 tỉnh): Gini EBI 0,245; đóng góp NNL&HT 24,4%, B2C 29,8%, B2B 45,9% (Bảng 4). Lưu ý: hàm Gini trong Excel ở đây dùng RANK.AVG nên đúng khi không có giá trị trùng; trong ví dụ này không trùng.')
for i in range(1, 14): ws.column_dimensions[get_column_letter(i)].width = 13
ws.column_dimensions['A'].width = 30; ws.column_dimensions['D'].width = 40; ws.column_dimensions['B'].width = 26

# ---------- Sheet 7: Beta
ws = sheet('7_Beta_hoi_tu'); rr = tieu_de(ws, 1, 'Hội tụ beta tuyệt đối 2021→2024: g_i = (1/T)·ln(EBI_2024/EBI_2021) = α + β·ln(EBI_2021)   (Bảng 5, Hình 7, phương trình (2))')
rr = chu_thich(ws, rr, 'β < 0 nghĩa là tỉnh xuất phát thấp (ln EBI_2021 nhỏ) tăng nhanh hơn → hội tụ. Tốc độ hội tụ λ = −ln(1+β·T)/T; thời gian bán hội tụ = ln2/λ. T = 3 năm.')
rr = header(ws, rr, ['Tỉnh', 'EBI 2021 (y0)', 'EBI 2024 (y1)', 'x = ln(y0)', 'ln(y1)', 'g = (ln y1 − ln y0)/T', 'x − x̄', 'g − ḡ', '(x−x̄)(g−ḡ)', '(x−x̄)²'])
B0 = rr
for i, g in enumerate(goc):
    ghi(ws, rr, 1, g['tinh']); ghi(ws, rr, 2, f'={DL}H{R0 + i}', fill=VANG, fmt='0.00'); ghi(ws, rr, 3, f'={DL}D{R0 + i}', fill=VANG, fmt='0.0')
    ghi(ws, rr, 4, f'=LN(B{rr})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 5, f'=LN(C{rr})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 6, f'=(E{rr}-D{rr})/$B$16', fill=XANH, fmt='0.0000')
    ghi(ws, rr, 7, f'=D{rr}-$D$13', fill=XANH, fmt='0.0000'); ghi(ws, rr, 8, f'=F{rr}-$F$13', fill=XANH, fmt='0.0000'); ghi(ws, rr, 9, f'=G{rr}*H{rr}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 10, f'=G{rr}^2', fill=XANH, fmt='0.0000'); rr += 1
B1 = rr - 1
ws.cell(row=rr, column=1, value='(hàng trống)'); rr += 1
# rr == 13 ? bảo đảm: B0=5 -> 6 dòng -> B1=10, rr=11 sau trống -> 12. Ta ghi trung bình ở dòng 13 và T ở B16 cố định
rr = 13
ghi(ws, rr, 1, 'Trung bình', bold=True); ghi(ws, rr, 4, f'=AVERAGE(D{B0}:D{B1})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 6, f'=AVERAGE(F{B0}:F{B1})', fill=XANH, fmt='0.0000')
ghi(ws, rr, 9, f'=SUM(I{B0}:I{B1})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 10, f'=SUM(J{B0}:J{B1})', fill=XANH, fmt='0.0000'); rr += 3
ghi(ws, rr, 1, 'T (số năm)', bold=True); ghi(ws, rr, 2, 3, fill=VANG); rr += 1
ghi(ws, rr, 1, 'β = Σ(x−x̄)(g−ḡ) / Σ(x−x̄)²', bold=True); ghi(ws, rr, 2, f'=I13/J13', fill=XAM, fmt='0.0000'); ghi(ws, rr, 3, 'kiểm tra bằng SLOPE:'); ghi(ws, rr, 4, f'=SLOPE(F{B0}:F{B1},D{B0}:D{B1})', fill=XANH, fmt='0.0000'); rB = rr; rr += 1
ghi(ws, rr, 1, 'α = ḡ − β·x̄', bold=True); ghi(ws, rr, 2, f'=F13-B{rB}*D13', fill=XAM, fmt='0.0000'); ghi(ws, rr, 3, 'kiểm tra bằng INTERCEPT:'); ghi(ws, rr, 4, f'=INTERCEPT(F{B0}:F{B1},D{B0}:D{B1})', fill=XANH, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 'R²', bold=True); ghi(ws, rr, 2, f'=RSQ(F{B0}:F{B1},D{B0}:D{B1})', fill=XANH, fmt='0.000'); rr += 1
ghi(ws, rr, 1, 'λ = −ln(1+β·T)/T', bold=True); ghi(ws, rr, 2, f'=-LN(1+B{rB}*B16)/B16', fill=XAM, fmt='0.0000'); rL = rr; rr += 1
ghi(ws, rr, 1, 'Tốc độ hội tụ (%/năm) = λ×100', bold=True); ghi(ws, rr, 2, f'=B{rL}*100', fill=XAM, fmt='0.0'); rr += 1
ghi(ws, rr, 1, 'Thời gian bán hội tụ (năm) = ln2/λ', bold=True); ghi(ws, rr, 2, f'=LN(2)/B{rL}', fill=XAM, fmt='0.0'); rr += 1
ghi(ws, rr, 1, 'Kết luận', bold=True); ghi(ws, rr, 2, f'=IF(B{rB}<0,"β âm → tỉnh thấp tăng nhanh hơn → HỘI TỤ","β dương → PHÂN KỲ")', fill=XANH); ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=6); rr += 2
rr = chu_thich(ws, rr, 'Trong bài báo (56 tỉnh, 2021–2024): β = −0,0888 (SE 0,0180), λ = 10,3%/năm, bán hội tụ 6,7 năm, R² = 0,528 (Bảng 5, Hình 7). Bài dùng sai số chuẩn vững HC1 để tính p-value; Excel không có sẵn HC1 nên chỉ minh hoạ β, α, R².')
for col, wd in zip('ABCDEFGHIJ', [36, 14, 14, 12, 12, 20, 10, 10, 14, 10]): ws.column_dimensions[col].width = wd

# ---------- Sheet 8: OLS
ws = sheet('8_OLS'); rr = tieu_de(ws, 1, 'Hồi quy OLS đơn giản: ln(EBI) = b0 + b1·(DN/1.000 dân)   (Bảng 6 – mô hình đầy đủ có 8 biến + biến giả năm, phương trình (3))')
rr = chu_thich(ws, rr, 'Bài dùng 8 biến + hiệu ứng năm trên 331 quan sát tỉnh–năm. Ở đây minh hoạ với 1 biến, 6 tỉnh, năm 2024 để thấy cách OLS chọn đường thẳng "khớp nhất" (tổng bình phương phần dư nhỏ nhất) và cách đọc hệ số.')
rr = header(ws, rr, ['Tỉnh', 'x = DN/1.000 dân', 'y = ln(EBI 2024)', 'x − x̄', 'y − ȳ', '(x−x̄)(y−ȳ)', '(x−x̄)²', 'ŷ = b0 + b1·x', 'phần dư = y − ŷ', 'phần dư²'])
O0 = rr
for i, g in enumerate(goc):
    ghi(ws, rr, 1, g['tinh']); ghi(ws, rr, 2, f'={DL}I{R0 + i}', fill=VANG, fmt='0.00'); ghi(ws, rr, 3, f'=LN({DL}D{R0 + i})', fill=XANH, fmt='0.0000')
    ghi(ws, rr, 4, f'=B{rr}-$B$12', fill=XANH, fmt='0.000'); ghi(ws, rr, 5, f'=C{rr}-$C$12', fill=XANH, fmt='0.0000'); ghi(ws, rr, 6, f'=D{rr}*E{rr}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 7, f'=D{rr}^2', fill=XANH, fmt='0.000')
    ghi(ws, rr, 8, f'=$B$15+$B$14*B{rr}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 9, f'=C{rr}-H{rr}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 10, f'=I{rr}^2', fill=XANH, fmt='0.0000'); rr += 1
O1 = rr - 1
rr = 12
ghi(ws, rr, 1, 'Trung bình', bold=True); ghi(ws, rr, 2, f'=AVERAGE(B{O0}:B{O1})', fill=XANH, fmt='0.000'); ghi(ws, rr, 3, f'=AVERAGE(C{O0}:C{O1})', fill=XANH, fmt='0.0000')
ghi(ws, rr, 6, f'=SUM(F{O0}:F{O1})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 7, f'=SUM(G{O0}:G{O1})', fill=XANH, fmt='0.000'); ghi(ws, rr, 10, f'=SUM(J{O0}:J{O1})', fill=XANH, fmt='0.0000'); rr += 2
ghi(ws, rr, 1, 'b1 (độ dốc) = Σ(x−x̄)(y−ȳ)/Σ(x−x̄)²', bold=True); ghi(ws, rr, 2, '=F12/G12', fill=XAM, fmt='0.0000'); ghi(ws, rr, 3, 'kiểm tra SLOPE:'); ghi(ws, rr, 4, f'=SLOPE(C{O0}:C{O1},B{O0}:B{O1})', fill=XANH, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 'b0 (hằng số) = ȳ − b1·x̄', bold=True); ghi(ws, rr, 2, '=C12-B14*B12', fill=XAM, fmt='0.0000'); ghi(ws, rr, 3, 'kiểm tra INTERCEPT:'); ghi(ws, rr, 4, f'=INTERCEPT(C{O0}:C{O1},B{O0}:B{O1})', fill=XANH, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 'R² = 1 − Σphần dư²/Σ(y−ȳ)²', bold=True); ghi(ws, rr, 2, f'=1-J12/SUMSQ(E{O0}:E{O1})', fill=XAM, fmt='0.000'); ghi(ws, rr, 3, 'kiểm tra RSQ:'); ghi(ws, rr, 4, f'=RSQ(C{O0}:C{O1},B{O0}:B{O1})', fill=XANH, fmt='0.000'); rr += 1
ghi(ws, rr, 1, 'Sai số chuẩn thường của b1 (LINEST)', bold=True); ghi(ws, rr, 2, f'=INDEX(LINEST(C{O0}:C{O1},B{O0}:B{O1},TRUE,TRUE),2,1)', fill=XANH, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 't = b1 / SE', bold=True); ghi(ws, rr, 2, f'=B14/B17', fill=XANH, fmt='0.00'); rr += 1
ghi(ws, rr, 1, 'p-value (2 phía, df = N−2)', bold=True); ghi(ws, rr, 2, f'=_xlfn.T.DIST.2T(ABS(B18),COUNT(B{O0}:B{O1})-2)', fill=XAM, fmt='0.0000'); rr += 1
ghi(ws, rr, 1, 'Đọc hệ số: thêm 1 DN/1.000 dân → EBI tăng ≈ (e^b1 − 1)×100 %', bold=True); ghi(ws, rr, 2, '=(EXP(B14)-1)*100', fill=XAM, fmt='0.0'); rr += 2
rr = chu_thich(ws, rr, 'Trong bài báo (Bảng 6, cột 2): hệ số mật độ DN = 0,070 (SE cụm 0,007, p < 0,001) → mỗi DN thêm trên 1.000 dân gắn với EBI cao hơn ≈ 7,3%; R² = 0,936; VIF lớn nhất 5,37. Ví dụ 1 biến ở đây cho hệ số khác vì chưa kiểm soát 7 biến còn lại và hiệu ứng năm.')
for col, wd in zip('ABCDEFGHIJ', [44, 14, 14, 10, 10, 12, 10, 14, 14, 10]): ws.column_dimensions[col].width = wd

# ---------- Sheet 9: Oaxaca
ws = sheet('9_Oaxaca'); rr = tieu_de(ws, 1, 'Phân tách Oaxaca–Blinder: khoảng cách ln(EBI) lõi − ngoại vi = phần NGUỒN LỰC + phần HỆ SỐ   (Bảng 8, 8b, Hình 8, phương trình (4))')
rr = chu_thich(ws, rr, 'Ví dụ 1 biến (x = DN/1.000 dân), hệ số tham chiếu = hệ số của nhóm lõi (bài dùng hệ số gộp theo Jann 2008 – ý tưởng giống nhau). Nguồn lực = (x̄_L − x̄_N)·β_L; Hệ số = phần còn lại.')
rr = header(ws, rr, ['Tỉnh', 'Nhóm', 'x = DN/1.000 dân', 'y = ln(EBI 2024)'])
X0 = rr
for i, g in enumerate(goc):
    ghi(ws, rr, 1, g['tinh']); ghi(ws, rr, 2, g['nhom'], fill=VANG); ghi(ws, rr, 3, f'={DL}I{R0 + i}', fill=VANG, fmt='0.00'); ghi(ws, rr, 4, f'=LN({DL}D{R0 + i})', fill=XANH, fmt='0.0000'); rr += 1
X1 = rr - 1; rr += 1
rr = header(ws, rr, ['Đại lượng', 'Nhóm Lõi (L)', 'Nhóm Ngoại vi (N)', 'Ghi chú'])
rL = rr
ghi(ws, rr, 1, 'x̄ (trung bình mật độ DN)', bold=True); ghi(ws, rr, 2, f'=AVERAGEIF($B${X0}:$B${X1},"Lõi",$C${X0}:$C${X1})', fill=XANH, fmt='0.000'); ghi(ws, rr, 3, f'=AVERAGEIF($B${X0}:$B${X1},"Ngoại vi",$C${X0}:$C${X1})', fill=XANH, fmt='0.000'); rr += 1
ghi(ws, rr, 1, 'ȳ (trung bình ln EBI)', bold=True); ghi(ws, rr, 2, f'=AVERAGEIF($B${X0}:$B${X1},"Lõi",$D${X0}:$D${X1})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 3, f'=AVERAGEIF($B${X0}:$B${X1},"Ngoại vi",$D${X0}:$D${X1})', fill=XANH, fmt='0.0000'); rr += 1
# hồi quy riêng từng nhóm: Lõi = 2 dòng đầu (Hà Nội, Bình Dương), Ngoại vi = 4 dòng sau
ghi(ws, rr, 1, 'β (độ dốc) hồi quy riêng từng nhóm', bold=True); ghi(ws, rr, 2, f'=SLOPE(D{X0}:D{X0 + 1},C{X0}:C{X0 + 1})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 3, f'=SLOPE(D{X0 + 2}:D{X1},C{X0 + 2}:C{X1})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 4, 'Lõi chỉ có 2 tỉnh nên đường thẳng đi qua đúng 2 điểm (minh hoạ)'); rr += 1
ghi(ws, rr, 1, 'α (hằng số) hồi quy riêng từng nhóm', bold=True); ghi(ws, rr, 2, f'=INTERCEPT(D{X0}:D{X0 + 1},C{X0}:C{X0 + 1})', fill=XANH, fmt='0.0000'); ghi(ws, rr, 3, f'=INTERCEPT(D{X0 + 2}:D{X1},C{X0 + 2}:C{X1})', fill=XANH, fmt='0.0000'); rr += 2
ghi(ws, rr, 1, 'Khoảng cách = ȳ_L − ȳ_N', bold=True); ghi(ws, rr, 2, f'=B{rL + 1}-C{rL + 1}', fill=XAM, fmt='0.0000'); ghi(ws, rr, 4, f'=TEXT((EXP(B{rr})-1)*100,"0")&"% – EBI vùng lõi cao hơn ngoại vi chừng này (e^khoảng cách − 1)"', fill=XANH); rG = rr; rr += 1
ghi(ws, rr, 1, 'Phần NGUỒN LỰC = (x̄_L − x̄_N) × β_L', bold=True); ghi(ws, rr, 2, f'=(B{rL}-C{rL})*B{rL + 2}', fill=XAM, fmt='0.0000'); ghi(ws, rr, 4, '"nếu ngoại vi có mật độ DN như lõi (và vẫn quy đổi theo hệ số của lõi) thì khoảng cách thu hẹp bấy nhiêu"'); rE = rr; rr += 1
ghi(ws, rr, 1, 'Phần HỆ SỐ = khoảng cách − phần nguồn lực', bold=True); ghi(ws, rr, 2, f'=B{rG}-B{rE}', fill=XAM, fmt='0.0000'); ghi(ws, rr, 4, '"cùng mật độ DN nhưng chuyển hoá thành EBI khác nhau (hằng số + độ dốc khác)"'); rC = rr; rr += 1
ghi(ws, rr, 1, '% nguồn lực', bold=True); ghi(ws, rr, 2, f'=B{rE}/B{rG}*100', fill=XAM, fmt='0.0'); rr += 1
ghi(ws, rr, 1, '% hệ số', bold=True); ghi(ws, rr, 2, f'=B{rC}/B{rG}*100', fill=XAM, fmt='0.0'); rr += 1
ghi(ws, rr, 1, 'Kiểm tra: nguồn lực + hệ số = khoảng cách', bold=True); ghi(ws, rr, 2, f'=B{rE}+B{rC}', fill=XANH, fmt='0.0000'); ghi(ws, rr, 3, f'=IF(ABS(B{rr}-B{rG})<0.00001,"✓","✗")', fill=XANH); rr += 2
rr = chu_thich(ws, rr, 'Trong bài báo (331 quan sát 2019–2024, 8 biến + biến giả năm, hệ số gộp): khoảng cách 0,455 (≈ 58% EBI); nguồn lực 0,360 = 79,1%; hệ số 0,095 = 20,9%; riêng mật độ DN giải thích 69,4% khoảng cách (Bảng 8, 8b, Hình 8).')
for col, wd in zip('ABCD', [40, 16, 18, 70]): ws.column_dimensions[col].width = wd

wb.save(OUT_XLSX)
print('Đã ghi', OUT_XLSX)

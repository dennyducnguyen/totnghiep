# -*- coding: utf-8 -*-
"""14 - Dựng SỔ TÍNH TOÁN ĐẦY ĐỦ, TOÀN QUY MÔ của bài báo: SO-TINH-TOAN-DAY-DU-BAI-BAO.xlsx

Nguyên tắc:
  * 1 sheet dữ liệu gốc (504 dòng tỉnh–năm) + mỗi bảng/hình của bài một sheet.
  * Mọi thứ Excel tính được là CÔNG THỨC SỐNG (Gini, Theil, MLD, CV, EBI tương đối, phân rã, hội tụ beta qua
    SLOPE/LINEST, hệ số OLS qua LINEST, tương quan…), tham chiếu ngược về sheet dữ liệu.
  * Những gì Excel không có (SE cụm theo tỉnh, p-value vững, VIF, Wald, FE tỉnh, Oaxaca bootstrap) được DÁN GIÁ TRỊ
    từ kết quả script 01 (thư mục ket-qua/), tô màu cam và ghi rõ.
  * Sheet cuối 'Doi_chieu' so từng con số Excel với Python để chứng minh khớp.
Chạy: python scripts/14_so_tinh_day_du.py   (sau khi đã chạy 01_phan_tich.py)
"""
import json, os
import numpy as np, pandas as pd
import statsmodels.api as sm
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, ScatterChart, Reference, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KQ = os.path.join(ROOT, 'ket-qua')
OUT = os.path.join(ROOT, 'SO-TINH-TOAN-DAY-DU-BAI-BAO.xlsx')

REG = ['ĐNB', 'ĐBSH', 'BTB&DHMT', 'ĐBSCL', 'TN', 'TDMNPB']
REG_FULL = {'ĐNB': 'Đông Nam Bộ', 'ĐBSH': 'Đồng bằng sông Hồng', 'BTB&DHMT': 'Bắc Trung Bộ và Duyên hải miền Trung',
            'ĐBSCL': 'Đồng bằng sông Cửu Long', 'TN': 'Tây Nguyên', 'TDMNPB': 'Trung du và miền núi phía Bắc'}
YEARS = list(range(2017, 2025))
W = {2017: dict(NNL_HT=.30, B2C=.30, B2B=.30, G2B=.10), 2018: dict(NNL_HT=.20, B2C=.35, B2B=.35, G2B=.10),
     2019: dict(NNL_HT=.20, B2C=.35, B2B=.35, G2B=.10), 2020: dict(NNL_HT=.20, B2C=.35, B2B=.35, G2B=.10),
     2021: dict(NNL_HT=.20, B2C=.40, B2B=.40), 2022: dict(NNL_HT=.20, B2C=.40, B2B=.40),
     2023: dict(NNL_HT=.20, B2C=.40, B2B=.40), 2024: dict(NNL_HT=.20, B2C=.40, B2B=.40)}
X1 = ['ln_grdp', 'dn_per_1000', 'urban_share_pct', 'trained_labor_pct', 'ln_freight', 'modern_retail_100k', 'ln_fdi_pc']
X2 = X1 + ['hh_computer_pct']
TEN_X = {'ln_grdp': 'ln(GRDP/người)', 'dn_per_1000': 'Mật độ DN (DN/1.000 dân)', 'urban_share_pct': 'Tỷ lệ dân số thành thị (%)',
         'trained_labor_pct': 'Tỷ lệ LĐ qua đào tạo (%)', 'ln_freight': 'ln(Hàng hóa vận chuyển đường bộ/người)',
         'modern_retail_100k': 'Siêu thị + TTTM/100.000 dân', 'ln_fdi_pc': 'ln(1 + FDI lũy kế/người)', 'hh_computer_pct': 'Tỷ lệ hộ có máy tính (%)',
         'internet_user_pct': 'Tỷ lệ người sử dụng Internet (%)', 'it_skills_pct': 'Tỷ lệ người có kỹ năng CNTT (%)'}

# ====================================================================== DỮ LIỆU & KẾT QUẢ PYTHON
df = pd.read_csv(os.path.join(KQ, 'panel_phan_tich.csv')).sort_values(['nam', 'tinh']).reset_index(drop=True)
assert len(df) == 504
kq = json.load(open(os.path.join(KQ, 'ket_qua.json'), encoding='utf-8'))
b3 = pd.read_csv(os.path.join(KQ, 'bang3_theil.csv')); b4 = pd.read_csv(os.path.join(KQ, 'bang4_gini_tru_cot.csv'))
b5 = pd.read_csv(os.path.join(KQ, 'bang5b_beta.csv')); b6 = pd.read_csv(os.path.join(KQ, 'bang6_ols.csv'))
vif = pd.read_csv(os.path.join(KQ, 'bang6_vif.csv')); b7 = pd.read_csv(os.path.join(KQ, 'bang7_vung.csv'))
b9 = pd.read_csv(os.path.join(KQ, 'bang6b_2024_cross.csv')); b8 = pd.read_csv(os.path.join(KQ, 'bang8_oaxaca.csv'))
b8d = pd.read_csv(os.path.join(KQ, 'bang8_oaxaca_chi_tiet.csv'), index_col=0)
ex2 = pd.read_csv(os.path.join(KQ, 'phu_theil_bo_HN_HCM.csv'))
dig = pd.read_excel(os.path.join(ROOT, 'du-lieu-dau-vao', 'EBI_NGTK_Panel_2017_2024.xlsx'), sheet_name='Digital_2024_cross')

d = df[df.EBI.notna()].copy()                       # 445 dòng có EBI
ROW1 = {}                                            # (nam, tinh) -> dòng trong 1_Du_lieu
R1_0 = 5
for i, r in df.iterrows():
    ROW1[(int(r.nam), r.tinh)] = R1_0 + i
YB1 = {y: (R1_0 + 63 * k, R1_0 + 63 * k + 62) for k, y in enumerate(YEARS)}   # khối năm trong 1_Du_lieu

# ====================================================================== ĐỊNH DẠNG
wb = Workbook()
F_IN = PatternFill('solid', fgColor='FFF2CC'); F_FX = PatternFill('solid', fgColor='E2EFDA'); F_RES = PatternFill('solid', fgColor='EDEDED')
F_PY = PatternFill('solid', fgColor='FCE4D6'); F_HD = PatternFill('solid', fgColor='1F4E79'); F_SUB = PatternFill('solid', fgColor='D9E2F3')
BR = Border(*[Side(style='thin', color='BFBFBF')] * 4)
FN_HD = Font(bold=True, color='FFFFFF'); FN_B = Font(bold=True); FN_T = Font(bold=True, size=13); FN_NOTE = Font(italic=True, color='595959')


def put(ws, r, c, v, kind=None, fmt=None, bold=False, wrap=False):
    cell = ws.cell(row=r, column=c, value=v)
    if kind == 'in': cell.fill = F_IN
    elif kind == 'fx': cell.fill = F_FX
    elif kind == 'res': cell.fill = F_RES
    elif kind == 'py': cell.fill = F_PY
    elif kind == 'hd': cell.fill = F_HD; cell.font = FN_HD; cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    elif kind == 'sub': cell.fill = F_SUB; cell.font = FN_B
    if kind and kind != 'txt': cell.border = BR
    if fmt: cell.number_format = fmt
    if bold: cell.font = FN_B
    if wrap: cell.alignment = Alignment(wrap_text=True, vertical='top')
    return cell


def hdr(ws, r, cols, c0=1):
    for i, h in enumerate(cols): put(ws, r, c0 + i, h, 'hd')
    ws.row_dimensions[r].height = 32
    return r + 1


def title(ws, r, t, note=None):
    ws.cell(row=r, column=1, value=t).font = FN_T; r += 1
    if note:
        c = ws.cell(row=r, column=1, value=note); c.font = FN_NOTE; c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); ws.row_dimensions[r].height = 44; r += 1
    return r + 1


def widths(ws, ws_list):
    for i, w in enumerate(ws_list, 1): ws.column_dimensions[L(i)].width = w


def axes(chart):
    """openpyxl 3.1: trục bị ẩn mặc định trong Excel 365 -> bật lại, hiện nhãn số."""
    chart.x_axis.delete = False; chart.y_axis.delete = False
    chart.x_axis.majorTickMark = 'out'; chart.y_axis.majorTickMark = 'out'
    return chart


def stars(pcell):
    return f'=IF({pcell}="","",IF({pcell}<0.01,"***",IF({pcell}<0.05,"**",IF({pcell}<0.1,"*",""))))'


A = lambda col, r0, r1: f'${col}${r0}:${col}${r1}'          # dải cột tuyệt đối
DL = "'1_Du_lieu'!"
DOI_CHIEU = []   # (nhóm, tên, ô Excel, giá trị Python)

# ====================================================================== 0. HƯỚNG DẪN
ws0 = wb.active; ws0.title = '0_Huong_dan'
r = title(ws0, 1, 'SỔ TÍNH TOÁN ĐẦY ĐỦ – TOÀN BỘ SỐ LIỆU CỦA BÀI BÁO "KHOẢNG CÁCH TMĐT GIỮA 63 TỈNH, 2017–2024"')
for t in [
    'Đây là bản "làm thật": mọi con số trong bài đều truy được về sheet 1_Du_lieu qua công thức, trừ các thống kê Excel không có hàm (đánh dấu màu cam).',
    'MÀU Ô:  VÀNG = số liệu gốc nhập vào (VECOM, Niên giám);  XANH LÁ = công thức Excel (bấm vào ô để xem);  XÁM = kết quả cuối được dẫn trong bài;  CAM = giá trị dán từ script 01_phan_tich.py (Excel không có hàm: sai số chuẩn cụm theo tỉnh, p-value vững HC1, VIF, kiểm định Wald, hiệu ứng cố định tỉnh, Oaxaca–Blinder bootstrap).',
    'Sheet cuối "11_Doi_chieu" so sánh từng con số Excel tính được với Python: cột "Chênh lệch" phải ≈ 0.',
    'Số trong Excel dùng dấu chấm thập phân (định dạng máy); bài báo in dấu phẩy. Các số trong bài đã làm tròn 3–4 chữ số.',
    'Hàng dữ liệu được sắp theo NĂM rồi TỈNH (63 tỉnh mỗi năm, cùng thứ tự) nên "năm trước" của một dòng luôn nằm đúng 63 dòng phía trên – các biến trễ (_L1) dựa vào quy tắc đó.',
]:
    c = ws0.cell(row=r, column=1, value=t); c.alignment = Alignment(wrap_text=True, vertical='top'); ws0.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8); ws0.row_dimensions[r].height = 46; r += 1
r += 1
r = hdr(ws0, r, ['Sheet', 'Nội dung', 'Tương ứng trong bài báo', 'Loại tính toán'])
for row in [
    ['1_Du_lieu', '504 dòng tỉnh–năm: EBI + 4 trụ cột (VECOM), biến Niên giám, biến dẫn xuất ln(…), biến trễ 1 năm', 'Bảng 1 (mô tả biến), Mục 3.1–3.2', 'Nhập + công thức dẫn xuất'],
    ['2_Bang1_Mo_ta', 'Thống kê mô tả 16 biến', 'Bảng 1', 'Công thức'],
    ['3_Bang2_EBI_tuong_doi', 'EBI tương đối từng tỉnh–năm; trung bình vùng theo giai đoạn; ma trận vùng × năm; biểu đồ', 'Bảng 2, Hình 2, Mục 3.1, 4.1', 'Công thức + biểu đồ'],
    ['4_Bang3_Gini_Theil', 'Gini, CV, SD ln, Theil-T, MLD và phân rã giữa/trong 6 vùng từng năm; kiểm tra bỏ HN & TP.HCM; biểu đồ', 'Bảng 3, Hình 3, Hình 4, phương trình (1)', 'Công thức + biểu đồ'],
    ['5_Bang4_Gini_tru_cot', 'Phân rã Gini theo trụ cột S·G·R từng năm', 'Bảng 4, Mục 4.2', 'Công thức'],
    ['6_Bang5a_Sigma', 'Hệ số biến thiên EBI và 3 trụ cột theo năm; biểu đồ', 'Hình 5, Mục 4.3', 'Công thức + biểu đồ'],
    ['7_Bang5_Beta', 'Hội tụ beta tuyệt đối và có điều kiện 3 cửa sổ; λ, bán hội tụ; biểu đồ điểm + đường xu hướng', 'Bảng 5, 5b, Hình 6, 7, phương trình (2)', 'Công thức (LINEST) + dán SE vững, bảng động'],
    ['8_Bang6_7_9_OLS', 'Hồi quy OLS gộp M1–M4, hệ số chuẩn hóa, ma trận tương quan, biến giả vùng M5–M6, cắt ngang 2024', 'Bảng 6, 7, 9, phương trình (3)', 'Công thức (LINEST) + dán SE cụm, p, VIF, Wald, FE tỉnh'],
    ['9_Bang8_Oaxaca', 'Phân tách Oaxaca–Blinder lõi – ngoại vi: trung bình nhóm, phần nguồn lực theo biến, tổng; biểu đồ', 'Bảng 8, 8b, Hình 8, phương trình (4)', 'Công thức (trung bình, tích) + dán β*, bootstrap'],
    ['10_So_lieu_khac', 'Top/bottom 5 tỉnh 2024, tỷ số max/trung vị, P90/P10, lõi – ngoại vi theo năm, tỉnh thiếu số liệu', 'Mục 3.1, 4.1, 4.3', 'Công thức'],
    ['11_Doi_chieu', 'Đối chiếu Excel ↔ Python cho mọi con số then chốt', '—', 'Kiểm tra'],
]:
    for i, v in enumerate(row): put(ws0, r, i + 1, v, 'txt', wrap=True)
    r += 1
widths(ws0, [24, 80, 34, 34])

# ====================================================================== 1. DỮ LIỆU
ws1 = wb.create_sheet('1_Du_lieu')
title(ws1, 1, 'DỮ LIỆU GỐC 63 TỈNH × 8 NĂM (sắp theo năm, rồi tỉnh)', 'Cột vàng: nhập từ VECOM (EBI, trụ cột) và Niên giám Thống kê. Cột xanh: dẫn xuất bằng công thức. Ô EBI trống = VECOM không xếp hạng tỉnh đó năm đó.')
COLS1 = [('tinh', 'Tỉnh', 'in'), ('nam', 'Năm', 'in'), ('vung6', 'Vùng (6 vùng NGTK)', 'in'), ('core', 'Lõi=1 (ĐNB, ĐBSH)', 'fx'),
         ('EBI', 'EBI', 'in'), ('NNL_HT', 'NNL&HT', 'in'), ('B2C', 'B2C', 'in'), ('B2B', 'B2B', 'in'), ('G2B', 'G2B (≤2020)', 'in'),
         ('pop_th', 'Dân số (nghìn người)', 'in'), ('grdp_pc_tr', 'GRDP/người (tr.đ)', 'in'), ('dn_per_1000', 'DN/1.000 dân', 'in'),
         ('urban_share_pct', '% thành thị', 'in'), ('trained_labor_pct', '% LĐ qua đào tạo', 'in'), ('hh_computer_pct', '% hộ có máy tính', 'in'),
         ('freight_road_kt', 'Hàng hóa đường bộ (nghìn tấn)', 'in'), ('freight_pc', 'Hàng hóa/người (tấn)', 'fx'), ('sieuthi', 'Siêu thị', 'in'),
         ('tttm', 'TTTM', 'in'), ('modern_retail_100k', 'ST+TTTM/100.000 dân', 'fx'), ('fdi_pc_usd', 'FDI lũy kế/người (USD)', 'in'),
         ('ln_ebi', 'ln(EBI)', 'fx'), ('ln_grdp', 'ln(GRDP/người)', 'fx'), ('ln_freight', 'ln(hàng hóa/người)', 'fx'), ('ln_fdi_pc', 'ln(1+FDI/người)', 'fx')]
LAGS = ['ln_grdp', 'dn_per_1000', 'urban_share_pct', 'trained_labor_pct', 'hh_computer_pct', 'ln_freight', 'modern_retail_100k', 'ln_fdi_pc']
C1 = {k: i + 1 for i, (k, _, _) in enumerate(COLS1)}
for j, v in enumerate(LAGS): C1[v + '_L1'] = len(COLS1) + 1 + j
hdr(ws1, 4, [h for _, h, _ in COLS1] + [f'{TEN_X[v]} (t−1)' for v in LAGS])
CL = lambda k: L(C1[k])
for i, rw in df.iterrows():
    r = R1_0 + i
    for k, _, kind in COLS1:
        c = C1[k]
        if kind == 'in':
            v = rw[k]; put(ws1, r, c, (None if pd.isna(v) else (int(v) if k == 'nam' else v)), 'in')
    put(ws1, r, C1['core'], f'=IF(OR({CL("vung6")}{r}="ĐNB",{CL("vung6")}{r}="ĐBSH"),1,0)', 'fx')
    put(ws1, r, C1['freight_pc'], f'=IF({CL("freight_road_kt")}{r}="","",{CL("freight_road_kt")}{r}/{CL("pop_th")}{r})', 'fx', '0.00')
    put(ws1, r, C1['modern_retail_100k'], f'=({CL("sieuthi")}{r}+{CL("tttm")}{r})/{CL("pop_th")}{r}*100', 'fx', '0.00')
    put(ws1, r, C1['ln_ebi'], f'=IF({CL("EBI")}{r}="","",LN({CL("EBI")}{r}))', 'fx', '0.0000')
    put(ws1, r, C1['ln_grdp'], f'=IF({CL("grdp_pc_tr")}{r}="","",LN({CL("grdp_pc_tr")}{r}))', 'fx', '0.0000')
    put(ws1, r, C1['ln_freight'], f'=IF({CL("freight_pc")}{r}="","",LN({CL("freight_pc")}{r}))', 'fx', '0.0000')
    put(ws1, r, C1['ln_fdi_pc'], f'=LN(1+{CL("fdi_pc_usd")}{r})', 'fx', '0.0000')
    for v in LAGS:
        src = L(C1[v]); put(ws1, r, C1[v + '_L1'], f'=IF({CL("nam")}{r}=2017,"",{src}{r - 63})', 'fx', '0.0000')
ws1.freeze_panes = 'C5'
widths(ws1, [20, 6, 11, 8, 7, 8, 7, 7, 8, 10, 10, 10, 9, 9, 9, 11, 10, 8, 7, 11, 11, 9, 9, 9, 9] + [11] * len(LAGS))
R1_END = R1_0 + 503

# ====================================================================== 2. BẢNG 1
ws2 = wb.create_sheet('2_Bang1_Mo_ta')
r = title(ws2, 1, 'BẢNG 1. Thống kê mô tả (gộp 2017–2024; N = số ô không trống)')
r = hdr(ws2, r, ['Biến', 'Cột trong 1_Du_lieu', 'N', 'Trung bình', 'Độ lệch chuẩn', 'Nhỏ nhất', 'Lớn nhất'])
B1_VARS = [('EBI', 'Chỉ số TMĐT tổng hợp (EBI)'), ('NNL_HT', 'Trụ cột NNL&HT'), ('B2C', 'Trụ cột B2C'), ('B2B', 'Trụ cột B2B'), ('G2B', 'Trụ cột G2B (2017–2020)'),
           ('grdp_pc_tr', 'GRDP bình quân đầu người (tr.đ)'), ('dn_per_1000', 'Mật độ doanh nghiệp (DN/1.000 dân)'), ('urban_share_pct', 'Tỷ lệ dân số thành thị (%)'),
           ('trained_labor_pct', 'Tỷ lệ LĐ 15+ qua đào tạo (%)'), ('hh_computer_pct', 'Tỷ lệ hộ có máy tính (%)'), ('freight_pc', 'Hàng hóa đường bộ/người (tấn)'),
           ('modern_retail_100k', 'Siêu thị + TTTM/100.000 dân'), ('fdi_pc_usd', 'FDI lũy kế/người (USD)')]
for k, ten in B1_VARS:
    rg = f"{DL}{A(CL(k), R1_0, R1_END)}"
    put(ws2, r, 1, ten, 'txt'); put(ws2, r, 2, CL(k), 'txt')
    put(ws2, r, 3, f'=COUNT({rg})', 'fx'); put(ws2, r, 4, f'=AVERAGE({rg})', 'fx', '0.00'); put(ws2, r, 5, f'=_xlfn.STDEV.S({rg})', 'fx', '0.00')
    put(ws2, r, 6, f'=MIN({rg})', 'fx', '0.00'); put(ws2, r, 7, f'=MAX({rg})', 'fx', '0.00')
    if k == 'EBI': DOI_CHIEU += [('Bảng 1', 'N quan sát EBI', f"'2_Bang1_Mo_ta'!C{r}", 445), ('Bảng 1', 'EBI trung bình', f"'2_Bang1_Mo_ta'!D{r}", round(float(d.EBI.mean()), 4))]
    r += 1
widths(ws2, [40, 12, 8, 12, 14, 10, 10])

# ====================================================================== 3. BẢNG 2 + HÌNH 2
ws3 = wb.create_sheet('3_Bang2_EBI_tuong_doi')
r = title(ws3, 1, 'EBI TƯƠNG ĐỐI = EBI / trung bình các tỉnh có số liệu trong năm × 100', 'Trung bình năm t tính bằng AVERAGE trên khối 63 dòng của năm t trong 1_Du_lieu (ô trống bị bỏ qua, nên mẫu số là số tỉnh có EBI).')
r = hdr(ws3, r, ['Tỉnh', 'Năm', 'Vùng', 'Giai đoạn', 'EBI', 'Trung bình năm (N tỉnh có EBI)', 'EBI tương đối', 'NNL&HT tương đối', 'B2C tương đối', 'B2B tương đối', 'Lõi'])
R3_0 = r; ROW3 = {}
for _, rw in d.iterrows():
    y, t = int(rw.nam), rw.tinh; r1 = ROW1[(y, t)]; a, b = YB1[y]
    ROW3[(y, t)] = r
    put(ws3, r, 1, t, 'txt'); put(ws3, r, 2, y, 'txt'); put(ws3, r, 3, f'={DL}{CL("vung6")}{r1}', 'fx'); put(ws3, r, 4, f'=IF(B{r}<=2020,"2017–2020","2021–2024")', 'fx')
    put(ws3, r, 5, f'={DL}{CL("EBI")}{r1}', 'fx', '0.0'); put(ws3, r, 6, f'=AVERAGE({DL}{A(CL("EBI"), a, b)})', 'fx', '0.000')
    put(ws3, r, 7, f'=E{r}/F{r}*100', 'fx', '0.0')
    for j, p in enumerate(['NNL_HT', 'B2C', 'B2B']):
        put(ws3, r, 8 + j, f'={DL}{CL(p)}{r1}/AVERAGE({DL}{A(CL(p), a, b)})*100', 'fx', '0.0')
    put(ws3, r, 11, f'={DL}{CL("core")}{r1}', 'fx')
    r += 1
R3_1 = r - 1
# Bảng 2
r += 2; put(ws3, r, 1, 'BẢNG 2. EBI tương đối theo 6 vùng (trung bình giản đơn các tỉnh–năm trong vùng)', 'txt', bold=True); r += 1
r = hdr(ws3, r, ['Vùng', 'Số tỉnh', 'EBI 2017–2020', 'EBI 2021–2024', 'EBI 2024', 'NNL&HT 2024', 'B2C 2024', 'B2B 2024'])
RB2_0 = r
for k in REG:
    put(ws3, r, 1, REG_FULL[k], 'txt'); put(ws3, r, 2, f'=COUNTIFS({A("C", R3_0, R3_1)},"{k}",{A("B", R3_0, R3_1)},2024)', 'fx')
    put(ws3, r, 3, f'=AVERAGEIFS({A("G", R3_0, R3_1)},{A("C", R3_0, R3_1)},"{k}",{A("D", R3_0, R3_1)},"2017–2020")', 'fx', '0.0')
    put(ws3, r, 4, f'=AVERAGEIFS({A("G", R3_0, R3_1)},{A("C", R3_0, R3_1)},"{k}",{A("D", R3_0, R3_1)},"2021–2024")', 'fx', '0.0')
    for j, col in enumerate('GHIJ'):
        put(ws3, r, 5 + j, f'=AVERAGEIFS({A(col, R3_0, R3_1)},{A("C", R3_0, R3_1)},"{k}",{A("B", R3_0, R3_1)},2024)', 'fx', '0.0')
    r += 1
RB2_1 = r - 1
put(ws3, r, 1, 'Tỷ số vùng cao nhất / vùng thấp nhất', 'txt', bold=True)
for c in range(3, 9): put(ws3, r, c, f'=MAX({L(c)}{RB2_0}:{L(c)}{RB2_1})/MIN({L(c)}{RB2_0}:{L(c)}{RB2_1})', 'res', '0.00')
b2 = pd.read_csv(os.path.join(KQ, 'bang2_vung.csv'), index_col=0)
DOI_CHIEU += [('Bảng 2', 'ĐNB 2021–2024', f"'3_Bang2_EBI_tuong_doi'!D{RB2_0}", round(float(b2.loc['ĐNB', '2021–2024']), 3)),
              ('Bảng 2', 'TDMNPB EBI 2024', f"'3_Bang2_EBI_tuong_doi'!E{RB2_0 + 5}", round(float(b2.loc['TDMNPB', 'EBI_rel']), 3)),
              ('Bảng 2', 'Tỷ số cao/thấp B2B 2024', f"'3_Bang2_EBI_tuong_doi'!H{r}", round(float(b2['B2B_rel'].max() / b2['B2B_rel'].min()), 3))]
# ma trận vùng × năm cho Hình 2
r += 3; put(ws3, r, 1, 'EBI tương đối trung bình vùng theo từng năm (dữ liệu Hình 2)', 'txt', bold=True); r += 1
r = hdr(ws3, r, ['Vùng'] + [str(y) for y in YEARS])
RH2_0 = r
for k in REG:
    put(ws3, r, 1, REG_FULL[k], 'txt')
    for j, y in enumerate(YEARS):
        put(ws3, r, 2 + j, f'=AVERAGEIFS({A("G", R3_0, R3_1)},{A("C", R3_0, R3_1)},"{k}",{A("B", R3_0, R3_1)},{y})', 'fx', '0.0')
    r += 1
RH2_1 = r - 1
ch = LineChart(); ch.title = 'Hình 2. EBI tương đối theo 6 vùng, 2017–2024 (trung bình = 100)'; ch.height = 9; ch.width = 18
ch.add_data(Reference(ws3, min_col=1, max_col=9, min_row=RH2_0, max_row=RH2_1), from_rows=True, titles_from_data=True)
ch.set_categories(Reference(ws3, min_col=2, max_col=9, min_row=RH2_0 - 1)); ch.y_axis.title = 'EBI tương đối'; ch.x_axis.title = 'Năm'
for sr in ch.series: sr.smooth = False
ws3.add_chart(axes(ch), f'K{RB2_0 - 2}')
ws3.freeze_panes = f'A{R3_0}'; widths(ws3, [20, 6, 11, 11, 8, 14, 12, 12, 12, 12, 6])

# ====================================================================== 4. BẢNG 3 + HÌNH 3, 4
ws4 = wb.create_sheet('4_Bang3_Gini_Theil')
r = title(ws4, 1, 'BẢNG 3. Gini, CV, Theil-T, MLD và phân rã theo 6 vùng – tính riêng từng năm trên các tỉnh có EBI',
          'Mỗi dòng = một tỉnh–năm (445 dòng, sắp theo năm). Cột "j" = thứ hạng trong năm (RANK.AVG, tăng dần). Gini = 2Σ(j·y)/(N·Σy) − (N+1)/N. Theil-T = Σ(y/ȳ)ln(y/ȳ)/N; MLD = Σln(ȳ/y)/N. Cột "f" = 1 nếu là Hà Nội hoặc TP.HCM (dùng cho kiểm tra vững).')
HD4 = ['Tỉnh', 'Năm', 'Vùng', 'y = EBI', 'f (HN/HCM)', 'ȳ năm', 'N năm', 'j = hạng', 'j·y', 'y/ȳ', 'ln(y/ȳ)', '(y/ȳ)ln(y/ȳ)', 'ln(ȳ/y)  [MLD]', 'ln(y)', 'ȳ_ex (bỏ HN,HCM)', '(1−f)(y/ȳ_ex)ln(y/ȳ_ex)']
r = hdr(ws4, r, HD4); R4_0 = r
YB4 = {}
for y in YEARS:
    dy = d[d.nam == y]; a = r
    for _, rw in dy.iterrows():
        r1 = ROW1[(y, rw.tinh)]
        put(ws4, r, 1, rw.tinh, 'txt'); put(ws4, r, 2, y, 'txt'); put(ws4, r, 3, f'={DL}{CL("vung6")}{r1}', 'fx'); put(ws4, r, 4, f'={DL}{CL("EBI")}{r1}', 'fx', '0.0')
        put(ws4, r, 5, f'=IF(OR(A{r}="Hà Nội",A{r}="TP. Hồ Chí Minh"),1,0)', 'fx')
        r += 1
    b = r - 1; YB4[y] = (a, b)
    for rr in range(a, b + 1):
        put(ws4, rr, 6, f'=AVERAGE({A("D", a, b)})', 'fx', '0.000'); put(ws4, rr, 7, f'=COUNT({A("D", a, b)})', 'fx')
        put(ws4, rr, 8, f'=_xlfn.RANK.AVG(D{rr},{A("D", a, b)},1)', 'fx', '0.0'); put(ws4, rr, 9, f'=H{rr}*D{rr}', 'fx', '0.0')
        put(ws4, rr, 10, f'=D{rr}/F{rr}', 'fx', '0.0000'); put(ws4, rr, 11, f'=LN(J{rr})', 'fx', '0.0000'); put(ws4, rr, 12, f'=J{rr}*K{rr}', 'fx', '0.0000')
        put(ws4, rr, 13, f'=LN(F{rr}/D{rr})', 'fx', '0.0000'); put(ws4, rr, 14, f'=LN(D{rr})', 'fx', '0.0000')
        put(ws4, rr, 15, f'=SUMPRODUCT((1-{A("E", a, b)})*{A("D", a, b)})/({A("G", a, b)}-2)'.replace(f'/({A("G", a, b)}-2)', f'/(G{rr}-SUM({A("E", a, b)}))'), 'fx', '0.000')
        put(ws4, rr, 16, f'=(1-E{rr})*(D{rr}/O{rr})*LN(D{rr}/O{rr})', 'fx', '0.0000')
R4_1 = r - 1
# --- bảng vùng × năm (phân rã)
r += 2; put(ws4, r, 1, 'PHÂN RÃ THEO VÙNG (từng năm): s_k = (N_k/N)(ȳ_k/ȳ); T_k = Theil trong vùng; phần trong = s_k·T_k; phần giữa = s_k·ln(ȳ_k/ȳ). MLD: p_k = N_k/N; L_k; giữa = p_k·ln(ȳ/ȳ_k). Cột cuối: bỏ HN & TP.HCM.', 'txt', bold=True); r += 1
r = hdr(ws4, r, ['Năm', 'Vùng', 'N_k', 'ȳ_k', 's_k', 'T_k', 's_k·T_k (trong)', 'ln(ȳ_k/ȳ)', 's_k·ln(ȳ_k/ȳ) (giữa)', 'p_k', 'L_k (MLD vùng)', 'p_k·L_k', 'p_k·ln(ȳ/ȳ_k) (giữa MLD)', 'N_k,ex', 'ȳ_k,ex', 's_k,ex·ln(ȳ_k,ex/ȳ_ex)'])
RV = {}
for y in YEARS:
    a, b = YB4[y]; V = lambda col: A(col, a, b)
    for k in REG:
        RV[(y, k)] = r
        put(ws4, r, 1, y, 'txt'); put(ws4, r, 2, k, 'txt')
        put(ws4, r, 3, f'=COUNTIFS({V("C")},B{r})', 'fx'); put(ws4, r, 4, f'=AVERAGEIFS({V("D")},{V("C")},B{r})', 'fx', '0.000')
        put(ws4, r, 5, f'=C{r}/G{a}*D{r}/F{a}', 'fx', '0.0000')
        put(ws4, r, 6, f'=SUMPRODUCT(({V("C")}=B{r})*({V("D")}/D{r})*LN({V("D")}/D{r}))/C{r}', 'fx', '0.0000')
        put(ws4, r, 7, f'=E{r}*F{r}', 'fx', '0.0000'); put(ws4, r, 8, f'=LN(D{r}/F{a})', 'fx', '0.0000'); put(ws4, r, 9, f'=E{r}*H{r}', 'fx', '0.0000')
        put(ws4, r, 10, f'=C{r}/G{a}', 'fx', '0.0000'); put(ws4, r, 11, f'=SUMPRODUCT(({V("C")}=B{r})*LN(D{r}/{V("D")}))/C{r}', 'fx', '0.0000')
        put(ws4, r, 12, f'=J{r}*K{r}', 'fx', '0.0000'); put(ws4, r, 13, f'=J{r}*LN(F{a}/D{r})', 'fx', '0.0000')
        put(ws4, r, 14, f'=COUNTIFS({V("C")},B{r},{V("E")},0)', 'fx'); put(ws4, r, 15, f'=AVERAGEIFS({V("D")},{V("C")},B{r},{V("E")},0)', 'fx', '0.000')
        put(ws4, r, 16, f'=(N{r}/(G{a}-SUM({V("E")})))*(O{r}/O{a})*LN(O{r}/O{a})', 'fx', '0.0000')
        r += 1
RV_1 = r - 1
# --- Bảng 3 tổng hợp
r += 2; put(ws4, r, 1, 'BẢNG 3 (tổng hợp theo năm)', 'txt', bold=True); r += 1
r = hdr(ws4, r, ['Năm', 'N', 'Gini', 'CV', 'SD ln(EBI)', 'Theil-T', 'Giữa các vùng', 'Trong vùng', 'Tỷ trọng giữa vùng (%)', 'MLD', 'MLD giữa vùng', 'Tỷ trọng giữa vùng MLD (%)', 'Theil-T bỏ HN&HCM', 'Giữa vùng bỏ HN&HCM', 'Tỷ trọng giữa vùng bỏ HN&HCM (%)', 'Gini bỏ HN&HCM'])
RB3 = {}
for y in YEARS:
    a, b = YB4[y]; va, vb = RV[(y, REG[0])], RV[(y, REG[-1])]; RB3[y] = r
    put(ws4, r, 1, y, 'txt'); put(ws4, r, 2, f'=G{a}', 'fx')
    put(ws4, r, 3, f'=2*SUM({A("I", a, b)})/(B{r}*SUM({A("D", a, b)}))-(B{r}+1)/B{r}', 'res', '0.0000')
    put(ws4, r, 4, f'=_xlfn.STDEV.S({A("D", a, b)})/F{a}', 'res', '0.0000'); put(ws4, r, 5, f'=_xlfn.STDEV.S({A("N", a, b)})', 'res', '0.0000')
    put(ws4, r, 6, f'=SUM({A("L", a, b)})/B{r}', 'res', '0.0000'); put(ws4, r, 7, f'=SUM({A("I", va, vb)})', 'res', '0.0000'); put(ws4, r, 8, f'=SUM({A("G", va, vb)})', 'res', '0.0000')
    put(ws4, r, 9, f'=G{r}/F{r}*100', 'res', '0.0'); put(ws4, r, 10, f'=SUM({A("M", a, b)})/B{r}', 'res', '0.0000'); put(ws4, r, 11, f'=SUM({A("M", va, vb)})', 'res', '0.0000')
    put(ws4, r, 12, f'=K{r}/J{r}*100', 'res', '0.0'); put(ws4, r, 13, f'=SUM({A("P", a, b)})/(B{r}-SUM({A("E", a, b)}))', 'res', '0.0000')
    put(ws4, r, 14, f'=SUM({A("P", va, vb)})', 'res', '0.0000'); put(ws4, r, 15, f'=N{r}/M{r}*100', 'res', '0.0')
    # Gini bỏ HN & HCM: rank trong tập con -> dùng SUMPRODUCT trên các tỉnh f=0 với hạng tính lại
    put(ws4, r, 16, f'=2*SUMPRODUCT((1-{A("E", a, b)})*(COUNTIFS({A("D", a, b)},"<"&{A("D", a, b)},{A("E", a, b)},0)+(COUNTIFS({A("D", a, b)},{A("D", a, b)},{A("E", a, b)},0)+1)/2)*{A("D", a, b)})/((B{r}-2)*SUMPRODUCT((1-{A("E", a, b)})*{A("D", a, b)}))-(B{r}-1)/(B{r}-2)', 'res', '0.0000')
    row3 = b3[b3.nam == y].iloc[0]; rowx = ex2[ex2.nam == y].iloc[0]
    DOI_CHIEU += [('Bảng 3', f'Gini {y}', f"'4_Bang3_Gini_Theil'!C{r}", round(float(row3.Gini), 4)), ('Bảng 3', f'Theil-T {y}', f"'4_Bang3_Gini_Theil'!F{r}", round(float(row3.Theil), 4)),
                  ('Bảng 3', f'Tỷ trọng giữa vùng {y} (%)', f"'4_Bang3_Gini_Theil'!I{r}", round(float(row3.Theil_between_share * 100), 2)),
                  ('Bảng 3', f'Tỷ trọng giữa vùng MLD {y} (%)', f"'4_Bang3_Gini_Theil'!L{r}", round(float(row3.MLD_between_share * 100), 2)),
                  ('Bảng 3', f'CV {y}', f"'4_Bang3_Gini_Theil'!D{r}", round(float(row3.CV), 4)), ('Bảng 3', f'Theil bỏ HN&HCM {y}', f"'4_Bang3_Gini_Theil'!M{r}", round(float(rowx.Theil_ex2), 4)),
                  ('Bảng 3', f'Gini bỏ HN&HCM {y}', f"'4_Bang3_Gini_Theil'!P{r}", round(float(rowx.Gini_ex2), 4))]
    r += 1
RB3_0, RB3_1 = RB3[2017], RB3[2024]
c1 = BarChart(); c1.type = 'col'; c1.grouping = 'stacked'; c1.overlap = 100; c1.title = 'Hình 3. Theil-T và hai thành phần giữa vùng / trong vùng'; c1.height = 9; c1.width = 16
c1.add_data(Reference(ws4, min_col=7, max_col=8, min_row=RB3_0 - 1, max_row=RB3_1), titles_from_data=True); c1.set_categories(Reference(ws4, min_col=1, min_row=RB3_0, max_row=RB3_1))
c1.y_axis.title = 'Theil-T'; ws4.add_chart(axes(c1), f'A{RB3_1 + 3}')
c2 = LineChart(); c2.title = 'Hình 4. Tỷ trọng giữa các vùng trong tổng bất bình đẳng (%)'; c2.height = 9; c2.width = 16
c2.add_data(Reference(ws4, min_col=9, min_row=RB3_0 - 1, max_row=RB3_1), titles_from_data=True); c2.add_data(Reference(ws4, min_col=12, min_row=RB3_0 - 1, max_row=RB3_1), titles_from_data=True)
c2.set_categories(Reference(ws4, min_col=1, min_row=RB3_0, max_row=RB3_1))
for sr in c2.series: sr.smooth = False
c2.y_axis.scaling.min = 0; c2.y_axis.scaling.max = 60; ws4.add_chart(axes(c2), f'J{RB3_1 + 3}')
ws4.freeze_panes = f'A{R4_0}'; widths(ws4, [20, 6, 10, 9, 8, 9, 7, 8, 9, 9, 9, 11, 11, 9, 12, 14])

# ====================================================================== 5. BẢNG 4
ws5 = wb.create_sheet('5_Bang4_Gini_tru_cot')
r = title(ws5, 1, 'BẢNG 4. Phân rã Gini theo trụ cột: G(EBI) = Σ S_k·G_k·R_k', 'Y = Σ w_k·X_k (EBI tái tạo từ trọng số phục hồi). S_k = TB(w_k X_k)/TB(Y); G_k = Gini của w_k X_k; R_k = COVARIANCE.S(w_k X_k, hạng Y/N) / COVARIANCE.S(w_k X_k, hạng w_k X_k/N). Đóng góp % = S·G·R / Σ(S·G·R). Trọng số theo năm ở bảng nhỏ bên phải.')
put(ws5, r - 1, 28, 'Trọng số', 'sub'); put(ws5, r - 1, 29, 'NNL&HT', 'sub'); put(ws5, r - 1, 30, 'B2C', 'sub'); put(ws5, r - 1, 31, 'B2B', 'sub'); put(ws5, r - 1, 32, 'G2B', 'sub')
RW = {}
for i, y in enumerate(YEARS):
    rr = r + i; RW[y] = rr; put(ws5, rr, 28, y, 'txt')
    for j, p in enumerate(['NNL_HT', 'B2C', 'B2B', 'G2B']): put(ws5, rr, 29 + j, W[y].get(p, 0), 'in', '0.00')
r = hdr(ws5, r, ['Tỉnh', 'Năm', 'NNL&HT', 'B2C', 'B2B', 'G2B', 'w·NNL&HT', 'w·B2C', 'w·B2B', 'w·G2B', 'Y tái tạo', 'hạng Y/N', 'hạng w·NNL/N', 'hạng w·B2C/N', 'hạng w·B2B/N', 'hạng w·G2B/N'])
R5_0 = r; YB5 = {}
for y in YEARS:
    cols = list(W[y].keys()); dy = d[d.nam == y].dropna(subset=cols); a = r
    for _, rw in dy.iterrows():
        r1 = ROW1[(y, rw.tinh)]
        put(ws5, r, 1, rw.tinh, 'txt'); put(ws5, r, 2, y, 'txt')
        for j, p in enumerate(['NNL_HT', 'B2C', 'B2B', 'G2B']):
            put(ws5, r, 3 + j, f'={DL}{CL(p)}{r1}' if p in cols else 0, 'fx' if p in cols else 'txt', '0.0')
            put(ws5, r, 7 + j, f'={L(3 + j)}{r}*${L(29 + j)}${RW[y]}', 'fx', '0.00')
        put(ws5, r, 11, f'=SUM(G{r}:J{r})', 'fx', '0.00')
        r += 1
    b = r - 1; YB5[y] = (a, b, len(dy), cols)
    for rr in range(a, b + 1):
        put(ws5, rr, 12, f'=_xlfn.RANK.AVG(K{rr},{A("K", a, b)},1)/{len(dy)}', 'fx', '0.000')
        for j in range(4):
            col = L(7 + j)
            put(ws5, rr, 13 + j, (f'=_xlfn.RANK.AVG({col}{rr},{A(col, a, b)},1)/{len(dy)}' if ['NNL_HT', 'B2C', 'B2B', 'G2B'][j] in cols else ''), 'fx' if ['NNL_HT', 'B2C', 'B2B', 'G2B'][j] in cols else 'txt', '0.000')
R5_1 = r - 1
r += 2; put(ws5, r, 1, 'BẢNG 4 (tổng hợp theo năm)', 'txt', bold=True); r += 1
HD5 = ['Năm', 'N', 'Gini Y (tái tạo)', 'Gini EBI công bố'] + [f'{s} {p}' for p in ['NNL&HT', 'B2C', 'B2B', 'G2B'] for s in ['S', 'G', 'R', 'S·G·R']] + ['Σ S·G·R'] + [f'Đóng góp {p} (%)' for p in ['NNL&HT', 'B2C', 'B2B', 'G2B']]
r = hdr(ws5, r, HD5)
gini_rng = lambda vals, ranks, n: f'(2*SUMPRODUCT({ranks}*{n},{vals})/({n}*SUM({vals}))-({n}+1)/{n})'
for y in YEARS:
    a, b, n, cols = YB5[y]; r4a, r4b = YB4[y]
    put(ws5, r, 1, y, 'txt'); put(ws5, r, 2, n, 'txt'); put(ws5, r, 3, gini_rng(A("K", a, b), A("L", a, b), n).strip('()').join(['=', '']), 'res', '0.0000')
    put(ws5, r, 4, f"='4_Bang3_Gini_Theil'!C{RB3[y]}", 'fx', '0.0000')
    for j, p in enumerate(['NNL_HT', 'B2C', 'B2B', 'G2B']):
        c0 = 5 + 4 * j; vcol = L(7 + j); rcol = L(13 + j)
        if p in cols:
            put(ws5, r, c0, f'=AVERAGE({A(vcol, a, b)})/AVERAGE({A("K", a, b)})', 'fx', '0.0000')
            put(ws5, r, c0 + 1, '=' + gini_rng(A(vcol, a, b), A(rcol, a, b), n), 'fx', '0.0000')
            put(ws5, r, c0 + 2, f'=_xlfn.COVARIANCE.S({A(vcol, a, b)},{A("L", a, b)})/_xlfn.COVARIANCE.S({A(vcol, a, b)},{A(rcol, a, b)})', 'fx', '0.0000')
            put(ws5, r, c0 + 3, f'={L(c0)}{r}*{L(c0 + 1)}{r}*{L(c0 + 2)}{r}', 'fx', '0.0000')
        else:
            for q in range(4): put(ws5, r, c0 + q, '', 'txt')
    put(ws5, r, 21, f'=SUM(H{r},L{r},P{r},T{r})', 'res', '0.0000')
    for j, p in enumerate(['NNL_HT', 'B2C', 'B2B', 'G2B']):
        put(ws5, r, 22 + j, (f'={L(8 + 4 * j)}{r}/U{r}*100' if p in cols else ''), 'res' if p in cols else 'txt', '0.0')
    rb = b4[b4.nam == y].iloc[0]
    DOI_CHIEU += [('Bảng 4', f'Gini Y tái tạo {y}', f"'5_Bang4_Gini_tru_cot'!C{r}", round(float(rb.Gini_EBI), 4)),
                  ('Bảng 4', f'Đóng góp B2B {y} (%)', f"'5_Bang4_Gini_tru_cot'!X{r}", round(float(rb.share_B2B * 100), 2)),
                  ('Bảng 4', f'Gini NNL&HT {y}', f"'5_Bang4_Gini_tru_cot'!F{r}", round(float(rb.Gini_NNL_HT), 4))]
    r += 1
ws5.freeze_panes = f'A{R5_0}'; widths(ws5, [20, 6, 8, 8, 8, 8, 9, 9, 9, 9, 9, 9, 10, 10, 10, 10, 10, 10, 10, 10, 10, 12, 12, 12, 12])

# ====================================================================== 6. BẢNG 5a / HÌNH 5
ws6 = wb.create_sheet('6_Bang5a_Sigma')
r = title(ws6, 1, 'HỘI TỤ SIGMA: hệ số biến thiên CV = STDEV.S / AVERAGE theo năm (Hình 5)', 'Tính trên khối 63 dòng của mỗi năm trong 1_Du_lieu; ô trống được bỏ qua.')
r = hdr(ws6, r, ['Năm', 'N (EBI)', 'CV EBI', 'CV NNL&HT', 'CV B2C', 'CV B2B', 'SD ln(EBI)', 'Trung bình EBI', 'SD EBI'])
R6_0 = r; sig = pd.read_csv(os.path.join(KQ, 'bang5a_sigma.csv'))
for y in YEARS:
    a, b = YB1[y]
    put(ws6, r, 1, y, 'txt'); put(ws6, r, 2, f'=COUNT({DL}{A(CL("EBI"), a, b)})', 'fx')
    for j, p in enumerate(['EBI', 'NNL_HT', 'B2C', 'B2B']):
        rg = f'{DL}{A(CL(p), a, b)}'; put(ws6, r, 3 + j, f'=_xlfn.STDEV.S({rg})/AVERAGE({rg})', 'res', '0.0000')
    put(ws6, r, 7, f"='4_Bang3_Gini_Theil'!E{RB3[y]}", 'fx', '0.0000')
    put(ws6, r, 8, f'=AVERAGE({DL}{A(CL("EBI"), a, b)})', 'fx', '0.000'); put(ws6, r, 9, f'=_xlfn.STDEV.S({DL}{A(CL("EBI"), a, b)})', 'fx', '0.000')
    rs = sig[sig.nam == y].iloc[0]
    DOI_CHIEU += [('Hình 5', f'CV B2C {y}', f"'6_Bang5a_Sigma'!E{r}", round(float(rs.CV_B2C), 4)), ('Hình 5', f'CV NNL&HT {y}', f"'6_Bang5a_Sigma'!D{r}", round(float(rs.CV_NNL_HT), 4))]
    r += 1
c = LineChart(); c.title = 'Hình 5. Hội tụ sigma: hệ số biến thiên EBI và 3 trụ cột'; c.height = 9; c.width = 16
c.add_data(Reference(ws6, min_col=3, max_col=6, min_row=R6_0 - 1, max_row=r - 1), titles_from_data=True); c.set_categories(Reference(ws6, min_col=1, min_row=R6_0, max_row=r - 1))
for sr in c.series: sr.smooth = False
ws6.add_chart(axes(c), f'A{r + 2}')
widths(ws6, [8, 8, 10, 11, 10, 10, 11, 13, 10])

# ====================================================================== 7. BẢNG 5 / HÌNH 6, 7
ws7 = wb.create_sheet('7_Bang5_Beta')
r = title(ws7, 1, 'BẢNG 5. Hội tụ beta: g_i = (1/T)·ln(EBI_T/EBI_0) = α + β·ln(EBI_0) [+ γ′Z_0]', 'β, α, R² và SE cổ điển bằng SLOPE/INTERCEPT/RSQ/LINEST. SE vững HC1 và p-value (dùng trong bài) dán từ Python (ô cam). λ = −ln(1+βT)/T; bán hội tụ = ln2/λ. Biến điều kiện Z lấy năm gốc (GRDP: 2018 cho gốc 2017). EBI tương đối lấy từ sheet 3.')
COND = ['ln_grdp', 'dn_per_1000', 'urban_share_pct', 'trained_labor_pct']
WIN = [(2017, 2020, 'EBI', 'Tuyệt đối, 2017–2020 (thang cũ)'), (2021, 2024, 'EBI', 'Tuyệt đối, 2021–2024 (thang mới)'), (2017, 2024, 'EBI_rel', 'Tuyệt đối, 2017–2024 (EBI tương đối)')]
SUMROWS = []
for (y0, y1, ycol, ten) in WIN:
    T = y1 - y0; cy = 2018 if y0 == 2017 else y0
    a0 = d[d.nam == y0].set_index('tinh')[ycol]; a1 = d[d.nam == y1].set_index('tinh')[ycol]
    both = sorted(set(a0.index) & set(a1.index))
    z = d[d.nam == cy].set_index('tinh')[COND]; cond_ok = [t for t in both if (t in z.index) and not z.loc[t].isna().any()]
    put(ws7, r, 1, ten + f'  (T = {T})', 'sub'); r += 1
    r = hdr(ws7, r, ['Tỉnh', f'y0 = {ycol} {y0}', f'y1 = {ycol} {y1}', 'x = ln y0', 'ln y1', 'g = (ln y1 − ln y0)/T', 'Có đủ Z?'] + [f'{TEN_X[c]} ({cy})' for c in COND])
    a = r
    for t in both:
        if ycol == 'EBI':
            put(ws7, r, 2, f'={DL}{CL("EBI")}{ROW1[(y0, t)]}', 'fx', '0.00'); put(ws7, r, 3, f'={DL}{CL("EBI")}{ROW1[(y1, t)]}', 'fx', '0.00')
        else:
            put(ws7, r, 2, f"='3_Bang2_EBI_tuong_doi'!G{ROW3[(y0, t)]}", 'fx', '0.00'); put(ws7, r, 3, f"='3_Bang2_EBI_tuong_doi'!G{ROW3[(y1, t)]}", 'fx', '0.00')
        put(ws7, r, 1, t, 'txt'); put(ws7, r, 4, f'=LN(B{r})', 'fx', '0.0000'); put(ws7, r, 5, f'=LN(C{r})', 'fx', '0.0000'); put(ws7, r, 6, f'=(E{r}-D{r})/{T}', 'fx', '0.0000')
        put(ws7, r, 7, 1 if t in cond_ok else 0, 'txt')
        for j, c in enumerate(COND): put(ws7, r, 8 + j, f'={DL}{CL(c)}{ROW1[(cy, t)]}', 'fx', '0.000')
        r += 1
    b = r - 1
    # khối có điều kiện: các dòng có đủ Z (liên tục) -> tạo bản sao gọn ở cột N.. để LINEST không gặp ô trống
    cr = a
    for t in cond_ok:
        src = a + both.index(t)
        put(ws7, cr, 14, f'=F{src}', 'fx', '0.0000'); put(ws7, cr, 15, f'=D{src}', 'fx', '0.0000')
        for j in range(4): put(ws7, cr, 16 + j, f'={L(8 + j)}{src}', 'fx', '0.000')
        cr += 1
    cb = cr - 1
    put(ws7, a - 1, 14, 'g (đủ Z)', 'hd'); put(ws7, a - 1, 15, 'x = ln y0', 'hd')
    for j, c in enumerate(COND): put(ws7, a - 1, 16 + j, TEN_X[c], 'hd')
    r += 1
    rows_abs = b5[(b5.period == f'{y0}–{y1}') & (b5.model == 'Tuyệt đối') & (b5.ycol == ycol)].iloc[0]
    rows_cond = b5[(b5.period == f'{y0}–{y1}') & (b5.model == 'Có điều kiện') & (b5.ycol == ycol)].iloc[0]
    put(ws7, r, 1, 'KẾT QUẢ', 'sub'); put(ws7, r, 2, 'Tuyệt đối', 'sub'); put(ws7, r, 3, 'Có điều kiện', 'sub'); put(ws7, r, 4, 'Ghi chú', 'sub'); r += 1
    gx, xx = A("F", a, b), A("D", a, b); gc, Xc = A("N", a, cb), f'$O${a}:$S${cb}'
    lines = [('N', f'=COUNT({gx})', f'=COUNT({gc})', ''),
             ('β (hệ số ln y0)', f'=SLOPE({gx},{xx})', f'=INDEX(LINEST({gc},{Xc},TRUE,TRUE),1,5)', 'SLOPE / LINEST (cột 5 = ln y0 vì LINEST trả hệ số theo thứ tự ngược)'),
             ('α (hằng số)', f'=INTERCEPT({gx},{xx})', f'=INDEX(LINEST({gc},{Xc},TRUE,TRUE),1,6)', ''),
             ('SE cổ điển của β', f'=INDEX(LINEST({gx},{xx},TRUE,TRUE),2,1)', f'=INDEX(LINEST({gc},{Xc},TRUE,TRUE),2,5)', 'giả định phương sai sai số đồng đều'),
             ('SE vững HC1 (bài dùng)', float(rows_abs.se), float(rows_cond.se), 'DÁN từ Python – Excel không có'),
             ('p-value (HC1)', float(rows_abs.p), float(rows_cond.p), 'DÁN từ Python'),
             ('R²', f'=RSQ({gx},{xx})', f'=INDEX(LINEST({gc},{Xc},TRUE,TRUE),3,1)', ''),
             ('λ = −ln(1+β·T)/T', None, None, 'chỉ khi 1+βT > 0 và λ > 0; nếu β > 0 ghi —'),
             ('Tốc độ hội tụ (%/năm)', None, None, ''), ('Thời gian bán hội tụ (năm)', None, None, 'ln2/λ')]
    rN = r
    for i, (lab, fa, fc, note) in enumerate(lines):
        rr = r + i; put(ws7, rr, 1, lab, 'txt', bold=True)
        if lab.startswith('λ'):
            put(ws7, rr, 2, f'=IF(1+B{rN + 1}*{T}>0,IF(-LN(1+B{rN + 1}*{T})/{T}>0,-LN(1+B{rN + 1}*{T})/{T},"—"),"—")', 'res', '0.0000'); put(ws7, rr, 3, f'=IF(1+C{rN + 1}*{T}>0,IF(-LN(1+C{rN + 1}*{T})/{T}>0,-LN(1+C{rN + 1}*{T})/{T},"—"),"—")', 'res', '0.0000')
        elif lab.startswith('Tốc'):
            put(ws7, rr, 2, f'=IF(ISNUMBER(B{rr - 1}),B{rr - 1}*100,"—")', 'res', '0.0'); put(ws7, rr, 3, f'=IF(ISNUMBER(C{rr - 1}),C{rr - 1}*100,"—")', 'res', '0.0')
        elif lab.startswith('Thời'):
            put(ws7, rr, 2, f'=IF(ISNUMBER(B{rr - 2}),LN(2)/B{rr - 2},"—")', 'res', '0.0'); put(ws7, rr, 3, f'=IF(ISNUMBER(C{rr - 2}),LN(2)/C{rr - 2},"—")', 'res', '0.0')
        else:
            kind = 'py' if isinstance(fa, float) else ('res' if lab.startswith(('β', 'R²')) else 'fx')
            put(ws7, rr, 2, fa, kind, '0.0000'); put(ws7, rr, 3, fc, kind, '0.0000')
        put(ws7, rr, 4, note, 'txt')
    # hệ số biến điều kiện (Bảng 5b)
    rr = r + len(lines); put(ws7, rr, 1, 'Hệ số các biến điều kiện (Bảng 5b):', 'txt', bold=True); rr += 1
    for j, c in enumerate(COND):
        put(ws7, rr, 1, TEN_X[c], 'txt'); put(ws7, rr, 3, f'=INDEX(LINEST({gc},{Xc},TRUE,TRUE),1,{4 - j})', 'res', '0.0000')
        put(ws7, rr, 4, f'p (HC1) = {float(rows_cond[f"p_{c}"]):.4f}  [DÁN]', 'py'); rr += 1
    SUMROWS.append((ten, rN, a, b, y0, y1))
    DOI_CHIEU += [('Bảng 5', f'β tuyệt đối {y0}–{y1} ({ycol})', f"'7_Bang5_Beta'!B{rN + 1}", round(float(rows_abs.beta), 4)),
                  ('Bảng 5', f'β có điều kiện {y0}–{y1} ({ycol})', f"'7_Bang5_Beta'!C{rN + 1}", round(float(rows_cond.beta), 4)),
                  ('Bảng 5', f'R² tuyệt đối {y0}–{y1}', f"'7_Bang5_Beta'!B{rN + 6}", round(float(rows_abs.R2), 4))]
    if not np.isnan(rows_abs.lam) and rows_abs.lam > 0:
        DOI_CHIEU += [('Bảng 5', f'λ tuyệt đối {y0}–{y1}', f"'7_Bang5_Beta'!B{rN + 7}", round(float(rows_abs.lam), 4)), ('Bảng 5', f'Bán hội tụ {y0}–{y1}', f"'7_Bang5_Beta'!B{rN + 9}", round(float(rows_abs.half_life), 2))]
    r = rr + 2
# bảng động
bp = kq['beta_panel']
put(ws7, r, 1, 'Bảng động (tăng trưởng hằng năm EBI tương đối, hiệu ứng năm, loại bước nhảy 2020→2021) – DÁN từ Python (hồi quy với biến giả năm và SE cụm):', 'txt', bold=True); r += 1
for lab, v in [('β', bp['beta']), ('SE cụm theo tỉnh', bp['se']), ('p', bp['p']), ('N', bp['N']), ('λ = −ln(1+β)', -np.log(1 + bp['beta'])), ('Bán hội tụ', np.log(2) / -np.log(1 + bp['beta']))]:
    put(ws7, r, 1, lab, 'txt'); put(ws7, r, 2, float(v), 'py', '0.0000'); r += 1
# Hình 6, 7
for (ten, rN, a, b, y0, y1), anchor in zip(SUMROWS[:2], ['V4', 'V26']):
    sc = ScatterChart(); sc.title = f'Hình {6 if y0 == 2017 else 7}. Hội tụ beta tuyệt đối {y0}–{y1}: g theo ln(EBI {y0})'; sc.height = 10; sc.width = 16
    s = Series(Reference(ws7, min_col=6, min_row=a, max_row=b), Reference(ws7, min_col=4, min_row=a, max_row=b), title=f'{b - a + 1} tỉnh')
    s.marker.symbol = 'circle'; s.marker.size = 6; s.graphicalProperties.line.noFill = True; s.marker.graphicalProperties.solidFill = '2E75B6'; s.marker.graphicalProperties.line.solidFill = '1F4E79'; s.trendline = Trendline(trendlineType='linear', dispEq=True, dispRSqr=True)
    sc.varyColors = False
    sc.series.append(s); sc.x_axis.title = f'ln(EBI năm {y0})'; sc.y_axis.title = 'Tăng trưởng bình quân năm (log)'; ws7.add_chart(axes(sc), anchor)
widths(ws7, [30, 13, 13, 12, 10, 16, 8, 12, 12, 12, 12, 3, 3, 10, 10, 10, 10, 10, 10])

# ====================================================================== 8. BẢNG 6, 7, 9
ws8 = wb.create_sheet('8_Bang6_7_9_OLS')
r = title(ws8, 1, 'BẢNG 6, 7, 9. Hồi quy OLS: ln(EBI) = β0 + β′X + δ_năm (+ vùng)', 'Hệ số, SE cổ điển và R² từ LINEST (công thức). SE cụm theo tỉnh, p-value, VIF, Wald, hiệu ứng cố định tỉnh: DÁN từ Python (ô cam). LINEST trả hệ số theo thứ tự ngược với cột X, nên hệ số của cột thứ j (j = 1..m) = INDEX(LINEST(...),1,m−j+1); hằng số ở cột m+1; SE ở hàng 2; R² ở INDEX(...,3,1).')


def khoi_mau(ws, r, ten, dm, Xcols, dummy_years, extra_cols=None):
    """Ghi khối dữ liệu mẫu: cột tỉnh, năm, y=ln_ebi, X..., biến giả năm, (biến giả vùng). Trả về (a, b, cột đầu X, số cột X)."""
    put(ws, r, 1, ten, 'sub'); r += 1
    heads = ['Tỉnh', 'Năm', 'y = ln(EBI)'] + [TEN_X.get(c, c) for c in Xcols] + [f'D{y}' for y in dummy_years] + ([f'D_{k}' for k in extra_cols] if extra_cols else [])
    r = hdr(ws, r, heads); a = r
    for _, rw in dm.iterrows():
        r1 = ROW1[(int(rw.nam), rw.tinh)]
        put(ws, r, 1, rw.tinh, 'txt'); put(ws, r, 2, int(rw.nam), 'txt'); put(ws, r, 3, f'={DL}{CL("ln_ebi")}{r1}', 'fx', '0.0000')
        for j, c in enumerate(Xcols): put(ws, r, 4 + j, f'={DL}{L(C1[c])}{r1}', 'fx', '0.000')
        for j, y in enumerate(dummy_years): put(ws, r, 4 + len(Xcols) + j, f'=IF(B{r}={y},1,0)', 'fx')
        if extra_cols:
            for j, k in enumerate(extra_cols): put(ws, r, 4 + len(Xcols) + len(dummy_years) + j, f'=IF({DL}{CL("vung6")}{r1}="{k}",1,0)', 'fx')
        r += 1
    return a, r - 1, 4, len(Xcols) + len(dummy_years) + (len(extra_cols) if extra_cols else 0)


def bang_he_so(ws, r, ten, a, b, c0, m, names, pyrows, note=''):
    """Bảng hệ số: LINEST + dán SE cụm, p, sao."""
    put(ws, r, 1, ten, 'sub'); put(ws, r, 2, note, 'txt'); r += 1
    r = hdr(ws, r, ['Biến', 'Hệ số (LINEST)', 'SE cổ điển (LINEST)', 'SE cụm theo tỉnh [DÁN]', 'p-value cụm [DÁN]', 'Dấu sao', 'Hệ số Python [DÁN – đối chiếu]'])
    yr, Xr = A("C", a, b), f'${L(c0)}${a}:${L(c0 + m - 1)}${b}'
    r0 = r
    for j, (nm, term) in enumerate(names):
        pos = m - j if term != 'Intercept' else m + 1
        put(ws, r, 1, nm, 'txt'); put(ws, r, 2, f'=INDEX(LINEST({yr},{Xr},TRUE,TRUE),1,{pos})', 'res', '0.0000'); put(ws, r, 3, f'=INDEX(LINEST({yr},{Xr},TRUE,TRUE),2,{pos})', 'fx', '0.0000')
        pr = pyrows[pyrows.term == term]
        if len(pr):
            put(ws, r, 4, float(pr.se.iloc[0]), 'py', '0.0000'); put(ws, r, 5, float(pr.p.iloc[0]), 'py', '0.0000'); put(ws, r, 7, float(pr.coef.iloc[0]), 'py', '0.0000')
        put(ws, r, 6, stars(f'E{r}'), 'fx'); r += 1
    put(ws, r, 1, 'N', 'txt', bold=True); put(ws, r, 2, f'=COUNT({yr})', 'res'); r += 1
    put(ws, r, 1, 'R²', 'txt', bold=True); put(ws, r, 2, f'=INDEX(LINEST({yr},{Xr},TRUE,TRUE),3,1)', 'res', '0.0000'); rR2 = r; r += 1
    put(ws, r, 1, 'R² hiệu chỉnh = 1 − (1−R²)(N−1)/(N−m−1)', 'txt', bold=True); put(ws, r, 2, f'=1-(1-B{rR2})*(B{rR2 - 1}-1)/(B{rR2 - 1}-{m}-1)', 'res', '0.0000'); r += 1
    return r, r0, rR2


d6a = d.dropna(subset=X1 + ['ln_ebi']); d6b = d.dropna(subset=X2 + ['ln_ebi']); d6c = d.dropna(subset=[v + '_L1' for v in X2] + ['ln_ebi']); d6d = d6b[d6b.nam >= 2021]
# ---- M2 (chính)
a2, b2_, c02, m2 = khoi_mau(ws8, r, 'MẪU M2: 2019–2024, 8 biến X + biến giả năm (gốc 2019) – 331 quan sát', d6b, X2, [2020, 2021, 2022, 2023, 2024])
r = b2_ + 2
names2 = [(TEN_X[c], c) for c in X2] + [(f'Năm {y}', f'C(nam)[T.{y}]') for y in [2020, 2021, 2022, 2023, 2024]] + [('Hằng số', 'Intercept')]
r, r0_M2, rR2_M2 = bang_he_so(ws8, r, 'BẢNG 6 cột (2) – M2', a2, b2_, c02, m2, names2, b6[b6.model.str.startswith('M2: ')])
DOI_CHIEU += [('Bảng 6', 'M2 hệ số mật độ DN', f"'8_Bang6_7_9_OLS'!B{r0_M2 + 1}", round(float(b6[(b6.model.str.startswith('M2: ')) & (b6.term == 'dn_per_1000')].coef.iloc[0]), 4)),
              ('Bảng 6', 'M2 hệ số ln GRDP', f"'8_Bang6_7_9_OLS'!B{r0_M2}", round(float(b6[(b6.model.str.startswith('M2: ')) & (b6.term == 'ln_grdp')].coef.iloc[0]), 4)),
              ('Bảng 6', 'M2 R²', f"'8_Bang6_7_9_OLS'!B{rR2_M2}", round(float(b6[(b6.model.str.startswith('M2: ')) & (b6.term == 'R2')].coef.iloc[0]), 4))]
# VIF + tương quan (mẫu M2)
r += 1; put(ws8, r, 1, 'VIF của cột (2) – DÁN từ Python (VIF_j = 1/(1−R²_j), hồi quy X_j lên 7 biến X còn lại)', 'txt', bold=True); r += 1
for _, rv in vif.iterrows(): put(ws8, r, 1, TEN_X[rv.bien], 'txt'); put(ws8, r, 2, float(rv.VIF), 'py', '0.00'); r += 1
r += 1; put(ws8, r, 1, 'Ma trận tương quan Pearson (mẫu M2, hàm CORREL)', 'txt', bold=True); r += 1
vars9 = ['ln_ebi'] + X2; r = hdr(ws8, r, [''] + [TEN_X.get(v, 'ln(EBI)') for v in vars9]); rc0 = r
for i, vi in enumerate(vars9):
    put(ws8, r, 1, TEN_X.get(vi, 'ln(EBI)'), 'txt')
    for j, vj in enumerate(vars9):
        ci, cj = (3 if vi == 'ln_ebi' else c02 + X2.index(vi)), (3 if vj == 'ln_ebi' else c02 + X2.index(vj))
        put(ws8, r, 2 + j, f'=CORREL({A(L(ci), a2, b2_)},{A(L(cj), a2, b2_)})', 'fx', '0.000')
    r += 1
DOI_CHIEU.append(('Mục 4.4', 'r(ln GRDP, mật độ DN)', f"'8_Bang6_7_9_OLS'!D{rc0 + 1}", 0.641))
# hệ số chuẩn hóa (cột 5): chuẩn hoá y và X (không chuẩn hoá biến giả năm)
r += 1; put(ws8, r, 1, 'HỆ SỐ CHUẨN HÓA – cột (5): z = (x − TB)/SD cho y và 8 biến X trên mẫu M2, biến giả năm giữ nguyên', 'sub'); r += 1
r = hdr(ws8, r, ['Tỉnh', 'Năm', 'z(y)'] + [f'z({TEN_X[c]})' for c in X2] + [f'D{y}' for y in [2020, 2021, 2022, 2023, 2024]]); az = r
for rr in range(a2, b2_ + 1):
    put(ws8, az + rr - a2, 1, f'=A{rr}', 'fx'); put(ws8, az + rr - a2, 2, f'=B{rr}', 'fx')
    put(ws8, az + rr - a2, 3, f'=STANDARDIZE(C{rr},AVERAGE({A("C", a2, b2_)}),_xlfn.STDEV.S({A("C", a2, b2_)}))', 'fx', '0.000')
    for j in range(8):
        col = L(c02 + j); put(ws8, az + rr - a2, 4 + j, f'=STANDARDIZE({col}{rr},AVERAGE({A(col, a2, b2_)}),_xlfn.STDEV.S({A(col, a2, b2_)}))', 'fx', '0.000')
    for j in range(5): put(ws8, az + rr - a2, 12 + j, f'={L(c02 + 8 + j)}{rr}', 'fx')
bz = az + (b2_ - a2); r = bz + 2
r, r0_z, _ = bang_he_so(ws8, r, 'BẢNG 6 cột (5) – hệ số chuẩn hóa của M2', az, bz, 4, 13, names2, b6[b6.model == 'M2 chuẩn hóa'])
DOI_CHIEU.append(('Bảng 6', 'Hệ số chuẩn hóa mật độ DN', f"'8_Bang6_7_9_OLS'!B{r0_z + 1}", round(float(b6[(b6.model == 'M2 chuẩn hóa') & (b6.term == 'dn_per_1000')].coef.iloc[0]), 4)))
# ---- M1, M3, M4
for ten, dm, Xc, dys, key, colno in [('MẪU M1: 2018–2024, 7 biến X (không có hộ có máy tính) + biến giả năm (gốc 2018) – 384 quan sát', d6a, X1, [2019, 2020, 2021, 2022, 2023, 2024], 'M1: ', '(1)'),
                                     ('MẪU M3: biến X trễ 1 năm, 2020–2024 + biến giả năm (gốc 2020) – 278 quan sát', d6c, [v + '_L1' for v in X2], [2021, 2022, 2023, 2024], 'M3: ', '(3)'),
                                     ('MẪU M4: 2021–2024, 8 biến X + biến giả năm (gốc 2021) – 224 quan sát', d6d, X2, [2022, 2023, 2024], 'M4: ', '(4)')]:
    r += 1; a_, b_, c0_, m_ = khoi_mau(ws8, r, ten, dm, Xc, dys); r = b_ + 2
    nms = [(TEN_X.get(c.replace('_L1', ''), c) + (' (t−1)' if c.endswith('_L1') else ''), c) for c in Xc] + [(f'Năm {y}', f'C(nam)[T.{y}]') for y in dys] + [('Hằng số', 'Intercept')]
    r, r0_, rR2_ = bang_he_so(ws8, r, f'BẢNG 6 cột {colno} – {key.strip(": ")}', a_, b_, c0_, m_, nms, b6[b6.model.str.startswith(key)])
    pr = b6[(b6.model.str.startswith(key)) & (b6.term.isin(['dn_per_1000', 'dn_per_1000_L1']))].coef.iloc[0]
    DOI_CHIEU.append(('Bảng 6', f'{key.strip(": ")} hệ số mật độ DN', f"'8_Bang6_7_9_OLS'!B{r0_ + 1}", round(float(pr), 4)))
# ---- BẢNG 7: M5, M6 (mẫu M2 + 5 biến giả vùng, gốc ĐNB), M7 dán
REGD = [k for k in REG if k != 'ĐNB']
r += 1; a7, b7_, c07, m7 = khoi_mau(ws8, r, 'MẪU cho BẢNG 7: mẫu M2 + 5 biến giả vùng (gốc Đông Nam Bộ)', d6b, X2, [2020, 2021, 2022, 2023, 2024], REGD); r = b7_ + 2
names7 = [(TEN_X[c], c) for c in X2] + [(f'Năm {y}', f'C(nam)[T.{y}]') for y in [2020, 2021, 2022, 2023, 2024]] + [(f'{REG_FULL[k]} (so với ĐNB)', f'C(vung6)[T.{k}]') for k in REGD] + [('Hằng số', 'Intercept')]
r, r0_M6, _ = bang_he_so(ws8, r, 'BẢNG 7 cột (7) – M6: X + biến giả vùng', a7, b7_, c07, m7, names7, b7[b7.model.str.startswith('M6')])
DOI_CHIEU.append(('Bảng 7', 'M6 hệ số ĐBSCL', f"'8_Bang6_7_9_OLS'!B{r0_M6 + 13 + REGD.index('ĐBSCL')}", round(float(b7[(b7.model.str.startswith('M6')) & (b7.term == 'C(vung6)[T.ĐBSCL]')].coef.iloc[0]), 4)))
# M5: chỉ vùng + năm -> LINEST trên cột biến giả (năm 5 cột + vùng 5 cột liền nhau)
r += 1; put(ws8, r, 1, 'BẢNG 7 cột (6) – M5: chỉ biến giả vùng + năm', 'sub'); r += 1
r = hdr(ws8, r, ['Biến', 'Hệ số (LINEST)', 'SE cổ điển', 'SE cụm [DÁN]', 'p cụm [DÁN]', 'Dấu sao', 'Hệ số Python [DÁN]'])
Xr5 = f'${L(c07 + 8)}${a7}:${L(c07 + 17)}${b7_}'; yr5 = A("C", a7, b7_); m5 = 10
names5 = [(f'Năm {y}', f'C(nam)[T.{y}]') for y in [2020, 2021, 2022, 2023, 2024]] + [(f'{REG_FULL[k]} (so với ĐNB)', f'C(vung6)[T.{k}]') for k in REGD] + [('Hằng số', 'Intercept')]
p5 = b7[b7.model.str.startswith('M5')]
for j, (nm, term) in enumerate(names5):
    pos = m5 - j if term != 'Intercept' else m5 + 1
    put(ws8, r, 1, nm, 'txt'); put(ws8, r, 2, f'=INDEX(LINEST({yr5},{Xr5},TRUE,TRUE),1,{pos})', 'res', '0.0000'); put(ws8, r, 3, f'=INDEX(LINEST({yr5},{Xr5},TRUE,TRUE),2,{pos})', 'fx', '0.0000')
    pr = p5[p5.term == term]
    if len(pr): put(ws8, r, 4, float(pr.se.iloc[0]), 'py', '0.0000'); put(ws8, r, 5, float(pr.p.iloc[0]), 'py', '0.0000'); put(ws8, r, 7, float(pr.coef.iloc[0]), 'py', '0.0000')
    put(ws8, r, 6, stars(f'E{r}'), 'fx')
    if term == 'C(vung6)[T.ĐBSCL]': DOI_CHIEU.append(('Bảng 7', 'M5 hệ số ĐBSCL', f"'8_Bang6_7_9_OLS'!B{r}", round(float(pr.coef.iloc[0]), 4)))
    r += 1
put(ws8, r, 1, 'R²', 'txt', bold=True); put(ws8, r, 2, f'=INDEX(LINEST({yr5},{Xr5},TRUE,TRUE),3,1)', 'res', '0.0000'); r += 2
put(ws8, r, 1, 'Kiểm định Wald (M6): H0: 5 hệ số vùng = 0 – DÁN từ Python', 'txt', bold=True); put(ws8, r, 2, float(kq['wald_vung_M6']['stat']), 'py', '0.00'); put(ws8, r, 3, 'χ²(5)', 'txt'); put(ws8, r, 4, float(kq['wald_vung_M6']['p']), 'py', '0.0000'); put(ws8, r, 5, 'p', 'txt'); r += 2
put(ws8, r, 1, 'BẢNG 7 cột (8) – M7: hiệu ứng cố định tỉnh (62 biến giả tỉnh) – DÁN toàn bộ từ Python', 'sub'); r += 1
r = hdr(ws8, r, ['Biến', 'Hệ số [DÁN]', 'SE cụm [DÁN]', 'p [DÁN]', 'Dấu sao'])
for _, rw in b7[b7.model.str.startswith('M7')].iterrows():
    if rw.term in ('N', 'R2', 'R2_adj') or rw.term.startswith('C(nam)') or rw.term == 'Intercept':
        if rw.term in ('N', 'R2', 'R2_adj'): put(ws8, r, 1, {'N': 'N', 'R2': 'R²', 'R2_adj': 'R² hiệu chỉnh'}[rw.term], 'txt', bold=True); put(ws8, r, 2, float(rw.coef), 'py', '0.0000'); r += 1
        continue
    put(ws8, r, 1, TEN_X[rw.term], 'txt'); put(ws8, r, 2, float(rw.coef), 'py', '0.0000'); put(ws8, r, 3, float(rw.se), 'py', '0.0000'); put(ws8, r, 4, float(rw.p), 'py', '0.0000'); put(ws8, r, 5, stars(f'D{r}'), 'fx'); r += 1
# ---- BẢNG 9: cắt ngang 2024
r += 1; put(ws8, r, 1, 'BẢNG 9. Cắt ngang 2024 (58 tỉnh) với chỉ tiêu số chi tiết (sheet Digital_2024_cross của file gốc)', 'sub'); r += 1
d24 = d[d.nam == 2024].merge(dig[['tinh', 'internet_user_pct', 'it_skills_pct']], on='tinh')
C9 = ['internet_user_pct', 'it_skills_pct', 'hh_computer_pct', 'ln_grdp', 'dn_per_1000', 'urban_share_pct', 'trained_labor_pct']
r = hdr(ws8, r, ['Tỉnh', 'y = ln(EBI 2024)'] + [TEN_X[c] for c in C9]); a9 = r
for _, rw in d24.iterrows():
    r1 = ROW1[(2024, rw.tinh)]
    put(ws8, r, 1, rw.tinh, 'txt'); put(ws8, r, 2, f'={DL}{CL("ln_ebi")}{r1}', 'fx', '0.0000')
    put(ws8, r, 3, float(rw.internet_user_pct), 'in', '0.0'); put(ws8, r, 4, float(rw.it_skills_pct), 'in', '0.0')
    for j, c in enumerate(C9[2:]): put(ws8, r, 5 + j, f'={DL}{L(C1[c])}{r1}', 'fx', '0.000')
    r += 1
b9_ = r - 1; r += 1
for ten, cols, key in [('Bảng 9 cột (1): chỉ biến số', ['internet_user_pct', 'it_skills_pct'], 'C1'), ('Bảng 9 cột (2): biến số + kinh tế', ['internet_user_pct', 'it_skills_pct', 'ln_grdp', 'dn_per_1000', 'urban_share_pct'], 'C2'),
                       ('Bảng 9 cột (3): máy tính + kinh tế', ['hh_computer_pct', 'ln_grdp', 'dn_per_1000', 'urban_share_pct', 'trained_labor_pct'], 'C3')]:
    put(ws8, r, 1, ten + ' (LINEST cần cột liền nhau nên dùng CHOOSE để ghép cột)', 'sub'); r += 1
    r = hdr(ws8, r, ['Biến', 'Hệ số (LINEST)', 'SE cổ điển', 'SE HC1 [DÁN]', 'p HC1 [DÁN]', 'Dấu sao', 'Hệ số Python [DÁN]'])
    idx = ','.join(str(i + 1) for i in range(len(cols))); rngs = ','.join(A(L(3 + C9.index(c)), a9, b9_) for c in cols)
    Xr9 = f'CHOOSE({{{idx}}},{rngs})'; yr9 = A("B", a9, b9_); m9 = len(cols); p9 = b9[b9.model.str.startswith(key)]
    for j, c in enumerate(cols + ['Intercept']):
        pos = m9 - j if c != 'Intercept' else m9 + 1
        put(ws8, r, 1, TEN_X.get(c, 'Hằng số'), 'txt'); put(ws8, r, 2, f'=INDEX(LINEST({yr9},{Xr9},TRUE,TRUE),1,{pos})', 'res', '0.0000'); put(ws8, r, 3, f'=INDEX(LINEST({yr9},{Xr9},TRUE,TRUE),2,{pos})', 'fx', '0.0000')
        pr = p9[p9.term == c]
        if len(pr): put(ws8, r, 4, float(pr.se.iloc[0]), 'py', '0.0000'); put(ws8, r, 5, float(pr.p.iloc[0]), 'py', '0.0000'); put(ws8, r, 7, float(pr.coef.iloc[0]), 'py', '0.0000')
        put(ws8, r, 6, stars(f'E{r}'), 'fx')
        if c == 'internet_user_pct' and key == 'C2': DOI_CHIEU.append(('Bảng 9', 'cột (2) hệ số Internet', f"'8_Bang6_7_9_OLS'!B{r}", round(float(pr.coef.iloc[0]), 4)))
        r += 1
    put(ws8, r, 1, 'R²', 'txt', bold=True); put(ws8, r, 2, f'=INDEX(LINEST({yr9},{Xr9},TRUE,TRUE),3,1)', 'res', '0.0000'); r += 1
    put(ws8, r, 1, 'R² hiệu chỉnh', 'txt', bold=True); put(ws8, r, 2, f'=1-(1-B{r - 1})*(58-1)/(58-{m9}-1)', 'res', '0.0000')
    DOI_CHIEU.append(('Bảng 9', f'{key} R² hiệu chỉnh', f"'8_Bang6_7_9_OLS'!B{r}", round(float(p9[p9.term == 'R2_adj'].coef.iloc[0]), 4)))
    r += 2
ws8.freeze_panes = 'C5'; widths(ws8, [34, 12, 12, 12, 12, 10, 12, 12, 12, 12, 12, 10, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8])

# ====================================================================== 9. BẢNG 8 / HÌNH 8
ws9 = wb.create_sheet('9_Bang8_Oaxaca')
r = title(ws9, 1, 'BẢNG 8. Oaxaca–Blinder: Ȳ_L − Ȳ_N = (X̄_L − X̄_N)′β*  +  [phần hệ số]', 'Trung bình nhóm X̄_L, X̄_N và khoảng cách tính bằng AVERAGEIFS trên mẫu M2 (sheet 8). β* (hệ số tham chiếu từ mô hình gộp có biến nhóm, Jann 2008) và SE bootstrap: DÁN từ Python. Phần nguồn lực từng biến = (X̄_L − X̄_N)·β* là công thức; phần hệ số = khoảng cách − phần nguồn lực.')
# β* tính lại (giống hàm oaxaca trong 01, ref='pooled')
def design(dd, X):
    Xm = dd[X].copy(); Xm = pd.concat([Xm, pd.get_dummies(dd['nam'], prefix='y', drop_first=True).astype(float)], axis=1); return sm.add_constant(Xm)
P = d6b.copy(); XP = design(P, X2); XP['grp'] = P['core'].values; bR = sm.OLS(P['ln_ebi'], XP).fit().params.drop('grp')
r = hdr(ws9, r, ['Biến', 'X̄ lõi (core=1)', 'X̄ ngoại vi (core=0)', 'Chênh X̄', 'β* [DÁN]', 'Phần nguồn lực = chênh × β*', '% khoảng cách', 'Phần hệ số theo biến [DÁN – chỉ có nghĩa ở tổng]'])
cores = A(L(C1['core']), R1_0, R1_END)
r0 = r; items = [(TEN_X[c], c02 + X2.index(c), c) for c in X2] + [(f'Năm {y}', c02 + 8 + j, f'y_{y}') for j, y in enumerate([2020, 2021, 2022, 2023, 2024])]
for nm, col, key in items:
    rg = A(L(col), a2, b2_); core_rg = f"'1_Du_lieu'!{cores}"
    put(ws9, r, 1, nm, 'txt')
    put(ws9, r, 2, f"=AVERAGEIFS('8_Bang6_7_9_OLS'!{rg},'8_Bang6_7_9_OLS'!{A('AA', a2, b2_)},1)", 'fx', '0.000')
    put(ws9, r, 3, f"=AVERAGEIFS('8_Bang6_7_9_OLS'!{rg},'8_Bang6_7_9_OLS'!{A('AA', a2, b2_)},0)", 'fx', '0.000')
    put(ws9, r, 4, f'=B{r}-C{r}', 'fx', '0.0000'); put(ws9, r, 5, float(bR[key]), 'py', '0.0000'); put(ws9, r, 6, f'=D{r}*E{r}', 'res', '0.0000')
    put(ws9, r, 8, float(b8d.loc[key, 'coef']), 'py', '0.0000'); r += 1
r1e = r - 1
# cột lõi (AA) trong sheet 8 cho mẫu M2: ghi thêm cột core vào khối M2
for rr in range(a2, b2_ + 1): put(ws8, rr, 27, f'={DL}{CL("core")}{ROW1[(int(d6b.iloc[rr - a2].nam), d6b.iloc[rr - a2].tinh)]}', 'fx')
put(ws8, a2 - 1, 27, 'Lõi (core)', 'hd')
put(ws9, r, 1, 'Hằng số', 'txt'); put(ws9, r, 8, float(b8d.loc['const', 'coef']), 'py', '0.0000'); r += 1
put(ws9, r, 1, 'ln(EBI) trung bình', 'txt', bold=True)
put(ws9, r, 2, f"=AVERAGEIFS('8_Bang6_7_9_OLS'!{A('C', a2, b2_)},'8_Bang6_7_9_OLS'!{A('AA', a2, b2_)},1)", 'fx', '0.0000'); put(ws9, r, 3, f"=AVERAGEIFS('8_Bang6_7_9_OLS'!{A('C', a2, b2_)},'8_Bang6_7_9_OLS'!{A('AA', a2, b2_)},0)", 'fx', '0.0000')
rY = r; r += 1
put(ws9, r, 1, 'N lõi / N ngoại vi', 'txt', bold=True); put(ws9, r, 2, f"=COUNTIF('8_Bang6_7_9_OLS'!{A('AA', a2, b2_)},1)", 'fx'); put(ws9, r, 3, f"=COUNTIF('8_Bang6_7_9_OLS'!{A('AA', a2, b2_)},0)", 'fx'); r += 2
put(ws9, r, 1, 'KHOẢNG CÁCH = Ȳ_L − Ȳ_N', 'txt', bold=True); put(ws9, r, 2, f'=B{rY}-C{rY}', 'res', '0.0000'); put(ws9, r, 3, f'=(EXP(B{r})-1)*100', 'fx', '0.0'); put(ws9, r, 4, '% EBI cao hơn (e^gap − 1)', 'txt'); rG = r; r += 1
put(ws9, r, 1, 'PHẦN NGUỒN LỰC = Σ (chênh X̄ × β*)', 'txt', bold=True); put(ws9, r, 2, f'=SUM({A("F", r0, r1e)})', 'res', '0.0000'); put(ws9, r, 3, f'=B{r}/B{rG}*100', 'res', '0.0'); put(ws9, r, 4, '% khoảng cách', 'txt'); put(ws9, r, 5, float(b8[(b8['sample'].str.startswith('2019')) & (b8.ref == 'pooled')].endow_se.iloc[0]), 'py', '0.0000'); put(ws9, r, 6, 'SE bootstrap 300 lần [DÁN]', 'txt'); rE = r; r += 1
put(ws9, r, 1, 'PHẦN HỆ SỐ = khoảng cách − phần nguồn lực', 'txt', bold=True); put(ws9, r, 2, f'=B{rG}-B{rE}', 'res', '0.0000'); put(ws9, r, 3, f'=B{r}/B{rG}*100', 'res', '0.0'); put(ws9, r, 4, '% khoảng cách', 'txt'); put(ws9, r, 5, float(b8[(b8['sample'].str.startswith('2019')) & (b8.ref == 'pooled')].coef_se.iloc[0]), 'py', '0.0000'); put(ws9, r, 6, 'SE bootstrap [DÁN]', 'txt'); rC = r; r += 1
for rr in range(r0, r1e + 1): put(ws9, rr, 7, f'=F{rr}/$B${rG}*100', 'res', '0.0')
b8m = b8[(b8['sample'].str.startswith('2019')) & (b8.ref == 'pooled')].iloc[0]
DOI_CHIEU += [('Bảng 8', 'Khoảng cách ln EBI (M2)', f"'9_Bang8_Oaxaca'!B{rG}", round(float(b8m.gap), 4)), ('Bảng 8', 'Phần nguồn lực (M2, gộp)', f"'9_Bang8_Oaxaca'!B{rE}", round(float(b8m.endow), 4)),
              ('Bảng 8', '% nguồn lực', f"'9_Bang8_Oaxaca'!C{rE}", round(float(b8m.endow_share * 100), 2)), ('Bảng 8b', 'Nguồn lực do mật độ DN', f"'9_Bang8_Oaxaca'!F{r0 + 1}", round(float(b8d.loc['dn_per_1000', 'endow']), 4))]
r += 2; put(ws9, r, 1, 'BẢNG 8 đầy đủ (3 mẫu × 2 hệ số tham chiếu) – DÁN từ Python', 'sub'); r += 1
r = hdr(ws9, r, ['Mẫu', 'Hệ số tham chiếu', 'N lõi', 'N ngoại vi', 'ln EBI lõi', 'ln EBI ngoại vi', 'Khoảng cách', 'Phần nguồn lực', 'SE bootstrap', '% nguồn lực', 'Phần hệ số', 'SE bootstrap', '% hệ số'])
for _, rw in b8.iterrows():
    vals = [rw['sample'], 'Hệ số gộp (Jann 2008)' if rw.ref == 'pooled' else 'Hệ số vùng lõi', int(rw.nA), int(rw.nB), rw.yA, rw.yB, rw.gap, rw.endow, rw.endow_se, rw.endow_share * 100, rw.coef, rw.coef_se, rw.coef_share * 100]
    for j, v in enumerate(vals): put(ws9, r, 1 + j, (None if (isinstance(v, float) and np.isnan(v)) else v), 'py' if j >= 2 else 'txt', '0.0000' if j in (4, 5, 6, 7, 8, 10, 11) else ('0.0' if j in (9, 12) else None))
    r += 1
ch8 = BarChart(); ch8.type = 'bar'; ch8.title = 'Hình 8. Đóng góp của từng biến vào phần nguồn lực (% khoảng cách)'; ch8.height = 9; ch8.width = 16
ch8.add_data(Reference(ws9, min_col=7, min_row=r0 - 1, max_row=r0 + 7), titles_from_data=True); ch8.set_categories(Reference(ws9, min_col=1, min_row=r0, max_row=r0 + 7)); ch8.legend = None; ch8.y_axis.title = '% khoảng cách'; ws9.add_chart(axes(ch8), f'J{r0}')
widths(ws9, [40, 14, 14, 12, 12, 16, 12, 18, 12, 12, 12, 12, 10])

# ====================================================================== 10. SỐ LIỆU KHÁC
ws10 = wb.create_sheet('10_So_lieu_khac')
r = title(ws10, 1, 'CÁC CON SỐ KHÁC ĐƯỢC DẪN TRONG THÂN BÀI')
put(ws10, r, 1, 'Top 5 và Bottom 5 tỉnh theo EBI 2024 (LARGE/SMALL + INDEX/MATCH; các tỉnh đồng điểm 13 sẽ trả tên tỉnh đầu tiên gặp)', 'sub'); r += 1
a, b = YB1[2024]; ebi24, tinh24 = f'{DL}{A(CL("EBI"), a, b)}', f'{DL}{A(CL("tinh"), a, b)}'
r = hdr(ws10, r, ['Hạng', 'Tỉnh cao nhất', 'EBI', 'Tỉnh thấp nhất', 'EBI'])
for i in range(1, 6):
    put(ws10, r, 1, i, 'txt'); put(ws10, r, 3, f'=LARGE({ebi24},{i})', 'fx', '0.0'); put(ws10, r, 2, f'=INDEX({tinh24},MATCH(C{r},{ebi24},0))', 'fx')
    put(ws10, r, 5, f'=SMALL({ebi24},{i})', 'fx', '0.0'); put(ws10, r, 4, f'=INDEX({tinh24},MATCH(E{r},{ebi24},0))', 'fx'); r += 1
r += 1; put(ws10, r, 1, 'Độ lệch phải của phân phối EBI theo năm (Mục 3.1)', 'sub'); r += 1
r = hdr(ws10, r, ['Năm', 'Cao nhất', 'Trung vị', 'Tỷ số cao nhất/trung vị', 'P90', 'P10', 'P90/P10', 'Trung bình EBI'])
for y in YEARS:
    a, b = YB1[y]; rg = f'{DL}{A(CL("EBI"), a, b)}'
    put(ws10, r, 1, y, 'txt'); put(ws10, r, 2, f'=MAX({rg})', 'fx', '0.0'); put(ws10, r, 3, f'=MEDIAN({rg})', 'fx', '0.00'); put(ws10, r, 4, f'=B{r}/C{r}', 'res', '0.00')
    put(ws10, r, 5, f'=PERCENTILE({rg},0.9)', 'fx', '0.00'); put(ws10, r, 6, f'=PERCENTILE({rg},0.1)', 'fx', '0.00'); put(ws10, r, 7, f'=E{r}/F{r}', 'res', '0.00'); put(ws10, r, 8, f'=AVERAGE({rg})', 'fx', '0.00')
    DOI_CHIEU.append(('Mục 3.1', f'Tỷ số max/trung vị {y}', f"'10_So_lieu_khac'!D{r}", round(float(kq['ratio_max_median'][str(y)]), 3)))
    r += 1
r += 1; put(ws10, r, 1, 'EBI tương đối trung bình vùng lõi và vùng ngoại vi theo năm (Mục 4.3)', 'sub'); r += 1
r = hdr(ws10, r, ['Năm', 'Lõi (ĐNB + ĐBSH)', 'Ngoại vi (4 vùng)'])
for y in YEARS:
    put(ws10, r, 1, y, 'txt')
    put(ws10, r, 2, f"=AVERAGEIFS('3_Bang2_EBI_tuong_doi'!{A('G', R3_0, R3_1)},'3_Bang2_EBI_tuong_doi'!{A('B', R3_0, R3_1)},{y},'3_Bang2_EBI_tuong_doi'!{A('K', R3_0, R3_1)},1)", 'res', '0.0')
    put(ws10, r, 3, f"=AVERAGEIFS('3_Bang2_EBI_tuong_doi'!{A('G', R3_0, R3_1)},'3_Bang2_EBI_tuong_doi'!{A('B', R3_0, R3_1)},{y},'3_Bang2_EBI_tuong_doi'!{A('K', R3_0, R3_1)},0)", 'res', '0.0')
    DOI_CHIEU.append(('Mục 4.3', f'Lõi EBI tương đối {y}', f"'10_So_lieu_khac'!B{r}", round(float(kq['core_vs_periphery_rel']['1'][str(y)]), 1)))
    r += 1
r += 1; put(ws10, r, 1, 'Tỉnh không có EBI theo năm (COUNTBLANK) và danh sách (từ Python)', 'sub'); r += 1
r = hdr(ws10, r, ['Năm', 'Số tỉnh thiếu', 'Danh sách'])
for y in YEARS:
    a, b = YB1[y]; put(ws10, r, 1, y, 'txt'); put(ws10, r, 2, f'=COUNTBLANK({DL}{A(CL("EBI"), a, b)})', 'fx'); put(ws10, r, 3, ', '.join(kq['so_tinh_thieu_EBI'][str(y)]), 'py'); r += 1
widths(ws10, [10, 22, 12, 22, 12, 10, 10, 14])

# ====================================================================== 11. ĐỐI CHIẾU
ws11 = wb.create_sheet('11_Doi_chieu')
r = title(ws11, 1, 'ĐỐI CHIẾU: giá trị Excel (công thức) so với giá trị Python (script 01) – cột "Chênh lệch" phải ≈ 0')
r = hdr(ws11, r, ['Nhóm', 'Đại lượng', 'Ô Excel', 'Giá trị Excel', 'Giá trị Python', 'Chênh lệch', 'Khớp?'])
for nhom, ten, ref, py in DOI_CHIEU:
    put(ws11, r, 1, nhom, 'txt'); put(ws11, r, 2, ten, 'txt'); put(ws11, r, 3, ref.replace("'", ''), 'txt'); put(ws11, r, 4, f'={ref}', 'fx', '0.0000'); put(ws11, r, 5, py, 'py', '0.0000')
    put(ws11, r, 6, f'=D{r}-E{r}', 'fx', '0.0000'); put(ws11, r, 7, f'=IF(ABS(F{r})<=MAX(0.0006,ABS(E{r})*0.001),"✓","✗ KIỂM TRA")', 'res'); r += 1
put(ws11, r + 1, 1, 'Số dòng khớp / tổng:', 'txt', bold=True); put(ws11, r + 1, 4, f'=COUNTIF(G5:G{r - 1},"✓")&" / "&COUNTA(G5:G{r - 1})', 'res')
widths(ws11, [10, 40, 30, 14, 14, 12, 14])

wb.save(OUT)
print('Đã ghi', OUT, '| số dòng đối chiếu:', len(DOI_CHIEU))

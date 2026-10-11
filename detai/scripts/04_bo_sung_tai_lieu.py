# -*- coding: utf-8 -*-
"""Bổ sung 2 sheet vào Nghien_Cuu_Tai_Lieu_EBI.xlsx:
 - Kiem_Chung_E1_E6: kết quả kiểm chứng 6 tài liệu "bổ sung 2020–2026" (05/09/2026)
 - Tai_Lieu_Thay_The_2026: 12 tài liệu thay thế đã xác minh DOI, trình bày theo 17 cột như sheet Tai_Lieu_Bo_Sung_2020_2026
"""
import os, copy
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
F = os.path.join(ROOT, "Nghien_Cuu_Tai_Lieu_EBI.xlsx")
wb = load_workbook(F)
for s in ["Kiem_Chung_E1_E6", "Tai_Lieu_Thay_The_2026"]:
    if s in wb.sheetnames: del wb[s]
src = wb["Tai_Lieu_Bo_Sung_2020_2026"]
HEAD = [c.value for c in src[3]]           # 17 tiêu đề cột gốc
hdr_font, hdr_fill = copy.copy(src["A3"].font), copy.copy(src["A3"].fill)
title_font = copy.copy(src["A1"].font)
wrap = Alignment(wrap_text=True, vertical="top")

# ---------------- Sheet 1: kiểm chứng ----------------
ws = wb.create_sheet("Kiem_Chung_E1_E6")
ws["A1"] = "KẾT QUẢ KIỂM CHỨNG 6 TÀI LIỆU 'BỔ SUNG ĐƯƠNG ĐẠI 2020–2026' (E1–E6)"; ws["A1"].font = title_font
ws["A2"] = "Kiểm tra ngày 05/09/2026 qua Crossref (metadata DOI), trang nhà xuất bản, OpenAlex/Semantic Scholar, PMC/PLOS. Kết luận: chỉ E1 đúng; E2–E6 sai nội dung hoặc không tồn tại → KHÔNG dùng trong bài báo; thay bằng các tài liệu ở sheet Tai_Lieu_Thay_The_2026."
kc_head = ["Mã", "Tài liệu như ghi trong sheet Tai_Lieu_Bo_Sung_2020_2026", "Kết luận kiểm chứng", "Bằng chứng / chi tiết", "Hành động", "Thay bằng (mã ở sheet Tai_Lieu_Thay_The_2026)"]
KC = [
    ["E1", "van Dijk (2020). The Digital Divide. Polity Press.", "TỒN TẠI & ĐÚNG NỘI DUNG",
     "Kho lưu trữ ĐH Twente xác nhận sách Polity Press 2020, ISBN 978-1-5095-3444-9/-3445-6 (https://research.utwente.nl/en/publications/the-digital-divide-2/); điểm sách trên JASIST 72(1), 136–138, DOI 10.1002/asi.24355. Nội dung khung đa tầng + third-level/outcome divide đúng.",
     "GIỮ, dùng trong bài (trích dẫn van Dijk, 2020).", "—"],
    ["E2", "Pan, Xie, Wang & Ma (2022). 'Digital economy: An innovation driving factor for low-carbon development' — mô tả là hội tụ σ/β kinh tế số 30 tỉnh TQ.", "TỒN TẠI NHƯNG NỘI DUNG MÔ TẢ SAI (ghép chéo 2 bài)",
     "Bài thật của Pan, Xie, Wang & Ma (2022): 'Digital economy: An innovation driver for total factor productivity', Journal of Business Research 139, 303–311, DOI 10.1016/j.jbusres.2021.09.061 — về TFP, KHÔNG có hội tụ. Tiêu đề 'low-carbon development' là của Zhang, Lyu, Li & Geng (2022), Environmental Impact Assessment Review 96, 106821, DOI 10.1016/j.eiar.2022.106821 — cũng không phải hội tụ.",
     "BỎ.", "R2a (Wu et al., 2025); R2b (Li et al., 2025)"],
    ["E3", "Luo, Yimamu, Li, Wu & Irfan (2023). 'Digital economy, spatial spillovers, and regional economic d…' — mô tả là phân rã Theil between >60%.", "TỒN TẠI NHƯNG NỘI DUNG MÔ TẢ SAI",
     "Bài thật: Luo, Yimamu, Li, Wu, Irfan & Hao (2023). 'Digitalization and sustainable development: How could digital economy development improve green innovation in China?', Business Strategy and the Environment 32(4), 1847–1871, DOI 10.1002/bse.3223 — 278 thành phố, SDM/DID về đổi mới xanh; KHÔNG có phân rã Theil, không có 'between >60%'.",
     "BỎ.", "R3a (Fan et al., 2024); R3b (Wang, 2024)"],
    ["E4", "Cieślik & Rokicki (2021). 'The effect of infrastructure investments on regional e-commerce…' (NUTS-2 EU).", "KHÔNG TÌM THẤY",
     "Crossref không có bài nào của hai tác giả về TMĐT/băng rộng/logistics. Công trình thật của Rokicki 2018–2021 là về hạ tầng giao thông/CGE vùng Ba Lan (vd. Transportation 48(2), 747–772, DOI 10.1007/s11116-019-10076-w).",
     "BỎ.", "R4a (Leogrande et al., 2026); R4b (Hao, Rasiah & Mustafa, 2026); phụ: Sadowski et al. (2021)"],
    ["E5", "Zhang, Liu, Wang & Zhou (2022). 'Digital economy and regional income inequality: A spatial econometric…' — mô tả Oaxaca–Blinder, endowments >65%.", "KHÔNG TÌM THẤY",
     "Crossref không có bài khớp tiêu đề. Các bài thật gần chủ đề (vd. 'Digital economy development and the urban-rural income gap: Evidence from Chinese cities', PLOS ONE 2023) KHÔNG dùng Oaxaca–Blinder, không có con số 65%.",
     "BỎ. Lưu ý: không tìm thấy bài bình duyệt nào áp dụng Oaxaca–Blinder cho khoảng cách số GIỮA CÁC VÙNG; các bài thay thế phân tách khoảng cách số nông thôn–thành thị/giới ở cấp cá nhân (tiền lệ phương pháp).", "R5a (Zhao et al., 2022); R5b (Lin et al., 2023); phụ: Galperin & Arcidiacono (2021)"],
    ["E6", "Nguyen, T.T.H. & Pham, T.L. (2023). 'Provincial digital transformation and economic performance in Vietnam…', Journal of Asian Economics 88, 101642 — 'PCI cao khai thác Internet hiệu quả gấp 1,8 lần'.", "KHÔNG TÌM THẤY — DOI KHÔNG TỒN TẠI",
     "https://doi.org/10.1016/j.asieco.2023.101642 → 404; Crossref works/10.1016/j.asieco.2023.101642 → 404 (DOI chưa từng được đăng ký). Crossref không có bài nào trong Journal of Asian Economics với tiêu đề này. Không có nguồn nào cho con số '1,8 lần'.",
     "BỎ; không dùng con số 1,8 lần.", "R6a (Vu, 2026); R6b (Vu, Nguyen & Le, 2025)"],
]
for j, h in enumerate(kc_head, 1):
    c = ws.cell(row=4, column=j, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = wrap
for i, r in enumerate(KC, 5):
    for j, v in enumerate(r, 1):
        c = ws.cell(row=i, column=j, value=v); c.alignment = wrap
        if j == 3: c.font = Font(bold=True, color=("006100" if "ĐÚNG" in v else "9C0006"))
for j, w in enumerate([6, 45, 28, 70, 30, 34], 1): ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = "A5"

# ---------------- Sheet 2: tài liệu thay thế ----------------
ws = wb.create_sheet("Tai_Lieu_Thay_The_2026")
ws["A1"] = "MA TRẬN 12 CÔNG TRÌNH THAY THẾ ĐÃ XÁC MINH DOI (bổ sung 05/09/2026) — thay cho E2–E6"; ws["A1"].font = title_font
ws["A2"] = "Mỗi DOI đã được phân giải qua Crossref/nhà xuất bản. Cột 'Key Finding (EN)' là diễn đạt lại ngắn từ abstract kèm số liệu gốc (không sao chép nguyên văn); người dùng tra câu gốc tại DOI. Cột cuối cho biết bài này thay cho mã nào và được dùng ở đâu trong bài báo."
head = HEAD + ["Thay cho mã", "DOI / xác minh"]
for j, h in enumerate(head, 1):
    c = ws.cell(row=3, column=j, value=h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = wrap

def row(ma, tg, ten, rq_vn, rq_en, ly_vn, ly_en, mh_vn, mh_en, dv, dv_bien, iv, pp_vn, pp_en, kq_vn, kq_en, ud, thay, doi):
    return [ma, tg, ten, rq_vn, rq_en, ly_vn, ly_en, mh_vn, mh_en, dv, dv_bien, iv, pp_vn, pp_en, kq_vn, kq_en, ud, thay, doi]

R = [
 row("R2a", "Wu, Y., Wang, H., Guo, D., Li, Y., Zhang, H. & Shan, T.\n(2025)",
     "Spatio-temporal evolution and convergence patterns of E-commerce development in China: Regional disparities and policy implications\n/ PLOS ONE, 20(8), e0330066",
     "TMĐT ở 31 tỉnh Trung Quốc phân bố không gian–thời gian ra sao và có hội tụ giữa các vùng không?",
     "Whether provincial e-commerce shows spatio-temporal disparities and converges across regions.",
     "Hội tụ tân cổ điển (β); phân cực – lan tỏa không gian.", "Neoclassical convergence; spatial polarisation and spillovers.",
     "Mô hình hội tụ TMĐT cấp tỉnh thực nghiệm (σ, β, SDM). Bảo trợ trực tiếp cho RQ2/H3.", "Empirical provincial e-commerce convergence model (σ, β, SDM); direct template for RQ2/H3.",
     "Dữ liệu bảng 31 tỉnh Trung Quốc, 2013–2023.", "Mức phát triển TMĐT cấp tỉnh.", "GDP, số huyện/hạ tầng, ma trận trọng số không gian.",
     "Dagum Gini; Moran's I; σ-convergence (CV); β tuyệt đối & có điều kiện; Durbin không gian (SDM).", "Dagum Gini, Moran's I, σ/β convergence, spatial Durbin model.",
     "Gini tổng giảm 0,505 (2013) → 0,471 (2021); CV ≈1,2 → ≈1,0; β tuyệt đối = −0,01766 (p = 0,005); SDM: ln_gdp 0,071 (p = 0,010), lan tỏa 0,209 (p = 0,006). Các vùng khởi động muộn phát triển nhanh hơn.",
     "Overall Gini fell from 0.505 (2013) to 0.471 (2021); absolute β = −0.01766 (p = 0.005); regions that initiated later develop faster.",
     "Bảng 5 (hội tụ σ/β): mẫu chuẩn để tính CV/Gini theo năm và hồi quy tăng trưởng EBI lên EBI gốc; mở rộng không gian.", "E2", "10.1371/journal.pone.0330066 (PMC12396765)"),
 row("R2b", "Li, S., Lin, Z., Wu, Y. & Hu, Y.\n(2025)",
     "Measurement, Regional Disparities, and Spatial Convergence in the Symbiotic Level of China's Digital Innovation Ecosystem\n/ Systems, 13(4), 254",
     "Mức cộng sinh hệ sinh thái đổi mới số của 30 tỉnh khác biệt và hội tụ thế nào?", "Measure, compare and test convergence of provincial digital innovation ecosystems.",
     "Hội tụ có điều kiện; lan tỏa không gian.", "Conditional convergence with spatial spillovers.",
     "Mô hình đo lường + Dagum Gini + hội tụ σ/β không gian. Minh họa tình huống 'σ phân kỳ nhưng β hội tụ'.", "Composite index + Dagum Gini + spatial σ/β convergence; illustrates 'σ-divergence with β-convergence'.",
     "30 tỉnh Trung Quốc, 2013–2022.", "Chỉ số tổng hợp 3 chiều (chủ thể, môi trường, tương tác đổi mới số).", "Cơ cấu kinh tế, ứng dụng đổi mới (biến điều kiện).",
     "Entropy-TOPSIS; Dagum Gini; σ-convergence; β tuyệt đối/có điều kiện không gian.", "Entropy-TOPSIS, Dagum Gini, spatial σ/β convergence.",
     "Mô hình 'cao ở Đông, phẳng ở Trung, thấp ở Tây'; chênh lệch giữa vùng là nguồn chính; σ cho thấy chênh lệch mở rộng nhưng β cho thấy vùng thấp tăng nhanh hơn; cơ cấu KT và ứng dụng đổi mới đẩy nhanh hội tụ.",
     "Between-region differences are the main source of inequality; σ widens while β-convergence holds; economic structure and innovation application accelerate convergence.",
     "Bảng 5 và Mục 5 (thảo luận): đối chiếu kết quả 'β hội tụ nhưng phân tán vẫn cao' của EBI.", "E2", "10.3390/systems13040254"),
 row("R3a", "Fan, R., Nie, C., Zhao, Y., Hao, C. & Peng, C.\n(2024)",
     "Spatiotemporal Distribution and Regional Imbalance of China's Digital Economy\n/ Sustainability, 16(16), 6738",
     "Kinh tế số 30 tỉnh tiến hóa không gian–thời gian và mất cân đối vùng ra sao?", "Spatiotemporal evolution and regional imbalance of the provincial digital economy.",
     "Tích tụ không gian; chênh lệch duyên hải – nội địa.", "Spatial agglomeration; coast–inland gradient.",
     "Mô hình phân rã Dagum Gini (within/between/transvariation) + phân tích không gian. Bảo trợ RQ1/H1.", "Dagum Gini decomposition plus spatial analysis; supports RQ1/H1.",
     "30 tỉnh Trung Quốc, 2013–2021.", "Chỉ số kinh tế số (trọng số entropy).", "—",
     "Kernel density 3D; Dagum Gini; Moran's I; hot-spot.", "3D kernel density, Dagum Gini, Moran's I, hot-spot analysis.",
     "Mức KTS tăng đều, giảm dần Đông→Tây; chênh lệch vùng lớn nhưng thu hẹp; chênh lệch tổng CHỦ YẾU TỪ NỘI VÙNG, đóng góp liên vùng thấp.",
     "Regional disparity is large but narrowing; overall disparity stems mainly from within-region differences, with a low between-region contribution.",
     "Bảng 3 (phân rã Theil) & Mục 5: kết quả 'nội vùng chi phối' của EBI trùng với bài này (và ngược với mô tả sai của E3).", "E3", "10.3390/su16166738"),
 row("R3b", "Wang, W.\n(2024)",
     "Regional measurement and dynamic evolution of the development level of China's digital economy\n/ Frontiers in Environmental Economics, 3, 1396615",
     "Đo lường và theo dõi chênh lệch vùng của kinh tế số Trung Quốc.", "Measure and track regional disparities of China's digital economy.",
     "Khoảng cách số vùng; phụ thuộc không gian dương.", "Regional digital divide with positive spatial dependence.",
     "Mô hình Theil between/within 4 vùng + Moran's I. Mốc so sánh quốc tế cho tỷ trọng giữa vùng.", "Theil between/within decomposition across four regions plus Moran's I; international benchmark for the between-region share.",
     "30 tỉnh Trung Quốc, 2013–2021 (NBS, niên giám, CAICT).", "Chỉ số DEDCI 3 chiều (hạ tầng số, công nghiệp số, tiềm năng số hóa).", "—",
     "Trọng số entropy; chỉ số Theil phân rã between/within; Moran's I.", "Entropy weighting; Theil decomposition; Moran's I.",
     "DEDCI 0,0969 → 0,1765 (7,78%/năm); Theil 0,267 (2013) → 0,224 (2020); tỷ trọng LIÊN VÙNG 54,40% (2013), 58,16% (2020), 53,95% (2021); Quảng Đông gấp 19,65 lần Ninh Hạ (2021).",
     "Theil fell from 0.267 (2013) to 0.224 (2020); the between-region share was 54.40% (2013), 58.16% (2020) and 53.95% (2021).",
     "Bảng 3 & Mục 5: mốc 54–58% để đối chiếu với 31–37% của Việt Nam.", "E3", "10.3389/frevc.2024.1396615"),
 row("R4a", "Leogrande, A., Drago, C., Costantiello, A. & Arnone, M.\n(2026)",
     "From Connectivity to Commerce: A Multi-Technique Investigation of E-Commerce Drivers in Italy's Regional Landscape\n/ Journal of Theoretical and Applied Electronic Commerce Research, 21(5), 137",
     "Yếu tố nào quyết định chênh lệch vùng về tỷ lệ doanh nghiệp bán hàng trực tuyến ở Ý?", "Drivers of regional disparities in enterprise e-commerce adoption in Italy.",
     "Năng lực số (kết nối, kỹ năng, làm việc từ xa) quan trọng hơn mức giàu có.", "Digital capacity rather than affluence drives adoption.",
     "Mô hình yếu tố quyết định TMĐT cấp vùng (IV panel + ML). Bảo trợ RQ3/H4.", "Regional e-commerce determinants model (IV panel + machine learning); supports RQ3/H4.",
     "Các vùng của Ý (Istat), dữ liệu bảng.", "Tỷ lệ doanh nghiệp (≥10 lao động) bán hàng trực tuyến.", "Sử dụng Internet, kết nối hộ gia đình (băng rộng), làm việc từ xa, thu nhập, đổi mới.",
     "Panel biến công cụ (độ dốc địa hình làm IV cho băng rộng); regularized regression, random forest, boosting; phân cụm; phân tích mạng.", "IV panel regressions; regularised regression, random forest, boosting; clustering; network analysis.",
     "Sử dụng Internet, kết nối số hộ gia đình và làm việc từ xa là dự báo mạnh nhất; mức giàu có không quyết định; khoảng cách bền vững Bắc–Trung vs Nam–đảo.",
     "Internet use, household connectivity and remote work are the strongest predictors; affluence is not decisive; a persistent North–South divide.",
     "Bảng 6 (OLS) & Mục 2.4: bảo trợ luận điểm 'giàu có không đủ; năng lực số và hệ sinh thái quyết định'.", "E4", "10.3390/jtaer21050137"),
 row("R4b", "Hao, T., Rasiah, R. & Mustafa, S.\n(2026)",
     "Digital Infrastructure, SME E-Commerce, and Economic Growth: Evidence from China's Platform Economy\n/ Economies, 14(2), 40",
     "Số hóa DNNVV và nền tảng ảnh hưởng tăng trưởng vùng ra sao; hạ tầng số điều tiết thế nào?", "Effects of SME digitalisation on regional growth, moderated by digital infrastructure readiness.",
     "Resource-Based View; hạ tầng số là năng lực bổ trợ.", "Resource-based view; digital infrastructure readiness as enabling capability.",
     "Mô hình tương tác 'hạ tầng số × TMĐT DNNVV' (system GMM). Bảo trợ vai trò hệ sinh thái doanh nghiệp (H4).", "Interaction model of digital infrastructure and SME e-commerce (system GMM); supports the enterprise-ecosystem channel (H4).",
     "30 tỉnh Trung Quốc, 2015–2024.", "Tăng trưởng GDP vùng.", "Số hóa DNNVV, hiệu quả chuỗi cung ứng/logistics, thanh toán di động, việc làm công nghệ, kinh tế nền tảng, mức sẵn sàng hạ tầng số (DIR).",
     "System GMM hai bước.", "Two-step system GMM.",
     "+10 điểm % số hóa DNNVV → ≈ +0,3 điểm % tăng trưởng; +10 điểm DIR → ≈ +0,4 điểm %; DIR khuếch đại tác động của TMĐT, thanh toán di động, marketing số.",
     "A 10-point rise in SME digitalisation adds about 0.3 percentage points to growth; a 10-point rise in DIR adds about 0.4; DIR amplifies e-commerce effects.",
     "Mục 2.4 & 5: cơ chế 'hạ tầng số nhân với hệ sinh thái doanh nghiệp'.", "E4", "10.3390/economies14020040"),
 row("R4c", "Sadowski, A., Lewandowska-Gwarda, K., Pisarek-Bartoszewska, R. & Engelseth, P.\n(2021)",
     "A longitudinal study of e-commerce diversity in Europe\n/ Electronic Commerce Research, 21(1), 169–194",
     "Mua sắm trực tuyến ở các vùng châu Âu phân bố và biến động ra sao 2010–2019?", "How online shopping varies across European regions, 2010–2019.",
     "Phụ thuộc không gian; đa dạng TMĐT vùng.", "Spatial dependence; regional e-commerce diversity.",
     "Phân tích ESDA/Moran cục bộ cấp vùng (tham chiếu phụ).", "Regional ESDA / local Moran analysis (supplementary reference).",
     "Vùng NUTS châu Âu, Eurostat, 2010–2019.", "Tỷ lệ cá nhân mua sắm trực tuyến.", "—",
     "ESDA, Moran cục bộ, phân cụm vùng.", "ESDA, local Moran's I, regional clustering.",
     "Các cụm vùng theo sở thích mua sắm trực tuyến bền vững theo thời gian; không hồi quy hạ tầng.", "Persistent regional clusters of online-shopping preference over time.",
     "Chỉ dùng cho phần 'chênh lệch vùng TMĐT ở EU' (không trích trong bản hiện tại).", "E4 (phụ)", "10.1007/s10660-021-09466-z"),
 row("R5a", "Zhao, L., Cao, C., Li, Y. & Li, Y.\n(2022)",
     "Determinants of the digital outcome divide in E-learning between rural and urban students: Empirical evidence from the COVID-19 pandemic based on capital theory\n/ Computers in Human Behavior, 130, 107177",
     "Vì sao học sinh nông thôn hưởng lợi từ e-learning ít hơn thành thị?", "What explains the rural–urban digital outcome divide in e-learning.",
     "Lý thuyết vốn (kinh tế/văn hóa/xã hội) + khoảng cách kết quả số (tầng 3 của van Dijk).", "Capital theory; third-level (outcome) digital divide.",
     "Mô hình phân tách Blinder–Oaxaca cho khoảng cách kết quả số. Bảo trợ RQ3/H5.", "Blinder–Oaxaca decomposition of the digital outcome divide; supports RQ3/H5.",
     "Khảo sát 492 học sinh THCS Trung Quốc (46,75% nông thôn).", "Mức tham gia hành vi e-learning.", "Tự tin e-learning, động cơ nội tại, hỗ trợ cha mẹ, nguồn lực.",
     "Blinder–Oaxaca.", "Blinder–Oaxaca decomposition.",
     "Chênh lệch −0,393 điểm; HIỆU ỨNG ĐẶC ĐIỂM (explained) 72,9%, hiệu ứng hệ số 27,1%; đóng góp lớn nhất: tự tin e-learning 33,4%, hỗ trợ cha mẹ 19,5%, động cơ nội tại 16,8%.",
     "The gap of −0.393 is 72.9% explained by characteristics and 27.1% by coefficients; e-learning self-efficacy (33.4%) and parental support (19.5%) contribute most.",
     "Bảng 8 (Oaxaca–Blinder) & Mục 5: mốc 73% do nguồn lực để đối chiếu với 79% của EBI.", "E5", "10.1016/j.chb.2021.107177 (PMC9758626)"),
 row("R5b", "Lin, R., Chu, J., Yang, L., Lou, L., Yu, H. & Yang, J.\n(2023)",
     "What are the determinants of rural-urban divide in teachers' digital teaching competence? Empirical evidence from a large sample\n/ Humanities and Social Sciences Communications, 10, 422",
     "Yếu tố nào tạo nên khoảng cách năng lực dạy học số giữa giáo viên nông thôn – thành thị?", "Determinants of the rural–urban gap in teachers' digital teaching competence.",
     "Khoảng cách kỹ năng (thái độ ICT, kỹ năng ICT, hiểu biết dữ liệu).", "Skills-based digital divide.",
     "Phân tách Blinder–Oaxaca trên mẫu lớn (tiền lệ phương pháp cho H5).", "Blinder–Oaxaca decomposition on a large sample (methodological precedent for H5).",
     "11.784 giáo viên K-12 Trung Quốc (43,40% nông thôn).", "Năng lực dạy học số.", "Thái độ ICT, kỹ năng ICT, data literacy.",
     "So sánh nhóm + Blinder–Oaxaca.", "Group comparison + Blinder–Oaxaca decomposition.",
     "Giáo viên nông thôn thấp hơn có ý nghĩa ở mọi chỉ số; data literacy và kỹ năng ICT là hai yếu tố giải thích lớn nhất.",
     "Rural teachers score significantly lower on all indicators; data literacy and ICT skills explain most of the gap.",
     "Mục 2.5: bảo trợ đưa biến kỹ năng/nhân lực số vào phân tách.", "E5", "10.1057/s41599-023-01933-2"),
 row("R5c", "Galperin, H. & Arcidiacono, M.\n(2021)",
     "Employment and the gender digital divide in Latin America: A decomposition analysis\n/ Telecommunications Policy, 45(7), 102166",
     "Việc làm giải thích bao nhiêu phần khoảng cách số theo giới ở Mỹ Latinh?", "How much of the gender digital divide in Latin America is explained by employment.",
     "Bất bình đẳng số gắn với bất bình đẳng kinh tế – xã hội.", "Digital inequality embedded in socio-economic inequality.",
     "Phân tách phi tuyến (Fairlie/Yun) cho khoảng cách số (tham chiếu phụ).", "Non-linear (Fairlie/Yun) decomposition of the digital divide (supplementary).",
     "Khảo sát hộ gia đình nhiều nước Mỹ Latinh.", "Sử dụng Internet.", "Việc làm, thu nhập, học vấn, tuổi.",
     "Phân tách phi tuyến kiểu Oaxaca–Blinder.", "Non-linear Oaxaca–Blinder-type decomposition.",
     "Nếu nữ có việc làm bằng nam, khoảng cách số giới giảm ít nhất 1/4 (cận dưới).", "Equalising employment would close at least one quarter of the gender digital divide.",
     "Mục 2.5: dẫn phương pháp phân tách cho khoảng cách số.", "E5 (phụ)", "10.1016/j.telpol.2021.102166"),
 row("R6a", "Vu, H.Q.\n(2026; online 2025)",
     "Institutional quality and infrastructure as catalyst to the effects of ICT on provincial total factor productivity\n/ Competitiveness Review: An International Business Journal, 36(4), 748–768",
     "ICT, chất lượng thể chế và hạ tầng tác động TFP cấp tỉnh Việt Nam ra sao; năng lực cạnh tranh cấp tỉnh có điều tiết không?", "Effects of ICT, institutions and infrastructure on provincial TFP in Vietnam, with provincial competitiveness as moderator.",
     "Kinh tế tri thức; bổ trợ giữa công nghệ và thể chế.", "Knowledge economy; technology–institution complementarity.",
     "Mô hình điều tiết thể chế cấp tỉnh Việt Nam (GMM). Thay thế trực tiếp luận điểm 'PCI điều tiết' của E6.", "Provincial institutional-moderation model for Vietnam (GMM); replaces the 'PCI moderation' claim of E6.",
     "Các tỉnh Việt Nam, 2010–2022.", "TFP cấp tỉnh.", "Chỉ số ICT, chất lượng thể chế (năng lực cạnh tranh cấp tỉnh), hạ tầng; tương tác ICT×thể chế, ICT×hạ tầng.",
     "Panel động, GMM.", "Dynamic panel GMM.",
     "ICT tăng TFP; quản trị tốt và hạ tầng tăng TFP; tỉnh có quản trị và hạ tầng tốt hơn thu được lợi ích TFP từ ICT lớn hơn (tương tác dương). KHÔNG có con số '1,8 lần'.",
     "ICT raises TFP; provinces with better governance and infrastructure obtain larger TFP gains from ICT (positive interaction).",
     "Mục 2.4 & 5: vai trò thể chế địa phương; hướng nghiên cứu tiếp theo với PCI.", "E6", "10.1108/CR-03-2025-0105"),
 row("R6b", "Vu, H.Q., Nguyen, H.M.T. & Le, T.T.\n(2025)",
     "Digital transformation and economic growth: empirical evidence from Vietnam\n/ Post-Communist Economies, 37(7), 910–932",
     "Chuyển đổi số tác động tăng trưởng cấp tỉnh Việt Nam thế nào?", "Impact of digital transformation on provincial growth in Vietnam.",
     "ICT nâng năng suất vốn và lao động; khuếch đại FDI; cải thiện quản trị.", "ICT raises factor productivity and amplifies FDI and governance effects.",
     "Mô hình tăng trưởng có tương tác, dữ liệu bảng cấp tỉnh từ Niên giám Thống kê (cùng nguồn với đề tài).", "Growth model with interactions on a provincial panel built from the Statistical Yearbook (same source as this study).",
     "Các tỉnh Việt Nam, 2010–2022; Niên giám Thống kê.", "Tăng trưởng kinh tế cấp tỉnh.", "Chỉ số ICT, vốn, lao động, FDI, quản trị; tương tác ICT×vốn/lao động/FDI/quản trị.",
     "Panel động GMM.", "Dynamic panel GMM.",
     "ICT tăng năng suất vốn và lao động, đẩy nhanh tăng trưởng cấp tỉnh; khuếch đại hiệu ứng FDI và cải thiện quản trị.",
     "ICT raises capital and labour productivity, accelerates provincial growth and amplifies FDI and governance effects.",
     "Mục 2.4: bằng chứng Việt Nam cùng nguồn dữ liệu; mẫu mô hình tương tác.", "E6", "10.1080/14631377.2025.2511511"),
]
for i, r in enumerate(R, 4):
    for j, v in enumerate(r, 1):
        c = ws.cell(row=i, column=j, value=v); c.alignment = wrap
    ws.row_dimensions[i].height = 210
for j in range(1, len(head) + 1):
    ws.column_dimensions[get_column_letter(j)].width = src.column_dimensions[get_column_letter(j)].width or 28 if j <= 17 else 22
ws.freeze_panes = "C4"
wb.save(F); print("Đã thêm 2 sheet:", wb.sheetnames)

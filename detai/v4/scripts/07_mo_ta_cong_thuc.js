// Sinh file Word mô tả công thức tính của mọi biến/chỉ số đã tính thêm: detai/MO-TA-CONG-THUC-TINH.docx
// Chạy trong thư mục detai:  node scripts/07_mo_ta_cong_thuc.js
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, HeadingLevel, ShadingType, LevelFormat, PageNumber, Footer } = require("docx");

const ROOT = path.resolve(__dirname, "..");
const FONT = "Times New Roman", MONO = "Consolas";
const PAGE_W = 11906, MARGIN = 1418, CW = PAGE_W - 2 * MARGIN;

// markup: **đậm**, *nghiêng*, `mã`
function runs(text, base = {}) {
  const out = []; const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g; let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: 24, ...base }));
    const s = m[0];
    if (s.startsWith("**")) out.push(new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, size: 24, ...base }));
    else if (s.startsWith("*")) out.push(new TextRun({ text: s.slice(1, -1), italics: true, font: FONT, size: 24, ...base }));
    else out.push(new TextRun({ text: s.slice(1, -1), font: MONO, size: 21, ...base }));
    last = m.index + s.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size: 24, ...base }));
  return out;
}
const P = (t, o = {}) => new Paragraph({ children: runs(t), alignment: AlignmentType.JUSTIFIED, spacing: { after: o.after ?? 120, line: 276 }, indent: o.noindent ? undefined : { firstLine: 0 } });
const H = (t, lv) => new Paragraph({ heading: lv === 1 ? HeadingLevel.HEADING_1 : lv === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
  children: runs(t, { bold: true, size: lv === 1 ? 28 : 24, color: "000000" }),
  spacing: { before: lv === 1 ? 280 : 180, after: 120 }, keepNext: true });
const EQ = (t) => new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 24 })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 120 } });
const CODE = (lines) => lines.map((l) => new Paragraph({ children: [new TextRun({ text: l, font: MONO, size: 19 })],
  shading: { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" }, spacing: { after: 0, line: 250 }, indent: { left: 400 } }));
const BUL = (items) => items.map((it) => new Paragraph({ children: runs(it), numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 } }));
const thin = { style: BorderStyle.SINGLE, size: 6, color: "000000" }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function TBL(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0); const w = widths.map((x) => Math.round(x / total * CW));
  const cell = (t, i, hd, last) => new TableCell({
    children: String(t).split("\n").map((l) => new Paragraph({ children: runs(l, { size: 19, bold: hd }), spacing: { after: 0, line: 245 } })),
    width: { size: w[i], type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 70, right: 70 },
    borders: { top: hd ? thin : none, bottom: hd || last ? thin : none, left: none, right: none },
    shading: hd ? { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" } : undefined,
  });
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: w, layout: "fixed", rows: [
    new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true, false)) }),
    ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false, ri === rows.length - 1)) })),
  ] });
}
const GAP = () => new Paragraph({ children: [], spacing: { after: 80 } });

const ch = [];
ch.push(new Paragraph({ children: [new TextRun({ text: "MÔ TẢ CÔNG THỨC TÍNH CÁC BIẾN VÀ CHỈ SỐ ĐÃ TÍNH THÊM", bold: true, font: FONT, size: 30 })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "Bài báo: “Khoảng cách thương mại điện tử giữa các tỉnh, thành phố Việt Nam giai đoạn 2017–2024” — tác giả Đào Thị Hồng Vân", font: FONT, size: 22, italics: true })], alignment: AlignmentType.CENTER, spacing: { after: 40 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "Tài liệu kỹ thuật kèm theo mã nguồn detai\\v2\\scripts\\01_phan_tich.py và detai\\v2\\hinh-bai-bao\\tao_hinh.py — ngày 06/09/2026", font: FONT, size: 22 })], alignment: AlignmentType.CENTER, spacing: { after: 240 } }));

// ===================== 0 =====================
ch.push(H("0. Cách đọc tài liệu này", 1));
ch.push(P("Tài liệu liệt kê mọi đại lượng KHÔNG có sẵn trong file dữ liệu gốc `EBI_NGTK_Panel_2017_2024.xlsx` mà đã được tính thêm trong quá trình phân tích. Mỗi mục ghi: tên biến trong mã nguồn, công thức, đơn vị, dữ liệu đầu vào (cột nào của file gốc hoặc bảng nào của Niên giám Thống kê), ý nghĩa, và nơi dùng trong bài (bảng/hình). Ký hiệu chung:"));
ch.push(...BUL([
  "*i* = tỉnh (63 tỉnh, thành phố theo đơn vị hành chính trước sáp nhập), *t* = năm (2017–2024), *k* = vùng (6 vùng kinh tế – xã hội theo Niên giám, cột `vung6`).",
  "*N_t* = số tỉnh có EBI trong năm *t* (54–58 tỉnh); mọi chỉ số bất bình đẳng chỉ tính trên các tỉnh có EBI, mỗi tỉnh một quan sát có trọng số bằng nhau.",
  "ln = logarit tự nhiên; ȳ = trung bình giản đơn; σ = độ lệch chuẩn mẫu (chia *N* − 1).",
  "Số trong bài dùng dấu phẩy thập phân; trong file CSV kết quả dùng dấu chấm.",
]));
ch.push(P("Thứ tự trình bày: (1) biến bổ sung từ Niên giám; (2) biến dẫn xuất từ file gốc; (3) chỉ số bất bình đẳng; (4) chỉ số hội tụ; (5) mô hình hồi quy và các thống kê kèm theo; (6) phân tách Oaxaca–Blinder; (7) bảng tra nhanh."));

// ===================== 1 =====================
ch.push(H("1. Ba biến bổ sung lấy từ kho Niên giám Thống kê", 1));
ch.push(P("Đây là các chỉ tiêu KHÔNG có trong file gốc, được lấy từ các bảng “phân theo địa phương” đã trích sẵn của 8 cuốn Niên giám 2017–2024 (thư mục `data\\csv\\tonghop`), ghép vào panel theo khóa (tỉnh, năm). Mục đích: bổ sung biến đại diện cho logistics và hạ tầng bán lẻ mà đề cương yêu cầu nhưng file gốc chưa có."));

ch.push(H("1.1. Khối lượng hàng hóa vận chuyển đường bộ bình quân đầu người — `freight_pc`, `ln_freight`", 2));
ch.push(EQ("freight_pc_it = HH_it / pop_it        ln_freight_it = ln(freight_pc_it)"));
ch.push(...BUL([
  "**HH_it**: khối lượng hàng hóa vận chuyển bằng đường bộ của tỉnh *i* năm *t*, đơn vị *nghìn tấn*. Nguồn: bảng “Khối lượng hàng hoá vận chuyển bằng đường bộ phân theo địa phương”, file `khoi-luong-hang-hoa-van-chuyen-bang-duong-bo-phan-theo-dia-p_tu-2010-2024.csv`.",
  "**pop_it**: dân số trung bình, đơn vị *nghìn người* (cột `pop_th` của file gốc).",
  "Vì tử số là nghìn tấn và mẫu số là nghìn người nên **đơn vị của freight_pc là tấn/người/năm** (không cần nhân hệ số).",
  "Ý nghĩa: biến đại diện (proxy) cho mức độ hoạt động logistics – giao nhận hàng hóa của tỉnh. Lấy ln để giảm ảnh hưởng của các tỉnh cực đoan và để hệ số hồi quy đọc theo % (tăng 1% vận tải/người gắn với thay đổi β% EBI).",
  "Dùng ở: Bảng 1 (mô tả), Bảng 6–8, Hình 8.",
]));

ch.push(H("1.2. Mật độ cơ sở bán lẻ hiện đại — `modern_retail_100k`", 2));
ch.push(EQ("modern_retail_100k_it = (ST_it + TTTM_it) / pop_it × 100"));
ch.push(...BUL([
  "**ST_it**: số siêu thị có đến 31/12; **TTTM_it**: số trung tâm thương mại có đến 31/12 (hai bảng riêng trong Niên giám, file `so-luong-sieu-thi-…csv` và `so-luong-trung-tam-thuong-mai-…csv`).",
  "pop_it tính bằng nghìn người, nên (cơ sở / nghìn người) × 100 = **cơ sở trên 100.000 dân**.",
  "Quy ước ô trống: trong Niên giám, tỉnh không có trung tâm thương mại được in trống hoặc “-”. Các ô này được gán **0** (không phải thiếu số liệu) trước khi cộng. Lệnh: `df[[\"sieuthi\", \"tttm\"]].fillna(0)`.",
  "Ý nghĩa: hạ tầng bán lẻ hiện đại tại chỗ; trong bài dùng để kiểm tra quan hệ bổ sung/thay thế giữa kênh bán lẻ hiện đại và giao dịch trực tuyến.",
  "Dùng ở: Bảng 1, Bảng 6–8, Hình 8.",
]));

ch.push(H("1.3. Doanh nghiệp đăng ký thành lập mới trên 1.000 dân — `dn_moi_1000`", 2));
ch.push(EQ("dn_moi_1000_it = DNmoi_it / pop_it"));
ch.push(...BUL([
  "DNmoi_it: số doanh nghiệp đăng ký thành lập mới trong năm (bảng “Số doanh nghiệp đăng ký thành lập mới phân theo địa phương”, 2016–2024); chia cho nghìn người cho ra **doanh nghiệp mới trên 1.000 dân**.",
  "Biến này được tính và lưu trong `panel_phan_tich.csv` nhưng **không đưa vào mô hình cuối** vì trùng nội dung với mật độ doanh nghiệp đang hoạt động (`dn_per_1000`) và tương quan cao với nó; giữ lại để phân tích bổ sung nếu cần.",
]));

ch.push(H("1.4. Các chỉ tiêu số chỉ có năm 2024", 2));
ch.push(P("Sheet `Digital_2024_cross` của file gốc (tỷ lệ người sử dụng Internet `internet_user_pct`, tỷ lệ hộ có kết nối Internet `hh_internet_pct`, tỷ lệ người dùng điện thoại di động, tỷ lệ tham gia mạng xã hội, tỷ lệ người có kỹ năng CNTT `it_skills_pct`) được ghép với panel năm 2024 theo tên tỉnh, không biến đổi thêm. Chỉ dùng cho mô hình cắt ngang ở Bảng 9. Đã tìm trong toàn bộ 229 bảng Niên giám đã trích: các chỉ tiêu này chỉ được công bố theo tỉnh từ Niên giám 2024, và không có PCI cũng như chỉ số logistics cấp tỉnh; hai chỉ tiêu này được ghi là hạn chế ở Mục 3.2 và Mục 5 của bài."));

// ===================== 2 =====================
ch.push(H("2. Biến dẫn xuất từ các cột có sẵn trong file gốc", 1));
ch.push(TBL(["Biến mới", "Công thức", "Đầu vào (cột file gốc)", "Đơn vị / ý nghĩa", "Dùng ở"], [
  ["`ln_ebi`", "ln(EBI_it)", "EBI", "Biến phụ thuộc của mọi hồi quy; hệ số đọc theo % thay đổi EBI. Chỉ tính khi EBI > 0 (mọi giá trị EBI đều dương).", "Bảng 6–9"],
  ["`ln_grdp`", "ln(grdp_pc_tr)", "grdp_pc_tr (triệu đồng/người, có từ 2018)", "Quy mô kinh tế; trống năm 2017.", "Bảng 5b, 6–8"],
  ["`ln_income`", "ln(income_month_th)", "income_month_th (nghìn đồng/tháng)", "Biến thu nhập thay thế GRDP; tính sẵn, không dùng trong mô hình cuối.", "—"],
  ["`ln_retail_pc`", "ln(retail_pc_tr)", "retail_pc_tr (triệu đồng/người)", "Bán lẻ bình quân đầu người; tính sẵn, không dùng trong mô hình cuối vì trùng GRDP.", "—"],
  ["`ln_fdi_pc`", "ln(1 + fdi_pc_usd)", "fdi_pc_usd (USD/người, lũy kế)", "Độ mở; cộng 1 trước khi lấy ln để không lỗi khi FDI = 0 (một số tỉnh miền núi có FDI ≈ 0).", "Bảng 6–8, Hình 6"],
  ["`ln_dn`", "ln(dn_per_1000)", "dn_per_1000", "Tính sẵn để kiểm tra dạng hàm; mô hình chính dùng `dn_per_1000` ở dạng mức.", "—"],
  ["`EBI_rel`", "EBI_it / ȳ_t × 100, với ȳ_t = trung bình giản đơn EBI của N_t tỉnh có số liệu năm t", "EBI, nam", "**EBI tương đối**: trung bình các tỉnh trong năm = 100. Cho phép so sánh vị thế tương đối qua điểm gãy phương pháp năm 2021 (VECOM đổi thang điểm).", "Bảng 2, Bảng 5 (dòng toàn kỳ và dòng bảng động), Hình 2"],
  ["`NNL_HT_rel`, `B2C_rel`, `B2B_rel`", "Trụ cột_it / trung bình trụ cột năm t × 100", "NNL_HT, B2C, B2B", "Trụ cột tương đối, cùng cách chuẩn hóa như EBI_rel.", "Bảng 2 (năm 2024)"],
  ["`core`", "1 nếu vung6 ∈ {ĐNB, ĐBSH}, ngược lại 0", "vung6", "Biến chỉ báo **vùng lõi** (hai vùng động lực chứa TP.HCM và Hà Nội: 17 tỉnh) so với **vùng ngoại vi** (4 vùng còn lại: 46 tỉnh).", "Bảng 8, Hình 8"],
  ["`<biến>_L1`", "X_{i,t−1}", "8 biến giải thích", "Biến trễ một năm, tạo bằng `groupby('tinh').shift(1)` sau khi sắp xếp theo năm; năm đầu của mỗi tỉnh trống. Dùng để giảm nội sinh đồng thời.", "Bảng 6 cột (3)"],
], [1500, 2000, 1700, 2870, 1000]));
ch.push(GAP());
ch.push(P("Lưu ý về `EBI_rel`: mẫu số là trung bình của **các tỉnh có số liệu** trong năm (không phải 63 tỉnh), nên khi VECOM bổ sung tỉnh mới (từ 54 lên 58 tỉnh) trung bình có thể dịch nhẹ. Đây là lý do bài chỉ dùng EBI_rel để so sánh vị thế vùng, còn kiểm định hội tụ chính được tách theo hai giai đoạn có thang đo ổn định."));

// ===================== 3 =====================
ch.push(H("3. Chỉ số bất bình đẳng (RQ1 — Bảng 3, Bảng 4, Hình 3, Hình 4)", 1));
ch.push(P("Tất cả tính theo từng năm trên véc-tơ EBI của N_t tỉnh. Các chỉ số này **bất biến với phép nhân thang đo** (nhân mọi EBI với cùng một hằng số thì Gini, Theil, MLD, CV không đổi), nên vẫn có ý nghĩa dù VECOM đổi thang từ 2021; tuy nhiên chúng nhạy với thay đổi **hình dạng** phân phối, và đó là điều xảy ra năm 2021."));

ch.push(H("3.1. Hệ số Gini — hàm `gini`", 2));
ch.push(EQ("G = [2 Σ_{j=1..N} j·y_(j)] / [N Σ y_(j)] − (N + 1)/N"));
ch.push(P("với y_(1) ≤ y_(2) ≤ … ≤ y_(N) là EBI đã sắp tăng dần và *j* là thứ hạng. Đây là công thức Gini theo thứ hạng, tương đương với nửa chênh lệch trung bình tương đối |y_i − y_j| / (2N²ȳ) trong đề cương. G = 0 khi mọi tỉnh bằng nhau, tiến tới 1 khi tập trung tuyệt đối."));

ch.push(H("3.2. Chỉ số Theil-T và MLD — hàm `theil_T`, `mld`", 2));
ch.push(EQ("T = (1/N) Σ_i (y_i/ȳ) · ln(y_i/ȳ)          L (MLD) = (1/N) Σ_i ln(ȳ/y_i)"));
ch.push(P("Theil-T là chỉ số entropy tổng quát GE(1), nhạy hơn với đuôi trên (các tỉnh rất cao như Hà Nội, TP.HCM); MLD là GE(0), nhạy hơn với đuôi dưới. Bài báo cáo cả hai để kiểm tra vững."));

ch.push(H("3.3. Phân rã Theil và MLD theo 6 vùng — hàm `theil_decomp`, `mld_decomp`", 2));
ch.push(EQ("T = Σ_k s_k T_k  +  Σ_k s_k ln(ȳ_k/ȳ),      s_k = (N_k/N)(ȳ_k/ȳ)"));
ch.push(EQ("L = Σ_k p_k L_k  +  Σ_k p_k ln(ȳ/ȳ_k),      p_k = N_k/N"));
ch.push(...BUL([
  "T_k, L_k: Theil/MLD tính riêng trong vùng *k*; ȳ_k: EBI trung bình vùng; N_k: số tỉnh có EBI trong vùng.",
  "Số hạng thứ nhất = **phần trong nội bộ vùng** (within); số hạng thứ hai = **phần giữa các vùng** (between). Tỷ trọng giữa vùng (%) = between / T × 100.",
  "Trọng số s_k của Theil-T là tỷ trọng “thu nhập” (ở đây là tổng EBI) của vùng; trọng số p_k của MLD là tỷ trọng số tỉnh.",
  "Kiểm tra vững: tính lại toàn bộ sau khi loại Hà Nội và TP. Hồ Chí Minh khỏi mẫu (file `phu_theil_bo_HN_HCM.csv`).",
]));

ch.push(H("3.4. Hệ số biến thiên và độ lệch chuẩn của ln(EBI)", 2));
ch.push(EQ("CV_t = σ_t / ȳ_t          SDln_t = độ lệch chuẩn của ln(EBI_it) trong năm t"));
ch.push(P("CV là thước đo hội tụ sigma kinh điển; SDln là thước đo thay thế ít nhạy với giá trị cực đoan. Tính cho EBI và từng trụ cột (Hình 5)."));

ch.push(H("3.5. Trọng số trụ cột của EBI (phục hồi từ dữ liệu)", 2));
ch.push(P("EBI của VECOM là tổng có trọng số của các trụ cột nhưng trọng số không in kèm số liệu. Trọng số được phục hồi bằng hồi quy không hằng số EBI lên các trụ cột trong từng năm: R² = 1,000 với sai số tối đa 0,07 điểm (riêng 2023 là 2,2 điểm và 2024 là 0,65 điểm ở một tỉnh, do làm tròn của VECOM). Kết quả và giá trị làm tròn dùng trong phân rã:"));
ch.push(TBL(["Năm", "NNL&HT", "B2C", "B2B", "G2B"], [
  ["2017", "0,30", "0,30", "0,30", "0,10"],
  ["2018–2020", "0,20", "0,35", "0,35", "0,10"],
  ["2021–2024", "0,20", "0,40", "0,40", "— (bỏ)"],
], [1600, 1800, 1800, 1800, 2070]));
ch.push(GAP());

ch.push(H("3.6. Phân rã Gini theo trụ cột (Lerman – Yitzhaki, 1985) — hàm `gini_source_decomp`", 2));
ch.push(EQ("Y = Σ_k w_k X_k        G(Y) = Σ_k S_k · G_k · R_k"));
ch.push(...BUL([
  "X_k: điểm trụ cột *k*; w_k: trọng số ở mục 3.5; Y = EBI tái tạo từ các trụ cột.",
  "**S_k = w_k·X̄_k / Ȳ**: tỷ trọng của trụ cột trong EBI trung bình.",
  "**G_k**: hệ số Gini của thành phần w_k·X_k (bằng Gini của X_k vì Gini bất biến với phép nhân).",
  "**R_k = cov(w_k X_k, F(Y)) / cov(w_k X_k, F(w_k X_k))**: “tương quan Gini” giữa trụ cột với thứ hạng EBI, trong đó F(·) là hàm phân phối tích lũy thực nghiệm (thứ hạng chia N). R_k = 1 khi thứ tự các tỉnh theo trụ cột trùng hoàn toàn với thứ tự theo EBI.",
  "**Đóng góp (%) của trụ cột k = S_k G_k R_k / Σ_k S_k G_k R_k × 100**. Tổng Σ_k S_k G_k R_k bằng Gini của Y (sai khác nhỏ với Gini EBI công bố do làm tròn trọng số).",
]));

// ===================== 4 =====================
ch.push(H("4. Chỉ số hội tụ (RQ2 — Bảng 5, Bảng 5b, Hình 5, Hình 6, Hình 7)", 1));
ch.push(H("4.1. Hội tụ beta tuyệt đối và có điều kiện — hàm `beta_conv`", 2));
ch.push(EQ("g_i = (1/T) · ln(EBI_{i,T} / EBI_{i,0}) = α + β · ln(EBI_{i,0}) + γ′Z_i + ε_i"));
ch.push(...BUL([
  "g_i: tốc độ tăng EBI bình quân năm (dạng log) của tỉnh *i* từ năm gốc 0 đến năm cuối T; chỉ tính cho tỉnh có EBI ở **cả hai** năm.",
  "β < 0 và có ý nghĩa → tỉnh xuất phát thấp tăng nhanh hơn (hội tụ). Ước lượng bằng OLS, sai số chuẩn vững phương sai thay đổi HC1.",
  "**Tuyệt đối**: không có Z. **Có điều kiện**: Z gồm ln(GRDP/người), mật độ doanh nghiệp, tỷ lệ dân số thành thị, tỷ lệ lao động qua đào tạo, lấy giá trị **năm gốc**; riêng GRDP/người chỉ có từ 2018 nên với gốc 2017 dùng giá trị 2018.",
  "Ba cửa sổ: 2017–2020 (T = 3, thang cũ), 2021–2024 (T = 3, thang mới), 2017–2024 (T = 7) trên EBI tương đối (`EBI_rel`).",
]));
ch.push(H("4.2. Tốc độ hội tụ và thời gian bán hội tụ", 2));
ch.push(EQ("λ = − ln(1 + β·T) / T            t_½ = ln 2 / λ"));
ch.push(P("λ là tốc độ hội tụ hằng năm (từ nghiệm của mô hình tăng trưởng tân cổ điển: β = −(1 − e^(−λT))/T); t_½ là số năm để một nửa khoảng cách ban đầu được xóa bỏ. Chỉ tính khi 1 + βT > 0 và λ > 0; nếu β > 0 (phân kỳ) bảng ghi “—”."));
ch.push(H("4.3. Hội tụ beta dạng bảng động (kiểm tra vững)", 2));
ch.push(EQ("g_{i,t} = ln(EBI_rel_{i,t}) − ln(EBI_rel_{i,t−1}) = α_t + β · ln(EBI_rel_{i,t−1}) + ε_{i,t}"));
ch.push(P("Hồi quy trên toàn bộ các cặp năm liên tiếp, có hiệu ứng cố định năm α_t, sai số chuẩn cụm theo tỉnh; **loại quan sát tăng trưởng 2020→2021** vì bước nhảy này do đổi phương pháp. Với T = 1, λ = −ln(1 + β)."));
ch.push(H("4.4. Hội tụ sigma", 2));
ch.push(P("Dùng CV_t và SDln_t ở mục 3.4; giảm theo thời gian trong một giai đoạn có thang đo ổn định = hội tụ sigma. Tính riêng cho EBI, NNL&HT, B2C, B2B."));

// ===================== 5 =====================
ch.push(H("5. Mô hình hồi quy các yếu tố quyết định (RQ3 — Bảng 6, 7, 9)", 1));
ch.push(H("5.1. OLS gộp với hiệu ứng năm — hàm `fit_ols`", 2));
ch.push(EQ("ln(EBI_{i,t}) = β_0 + β′X_{i,t} + δ_t + ε_{i,t}"));
ch.push(...BUL([
  "X gồm 8 biến: `ln_grdp`, `dn_per_1000`, `urban_share_pct`, `trained_labor_pct`, `ln_freight`, `modern_retail_100k`, `ln_fdi_pc`, `hh_computer_pct` (cột (2) trở đi; cột (1) bỏ `hh_computer_pct` để có thêm năm 2018).",
  "δ_t: biến giả năm (`C(nam)`); hấp thụ mọi thay đổi chung theo năm, trong đó có việc đổi thang EBI 2021 (nếu đổi thang là phép nhân thì ln biến nó thành phép cộng và δ_t hấp thụ trọn).",
  "Sai số chuẩn **cụm theo tỉnh** (`cov_type='cluster'`): cho phép sai số của cùng một tỉnh tương quan qua các năm.",
  "Mẫu: mọi tỉnh–năm có đủ EBI và X; cột (3) dùng X trễ 1 năm; cột (4) chỉ 2021–2024.",
  "Đọc hệ số: với biến dạng mức (ví dụ `dn_per_1000`), tăng 1 đơn vị gắn với EBI thay đổi ≈ 100·(e^β − 1)% (β = 0,070 → +7,3%); với biến dạng ln, β là độ co giãn.",
]));
ch.push(H("5.2. Hệ số chuẩn hóa (cột (5) Bảng 6)", 2));
ch.push(EQ("z = (x − x̄) / σ_x  cho cả biến phụ thuộc và mọi biến giải thích, rồi ước lượng lại mô hình (2)"));
ch.push(P("Hệ số chuẩn hóa cho biết ln(EBI) thay đổi bao nhiêu độ lệch chuẩn khi biến giải thích tăng một độ lệch chuẩn, dùng để so sánh độ lớn tác động giữa các biến có đơn vị khác nhau."));
ch.push(H("5.3. Hệ số phóng đại phương sai VIF", 2));
ch.push(EQ("VIF_j = 1 / (1 − R²_j)"));
ch.push(P("R²_j là R² khi hồi quy biến giải thích *j* lên các biến giải thích còn lại (không gồm biến giả năm). VIF < 10 được coi là chấp nhận được; giá trị lớn nhất trong bài là 5,4 (mật độ doanh nghiệp). Hàm `variance_inflation_factor` của statsmodels."));
ch.push(H("5.4. Biến giả vùng và kiểm định Wald (Bảng 7)", 2));
ch.push(EQ("ln(EBI_{i,t}) = β_0 + β′X_{i,t} + Σ_{k≠ĐNB} θ_k · D_{k,i} + δ_t + ε_{i,t}"));
ch.push(P("D_{k,i} = 1 nếu tỉnh *i* thuộc vùng *k*; Đông Nam Bộ là nhóm gốc nên θ_k đo chênh lệch ln(EBI) của vùng *k* so với Đông Nam Bộ **sau khi** giữ X cố định. Kiểm định Wald H₀: θ_ĐBSH = θ_BTB&DHMT = θ_TN = θ_ĐBSCL = θ_TDMNPB = 0 với ma trận phương sai cụm, thống kê χ²(5). Cột (8) thay biến giả vùng bằng biến giả tỉnh (hiệu ứng cố định tỉnh, `C(tinh)`), khi đó chỉ còn biến thiên theo thời gian trong nội bộ tỉnh."));
ch.push(H("5.5. Mô hình cắt ngang 2024 (Bảng 9)", 2));
ch.push(EQ("ln(EBI_{i,2024}) = β_0 + β_1·Internet_i + β_2·Kỹ năng CNTT_i + β′X_i + ε_i"));
ch.push(P("OLS trên 58 tỉnh, sai số chuẩn HC1. Ba đặc tả: chỉ biến số; biến số + kinh tế (ln GRDP, mật độ DN, đô thị hóa); tỷ lệ hộ có máy tính + kinh tế + lao động qua đào tạo."));

// ===================== 6 =====================
ch.push(H("6. Phân tách Oaxaca – Blinder (Bảng 8, 8b, Hình 8) — hàm `oaxaca`, `oaxaca_boot`", 1));
ch.push(P("Mục tiêu: tách chênh lệch ln(EBI) trung bình giữa vùng lõi (L, `core` = 1) và vùng ngoại vi (N, `core` = 0) thành phần do khác biệt **nguồn lực** (X̄ khác nhau) và phần do khác biệt **hệ số** (cùng X nhưng chuyển hóa thành EBI khác nhau)."));
ch.push(EQ("Ȳ_L − Ȳ_N = (X̄_L − X̄_N)′β⋆   +   [ X̄_L′(β̂_L − β⋆) + X̄_N′(β⋆ − β̂_N) ]"));
ch.push(...BUL([
  "β̂_L, β̂_N: hệ số OLS ước lượng riêng trên từng nhóm (mô hình gồm 8 biến X và biến giả năm).",
  "**β\\* (hệ số tham chiếu)**: theo Jann (2008), lấy từ hồi quy trên mẫu gộp có thêm biến chỉ báo nhóm `core` (để tránh phần hệ số bị nhiễm bởi chính khác biệt nhóm); kiểm tra vững dùng β⋆ = β̂_L.",
  "Số hạng thứ nhất = **phần nguồn lực** (endowments), tách được theo từng biến: (X̄_{L,j} − X̄_{N,j})·β⋆_j; chia cho tổng khoảng cách và nhân 100 cho ra “% khoảng cách do biến j” (Bảng 8b, Hình 8).",
  "Số hạng trong ngoặc vuông = **phần hệ số** (coefficients), gồm cả chênh lệch hằng số; phần hệ số theo từng biến phụ thuộc vào gốc đo nên bài chỉ diễn giải ở tổng.",
  "**Sai số chuẩn bootstrap**: lặp 300 lần, mỗi lần rút có hoàn lại 63 tỉnh (giữ nguyên toàn bộ các năm của tỉnh được rút, tức lấy mẫu theo cụm), tính lại hai thành phần; SE = độ lệch chuẩn của 300 giá trị.",
  "Ba mẫu: 2019–2024 với 8 biến (chính), 2018–2024 với 7 biến (không có hộ có máy tính), 2021–2024 với 8 biến.",
]));

// ===================== 7 =====================
ch.push(H("7. Bảng tra nhanh: biến/chỉ số → hàm trong mã → file kết quả", 1));
ch.push(TBL(["Đại lượng", "Hàm / đoạn mã (01_phan_tich.py)", "File kết quả trong ket-qua/"], [
  ["freight_pc, modern_retail_100k, dn_moi_1000, ln_*, EBI_rel, *_rel, core, *_L1", "Mục “1. DỮ LIỆU”: `load_tonghop`, các dòng `df[...] =`", "panel_phan_tich.csv (và bản sao hinh-bai-bao/du_lieu/panel_63_tinh_2017_2024.csv)"],
  ["Gini, Theil-T, MLD, CV, SDln; phân rã theo vùng", "`gini`, `theil_T`, `mld`, `theil_decomp`, `mld_decomp`; mục 5", "bang3_theil.csv, phu_theil_vung.csv, phu_theil_bo_HN_HCM.csv"],
  ["Trọng số trụ cột; phân rã Gini theo trụ cột", "`WEIGHTS`, `gini_source_decomp`; mục 6", "bang4_gini_tru_cot.csv"],
  ["CV theo trụ cột (sigma)", "mục 7, biến `sig`", "bang5a_sigma.csv"],
  ["Beta tuyệt đối / có điều kiện, λ, t½; beta bảng động", "`beta_conv`; mục 7, biến `mp`", "bang5b_beta.csv, phu_beta_*.csv, ket_qua.json (khóa beta_panel)"],
  ["OLS gộp, hệ số chuẩn hóa, VIF, tương quan", "`fit_ols`, `tidy`; mục 8", "bang6_ols.csv, bang6_vif.csv, phu_tuong_quan.csv"],
  ["Cắt ngang 2024", "mục 8, biến `mc1`–`mc3`", "bang6b_2024_cross.csv"],
  ["Biến giả vùng, FE tỉnh, Wald", "mục 9, `m7a`–`m7c`, `wald_test`", "bang7_vung.csv, ket_qua.json (khóa wald_vung_M6)"],
  ["Oaxaca–Blinder, bootstrap", "`oaxaca`, `oaxaca_boot`; mục 10", "bang8_oaxaca.csv, bang8_oaxaca_chi_tiet.csv"],
], [3000, 3200, 2870]));
ch.push(GAP());
ch.push(P("Các hình lặp lại một phần các phép tính trên (Theil, CV, beta, Oaxaca) bằng mã độc lập trong `hinh-bai-bao\\tao_hinh.py` để thư mục hình tự chứa; kết quả đã được đối chiếu trùng khớp với các file CSV ở trên."));

const doc = new Document({
  creator: "Đào Thị Hồng Vân", title: "Mô tả công thức tính các biến và chỉ số",
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] }) },
    children: ch,
  }],
});
Packer.toBuffer(doc).then((buf) => { const out = path.join(ROOT, "MO-TA-CONG-THUC-TINH.docx"); fs.writeFileSync(out, buf); console.log("Đã ghi", out, (buf.length / 1024).toFixed(0), "KB"); });

// Tạo file Word hướng dẫn dùng script vẽ hình: hinh-bai-bao/HUONG-DAN-TAO-HINH.docx
// Chạy: NODE_PATH=<thư mục có node_modules docx> node scripts/05_huong_dan_hinh.js
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, HeadingLevel, ShadingType, LevelFormat, ImageRun, PageNumber, Footer } = require("docx");

const ROOT = path.resolve(__dirname, "..");
const DIR = path.join(ROOT, "hinh-bai-bao");
const FONT = "Times New Roman", MONO = "Consolas";
const PAGE_W = 11906, MARGIN = 1418, CW = PAGE_W - 2 * MARGIN;

const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 24, ...o });
const P = (parts, o = {}) => new Paragraph({
  children: (Array.isArray(parts) ? parts : [run(parts)]),
  alignment: o.align || AlignmentType.JUSTIFIED, spacing: { after: o.after ?? 120, line: 276 }, ...(o.extra || {}),
});
const H = (t, lv) => new Paragraph({ heading: lv === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: lv === 1 ? 28 : 24, color: "000000" })],
  spacing: { before: lv === 1 ? 280 : 180, after: 120 }, keepNext: true });
const CODE = (lines) => lines.map((l) => new Paragraph({
  children: [new TextRun({ text: l, font: MONO, size: 20 })], shading: { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" },
  spacing: { after: 0, line: 260 }, indent: { left: 400 },
}));
const BUL = (items) => items.map((it) => new Paragraph({
  children: Array.isArray(it) ? it : [run(it)], numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 },
}));
const NUM = (items, id) => items.map((it) => new Paragraph({
  children: Array.isArray(it) ? it : [run(it)], numbering: { reference: "num" + id, level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 },
}));
const thin = { style: BorderStyle.SINGLE, size: 6, color: "000000" }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function TBL(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0); const w = widths.map((x) => Math.round(x / total * CW));
  const cell = (t, i, hd, last) => new TableCell({
    children: String(t).split("\n").map((l) => new Paragraph({ children: [new TextRun({ text: l, font: l.startsWith("`") ? MONO : FONT, size: 20, bold: hd })], spacing: { after: 0, line: 250 } })),
    width: { size: w[i], type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 70, right: 70 },
    borders: { top: hd ? thin : none, bottom: hd || last ? thin : none, left: none, right: none },
    shading: hd ? { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" } : undefined,
  });
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: w, layout: "fixed", rows: [
    new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true, false)) }),
    ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false, ri === rows.length - 1)) })),
  ] });
}
function IMG(file, cm) {
  const buf = fs.readFileSync(path.join(DIR, "hinh", file)); const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const W = Math.round(cm * 37.8);
  return new Paragraph({ children: [new ImageRun({ type: "png", data: buf, transformation: { width: W, height: Math.round(W * h / w) } })], alignment: AlignmentType.CENTER, spacing: { before: 80, after: 40 } });
}
const CAP = (t) => new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 20, italics: true })], alignment: AlignmentType.CENTER, spacing: { after: 160 } });

const ch = [];
ch.push(new Paragraph({ children: [new TextRun({ text: "HƯỚNG DẪN CHẠY SCRIPT TẠO HÌNH CHO BÀI BÁO", bold: true, font: FONT, size: 30 })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "Bài báo: “Khoảng cách thương mại điện tử giữa các tỉnh, thành phố Việt Nam giai đoạn 2017–2024: mức độ, cấu trúc, hội tụ và các yếu tố quyết định” (tác giả Đào Thị Hồng Vân)", font: FONT, size: 22, italics: true })], alignment: AlignmentType.CENTER, spacing: { after: 60 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "Thư mục: detai\\hinh-bai-bao\\   —   cập nhật 05/09/2026", font: FONT, size: 22 })], alignment: AlignmentType.CENTER, spacing: { after: 240 } }));

ch.push(H("1. Thư mục này để làm gì", 1));
ch.push(P("Thư mục hinh-bai-bao là một bộ tự chứa: chỉ cần mở file Word này, cài thư viện Python một lần, rồi chạy một file lập trình duy nhất là toàn bộ 6 hình dùng trong bài báo được tạo lại vào thư mục con hinh. Mọi thứ cần thiết (dữ liệu đầu vào, mã nguồn, hướng dẫn) đều nằm trong thư mục này, không phụ thuộc vào các thư mục khác của dự án."));
ch.push(P("Cấu trúc thư mục:"));
ch.push(TBL(["Tên", "Loại", "Nội dung"], [
  ["tao_hinh.py", "Script Python", "File lập trình DUY NHẤT tạo 6 hình. Mỗi hình là một hàm riêng (hinh1_… đến hinh6_…), có chú thích tiếng Việt ở đầu mỗi hàm."],
  ["chay_tao_hinh.bat", "File chạy nhanh", "Nhấp đúp để chạy tao_hinh.py mà không cần gõ lệnh (Windows)."],
  ["du_lieu\\panel_63_tinh_2017_2024.csv", "Dữ liệu đầu vào", "Bảng dữ liệu 63 tỉnh × 8 năm (504 dòng, 56 cột) đã ghép EBI của VECOM với các chỉ tiêu Niên giám Thống kê. Xem mục 5."],
  ["hinh\\", "Thư mục kết quả", "Nơi script ghi 6 file PNG (được tạo tự động nếu chưa có). Có thể xóa cả thư mục rồi chạy lại."],
  ["HUONG-DAN-TAO-HINH.docx", "Hướng dẫn", "File này."],
], [2600, 1400, 5070]));
ch.push(P("", { after: 60 }));

ch.push(H("2. Cài đặt một lần", 1));
ch.push(P("Máy cần có Python 3.10 trở lên (khuyến nghị 3.11–3.14) và bốn thư viện: pandas, numpy, matplotlib, statsmodels. Kiểm tra và cài bằng cách mở PowerShell (hoặc CMD) và gõ:"));
ch.push(...CODE(["python --version", "pip install pandas numpy matplotlib statsmodels"]));
ch.push(P("Nếu máy đã chạy được các script trong detai\\scripts thì các thư viện này đã có sẵn; chỉ cần chạy lệnh thứ hai khi báo lỗi “No module named …”.", { extra: { spacing: { before: 120, after: 120 } } }));

ch.push(H("3. Cách chạy", 1));
ch.push(H("3.1. Cách nhanh nhất: nhấp đúp", 2));
ch.push(P("Nhấp đúp file chay_tao_hinh.bat. Một cửa sổ đen hiện ra, in đường dẫn 6 hình vừa tạo rồi dừng lại chờ bấm phím. Kết quả nằm trong thư mục hinh."));
ch.push(H("3.2. Chạy bằng dòng lệnh", 2));
ch.push(P("Mở PowerShell tại thư mục hinh-bai-bao (giữ Shift, nhấp chuột phải vào khoảng trống trong thư mục, chọn “Open PowerShell window here”), rồi gõ:"));
ch.push(...CODE(["python tao_hinh.py"]));
ch.push(P("Muốn tạo lại chỉ một vài hình (ví dụ Hình 3 và Hình 5), thêm số thứ tự hình phía sau:", { extra: { spacing: { before: 120, after: 120 } } }));
ch.push(...CODE(["python tao_hinh.py 3 5"]));
ch.push(H("3.3. Kết quả mong đợi trên màn hình", 2));
ch.push(...CODE([
  "Đọc ...\\du_lieu\\panel_63_tinh_2017_2024.csv: 445 quan sát tỉnh–năm có EBI, 58 tỉnh, 2017–2024",
  "  Hình 1 → ...\\hinh\\Hinh1_khung_phan_tich.png",
  "  Hình 2 → ...\\hinh\\Hinh2_ebi_vung.png",
  "  ...",
  "  Hình 6 → ...\\hinh\\Hinh6_oaxaca_dong_gop.png",
  "XONG. Mở thư mục 'hinh' để xem kết quả.",
]));
ch.push(P("Thời gian chạy khoảng 10–20 giây. Số quan sát 445 và 58 tỉnh là đúng: mỗi năm VECOM chỉ xếp hạng 54–58 tỉnh, 5 tỉnh miền núi (Cao Bằng, Hà Giang, Lai Châu, Điện Biên, Sơn La) không có EBI năm nào.", { extra: { spacing: { before: 120, after: 120 } } }));

ch.push(H("4. Sáu hình được tạo và vị trí trong bài báo", 1));
ch.push(TBL(["File trong thư mục hinh", "Tên hình trong bài", "Mục", "Dữ liệu và cách tính", "Hàm trong tao_hinh.py"], [
  ["Hinh1_khung_phan_tich.png", "Hình 1. Khung phân tích", "2.1", "Không dùng dữ liệu; sơ đồ khái niệm vẽ bằng các hộp và mũi tên.", "hinh1_khung_phan_tich"],
  ["Hinh2_ebi_vung.png", "Hình 2. EBI tương đối theo 6 vùng, 2017–2024", "4.1", "Cột EBI_rel (EBI của tỉnh / trung bình các tỉnh trong năm × 100), lấy trung bình giản đơn theo vùng (cột vung6) và năm.", "hinh2_ebi_vung"],
  ["Hinh3_theil.png", "Hình 3. Theil-T và phân rã theo vùng", "4.1", "Chỉ số Theil-T và MLD của EBI theo năm, phân rã thành phần giữa các vùng và trong vùng (hàm phan_ra_theil, phan_ra_mld).", "hinh3_theil"],
  ["Hinh4_sigma_cv.png", "Hình 4. Hội tụ sigma (CV)", "4.3", "Hệ số biến thiên = độ lệch chuẩn / trung bình của EBI, NNL_HT, B2C, B2B theo năm.", "hinh4_sigma"],
  ["Hinh5_beta.png", "Hình 5. Hội tụ beta tuyệt đối", "4.3", "Hồi quy (1/T)·ln(EBI cuối/EBI gốc) lên ln(EBI gốc) cho 2017–2020 và 2021–2024, sai số chuẩn HC1 (hàm hoi_tu_beta, dùng statsmodels).", "hinh5_beta"],
  ["Hinh6_oaxaca_dong_gop.png", "Hình 6. Đóng góp từng biến vào phần nguồn lực (Oaxaca–Blinder)", "4.6", "Phân tách khoảng cách ln(EBI) giữa vùng lõi (ĐNB + ĐBSH) và ngoại vi, mẫu 2019–2024, hệ số tham chiếu gộp (hàm oaxaca_chi_tiet).", "hinh6_oaxaca"],
], [1900, 2000, 500, 3170, 1500]));
ch.push(P("", { after: 40 }));
ch.push(P("Xem nhanh hai hình mẫu do script tạo ra:"));
ch.push(IMG("Hinh2_ebi_vung.png", 12)); ch.push(CAP("Hình 2 trong bài báo (file Hinh2_ebi_vung.png)"));
ch.push(IMG("Hinh5_beta.png", 13)); ch.push(CAP("Hình 5 trong bài báo (file Hinh5_beta.png)"));

ch.push(H("5. Dữ liệu đầu vào", 1));
ch.push(P("File du_lieu\\panel_63_tinh_2017_2024.csv (UTF-8, dấu phẩy ngăn cột, dấu chấm thập phân) là bản sao của detai\\ket-qua\\panel_phan_tich.csv do script detai\\scripts\\01_phan_tich.py tạo ra từ EBI_NGTK_Panel_2017_2024.xlsx và bốn bảng Niên giám trong data\\csv\\tonghop. Các cột mà tao_hinh.py sử dụng:"));
ch.push(TBL(["Cột", "Ý nghĩa", "Dùng cho hình"], [
  ["tinh, nam, vung6", "Tên tỉnh (63 tỉnh trước sáp nhập), năm 2017–2024, mã 6 vùng kinh tế – xã hội (ĐBSH, TDMNPB, BTB&DHMT, TN, ĐNB, ĐBSCL)", "2, 3, 5, 6"],
  ["EBI, NNL_HT, B2C, B2B", "Chỉ số TMĐT tổng hợp và ba trụ cột (điểm, VECOM); trống ở tỉnh không được xếp hạng", "2–6"],
  ["EBI_rel", "EBI tương đối = EBI / trung bình các tỉnh có số liệu trong năm × 100", "2"],
  ["ln_ebi", "ln(EBI)", "6"],
  ["ln_grdp, dn_per_1000, urban_share_pct, trained_labor_pct", "ln(GRDP/người); doanh nghiệp/1.000 dân; % dân số thành thị; % lao động 15+ qua đào tạo (Niên giám)", "6"],
  ["ln_freight, modern_retail_100k, ln_fdi_pc, hh_computer_pct", "ln(hàng hóa vận chuyển đường bộ/người); siêu thị + TTTM/100.000 dân; ln(1 + FDI lũy kế/người); % hộ có máy tính (Niên giám)", "6"],
], [2600, 5070, 1400]));
ch.push(P("Các cột còn lại (dân số, thu nhập, HDI, biến trễ _L1, …) được giữ nguyên để tiện phân tích thêm nhưng script không dùng.", { extra: { spacing: { before: 120, after: 120 } } }));
ch.push(P([run("Muốn cập nhật dữ liệu (ví dụ thêm năm 2025): "), run("chỉ cần thay file CSV bằng file mới có cùng tên cột; script tự tính lại EBI_rel nếu cột này thiếu. Vạch dọc đánh dấu năm đổi phương pháp EBI được đặt ở hằng số NAM_DOI_PHUONG_PHAP đầu script.")]));

ch.push(H("6. Tùy chỉnh thường gặp", 1));
ch.push(P("Tất cả thông số nằm ở mục “0. THIẾT LẬP CHUNG” đầu file tao_hinh.py; sửa bằng bất kỳ trình soạn thảo nào (Notepad, VS Code) rồi chạy lại."));
ch.push(TBL(["Muốn", "Sửa ở đâu"], [
  ["Đổi độ phân giải ảnh", "DPI = 200 (tạp chí thường yêu cầu 300)"],
  ["Đổi màu 6 vùng / kiểu nét", "Danh sách MAU và NET (theo thứ tự VUNG); giữ thứ tự cố định để cùng một vùng có cùng màu trên mọi hình"],
  ["Đổi font chữ trên hình", "Danh sách font trong vòng lặp for ten_font in [...]; font phải có dấu tiếng Việt (Arial, Segoe UI, DejaVu Sans, Times New Roman)"],
  ["Đổi tên vùng hiển thị", "Từ điển TEN_VUNG"],
  ["Đổi định nghĩa vùng lõi ở Hình 6", "VUNG_LOI = [\"ĐNB\", \"ĐBSH\"]"],
  ["Đổi/bớt tỉnh được gắn nhãn ở Hình 5", "Danh sách gan_nhan trong hàm hinh5_beta"],
  ["Đổi kích thước một hình", "figsize=(rộng, cao) tính bằng inch trong hàm tương ứng"],
  ["Xuất thêm định dạng PDF/SVG", "Trong phần “4. CHẠY”, thêm dòng fig.savefig(duong_dan.replace('.png', '.pdf'))"],
], [3200, 5870]));

ch.push(H("7. Lỗi thường gặp", 1));
ch.push(TBL(["Thông báo / hiện tượng", "Nguyên nhân và cách xử lý"], [
  ["'python' is not recognized…", "Chưa cài Python hoặc chưa thêm vào PATH. Cài từ python.org, tích “Add python.exe to PATH”, mở lại PowerShell."],
  ["ModuleNotFoundError: No module named 'statsmodels' (hoặc pandas, matplotlib)", "Chạy: pip install pandas numpy matplotlib statsmodels"],
  ["FileNotFoundError: …du_lieu\\panel_63_tinh_2017_2024.csv", "File dữ liệu bị đổi tên hoặc di chuyển; đặt lại đúng tên trong thư mục du_lieu, hoặc sửa hằng FILE_DU_LIEU."],
  ["Chữ tiếng Việt trên hình hiện ô vuông", "Máy thiếu font có dấu; cài Arial/Segoe UI hoặc đổi danh sách font (mục 6)."],
  ["Cửa sổ .bat hiện chữ lỗi mã hóa", "Chỉ ảnh hưởng hiển thị thông báo, không ảnh hưởng hình. Có thể chạy bằng PowerShell theo mục 3.2."],
  ["Hình khác với hình trong bài", "Đã sửa dữ liệu hoặc thông số. Để tái tạo đúng bản gốc, dùng lại file CSV gốc và các thiết lập mặc định."],
], [3600, 5470]));

ch.push(H("8. Liên hệ với các file khác của đề tài", 1));
ch.push(...BUL([
  [run("detai\\scripts\\01_phan_tich.py", { font: MONO, size: 22 }), run(": tính toàn bộ bảng số liệu (Bảng 1–9) và tạo file panel_phan_tich.csv; không vẽ hình.")],
  [run("detai\\scripts\\02_bang_docx.py", { font: MONO, size: 22 }), run(" và "), run("03_build_docx.js", { font: MONO, size: 22 }), run(": định dạng bảng và dựng file Word bài báo; file Word lấy hình trực tiếp từ hinh-bai-bao\\hinh, nên sau khi chạy lại tao_hinh.py chỉ cần dựng lại bài là hình được cập nhật.")],
  [run("detai\\ket-qua\\", { font: MONO, size: 22 }), run(": các bảng kết quả dạng CSV dùng để đối chiếu số liệu với hình.")],
]));
ch.push(P("Thứ tự tái tạo toàn bộ từ đầu: 01_phan_tich.py → sao chép ket-qua\\panel_phan_tich.csv sang hinh-bai-bao\\du_lieu\\panel_63_tinh_2017_2024.csv → hinh-bai-bao\\tao_hinh.py → 02_bang_docx.py → 03_build_docx.js.", { extra: { spacing: { before: 120 } } }));

const doc = new Document({
  creator: "Đào Thị Hồng Vân", title: "Hướng dẫn chạy script tạo hình cho bài báo",
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "num1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] }) },
    children: ch,
  }],
});
Packer.toBuffer(doc).then((buf) => { const out = path.join(DIR, "HUONG-DAN-TAO-HINH.docx"); fs.writeFileSync(out, buf); console.log("Đã ghi", out, (buf.length / 1024).toFixed(0), "KB"); });

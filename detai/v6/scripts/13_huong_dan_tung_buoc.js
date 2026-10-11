// 13 - Dựng file Word HUONG-DAN-TINH-TUNG-BUOC.docx: giải thích từng phương pháp trong bài cho người mới,
//      song ngữ thuật ngữ Việt – chuyên môn, có ví dụ số nhỏ (6 tỉnh) tính từng bước.
//      Số liệu ví dụ lấy từ ket-qua/vi_du_tung_buoc.json (script 12); số trang lấy từ ket-qua/_so_trang_bai_bao.json.
// Chạy trong thư mục v6:  node scripts/13_huong_dan_tung_buoc.js
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, HeadingLevel, ShadingType, LevelFormat, PageNumber, Footer, TableOfContents, PageBreak } = require("docx");

const ROOT = path.resolve(__dirname, "..");
const VD = JSON.parse(fs.readFileSync(path.join(ROOT, "ket-qua", "vi_du_tung_buoc.json"), "utf8"));
const TR = JSON.parse(fs.readFileSync(path.join(ROOT, "ket-qua", "so_trang_bai_bao.json"), "utf8"));
const FONT = "Times New Roman", MONO = "Consolas";
const PAGE_W = 11906, MARGIN = 1418, CW = PAGE_W - 2 * MARGIN;

// ---------- tiện ích định dạng (giống 07) ----------
const LIT = ""; // dấu sao chữ nghĩa (dấu sao ý nghĩa thống kê) viết là {*} trong nguồn
function runs(text, base = {}) {
  text = String(text).split("{*}").join(LIT);
  const out = []; const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g; let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRunLit({ text: text.slice(last, m.index), font: FONT, size: 24, ...base }));
    const s = m[0];
    if (s.startsWith("**")) out.push(new TextRunLit({ text: s.slice(2, -2), bold: true, font: FONT, size: 24, ...base }));
    else if (s.startsWith("*")) out.push(new TextRunLit({ text: s.slice(1, -1), italics: true, font: FONT, size: 24, ...base }));
    else out.push(new TextRunLit({ text: s.slice(1, -1), font: MONO, size: 21, ...base }));
    last = m.index + s.length;
  }
  if (last < text.length) out.push(new TextRunLit({ text: text.slice(last), font: FONT, size: 24, ...base }));
  return out.map((r) => { const o = r.options ?? r; return r; }).map((r) => r);
}
// đổi LIT -> "*" ngay trong constructor TextRun (bọc lại lớp)
const _TR = TextRun;
function TextRunLit(opts) { if (opts && typeof opts.text === "string") opts.text = opts.text.split(LIT).join("*"); return new _TR(opts); }
const P = (t, o = {}) => new Paragraph({ children: runs(t), alignment: AlignmentType.JUSTIFIED, spacing: { after: o.after ?? 120, line: 276 } });
const H = (t, lv) => new Paragraph({ heading: lv === 1 ? HeadingLevel.HEADING_1 : lv === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
  children: runs(t, { bold: true, size: lv === 1 ? 28 : lv === 2 ? 25 : 24, color: "000000" }),
  spacing: { before: lv === 1 ? 320 : 200, after: 120 }, keepNext: true });
const EQ = (t) => new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 24, italics: false })], alignment: AlignmentType.CENTER,
  spacing: { before: 60, after: 120 }, shading: { type: ShadingType.CLEAR, fill: "F7F7F7", color: "auto" } });
const BUL = (items) => items.map((it) => new Paragraph({ children: runs(it), numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 } }));
const STEP = (items) => items.map((it) => new Paragraph({ children: runs(it), numbering: { reference: "num", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 276 } }));
const thin = { style: BorderStyle.SINGLE, size: 6, color: "000000" }, none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
function TBL(header, rows, widths, opt = {}) {
  const total = widths.reduce((a, b) => a + b, 0); const w = widths.map((x) => Math.round(x / total * CW));
  const cell = (t, i, hd, last, boldRow) => new TableCell({
    children: String(t).split("\n").map((l) => new Paragraph({ children: runs(l, { size: 19, bold: hd || boldRow }), spacing: { after: 0, line: 245 },
      alignment: (i > 0 && opt.numRight && !hd) ? AlignmentType.RIGHT : AlignmentType.LEFT })),
    width: { size: w[i], type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 70, right: 70 },
    borders: { top: hd ? thin : none, bottom: hd || last ? thin : none, left: none, right: none },
    shading: hd ? { type: ShadingType.CLEAR, fill: "DCE6F1", color: "auto" } : (boldRow ? { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" } : undefined),
  });
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: w, layout: "fixed", rows: [
    new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true, false)) }),
    ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false, ri === rows.length - 1, opt.boldLast && ri === rows.length - 1)) })),
  ] });
}
const GAP = () => new Paragraph({ children: [], spacing: { after: 100 } });
// Hộp "Nói đơn giản" / "Trong bài báo"
const BOX = (title, lines, fill) => [
  new Paragraph({ children: [new TextRun({ text: title, bold: true, font: FONT, size: 22 })], shading: { type: ShadingType.CLEAR, fill, color: "auto" },
    spacing: { before: 100, after: 0 }, indent: { left: 300, right: 300 }, keepNext: true }),
  ...lines.map((l, i) => new Paragraph({ children: runs(l, { size: 22 }), shading: { type: ShadingType.CLEAR, fill, color: "auto" }, alignment: AlignmentType.JUSTIFIED,
    spacing: { after: i === lines.length - 1 ? 140 : 40, line: 260 }, indent: { left: 300, right: 300 } })),
];
const DON_GIAN = (lines) => BOX("Nói đơn giản", lines, "EAF3E3");
const TRONG_BAI = (lines) => BOX("Con số tương ứng trong bài báo", lines, "FFF4DC");
const LUU_Y = (lines) => BOX("Lưu ý / bẫy thường gặp", lines, "FCE9E9");
const VI = (x, n = 4) => Number(x).toLocaleString("vi-VN", { minimumFractionDigits: n, maximumFractionDigits: n });
const v1 = (x) => VI(x, 1), v2 = (x) => VI(x, 2), v3 = (x) => VI(x, 3), v4 = (x) => VI(x, 4);
const tr = (k) => TR[k] ? `tr. ${TR[k]}` : "";
// hai muc: gop trang neu trung, tranh in "15, 15"
const tr2 = (a, b) => (TR[a] === TR[b] ? `tr. ${TR[a]}` : `tr. ${TR[a]}, ${TR[b]}`);

const ch = [];
// ================= TRANG BÌA =================
ch.push(new Paragraph({ children: [new TextRun({ text: "HƯỚNG DẪN ĐỌC HIỂU CÁCH TÍNH TỪNG BƯỚC", bold: true, font: FONT, size: 34 })], alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 120 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "cho bài báo “Khoảng cách thương mại điện tử giữa các tỉnh, thành phố Việt Nam giai đoạn 2017–2024”", font: FONT, size: 24, italics: true })], alignment: AlignmentType.CENTER, spacing: { after: 400 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "Dành cho người chưa quen thống kê: mỗi phương pháp được trình bày theo cùng một khung — (1) nằm ở đâu trong bài, (2) trả lời câu hỏi gì, (3) công thức và nghĩa từng ký hiệu, (4) tính từng bước trên ví dụ 6 tỉnh với số thật, (5) đọc kết quả, (6) con số thật trong bài, (7) bẫy cần tránh.", font: FONT, size: 22 })], alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, indent: { left: 600, right: 600 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "Đọc kèm sổ Excel HUONG-DAN-TINH-TUNG-BUOC.xlsx (cùng bộ số, mỗi ô là công thức Excel có thể bấm xem) và MO-TA-CONG-THUC-TINH.docx (bản kỹ thuật đầy đủ, đối chiếu với mã nguồn).", font: FONT, size: 22 })], alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, indent: { left: 600, right: 600 } }));
ch.push(new Paragraph({ children: [new TextRun({ text: "Số trang dẫn trong tài liệu này là số trang của bản in PDF xuất từ Bai-bao-EBI-63-tinh-2017-2024.docx (25 trang, khổ A4) ngày 12/09/2026; nếu bài được chỉnh sửa, chạy lại scripts/15_do_so_trang.py rồi dựng lại tài liệu này để số trang khớp trở lại.", font: FONT, size: 20, italics: true })], alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, indent: { left: 600, right: 600 } }));
ch.push(new Paragraph({ children: [new PageBreak()] }));
ch.push(new Paragraph({ children: [new TextRun({ text: "MỤC LỤC", bold: true, font: FONT, size: 26 })], spacing: { after: 120 } }));
ch.push(new TableOfContents("Mục lục", { hyperlink: true, headingStyleRange: "1-2" }));
ch.push(GAP());

// ================= 0. BẢN ĐỒ =================
ch.push(H("0. Bản đồ: bài báo tính những gì, ở đâu, bằng phương pháp nào", 1));
ch.push(P("Bài báo trả lời ba câu hỏi nghiên cứu (RQ). Mỗi câu hỏi dùng một nhóm công cụ. Bảng dưới đây là “mục lục theo phép tính”: đọc từ trái sang phải để biết một con số trong bài sinh ra từ đâu."));
ch.push(TBL(["Câu hỏi của bài", "Phép tính", "Trả lời câu hỏi thường ngày nào?", "Ở mục / trang nào", "Bảng, hình, phương trình", "Xem mục nào của tài liệu này"], [
  ["RQ1 – Khoảng cách TMĐT giữa các tỉnh lớn đến đâu, do đâu?", "EBI tương đối", "Tỉnh này đứng ở đâu so với mức trung bình cả nước?", `3.1 (${tr("3.1.")}), 4.1 (${tr("4.1.")})`, `Bảng 2 (${tr("Bảng 2.")}), Hình 2 (${tr("Hình 2.")})`, "Mục 2"],
  ["", "Gini, CV", "Các tỉnh chênh nhau nhiều hay ít?", `3.3 (${tr("3.3.")}), 4.1`, `Bảng 3 (${tr("Bảng 3.")})`, "Mục 3, 4"],
  ["", "Theil-T và phân rã giữa vùng / trong vùng", "Chênh lệch là do các VÙNG khác nhau, hay do các tỉnh TRONG CÙNG VÙNG khác nhau?", "3.3, 4.1", `Phương trình (1) (${tr("(1)")}), Bảng 3, Hình 3–4 (${tr2("Hình 3.", "Hình 4.")})`, "Mục 5"],
  ["", "Phân rã Gini theo trụ cột", "Trong ba trụ cột NNL&HT, B2C, B2B, trụ cột nào gây ra khoảng cách nhiều nhất?", `3.3, 4.2 (${tr("4.2.")})`, `Bảng 4 (${tr("Bảng 4.")})`, "Mục 6"],
  ["RQ2 – Khoảng cách đang thu hẹp hay doãng ra?", "Hội tụ sigma (CV theo năm)", "Độ chênh lệch có giảm dần theo năm không?", `3.3, 4.3 (${tr("4.3.")})`, `Hình 5 (${tr("Hình 5.")})`, "Mục 4"],
  ["", "Hội tụ beta", "Tỉnh đi sau có tăng nhanh hơn tỉnh đi trước không? Bao lâu thì bắt kịp một nửa?", "3.3, 4.3", `Phương trình (2) (${tr("(2)")}), Bảng 5, 5b (${tr2("Bảng 5.", "Bảng 5b.")}), Hình 6–7 (${tr2("Hình 6.", "Hình 7.")})`, "Mục 7"],
  ["RQ3 – Yếu tố nào quyết định EBI và khoảng cách vùng?", "Hồi quy OLS gộp (+VIF, biến giả vùng, hiệu ứng cố định)", "Tỉnh có nhiều doanh nghiệp hơn, giàu hơn, đô thị hơn… thì EBI cao hơn bao nhiêu?", `3.3, 4.4 (${tr("4.4.")}), 4.5 (${tr("4.5.")})`, `Phương trình (3) (${tr("(3)")}), Bảng 6 (${tr("Bảng 6.")}), Bảng 7 (${tr("Bảng 7.")}), Bảng 9 (${tr("Bảng 9.")})`, "Mục 8"],
  ["", "Phân tách Oaxaca–Blinder", "Vùng lõi hơn vùng ngoại vi vì CÓ NHIỀU NGUỒN LỰC hơn, hay vì DÙNG NGUỒN LỰC HIỆU QUẢ hơn?", `3.3, 4.6 (${tr("4.6.")})`, `Phương trình (4) (${tr("(4)")}), Bảng 8, 8b (${tr2("Bảng 8.", "Bảng 8b.")}), Hình 8 (${tr("Hình 8.")})`, "Mục 9"],
], [2100, 1700, 2400, 1250, 2000, 1000]));
ch.push(GAP());
ch.push(P("**Ký hiệu chung dùng suốt tài liệu:** *i* = tỉnh; *t* = năm; *k* = vùng (hoặc trụ cột, tuỳ mục); *N* = số tỉnh có số liệu trong năm (54–58); ȳ (đọc “y gạch”) = trung bình; σ = độ lệch chuẩn; ln = logarit tự nhiên (nút LN trên máy tính / hàm LN trong Excel); Σ = “cộng tất cả lại”."));

// ================= 1. DỮ LIỆU VÍ DỤ =================
ch.push(H("1. Bộ số ví dụ dùng chung: 6 tỉnh, số thật", 1));
ch.push(P("Để mọi phép tính trong tài liệu này có thể **kiểm tra bằng máy tính bỏ túi**, tất cả ví dụ dùng chung 6 tỉnh dưới đây (số lấy đúng từ dữ liệu của bài, file `ket-qua/panel_phan_tich.csv`). Hai tỉnh đầu thuộc **vùng lõi** (Đồng bằng sông Hồng, Đông Nam Bộ), bốn tỉnh sau thuộc **vùng ngoại vi**."));
ch.push(TBL(["Tỉnh", "Vùng (Niên giám)", "Nhóm ví dụ", "EBI 2024", "NNL&HT 2024", "B2C 2024", "B2B 2024", "EBI 2021", "DN/1.000 dân 2024"],
  VD.bang_goc.map((g) => [g.tinh, g.vung, g.nhom, v1(g.EBI_2024), v1(g.NNL_HT), v1(g.B2C), v1(g.B2B), v2(g.EBI_2021), v2(g.dn_per_1000)]),
  [1300, 1600, 900, 800, 900, 800, 800, 900, 1100], { numRight: true }));
ch.push(GAP());
ch.push(...LUU_Y([
  "Kết quả tính trên 6 tỉnh **không phải** là kết quả của bài báo (bài tính trên 54–58 tỉnh mỗi năm). Chúng chỉ để thấy phép tính chạy thế nào. Cuối mỗi mục có hộp màu vàng ghi con số thật của bài để đối chiếu.",
  "EBI 2021 và 2024 nằm trên **cùng thang đo mới** (VECOM đổi cách tính từ 2021), nên so sánh được với nhau. Không được so sánh thô EBI 2020 với EBI 2021 — đây là điểm gãy quan trọng nhất của toàn bài (Mục 3.1 của bài, " + tr("3.1.") + ").",
]));

// ================= 2. EBI TƯƠNG ĐỐI =================
const er = VD.ebi_rel;
ch.push(H("2. EBI tương đối — “tỉnh này bằng bao nhiêu phần trăm mức trung bình cả nước?”", 1));
ch.push(H("2.1. Ở đâu trong bài", 2));
ch.push(P(`Định nghĩa ở Mục 3.1 (${tr("3.1.")}); dùng ở Bảng 2 (${tr("Bảng 2.")}), Hình 2 (${tr("Hình 2.")}), dòng “2017–2024 (EBI tương đối)” của Bảng 5 và mô tả khoảng cách lõi – ngoại vi ở Mục 4.3.`));
ch.push(H("2.2. Công thức", 2));
ch.push(EQ("EBI tương đối của tỉnh i năm t  =  EBI_it / ( trung bình EBI của các tỉnh có số liệu năm t ) × 100"));
ch.push(...DON_GIAN(["Lấy điểm của tỉnh chia cho điểm trung bình của tất cả các tỉnh cùng năm, nhân 100. Kết quả 100 = đúng bằng trung bình; 150 = cao hơn trung bình 50%; 70 = chỉ bằng 70% trung bình.",
  "Vì sao cần nó: VECOM đổi thang điểm năm 2021 (trung bình cả nước rơi từ 41,6 xuống 8,5 điểm). Nếu so điểm thô thì mọi tỉnh đều “tụt”. Chia cho trung bình cùng năm sẽ **xoá tác động của thang đo**, chỉ còn lại vị thế tương đối."]));
ch.push(H("2.3. Tính từng bước (năm 2024)", 2));
ch.push(...STEP([
  `**Bước 1 – Lấy trung bình cả nước.** Năm 2024 có ${er.so_tinh_2024} tỉnh được VECOM chấm điểm; cộng 58 điểm số rồi chia 58 được **${v3(er.trung_binh_2024)}** điểm. (Chú ý: chia cho số tỉnh *có số liệu*, không phải 63.)`,
  "**Bước 2 – Chia và nhân 100{*}{*} cho từng tỉnh:",
]));
ch.push(TBL(["Tỉnh", "EBI 2024", "Phép tính", "EBI tương đối", "Đọc"], er.vi_du.map((x) => [x.tinh, v1(x.EBI), `${v1(x.EBI)} / ${v3(er.trung_binh_2024)} × 100`, v1(x.EBI_rel),
  x.EBI_rel >= 100 ? `cao hơn trung bình ${v1(x.EBI_rel - 100)} điểm` : `thấp hơn trung bình ${v1(100 - x.EBI_rel)} điểm`]), [1300, 900, 2200, 1200, 2400], { numRight: true }));
ch.push(GAP());
ch.push(P("**Bước 3 – Trung bình vùng** (Bảng 2): cộng EBI tương đối của các tỉnh trong vùng rồi chia cho số tỉnh của vùng. Cột “2017–2020” của Bảng 2 là trung bình của mọi cặp tỉnh–năm thuộc vùng trong 4 năm đó."));
ch.push(...TRONG_BAI([`Bảng 2 (${tr("Bảng 2.")}): Đông Nam Bộ 166,8 và Đồng bằng sông Hồng 135,1 điểm năm 2024; ba vùng thấp nhất 71–78 điểm. Tỷ số vùng cao nhất/thấp nhất = 166,8/71,2 = 2,34.`,
  "Hà Nội 84/23,086 = 364 điểm tức gấp 3,64 lần trung bình, TP. Hồ Chí Minh 87/23,086 = 377 điểm tức 3,77 lần — đúng câu “EBI của TP. Hồ Chí Minh và Hà Nội gấp 3,6–3,8 lần trung bình các tỉnh năm 2024” ở Mục 4.1."]));

// ================= 3. GINI =================
const gi = VD.gini;
ch.push(H("3. Hệ số Gini — “các tỉnh chênh nhau nhiều hay ít?”", 1));
ch.push(H("3.1. Ở đâu trong bài", 2));
ch.push(P(`Mục 3.3 (${tr("3.3.")}); kết quả ở Bảng 3 (${tr("Bảng 3.")}) cột “Gini” và Bảng 4 (${tr("Bảng 4.")}) cho từng trụ cột.`));
ch.push(H("3.2. Công thức (dạng thứ hạng, đúng như hàm `gini` trong mã)", 2));
ch.push(EQ("G = 2·Σ( j × y_(j) ) / ( N × Σ y_(j) )  −  (N + 1)/N"));
ch.push(...BUL(["y_(j): EBI của các tỉnh **sau khi sắp xếp từ thấp đến cao**; j = thứ hạng (1 = thấp nhất, N = cao nhất).",
  "G = 0 khi mọi tỉnh bằng nhau; G tiến tới 1 khi một tỉnh chiếm toàn bộ. Với EBI cấp tỉnh, G khoảng 0,13–0,15 (thang cũ) là chênh lệch vừa; 0,35 (2021) là chênh lệch lớn."]));
ch.push(...DON_GIAN(["Gini hỏi: nếu bốc ngẫu nhiên hai tỉnh, chúng chênh nhau trung bình bao nhiêu **so với mức trung bình**? Công thức thứ hạng ở trên chỉ là cách tính nhanh của ý đó: tỉnh càng cao (j lớn) càng được “nhân” nhiều, nên nếu điểm dồn vào các tỉnh top thì tổng Σ(j·y) lớn và G lớn.",
  "Một tính chất then chốt: nhân mọi EBI với cùng một số (đổi thang) thì G **không đổi** — nhờ vậy bài vẫn so sánh được Gini trước và sau 2021 (dù thận trọng)."]));
ch.push(H("3.3. Tính từng bước (6 tỉnh, 2024)", 2));
ch.push(...STEP(["**Bước 1 – Sắp xếp EBI từ thấp đến cao và đánh thứ hạng j.**", "**Bước 2 – Nhân j với y rồi cộng lại** (cột cuối)."]));
ch.push(TBL(["Thứ hạng j", "Tỉnh", "EBI (y)", "j × y"], [...gi.sap_xep.map((s, i) => [String(s.j), VD.bang_goc.slice().sort((a, b) => a.EBI_2024 - b.EBI_2024)[i].tinh, v1(s.y), v1(s.j_y)]),
  ["Tổng", "", `Σy = ${v1(gi.tong_y)}`, `Σ(j·y) = ${v1(gi.tong_j_y)}`]], [1200, 1800, 1500, 1500], { numRight: true, boldLast: true }));
ch.push(GAP());
ch.push(...STEP([
  `**Bước 3 – Số hạng thứ nhất:** 2 × ${v1(gi.tong_j_y)} / (${gi.n} × ${v1(gi.tong_y)}) = ${v1(2 * gi.tong_j_y)} / ${v1(gi.n * gi.tong_y)} = **${v4(gi.so_hang_1)}**.`,
  `**Bước 4 – Số hạng thứ hai:** (N + 1)/N = 7/6 = **${v4(gi.so_hang_2)}**.`,
  `**Bước 5 – Trừ:** G = ${v4(gi.so_hang_1)} − ${v4(gi.so_hang_2)} = **${v4(gi.G)}**.`,
  `**Kiểm tra bằng cách hiểu “chênh lệch từng cặp”:** lấy mọi cặp tỉnh, tính |y_i − y_j|, cộng hết, chia cho 2N²ȳ → cũng ra ${v4(gi.G_kiem_tra_cach_2)}. Hai cách cho cùng kết quả, nên bạn có thể dùng cách nào dễ hình dung hơn.`,
]));
ch.push(...TRONG_BAI([`Bảng 3: Gini EBI 2024 = 0,245 (58 tỉnh); giảm đều từ 0,353 (2021). Thang cũ 2017–2020: 0,129–0,145. Ví dụ 6 tỉnh cho ${v3(gi.G)}, cao hơn vì mẫu nhỏ gồm cả Hà Nội (84) lẫn Thái Nguyên (20).`,
  "Bảng 4: Gini của từng trụ cột tính y hệt, chỉ thay cột EBI bằng cột trụ cột. Năm 2024 B2B có Gini cao nhất (0,340), nhưng không phải năm nào cũng vậy: 2021 là B2C (0,480) và 2023 là NNL&HT (0,299) — xem Mục 4.2 của bài."]));

// ================= 4. CV, SIGMA =================
const cv = VD.cv, sg = VD.sigma;
ch.push(H("4. Hệ số biến thiên CV và hội tụ sigma — “độ chênh lệch có giảm dần theo năm không?”", 1));
ch.push(H("4.1. Ở đâu trong bài", 2));
ch.push(P(`Định nghĩa ở Mục 3.3 (${tr("3.3.")}); Bảng 3 cột “CV”; Hình 5 (${tr("Hình 5.")}) vẽ CV theo năm cho EBI và ba trụ cột; kết quả đọc ở Mục 4.3 (${tr("4.3.")}).`));
ch.push(H("4.2. Công thức", 2));
ch.push(EQ("CV_t = σ_t / ȳ_t          (độ lệch chuẩn chia trung bình, tính riêng từng năm)"));
ch.push(EQ("Hội tụ sigma  ⇔  CV giảm dần theo thời gian (trong một giai đoạn có thang đo ổn định)"));
ch.push(...DON_GIAN(["Độ lệch chuẩn σ đo “các tỉnh cách xa trung bình bao nhiêu điểm”. Nhưng 5 điểm chênh lệch khi trung bình là 40 thì nhỏ, còn khi trung bình là 8 thì rất lớn. Chia cho trung bình để so sánh công bằng giữa các năm: đó là CV.",
  "“Sigma” chỉ là tên gọi (σ = độ lệch chuẩn). Hội tụ sigma = các tỉnh xích lại gần nhau. Bài còn dùng thêm độ lệch chuẩn của ln(EBI) làm thước đo phụ, ít nhạy với tỉnh cực đoan."]));
ch.push(H("4.3. Tính từng bước (6 tỉnh, 2024)", 2));
ch.push(...STEP([
  `**Bước 1 – Trung bình:** ȳ = (${VD.bang_goc.map((g) => v1(g.EBI_2024)).join(" + ")}) / 6 = **${v3(cv.mean)}**.`,
  "**Bước 2 – Độ lệch từng tỉnh so với trung bình và bình phương:**",
]));
ch.push(TBL(["Tỉnh", "y", "y − ȳ", "(y − ȳ)²"], [...VD.bang_goc.map((g, i) => [g.tinh, v1(g.EBI_2024), v3(cv.lech[i]), v2(cv.binh_phuong[i])]),
  ["Tổng", "", "", v2(cv.tong_bp)]], [1500, 1000, 1200, 1300], { numRight: true, boldLast: true }));
ch.push(GAP());
ch.push(...STEP([
  `**Bước 3 – Độ lệch chuẩn mẫu:** σ = √( ${v2(cv.tong_bp)} / (6 − 1) ) = √${v2(cv.tong_bp / 5)} = **${v3(cv.sd)}**. (Chia N − 1 = 5, không chia 6: quy ước “mẫu” của thống kê; Excel: STDEV.S.)`,
  `**Bước 4 – CV:** ${v3(cv.sd)} / ${v3(cv.mean)} = **${v3(cv.CV)}**.`,
  `**Bước 5 – So với năm 2021 (cùng 6 tỉnh):** ȳ₂₀₂₁ = ${v3(sg.mean_2021)}, σ₂₀₂₁ = ${v3(sg.sd_2021)} → CV₂₀₂₁ = **${v3(sg.CV_2021)}**. CV giảm từ ${v3(sg.CV_2021)} xuống ${v3(sg.CV_2024)} → **có hội tụ sigma** trong nhóm 6 tỉnh này.`,
  `**Thước đo phụ:** SD của ln(EBI) 2024 = ${v3(cv.SD_ln)} (lấy ln từng EBI: ${cv.ln_y.map(v3).join("; ")} rồi tính độ lệch chuẩn như trên).`,
]));
ch.push(...TRONG_BAI(["Mục 4.3 và Bảng 3: CV EBI giảm 0,348 → 0,283 (2017–2020) và 1,253 → 0,608 (2021–2024); SD ln(EBI) 0,528 → 0,410. Hình 5: theo trụ cột, B2C thu hẹp mạnh nhất (CV 1,95 → 0,43), B2B vẫn 0,85, NNL&HT lại tăng 0,61 → 0,70.",
  "Cách đọc: có hội tụ sigma trong từng giai đoạn, nhưng mức phân tán theo thang mới (0,6) vẫn gấp đôi thang cũ (0,3)."]));
ch.push(...LUU_Y(["Không nối CV 2020 (0,283) với CV 2021 (1,253) để kết luận “phân kỳ”: bước nhảy đó do đổi cách tính EBI. Đây là lý do bài luôn tách hai giai đoạn."]));

// ================= 5. THEIL =================
const th = VD.theil, pr = th.phan_ra;
ch.push(H("5. Chỉ số Theil-T và phân rã giữa vùng / trong vùng — “chênh lệch nằm ở đâu?”", 1));
ch.push(H("5.1. Ở đâu trong bài", 2));
ch.push(P(`Phương trình (1) ở Mục 3.3 (${tr("(1)")}); Bảng 3 (${tr("Bảng 3.")}) các cột Theil-T, giữa các vùng, trong vùng, tỷ trọng; Hình 3 và Hình 4 (${tr2("Hình 3.", "Hình 4.")}); diễn giải ở Mục 4.1 (${tr("4.1.")}). Đây là phép tính trả lời giả thuyết H1 và cho kết quả **trái với H1{*}{*} (chênh lệch trong vùng lớn hơn giữa vùng).`));
ch.push(H("5.2. Công thức", 2));
ch.push(EQ("Theil-T:   T = (1/N) · Σ_i (y_i/ȳ) · ln(y_i/ȳ)"));
ch.push(EQ("Phân rã:   T = Σ_k s_k·T_k   +   Σ_k s_k·ln(ȳ_k/ȳ),      s_k = (N_k/N)·(ȳ_k/ȳ)"));
ch.push(...BUL(["Số hạng thứ nhất = **phần trong vùng** (within): tổng Theil riêng của từng vùng, có trọng số s_k.", "Số hạng thứ hai = **phần giữa vùng** (between): chỉ dùng trung bình của từng vùng, coi như mọi tỉnh trong vùng bằng nhau.",
  "s_k = tỷ trọng “tổng EBI” của vùng k trong tổng EBI cả nước (số tỉnh × trung bình vùng, chia tương ứng).", "**Tỷ trọng giữa vùng (%) = between / T × 100{*}{*} — con số bài báo nhấn mạnh (31–37%)."]));
ch.push(...DON_GIAN(["Theil giống Gini ở chỗ đo chênh lệch, nhưng có một ưu điểm: **tách được thành hai phần cộng lại đúng bằng tổng**. Tưởng tượng chia cả nước thành 6 “lớp học” (vùng). Chênh lệch điểm cả trường = chênh lệch **giữa điểm trung bình các lớp** + chênh lệch **giữa các học sinh trong cùng lớp**.",
  "Bài phát hiện phần “trong lớp” chiếm 63–69%: tức là Hà Nội cách xa Nam Định, Thái Bình (cùng vùng) còn hơn là vùng Đồng bằng sông Hồng cách xa vùng Tây Nguyên."]));
ch.push(H("5.3. Tính từng bước (6 tỉnh, 2 nhóm Lõi / Ngoại vi, 2024)", 2));
ch.push(...STEP([`**Bước 1 – Trung bình chung** ȳ = ${v3(th.mean)} (đã tính ở Mục 4).`, "**Bước 2 – Với từng tỉnh, tính y/ȳ, lấy ln, rồi nhân hai số đó:**"]));
ch.push(TBL(["Tỉnh", "Nhóm", "y", "y/ȳ", "ln(y/ȳ)", "(y/ȳ)·ln(y/ȳ)"], [...th.buoc.map((b) => [b.tinh, VD.nhom[b.tinh], v1(b.y), v4(b.y_tren_mu), v4(b.ln), v4(b.tich)]),
  ["Tổng", "", "", "", "", v4(th.tong)]], [1400, 1000, 800, 1000, 1100, 1400], { numRight: true, boldLast: true }));
ch.push(GAP());
ch.push(...STEP([
  `**Bước 3 – Theil-T** = ${v4(th.tong)} / 6 = **${v4(th.T)}**. (Tỉnh trên trung bình cho số hạng dương, tỉnh dưới trung bình cho số hạng âm nhưng nhỏ hơn về trị tuyệt đối, nên tổng luôn ≥ 0.)`,
  "**Bước 4 – Tính riêng cho từng nhóm:** N_k, trung bình nhóm ȳ_k, trọng số s_k, Theil trong nhóm T_k (tính lại công thức Bước 2–3 nhưng chỉ với các tỉnh của nhóm và dùng ȳ_k thay ȳ).",
]));
ch.push(TBL(["Nhóm k", "N_k", "ȳ_k", "s_k = (N_k/N)(ȳ_k/ȳ)", "T_k", "s_k·T_k (trong)", "ln(ȳ_k/ȳ)", "s_k·ln(ȳ_k/ȳ) (giữa)"],
  [...pr.nhom.map((x) => [x.nhom, String(x.n), v3(x.mean), `(${x.n}/6)(${v3(x.mean)}/${v3(th.mean)}) = ${v4(x.s_k)}`, v4(x.T_k), v4(x.s_T), v4(x.ln_mu_ratio), v4(x.s_ln)]),
  ["Tổng", "", "", "", "", `within = ${v4(pr.within)}`, "", `between = ${v4(pr.between)}`]], [1000, 500, 800, 2200, 800, 1200, 900, 1300], { numRight: true, boldLast: true }));
ch.push(GAP());
ch.push(...STEP([
  `**Bước 5 – Kiểm tra:** within + between = ${v4(pr.within)} + ${v4(pr.between)} = ${v4(pr.tong)} = T ở Bước 3 ✓. Phân rã Theil luôn cộng lại đúng bằng tổng — đó là lý do bài chọn Theil chứ không phải Gini để tách vùng.`,
  `**Bước 6 – Tỷ trọng giữa nhóm** = ${v4(pr.between)} / ${v4(th.T)} × 100 = **${v1(pr.ty_trong_giua)}%**. Trong ví dụ này phần giữa nhóm chiếm phần lớn vì hai tỉnh lõi (84 và 51) tách rất xa bốn tỉnh còn lại; với 58 tỉnh và 6 vùng thực, bài cho kết quả ngược lại (36,9%).`,
]));
ch.push(...TRONG_BAI(["Bảng 3, năm 2024: T = 0,128 = 0,047 (giữa vùng) + 0,081 (trong vùng); tỷ trọng giữa vùng 36,9%. Suốt 8 năm tỷ trọng này chỉ 31–37% (Hình 4), tức chênh lệch trong vùng luôn gấp ~2 lần chênh lệch giữa vùng.",
  "Mục 4.1 còn tách phần trong vùng theo từng vùng (s_k·T_k): Đông Nam Bộ 0,031 và Đồng bằng sông Hồng 0,029 — chính hai vùng chứa TP.HCM và Hà Nội tạo ra phần lớn chênh lệch trong vùng.",
  "Kiểm tra vững: bỏ Hà Nội và TP.HCM ra thì T 2024 giảm từ 0,128 xuống 0,057 (cột cuối Bảng 3) nhưng tỷ trọng giữa vùng vẫn ~34%. MLD (cột “Tỷ trọng giữa vùng, MLD”) là chỉ số anh em, tính bằng (1/N)·Σ ln(ȳ/y_i), nhạy hơn với tỉnh thấp, cho tỷ trọng 37–44% — vẫn dưới 50%."]));

// ================= 6. GINI TRỤ CỘT =================
const gn = VD.gini_nguon;
ch.push(H("6. Phân rã Gini theo trụ cột — “trụ cột nào gây ra khoảng cách?”", 1));
ch.push(H("6.1. Ở đâu trong bài", 2));
ch.push(P(`Công thức ở Mục 3.3 (${tr("3.3.")}); kết quả ở Bảng 4 (${tr("Bảng 4.")}), diễn giải ở Mục 4.2 (${tr("4.2.")}) để kiểm định H2 (khoảng cách ở tầng giao dịch lớn hơn tầng hạ tầng).`));
ch.push(H("6.2. Công thức", 2));
ch.push(EQ("EBI = w₁·NNL&HT + w₂·B2C + w₃·B2B       (từ 2021: w = 0,2 / 0,4 / 0,4)"));
ch.push(EQ("G(EBI) = Σ_k  S_k · G_k · R_k"));
ch.push(...BUL(["**S_k** = tỷ trọng của trụ cột k trong EBI trung bình = trung bình(w_k·X_k) / trung bình(EBI). Ba S_k cộng lại bằng 1.",
  "**G_k** = Gini của thành phần w_k·X_k (tính đúng như Mục 3; nhân hằng số không đổi Gini nên bằng Gini của X_k).",
  "**R_k** = “tương quan Gini” = cov(w_k·X_k, hạng EBI) / cov(w_k·X_k, hạng của chính nó). R_k = 1 nếu thứ tự tỉnh theo trụ cột trùng khớp thứ tự theo EBI; R_k nhỏ nếu trụ cột đó xếp hạng khác EBI.",
  "**Đóng góp (%) của trụ cột k** = S_k·G_k·R_k / Σ(S·G·R) × 100."]));
ch.push(...DON_GIAN(["EBI là tổng có trọng số của ba trụ cột, nên bất bình đẳng của EBI cũng “góp” từ ba nguồn. Một trụ cột góp nhiều khi nó (a) chiếm tỷ trọng lớn trong EBI (S), (b) tự nó chênh lệch nhiều giữa các tỉnh (G), và (c) tỉnh nào cao ở trụ cột đó cũng cao ở EBI (R).",
  "Trọng số w không được VECOM công bố; bài phục hồi bằng cách hồi quy EBI lên các trụ cột không có hằng số (R² = 1,000) — xem Mục 3.5 của MO-TA-CONG-THUC-TINH.docx."]));
ch.push(H("6.3. Tính từng bước (6 tỉnh, 2024)", 2));
ch.push(...STEP(["**Bước 1 – Nhân trọng số và tái tạo EBI:** w₁X₁ = 0,2·NNL&HT; w₂X₂ = 0,4·B2C; w₃X₃ = 0,4·B2B; cộng ba cột được EBI tái tạo (so với EBI công bố để kiểm tra)."]));
ch.push(TBL(["Tỉnh", "0,2·NNL&HT", "0,4·B2C", "0,4·B2B", "EBI tái tạo", "EBI công bố", "Hạng EBI / N"],
  VD.bang_goc.map((g, i) => [g.tinh, v2(gn.tru_cot[0].wX[i]), v2(gn.tru_cot[1].wX[i]), v2(gn.tru_cot[2].wX[i]), v2(gn.Y_tai_tao[i]), v1(gn.EBI_cong_bo[i]), v3(gn.rank_Y[i])]),
  [1400, 1100, 1000, 1000, 1100, 1100, 1100], { numRight: true }));
ch.push(GAP());
ch.push(...STEP(["**Bước 2 – Với từng trụ cột tính S, G, R rồi nhân ba số:**"]));
ch.push(TBL(["Trụ cột", "S_k", "G_k", "R_k", "S·G·R", "Đóng góp (%)"], [...gn.tru_cot.map((x) => [x.tru_cot, v4(x.S), v4(x.G), v4(x.R), v4(x.tich), v1(x.dong_gop_pct)]),
  ["Tổng", "1,0000", "", "", v4(gn.tong_tich), "100,0"]], [1500, 1200, 1200, 1200, 1300, 1400], { numRight: true, boldLast: true }));
ch.push(GAP());
ch.push(...STEP([`**Bước 3 – Kiểm tra:** Σ(S·G·R) = ${v4(gn.tong_tich)} phải bằng Gini của EBI tái tạo = ${v4(gn.G_Y)} ✓ (sai khác chỉ do làm tròn trọng số).`,
  `**Đọc:** trong 6 tỉnh, B2B đóng góp ${v1(gn.tru_cot[2].dong_gop_pct)}% vào Gini của EBI — vừa vì trọng số 0,4, vừa vì B2B chênh lệch mạnh (Hà Nội 85 so với Thái Nguyên 14).`]));
ch.push(...TRONG_BAI(["Bảng 4, 2024 (58 tỉnh): Gini EBI 0,245; đóng góp NNL&HT 24,4%, B2C 29,8%, B2B 45,9%. Giai đoạn 2018–2020: hai trụ cột giao dịch (B2C + B2B) góp 78–80%, NNL&HT chỉ 15–16% → ủng hộ H2: khoảng cách ở tầng kết quả lớn hơn tầng tiếp cận.",
  "Giai đoạn 2017–2020 còn trụ cột thứ tư G2B (trọng số 0,1), đóng góp dưới 6%."]));

// ================= 7. BETA =================
const be = VD.beta;
ch.push(H("7. Hội tụ beta — “tỉnh đi sau có tăng nhanh hơn không? Bao lâu bắt kịp một nửa?”", 1));
ch.push(H("7.1. Ở đâu trong bài", 2));
ch.push(P(`Phương trình (2) ở Mục 3.3 (${tr("(2)")}); Bảng 5 và 5b (${tr2("Bảng 5.", "Bảng 5b.")}); Hình 6 và Hình 7 (${tr2("Hình 6.", "Hình 7.")}) là biểu đồ điểm với đường hồi quy; diễn giải ở Mục 4.3 (${tr("4.3.")}) cho H3.`));
ch.push(H("7.2. Công thức", 2));
ch.push(EQ("g_i = (1/T)·ln( EBI_{i,T} / EBI_{i,0} )  =  α + β·ln(EBI_{i,0}) + γ′Z_i + ε_i"));
ch.push(EQ("λ = −ln(1 + β·T)/T  (tốc độ hội tụ mỗi năm)          t_½ = ln 2 / λ  (số năm để xoá một nửa khoảng cách)"));
ch.push(...BUL(["g_i: tốc độ tăng EBI bình quân mỗi năm của tỉnh i (dạng log), từ năm gốc 0 đến năm cuối T.", "ln(EBI_{i,0}): mức xuất phát. **β < 0{*}{*} = tỉnh xuất phát thấp tăng nhanh hơn = hội tụ; **β > 0{*}{*} = phân kỳ.",
  "**Tuyệt đối**: không có Z. **Có điều kiện**: thêm Z = đặc trưng năm gốc (ln GRDP/người, mật độ DN, % đô thị, % lao động qua đào tạo) — hỏi “sau khi tính đến nền tảng kinh tế, tỉnh thấp có còn tăng nhanh hơn không?”.",
  "T = 3 cho mỗi giai đoạn (2017→2020, 2021→2024)."]));
ch.push(...DON_GIAN(["Vẽ mỗi tỉnh thành một chấm: trục ngang = điểm xuất phát (ln EBI năm đầu), trục dọc = tốc độ tăng mỗi năm. Nếu chấm bên trái (xuất phát thấp) nằm cao hơn chấm bên phải, đường thẳng khớp qua các chấm sẽ **dốc xuống** → β âm → tỉnh nghèo TMĐT đang đuổi kịp. Hình 6, 7 chính là bức tranh này.",
  "λ và t_½ chỉ là cách đổi β sang ngôn ngữ dễ hiểu: “mỗi năm khoảng cách xoá được λ%” và “mất t_½ năm để xoá một nửa”."]));
ch.push(H("7.3. Tính từng bước (6 tỉnh, 2021 → 2024, hội tụ tuyệt đối)", 2));
ch.push(...STEP(["**Bước 1 – Với từng tỉnh:** lấy ln của EBI hai năm, trừ nhau, chia T = 3 → tốc độ tăng g."]));
ch.push(TBL(["Tỉnh", "EBI 2021 (y₀)", "EBI 2024 (y₁)", "x = ln y₀", "ln y₁", "g = (ln y₁ − ln y₀)/3", "Đọc g"],
  be.buoc.map((b) => [b.tinh, v2(b.y0), v1(b.y1), v4(b.ln_y0), v4(b.ln_y1), v4(b.g), `≈ +${v1(b.g * 100)}%/năm`]), [1300, 1000, 1000, 1000, 900, 1500, 1200], { numRight: true }));
ch.push(GAP());
ch.push(...STEP([
  `**Bước 2 – Hồi quy đường thẳng g theo x** (OLS, chi tiết cách tính ở Mục 8): x̄ = ${v4(be.mean_x)}, ḡ = ${v4(be.mean_g)}; Σ(x−x̄)(g−ḡ) = ${v4(be.Sxy)}; Σ(x−x̄)² = ${v4(be.Sxx)}.`,
  `**β = ${v4(be.Sxy)} / ${v4(be.Sxx)} = ${v4(be.beta)}**; α = ḡ − β·x̄ = ${v4(be.alpha)}; R² = ${v3(be.R2)}.`,
  `**Bước 3 – Đọc dấu:** β = ${v4(be.beta)} < 0 → tỉnh xuất phát thấp (Thái Nguyên, Lâm Đồng: ln EBI ≈ 1,9) tăng 36–45%/năm, tỉnh xuất phát cao (Hà Nội: 4,0) chỉ tăng 14%/năm → **hội tụ**.`,
  `**Bước 4 – Tốc độ hội tụ:** λ = −ln(1 + (${v4(be.beta)})×3)/3 = −ln(${v4(1 + be.beta * 3)})/3 = **${v4(be.lam)}** ≈ ${v1(be.lam_pct)}%/năm.`,
  `**Bước 5 – Thời gian bán hội tụ:** ln 2 / ${v4(be.lam)} = 0,6931 / ${v4(be.lam)} = **${v1(be.half_life)} năm**.`,
]));
ch.push(...TRONG_BAI([`Bảng 5 (56 tỉnh, 2021–2024): β = −0,0888 (SE 0,0180, p < 0,001), λ = 10,3%/năm, bán hội tụ 6,7 năm, R² = 0,528. Giai đoạn 2017–2020: β = −0,0425, λ = 4,6%/năm, bán hội tụ 15,2 năm. Ví dụ 6 tỉnh cho β dốc hơn (${v3(be.beta)}) vì thiếu các tỉnh trung bình.`,
  "Bảng 5, dòng “2017–2024 (EBI tương đối)”: β = +0,0705 — DƯƠNG. Không phải phân kỳ thật mà vì nối hai thang đo: bài dùng dòng này để chứng minh không được nối thô hai giai đoạn.",
  "Bảng 5, hội tụ có điều kiện: β mất ý nghĩa (0,0025 với p = 0,94; −0,0492 với p = 0,18) → sau khi tính đến GRDP, mật độ DN…, tỉnh thấp không còn tự động tăng nhanh hơn. Bài diễn giải là “bắt kịp có chọn lọc”.",
  "Dòng cuối Bảng 5 (bảng động): dùng mọi cặp năm liên tiếp, T = 1, có hiệu ứng năm, bỏ cặp 2020→2021; β = −0,1057."]));
ch.push(...LUU_Y(["Sai số chuẩn (SE) và p-value trong Bảng 5 dùng phương pháp vững HC1 (không giả định phương sai sai số bằng nhau); Excel thường không có sẵn, nên sổ Excel kèm theo chỉ tính β, α, R².",
  "Một phần β âm có thể chỉ là “hồi quy về trung bình” do sai số đo lường của chỉ số khảo sát — bài tự nêu ở cuối Mục 4.3."]));

// ================= 8. OLS =================
const ol = VD.ols, vf = VD.vif;
ch.push(H("8. Hồi quy OLS gộp — “yếu tố nào làm EBI cao hơn, và bao nhiêu?”", 1));
ch.push(H("8.1. Ở đâu trong bài", 2));
ch.push(P(`Phương trình (3) ở Mục 3.3 (${tr("(3)")}); Bảng 6 (${tr("Bảng 6.")}) là bảng chính, Bảng 9 (${tr("Bảng 9.")}) là cắt ngang 2024, Bảng 7 (${tr("Bảng 7.")}) thêm biến giả vùng và hiệu ứng cố định tỉnh; diễn giải ở Mục 4.4 và 4.5 (${tr2("4.4.", "4.5.")}) cho H4.`));
ch.push(H("8.2. Công thức", 2));
ch.push(EQ("ln(EBI_{i,t}) = β₀ + β₁·X₁ + β₂·X₂ + … + β₈·X₈ + δ_t + ε_{i,t}"));
ch.push(...BUL(["8 biến X: ln GRDP/người; mật độ DN (DN/1.000 dân); % dân thành thị; % lao động qua đào tạo; ln hàng hoá vận chuyển/người; siêu thị + TTTM/100.000 dân; ln(1 + FDI/người); % hộ có máy tính.",
  "δ_t = **hiệu ứng cố định năm**: một hằng số riêng cho mỗi năm, “hút” mọi thay đổi chung (kể cả việc đổi thang EBI 2021).",
  "OLS = chọn các β sao cho tổng bình phương sai số (khoảng cách dọc từ điểm tới đường) nhỏ nhất. **Sai số chuẩn cụm theo tỉnh**: thừa nhận 6–7 quan sát của cùng một tỉnh không độc lập với nhau.",
  "**Đọc hệ số** khi biến phụ thuộc là ln: X dạng mức tăng 1 đơn vị → EBI đổi ≈ 100·(e^β − 1)%; X dạng ln → β là % đổi của EBI khi X tăng 1%."]));
ch.push(...DON_GIAN(["Hồi quy trả lời: “giữ mọi thứ khác như nhau, tỉnh có thêm 1 doanh nghiệp trên 1.000 dân thì EBI cao hơn bao nhiêu %?”. Từ khoá là **giữ mọi thứ khác như nhau** — đó là điều mà nhìn bảng số thô không làm được.",
  "Vì vậy trong Bảng 6, GRDP không có ý nghĩa: không phải vì tỉnh giàu không có EBI cao, mà vì khi đã tính đến mật độ DN (hai biến tương quan 0,64) thì GRDP không giải thích thêm được gì."]));
ch.push(H("8.3. Tính từng bước: phiên bản 1 biến để thấy cơ chế (6 tỉnh, 2024)", 2));
ch.push(P("Mô hình rút gọn: ln(EBI) = b₀ + b₁·(mật độ DN). Với một biến, OLS có công thức tay:"));
ch.push(EQ("b₁ = Σ(x−x̄)(y−ȳ) / Σ(x−x̄)²          b₀ = ȳ − b₁·x̄"));
ch.push(TBL(["Tỉnh", "x = DN/1.000 dân", "y = ln(EBI)", "ŷ = b₀ + b₁x", "phần dư y − ŷ"], ol.buoc.map((b) => [b.tinh, v2(b.x), v4(b.ln_ebi), v4(b.du_bao), v4(b.phan_du)]), [1500, 1400, 1300, 1400, 1400], { numRight: true }));
ch.push(GAP());
ch.push(...STEP([
  `**Bước 1:** x̄ = ${v3(ol.mean_x)}; ȳ = ${v4(ol.mean_y)}. **Bước 2:** Σ(x−x̄)(y−ȳ) = ${v4(ol.Sxy)}; Σ(x−x̄)² = ${v4(ol.Sxx)}.`,
  `**Bước 3:** b₁ = ${v4(ol.Sxy)} / ${v4(ol.Sxx)} = **${v4(ol.b1)}**; b₀ = ${v4(ol.mean_y)} − ${v4(ol.b1)} × ${v3(ol.mean_x)} = **${v4(ol.b0)}**.`,
  `**Bước 4 – Dự báo và phần dư** (hai cột cuối bảng): phần dư là phần đường thẳng chưa giải thích được; R² = 1 − Σ(phần dư)²/Σ(y−ȳ)² = **${v3(ol.R2)}** → mật độ DN một mình giải thích ${v1(ol.R2 * 100)}% chênh lệch ln(EBI) giữa 6 tỉnh.`,
  `**Bước 5 – Đọc hệ số:** b₁ = ${v4(ol.b1)} → thêm 1 DN/1.000 dân gắn với EBI cao hơn e^${v4(ol.b1)} − 1 = **${v1(ol.doc_he_so_pct)}%**.`,
  `**Bước 6 – Có ý nghĩa thống kê không?** Sai số chuẩn của b₁ = ${v4(ol.se)} (Excel: LINEST); t = b₁/SE = ${v2(ol.t)}; p = ${v4(ol.p)} (phân phối t, bậc tự do N − 2 = 4). Quy ước dấu sao: p < 0,01 {*}{*}{*}, < 0,05 {*}{*}, < 0,10 {*}. Ở đây SE là loại “cổ điển”; bài báo dùng SE cụm theo tỉnh (Bảng 6) và HC1 (Bảng 5, 9) — cùng ý nghĩa, chỉ khác cách tính độ không chắc chắn để không giả định sai số đồng đều.`,
]));
ch.push(H("8.4. Ba thứ đi kèm trong Bảng 6 và 7", 2));
ch.push(...BUL([
  `**VIF (đa cộng tuyến):** VIF_j = 1/(1 − R²_j), với R²_j là R² khi hồi quy biến j lên các biến X còn lại. Ví dụ hồi quy mật độ DN lên ln GRDP trong 6 tỉnh cho R² = ${v3(vf.R2_dn_tren_grdp)} → VIF = 1/(1 − ${v3(vf.R2_dn_tren_grdp)}) = ${v2(vf.VIF)}. VIF < 10 là chấp nhận được; bài: lớn nhất 5,37 (mật độ DN).`,
  "**Hệ số chuẩn hoá** (cột (5) Bảng 6): trừ trung bình, chia độ lệch chuẩn cho cả y lẫn mọi X rồi hồi quy lại; hệ số đọc là “X tăng 1 độ lệch chuẩn thì y tăng bao nhiêu độ lệch chuẩn”, dùng để so độ lớn giữa các biến khác đơn vị. Mật độ DN: 0,481, gấp nhiều lần các biến khác.",
  "**Biến giả vùng** (Bảng 7): thêm 5 biến 0/1 (Đông Nam Bộ làm gốc). Cột (6) chỉ có vùng: bốn vùng ngoại vi có ln(EBI) thấp hơn ĐNB 0,42–0,62 điểm logarit, tức EBI thấp hơn 34–46% (đổi bằng e^β − 1; nhớ rằng 0,62 điểm logarit KHÔNG phải 62%). Cột (7) thêm 8 biến X: các hệ số vùng co về −0,05…+0,12 và mất ý nghĩa → khoảng cách vùng gần như được giải thích trọn bởi cấu trúc kinh tế. **Kiểm định Wald** hỏi “5 hệ số vùng có đồng thời bằng 0 không?”: χ²(5) = 17,47, p = 0,004 → vẫn bác bỏ, nhưng phần còn lại nhỏ.",
  "**Hiệu ứng cố định tỉnh** (cột (8)): thay 5 biến giả vùng bằng 62 biến giả tỉnh; chỉ còn dùng sự thay đổi *trong nội bộ* mỗi tỉnh qua các năm. Mật độ DN vẫn dương mạnh (0,253) — kết quả không phụ thuộc vào việc so sánh tỉnh này với tỉnh khác.",
]));
ch.push(...TRONG_BAI(["Bảng 6, cột (2), 331 quan sát 2019–2024: mật độ DN 0,070{*}{*}{*} (SE 0,007) → ≈ +7,3% EBI mỗi DN thêm trên 1.000 dân; ln GRDP 0,054 (không ý nghĩa); hộ có máy tính 0,002 (p = 0,54); siêu thị + TTTM −0,033{*}{*}; R² 0,936 (cao chủ yếu nhờ biến giả năm). VIF lớn nhất 5,37.",
  "Bảng 9 (58 tỉnh, 2024): tỷ lệ người dùng Internet và kỹ năng CNTT đứng riêng giải thích 64% phương sai; thêm biến kinh tế thì vẫn có ý nghĩa nhưng nhỏ đi 2/3, mật độ DN nổi lên 0,057{*}{*}{*}."]));
ch.push(...LUU_Y(["Hệ số hồi quy là **tương quan có điều kiện**, không phải nhân quả: EBI cao cũng có thể hút doanh nghiệp về. Bài giảm bớt bằng biến trễ 1 năm (cột (3)) và hiệu ứng cố định tỉnh, nhưng tự nhận là hạn chế ở Mục 5."]));

// ================= 9. OAXACA =================
const ox = VD.oaxaca;
ch.push(H("9. Phân tách Oaxaca–Blinder — “lõi hơn ngoại vi vì có nhiều nguồn lực hơn, hay vì dùng hiệu quả hơn?”", 1));
ch.push(H("9.1. Ở đâu trong bài", 2));
ch.push(P(`Phương trình (4) ở Mục 3.3 (${tr("(4)")}); Bảng 8 và 8b (${tr("Bảng 8.")}), Hình 8 (${tr("Hình 8.")}); diễn giải ở Mục 4.6 (${tr("4.6.")}) cho H5 — kết luận mạnh nhất của bài: 79% khoảng cách là do nguồn lực, riêng mật độ doanh nghiệp 69%.`));
ch.push(H("9.2. Công thức", 2));
ch.push(EQ("Ȳ_L − Ȳ_N  =  (X̄_L − X̄_N)′β*   +   [ X̄_L′(β_L − β*) + X̄_N′(β* − β_N) ]"));
ch.push(...BUL(["Ȳ_L, Ȳ_N: trung bình ln(EBI) của vùng lõi (L) và ngoại vi (N). X̄: trung bình các biến X của từng vùng.",
  "β_L, β_N: hệ số hồi quy chạy **riêng** trên từng vùng. β*: hệ số **tham chiếu** — bài lấy từ hồi quy gộp cả hai vùng có thêm biến nhóm (Jann, 2008); kiểm tra vững dùng β* = β_L.",
  "Số hạng thứ nhất = **phần nguồn lực** (endowments): “nếu ngoại vi có X như lõi thì khoảng cách thu hẹp bao nhiêu”. Tách được theo từng biến → Bảng 8b, Hình 8.",
  "Số hạng trong ngoặc = **phần hệ số** (coefficients): cùng X nhưng chuyển hoá thành EBI khác nhau (bao gồm cả chênh lệch hằng số)."]));
ch.push(...DON_GIAN(["Hai vùng chênh nhau 0,455 điểm ln (≈ 58% EBI). Câu hỏi: nếu bốn vùng ngoại vi có mật độ doanh nghiệp, logistics, FDI… **bằng** vùng lõi, khoảng cách còn lại bao nhiêu? Phần biến mất là “do nguồn lực”; phần còn lại là “do hệ số” — tức là cùng nguồn lực nhưng lõi khai thác tốt hơn (thể chế, mạng lưới, quy mô thị trường…).",
  "Ví dụ so sánh lương nam – nữ kinh điển của Oaxaca (1973) và Blinder (1973) cũng đúng công thức này: phần do khác biệt học vấn/kinh nghiệm và phần “không giải thích được”."]));
ch.push(H("9.3. Tính từng bước: phiên bản 1 biến, tham chiếu β* = β_L (6 tỉnh, 2024)", 2));
ch.push(...STEP([
  `**Bước 1 – Trung bình từng nhóm:** Lõi (${ox.A.join(", ")}): x̄_L = ${v3(ox.xA)} DN/1.000 dân, ȳ_L = ${v4(ox.yA)}. Ngoại vi (${ox.B.join(", ")}): x̄_N = ${v3(ox.xB)}, ȳ_N = ${v4(ox.yB)}.`,
  `**Bước 2 – Khoảng cách:** ȳ_L − ȳ_N = ${v4(ox.yA)} − ${v4(ox.yB)} = **${v4(ox.gap)}** (≈ e^${v4(ox.gap)} − 1 = ${v1(ox.gap_pct_ebi)}% EBI).`,
  `**Bước 3 – Hồi quy riêng từng nhóm** (như Mục 8): lõi ln(EBI) = ${v4(ox.aA)} + ${v4(ox.bA)}·x; ngoại vi ln(EBI) = ${v4(ox.aB)} + ${v4(ox.bB)}·x. (Lõi chỉ có 2 tỉnh nên đường đi qua đúng 2 điểm — chỉ để minh hoạ.)`,
  `**Bước 4 – Phần nguồn lực** = (x̄_L − x̄_N)·β_L = (${v3(ox.xA)} − ${v3(ox.xB)}) × ${v4(ox.bA)} = **${v4(ox.endow)}** → chiếm ${v4(ox.endow)}/${v4(ox.gap)} = **${v1(ox.endow_pct)}%** khoảng cách.`,
  `**Bước 5 – Phần hệ số** = khoảng cách − phần nguồn lực = ${v4(ox.gap)} − ${v4(ox.endow)} = **${v4(ox.coef)}** → **${v1(ox.coef_pct)}%**. Kiểm tra: ${v4(ox.endow)} + ${v4(ox.coef)} = ${v4(ox.gap)} ✓.`,
  "**Bước 6 – Với 8 biến** (bài báo): phần nguồn lực là tổng của 8 số hạng (X̄_{L,j} − X̄_{N,j})·β*_j, mỗi số hạng chia cho khoảng cách là “% do biến j” — chính là Bảng 8b và Hình 8.",
]));
ch.push(...TRONG_BAI(["Bảng 8 (331 quan sát, 8 biến + biến giả năm, β* gộp): khoảng cách 0,455; nguồn lực 0,360 (79,1%), SE bootstrap 0,129; hệ số 0,095 (20,9%), SE 0,060 – không có ý nghĩa ở mức 10%. Kết quả ổn định khi đổi β* = β_L (82%/18%) và đổi mẫu (80%/20%).",
  "Bảng 8b / Hình 8: riêng mật độ DN góp 0,316 = 69,4% khoảng cách; logistics 5,1%; FDI 4,6%; hộ có máy tính 3,9%; GRDP, đô thị hoá, bán lẻ hiện đại góp âm nhỏ (nếu chỉ xét chúng, ngoại vi lẽ ra không kém đến vậy).",
  "**SE bootstrap**: lặp 300 lần, mỗi lần rút ngẫu nhiên có hoàn lại 63 tỉnh (giữ nguyên các năm của tỉnh được rút), tính lại hai phần; SE = độ lệch chuẩn của 300 kết quả. Đây là cách lấy sai số chuẩn khi không có công thức đóng."]));
ch.push(...LUU_Y(["Phần hệ số theo từng biến (cột cuối Bảng 8b) đổi giá trị khi đổi gốc đo của biến (ví dụ % → tỷ lệ), nên bài chỉ đọc **tổng** phần hệ số, không đọc từng dòng.",
  "Nếu đổi hệ số tham chiếu (β_L, β_N hay gộp) thì hai phần đổi chút ít; bài báo cáo cả hai để chứng minh kết luận không phụ thuộc lựa chọn này."]));

// ================= 10. THUẬT NGỮ =================
ch.push(H("10. Bảng thuật ngữ song ngữ", 1));
ch.push(TBL(["Trong bài (tiếng Việt)", "Thuật ngữ chuyên môn / tiếng Anh", "Hiểu đơn giản"], [
  ["Chỉ số TMĐT (EBI)", "E-Business Index (VECOM)", "Điểm 0–100 do VECOM chấm cho mỗi tỉnh, tổng có trọng số của 3–4 trụ cột."],
  ["Trụ cột NNL&HT / B2C / B2B / G2B", "Human resources & IT infrastructure / business-to-consumer / business-to-business / government-to-business", "Bốn mảng chấm điểm: nhân lực – hạ tầng; bán cho người tiêu dùng; bán giữa doanh nghiệp; giao dịch với cơ quan nhà nước (chỉ đến 2020)."],
  ["EBI tương đối", "Relative EBI (mean = 100)", "Điểm của tỉnh so với trung bình cả nước cùng năm."],
  ["Điểm gãy phương pháp 2021", "Methodological break", "VECOM đổi cách tính từ 2021 → không so thô hai giai đoạn."],
  ["Hệ số Gini", "Gini coefficient", "0 = bằng nhau hoàn toàn; càng gần 1 càng tập trung."],
  ["Chỉ số Theil-T / MLD", "Theil index (GE(1)) / mean log deviation (GE(0))", "Hai thước đo chênh lệch tách được thành phần giữa nhóm + trong nhóm; Theil nhạy với tỉnh cao, MLD nhạy với tỉnh thấp."],
  ["Phần giữa vùng / trong vùng", "Between-group / within-group component", "Chênh lệch do các vùng khác nhau / do các tỉnh trong cùng vùng khác nhau."],
  ["Phân rã Gini theo nguồn", "Gini decomposition by source (S·G·R)", "Mỗi trụ cột góp bao nhiêu % vào chênh lệch EBI."],
  ["Hệ số biến thiên (CV)", "Coefficient of variation = SD/mean", "Độ chênh lệch đã chia cho trung bình, so sánh được giữa các năm."],
  ["Hội tụ sigma", "σ-convergence", "CV giảm theo thời gian."],
  ["Hội tụ beta (tuyệt đối / có điều kiện)", "Absolute / conditional β-convergence", "Tỉnh thấp tăng nhanh hơn (không / có kiểm soát đặc trưng cấu trúc)."],
  ["Tốc độ hội tụ λ, thời gian bán hội tụ", "Speed of convergence, half-life", "% khoảng cách xoá mỗi năm; số năm xoá một nửa."],
  ["Câu lạc bộ hội tụ", "Convergence club (Phillips & Sul, 2007)", "Các nhóm tỉnh tiến về các mức đích khác nhau thay vì một mức chung."],
  ["Hồi quy OLS gộp", "Pooled OLS", "Đường/mặt phẳng khớp nhất qua mọi quan sát tỉnh–năm."],
  ["Hiệu ứng cố định năm / tỉnh", "Year / province fixed effects", "Một hằng số riêng cho mỗi năm / mỗi tỉnh để hút khác biệt chung."],
  ["Sai số chuẩn cụm theo tỉnh; HC1", "Cluster-robust SE; heteroskedasticity-consistent SE", "Cách tính độ không chắc chắn của hệ số cho phép sai số không đồng đều / tương quan trong tỉnh."],
  ["Giá trị p; dấu {*}{*}{*}, {*}{*}, {*}", "p-value; significance at 1%, 5%, 10%", "Xác suất thấy hệ số lớn như vậy nếu thật ra bằng 0; càng nhỏ càng chắc."],
  ["R², R² hiệu chỉnh", "R-squared, adjusted R-squared", "% biến động của y mà mô hình giải thích được (hiệu chỉnh: phạt số biến)."],
  ["VIF", "Variance inflation factor", "Cảnh báo biến giải thích trùng lặp nhau; < 10 là ổn."],
  ["Biến giả vùng; kiểm định Wald", "Region dummies; Wald test", "Biến 0/1 đánh dấu vùng; kiểm định “tất cả hệ số vùng có = 0 không?”."],
  ["Biến trễ một năm", "One-year lag (X_{t−1})", "Dùng giá trị năm trước của X để giảm khả năng EBI ‘gây ngược’ lên X."],
  ["Phân tách Oaxaca–Blinder", "Oaxaca–Blinder decomposition", "Tách khoảng cách hai nhóm thành phần nguồn lực + phần hệ số."],
  ["Phần nguồn lực / phần hệ số", "Endowment (explained) / coefficient (unexplained) component", "Do khác biệt X / do cùng X nhưng chuyển hoá khác."],
  ["Hệ số tham chiếu (Jann 2008)", "Reference coefficients from pooled model with group indicator", "Hệ số ‘trung lập’ dùng để định giá khác biệt X."],
  ["Bootstrap theo tỉnh", "Cluster bootstrap (300 replications)", "Rút mẫu lại nhiều lần để ước lượng sai số chuẩn."],
  ["Vùng lõi / ngoại vi", "Core / periphery", "ĐNB + ĐBSH (17 tỉnh) / 4 vùng còn lại (46 tỉnh)."],
], [2400, 3000, 3670]));
ch.push(GAP());
ch.push(P("**Cách tự kiểm tra:** mở HUONG-DAN-TINH-TUNG-BUOC.xlsx, thay số ở ô vàng (ví dụ đổi EBI của Thái Nguyên) và xem mọi kết quả cập nhật; rồi mở `ket-qua/panel_phan_tich.csv` chọn 6 tỉnh khác thay vào sheet 1_Du_lieu. Khi đã tin phép tính, chạy `python scripts/01_phan_tich.py` để tái tạo toàn bộ bảng của bài."));

const doc = new Document({
  creator: "Đào Thị Hồng Vân", title: "Hướng dẫn đọc hiểu cách tính từng bước",
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  features: { updateFields: true },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.BULLET, text: "▸", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] }) },
    children: ch,
  }],
});
Packer.toBuffer(doc).then((buf) => { const out = path.join(ROOT, "HUONG-DAN-TINH-TUNG-BUOC.docx"); fs.writeFileSync(out, buf); console.log("Đã ghi", out, (buf.length / 1024).toFixed(0), "KB"); });

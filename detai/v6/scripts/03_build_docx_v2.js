// Dựng file Word bài báo BẢN v2: công thức là phương trình Word gốc (OMML), đánh số tự động, có giải thích ký hiệu.
// Nguồn: noi_dung_v2.js + cong_thuc_v2.js + ghi_chu_bang_v2.js + ket-qua/bang_docx.json + hinh-bai-bao/hinh/*.png
// Chạy: node scripts/03_build_docx_v2.js   (gói docx có sẵn trong scripts/node_modules)
// Bản v1 (03_build_docx.js, noi_dung.js) giữ nguyên; script này không bao giờ ghi đè file bài báo v1.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, ImageRun, HeadingLevel, ShadingType, PageNumber, Footer, LevelFormat,
  ImportedXmlComponent, VerticalAlign,
} = require("docx");
const B = require("./cong_thuc_omml.js");
const { M } = require("./cong_thuc_v2.js");
const NOTES = require("./ghi_chu_bang_v2.js");

const ROOT = path.resolve(__dirname, "..");
const KQ = path.join(ROOT, "ket-qua");
const HINH = path.join(ROOT, "hinh-bai-bao", "hinh");   // hình do hinh-bai-bao/tao_hinh.py tạo ra
const T = JSON.parse(fs.readFileSync(path.join(KQ, "bang_docx.json"), "utf8"));
const ND = require("./noi_dung_v2.js");
const V1_FILE = "Bai-bao-EBI-63-tinh-2017-2024.docx";

const FONT = "Times New Roman";
const BODY = 24;      // 12 pt (half-points)
const SMALL = 20;     // 10 pt
const TABLE = 18;     // 9 pt
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1418; // A4, lề 2,5 cm
const CONTENT_W = PAGE_W - 2 * MARGIN;               // 9070 DXA

// ---------- công thức ----------
// Số phương trình đánh theo thứ tự xuất hiện của các khối eq (trừ khối n: false).
const EQNO = {};
{ let n = 0; for (const b of ND.blocks) if (b.t === "eq" && b.n !== false) EQNO[b.key] = ++n; }
const eqRef = (key) => { if (!EQNO[key]) throw new Error("Không có phương trình đánh số: " + key); return `(${EQNO[key]})`; };
const formula = (key) => { if (!M[key]) throw new Error("Thiếu công thức: " + key); return M[key]; };
const xml = (s) => ImportedXmlComponent.fromXmlString(s).root[0];

// Ký hiệu ngắn trong câu văn: $...$. Chữ cái đơn in nghiêng, tên nhiều chữ (EBI, Th, ln, CV) in đứng;
// _x / ^x (hoặc _{..} / ^{..}) chỉ số dưới/trên; \bar{..}, \hat{..}, \over{..} (gạch trên cả cụm), \rm{..} (in đứng).
function lin(src) {
  let i = 0; const out = [];
  const readGroup = () => {
    let depth = 1, j = i;
    while (j < src.length && depth) { if (src[j] === "{") depth++; else if (src[j] === "}") depth--; j++; }
    const inner = src.slice(i, j - 1); i = j; return lin(inner);
  };
  const readAtom = () => {
    if (src[i] === "{") { i++; return readGroup(); }
    for (const [cmd, fn] of [["\\bar{", B.bar], ["\\hat{", B.hat], ["\\over{", B.over]]) {
      if (src.startsWith(cmd, i)) { i += cmd.length; return fn(readGroup()); }
    }
    if (src.startsWith("\\rm{", i)) { i += 4; const j = src.indexOf("}", i); const w = src.slice(i, j); i = j + 1; return B.t(w); }
    const m = /^[A-Za-z]+/.exec(src.slice(i));
    if (m) {
      i += m[0].length;
      if (m[0] === "ln" && src[i] === " ") {   // "ln x": đối tượng hàm của Word để có khoảng cách chuẩn giữa ln và đối số
        while (src[i] === " ") i++;
        return B.ln(readTerm());
      }
      return m[0].length > 1 ? B.t(m[0]) : B.v(m[0]);
    }
    return B.v(src[i++]);   // chữ Hy Lạp, chữ số, toán tử: mỗi ký tự một run
  };
  const readTerm = () => {
    const base = readAtom(); let lo = null, hi = null;
    while (src[i] === "_" || src[i] === "^") { const k = src[i++]; const a = readAtom(); if (k === "_") lo = a; else hi = a; }
    return lo && hi ? B.subsup(base, lo, hi) : lo ? B.sub(base, lo) : hi ? B.sup(base, hi) : base;
  };
  while (i < src.length) {
    if (src[i] === " ") { i++; continue; }
    out.push(readTerm());
  }
  return out.join("");
}

// ---------- inline markup: {{m:khóa}}, {{eq:khóa}}, $ký hiệu$, **bold**, *italic*, ^{sup}, _{sub} ----------
// emph: false tắt **/* (ghi chú bảng có chuỗi "***, **, *").
function runs(text, base = {}, emph = true) {
  const out = [];
  const re = emph
    ? /(\{\{m:[\w-]+\}\}|\{\{eq:[\w-]+\}\}|\$[^$]+\$|\*\*[^*]+\*\*|\*[^*]+\*|\^\{[^}]+\}|_\{[^}]+\})/g
    : /(\{\{m:[\w-]+\}\}|\{\{eq:[\w-]+\}\}|\$[^$]+\$|\^\{[^}]+\}|_\{[^}]+\})/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const s = m[0];
    if (s.startsWith("{{m:")) out.push(xml(B.oMath(formula(s.slice(4, -2)))));
    else if (s.startsWith("{{eq:")) out.push(new TextRun({ text: eqRef(s.slice(5, -2)), ...base }));
    else if (s.startsWith("$")) out.push(xml(B.oMath(lin(s.slice(1, -1)))));
    else if (s.startsWith("**")) out.push(new TextRun({ text: s.slice(2, -2), bold: true, ...base }));
    else if (s.startsWith("*")) out.push(new TextRun({ text: s.slice(1, -1), italics: true, ...base }));
    else if (s.startsWith("^{")) out.push(new TextRun({ text: s.slice(2, -1), superScript: true, ...base }));
    else out.push(new TextRun({ text: s.slice(2, -1), subScript: true, ...base }));
    last = m.index + s.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}

const P = (text, opt = {}) => new Paragraph({
  children: runs(text, { font: FONT, size: opt.size || BODY, ...(opt.run || {}) }),
  alignment: opt.align || AlignmentType.JUSTIFIED,
  spacing: { after: opt.after ?? 120, before: opt.before ?? 0, line: opt.line ?? 276 },
  indent: opt.indent === false ? undefined : { firstLine: opt.firstLine ?? 567 },
  keepNext: opt.keepNext, ...(opt.extra || {}),
});

function heading(text, level) {
  const sizes = { 1: 26, 2: 24, 3: 24 };
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, italics: level === 3, font: FONT, size: sizes[level], color: "000000" })],
    spacing: { before: level === 1 ? 240 : 160, after: 120 },
    keepNext: true,
  });
}

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const NO_BORDERS = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// Phương trình hiển thị: bảng 3 cột không viền — công thức căn giữa ở cột giữa, số (n) căn phải ở cột phải.
function eqBlock(b) {
  const side = 850, mid = CONTENT_W - 2 * side;
  const cell = (children, w) => new TableCell({
    children, width: { size: w, type: WidthType.DXA }, borders: NO_BORDERS,
    verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 },
  });
  const para = (kids, align) => new Paragraph({ children: kids, alignment: align, spacing: { before: 80, after: 80, line: 240 } });
  const num = b.n === false ? [] : [new TextRun({ text: eqRef(b.key), font: FONT, size: BODY })];
  return new Table({
    rows: [new TableRow({ cantSplit: true, children: [
      cell([para([], AlignmentType.LEFT)], side),
      cell([para([xml(B.oMathPara(formula(b.key)))], AlignmentType.CENTER)], mid),
      cell([para(num, AlignmentType.RIGHT)], side),
    ] })],
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [side, mid, side], layout: "fixed",
    borders: { ...NO_BORDERS, insideHorizontal: noBorder, insideVertical: noBorder },
  });
}

function tableBlock(key) {
  const t = T[key];
  const total = t.widths.reduce((a, b) => a + b, 0);
  const widths = t.widths.map((w) => Math.round((w / total) * CONTENT_W));
  const cellPara = (txt, bold, alignRight) => txt.split("\n").map((line) => new Paragraph({
    children: [new TextRun({ text: line, font: FONT, size: TABLE, bold })],
    alignment: alignRight ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { after: 0, before: 0, line: 240 },
  }));
  const mkCell = (txt, i, bold, isHeader, isLast) => new TableCell({
    children: cellPara(txt, bold, i > 0),
    width: { size: widths[i], type: WidthType.DXA },
    borders: { top: isHeader ? thin : noBorder, bottom: isHeader || isLast ? thin : noBorder, left: noBorder, right: noBorder },
    margins: { top: 40, bottom: 40, left: 60, right: 60 },
    verticalAlign: "center",
    shading: isHeader ? { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" } : undefined,
  });
  const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: t.header.map((h, i) => mkCell(h, i, true, true, false)) })];
  t.rows.forEach((r, ri) => {
    const isLast = ri === t.rows.length - 1;
    rows.push(new TableRow({ cantSplit: true, children: r.map((c, i) => mkCell(String(c ?? ""), i, false, false, isLast)) }));
  });
  const out = [
    new Paragraph({ children: runs(t.title, { font: FONT, size: SMALL, bold: true }), alignment: AlignmentType.LEFT, spacing: { before: 200, after: 80 }, keepNext: true }),
    new Table({ rows, width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, layout: "fixed" }),
  ];
  const note = NOTES[key] ?? t.note;   // ghi chú bảng v2 dẫn chiếu số phương trình
  if (note) out.push(new Paragraph({ children: runs(note, { font: FONT, size: TABLE }, false), alignment: AlignmentType.JUSTIFIED, spacing: { before: 60, after: 200 } }));
  return out;
}

function figureBlock(file, caption, widthCm) {
  const buf = fs.readFileSync(path.join(HINH, file));
  // đọc kích thước PNG
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const W = Math.round((widthCm || 15) * 37.8), H = Math.round(W * h / w);
  return [
    new Paragraph({ children: [new ImageRun({ type: "png", data: buf, transformation: { width: W, height: H } })], alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, keepNext: true }),
    new Paragraph({ children: runs(caption, { font: FONT, size: SMALL, bold: false }), alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
  ];
}

function build() {
  const children = [];
  for (const b of ND.blocks) {
    switch (b.t) {
      case "title":
        children.push(new Paragraph({ children: [new TextRun({ text: b.text, bold: true, font: FONT, size: 28 })], alignment: AlignmentType.CENTER, spacing: { after: 200, line: 276 } }));
        break;
      case "center":
        children.push(new Paragraph({ children: runs(b.text, { font: FONT, size: b.size || BODY }), alignment: AlignmentType.CENTER, spacing: { after: b.after ?? 60 } }));
        break;
      case "h1": children.push(heading(b.text, 1)); break;
      case "h2": children.push(heading(b.text, 2)); break;
      case "h3": children.push(heading(b.text, 3)); break;
      case "p": children.push(P(b.text, b.opt || {})); break;
      case "abstract":
        children.push(new Paragraph({ children: runs(b.text, { font: FONT, size: SMALL }), alignment: AlignmentType.JUSTIFIED, spacing: { after: 100, line: 260 }, indent: { left: 567, right: 567 } }));
        break;
      case "eq": children.push(eqBlock(b)); break;
      case "where":   // "trong đó:" + danh sách giải thích ký hiệu
        children.push(new Paragraph({ children: runs(b.lead || "trong đó:", { font: FONT, size: BODY }), alignment: AlignmentType.LEFT, spacing: { after: 40, line: 276 }, keepNext: true }));
        b.items.forEach((it) => children.push(new Paragraph({
          children: runs(it, { font: FONT, size: BODY }), numbering: { reference: "bul", level: 0 },
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 40, line: 276 },
        })));
        children.push(new Paragraph({ children: [], spacing: { after: 40, line: 240 } }));
        break;
      case "bullets":
        b.items.forEach((it) => children.push(new Paragraph({
          children: runs(it, { font: FONT, size: BODY }), numbering: { reference: "bul", level: 0 },
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 },
        })));
        break;
      case "numbered":
        b.items.forEach((it) => children.push(new Paragraph({
          children: runs(it, { font: FONT, size: BODY }), numbering: { reference: "num" + b.id, level: 0 },
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 },
        })));
        break;
      case "table": children.push(...tableBlock(b.key)); break;
      case "fig": children.push(...figureBlock(b.file, b.caption, b.width)); break;
      case "refs":
        b.items.forEach((r) => children.push(new Paragraph({
          children: runs(r, { font: FONT, size: SMALL }), alignment: AlignmentType.JUSTIFIED,
          indent: { left: 567, hanging: 567 }, spacing: { after: 80, line: 260 },
        })));
        break;
      case "pagebreak": children.push(new Paragraph({ children: [new TextRun({ break: 1 })], pageBreakBefore: true })); break;
      default: throw new Error("Unknown block " + b.t);
    }
  }
  const numbering = {
    config: [
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ...["1", "2", "3", "4", "5", "6"].map((id) => ({ reference: "num" + id, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "(%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] })),
    ],
  };
  const doc = new Document({
    creator: ND.meta.author, title: ND.meta.title, description: ND.meta.title,
    styles: { default: { document: { run: { font: FONT, size: BODY } } } },
    numbering,
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SMALL })] })] }) },
      children,
    }],
  });
  return Packer.toBuffer(doc);
}

build().then((buf) => {
  const out = process.env.OUT_FILE || path.join(ROOT, ND.meta.file);   // OUT_FILE: ghi ra chỗ khác khi file đang mở trong Word
  if (path.basename(out) === V1_FILE) throw new Error("Không ghi đè bài báo v1: " + out);
  fs.writeFileSync(out, buf);
  console.log("Đã ghi", out, (buf.length / 1024).toFixed(0), "KB;", Object.keys(EQNO).length, "phương trình đánh số:", JSON.stringify(EQNO));
});

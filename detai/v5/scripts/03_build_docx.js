// Dựng file Word bài báo từ noi_dung.js + ket-qua/bang_docx.json + hinh-bai-bao/hinh/*.png
// Chạy: NODE_PATH=<thư mục có node_modules docx> node scripts/03_build_docx.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, ImageRun, HeadingLevel, ShadingType, PageNumber, Footer, TabStopType, LevelFormat,
} = require("docx");

const ROOT = path.resolve(__dirname, "..");
const KQ = path.join(ROOT, "ket-qua");
const HINH = path.join(ROOT, "hinh-bai-bao", "hinh");   // hình do hinh-bai-bao/tao_hinh.py tạo ra
const T = JSON.parse(fs.readFileSync(path.join(KQ, "bang_docx.json"), "utf8"));
const ND = require("./noi_dung.js");

const FONT = "Times New Roman";
const BODY = 24;      // 12 pt (half-points)
const SMALL = 20;     // 10 pt
const TABLE = 18;     // 9 pt
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1418; // A4, lề 2,5 cm
const CONTENT_W = PAGE_W - 2 * MARGIN;               // 9070 DXA

// ---------- inline markup: **bold**, *italic*, ^{sup}, _{sub} ----------
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|\^\{[^}]+\}|_\{[^}]+\})/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const s = m[0];
    if (s.startsWith("**")) out.push(new TextRun({ text: s.slice(2, -2), bold: true, ...base }));
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

function tableBlock(key) {
  const t = T[key];
  const total = t.widths.reduce((a, b) => a + b, 0);
  const widths = t.widths.map((w) => Math.round((w / total) * CONTENT_W));
  const ncol = widths.length;
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
  if (t.note) out.push(new Paragraph({ children: [new TextRun({ text: t.note, font: FONT, size: TABLE })], alignment: AlignmentType.JUSTIFIED, spacing: { before: 60, after: 200 } }));
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
      case "eq":
        children.push(new Paragraph({
          children: [...runs(b.text, { font: FONT, size: BODY }), new TextRun({ text: `\t(${b.n})`, font: FONT, size: BODY })],
          alignment: AlignmentType.CENTER, tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }], spacing: { before: 60, after: 120 },
        }));
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
  const out = process.env.OUT_FILE || path.join(ROOT, ND.meta.file);   // OUT_FILE: ghi ra chỗ khác khi file gốc đang mở trong Word
  fs.writeFileSync(out, buf);
  console.log("Đã ghi", out, (buf.length / 1024).toFixed(0), "KB");
});

// BƯỚC 4 – DỰNG FILE WORD BÀI BÁO BẢN v8 từ noi_dung.js (lời văn), ket-qua/bang_docx.json (bảng),
// ket-qua/so_lieu_bai.json (các con số trong lời văn) và hinh/*.png (hình).
// Chạy (trong thư mục v8):   node scripts/04_dung_word.js
//   Cần Node.js và gói "docx" (cài một lần: cd scripts && npm install).
//   Nếu file Word đang mở, đặt biến môi trường OUT_FILE để ghi ra tên khác.
// Khối nội dung: title, center, h1, h2, h3, p, abstract, eq (key, tag?: "A", n?: false), where (items), bullets (items),
//       numbered (id, items), table (key -> bang_docx.json), table_inline ({title, header, rows, widths, note}),
//       fig (file, caption, width cm), refs (items), pagebreak.
// Markup nội dòng: **đậm**, *nghiêng*, ^{trên}, _{dưới}, $ký hiệu$, {{m:khóa}} (công thức), {{eq:khóa}} (số phương trình),
//       {{n:khóa}} (con số lấy từ so_lieu_bai.json – bảo đảm lời văn khớp bảng kết quả).
// Công thức là phương trình Word gốc (OMML) từ cong_thuc.js; số phương trình tự đánh (1)…, phụ lục (A1)….
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, ImageRun, HeadingLevel, ShadingType, PageNumber, Footer, LevelFormat,
  ImportedXmlComponent, VerticalAlign, PageOrientation,
} = require("docx");
const B = require("./cong_thuc_omml.js");
const { M } = require("./cong_thuc.js");

const ROOT = path.resolve(__dirname, "..");
const modName = process.argv[2] || "noi_dung.js";
const ND = require(path.join(__dirname, modName));
const T = ND.meta.tables ? JSON.parse(fs.readFileSync(path.join(ROOT, ND.meta.tables), "utf8")) : {};
const HINH = path.join(ROOT, ND.meta.figdir || "hinh");
const SO = JSON.parse(fs.readFileSync(path.join(ROOT, "ket-qua", "so_lieu_bai.json"), "utf8"));
const PROTECTED = ["phantichketqua.docx"];
const so = (s) => String(s).replace(/\{\{n:([^}]+)\}\}/g, (_, k) => { if (!(k in SO)) throw new Error("Thiếu con số: " + k); return SO[k]; });

const FONT = "Times New Roman";
const BODY = 24, SMALL = 20, TABLE = 18;
const landscape = !!ND.meta.landscape;
const PAGE_W = landscape ? 16838 : 11906, PAGE_H = landscape ? 11906 : 16838, MARGIN = 1418;
const CONTENT_W = PAGE_W - 2 * MARGIN;

// ---------- số phương trình ----------
const EQNO = {};
{ let n = 0, a = 0;
  for (const b of ND.blocks) if (b.t === "eq" && b.n !== false) EQNO[b.key] = b.tag === "A" ? `A${++a}` : `${++n}`; }
const eqRef = (key) => { if (!EQNO[key]) throw new Error("Không có phương trình đánh số: " + key); return `(${EQNO[key]})`; };
const formula = (key) => { if (!M[key]) throw new Error("Thiếu công thức: " + key); return M[key]; };
const xml = (s) => ImportedXmlComponent.fromXmlString(s).root[0];

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
      if (m[0] === "ln" && src[i] === " ") { while (src[i] === " ") i++; return B.ln(readTerm()); }
      return m[0].length > 1 ? B.t(m[0]) : B.v(m[0]);
    }
    return B.v(src[i++]);
  };
  const readTerm = () => {
    const base = readAtom(); let lo = null, hi = null;
    while (src[i] === "_" || src[i] === "^") { const k = src[i++]; const a = readAtom(); if (k === "_") lo = a; else hi = a; }
    return lo && hi ? B.subsup(base, lo, hi) : lo ? B.sub(base, lo) : hi ? B.sup(base, hi) : base;
  };
  while (i < src.length) { if (src[i] === " ") { i++; continue; } out.push(readTerm()); }
  return out.join("");
}

function runs(text, base = {}, emph = true) {
  text = so(text);
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
    spacing: { before: level === 1 ? 240 : 160, after: 120 }, keepNext: true,
  });
}
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const NO_BORDERS = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function eqBlock(b) {
  const side = 850, mid = CONTENT_W - 2 * side;
  const cell = (children, w) => new TableCell({ children, width: { size: w, type: WidthType.DXA }, borders: NO_BORDERS,
    verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 } });
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

function tableFromSpec(t, opt = {}) {
  const total = t.widths.reduce((a, b) => a + b, 0);
  const widths = t.widths.map((w) => Math.round((w / total) * CONTENT_W));
  const size = opt.size || TABLE;
  const cellPara = (txt, bold, center) => String(txt).split("\n").map((line) => new Paragraph({
    children: runs(line, { font: FONT, size, bold }, opt.emph !== false ? true : false),
    alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { after: 0, before: 0, line: 240 },
  }));
  const mkCell = (txt, i, bold, isHeader, isLast) => new TableCell({
    children: cellPara(txt, bold, opt.leftAll ? false : i > 0),
    width: { size: widths[i], type: WidthType.DXA },
    borders: { top: isHeader ? thin : noBorder, bottom: isHeader || isLast ? thin : noBorder, left: noBorder, right: noBorder },
    margins: { top: 40, bottom: 40, left: 60, right: 60 }, verticalAlign: "center",
    shading: isHeader ? { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" } : undefined,
  });
  const rows = [new TableRow({ tableHeader: true, cantSplit: true, children: t.header.map((h, i) => mkCell(h, i, true, true, false)) })];
  t.rows.forEach((r, ri) => {
    const isLast = ri === t.rows.length - 1;
    rows.push(new TableRow({ cantSplit: !opt.allowSplit, children: r.map((c, i) => mkCell(c ?? "", i, false, false, isLast)) }));
  });
  const out = [];
  if (t.title) out.push(new Paragraph({ children: runs(t.title, { font: FONT, size: SMALL, bold: true }), alignment: AlignmentType.LEFT, spacing: { before: 200, after: 80 }, keepNext: true }));
  out.push(new Table({ rows, width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, layout: "fixed" }));
  if (t.note) out.push(new Paragraph({ children: runs(t.note, { font: FONT, size: TABLE }, false), alignment: AlignmentType.JUSTIFIED, spacing: { before: 60, after: 200 } }));
  return out;
}
function tableBlock(key) { if (!T[key]) throw new Error("Thiếu bảng: " + key); return tableFromSpec(T[key]); }

function figureBlock(file, caption, widthCm) {
  const buf = fs.readFileSync(path.join(HINH, file));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const W = Math.round((widthCm || 15) * 37.8), H = Math.round(W * h / w);
  return [
    new Paragraph({ children: [new ImageRun({ type: "png", data: buf, transformation: { width: W, height: H } })], alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, keepNext: true }),
    new Paragraph({ children: runs(caption, { font: FONT, size: SMALL }), alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
  ];
}

function build() {
  const children = [];
  for (const b of ND.blocks) {
    switch (b.t) {
      case "title":
        children.push(new Paragraph({ children: [new TextRun({ text: b.text, bold: true, font: FONT, size: b.size || 28 })], alignment: AlignmentType.CENTER, spacing: { after: 200, line: 276 } })); break;
      case "center":
        children.push(new Paragraph({ children: runs(b.text, { font: FONT, size: b.size || BODY }), alignment: AlignmentType.CENTER, spacing: { after: b.after ?? 60, before: b.before ?? 0 } })); break;
      case "h1": children.push(heading(b.text, 1)); break;
      case "h2": children.push(heading(b.text, 2)); break;
      case "h3": children.push(heading(b.text, 3)); break;
      case "p": children.push(P(b.text, b.opt || {})); break;
      case "abstract":
        children.push(new Paragraph({ children: runs(b.text, { font: FONT, size: SMALL }), alignment: AlignmentType.JUSTIFIED, spacing: { after: 100, line: 260 }, indent: { left: 567, right: 567 } })); break;
      case "eq": children.push(eqBlock(b)); break;
      case "where":
        children.push(new Paragraph({ children: runs(b.lead || "trong đó:", { font: FONT, size: BODY }), alignment: AlignmentType.LEFT, spacing: { after: 40, line: 276 }, keepNext: true }));
        b.items.forEach((it) => children.push(new Paragraph({ children: runs(it, { font: FONT, size: BODY }), numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 40, line: 276 } })));
        children.push(new Paragraph({ children: [], spacing: { after: 40, line: 240 } })); break;
      case "bullets":
        b.items.forEach((it) => children.push(new Paragraph({ children: runs(it, { font: FONT, size: b.size || BODY }), numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 } }))); break;
      case "numbered":
        b.items.forEach((it) => children.push(new Paragraph({ children: runs(it, { font: FONT, size: BODY }), numbering: { reference: "num" + b.id, level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 } }))); break;
      case "table": children.push(...tableBlock(b.key)); break;
      case "table_inline": children.push(...tableFromSpec(b, b.opt || {})); break;
      case "fig": children.push(...figureBlock(b.file, b.caption, b.width)); break;
      case "refs":
        b.items.forEach((r) => children.push(new Paragraph({ children: runs(r, { font: FONT, size: SMALL }), alignment: AlignmentType.JUSTIFIED, indent: { left: 567, hanging: 567 }, spacing: { after: 80, line: 260 } }))); break;
      case "pagebreak": children.push(new Paragraph({ children: [new TextRun({ break: 1 })], pageBreakBefore: true })); break;
      default: throw new Error("Unknown block " + b.t);
    }
  }
  const numbering = { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ...["1", "2", "3", "4", "5", "6", "7", "8", "9"].map((id) => ({ reference: "num" + id, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "(%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] })),
  ] };
  const doc = new Document({
    creator: ND.meta.author, title: ND.meta.title, description: ND.meta.title,
    styles: { default: { document: { run: { font: FONT, size: BODY } } } }, numbering,
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H, orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SMALL })] })] }) },
      children,
    }],
  });
  return Packer.toBuffer(doc);
}

build().then((buf) => {
  const out = process.env.OUT_FILE || path.join(ROOT, ND.meta.file);
  if (PROTECTED.includes(path.basename(out))) throw new Error("Không ghi đè bài báo bản trước: " + out);
  fs.writeFileSync(out, buf);
  console.log("Đã ghi", out, (buf.length / 1024).toFixed(0), "KB;", Object.keys(EQNO).length, "phương trình:", JSON.stringify(EQNO));
});

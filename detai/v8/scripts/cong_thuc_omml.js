// Bộ dựng công thức Word (OMML) dạng chuỗi XML — dùng cho cong_thuc_v2.js và 03_build_docx_v2.js.
// Word mở ra là phương trình gốc, sửa được bằng Equation Editor.
// Quy ước: v() ký hiệu toán (chữ cái đơn tự in nghiêng), t() tên nhiều chữ in đứng (EBI, Th, CV, ln, cov),
// txt() chữ thường tiếng Việt đặt trong công thức.
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const WR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';

const run = (s, mrpr = "") => `<m:r>${mrpr}${WR}<m:t xml:space="preserve">${esc(s)}</m:t></m:r>`;
const v = (s) => run(s);
const t = (s) => run(s, '<m:rPr><m:sty m:val="p"/></m:rPr>');
const txt = (s) => `<m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><m:t xml:space="preserve">${esc(s)}</m:t></m:r>`;
const cat = (...xs) => xs.filter((x) => x !== null && x !== undefined).join("");

const sub = (b, s) => `<m:sSub><m:e>${b}</m:e><m:sub>${s}</m:sub></m:sSub>`;
const sup = (b, s) => `<m:sSup><m:e>${b}</m:e><m:sup>${s}</m:sup></m:sSup>`;
const subsup = (b, s, p) => `<m:sSubSup><m:e>${b}</m:e><m:sub>${s}</m:sub><m:sup>${p}</m:sup></m:sSubSup>`;
const frac = (n, d) => `<m:f><m:num>${n}</m:num><m:den>${d}</m:den></m:f>`;
const fracL = (n, d) => `<m:f><m:fPr><m:type m:val="lin"/></m:fPr><m:num>${n}</m:num><m:den>${d}</m:den></m:f>`; // a/b trên một dòng
const nary = (chr, lo, hi, body) =>
  `<m:nary><m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="undOvr"/>` +
  `${lo ? "" : '<m:subHide m:val="1"/>'}${hi ? "" : '<m:supHide m:val="1"/>'}</m:naryPr>` +
  `<m:sub>${lo || ""}</m:sub><m:sup>${hi || ""}</m:sup><m:e>${body}</m:e></m:nary>`;
const sum = (lo, hi, body) => nary("∑", lo, hi, body);
const d = (body, beg = "(", end = ")") =>
  `<m:d><m:dPr><m:begChr m:val="${beg}"/><m:endChr m:val="${end}"/></m:dPr><m:e>${body}</m:e></m:d>`;
const acc = (body, chr) => `<m:acc><m:accPr><m:chr m:val="${chr}"/></m:accPr><m:e>${body}</m:e></m:acc>`;
const bar = (body) => acc(body, "̅");                                             // gạch ngang trên: ȳ, X̄
const hat = (body) => acc(body, "̂");                                             // mũ: β̂
const over = (body) => `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${body}</m:e></m:bar>`; // gạch trên cả cụm: EBI
const func = (name, arg) => `<m:func><m:fName>${t(name)}</m:fName><m:e>${arg}</m:e></m:func>`;
const ln = (arg) => func("ln", arg);
const sqrt = (b) => `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${b}</m:e></m:rad>`;

const oMath = (x) => `<m:oMath>${x}</m:oMath>`;
const oMathPara = (x) => `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${oMath(x)}</m:oMathPara>`;

module.exports = { esc, v, t, txt, cat, sub, sup, subsup, frac, fracL, nary, sum, d, acc, bar, hat, over, func, ln, sqrt, oMath, oMathPara };

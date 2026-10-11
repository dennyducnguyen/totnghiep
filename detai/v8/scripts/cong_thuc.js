// Từ điển công thức của bài báo bản v8 — mỗi khóa là thân một công thức Word gốc (OMML), dựng bằng cong_thuc_omml.js.
// Khối hiển thị: { t: "eq", key } trong noi_dung.js; công thức nội dòng: {{m:khóa}}; số phương trình: {{eq:khóa}}.
// Ký hiệu: i = tỉnh, t = năm, k = vùng; y = EBI (hoặc điểm trụ cột); ȳ = trung bình; τ_t = hiệu ứng năm.
const { v, t, cat, sub, sup, subsup, frac, fracL, sum, d, bar, func, ln, sqrt } = require("./cong_thuc_omml.js");

const EQ = v("="), PLUS = v("+"), MINUS = v("−"), COMMA = t(",  "), PR = v("′");
const EBI = t("EBI"), DTF = t("DTF");
const y = v("y"), ybar = bar(v("y"));
const yi = sub(y, v("i")), yj = sub(y, v("j")), ybk = sub(ybar, v("k"));
const Nk = sub(v("N"), v("k"));
const sk = sub(v("s"), v("k"));
const Th = t("Th"), Thk = sub(t("Th"), v("k")), ThW = sub(t("Th"), t("W")), ThB = sub(t("Th"), t("B"));
const sumK = (body) => sum(v("k=1"), v("K"), body);
const sumN = (body) => sum(v("i=1"), v("N"), body);
const inv = (den) => frac(v("1"), den);
const lnDTF = (idx) => ln(sub(DTF, v(idx)));

const M = {
  // ===== 3.3 biến đo vị thế =====
  dtf: cat(sub(DTF, v("i,t")), EQ, v("100"), v("×"), frac(sub(EBI, v("i,t")), cat(sub(t("max"), v("j")), sub(EBI, v("j,t"))))),

  // ===== 3.4.1 mức độ khoảng cách =====
  ybar_def: cat(ybar, EQ, d(fracL(v("1"), v("N"))), sumN(yi)),
  cv: cat(t("CV"), EQ, frac(v("s"), ybar), COMMA, v("s"), EQ,
    sqrt(cat(inv(cat(v("N"), MINUS, v("1"))), sumN(sup(d(cat(yi, MINUS, ybar)), v("2")))))),
  gini: cat(v("G"), EQ, inv(cat(v("2"), sup(v("N"), v("2")), ybar)), sumN(sum(v("j=1"), v("N"), d(cat(yi, MINUS, yj), "|", "|")))),
  theil: cat(Th, EQ, inv(v("N")), sumN(cat(frac(yi, ybar), ln(d(frac(yi, ybar)))))),

  // ===== 3.4.2 phân rã Theil =====
  theil_dec: cat(Th, EQ, sumK(cat(sk, Thk)), PLUS, sumK(cat(sk, ln(d(frac(ybk, ybar))))), COMMA, sk, EQ, frac(Nk, v("N")), v("·"), frac(ybk, ybar)),
  theil_wb: cat(Th, EQ, ThW, PLUS, ThB),
  theil_k: cat(Thk, EQ, d(fracL(v("1"), Nk)), sum(v("i∈k"), null, cat(d(fracL(yi, ybk)), ln(d(fracL(yi, ybk)))))),
  ThW_share: cat(v("100"), v("·"), fracL(ThW, Th)),

  // ===== 3.4.3 hội tụ =====
  beta: cat(lnDTF("i,t"), MINUS, lnDTF("i,t−1"), EQ, v("α"), PLUS, v("β"), lnDTF("i,t−1"), PLUS, v("γ"), PR, sub(v("Z"), v("i,t−1")),
    PLUS, sub(v("τ"), v("t")), PLUS, sub(v("ε"), v("i,t"))),
  speed: cat(v("λ"), EQ, MINUS, ln(d(cat(v("1"), PLUS, v("β"))))),

  // ===== 3.4.4 các yếu tố liên quan =====
  pooled: cat(subsup(DTF, v("i,t"), d(v("p"))), EQ, v("α"), PLUS, v("β"), PR, sub(v("X"), v("i,t")), PLUS, sub(v("τ"), v("t")), PLUS, sub(v("ε"), v("i,t"))),
  cross: cat(sub(DTF, v("i,2024")), EQ, v("α"), PLUS, v("δ"), v("·"), sub(t("KN"), v("i")), PLUS, v("γ"), PR, sub(v("W"), v("i")), PLUS, sub(v("u"), v("i"))),
};

module.exports = { M };

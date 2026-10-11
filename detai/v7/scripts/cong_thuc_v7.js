// Từ điển công thức của bài báo bản v7 — mỗi khóa là thân một công thức OMML (dựng bằng cong_thuc_omml.js).
// Khối hiển thị: { t: "eq", key } trong noi_dung_v7.js; công thức nội dòng: {{m:key}}; số phương trình: {{eq:key}}.
// Ký hiệu: Th = chỉ số Theil-T; i = tỉnh, t = năm, k = vùng (hoặc trụ cột, nói rõ tại chỗ); ȳ = trung bình.
const { v, t, cat, sub, sup, subsup, frac, fracL, sum, d, bar, hat, over, func, ln, sqrt } = require("./cong_thuc_omml.js");

const EQ = v("="), PLUS = v("+"), MINUS = v("−"), COMMA = t(",  "), PR = v("′");
const EBI = t("EBI");
const EBIit = sub(EBI, v("i,t"));
const EBIrel = (idx) => subsup(EBI, v(idx), t("rel"));
const y = v("y"), ybar = bar(v("y"));
const yi = sub(y, v("i")), yj = sub(y, v("j")), ybk = sub(ybar, v("k"));
const Nk = sub(v("N"), v("k")), Nt = sub(v("N"), v("t"));
const sk = sub(v("s"), v("k")), pk = sub(v("p"), v("k"));
const Th = t("Th"), Thk = sub(t("Th"), v("k")), ThW = sub(t("Th"), t("W")), ThB = sub(t("Th"), t("B"));
const L = v("L"), Lk = sub(v("L"), v("k")), LW = sub(v("L"), t("W")), LB = sub(v("L"), t("B"));
const Sk = sub(v("S"), v("k")), Gk = sub(v("G"), v("k")), Rk = sub(v("R"), v("k"));
const Xk = sub(v("X"), v("k")), wk = sub(v("w"), v("k"));
const sumK = (body) => sum(v("k=1"), v("K"), body);
const sumN = (body) => sum(v("i=1"), v("N"), body);
const inv = (den) => frac(v("1"), den);

const M = {
  // ===== 4.2 chỉ số EBI =====
  ebi_w: cat(EBIit, EQ, sum(v("k"), null, cat(sub(v("w"), v("k,t")), sub(v("X"), v("k,i,t")))),
    COMMA, sum(v("k"), null, sub(v("w"), v("k,t"))), EQ, v("1")),
  ebi_rel: cat(EBIrel("i,t"), EQ, v("100"), v("×"), frac(EBIit, cat(inv(Nt), sum(v("j=1"), Nt, sub(EBI, v("j,t")))))),

  // ===== 4.4 RQ1 =====
  ybar_def: cat(ybar, EQ, d(fracL(v("1"), v("N"))), sumN(yi)),
  gini: cat(v("G"), EQ, inv(cat(v("2"), sup(v("N"), v("2")), ybar)), sumN(sum(v("j=1"), v("N"), d(cat(yi, MINUS, yj), "|", "|")))),
  theil: cat(Th, EQ, inv(v("N")), sumN(cat(frac(yi, ybar), ln(d(frac(yi, ybar)))))),
  sdln: cat(sub(v("σ"), v("t")), EQ, sqrt(cat(inv(cat(Nt, MINUS, v("1"))), sum(v("i=1"), Nt,
    sup(d(cat(ln(sub(y, v("i,t"))), MINUS, sub(over(cat(t("ln"), v("y"))), v("t")))), v("2")))))),

  // ===== 4.4 RQ2 =====
  theil_dec: cat(Th, EQ, sumK(cat(sk, Thk)), PLUS, sumK(cat(sk, ln(d(frac(ybk, ybar))))), COMMA, sk, EQ, frac(Nk, v("N")), v("·"), frac(ybk, ybar)),
  theil_wb: cat(Th, EQ, ThW, PLUS, ThB),
  theil_k: cat(Thk, EQ, d(fracL(v("1"), Nk)), sum(v("i∈k"), null, cat(d(fracL(yi, ybk)), ln(d(fracL(yi, ybk)))))),
  sk_share: cat(sk, EQ, fracL(cat(sum(v("i∈k"), null, yi)), cat(sum(v("i"), null, yi)))),
  ThW_def: cat(ThW, EQ, sumK(cat(sk, Thk))),
  ThB_def: cat(ThB, EQ, sumK(cat(sk, ln(d(fracL(ybk, ybar)))))),
  ThB_share: cat(v("100"), v("·"), fracL(ThB, Th)),
  contrib_k: cat(v("100"), v("·"), fracL(cat(sk, Thk), Th)),

  // ===== 4.4 RQ3 =====
  fe: cat(ln(EBIit), EQ, sub(v("α"), v("i")), PLUS, sub(v("λ"), v("t")), PLUS, v("β"), PR, sub(v("X"), v("i,t")), PLUS, sub(v("ε"), v("i,t"))),
  scale_abs: cat(ln(d(cat(sub(v("c"), v("t")), v("·"), EBIit))), EQ, ln(sub(v("c"), v("t"))), PLUS, ln(EBIit)),
  semi_el: cat(v("100"), v("·"), d(cat(sup(v("e"), sub(v("β"), v("j"))), MINUS, v("1")))),
  fe_rank: cat(sub(v("R"), v("i,t")), EQ, sub(v("α"), v("i")), PLUS, sub(v("λ"), v("t")), PLUS, v("β"), PR, sub(v("X"), v("i,t")), PLUS, sub(v("ε"), v("i,t"))),

  // ===== Phụ lục =====
  mld: cat(L, EQ, inv(v("N")), sumN(ln(d(frac(ybar, yi))))),
  mld_dec: cat(L, EQ, sumK(cat(pk, Lk)), PLUS, sumK(cat(pk, ln(d(frac(ybar, ybk))))), COMMA, pk, EQ, frac(Nk, v("N"))),
  LB_def: cat(LB, EQ, sumK(cat(pk, ln(d(fracL(ybar, ybk)))))),
  gini_src: cat(v("G"), EQ, sum(v("k"), null, cat(Sk, Gk, Rk)), COMMA, Sk, EQ, frac(cat(wk, sub(bar(v("X")), v("k"))), over(EBI)),
    COMMA, Rk, EQ, frac(func("cov", d(cat(Xk, t(", "), v("F"), d(EBI)))), func("cov", d(cat(Xk, t(", "), v("F"), d(Xk)))))),
  beta: cat(sub(v("g"), v("i")), EQ, inv(v("T")), ln(d(frac(sub(EBI, v("i,T")), sub(EBI, v("i,0"))))),
    EQ, v("α"), PLUS, v("β"), ln(sub(EBI, v("i,0"))), PLUS, sub(v("ε"), v("i"))),
  lambda: cat(v("λ"), EQ, MINUS, frac(ln(d(cat(v("1"), PLUS, v("βT")))), v("T")), COMMA, sub(v("t"), v("1/2")), EQ, frac(func("ln", v("2")), v("λ"))),
};

module.exports = { M };

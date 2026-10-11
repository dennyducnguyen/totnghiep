// Từ điển công thức của bài báo (bản v2) — mỗi khóa là thân một công thức OMML.
// Khối hiển thị: { t: "eq", key } trong noi_dung_v2.js; công thức nội dòng: {{m:key}}; số phương trình: {{eq:key}}.
// Ký hiệu đã thống nhất: Th = chỉ số Theil-T (tránh trùng với T = độ dài giai đoạn trong phương trình hội tụ).
const {
  v, t, cat, sub, sup, subsup, frac, fracL, sum, d, bar, hat, over, func, ln, sqrt,
} = require("./cong_thuc_omml.js");

// ---------- ký hiệu dùng lại ----------
const EQ = v("="), PLUS = v("+"), MINUS = v("−"), COMMA = t(",  "), PR = v("′");
const EBI = t("EBI");
const EBIit = sub(EBI, v("i,t"));
const EBIrel = (idx) => subsup(EBI, v(idx), t("rel"));
const lnEBIit = ln(EBIit);
const y = v("y"), ybar = bar(v("y"));
const yi = sub(y, v("i")), yj = sub(y, v("j"));
const ybk = sub(ybar, v("k"));
const Nk = sub(v("N"), v("k")), Nt = sub(v("N"), v("t"));
const sk = sub(v("s"), v("k")), pk = sub(v("p"), v("k"));
const Th = t("Th"), Thk = sub(t("Th"), v("k")), ThW = sub(t("Th"), t("W")), ThB = sub(t("Th"), t("B"));
const L = v("L"), Lk = sub(v("L"), v("k")), LW = sub(v("L"), t("W")), LB = sub(v("L"), t("B"));
const Sk = sub(v("S"), v("k")), Gk = sub(v("G"), v("k")), Rk = sub(v("R"), v("k"));
const Xk = sub(v("X"), v("k")), wk = sub(v("w"), v("k"));
const bstar = sup(v("β"), v("*"));
const Yb = (g) => sub(bar(v("Y")), v(g));
const Xb = (g) => sub(bar(v("X")), v(g));
const bh = (g) => sub(hat(v("β")), v(g));
const sumK = (body) => sum(v("k=1"), v("K"), body);
const sumN = (body) => sum(v("i=1"), v("N"), body);
const inv = (den) => frac(v("1"), den);

const M = {
  // ===== 3.1 =====
  ebi_w: cat(EBIit, EQ, sum(v("k"), null, cat(sub(v("w"), v("k,t")), sub(v("X"), v("k,i,t")))),
    COMMA, sum(v("k"), null, sub(v("w"), v("k,t"))), EQ, v("1")),
  ebi_rel: cat(EBIrel("i,t"), EQ, v("100"), v("×"),
    frac(EBIit, cat(inv(Nt), sum(v("j=1"), Nt, sub(EBI, v("j,t")))))),

  // ===== 3.3.1 RQ1 =====
  ybar_def: cat(ybar, EQ, d(fracL(v("1"), v("N"))), sumN(yi)),
  gini: cat(v("G"), EQ, inv(cat(v("2"), sup(v("N"), v("2")), ybar)),
    sumN(sum(v("j=1"), v("N"), d(cat(yi, MINUS, yj), "|", "|")))),
  theil: cat(Th, EQ, inv(v("N")), sumN(cat(frac(yi, ybar), ln(d(frac(yi, ybar)))))),
  mld: cat(L, EQ, inv(v("N")), sumN(ln(d(frac(ybar, yi))))),

  theil_dec: cat(Th, EQ, sumK(cat(sk, Thk)), PLUS, sumK(cat(sk, ln(d(frac(ybk, ybar))))),
    COMMA, sk, EQ, frac(Nk, v("N")), v("·"), frac(ybk, ybar)),
  theil_dec_exp: cat(Th, EQ,
    sumK(cat(frac(Nk, v("N")), frac(ybk, ybar), Thk)), PLUS,
    sumK(cat(frac(Nk, v("N")), frac(ybk, ybar), ln(d(frac(ybk, ybar)))))),
  theil_wb: cat(Th, EQ, ThW, PLUS, ThB),
  theil_k: cat(Thk, EQ, d(fracL(v("1"), Nk)), sum(v("i∈k"), null, cat(d(fracL(yi, ybk)), ln(d(fracL(yi, ybk)))))),
  sk_share: cat(sk, EQ, fracL(cat(sum(v("i∈k"), null, yi)), cat(sum(v("i"), null, yi)))),
  sum_sk: cat(sum(v("k"), null, sk), EQ, v("1")),
  ThW_def: cat(ThW, EQ, sumK(cat(sk, Thk))),
  ThB_def: cat(ThB, EQ, sumK(cat(sk, ln(d(fracL(ybk, ybar)))))),
  ThB_share: cat(v("100"), v("·"), fracL(ThB, Th)),

  mld_dec: cat(L, EQ, sumK(cat(pk, Lk)), PLUS, sumK(cat(pk, ln(d(frac(ybar, ybk))))),
    COMMA, pk, EQ, frac(Nk, v("N"))),
  LW_def: cat(LW, EQ, sumK(cat(pk, Lk))),
  LB_def: cat(LB, EQ, sumK(cat(pk, ln(d(fracL(ybar, ybk)))))),

  gini_src: cat(v("G"), EQ, sum(v("k"), null, cat(Sk, Gk, Rk))),
  gini_src_parts: cat(Sk, EQ, frac(cat(wk, sub(bar(v("X")), v("k"))), over(EBI)),
    COMMA, Rk, EQ, frac(func("cov", d(cat(Xk, t(", "), v("F"), d(EBI)))), func("cov", d(cat(Xk, t(", "), v("F"), d(Xk)))))),
  gini_contrib: cat(v("100"), v("·"), fracL(cat(Sk, Gk, Rk), v("G"))),

  // ===== 3.3.2 RQ2 =====
  cv: cat(sub(t("CV"), v("t")), EQ, frac(sub(v("σ"), v("t")), sub(ybar, v("t"))),
    COMMA, sub(v("σ"), v("t")), EQ,
    sqrt(cat(inv(cat(Nt, MINUS, v("1"))), sum(v("i=1"), Nt, sup(d(cat(sub(y, v("i,t")), MINUS, sub(ybar, v("t")))), v("2")))))),
  sdln: cat(subsup(t("SD"), v("t"), t("ln")), EQ,
    sqrt(cat(inv(cat(Nt, MINUS, v("1"))), sum(v("i=1"), Nt,
      sup(d(cat(ln(sub(y, v("i,t"))), MINUS, sub(over(cat(t("ln"), v("y"))), v("t")))), v("2")))))),

  beta: cat(sub(v("g"), v("i")), EQ, inv(v("T")), ln(d(frac(sub(EBI, v("i,T")), sub(EBI, v("i,0"))))),
    EQ, v("α"), PLUS, v("β"), ln(sub(EBI, v("i,0"))), PLUS, v("γ"), PR, sub(v("Z"), v("i")), PLUS, sub(v("ε"), v("i"))),
  beta_solow: cat(v("β"), EQ, MINUS, fracL(d(cat(v("1"), MINUS, sup(v("e"), cat(MINUS, v("λT"))))), v("T"))),
  lambda: cat(v("λ"), EQ, MINUS, frac(ln(d(cat(v("1"), PLUS, v("βT")))), v("T")),
    COMMA, sub(v("t"), v("1/2")), EQ, frac(func("ln", v("2")), v("λ"))),
  lambda_cond: cat(v("1"), PLUS, v("βT"), v(">"), v("0")),
  beta_panel: cat(sub(v("g"), v("i,t")), EQ, ln(EBIrel("i,t")), MINUS, ln(EBIrel("i,t−1")),
    EQ, sub(v("α"), v("t")), PLUS, v("β"), ln(EBIrel("i,t−1")), PLUS, sub(v("ε"), v("i,t"))),
  lambda1: cat(v("λ"), EQ, MINUS, ln(d(cat(v("1"), PLUS, v("β"))))),

  // ===== 3.3.3 RQ3 =====
  ols: cat(lnEBIit, EQ, sub(v("β"), v("0")), PLUS, v("β"), PR, sub(v("X"), v("i,t")), PLUS, sub(v("δ"), v("t")), PLUS, sub(v("ε"), v("i,t"))),
  scale_abs: cat(ln(d(cat(sub(v("c"), v("t")), v("·"), EBIit))), EQ, ln(sub(v("c"), v("t"))), PLUS, lnEBIit),
  semi_el: cat(v("100"), v("·"), d(cat(sup(v("e"), sub(v("β"), v("j"))), MINUS, v("1")))),
  vif: cat(sub(t("VIF"), v("j")), EQ, fracL(v("1"), d(cat(v("1"), MINUS, subsup(v("R"), v("j"), v("2")))))),
  ols_region: cat(lnEBIit, EQ, sub(v("β"), v("0")), PLUS, v("β"), PR, sub(v("X"), v("i,t")), PLUS,
    sum(v("k=1"), v("5"), cat(sub(v("θ"), v("k")), sub(v("D"), v("k,i")))), PLUS, sub(v("δ"), v("t")), PLUS, sub(v("ε"), v("i,t"))),
  theta_pct: cat(v("100"), v("·"), d(cat(sup(v("e"), sub(v("θ"), v("k"))), MINUS, v("1")))),
  wald: cat(sub(v("H"), v("0")), v(":"), sub(v("θ"), v("1")), EQ, v("…"), EQ, sub(v("θ"), v("5")), EQ, v("0")),

  oaxaca_g: cat(Yb("g"), EQ, Xb("g"), PR, bh("g"), t(", "), v("g"), v("∈"), d(cat(v("L"), t(", "), v("N")), "{", "}")),
  oaxaca: cat(Yb("L"), MINUS, Yb("N"), EQ, d(cat(Xb("L"), MINUS, Xb("N"))), PR, bstar, PLUS,
    d(cat(Xb("L"), PR, d(cat(bh("L"), MINUS, bstar)), PLUS, Xb("N"), PR, d(cat(bstar, MINUS, bh("N")))), "[", "]")),
  oaxaca_E: cat(v("E"), EQ, d(cat(Xb("L"), MINUS, Xb("N"))), PR, bstar),
  oaxaca_C: cat(v("C"), EQ, Xb("L"), PR, d(cat(bh("L"), MINUS, bstar)), PLUS, Xb("N"), PR, d(cat(bstar, MINUS, bh("N")))),
  oaxaca_EC: cat(Yb("L"), MINUS, Yb("N"), EQ, v("E"), PLUS, v("C")),
  bstar_L: cat(bstar, EQ, bh("L")),
  oaxaca_j: cat(sub(v("E"), v("j")), EQ, d(cat(sub(bar(v("X")), v("L,j")), MINUS, sub(bar(v("X")), v("N,j")))),
    subsup(v("β"), v("j"), v("*")), COMMA, v("E"), EQ, sum(v("j"), null, sub(v("E"), v("j")))),
  oaxaca_j_share: cat(v("100"), v("·"), fracL(sub(v("E"), v("j")), d(cat(Yb("L"), MINUS, Yb("N"))))),
};

module.exports = { M };

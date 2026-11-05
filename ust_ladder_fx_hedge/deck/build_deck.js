// 米国債ラダー × Dynamic FX Hedging マーケティング資料（社内承認前ドラフト）
// 実行: NODE_PATH=<pptxgenjs のある node_modules> node deck/build_deck.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const D = require(path.join(__dirname, "data.json"));

const THEME = {
  name: "UST Ladder FX",
  headFontFace: "Yu Gothic",
  bodyFontFace: "Yu Gothic",
  colors: {
    dk1: "0B0B0B", lt1: "FFFFFF", dk2: "14213D", lt2: "F2F4F7",
    accent1: "2A78D6", accent2: "EB6834", accent3: "1BAF7A", accent4: "898781",
    accent5: "C3C2B7", accent6: "6250D6", hlink: "2A78D6", folHlink: "6250D6",
  },
};
const HEX = THEME.colors;
const GRID = "E1E0D9", MUTED = "6B6A66", SEC = "52514E";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "米国債ラダー × Dynamic FX Hedging";
pres.author = "TBC";
const C = pres.SchemeColor;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 1.7, w: 11.5, h: 0.5, fontSize: 14, color: C.accent2, bold: true, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 11.7, h: 1.6, fontSize: 40, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.1, w: 11.5, h: 1.6, fontSize: 16, color: "CADCFC", margin: 0, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  margin: [0.5, 0.5, 0.7, 0.5],
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.5, y: 0.3, w: 9, h: 0.32, fontSize: 11, bold: true, color: C.accent1, margin: 0, charSpacing: 1 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.62, w: 12.3, h: 0.78, fontSize: 22, align: "left", bold: true, color: C.text2, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "src", type: "body", x: 0.5, y: 6.78, w: 11.2, h: 0.5, fontSize: 8, color: MUTED, margin: 0, valign: "top" }, text: "" } },
    { text: { text: "社内承認前ドラフト｜機関投資家向け｜無断転載禁止", options: { x: 8.3, y: 0.3, w: 4.53, h: 0.32, fontSize: 9, color: MUTED, align: "right", margin: 0 } } },
  ],
  slideNumber: { x: 12.33, y: 6.95, w: 0.5, h: 0.3, fontSize: 9, color: MUTED, align: "right" },
});

// ---------- helpers ----------
let n = 0;
const oname = (s) => `${s}-${++n}`;
function content(section, kicker, title, src) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  s.addText(src, { placeholder: "src" });
  return s;
}
function text(s, t, o) {
  s.addText(t, Object.assign({ isTextBox: true, margin: 0, fontSize: 12, color: C.text1, valign: "top", objectName: oname("txt") }, o));
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { color: fill || C.background2, width: 0 }, objectName: oname("card") });
}
function stat(s, x, y, w, big, label, sub, color) {
  card(s, x, y, w, 1.25);
  text(s, big, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: 0.55, fontSize: 26, bold: true, color: color || C.accent1 });
  text(s, label, { x: x + 0.2, y: y + 0.66, w: w - 0.4, h: 0.28, fontSize: 11, bold: true });
  if (sub) text(s, sub, { x: x + 0.2, y: y + 0.93, w: w - 0.4, h: 0.28, fontSize: 9, color: SEC });
}
function tbcBox(s, x, y, w, h, msg) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: "FFF6F1" }, line: { color: HEX.accent2, width: 1, dashType: "dash" }, objectName: oname("tbc") });
  text(s, [{ text: "TBC  ", options: { bold: true, color: C.accent2, fontSize: 14 } }, { text: msg, options: { fontSize: 11, color: C.text1 } }],
    { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: h - 0.3, valign: "middle" });
}
function lineChart(s, series, o) {
  const base = {
    x: o.x, y: o.y, w: o.w, h: o.h, objectName: oname("chart"),
    chartColors: o.colors, lineSize: 2, lineDataSymbol: o.symbols ? "circle" : "none", lineDataSymbolSize: 6,
    showTitle: true, title: o.title, titleFontSize: 11, titleColor: HEX.dk1, titleFontFace: "+mn-lt", titleAlign: "left",
    showLegend: series.length > 1, legendPos: "b", legendFontSize: 9, legendFontFace: "+mn-lt", legendColor: SEC,
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    catAxisLineShow: true, catAxisMajorTickMark: "none", valAxisMajorTickMark: "none", catAxisLineColor: "C3C2B7", valAxisLineShow: false,
    valGridLine: { color: GRID, size: 0.5 }, catGridLine: { style: "none" },
    catAxisLabelFrequency: o.freq || 1, catAxisOrientation: "minMax", catAxisLabelRotate: 0,
    valAxisMinVal: o.min, valAxisMaxVal: o.max, valAxisMajorUnit: o.unit, valAxisLabelFormatCode: o.fmt || "0.0",
    showValue: false,
  };
  s.addChart(pres.charts.LINE, series, base);
}
const TBL = { fontSize: 10.5, color: HEX.dk1, border: { type: "solid", pt: 0.5, color: GRID }, valign: "middle", margin: [0.04, 0.08, 0.04, 0.08] };
const th = (t, o) => ({ text: t, options: Object.assign({ bold: true, fill: { color: HEX.dk2 }, color: HEX.lt1 }, o) });
const td = (t, o) => ({ text: t, options: o || {} });
const tbc = (t) => ({ text: t || "TBC", options: { bold: true, color: HEX.accent2 } });

const K = D.k;
const SRC_MKT = "出所：米国財務省、FRB（H.15・H.10、FRED）、FOMC声明・SEP、財務省（日本）、日本銀行。データ取得日2026/11/5。";

// ================= Cover =================
pres.addSection({ title: "表紙" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "表紙" });
  s.addText("FOR REGIONAL BANK FIXED-INCOME INVESTORS", { placeholder: "kicker" });
  s.addText("米国債ラダー × Dynamic FX Hedging", { placeholder: "title" });
  s.addText([
    { text: "米国債の利回りを、円建てで取り込むための為替ヘッジの新しい選択肢", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "2026年11月｜社内承認前ドラフト｜市場データ基準日 2026/11/4（一部 11/2〜11/3）", options: { fontSize: 12 } },
  ], { placeholder: "body" });
  s.addShape(pres.shapes.OVAL, { x: 10.6, y: 4.9, w: 1.9, h: 1.9, fill: { color: HEX.accent1, transparency: 70 }, line: { color: HEX.accent1, width: 0 }, objectName: "cover-motif-1" });
  s.addShape(pres.shapes.OVAL, { x: 11.5, y: 5.6, w: 1.2, h: 1.2, fill: { color: HEX.accent2, transparency: 55 }, line: { color: HEX.accent2, width: 0 }, objectName: "cover-motif-2" });
  s.addNotes("表紙。顧客名・担当者名は記載しない。商品名・運用会社名は正式名称の確定後に記載（TBC）。");
}

// ================= 1. Why now =================
pres.addSection({ title: "Why Now" });
{
  const s = content("Why Now", "01 ｜ WHY US TREASURIES NOW?",
    "米国債利回りは2008年以降の最高圏。1〜10年ラダーでも平均4.7%程度",
    SRC_MKT + " ラダー利回りは1〜10年に均等投資したと仮定した参考値（パー利回りの線形補間・単純平均）で、ファンドの実際のYTMではありません（TBC）。過去の利回り水準は将来を示唆しません。");
  lineChart(s, [{ name: "米国10年国債", labels: D.c1.labels, values: D.c1.v }],
    { x: 0.5, y: 1.5, w: 6.2, h: 3.25, colors: [HEX.accent1], title: "米国10年国債利回り（月末、%）2000/1〜2026/11", freq: 60, min: 0, max: 7, unit: 1 });
  lineChart(s, [
    { name: "米国債 2026/11/04", labels: D.c2.labels, values: D.c2.ust },
    { name: "米国債 2025/11/04", labels: D.c2.labels, values: D.c2.ust25 },
    { name: "日本国債 2026/11/04", labels: D.c2.labels, values: D.c2.jgb },
  ], { x: 6.9, y: 1.5, w: 5.95, h: 3.25, colors: [HEX.accent1, "9CC0EC", HEX.accent2], symbols: true, title: "利回り曲線（%、横軸は残存年数）", min: 0, max: 6, unit: 1 });
  const y = 4.95, w = 2.95, g = 0.17;
  stat(s, 0.5, y, w, "4.90%", "米国10年国債（11/4）", "2008年以降の日次データの上位約1%");
  stat(s, 0.5 + (w + g), y, w, `${K.lad.toFixed(2)}%`, "米国債ラダー参考利回り", "1〜10年均等（仮定）・ドル建て");
  stat(s, 0.5 + 2 * (w + g), y, w, "3.75–4.00%", "FF金利誘導目標", "2026/9に利上げ再開、10/28据え置き");
  stat(s, 0.5 + 3 * (w + g), y, w, "+0.80%pt", "米10年債の1年間の上昇幅", "4.10%（2025/11/4）→ 4.90%", C.accent2);
  text(s, [
    { text: "リスク：", options: { bold: true, color: C.accent2 } },
    { text: "インフレ持続（FOMC見通しの2026年PCEは3.7%）による追加利上げ・債券価格下落。10年債は10/5の5.31%から約0.4%pt低下しており、値動きが大きい。" },
  ], { x: 0.5, y: 6.3, w: 12.3, h: 0.4, fontSize: 10.5 });
  s.addNotes("ラダー投資の意義：毎年一部が満期を迎え再投資されるため、金利上昇時は高い利回りで再投資でき、金利変動の影響が時間的に分散される。1年4.39%と10年4.90%の差は0.51%ptで、満期を分散しても利回りをあまり犠牲にしない。FOMC見通し中央値：2026年末4.1%、2027年末4.1%、2028年末3.9%、長期3.2%（2026/9/16）。市場が織り込む利上げ・利下げ予想（FedWatch／OIS）はTBC。");
}

// ================= 2. Cost of full hedging =================
pres.addSection({ title: "Investment Challenge" });
{
  const s = content("Investment Challenge", "02 ｜ THE COST OF FULL HEDGING",
    "フルヘッジ後の米国債ラダーは約2.0%。日本国債ラダー（2.30%）を下回る",
    "出所：米国財務省、財務省（日本）、FRB（FRED: DGS10, DFF, SOFR）、日本銀行（STRDCLUCON）より作成。ヘッジコストは日米の短期金利差による近似で、為替フォワードのベーシス・取引コスト・信託報酬は未反映。実際のヘッジコストは一般にこれより高い（フォワード実勢：TBC）。");
  // custom bars
  text(s, "米国債ラダーのヘッジ後参考利回り（%、年率）", { x: 0.5, y: 1.5, w: 6, h: 0.3, fontSize: 11, bold: true });
  const base = 5.35, scale = 0.62, bw = 1.05, x0 = 0.75, gap = 0.42;
  const bars = [
    { lab: "米国債ラダー\n1〜10年", v: K.lad, from: 0, col: HEX.accent1, txt: `${K.lad.toFixed(2)}%` },
    { lab: "ヘッジコスト目安\nSOFR−TONA", v: K.hc, from: K.hl, col: "B7B6B0", txt: `−${K.hc.toFixed(2)}%` },
    { lab: "フルヘッジ後\n参考利回り", v: K.hl, from: 0, col: "9CC0EC", txt: `${K.hl.toFixed(2)}%` },
    { lab: "日本国債ラダー\n1〜10年", v: K.jl, from: 0, col: HEX.accent2, txt: `${K.jl.toFixed(2)}%` },
  ];
  bars.forEach((b, i) => {
    const x = x0 + i * (bw + gap), top = base - (b.from + b.v) * scale;
    s.addShape(pres.shapes.RECTANGLE, { x, y: top, w: bw, h: b.v * scale, fill: { color: b.col }, line: { color: b.col, width: 0 }, objectName: oname("bar") });
    text(s, b.txt, { x: x - 0.2, y: top - 0.36, w: bw + 0.4, h: 0.32, fontSize: 14, bold: true, align: "center" });
    text(s, b.lab, { x: x - 0.25, y: base + 0.08, w: bw + 0.5, h: 0.5, fontSize: 9.5, align: "center", color: SEC });
  });
  s.addShape(pres.shapes.LINE, { x: 0.6, y: base, w: 5.9, h: 0, line: { color: "C3C2B7", width: 1 }, objectName: "bars-baseline" });
  lineChart(s, [
    { name: "米国10年債（フルヘッジ後の目安）", labels: D.c6.labels, values: D.c6.h },
    { name: "日本10年国債", labels: D.c6.labels, values: D.c6.j },
  ], { x: 6.9, y: 1.5, w: 5.95, h: 4.2, colors: [HEX.accent1, HEX.accent2], title: "米国10年債（ヘッジ後）と日本10年国債（月平均、%）", freq: 36, min: -2, max: 4, unit: 1 });
  card(s, 0.5, 6.0, 12.33, 0.66, "EEF4FC");
  text(s, [
    { text: "2022年11月から逆転が継続。", options: { bold: true } },
    { text: "差は−2.13%pt（2024/8）から−0.49%pt（2026/10）まで縮小したが、FRBも2026年9月に利上げを再開しており、ヘッジコストの低下は不確実。" },
  ], { x: 0.7, y: 6.08, w: 12.0, h: 0.5, fontSize: 11, valign: "middle" });
  s.addNotes("ヘッジコスト目安2.72%＝SOFR 3.95% − TONA 1.227%（2026/11/2）。月平均ベース（実効FF金利−TONA）では2023年8月の5.39%がピーク、2026年10月は2.65%。ヘッジなしとの比較は補足扱い：ドル円は2026/7/29に163.86円（1986年以来）、2000年以降の各日から1年間で見て10%以上円高になった期間は全体の11%。");
}

// ================= 3. New approach =================
pres.addSection({ title: "Proposed Solution" });
{
  const s = content("Proposed Solution", "03 ｜ A NEW APPROACH TO FX HEDGING",
    "ヘッジ比率は「利回り」と「円高への耐性」のトレードオフ。固定せず調整する",
    `出所：米国財務省、財務省（日本）、FRB（FRED）、日本銀行より作成。米国債ラダー参考利回り${K.lad.toFixed(2)}%、ヘッジコスト目安${K.hc.toFixed(2)}%で一定、債券価格の変動・費用なしとした1年間の簡易試算。将来の運用成果を示すものではありません。`);
  lineChart(s, [
    { name: "ヘッジなし（0%）", labels: D.c8.labels, values: D.c8.h0 },
    { name: "50%固定ヘッジ", labels: D.c8.labels, values: D.c8.h50 },
    { name: "フルヘッジ（100%）", labels: D.c8.labels, values: D.c8.h100 },
    { name: "日本国債ラダー", labels: D.c8.labels, values: D.c8.jgb },
  ], { x: 0.5, y: 1.5, w: 6.4, h: 4.6, colors: [HEX.accent2, "898781", HEX.accent1, HEX.accent3], title: "1年間のUSD/JPY変化率（横軸、マイナス＝円高）と円建てリターンの概算（%）", freq: 5, min: -12, max: 14, unit: 4 });
  text(s, `3方式のリターンが一致するのは、ドル円が${K.hc.toFixed(1)}%円高（157.5円→約153円）になったとき`, { x: 0.5, y: 6.15, w: 6.4, h: 0.45, fontSize: 10.5, bold: true, color: SEC });
  const rows = [
    [th("", { fill: { color: HEX.lt1 } }), th("ヘッジなし"), th("50%固定"), th("フルヘッジ"), th("動的ヘッジ", { fill: { color: HEX.accent1 } })],
    [td("参考利回り（ヘッジ後）", { bold: true }), td(`${K.lad.toFixed(2)}%`), td(`${(K.lad - K.hc / 2).toFixed(2)}%`), td(`${K.hl.toFixed(2)}%`), td("ヘッジ比率に応じて変動")],
    [td("10%円高時の為替損益（概算）", { bold: true }), td("−10%"), td("−5%"), td("0%"), tbc("バックテストで確認")],
    [td("ヘッジコストの負担", { bold: true }), td("なし"), td("一定（半分）"), td("最大"), td("ヘッジ比率に応じて変動")],
    [td("主な弱点", { bold: true }), td("円高で大きな損失"), td("どちらの局面でも中途半端"), td("利回りが円債以下"), td("判断の遅れ・急な反転への弱さ")],
  ];
  s.addTable(rows, Object.assign({}, TBL, { x: 7.15, y: 1.55, w: 5.68, colW: [1.45, 0.98, 1.0, 1.0, 1.25], rowH: [0.42, 0.6, 0.6, 0.6, 0.75], fontSize: 10, objectName: "approach-table" }));
  card(s, 7.15, 4.85, 5.68, 1.3, "EEF4FC");
  text(s, [
    { text: "動的ヘッジは為替リスクがない商品ではありません。", options: { bold: true, breakLine: true } },
    { text: "円高リスクの高まりに応じてヘッジ比率を変え、利回りの確保と円高時の損失抑制のバランスを取ることを目指す運用です。" },
  ], { x: 7.35, y: 4.98, w: 5.3, h: 1.05, fontSize: 11 });
  s.addNotes("「円高を回避できる」「損失を防ぐ」とは言わない。「抑制を目指す」と表現する。50%固定の参考利回り＝4.70−2.72/2。2000年以降の各日から1年間で見て、2.7%を超える円高になった期間は全体の34%。");
}

// ================= 4. Strategy at a glance =================
{
  const s = content("Proposed Solution", "04 ｜ STRATEGY AT A GLANCE",
    "米国債ラダー＋ヘッジ比率を調整する為替オーバーレイの円建てファンド",
    "出所：運用会社・戦略提供会社の資料（TBC）。図は商品構造の概念図であり、ヘッジ比率の範囲・変更基準・頻度は正式資料の受領後に記載します。");
  // structure flow
  const bx = (x, y, w, h, t, sub, fill, fc) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: oname("node") });
    text(s, [{ text: t, options: { bold: true, fontSize: 13, color: fc, breakLine: !!sub } }].concat(sub ? [{ text: sub, options: { fontSize: 10, color: fc } }] : []),
      { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: h - 0.2, align: "center", valign: "middle" });
  };
  const arrow = (x, y, w, h) => s.addShape(pres.shapes.LINE, { x, y, w, h, line: { color: "898781", width: 1.5, endArrowType: "triangle" }, objectName: oname("arrow") });
  text(s, "商品構造", { x: 0.5, y: 1.5, w: 6, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
  bx(0.5, 2.0, 2.2, 1.0, "投資家", "地域金融機関", HEX.lt2, HEX.dk1);
  arrow(2.75, 2.5, 0.55, 0);
  bx(3.35, 2.0, 2.6, 1.0, "円建てファンド", "形態・運用会社 TBC", HEX.dk2, HEX.lt1);
  arrow(4.65, 3.05, 0, 0.45); arrow(4.65, 3.05, -1.6, 0.45); arrow(4.65, 3.05, 1.4, 0.45);
  bx(2.1, 3.55, 2.0, 1.1, "米国債ラダー", "1〜10年（TBC）", HEX.accent1, HEX.lt1);
  bx(4.85, 3.55, 2.2, 1.1, "為替ヘッジ・オーバーレイ", "比率 TBC%〜TBC%", HEX.accent2, HEX.lt1);
  bx(4.85, 5.15, 2.2, 0.8, "戦略提供会社", "ヘッジ比率のシグナル", HEX.lt2, HEX.dk1);
  arrow(5.95, 5.1, 0, -0.42);
  // hedge ratio concept
  text(s, "ヘッジ比率の考え方（概念図）", { x: 7.6, y: 1.5, w: 5.2, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
  const lv = [
    ["低ヘッジ", "円高リスクが低いと判断される局面：利回りを重視", "D6E6F8"],
    ["中間", "シグナルに応じて段階的または連続的に調整（TBC）", "9CC0EC"],
    ["高ヘッジ", "円高リスクが高いと判断される局面：損失抑制を重視", HEX.accent1],
  ];
  lv.forEach((l, i) => {
    const y = 2.0 + i * 0.9;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.6, y, w: 1.5, h: 0.75, rectRadius: 0.08, fill: { color: l[2] }, line: { color: l[2], width: 0 }, objectName: oname("lvl") });
    text(s, l[0], { x: 7.6, y, w: 1.5, h: 0.75, align: "center", valign: "middle", bold: true, fontSize: 12, color: i === 2 ? C.background1 : C.text2 });
    text(s, l[1], { x: 9.3, y: y + 0.05, w: 3.5, h: 0.68, fontSize: 10.5, valign: "middle" });
  });
  text(s, [{ text: "TBC：", options: { bold: true, color: C.accent2 } }, { text: "ヘッジ比率の上限・下限、変更の頻度、シグナル算出から約定までの時間差" }],
    { x: 7.6, y: 4.75, w: 5.2, h: 0.35, fontSize: 10 });
  // merits / weaknesses
  card(s, 7.6, 5.2, 2.55, 1.45, "EEF4FC");
  text(s, [{ text: "期待されるメリット", options: { bold: true, breakLine: true } }, { text: "フルヘッジより高い利回りの確保と、ヘッジなしより小さい円高時の損失を目指す", options: { fontSize: 10 } }], { x: 7.75, y: 5.3, w: 2.3, h: 1.3, fontSize: 11 });
  card(s, 10.28, 5.2, 2.55, 1.45, "FFF1EA");
  text(s, [{ text: "想定される弱点", options: { bold: true, breakLine: true } }, { text: "急な円高への反応の遅れ、ヘッジ比率の変更が裏目に出る局面、ヘッジ変更に伴う取引コスト", options: { fontSize: 10 } }], { x: 10.43, y: 5.3, w: 2.3, h: 1.3, fontSize: 11 });
  text(s, "円建て・米国債のみ。損益の主な要因は米国金利、為替（ヘッジしていない部分）、ヘッジコスト、費用", { x: 0.5, y: 6.15, w: 6.6, h: 0.45, fontSize: 10, color: SEC });
  s.addNotes("シグナルの詳細やモデルは説明しない。ヘッジ比率の範囲・頻度・切替基準は提供会社の資料で確認し、確認できるまでTBCとする。");
}

// ================= 5. Provider =================
pres.addSection({ title: "Strategy Provider" });
{
  const s = content("Strategy Provider", "05 ｜ STRATEGY PROVIDER",
    "戦略提供会社のプロフィール：規模・実績・体制を正式資料で確認のうえ記載",
    "出所：戦略提供会社の正式資料（TBC）。運用資産・顧客数・実績は必ず出典と基準日を付けて記載し、未確認の数値・顧客名は記載しません。");
  const rows = [
    [th("項目", { w: 2.2 }), th("内容"), th("出典・基準日")],
    [td("会社名・設立年・本拠地", { bold: true }), tbc(), tbc()],
    [td("規制登録（国・登録の種類）", { bold: true }), tbc(), tbc()],
    [td("運用資産・助言資産", { bold: true }), tbc("TBC（運用ベースか助言ベースか、為替戦略に限った額を区別）"), tbc()],
    [td("為替戦略の提供実績", { bold: true }), tbc("TBC（実運用の期間、実績とシミュレーションの区別）"), tbc()],
    [td("主要な顧客層", { bold: true }), tbc("TBC（開示可能な範囲）"), tbc()],
    [td("運用チーム", { bold: true }), tbc("TBC（主要メンバーの経歴・年数）"), tbc()],
    [td("運用プロセス・リスク管理", { bold: true }), tbc("TBC（モデル管理、ヘッジ比率の制約、緊急時の対応）"), tbc()],
    [td("他社との違い", { bold: true }), tbc(), tbc()],
  ];
  s.addTable(rows, Object.assign({}, TBL, { x: 0.5, y: 1.55, w: 7.9, colW: [2.3, 4.2, 1.4], rowH: 0.5, objectName: "provider-table" }));
  text(s, "機関投資家が評価する4つの観点", { x: 8.8, y: 1.55, w: 4.0, h: 0.32, fontSize: 13, bold: true, color: C.text2 });
  const pts = [
    ["組織の安定性", "沿革、株主、人員の推移、規制登録"],
    ["実績の検証可能性", "実運用の期間、第三者による検証の有無"],
    ["プロセスの再現性", "開発日、モデル変更の履歴、事後調整の有無"],
    ["リスク管理", "ヘッジ比率の制約、モデルの監視体制"],
  ];
  pts.forEach((p, i) => {
    const y = 2.0 + i * 1.0;
    card(s, 8.8, y, 4.03, 0.85);
    s.addShape(pres.shapes.OVAL, { x: 8.95, y: y + 0.2, w: 0.45, h: 0.45, fill: { color: HEX.accent1 }, line: { color: HEX.accent1, width: 0 }, objectName: oname("dot") });
    text(s, String(i + 1), { x: 8.95, y: y + 0.2, w: 0.45, h: 0.45, align: "center", valign: "middle", bold: true, color: C.background1, fontSize: 13 });
    text(s, [{ text: p[0], options: { bold: true, breakLine: true } }, { text: p[1], options: { fontSize: 10, color: SEC } }], { x: 9.55, y: y + 0.1, w: 3.2, h: 0.7, fontSize: 12 });
  });
  s.addNotes("知名度の低さを補うページ。数値の誇張は禁止。未確認事項はTBCのまま残し、正式資料の受領後に出典・基準日と合わせて記載する。");
}

// ================= 6. Return & Yield =================
pres.addSection({ title: "Evidence" });
{
  const s = content("Evidence", "06 ｜ RETURN & YIELD COMPARISON",
    "「現時点の利回り」と「過去の実績」は性質が異なるため、分けて比較する",
    "出所：米国財務省、財務省（日本）、FRB（FRED: SOFR）、日本銀行より作成（2026/11/4、ヘッジコストは11/2）。①はファンドの実際の保有債券のYTMが確定するまで仮定値。ブロック2は過去の検証結果であり、将来の利回りや確定利回りを示すものではありません。");
  text(s, "ブロック1：現時点の利回り（推計値・為替の変動は含まない）", { x: 0.5, y: 1.5, w: 12.3, h: 0.32, fontSize: 13, bold: true, color: C.text2 });
  const r1 = [
    [th("指標"), th("定義"), th("現時点"), th("用途")],
    [td("① 米国債ラダーのYTM", { bold: true }), td("ファンドが保有する債券の時価加重の最終利回り（ドル建て）"), td(`${K.lad.toFixed(2)}%（仮定値）／実際は TBC`, { bold: true }), td("債券部分の利回り水準")],
    [td("② フルヘッジ後の参考利回り", { bold: true }), td("① − ヘッジコスト（正式版は3か月フォワードの実勢）"), td(`${K.hl.toFixed(2)}%（近似）`, { bold: true }), td("フルヘッジ外債との比較の基準")],
    [td("③ 現在のヘッジ比率を反映した参考利回り", { bold: true }), td("① − ヘッジ比率 × ヘッジコスト"), tbc(), td("その時点の利回り（確定利回りではない）")],
    [td("比較：日本国債ラダー", { bold: true }), td("日本国債1〜10年の均等平均（財務省公表の複利利回り）"), td(`${K.jl.toFixed(2)}%`, { bold: true }), td("円債の代替比較")],
  ];
  s.addTable(r1, Object.assign({}, TBL, { x: 0.5, y: 1.9, w: 12.33, colW: [3.0, 4.6, 2.3, 2.43], rowH: 0.46, objectName: "yield-now" }));
  text(s, "ブロック2：過去の実績（バックテスト・シミュレーション）", { x: 0.5, y: 4.4, w: 12.3, h: 0.32, fontSize: 13, bold: true, color: C.text2 });
  const r2 = [
    [th("指標"), th("定義"), th("フルヘッジ"), th("Dynamic"), th("ヘッジなし（補足）")],
    [td("④ 年率リターン（費用控除前）", { bold: true }), td("検証期間の円建てトータルリターンを年率換算"), tbc(), tbc(), tbc()],
    [td("⑤ 年率リターン（費用控除後）", { bold: true }), td("④ − ヘッジコスト・取引コスト・信託報酬"), tbc(), tbc(), tbc()],
  ];
  s.addTable(r2, Object.assign({}, TBL, { x: 0.5, y: 4.8, w: 12.33, colW: [3.0, 4.6, 1.55, 1.55, 1.63], rowH: 0.46, objectName: "yield-hist" }));
  text(s, "④⑤は「利回り」と呼ばず、ブロック1と同じ列に並べない。検証期間・データ・費用の前提は TBC。", { x: 0.5, y: 6.3, w: 12.3, h: 0.35, fontSize: 10, color: SEC });
  s.addNotes("③は「確定利回り」「期待利回り」と呼ばない。為替の損益を含まないことを口頭でも補足する。");
}

// ================= 7. Case study 1 =================
{
  const s = content("Evidence", "07 ｜ HISTORICAL CASE STUDY — JPY APPRECIATION",
    "大幅な円高局面での検証①：リーマン・ショック（2008年）と2022〜23年",
    "出所：FRB H.10（FRED: DEXJPUS、NY正午の値、日次）、財務省「外国為替平衡操作の実施状況」。局面は「2005年以降に3か月で10%以上の円高」などの基準で選定。戦略のヘッジ比率と損益はシグナルデータの受領後に算出（TBC）。");
  lineChart(s, [{ name: "USD/JPY", labels: D.e08.labels, values: D.e08.v }],
    { x: 0.5, y: 1.5, w: 6.05, h: 2.75, colors: [HEX.accent1], title: "局面B：2008/7〜2009/1 USD/JPY（円/ドル）", freq: 30, min: 80, max: 115, unit: 5, fmt: "0" });
  lineChart(s, [{ name: "USD/JPY", labels: D.e22.labels, values: D.e22.v }],
    { x: 6.78, y: 1.5, w: 6.05, h: 2.75, colors: [HEX.accent1], title: "局面E：2022/9〜2023/2 USD/JPY（円/ドル）", freq: 25, min: 120, max: 155, unit: 5, fmt: "0" });
  text(s, [{ text: "110.48 → 87.84（−20.5%）", options: { bold: true, breakLine: true } }, { text: "2008/8/15 → 12/17。介入なし。米10年債は同期間に−1.72%pt", options: { fontSize: 10, color: SEC } }], { x: 0.6, y: 4.3, w: 5.9, h: 0.6, fontSize: 12 });
  text(s, [{ text: "149.82 → 127.85（−14.7%）", options: { bold: true, breakLine: true } }, { text: "2022/10/20 → 2023/1/13。円買い介入 10/21・10/24（財務省）。日銀が長期金利操作を修正（2022/12）", options: { fontSize: 10, color: SEC } }], { x: 6.88, y: 4.3, w: 5.9, h: 0.6, fontSize: 12 });
  tbcBox(s, 0.5, 5.1, 12.33, 1.5, "同じ時間軸でヘッジ比率と、フルヘッジ外債との累積損益差を表示（シグナルデータの受領後に算出）。各局面で ①期間と背景 ②USD/JPY ③ヘッジ比率 ④フルヘッジとの差 ⑤損失抑制効果 ⑥十分に機能しなかった要因 を記載。");
  s.addNotes("候補局面：A 2007/6〜2008/3（−21.9%）、B 2008/8〜12（−20.5%）、C 2016/1〜6（−16.0%）、E 2022/10〜2023/1（−14.7%）、G 2024/7〜9（−13.0%）。選定基準を必ず併記し、都合の良い局面だけを選ばない。");
}

// ================= 8. Case study 2 =================
{
  const s = content("Evidence", "08 ｜ HISTORICAL CASE STUDY — FX INTERVENTION",
    "為替介入前後の検証②：円高が続いた2024年と、円安に戻った2026年",
    "出所：FRB H.10（FRED: DEXJPUS）、財務省「外国為替平衡操作の実施状況」（日次：〜2026年6月期、月次：〜2026/10/28）。NY正午の日次値のため、介入当日の日中の高値・安値は反映していません。介入日の表示は戦略の予測能力を示すものではありません。");
  lineChart(s, [{ name: "USD/JPY", labels: D.e24.labels, values: D.e24.v }],
    { x: 0.5, y: 1.5, w: 6.05, h: 2.6, colors: [HEX.accent1], title: "局面F・G：2024/4〜9 USD/JPY（円/ドル）", freq: 25, min: 135, max: 165, unit: 5, fmt: "0" });
  lineChart(s, [{ name: "USD/JPY", labels: D.e26.labels, values: D.e26.v }],
    { x: 6.78, y: 1.5, w: 6.05, h: 2.6, colors: [HEX.accent1], title: "局面H〜J：2026/4〜10 USD/JPY（円/ドル）", freq: 30, min: 150, max: 166, unit: 2, fmt: "0" });
  const r = [
    [th("局面"), th("円買い介入（財務省）"), th("ドル円の最大下落（期間内）"), th("その後")],
    [td("G 2024/7〜9", { bold: true }), td("7/11・7/12 計5.5兆円"), td("161.73→140.66（−13.0%）"), td("日銀利上げ（7/31）も重なり円高が継続")],
    [td("H 2026/4〜6", { bold: true }), td("4/30・5/4・5/6 計11.7兆円"), td("160.23→156.44（−2.4%）"), td("6月末に162.61まで円安に戻る（効果が限定的）")],
    [td("I 2026/7〜8", { bold: true }), td("7/30〜8/26の期間に計15.4兆円（日次は11/9公表）"), td("163.86→156.96（−4.2%）"), tbc("実施日は公表後に確定")],
    [td("J 2026/9", { bold: true }), td("なし（月次の公表で0円）"), td("160.03→153.50（−4.1%）"), td("介入なしでも約1週間で急な円高")],
  ];
  s.addTable(r, Object.assign({}, TBL, { x: 0.5, y: 4.2, w: 12.33, colW: [1.6, 3.6, 3.0, 4.13], rowH: 0.37, fontSize: 10, objectName: "intervention-table" }));
  tbcBox(s, 0.5, 6.1, 12.33, 0.6, "各局面のヘッジ比率の推移とフルヘッジとの損益差（効果が限定的だった局面 D・F・H を含めて表示）");
  s.addNotes("「介入を予測した」「介入を察知した」とは表現しない。シグナルが出た時点と約定の時点を明記する。2022年の介入（9/22・10/21・10/24、計9.2兆円）は局面Dとして補足。");
}

// ================= 9. Long-term backtest =================
{
  const s = content("Evidence", "09 ｜ LONG-TERM BACKTEST & RISK-RETURN PROFILE",
    "長期バックテスト：フルヘッジ・Dynamic・ヘッジなしの比較",
    "出所：バックテスト結果は運用会社（TBC）。右図はFRB（FRED: DFF）・日本銀行（STRDCLUCON）より作成した実際のデータ。バックテストはシミュレーションであり実際の運用実績ではありません。対象が為替部分だけかファンド全体か、費用の扱い、検証期間・開発日はTBC。");
  const r = [
    [th("指標"), th("フルヘッジ"), th("Dynamic", { fill: { color: HEX.accent1 } }), th("ヘッジなし（補足）")],
    ...["累積リターン", "年率リターン", "ボラティリティ", "最大ドローダウン", "円高局面での損失", "ヘッジコストの累計", "平均ヘッジ比率", "売買回転率"].map((k) => [td(k, { bold: true }), tbc(), tbc(), tbc()]),
  ];
  s.addTable(r, Object.assign({}, TBL, { x: 0.5, y: 1.5, w: 6.6, colW: [2.1, 1.5, 1.5, 1.5], rowH: 0.42, objectName: "bt-table" }));
  text(s, [{ text: "前提（TBC）：", options: { bold: true, color: C.accent2 } }, { text: "対象（為替部分のみ／ファンド全体）・検証期間・戦略の開発日・ベンチマーク系列・取引コスト・追加30bpの扱い" }],
    { x: 0.5, y: 5.45, w: 6.6, h: 0.6, fontSize: 10 });
  lineChart(s, [{ name: "短期金利差", labels: D.c4.labels, values: D.c4.v }],
    { x: 7.4, y: 1.5, w: 5.43, h: 3.6, colors: [HEX.accent1], title: "検証期間のヘッジコスト環境：米実効FF金利−無担保コール翌日物（月平均、%）", freq: 60, min: 0, max: 6, unit: 1 });
  card(s, 7.4, 5.25, 5.43, 0.95, "EEF4FC");
  text(s, "ヘッジコストの目安は2023/8の5.39%をピークに、2026/10は2.65%。ヘッジ比率を下げた期間にどれだけヘッジコストを節約できたかが、長期の成績を左右する。", { x: 7.55, y: 5.33, w: 5.15, h: 0.8, fontSize: 10.5 });
  s.addNotes("異なる対象・条件のパフォーマンスを同等に比較しない。データが最新でない場合は『旧データによる暫定分析』と明記する。");
}

// ================= 10. Fund overview =================
pres.addSection({ title: "Investment Case" });
{
  const s = content("Investment Case", "10 ｜ FUND OVERVIEW & INVESTMENT CASE",
    "商品概要と検討のポイント",
    "出所：目論見書・運用会社の資料（TBC）。本資料は情報提供を目的とした社内承認前のドラフトであり、特定の投資を勧誘するものではありません。");
  const r = [
    [th("項目"), th("内容")],
    [td("投資対象", { bold: true }), td("米国債ラダー（年限・構成 TBC）＋ 為替ヘッジ・オーバーレイ")],
    [td("運用構造・通貨", { bold: true }), td("円建てファンド（形態・運用会社 TBC）")],
    [td("戦略提供会社", { bold: true }), tbc()],
    [td("流動性（申込・解約）", { bold: true }), tbc()],
    [td("費用（信託報酬・その他）", { bold: true }), tbc()],
    [td("追加30bp", { bold: true }), tbc("TBC（適用対象・符号・年率換算）")],
    [td("会計・規制上の扱い", { bold: true }), tbc()],
  ];
  s.addTable(r, Object.assign({}, TBL, { x: 0.5, y: 1.5, w: 6.4, colW: [2.2, 4.2], rowH: 0.42, objectName: "fund-table" }));
  text(s, "主要なリスク", { x: 7.2, y: 1.5, w: 5.6, h: 0.32, fontSize: 13, bold: true, color: C.text2 });
  text(s, [
    { text: "金利リスク：米国金利の上昇による債券価格の下落", options: { bullet: true, breakLine: true } },
    { text: "為替リスク：ヘッジしていない部分の円高による損失", options: { bullet: true, breakLine: true } },
    { text: "ヘッジコストの変動：日米の金利差・ベーシスの拡大", options: { bullet: true, breakLine: true } },
    { text: "戦略リスク：シグナルの遅れ・誤判断、モデルリスク", options: { bullet: true, breakLine: true } },
    { text: "信用・流動性・カウンターパーティーのリスク（TBC）", options: { bullet: true } },
  ], { x: 7.2, y: 1.9, w: 5.6, h: 2.0, fontSize: 11, paraSpaceAfter: 4 });
  const f = [
    ["01", "米国債の高い利回り", "2008年以降で最も高い水準の圏内にある利回りを、ラダーで分散して取り込む"],
    ["02", "固定ではないヘッジ", "フルヘッジでは日本国債ラダーを下回る利回り構造に対し、ヘッジ比率を調整"],
    ["03", "検証の透明性", "円高・介入局面での挙動を、効果が限定的だった局面も含めて開示（TBC）"],
  ];
  f.forEach((x, i) => {
    const xx = 0.5 + i * 4.18;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xx, y: 5.0, w: 3.97, h: 1.62, rectRadius: 0.08, fill: { color: HEX.dk2 }, line: { color: HEX.dk2, width: 0 }, objectName: oname("feat") });
    text(s, x[0], { x: xx + 0.2, y: 5.1, w: 0.8, h: 0.45, fontSize: 20, bold: true, color: C.accent2 });
    text(s, x[1], { x: xx + 0.95, y: 5.15, w: 2.9, h: 0.4, fontSize: 13, bold: true, color: C.background1 });
    text(s, x[2], { x: xx + 0.2, y: 5.62, w: 3.6, h: 0.95, fontSize: 10.5, color: "CADCFC" });
  });
  text(s, "検討の際の確認事項：運用ニーズとの適合、ヘッジ比率の範囲、費用、会計区分、社内の承認手続き", { x: 7.2, y: 4.05, w: 5.6, h: 0.6, fontSize: 10, color: SEC });
  s.addNotes("メリットを断定しない。投資判断に必要な確認事項を示す。");
}

// ================= Disclosures =================
pres.addSection({ title: "Disclosures" });
{
  const s = content("Disclosures", "IMPORTANT NOTES", "データの出所と重要な注意事項",
    "本資料は社内コンプライアンス審査の前のドラフトです。顧客への提示前に必要な審査を受けてください。");
  text(s, [
    { text: "市場データの出所（取得日 2026/11/5）", options: { bold: true, breakLine: true } },
    { text: "米国財務省 Daily Treasury Par Yield Curve Rates（〜2026/11/4）／FRB H.15・H.10（FRED: DGS1〜30, DFF, DFEDTARU/L, DEXJPUS）／NY連銀 SOFR（FRED）／FOMC声明（2026/9/16、10/28）・経済見通し（2026/9/16）／財務省 国債金利情報（〜2026/11/4）／日本銀行 時系列統計（無担保コール翌日物、USD/JPY）・金融政策の公表文（2026/10/30他）／財務省 外国為替平衡操作の実施状況（日次：〜2026年6月期、月次：〜2026/10/28）", options: { breakLine: true, fontSize: 10 } },
    { text: " ", options: { breakLine: true } },
    { text: "計算上の注記", options: { bold: true, breakLine: true } },
    { text: "ラダーの参考利回りは1〜10年に均等投資したと仮定した単純平均で、ファンドの実際の利回りではありません。ヘッジコストは日米の短期金利差による近似で、為替フォワードのベーシス・取引コスト・費用を含みません。米国債はパー利回り、日本国債は財務省公表の複利利回りです。", options: { breakLine: true, fontSize: 10 } },
    { text: " ", options: { breakLine: true } },
    { text: "バックテストについて", options: { bold: true, breakLine: true } },
    { text: "バックテストは過去のデータに基づくシミュレーションであり、実際の運用実績ではありません。将来の運用成果を示唆・保証するものではありません。為替ヘッジによる損失の抑制は保証されません。", options: { breakLine: true, fontSize: 10 } },
    { text: " ", options: { breakLine: true } },
    { text: "TBC 表記", options: { bold: true, breakLine: true } },
    { text: "商品仕様、戦略提供会社の情報、バックテスト結果、為替フォワードの実勢ヘッジコスト、市場が織り込む政策金利の見通し、2026年7〜9月期の介入の日次データは、正式資料・データの受領後に更新します。", options: { fontSize: 10 } },
  ], { x: 0.5, y: 1.55, w: 12.3, h: 5.0, fontSize: 12 });
}

(async () => {
  const out = path.join(__dirname, "UST_Ladder_DynamicFXHedging_draft.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})();

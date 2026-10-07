const pptxgen = require("pptxgenjs");
// apply_theme.js は pptx スキル同梱のヘルパー(テーマ色をpptxへ書き込む)
const { applyTheme } = require(process.env.PPTX_APPLY_THEME || "/root/.claude/skills/synced/e702c3d0-e0be-4bd4-91a1-b7fb6e6b4102_55712a2a-44a7-48e7-b712-bd90fba3fd2f/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
// 数字は scripts/deck_metrics.py が書き出す JSON から読む
const M = JSON.parse(require("fs").readFileSync(process.argv[3] || "deck_metrics.json", "utf8"));
const A = M.annual, Q = M.q3, RC = M.recovery, AG = M.agents, AC = M.actions, WL = M.winloss, LB = M.label;
const f1 = (v) => Number(v).toFixed(1);
const pct = (v) => `${Math.round(v * 100)}%`;
const ABBR = { "富士フイルム和光": "和光", "池田理化": "池田", "ナカライテスク": "ナカライ", "バイオテック・ラボ": "BTL" };
const ab = (n) => ABBR[n] || n;
const THEME = {
  name: "Pipeline Recovery",
  headFontFace: "Meiryo UI",
  bodyFontFace: "Meiryo UI",
  colors: {
    dk1: "1F2937", lt1: "FFFFFF", dk2: "1B2A4A", lt2: "EEF2F7",
    accent1: "1B2A4A", accent2: "0F8B7E", accent3: "D9822B", accent4: "C0392B",
    accent5: "8A94A6", accent6: "D6E4F0", hlink: "0F8B7E", folHlink: "8A94A6",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "下期パイプライン挽回計画";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "TITLE_ONLY",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.33, h: 0.75,
      fontSize: 24, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1.0, w: 12.33, h: 0.4,
      fontSize: 13, color: C.accent5, valign: "top", margin: 0 }, text: "" } },
    { text: { text: `出典: 案件Excel(9FCT / ${LB}版)、FY20-25商談総括、FY26新規リード、LST MOC FY26 Sep報告。施設名・個人名は非掲載`,
      options: { x: 0.5, y: 7.08, w: 10.5, h: 0.3, fontSize: 9, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.4, y: 7.08, w: 0.45, h: 0.3, fontSize: 9, color: C.accent5, align: "right" },
});

const card = (slide, o) => slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.08, fill: { color: o.fill || C.background2 },
  line: { color: o.fill || C.background2 }, objectName: o.name,
});
const label = (slide, text, x, y, w, opts = {}) => slide.addText(text, {
  x, y, w, h: opts.h || 0.32, fontSize: opts.size || 13, bold: opts.bold !== false, color: opts.color || C.text2,
  margin: 0, valign: "middle", isTextBox: true, objectName: opts.name, align: opts.align || "left",
});
const axisFont = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
  legendFontFace: "+mn-lt", titleFontFace: "+mn-lt" };

pres.addSection({ title: "役員報告" });

// ---------------- Slide 1: 現状 ----------------
{
  const s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "役員報告" });
  s.addText(`下期パイプラインは9F目標に${f1(A.short)}台不足、必要リードは${A.need9_deck}件から${A.need}件に拡大`, { placeholder: "title" });
  s.addText(`9F以降、今期案件${A.moved}件が来期へ移動。上期▲7台の挽回を含む9F目標${A.target}台の達成には、代理店の既存案件の引き上げが前提`, { placeholder: "body" });

  // KPI cards
  const kpis = [
    ["9F目標", `${A.target}台`, `1H実績${A.h1}台 + 下期${A.h2}台`, C.text2],
    [`着地理論値(${LB})`, `${f1(A.land)}台`, `9F時点 ${f1(A.land9)}台から ▲${f1(A.land9 - A.land)}台`, C.text2],
    ["不足", `${f1(A.short)}台`, `9F時点 ${f1(A.short9)}台から拡大`, C.accent4],
    ["必要リード(FCT相当)", `${A.need}件`, `9F時点 ${A.need9_deck}件 = 不足 ÷ モデル別勝率`, C.accent4],
  ];
  kpis.forEach(([t, v, sub, col], i) => {
    const x = 0.5 + i * 3.13, y = 1.55, w = 2.93, h = 1.3;
    card(s, { x, y, w, h, name: `kpi-${i}` });
    label(s, t, x + 0.2, y + 0.1, w - 0.4, { size: 12, color: C.accent5, name: `kpi-t-${i}` });
    s.addText(v, { x: x + 0.2, y: y + 0.38, w: w - 0.4, h: 0.55, fontSize: 30, bold: true, color: col, margin: 0,
      isTextBox: true, objectName: `kpi-v-${i}` });
    label(s, sub, x + 0.2, y + 0.92, w - 0.4, { size: 10, bold: false, color: C.text1, h: 0.3, name: `kpi-s-${i}` });
  });

  const top = 3.1, colH = 3.85;
  // Column 1: pipeline
  card(s, { x: 0.5, y: top, w: 4.0, h: colH, fill: C.background1, name: "col-pipe" });
  label(s, "パイプライン数(今期・確度>0)", 0.5, top, 4.0, { size: 14, name: "pipe-h" });
  s.addChart(pres.charts.BAR, [
    { name: "9F時点", labels: ["FCT", "Backup"], values: [A.fct9, A.bk9] },
    { name: `${LB}時点`, labels: ["FCT", "Backup"], values: [A.fct, A.bk] },
  ], {
    x: 0.5, y: top + 0.35, w: 4.0, h: 2.3, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HEX.accent6, HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11,
    dataLabelColor: HEX.dk1, showLegend: true, legendPos: "t", legendFontSize: 10, catAxisLabelFontSize: 11,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelColor: HEX.dk1, ...axisFont, objectName: "pipe-chart",
  });
  s.addText([
    { text: `Backupが${A.bk9}件→${A.bk}件に減少`, options: { bold: true, breakLine: true } },
    { text: `9F以降の動き: 来期へ移動${A.moved}件(うちBackup${A.moved_bk}件)、新規追加${A.new}件、LOST等${A.lost}件`, options: { breakLine: true } },
    { text: `下期FCT${A.h2_fct}件のうち${A.mar}件が3月計上、${A.mar_c}件は納期未定(C)` },
  ], { x: 0.65, y: top + 2.7, w: 3.7, h: 1.1, fontSize: 11, color: C.text1, margin: 0, valign: "top",
    paraSpaceAfter: 3, isTextBox: true, objectName: "pipe-note" });

  // Column 2: win/loss
  const x2 = 4.75;
  card(s, { x: x2, y: top, w: 3.6, h: colH, fill: C.background1, name: "col-win" });
  label(s, "勝敗(FY20-25実績)", x2, top, 3.6, { size: 14, name: "win-h" });
  s.addText([
    { text: pct(WL.rate), options: { fontSize: 30, bold: true, color: C.text2 } },
    { text: `  ${WL.win}勝${WL.loss}敗`, options: { fontSize: 12, color: C.text1 } },
  ], { x: x2, y: top + 0.38, w: 3.6, h: 0.55, margin: 0, isTextBox: true, objectName: "win-big" });
  s.addChart(pres.charts.BAR, [
    { name: "勝率", labels: WL.models.map((m) => m[0]), values: WL.models.map((m) => m[1]) },
  ], {
    x: x2, y: top + 0.95, w: 3.6, h: 1.55, barDir: "bar", chartColors: [HEX.accent1],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0%", dataLabelFontSize: 11,
    dataLabelColor: HEX.dk1, showLegend: false, valAxisHidden: true, valAxisMaxVal: 1, valAxisMinVal: 0,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLabelFontSize: 11,
    catAxisLabelColor: HEX.dk1, catAxisOrientation: "maxMin", barGapWidthPct: 50, ...axisFont, objectName: "win-chart",
  });
  card(s, { x: x2 + 0.15, y: top + 2.6, w: 3.3, h: 1.15, fill: C.background2, name: "win-callout" });
  s.addText([
    { text: "IDはデモの有無で勝率が倍違う", options: { bold: true, color: C.accent4, breakLine: true } },
    { text: "デモ有 90%(9勝1敗) / デモ無 46%", options: { breakLine: true } },
    { text: "1H: 大型LOST 3台、競合負け2台" },
  ], { x: x2 + 0.3, y: top + 2.66, w: 3.05, h: 1.05, fontSize: 11, color: C.text1, margin: 0, valign: "middle",
    paraSpaceAfter: 3, isTextBox: true, objectName: "win-note" });

  // Column 3: agents
  const x3 = 8.6;
  card(s, { x: x3, y: top, w: 4.23, h: colH, fill: C.background1, name: "col-agent" });
  label(s, "代理店別(一次店、下期)", x3, top, 4.23, { size: 14, name: "agent-h" });
  const hdr = ["一次店", "FCT", "Backup", "着地寄与", "期外移動", "勝率"].map((t) => ({
    text: t, options: { bold: true, fontSize: 9, color: C.background1, fill: { color: C.text2 }, align: t === "一次店" ? "left" : "center" },
  }));
  const rows = AG.rows.map((r) => [r[0], r[1], r[2], f1(r[3]), r[4], r[5] == null ? "-" : pct(r[5])]).map((r, i) => r.map((v, j) => ({ text: String(v), options: {
    align: j === 0 ? "left" : "center", fill: { color: i % 2 ? C.background2 : C.background1 },
    bold: j === 0 && i === 0, color: j === 4 && v >= 6 ? C.accent4 : C.text1 } })));
  s.addTable([hdr, ...rows], {
    x: x3, y: top + 0.4, w: 4.23, colW: [1.45, 0.42, 0.6, 0.64, 0.64, 0.48], rowH: 0.3, fontSize: 10,
    fontFace: THEME.bodyFontFace, border: { type: "none" }, margin: [0, 0.04, 0, 0.04], valign: "middle",
    objectName: "agent-table",
  });
  s.addText([
    { text: `上位5社で下期着地寄与の${pct(AG.share)}`, options: { bold: true, breakLine: true } },
    { text: `和光に集中: FCT${AG.wako_fct}件、期外移動${AG.wako_moved}件、期日超過・アクション未記入${AG.wako_open}件`, options: { breakLine: true } },
    { text: "着地寄与 = FCT×モデル勝率 + Backup×10%(台)" , options: { color: C.accent5, fontSize: 9 } },
  ], { x: x3 + 0.15, y: top + 2.35, w: 3.95, h: 1.4, fontSize: 11, color: C.text1, margin: 0, valign: "top",
    paraSpaceAfter: 3, isTextBox: true, objectName: "agent-note" });

  s.addNotes(`必要リード${A.need}件は、Sep報告と同じ式(着地理論値=成約+FCT×勝率+Backup×10%、必要リード=不足÷勝率)を${LB}版の案件Excelに当てた値。9F版で再計算するとスライドの21件とほぼ一致する。`);
}


// ---------------- Slide 2: Q3 ----------------
{
  const s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "役員報告" });
  s.addText(Q.max_all + Q.bk < Q.bgt
    ? `Q3 BGT${Q.bgt}台に対し見込み${f1(Q.land)}台、Q3全件でも届かずQ4前倒しが必須`
    : `Q3 BGT${Q.bgt}台に対し見込み${f1(Q.land)}台、FCT決め切りとQ4前倒しで${f1(Q.gap)}台を埋める`, { placeholder: "title" });
  s.addText(`Q3 FCT${Q.fct}件のうち${Q.fct_c}件が納期未定(C)。MTO品は受注から5週超のため、12月納品には11月上旬の受注確定がリミット`, { placeholder: "body" });

  const kpis = [
    ["Q3 BGT(10〜12月)", `${Q.bgt}台`, `年間9F ${A.target}台のうち下期${A.h2}台の一部`, C.text2],
    [`Q3 着地理論値(${LB})`, `${f1(Q.land)}台`, `9F時点 ${f1(Q.land9)}台から ▲${f1(Q.land9 - Q.land)}台`, C.text2],
    ["Q3 ギャップ", `${f1(Q.gap)}台`, `FCT・Backupを全件取っても${Q.max_all + Q.bk}台`, C.accent4],
    ["Q3 FCT", `${Q.fct}件`, `うち納期未定(C) ${Q.fct_c}件・MTO ${Q.mto}件`, C.accent4],
  ];
  kpis.forEach(([t, v, sub, col], i) => {
    const x = 0.5 + i * 3.13, y = 1.55, w = 2.93, h = 1.3;
    card(s, { x, y, w, h, name: `q3kpi-${i}` });
    label(s, t, x + 0.2, y + 0.1, w - 0.4, { size: 12, color: C.accent5, name: `q3kpi-t-${i}` });
    s.addText(v, { x: x + 0.2, y: y + 0.38, w: w - 0.4, h: 0.55, fontSize: 30, bold: true, color: col, margin: 0,
      isTextBox: true, objectName: `q3kpi-v-${i}` });
    label(s, sub, x + 0.2, y + 0.92, w - 0.4, { size: 10, bold: false, color: C.text1, h: 0.3, name: `q3kpi-s-${i}` });
  });

  const top = 3.1, colH = 3.85;
  // Column 1: Q3 pipeline 9F -> 10/5
  card(s, { x: 0.5, y: top, w: 3.7, h: colH, fill: C.background1, name: "q3-pipe" });
  label(s, "Q3パイプライン(確度>0)", 0.5, top, 3.7, { size: 14, name: "q3-pipe-h" });
  s.addChart(pres.charts.BAR, [
    { name: "9F時点", labels: ["FCT", "Backup"], values: [Q.fct9, Q.bk9] },
    { name: `${LB}時点`, labels: ["FCT", "Backup"], values: [Q.fct, Q.bk] },
  ], {
    x: 0.5, y: top + 0.35, w: 3.7, h: 2.3, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HEX.accent6, HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11,
    dataLabelColor: HEX.dk1, showLegend: true, legendPos: "t", legendFontSize: 10, catAxisLabelFontSize: 11,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelColor: HEX.dk1, ...axisFont, objectName: "q3-pipe-chart",
  });
  s.addText([
    { text: `9Fから FCT${Q.fct9}→${Q.fct}件・Backup${Q.bk9}→${Q.bk}件`, options: { bold: true, breakLine: true } },
    { text: `FCT${Q.fct}件: ` + Object.entries(Q.models).map(([k, v]) => `${k}${v}`).join("・"), options: { breakLine: true } },
    { text: "顧客: " + Object.entries(Q.genre).map(([k, v]) => `${{ AC: "アカデミア", BT: "バイオ企業", PHM: "製薬" }[k]}${v}`).join("・") },
  ], { x: 0.65, y: top + 2.7, w: 3.4, h: 1.1, fontSize: 11, color: C.text1, margin: 0, valign: "top",
    paraSpaceAfter: 3, isTextBox: true, objectName: "q3-pipe-note" });

  // Column 2: Q3 by agent
  const x2 = 4.45;
  card(s, { x: x2, y: top, w: 4.1, h: colH, fill: C.background1, name: "q3-agent" });
  label(s, "代理店別(一次店、Q3)", x2, top, 4.1, { size: 14, name: "q3-agent-h" });
  const hdr = ["一次店", "FCT", "Backup", "うち納期未定", "見込み"].map((t) => ({
    text: t, options: { bold: true, fontSize: 9, color: C.background1, fill: { color: C.text2 }, align: t === "一次店" ? "left" : "center" },
  }));
  const last = Q.agents.length - 1;
  const rows = Q.agents.map((r) => [r[0], r[1], r[2], r[3], f1(r[4])]).map((r, i) => r.map((v, j) => ({ text: String(v), options: {
    align: j === 0 ? "left" : "center", fill: { color: i === last ? C.background2 : C.background1 },
    bold: i === last || (j === 0 && i === 0), color: j === 3 && v >= 3 && i < last ? C.accent4 : C.text1 } })));
  s.addTable([hdr, ...rows], {
    x: x2, y: top + 0.4, w: 4.1, colW: [1.45, 0.5, 0.6, 0.85, 0.7], rowH: 0.3, fontSize: 10,
    fontFace: THEME.bodyFontFace, border: { type: "solid", pt: 0.5, color: HEX.lt2 }, margin: [0, 0.04, 0, 0.04],
    valign: "middle", objectName: "q3-agent-table",
  });
  s.addText([
    { text: `${Q.top2.map(ab).join("・")}でQ3 FCTの${pct(Q.top2_fct / Q.fct)}(${Q.top2_fct}件)`, options: { bold: true, breakLine: true } },
    { text: "納期を月決(A)・Q決(B)に固められるかが勝負", options: { breakLine: true } },
    { text: "見込み = FCT×モデル勝率 + Backup×10%", options: { color: C.accent5, fontSize: 9 } },
  ], { x: x2 + 0.15, y: top + 2.65, w: 3.8, h: 1.1, fontSize: 11, color: C.text1, margin: 0, valign: "top",
    paraSpaceAfter: 3, isTextBox: true, objectName: "q3-agent-note" });

  // Column 3: path to 18
  const x3 = 8.8;
  card(s, { x: x3, y: top, w: 4.03, h: colH, fill: C.background1, name: "q3-path" });
  label(s, `${Q.bgt}台への道筋(必要水準)`, x3, top, 4.03, { size: 14, name: "q3-path-h" });
  const pcats = ["現状見込み", "必要水準"];
  s.addChart(pres.charts.BAR, [
    { name: "見込み", labels: pcats, values: [Q.land, 0] },
    { name: "Q3 FCT決め切り", labels: pcats, values: [0, Q.close] },
    { name: "Backup昇格", labels: pcats, values: [0, Q.bk_up] },
    { name: "Q4から前倒し", labels: pcats, values: [0, Q.pf] },
  ], {
    x: x3, y: top + 0.35, w: 1.95, h: 3.0, barDir: "col", barGrouping: "stacked", barGapWidthPct: 35,
    chartColors: [HEX.accent5, HEX.accent1, HEX.accent3, HEX.accent2], showValue: true, dataLabelPosition: "ctr",
    dataLabelFormatCode: "0.0;;;", dataLabelFontSize: 11, dataLabelColor: HEX.lt1, showLegend: false,
    valAxisHidden: true, valAxisMaxVal: Q.bgt + 1, valAxisMinVal: 0, valGridLine: { style: "none" },
    catGridLine: { style: "none" }, catAxisLabelFontSize: 10, catAxisLabelColor: HEX.dk1, ...axisFont,
    objectName: "q3-path-chart",
  });
  const legend = [
    [`Q3 FCT決め切り ${Q.close}台`, `${Q.fct}件中${Q.close}件(${pct(Q.close_rate)})。モデル勝率なら約${Q.close_expected}件`, C.accent1],
    [`Backup昇格 ${Q.bk_up}台`, `Q3 Backup ${Q.bk}件から${Q.bk_up}件`, C.accent3],
    [`Q4から前倒し ${Q.pf}台`, `Q4 FCTのうち確度60%以上・納期A/Bの${Q.pf_cand}件から半分`, C.accent2],
  ];
  legend.forEach(([t, d, col], i) => {
    const ly = top + 0.55 + i * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x3 + 1.98, y: ly + 0.05, w: 0.18, h: 0.18, rectRadius: 0.03,
      fill: { color: col }, line: { color: col }, objectName: `q3-lg-${i}` });
    s.addText([
      { text: t, options: { bold: true, breakLine: true } },
      { text: d, options: { fontSize: 9, color: C.text1 } },
    ], { x: x3 + 2.22, y: ly, w: 1.76, h: 0.85, fontSize: 10, color: C.text2, margin: 0, valign: "top",
      isTextBox: true, objectName: `q3-lg-t-${i}` });
  });
  label(s, "前倒し分は年間の台数を増やさない点に注意", x3 + 0.15, top + 3.4, 3.8,
    { size: 10, bold: false, color: C.accent5, name: "q3-path-note" });

  s.addNotes(`Q3=2026年10〜12月。見込みは案件Excel(${LB}版)のQ3案件(10FCT計上分を含む)に、年間と同じ式(FCT×モデル勝率+Backup×10%)を当てた値。必要水準の内訳(決め切り${Q.close}・Backup${Q.bk_up}・前倒し${Q.pf})は${Q.bgt}台に届く組み合わせの一例で、予測ではない。MTO>Week+5週は案件Excelの記載。`);
}

// ---------------- Slide 3: アクションと挽回 ----------------
{
  const s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "役員報告" });
  s.addText(`11/7までにQ3受注を確定し年間${f1(RC.filled)}台を挽回、9F達成には棚卸の倍増が必要`, { placeholder: "title" });
  s.addText(`Q3は11/7までに受注確定(12月納品のリミット)。年間は厳しめの率(昇格${pct(RC.rate_up)}・引き戻し${pct(RC.rate_pb)}・商談化${pct(RC.rate_lead)})で試算し、11/7に判定`, { placeholder: "body" });

  // Waterfall (stacked bar with invisible base)
  const top = 1.55;
  card(s, { x: 0.5, y: top, w: 5.6, h: 5.4, fill: C.background1, name: "col-wf" });
  label(s, "挽回見込み(台)", 0.5, top, 5.6, { size: 14, name: "wf-h" });
  // ウォーターフォールの位置: 不足から各アクションを順に差し引く
  const steps = [RC.up, RC.pb, RC.tm, RC.lead];
  const bases = []; let cur = RC.short;
  steps.forEach((v) => { cur = Math.round((cur - v) * 10) / 10; bases.push(cur); });
  const cats = ["不足", "①Backup昇格", "②引き戻し", "③テレマ", "④⑤リード・展示会", "残ギャップ"];
  const wfY = top + 0.35, wfH = 2.6, L = { x: 0.02, y: 0.12, w: 0.96, h: 0.72 }, VMAX = Math.ceil(RC.short + 1);
  s.addChart(pres.charts.BAR, [
    { name: "base", labels: cats, values: [0, ...bases, 0] },
    { name: "不足", labels: cats, values: [RC.short, 0, 0, 0, 0, RC.gap] },
    { name: "アクション", labels: cats, values: [0, ...steps, 0] },
  ], {
    x: 0.5, y: wfY, w: 5.6, h: wfH, barDir: "col", barGrouping: "stacked", barGapWidthPct: 40,
    chartColors: [HEX.lt1, HEX.accent4, HEX.accent2], showLegend: false, valAxisHidden: true, catAxisHidden: true,
    valAxisMinVal: 0, valAxisMaxVal: VMAX, layout: L,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...axisFont, objectName: "wf-chart",
  });
  const px = 0.5 + L.x * 5.6, pw = L.w * 5.6, py = wfY + L.y * wfH, ph = L.h * wfH, step = pw / 6;
  const tops = [RC.short, RC.short, ...bases.slice(0, 3), RC.gap];
  const vals = [f1(RC.short), ...steps.map((v) => `+${f1(v)}`), f1(RC.gap)];
  const catTxt = ["不足", "①Backup\n昇格", "②引き戻し", "③テレマ", "④⑤リード\n・展示会", "残ギャップ"];
  tops.forEach((t, i) => {
    s.addText(vals[i], { x: px + i * step, y: py + (VMAX - t) / VMAX * ph - 0.3, w: step, h: 0.28,
      fontSize: 12, bold: true, align: "center", color: i === 0 || i === 5 ? C.accent4 : C.accent2, margin: 0,
      isTextBox: true, objectName: `wf-val-${i}` });
    s.addText(catTxt[i], { x: px + i * step, y: py + ph + 0.06, w: step, h: 0.42, fontSize: 10, align: "center",
      valign: "top", color: C.text1, margin: 0, isTextBox: true, objectName: `wf-cat-${i}` });
  });
  label(s, `代理店の既存案件 ${f1(RC.agent)}台 / インハウス施策 ${f1(RC.inhouse)}台`, 0.65, top + 3.2, 5.3,
    { size: 11, bold: false, color: C.text1, name: "wf-split" });

  // scenario cards
  const sc = [
    ["厳しめ(確認済みの率)", `着地 ${f1(RC.land_after)}台`, `9F目標に ▲${f1(RC.gap)}台`, C.accent3],
    [`ストレッチ(昇格${pct(RC.stretch_up)}/引戻${pct(RC.stretch_pb)})`, `着地 ${A.target}台`, "9F達成 = 上期▲7台も挽回", C.accent2],
  ];
  sc.forEach(([t, v, sub, col], i) => {
    const x = 0.65 + i * 2.7, y = top + 3.65, w = 2.55, h = 1.55;
    card(s, { x, y, w, h, fill: C.background2, name: `sc-${i}` });
    label(s, t, x + 0.15, y + 0.1, w - 0.3, { size: 10, color: C.accent5, h: 0.45, name: `sc-t-${i}` });
    s.addText(v, { x: x + 0.15, y: y + 0.55, w: w - 0.3, h: 0.5, fontSize: 22, bold: true, color: col, margin: 0,
      isTextBox: true, objectName: `sc-v-${i}` });
    label(s, sub, x + 0.15, y + 1.08, w - 0.3, { size: 11, bold: false, color: C.text1, name: `sc-s-${i}` });
  });

  // Gantt: 10/6 - 11/7
  const gx = 6.35, gw = 6.48;
  card(s, { x: gx, y: top, w: gw, h: 5.4, fill: C.background1, name: "col-gantt" });
  label(s, "スケジュール(リミット 11/7)", gx, top, gw, { size: 14, name: "gantt-h" });
  const lw = 0.95, cw = (gw - lw) / 5, hy = top + 0.4;
  const weeks = ["10/6〜", "10/13〜", "10/20〜", "10/27〜", "11/4〜7"];
  weeks.forEach((w, i) => {
    const last = i === 4;
    s.addShape(pres.shapes.RECTANGLE, { x: gx + lw + i * cw, y: hy, w: cw - 0.04, h: 0.32,
      fill: { color: last ? C.accent4 : C.text2 }, line: { color: last ? C.accent4 : C.text2 }, objectName: `wk-${i}` });
    s.addText(w, { x: gx + lw + i * cw, y: hy, w: cw - 0.04, h: 0.32, fontSize: 11, bold: true, align: "center",
      color: C.background1, margin: 0, isTextBox: true, objectName: `wk-t-${i}` });
  });
  const lanes = [
    ["Q3刈り取り", C.accent4, [
      [0, 1, `Q3 FCT${Q.fct}件 納期確定`],
      [1, 2, `Q4前倒し${Q.pf_cand}件 年内受注交渉`],
      [3, 2, "PO回収・12月納品確定"],
    ]],
    ["代理店", C.accent1, [
      [0, 1, "上位5社 棚卸会"],
      [1, 1, `Backup昇格判定 ${AC.bk_cases}件`],
      [2, 2, "ID/FP デモ・見積(11月期限)"],
      [1, 2, `来期移動・前倒し ${AC.pull_cases}件の引き戻し交渉`],
    ]],
    ["インハウス", C.accent2, [
      [0, 4, "テレマ 422名(10/9〜、S/A 44名を先行)"],
      [1, 1, "NGS EXPO リード配分"],
      [2, 1, `滞留リード${AC.prospecting}件 案件化`],
    ]],
    ["管理", C.accent5, [
      [3, 1, "10月末FCTレビュー"],
      [4, 1, "Q3・9F判定"],
    ]],
  ];
  let y = hy + 0.45;
  const barH = 0.4, gap = 0.05;
  lanes.forEach(([name, col, bars], li) => {
    // pack bars into rows so they do not overlap
    const rowsEnd = [];
    const placed = bars.map(([st, len, t]) => {
      let r = rowsEnd.findIndex((e) => e <= st);
      if (r < 0) { r = rowsEnd.length; rowsEnd.push(0); }
      rowsEnd[r] = st + len;
      return { st, len, t, r };
    });
    const laneH = rowsEnd.length * (barH + gap) + 0.06;
    s.addShape(pres.shapes.RECTANGLE, { x: gx + 0.08, y, w: gw - 0.16, h: laneH,
      fill: { color: li % 2 ? C.background1 : C.background2 }, line: { color: li % 2 ? C.background1 : C.background2 },
      objectName: `lane-${li}` });
    s.addText(name, { x: gx + 0.12, y, w: lw - 0.12, h: laneH, fontSize: 10, bold: true, color: C.text2,
      margin: 0, valign: "middle", isTextBox: true, objectName: `lane-t-${li}` });
    placed.forEach((b, bi) => {
      const bx = gx + lw + b.st * cw + 0.03, by = y + 0.05 + b.r * (barH + gap);
      const isLimit = b.st === 4 && b.len === 1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w: b.len * cw - 0.1, h: barH, rectRadius: 0.06,
        fill: { color: isLimit ? C.accent4 : col }, line: { color: isLimit ? C.accent4 : col },
        objectName: `bar-${li}-${bi}` });
      s.addText(b.t, { x: bx + 0.05, y: by, w: b.len * cw - 0.2, h: barH, fontSize: 10, color: C.background1,
        bold: true, margin: 0, valign: "middle", isTextBox: true, objectName: `bar-t-${li}-${bi}` });
    });
    y += laneH + 0.04;
  });

  // decisions requested
  const dy = y + 0.08, dh = top + 5.4 - 0.15 - dy;
  card(s, { x: gx + 0.15, y: dy, w: gw - 0.3, h: dh, fill: C.background2, name: "ask" });
  s.addText([
    { text: "役員へのお願い", options: { bold: true, color: C.accent4, breakLine: true } },
    { text: "主力代理店(和光・池田・レスター・ナカライ)の幹部へ、Q3案件の年内受注と年度内前倒しの要請(10月中)", options: { bullet: { indent: 12 }, breakLine: true } },
    { text: "前倒し案件向けの価格・在庫条件の決裁枠(11/7の判定で使う)", options: { bullet: { indent: 12 } } },
  ], { x: gx + 0.3, y: dy + 0.05, w: gw - 0.6, h: dh - 0.1, fontSize: 11, color: C.text1, margin: 0,
    valign: "middle", paraSpaceAfter: 3, isTextBox: true, objectName: "ask-text" });

  s.addNotes(`①=下期Backup${AC.bk_cases}件×${pct(RC.rate_up)}×(モデル勝率−10%)、②=来期へ移った${A.moved}件×${pct(RC.rate_pb)}×モデル勝率、③=テレマ422名×${pct(RC.rate_lead)}×FCT化32%(Sep報告slide4: 30/94)×平均勝率0.54、④⑤=リード保証148件と展示会2回×110リードを同じ率で換算。ストレッチは①②の率を同じ倍率で引き上げ、不足を埋める水準。`);
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

// ICIMC 2026 oral talk — Color-based life prediction of carbon steel
// Build: node build.js  -> ICIMC2026_Woo_colorimetry.pptx
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const path = require("path");

const IMG = (f) => path.join(__dirname, "img", f);
const OUT = process.argv[2] || "ICIMC2026_Woo_colorimetry.pptx";
// Optional look: `node build.js out.pptx academic` = the html-ppt academic-paper
// theme (cream paper, square corners, serif type). Default = original rust-lab.
const THEME = process.argv[3] || "rust-lab";
const ACAD = THEME === "academic";

// ---------- palette ----------
const C = {
  dark: "1F2226", ink: "2B2D31", muted: "6B6F76", faint: "A3A7AD",
  line: "D9DCE0", panel: "F3F4F6", white: "FFFFFF",
  rust: "B5501F", rustLight: "FBEDE4", rustSoft: "E8A87C",
  amber: "D4961C", teal: "2A8C8C",
  inc: "3A6EA5", prop: "3C8D5A", mat: "B23A2E",
  incL: "E6EEF7", propL: "E5F2EA", matL: "F8E4E2",
};
if (ACAD) Object.assign(C, {
  bg: "FDFCF8", ink: "0A0A0A", muted: "333333", faint: "707070", line: "DDDCD8",
  panel: "EEEDEA", rustLight: "F4E7DE", incL: "E4E9ED", propL: "E4EEE3", matL: "F3E3DE",
});
C.bg = C.bg || C.white;
const HF = "Cambria", BF = ACAD ? "Cambria" : "Calibri";
const W = 13.333, H = 7.5, MX = 0.6;

// ---------- data (manuscript Tables 2–3, Figs 3, 6b, 7, 9b) ----------
const CYC = [0, 1, 3, 5, 10, 15, 20, 25, 30];
const RA1 = [0.0, 1.8, 4.1, 3.8, 25.9, 47.7, 63.0, 63.8, 62.8];
const RA2 = [0.0, 3.6, 3.7, 7.4, 15.6, 51.3, 62.0, 62.9, 63.9];
const RP = [519.1, 852.2, 428.7, 473.4, 441.1, 301.3, 265.6, 205.0, 191.2];
const CR = [0.324, 0.193, 0.447, 0.406, 0.524, 0.838, 1.176, 1.414, 1.745];
const Ls = [34.0, 30.4, 32.2, 30.3, 28.2, 28.0, 27.2, 27.3, 26.4];
const As = [0.1, 0.2, -0.2, 0.1, 1.5, 4.4, 5.8, 5.8, 3.4];
const Bs = [-0.2, 2.1, 2.7, 2.3, 6.6, 12.2, 13.2, 13.9, 12.1];
const SZ = [0.02, 0.03, 0.04, 0.05, 0.06, 0.06, 0.08, 0.09, 0.10];
const RGB = [[79,79,80],[72,71,68],[76,75,71],[72,71,68],[72,65,56],[78,63,47],[78,60,44],[78,60,43],[73,60,44]];
const O_CYC = [0, 5, 15, 20, 30], O_VAL = [2.3, 20.0, 31.8, 36.0, 37.9];
const hex = (r, g, b) => [r, g, b].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
const stageOf = (c) => (c <= 5 ? 0 : c <= 15 ? 1 : 2);
const STAGE = [
  { name: "Incubation", span: "0–5", col: C.inc, light: C.incL, x0: -1, x1: 7.5 },
  { name: "Propagation", span: "10–15", col: C.prop, light: C.propL, x0: 7.5, x1: 17.5 },
  { name: "Maturation", span: "20–30", col: C.mat, light: C.matL, x0: 17.5, x1: 31 },
];

// ---------- helpers ----------
async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const pres = new pptxgen();
// academic-paper uses square corners everywhere
const RR = ACAD ? "rect" : "roundRect";
pres.layout = "LAYOUT_WIDE";
pres.title = "Color-Based Life Prediction of Carbon Steel";
pres.author = "Seonghun Woo";

let pageNo = 0;
function base(title) {
  const s = pres.addSlide();
  pageNo += 1;
  s.background = { color: C.bg };
  if (title) {
    s.addText(title, { x: MX, y: 0.35, w: W - 2 * MX, h: 0.75, fontFace: HF, fontSize: 30, bold: true, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
  }
  s.addText("ICIMC 2026  |  S. Woo et al., Gyeongsang National University", { x: MX, y: 7.05, w: 7, h: 0.3, fontFace: BF, fontSize: 9, color: C.faint, margin: 0, isTextBox: true });
  s.addText(String(pageNo), { x: W - MX - 1, y: 7.05, w: 1, h: 0.3, fontFace: BF, fontSize: 9, color: C.faint, align: "right", margin: 0, isTextBox: true });
  return s;
}
function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: BF, fontSize: 14, color: C.ink, margin: 0, isTextBox: true, valign: "top" }, o));
}
function card(s, x, y, w, h, fill) {
  // academic-paper: thin rule around every card, as in the theme's .card
  s.addShape(RR, { x, y, w, h, fill: { color: fill || C.panel }, line: ACAD ? { color: C.line, width: 0.75 } : { color: fill || C.panel }, rectRadius: 0.08 });
}
function pill(s, text, x, y, w, color, fs = 11) {
  s.addShape(RR, { x, y, w, h: 0.3, fill: { color }, line: { color }, rectRadius: 0.15 });
  txt(s, text, { x, y, w, h: 0.3, fontSize: fs, bold: true, color: C.white, align: "center", valign: "middle" });
}
function iconDot(s, data, x, y, d, bg) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const p = d * 0.26;
  s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

// Scatter (x = cycle) chart with a manual inner plot area so annotations and
// stage bands line up exactly with data coordinates.
function xyChart(s, box, series, o) {
  const L = Object.assign({ x: 0.1, y: 0.06, w: 0.86, h: 0.76 }, o.layout || {});
  const pa = { x: box.x + L.x * box.w, y: box.y + L.y * box.h, w: L.w * box.w, h: L.h * box.h };
  const xmin = o.xmin ?? 0, xmax = o.xmax ?? 30, ymin = o.ymin, ymax = o.ymax;
  const map = {
    px: (v) => pa.x + ((v - xmin) / (xmax - xmin)) * pa.w,
    py: (v) => pa.y + (1 - (v - ymin) / (ymax - ymin)) * pa.h,
    pa,
  };
  if (o.bands !== false) {
    STAGE.forEach((st) => {
      const x0 = map.px(Math.max(st.x0, xmin)), x1 = map.px(Math.min(st.x1, xmax));
      s.addShape(pres.shapes.RECTANGLE, { x: x0, y: pa.y, w: x1 - x0, h: pa.h, fill: { color: st.light, transparency: 20 }, line: { color: st.light, transparency: 100 } });
    });
  }
  const data = [{ name: "X", values: o.xs || CYC }].concat(series.map((sr) => ({ name: sr.name, values: sr.values })));
  s.addChart(pres.charts.SCATTER, data, {
    x: box.x, y: box.y, w: box.w, h: box.h,
    layout: L,
    chartColors: series.map((sr) => sr.color),
    lineSize: o.lineSize ?? 2.25,
    lineDataSymbol: "circle",
    lineDataSymbolSize: o.symbol ?? 8,
    lineDataSymbolLineSize: 1,
    lineDataSymbolLineColor: C.bg,
    showLegend: false,
    valAxisMinVal: ymin, valAxisMaxVal: ymax, valAxisMajorUnit: o.yunit,
    valAxisLabelFormatCode: o.yfmt || "General",
    catAxisMinVal: xmin, catAxisMaxVal: xmax, catAxisMajorUnit: 5, catAxisLabelFormatCode: "0",
    catAxisLabelColor: C.muted, valAxisLabelColor: C.muted,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
    catAxisLabelFontFace: BF, valAxisLabelFontFace: BF,
    valGridLine: { color: "E6E8EB", size: 0.75 },
    catGridLine: { style: "none" },
    catAxisLineColor: C.faint, valAxisLineShow: false,
    showValAxisTitle: !!o.ytitle, valAxisTitle: o.ytitle || "", valAxisTitleFontSize: 11, valAxisTitleColor: C.muted, valAxisTitleFontFace: BF,
    showCatAxisTitle: o.xtitle !== false, catAxisTitle: o.xtitle || "CCT cycle", catAxisTitleFontSize: 11, catAxisTitleColor: C.muted, catAxisTitleFontFace: BF,
  });
  return map;
}
function stageLegend(s, x, y, w) {
  const gw = (w - 0.2) / 3;
  STAGE.forEach((st, i) => pill(s, `${st.name}  ${st.span}`, x + i * (gw + 0.1), y, gw, st.col, 10.5));
}
function note(s, en, ko) { s.addNotes(`[EN]\n${en}\n\n[KO 참고]\n${ko}`); }

(async () => {
  const I = {
    camera: await icon(fa.FaCamera, C.white), palette: await icon(fa.FaPalette, C.white),
    micro: await icon(fa.FaMicroscope, C.white), xrd: await icon(fa.FaWaveSquare, C.white),
    bolt: await icon(fa.FaBolt, C.white), layers: await icon(fa.FaLayerGroup, C.white),
    arrow: await icon(fa.FaArrowRight, C.faint), arrowR: await icon(fa.FaArrowRight, C.rust),
    search: await icon(fa.FaSearch, C.white), flask: await icon(fa.FaFlask, C.white),
    chart: await icon(fa.FaChartLine, C.white), eye: await icon(fa.FaEye, C.white),
    tint: await icon(fa.FaTint, C.white), mountain: await icon(fa.FaMountain, C.white),
    question: await icon(fa.FaQuestion, C.white), check: await icon(fa.FaCheck, C.white),
  };

  // ===== 1. Title =====
  {
    const s = pres.addSlide(); pageNo += 1;
    s.background = { color: C.dark };
    txt(s, "ICIMC 2026  ·  Seoul Olympic Parktel  ·  November 3–7, 2026", { x: MX, y: 0.5, w: 9, h: 0.3, fontSize: 12, color: C.faint });
    txt(s, "Color-Based Life Prediction of Carbon Steel", { x: MX, y: 1.0, w: 12.1, h: 0.9, fontFace: HF, fontSize: 40, bold: true, color: C.white, valign: "middle" });
    txt(s, "Linking CIE L*a*b* colorimetry to corrosion rate", { x: MX, y: 2.0, w: 11, h: 0.5, fontFace: HF, fontSize: 22, italic: true, color: C.rustSoft });
    s.addText([
      { text: "Seonghun Woo", options: { bold: true, color: C.white } },
      { text: ", Junbeom Park, Beomsoo Kim, Jaeseung Kwon, Jung-Pil Noh, Jeonghyeon Yang*", options: { color: "D5D7DA" } },
    ], { x: MX, y: 2.85, w: 12, h: 0.35, fontFace: BF, fontSize: 15, margin: 0, isTextBox: true });
    txt(s, "Department of Mechanical System Engineering, Gyeongsang National University, Tongyeong, Korea   ·   * Corresponding author", { x: MX, y: 3.25, w: 12, h: 0.3, fontSize: 11.5, color: C.faint });
    // real specimen strip, cycle 0 → 30
    const n = 9, gap = 0.12, tw = (W - 2 * MX - gap * (n - 1)) / n;
    CYC.forEach((c, i) => {
      const x = MX + i * (tw + gap);
      s.addImage({ path: IMG(`photo_${c}.png`), x, y: 4.55, w: tw, h: tw });
      txt(s, `${c} cyc`, { x, y: 4.55 + tw + 0.08, w: tw, h: 0.3, fontSize: 11, color: C.faint, align: "center" });
    });
    txt(s, "Specimen surface, 0 → 30 CCT cycles (240 h)", { x: MX, y: 4.15, w: 8, h: 0.3, fontSize: 11, italic: true, color: C.faint });
    note(s,
      "Hello everyone. I am Seonghun Woo from Gyeongsang National University. My talk asks one question. Can the color of rust tell us how fast the steel underneath is corroding? The strip at the bottom is the actual specimen surface from cycle 0 to cycle 30. We tested this question on bare carbon steel through 30 cycles of accelerated corrosion.",
      "안녕하십니까. 경상국립대학교 우성훈입니다. 오늘 발표는 질문 하나에서 출발합니다. 녹의 색으로 그 아래 강재의 부식 속도를 알 수 있을까요? 아래 사진은 0사이클부터 30사이클까지 실제 시편 표면입니다. 무도장 탄소강을 30사이클 촉진부식 시험에 넣고 이 질문을 검증했습니다.");
  }

  // ===== 2. Outline =====
  {
    const s = base("Outline");
    const items = [
      ["01", "Why rust color?", I.question],
      ["02", "How we measured", I.flask],
      ["03", "What changed", I.chart],
      ["04", "What color tells us", I.eye],
    ];
    const cw = 2.7, gap = (W - 2 * MX - 4 * cw) / 3;
    items.forEach(([n, t, ic], i) => {
      const x = MX + i * (cw + gap), y = 2.4;
      card(s, x, y, cw, 2.6, i === 3 ? C.rustLight : C.panel);
      iconDot(s, ic, x + 0.35, y + 0.4, 0.8, i === 3 ? C.rust : C.ink);
      txt(s, n, { x: x + 0.35, y: y + 1.4, w: 2, h: 0.4, fontFace: HF, fontSize: 20, bold: true, color: C.rust });
      txt(s, t, { x: x + 0.35, y: y + 1.85, w: cw - 0.5, h: 0.5, fontSize: 18, bold: true });
      if (i < 3) s.addImage({ data: I.arrow, x: x + cw + gap / 2 - 0.12, y: y + 1.18, w: 0.24, h: 0.24 });
    });
    note(s,
      "The talk has four parts. First, why rust color matters. Second, how we measured it. Third, what changed on the surface, measured six different ways. And last, what the color signal actually tells us, with its limits.",
      "발표는 네 부분입니다. 먼저 녹의 색이 왜 중요한지, 다음으로 측정 방법, 이어서 여섯 가지 방법으로 본 표면 변화, 마지막으로 색 신호가 실제로 알려 주는 것과 그 한계를 말씀드리겠습니다.");
  }

  // ===== 3. Background =====
  {
    const s = base("Rust color carries phase information");
    // stat
    txt(s, "3–4%", { x: MX, y: 1.5, w: 4.5, h: 1.2, fontFace: HF, fontSize: 72, bold: true, color: C.rust, valign: "middle" });
    txt(s, "of global GDP lost to corrosion", { x: MX, y: 2.75, w: 4.8, h: 0.4, fontSize: 16, color: C.muted });
    // flow: rust layer -> lab
    card(s, MX, 3.5, 5.3, 1.25);
    iconDot(s, I.micro, MX + 0.25, 3.72, 0.8, C.ink);
    txt(s, [
      { text: "Rust phases decide what happens next", options: { bold: true, breakLine: true } },
      { text: "Today: XRD / SEM — specimen goes to the lab", options: { color: C.muted, fontSize: 13 } },
    ], { x: MX + 1.25, y: 3.72, w: 3.9, h: 0.85, fontSize: 15, valign: "middle" });
    // phase swatches
    const ph = [
      ["γ-FeOOH", "Lepidocrocite", "Orange-yellow", "D38A2E"],
      ["α-FeOOH", "Goethite", "Reddish-brown", "7A3E1D"],
      ["Fe₃O₄", "Magnetite", "Black", "1B1B1B"],
    ];
    ph.forEach(([f, n, c, col], i) => {
      const x = 7.0 + i * 1.95;
      s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 1.55, w: 1.35, h: 1.35, fill: { color: col }, line: { color: C.line, width: 1 } });
      txt(s, f, { x, y: 3.05, w: 1.75, h: 0.4, fontFace: HF, fontSize: 17, bold: true, align: "center" });
      txt(s, `${n}\n${c}`, { x, y: 3.45, w: 1.75, h: 0.6, fontSize: 12, color: C.muted, align: "center" });
    });
    txt(s, "Swatch colors are illustrative.", { x: 7.0, y: 4.15, w: 5.8, h: 0.25, fontSize: 10, italic: true, color: C.faint, align: "center" });
    // research question
    s.addShape(RR, { x: MX, y: 5.2, w: W - 2 * MX, h: 1.4, fill: { color: C.dark }, line: { color: C.dark }, rectRadius: 0.08 });
    iconDot(s, I.camera, MX + 0.35, 5.5, 0.8, C.rust);
    txt(s, "RESEARCH QUESTION", { x: MX + 1.45, y: 5.45, w: 5, h: 0.3, fontSize: 11, bold: true, color: C.rustSoft, charSpacing: 2 });
    txt(s, "Can one color value from a photograph report the corrosion rate of bare carbon steel?", { x: MX + 1.45, y: 5.78, w: 10.3, h: 0.6, fontFace: HF, fontSize: 21, italic: true, color: C.white, valign: "middle" });
    note(s,
      "Corrosion costs an estimated 3 to 4 percent of global GDP. The state of the rust layer matters because that layer controls what happens next. The phases in rust differ in color. Lepidocrocite is orange-yellow. Goethite is darker and more reddish-brown. Magnetite is black. Today we identify these phases with XRD or SEM, and both need a specimen taken to the lab. If color carries phase information, a camera could take over part of that work. That is what we tested.",
      "부식 비용은 세계 GDP의 3~4%로 추산됩니다. 녹층의 상태가 중요한 이유는 이 층이 이후의 부식 거동을 좌우하기 때문입니다. 녹을 이루는 상들은 색이 다릅니다. 레피도크로사이트는 주황빛 노란색, 괴타이트는 더 어두운 적갈색, 마그네타이트는 검은색입니다. 지금은 XRD나 SEM으로 이 상들을 확인하는데, 둘 다 시편을 실험실로 가져가야 합니다. 색에 상 정보가 담겨 있다면 이 작업의 일부를 카메라가 맡을 수 있습니다. 이 가능성을 검증했습니다.");
  }

  // ===== 4. Methods overview =====
  {
    const s = base("Six measurements on one specimen set");
    // optical group
    card(s, MX, 1.5, 4.6, 4.3, C.rustLight);
    txt(s, "OPTICAL  ·  NON-DESTRUCTIVE", { x: MX + 0.35, y: 1.72, w: 4, h: 0.3, fontSize: 11, bold: true, color: C.rust, charSpacing: 1 });
    [[I.camera, "Rust area", "Image segmentation"], [I.palette, "CIE L*a*b*", "Color of rust pixels"]].forEach(([ic, t, d], i) => {
      const y = 2.35 + i * 1.6;
      iconDot(s, ic, MX + 0.35, y, 1.0, C.rust);
      txt(s, t, { x: MX + 1.6, y: y + 0.08, w: 2.9, h: 0.45, fontFace: HF, fontSize: 20, bold: true });
      txt(s, d, { x: MX + 1.6, y: y + 0.55, w: 2.9, h: 0.35, fontSize: 14, color: C.muted });
    });
    // arrow
    s.addImage({ data: I.arrowR, x: 5.55, y: 3.45, w: 0.5, h: 0.5 });
    txt(s, "checked\nagainst", { x: 5.25, y: 4.05, w: 1.1, h: 0.55, fontSize: 12, color: C.muted, align: "center" });
    // reference group
    card(s, 6.55, 1.5, W - MX - 6.55, 4.3);
    txt(s, "REFERENCE  ·  LAB-BASED", { x: 6.9, y: 1.72, w: 4, h: 0.3, fontSize: 11, bold: true, color: C.muted, charSpacing: 1 });
    const ref = [[I.micro, "SEM / EDS", "Morphology, O content"], [I.xrd, "XRD", "Crystalline phases"], [I.bolt, "LPR", "Rp, corrosion rate"], [I.layers, "Confocal", "Surface height Sz"]];
    ref.forEach(([ic, t, d], i) => {
      const x = 6.9 + (i % 2) * 2.95, y = 2.35 + Math.floor(i / 2) * 1.6;
      iconDot(s, ic, x, y, 0.85, C.ink);
      txt(s, t, { x: x + 1.0, y: y + 0.02, w: 1.9, h: 0.4, fontFace: HF, fontSize: 17, bold: true });
      txt(s, d, { x: x + 1.0, y: y + 0.44, w: 1.9, h: 0.5, fontSize: 12.5, color: C.muted });
    });
    // sampling strip
    txt(s, [
      { text: "9", options: { bold: true, color: C.rust, fontSize: 20 } }, { text: " sampling points   ·   ", options: { color: C.muted } },
      { text: "0–30", options: { bold: true, color: C.rust, fontSize: 20 } }, { text: " cycles (240 h)   ·   ", options: { color: C.muted } },
      { text: "2", options: { bold: true, color: C.rust, fontSize: 20 } }, { text: " specimens per point", options: { color: C.muted } },
    ], { x: MX, y: 6.15, w: W - 2 * MX, h: 0.5, fontSize: 15, align: "center", valign: "middle" });
    note(s,
      "We ran six measurements on the same set of specimens. Two are optical and non-destructive. Rust area comes from image segmentation, and color comes from the CIE L*a*b* values of rust pixels. The other four serve as references. SEM with EDS gives morphology and oxygen content. XRD identifies phases. Linear polarization gives Rp and corrosion rate. Confocal microscopy gives surface height. The goal was to check the optical data against each reference.",
      "같은 시편 세트에 여섯 가지 측정을 적용했습니다. 두 가지는 광학적 비파괴 측정입니다. 녹 면적은 이미지 분할로, 색은 녹 픽셀의 CIE L*a*b* 값으로 구했습니다. 나머지 네 가지는 기준 데이터입니다. SEM/EDS로 형상과 산소 함량을, XRD로 상을 확인했습니다. 선형분극으로 Rp와 부식속도를, 공초점 현미경으로 표면 높이를 측정했습니다. 광학 데이터를 각 기준 데이터와 대조하는 것이 목표였습니다.");
  }

  // ===== 5. Specimen & CCT =====
  {
    const s = base("Specimens and accelerated exposure");
    // doughnut of one cycle
    s.addChart(pres.charts.DOUGHNUT, [{ name: "CCT", labels: ["Fog", "Dry", "Humid"], values: [2, 4, 2] }], {
      x: MX, y: 1.35, w: 3.6, h: 3.6, holeSize: 62, chartColors: [C.inc, C.amber, C.prop],
      showLegend: false, showValue: false, showPercent: false, showLabel: false, dataBorder: { pt: 2, color: "FFFFFF" },
    });
    txt(s, "8 h", { x: MX + 0.8, y: 2.65, w: 2.0, h: 0.6, fontFace: HF, fontSize: 32, bold: true, align: "center", valign: "middle" });
    txt(s, "per cycle", { x: MX + 0.8, y: 3.22, w: 2.0, h: 0.3, fontSize: 12, color: C.muted, align: "center" });
    const legend = [[C.inc, "Fog  2 h", "35 °C  ·  UPW spray"], [C.amber, "Dry  4 h", "60 °C  ·  RH < 30%"], [C.prop, "Humid  2 h", "50 °C  ·  RH > 95%"]];
    legend.forEach(([col, t, d], i) => {
      const y = 1.75 + i * 0.95;
      s.addShape(pres.shapes.OVAL, { x: 4.45, y: y + 0.07, w: 0.26, h: 0.26, fill: { color: col }, line: { color: col } });
      txt(s, t, { x: 4.85, y, w: 2.2, h: 0.4, fontSize: 17, bold: true });
      txt(s, d, { x: 4.85, y: y + 0.4, w: 2.4, h: 0.3, fontSize: 12.5, color: C.muted });
    });
    txt(s, "Q-FOG CCT-600  ·  modified ISO 14993", { x: MX, y: 4.95, w: 6.5, h: 0.3, fontSize: 12, color: C.muted, italic: true });
    // specimen card
    card(s, 7.55, 1.45, W - MX - 7.55, 1.75);
    txt(s, "SPECIMEN", { x: 7.85, y: 1.62, w: 3, h: 0.3, fontSize: 11, bold: true, color: C.rust, charSpacing: 1 });
    txt(s, [
      { text: "Cold-rolled carbon steel", options: { bold: true, breakLine: true, fontSize: 16 } },
      { text: "30 × 70 × 1 mm  ·  0.12 wt.% C", options: { breakLine: true } },
      { text: "SiC #220 ground  ·  ethanol ultrasonic clean", options: {} },
    ], { x: 7.85, y: 1.98, w: 4.8, h: 1.5, fontSize: 13.5, color: C.ink, paraSpaceAfter: 4 });
    // UPW callout
    card(s, 7.55, 3.45, W - MX - 7.55, 1.4, C.rustLight);
    iconDot(s, I.tint, 7.8, 3.72, 0.8, C.rust);
    txt(s, [
      { text: "Ultrapure water, not 5% NaCl", options: { bold: true, breakLine: true } },
      { text: "Salt residue disturbed imaging → chloride-free atmospheric test", options: { color: C.muted, fontSize: 12.5 } },
    ], { x: 8.8, y: 3.65, w: 3.85, h: 1.0, fontSize: 15, valign: "middle" });
    // sampling timeline
    txt(s, "Sampling points (cycles)", { x: MX, y: 5.55, w: 5, h: 0.3, fontSize: 12, bold: true, color: C.muted });
    const tx0 = MX + 0.2, tx1 = W - MX - 0.2, ty = 6.2;
    s.addShape(pres.shapes.LINE, { x: tx0, y: ty, w: tx1 - tx0, h: 0, line: { color: C.line, width: 2 } });
    CYC.forEach((c) => {
      const x = tx0 + (c / 30) * (tx1 - tx0), col = STAGE[stageOf(c)].col;
      s.addShape(pres.shapes.OVAL, { x: x - 0.12, y: ty - 0.12, w: 0.24, h: 0.24, fill: { color: col }, line: { color: C.white, width: 1.5 } });
      txt(s, String(c), { x: x - 0.3, y: ty + 0.2, w: 0.6, h: 0.3, fontSize: 12, align: "center", color: C.ink });
    });
    txt(s, "30 cycles = 240 h  ·  2 specimens per point  ·  18 specimens", { x: W - MX - 6, y: 5.55, w: 6, h: 0.3, fontSize: 12, color: C.muted, align: "right" });
    note(s,
      "The specimens were cold-rolled carbon steel, 30 by 70 by 1 millimeters, with 0.12 percent carbon. We ground one face with 220-grit paper. One CCT cycle lasts eight hours: two hours of fog at 35 degrees, four hours of drying at 60 degrees, then two hours of humidity above 95 percent. We replaced the salt fog with ultrapure water, because salt deposits interfered with imaging. So this test reflects chloride-free atmospheric corrosion, not a marine one. Specimens came out at nine points up to 30 cycles, in duplicate.",
      "시편은 냉연 탄소강으로 크기 30×70×1 mm, 탄소 0.12%입니다. 한 면을 220번 연마지로 연마했습니다. CCT 한 사이클은 8시간으로, 35 °C 분무 2시간, 60 °C 건조 4시간, 습도 95% 이상 습윤 2시간 순서입니다. 염 잔류물이 이미지 분석을 방해해서 염수 대신 초순수를 썼습니다. 따라서 이 시험은 해양 환경이 아니라 염화물 없는 대기부식을 반영합니다. 시편은 30사이클까지 아홉 시점에서 두 개씩 회수했습니다.");
  }

  // ===== 6. Image pipeline =====
  {
    const s = base("Color is read only where rust is");
    const steps = [
      { t: "Photograph", d: "Diffuse, fixed lighting\nCentral 75% of image", img: "photo_15.png" },
      { t: "HSV mask", d: "H 10–60°, S > 24%, V > 16%\nOpen/close, drop small blobs", img: "mask_15.png" },
      { t: "Common region", d: "Pixels rust in ≥ 5 of 9\nsampling intervals", img: "mask_30.png" },
      { t: "L*, a*, b*", d: "Mean over that fixed\npixel set, every cycle", color: hex(...RGB[5]) },
    ];
    const cw = 2.55, gap = (W - 2 * MX - 4 * cw) / 3;
    steps.forEach((st, i) => {
      const x = MX + i * (cw + gap), y = 1.45;
      if (st.img) s.addImage({ path: IMG(st.img), x, y, w: cw, h: cw });
      else {
        s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: cw, fill: { color: st.color }, line: { color: C.line } });
        txt(s, "L* 28.0\na* 4.4\nb* 12.2", { x, y, w: cw, h: cw, fontFace: HF, fontSize: 22, bold: true, color: C.white, align: "center", valign: "middle" });
      }
      s.addShape(pres.shapes.OVAL, { x: x - 0.18, y: y - 0.18, w: 0.5, h: 0.5, fill: { color: C.rust }, line: { color: C.white, width: 2 } });
      txt(s, String(i + 1), { x: x - 0.18, y: y - 0.18, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
      txt(s, st.t, { x, y: y + cw + 0.15, w: cw, h: 0.4, fontFace: HF, fontSize: 18, bold: true });
      txt(s, st.d, { x, y: y + cw + 0.58, w: cw + 0.2, h: 0.6, fontSize: 12.5, color: C.muted });
      if (i < 3) s.addImage({ data: I.arrow, x: x + cw + gap / 2 - 0.13, y: y + cw / 2 - 0.13, w: 0.26, h: 0.26 });
    });
    txt(s, "Example: cycle 15. The same pixel set is tracked across all cycles, so a color change is not a coverage change.", { x: MX, y: 5.35, w: W - 2 * MX, h: 0.3, fontSize: 12, italic: true, color: C.muted });
    // electrochem chips
    card(s, MX, 5.85, W - 2 * MX, 0.9);
    iconDot(s, I.bolt, MX + 0.2, 5.97, 0.66, C.ink);
    txt(s, "Electrochemistry", { x: MX + 1.05, y: 5.85, w: 2.2, h: 0.9, fontSize: 15, bold: true, valign: "middle" });
    const chips = ["LPR  ±25 mV vs OCP", "0.167 mV/s", "3.5 wt.% NaCl, 1 cm²", "Ag/AgCl  ·  Pt mesh", "8 scans / point"];
    let cx = 3.35;
    chips.forEach((c) => {
      const w = 0.2 + c.length * 0.085;
      s.addShape(RR, { x: cx, y: 6.13, w, h: 0.36, fill: { color: C.white }, line: { color: C.line }, rectRadius: 0.18 });
      txt(s, c, { x: cx, y: 6.13, w, h: 0.36, fontSize: 12, align: "center", valign: "middle" });
      cx += w + 0.15;
    });
    note(s,
      "Color is only read where rust is. We photographed each specimen under diffuse lighting and used the central 75 percent of the image. Rust pixels were masked in HSV space, with hue between 10 and 60 degrees, and small noise blobs were removed. Then we defined a common rust region: the pixels classified as rust in at least five of the nine intervals. L*, a* and b* were averaged over that fixed pixel set at every cycle. This separates a change in rust color from a change in rust coverage, and bare steel does not dilute the signal. For electrochemistry we used 3.5 percent NaCl as a test electrolyte, because ultrapure water conducts too poorly for LPR.",
      "색은 녹이 있는 곳에서만 읽습니다. 확산 조명 아래에서 시편을 촬영하고 이미지 중앙 75% 영역을 사용했습니다. HSV 공간에서 색상각 10~60° 조건으로 녹 픽셀을 골라내고 작은 노이즈 영역은 제거했습니다. 이어서 아홉 시점 중 다섯 번 이상 녹으로 분류된 픽셀을 공통 녹 영역으로 정의했습니다. 모든 사이클에서 이 고정된 픽셀 집합의 L*, a*, b*를 평균했습니다. 이렇게 하면 녹 색의 변화와 녹 면적의 변화를 분리할 수 있고, 맨 강재가 신호를 희석하지 않습니다. 전기화학 측정에는 3.5% NaCl을 시험 전해질로 썼습니다. 초순수는 전도도가 너무 낮아 LPR 측정이 어렵기 때문입니다.");
  }

  // ===== 7. Rust area =====
  {
    const s = base("Coverage rose fastest between cycles 10 and 15");
    // masks 2x2
    const ms = [[5, 5.6], [10, 20.8], [15, 49.5], [30, 63.4]];
    const tw = 1.9, g = 0.2;
    ms.forEach(([c, v], i) => {
      const x = MX + (i % 2) * (tw + g), y = 1.4 + Math.floor(i / 2) * (tw + 0.55);
      s.addImage({ path: IMG(`mask_${c}.png`), x, y, w: tw, h: tw });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: tw, h: tw, fill: { color: C.white, transparency: 100 }, line: { color: C.line } });
      txt(s, [{ text: `${c} cyc  `, options: { color: C.muted } }, { text: `${v}%`, options: { bold: true, color: STAGE[stageOf(c)].col } }], { x, y: y + tw + 0.05, w: tw, h: 0.35, fontSize: 13, align: "center" });
    });
    txt(s, "Segmented rust mask; value = mean of 2 specimens", { x: MX, y: 6.3, w: 4.2, h: 0.3, fontSize: 10.5, italic: true, color: C.faint });
    // chart
    const box = { x: 4.95, y: 1.3, w: 7.8, h: 4.3 };
    const m = xyChart(s, box, [{ name: "Specimen 1", values: RA1, color: C.rust }, { name: "Specimen 2", values: RA2, color: C.ink }], { ymin: 0, ymax: 70, yunit: 10, ytitle: "Rust area (%)" });
    // legend
    [[C.rust, "Specimen 1"], [C.ink, "Specimen 2"]].forEach(([col, t], i) => {
      const x = m.pa.x + 0.2 + i * 1.6;
      s.addShape(pres.shapes.LINE, { x, y: m.pa.y + 0.25, w: 0.35, h: 0, line: { color: col, width: 2.5 } });
      txt(s, t, { x: x + 0.45, y: m.pa.y + 0.12, w: 1.2, h: 0.26, fontSize: 11, color: C.ink });
    });
    stageLegend(s, m.pa.x, 5.75, m.pa.w);
    // stat row
    const st = [["< 6%", "cycles 0–5", C.inc], ["49.5%", "cycle 15", C.prop], ["≈ 63%", "cycles 20–30", C.mat]];
    st.forEach(([v, l, col], i) => {
      const x = m.pa.x + i * (m.pa.w / 3);
      txt(s, [{ text: v + "  ", options: { fontFace: HF, fontSize: 22, bold: true, color: col } }, { text: l, options: { fontSize: 12.5, color: C.muted } }], { x, y: 6.2, w: m.pa.w / 3, h: 0.45, valign: "middle" });
    });
    note(s,
      "On the left are rust masks at four cycles. The spots at cycle 5 are isolated. By cycle 15 they have joined into a network. On the right is the coverage curve for both specimens. Through cycle 5 coverage stayed under 6 percent. Between cycles 10 and 15 it jumped to about 50 percent. After cycle 20 it flattened near 63 percent. The two specimens agree within 2 percent from cycle 20 on. We used these breaks to define three stages: incubation, propagation and maturation. The shaded bands on every chart in this talk use the same boundaries.",
      "왼쪽은 네 시점의 녹 마스크입니다. 5사이클의 녹은 고립된 점입니다. 15사이클에는 이 점들이 연결되어 망을 이룹니다. 오른쪽은 두 시편의 면적률 곡선입니다. 5사이클까지 면적률은 6% 미만이었습니다. 10~15사이클 사이에 약 50%까지 급증했습니다. 20사이클 이후에는 63% 부근에서 평탄해졌고, 두 시편의 차이는 2% 이내였습니다. 이 변곡점을 기준으로 잠복기, 전파기, 성숙기의 세 단계를 정의했습니다. 이후 모든 그래프의 배경 색 구간은 같은 경계를 사용합니다.");
  }

  // ===== 8. Rp & CR =====
  {
    const s = base("Polarization resistance fell as the layer grew");
    const bw = (W - 2 * MX - 0.4) / 2;
    const m1 = xyChart(s, { x: MX, y: 1.3, w: bw, h: 3.9 }, [{ name: "Rp", values: RP, color: C.ink }], { ymin: 0, ymax: 1000, yunit: 200, ytitle: "Rp (Ω·cm²)", layout: { x: 0.14, w: 0.82 } });
    const m2 = xyChart(s, { x: MX + bw + 0.4, y: 1.3, w: bw, h: 3.9 }, [{ name: "CR", values: CR, color: C.rust }], { ymin: 0, ymax: 2, yunit: 0.5, ytitle: "Corrosion rate (mmpy)", layout: { x: 0.14, w: 0.82 } });
    // annotations
    const ann = (m, xv, yv, t, col, dx = 0.12, dy = -0.42) => txt(s, t, { x: m.px(xv) + dx, y: m.py(yv) + dy, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: col });
    ann(m1, 1, 852.2, "852 (cycle 1)", C.ink);
    ann(m1, 10, 441.1, "441", C.prop, -0.2, -0.4);
    ann(m1, 15, 301.3, "301", C.prop, 0.05, -0.4);
    ann(m1, 30, 191.2, "191", C.mat, -0.35, -0.42);
    ann(m2, 3, 0.447, "0.45", C.inc, -0.25, -0.42);
    ann(m2, 30, 1.745, "1.75", C.mat, -0.55, -0.1);
    // stat row
    const st = [
      ["Cycle 1", "Transient Rp peak — thin initial oxide film", C.panel, C.ink],
      ["441 → 301", "Ω·cm², cycles 10 → 15: steepest Rp drop", C.propL, C.prop],
      ["3.9×", "Corrosion rate, 0.45 → 1.75 mmpy (cycle 3 → 30)", C.rustLight, C.rust],
    ];
    const cw = (W - 2 * MX - 0.4) / 3;
    st.forEach(([v, l, bg, col], i) => {
      const x = MX + i * (cw + 0.2);
      card(s, x, 5.55, cw, 1.2, bg);
      txt(s, v, { x: x + 0.3, y: 5.65, w: cw - 0.5, h: 0.55, fontFace: HF, fontSize: 26, bold: true, color: col, valign: "middle" });
      txt(s, l, { x: x + 0.3, y: 6.2, w: cw - 0.5, h: 0.45, fontSize: 12.5, color: C.muted });
    });
    note(s,
      "Rp and corrosion rate move in opposite directions. At cycle 1 Rp briefly rose to 852 ohm square centimeters. We attribute this to a thin initial oxide film. After that Rp fell. The steepest drop came between cycles 10 and 15, from 441 to 301. That is the same window where coverage jumped. Corrosion rate rose from 0.45 millimeters per year at cycle 3 to 1.75 at cycle 30, close to four times. After cycle 20 both curves change more slowly. This fits diffusion through a thicker layer.",
      "Rp와 부식속도는 반대 방향으로 움직입니다. 1사이클에서 Rp가 852 Ω·cm²로 일시적으로 올랐습니다. 얇은 초기 산화막 때문으로 봅니다. 그 뒤 Rp는 감소했습니다. 가장 가파른 감소는 10~15사이클 구간으로, 441에서 301로 떨어졌습니다. 면적률이 급증한 구간과 같습니다. 부식속도는 3사이클 0.45 mmpy에서 30사이클 1.75 mmpy로, 약 4배 올랐습니다. 20사이클 이후에는 두 곡선 모두 완만해집니다. 두꺼워진 층을 통한 확산 지배와 부합합니다.");
  }

  // ===== 9. SEM/EDS =====
  {
    const s = base("Surface oxygen climbed to 37.9 wt.%");
    const cs = [0, 5, 15, 30], tw = 1.6, th = tw * 301 / 401, g = 0.12;
    txt(s, "SEM ×500", { x: MX + 0.75, y: 1.3, w: tw, h: 0.3, fontSize: 12, bold: true, color: C.muted, align: "center" });
    txt(s, "O map", { x: MX + 0.75 + tw + g, y: 1.3, w: tw, h: 0.3, fontSize: 12, bold: true, color: C.muted, align: "center" });
    cs.forEach((c, i) => {
      const y = 1.65 + i * (th + g);
      txt(s, `${c}\ncyc`, { x: MX, y, w: 0.65, h: th, fontSize: 12, bold: true, color: STAGE[stageOf(c)].col, align: "center", valign: "middle" });
      s.addImage({ path: IMG(`sem_${c}.jpg`), x: MX + 0.75, y, w: tw, h: th });
      s.addImage({ path: IMG(`omap_${c}.jpg`), x: MX + 0.75 + tw + g, y, w: tw, h: th });
    });
    // O content chart: stacked by stage so each column takes its stage color
    const cats = O_CYC.map((c) => `${c} cyc`);
    const cx = 5.55;
    txt(s, "EDS oxygen content (map sum, wt.%)", { x: cx, y: 1.3, w: 6, h: 0.3, fontSize: 13, bold: true, color: C.ink });
    s.addChart(pres.charts.BAR, [{ name: "O wt.%", labels: cats, values: O_VAL }], {
      x: cx, y: 1.6, w: W - MX - cx, h: 3.85, barDir: "col", barGapWidthPct: 55,
      chartColors: O_CYC.map((c) => STAGE[stageOf(c)].col), showLegend: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelColor: C.ink, dataLabelFontSize: 14, dataLabelFontBold: true,
      valAxisMinVal: 0, valAxisMaxVal: 45, valAxisMajorUnit: 10, valAxisLabelColor: C.muted, catAxisLabelColor: C.ink, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
      valGridLine: { color: "E6E8EB", size: 0.75 }, catGridLine: { style: "none" }, catAxisLineColor: C.faint, valAxisLineShow: false,
    });
    const obs = [["0", "Smooth ground surface"], ["5", "Isolated oxide nodules"], ["15", "Porous, cracked layer spreads"], ["30", "Continuous layer, no bare metal"]];
    card(s, cx, 5.6, W - MX - cx, 1.15, C.rustLight);
    txt(s, [
      { text: "+1.9 wt.% O", options: { bold: true, color: C.rust, fontFace: HF, fontSize: 22 } },
      { text: "   from cycle 20 → 30:  surface composition near saturation", options: { color: C.ink, fontSize: 14 } },
    ], { x: cx + 0.3, y: 5.6, w: W - MX - cx - 0.5, h: 1.15, valign: "middle" });
    note(s,
      "SEM and EDS show the same trend at the microscale. The fresh surface holds 2.3 percent oxygen and shows only grinding marks. At cycle 5 oxide nodules appear, and oxygen rises to 20 percent. At cycle 15 a porous and cracked layer covers much of the field, at 31.8 percent. By cycle 30 there is no bare metal in view, and oxygen reaches 37.9 percent. From cycle 20 to 30 it rose by only 1.9 points. The surface composition was close to saturation.",
      "SEM/EDS는 미세 규모에서 같은 경향을 보여 줍니다. 초기 표면은 연마 자국만 보이고 산소는 2.3%입니다. 5사이클에서 산화물 노듈이 나타나고 산소가 20%로 오릅니다. 15사이클에서는 다공성의 균열진 층이 시야 대부분을 덮고 산소는 31.8%입니다. 30사이클에는 맨 금속이 보이지 않고 산소는 37.9%에 이릅니다. 20사이클에서 30사이클까지 증가폭은 1.9%p에 그쳤습니다. 표면 조성이 포화에 가까웠습니다.");
    void obs;
  }

  // ===== 10. XRD =====
  {
    const s = base("γ-FeOOH appeared first, then Fe₃O₄");
    const iw = 6.6, ih = iw * 2140 / 2967;
    s.addImage({ path: IMG("xrd.png"), x: MX, y: 1.3, w: iw, h: ih });
    txt(s, "No.1–9 = cycles 0, 1, 3, 5, 10, 15, 20, 25, 30", { x: MX, y: 1.3 + ih + 0.05, w: iw, h: 0.3, fontSize: 11, italic: true, color: C.muted, align: "center" });
    // phase onset timeline
    const gx = 8.95, gw = W - MX - gx, gy = 1.75, rh = 0.78;
    txt(s, "Phase detected by XRD", { x: 7.5, y: 1.3, w: 5, h: 0.3, fontSize: 13, bold: true });
    const ph = [["α-Fe", 0, C.faint], ["γ-FeOOH", 10, "D38A2E"], ["Fe₃O₄", 15, C.ink], ["α-FeOOH", null, null]];
    const X = (c) => gx + (c / 30) * gw;
    ph.forEach(([n, from, col], i) => {
      const y = gy + i * rh;
      txt(s, n, { x: 7.5, y, w: 1.4, h: 0.42, fontFace: HF, fontSize: 16, bold: true, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: gx, y: y + 0.06, w: gw, h: 0.3, fill: { color: C.panel }, line: { color: C.panel } });
      if (from !== null) s.addShape(pres.shapes.RECTANGLE, { x: X(from), y: y + 0.06, w: X(30) - X(from), h: 0.3, fill: { color: col }, line: { color: col } });
      else txt(s, "not resolved (no peak at ~21.2° / ~33.2°)", { x: gx, y: y + 0.06, w: gw, h: 0.3, fontSize: 11, italic: true, color: C.muted, align: "center", valign: "middle" });
    });
    const ay = gy + 4 * rh - 0.05;
    [0, 5, 10, 15, 20, 25, 30].forEach((c) => txt(s, String(c), { x: X(c) - 0.25, y: ay, w: 0.5, h: 0.25, fontSize: 11, color: C.muted, align: "center" }));
    txt(s, "CCT cycle", { x: gx, y: ay + 0.25, w: gw, h: 0.25, fontSize: 11, color: C.muted, align: "center" });
    // onset markers
    [[10, "D38A2E"], [15, C.ink]].forEach(([c, col]) => s.addShape(pres.shapes.LINE, { x: X(c), y: gy, w: 0, h: 3 * rh - 0.1, line: { color: col, width: 1, dashType: "dash" } }));
    card(s, 7.5, 5.35, W - MX - 7.5, 1.35, C.rustLight);
    txt(s, [
      { text: "No crystalline oxide through cycle 5", options: { bold: true, breakLine: true } },
      { text: "γ-FeOOH stayed dominant for all 240 h", options: { bold: true, color: C.rust } },
    ], { x: 7.8, y: 5.35, w: W - MX - 8.0, h: 1.35, fontSize: 15, valign: "middle", paraSpaceAfter: 4 });
    note(s,
      "XRD found no crystalline oxide through cycle 5. The early nodules were amorphous or too thin to detect. Lepidocrocite appeared at cycle 10. Magnetite followed at cycle 15. Both kept growing to cycle 30. We saw no clear goethite peak at 21 or 33 degrees. A broad feature near 35 degrees may overlap with goethite, so a small amount cannot be ruled out. The layer stayed dominated by lepidocrocite for all 240 hours. Keep this in mind for the color data.",
      "XRD에서는 5사이클까지 결정성 산화물이 검출되지 않았습니다. 초기 노듈은 비정질이거나 너무 얇았습니다. 레피도크로사이트는 10사이클에, 마그네타이트는 15사이클에 나타났고, 둘 다 30사이클까지 계속 성장했습니다. 21°와 33° 부근에서 뚜렷한 괴타이트 피크는 없었습니다. 35° 부근의 넓은 피크가 괴타이트와 겹칠 수 있어 소량 존재 가능성은 배제하지 못합니다. 240시간 동안 녹층은 레피도크로사이트가 지배했습니다. 이 점을 기억하고 색 데이터를 보시겠습니다.");
  }

  // ===== 11. Color evolution =====
  {
    const s = base("The surface turned yellow during propagation");
    // swatch strip
    const n = 9, g = 0.08, sw = (W - 2 * MX - g * (n - 1)) / n;
    CYC.forEach((c, i) => {
      const x = MX + i * (sw + g);
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: sw, h: 0.7, fill: { color: hex(...RGB[i]) }, line: { color: hex(...RGB[i]) } });
      txt(s, `${c} cyc`, { x, y: 2.03, w: sw, h: 0.25, fontSize: 11, color: C.muted, align: "center" });
    });
    txt(s, "Mean color of the common rust region (camera sRGB)", { x: MX, y: 1.0, w: 8, h: 0.28, fontSize: 11, italic: true, color: C.faint });
    const bw = 5.0;
    const mL = xyChart(s, { x: MX, y: 2.45, w: bw, h: 3.6 }, [{ name: "L*", values: Ls, color: C.muted }], { ymin: 24, ymax: 36, yunit: 4, ytitle: "L*  (lightness)", layout: { x: 0.15, w: 0.81, h: 0.72 } });
    const mB = xyChart(s, { x: MX + bw + 0.3, y: 2.45, w: W - 2 * MX - bw - 0.3, h: 3.6 }, [{ name: "a*", values: As, color: C.mat }, { name: "b*", values: Bs, color: C.amber }], { ymin: -2, ymax: 16, yunit: 4, ytitle: "a*, b*", layout: { x: 0.1, w: 0.86, h: 0.72 } });
    // labels
    txt(s, "34.0", { x: mL.px(0) + 0.1, y: mL.py(34) - 0.3, w: 0.6, h: 0.25, fontSize: 12, bold: true, color: C.muted });
    txt(s, "26.4", { x: mL.px(30) - 0.55, y: mL.py(26.4) + 0.05, w: 0.6, h: 0.25, fontSize: 12, bold: true, color: C.muted });
    txt(s, "b*", { x: mB.px(30) + 0.08, y: mB.py(12.1) - 0.15, w: 0.4, h: 0.3, fontSize: 14, bold: true, color: C.amber });
    txt(s, "a*", { x: mB.px(30) + 0.08, y: mB.py(3.4) - 0.15, w: 0.4, h: 0.3, fontSize: 14, bold: true, color: C.mat });
    txt(s, "2.3", { x: mB.px(5) - 0.2, y: mB.py(2.3) - 0.38, w: 0.5, h: 0.25, fontSize: 12, bold: true, color: C.amber });
    txt(s, "12.2", { x: mB.px(15) - 0.5, y: mB.py(12.2) - 0.36, w: 0.5, h: 0.25, fontSize: 12, bold: true, color: C.amber });
    // takeaways
    const tk = [["L*", "darkens in incubation  34.0 → 30.3", C.muted], ["b*", "jumps in propagation  2.3 → 12.2", C.amber], ["", "Both level off in maturation", C.mat]];
    tk.forEach(([k, t, col], i) => {
      const x = MX + i * ((W - 2 * MX) / 3);
      txt(s, [{ text: k ? k + "  " : "", options: { fontFace: HF, fontSize: 18, bold: true, color: col } }, { text: t, options: { fontSize: 14, color: C.ink } }], { x, y: 6.2, w: (W - 2 * MX) / 3, h: 0.45, valign: "middle" });
    });
    note(s,
      "These swatches are the mean colors of the common rust region at each cycle. In the incubation stage the surface only darkened. L* fell from 34 to about 30, and b* stayed below 3. Between cycles 5 and 15, b* jumped from 2.3 to 12.2. This is the same window where lepidocrocite and magnetite crystallized. In maturation L* and b* leveled off. So each channel flags a different stage. L* catches the early darkening, and b* catches the yellowing that follows.",
      "이 색 견본은 사이클별 공통 녹 영역의 평균색입니다. 잠복기에는 표면이 어두워지기만 했습니다. L*는 34에서 약 30으로 낮아졌고, b*는 3 미만에 머물렀습니다. 5사이클에서 15사이클 사이에 b*가 2.3에서 12.2로 뛰었습니다. 레피도크로사이트와 마그네타이트가 결정화된 구간과 같습니다. 성숙기에는 L*와 b*가 평탄해졌습니다. 즉 두 채널은 서로 다른 단계를 드러냅니다. L*는 초기의 어두워짐을, b*는 그 뒤의 황색화를 잡아냅니다.");
  }

  // ===== 12. Correlations =====
  {
    const s = base("b* tracks damage in the same direction");
    const lw = 6.0;
    txt(s, "Linear fit, R²", { x: MX, y: 1.3, w: 4, h: 0.3, fontSize: 14, bold: true });
    [[C.rust, "vs rust area"], [C.faint, "vs O content"]].forEach(([col, t], i) => {
      const x = MX + 2.2 + i * 1.8;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.38, w: 0.2, h: 0.18, fill: { color: col }, line: { color: col } });
      txt(s, t, { x: x + 0.28, y: 1.32, w: 1.5, h: 0.3, fontSize: 12, color: C.muted });
    });
    s.addChart(pres.charts.BAR, [
      { name: "vs rust area", labels: ["L*", "a*", "b*"], values: [0.783, 0.911, 0.970] },
      { name: "vs O content", labels: ["L*", "a*", "b*"], values: [0.894, 0.717, 0.894] },
    ], {
      x: MX, y: 1.7, w: lw, h: 4.4, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
      chartColors: [C.rust, C.faint], showLegend: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.000", dataLabelFontSize: 12, dataLabelColor: C.ink,
      valAxisMinVal: 0, valAxisMaxVal: 1.1, valAxisMajorUnit: 0.2, valAxisLabelFormatCode: "0.0", valAxisLabelColor: C.muted, valAxisLabelFontSize: 11,
      catAxisLabelColor: C.ink, catAxisLabelFontSize: 16, catAxisLabelFontFace: HF,
      valGridLine: { color: "E6E8EB", size: 0.75 }, catGridLine: { style: "none" }, catAxisLineColor: C.faint, valAxisLineShow: false,
    });
    // Spearman: signed horizontal bars, one series per channel for per-bar colour
    const rx = MX + lw + 0.5, rw = W - MX - rx;
    txt(s, "Spearman ρ with corrosion rate (mmpy)", { x: rx, y: 1.3, w: rw, h: 0.3, fontSize: 14, bold: true });
    const chans = ["a*", "b*", "L*"], rho = [0.740, 0.883, -0.900], cols = [C.mat, C.amber, C.muted];
    s.addChart(pres.charts.BAR, [{ name: "rho", labels: chans, values: rho }], {
      x: rx, y: 1.7, w: rw, h: 3.4, barDir: "bar", barGapWidthPct: 45,
      chartColors: cols, invertedColors: cols, showLegend: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.000;−0.000", dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: C.ink,
      valAxisMinVal: -1.6, valAxisMaxVal: 1.2, valAxisMajorUnit: 0.4, valAxisLabelFormatCode: "0.0", valAxisLabelColor: C.muted, valAxisLabelFontSize: 11,
      catAxisLabelColor: C.ink, catAxisLabelFontSize: 16, catAxisLabelFontFace: HF, catAxisLabelPos: "low",
      valGridLine: { color: "E6E8EB", size: 0.75 }, catGridLine: { style: "none" }, catAxisLineColor: C.faint, valAxisLineShow: false,
    });
    txt(s, "p = 0.0009 (L*)   ·   0.0016 (b*)   ·   0.0228 (a*)", { x: rx, y: 5.12, w: rw, h: 0.3, fontSize: 11.5, color: C.muted, align: "center" });
    card(s, rx, 5.55, rw, 1.15, C.rustLight);
    txt(s, [
      { text: "b* chosen as indicator", options: { bold: true, color: C.rust, breakLine: true } },
      { text: "Rises with damage; |ρ| gap to L* is small", options: { color: C.ink, fontSize: 13 } },
    ], { x: rx + 0.3, y: 5.55, w: rw - 0.5, h: 1.15, fontSize: 16, valign: "middle" });
    note(s,
      "Now the correlations. Against rust area, b* gave the highest R squared, 0.970. Against oxygen content, L* and b* tied at 0.894. Against corrosion rate we used Spearman's rank correlation, since the relation need not be linear. L* had the largest magnitude at minus 0.900. b* followed at 0.883. Both are significant at p below 0.002. We chose b* as the working indicator because it rises with damage, which makes it easier to read. The gap to L* is small and should not be over-read.",
      "상관관계입니다. 녹 면적과는 b*의 R²가 0.970으로 가장 높았습니다. 산소 함량과는 L*와 b*가 0.894로 같았습니다. 부식속도와는 선형 관계를 가정할 필요가 없어 스피어만 순위상관을 썼습니다. 절댓값은 L*가 −0.900으로 가장 컸고, b*가 0.883으로 뒤를 이었습니다. 둘 다 p < 0.002로 유의합니다. 실무 지표로는 b*를 택했습니다. 손상이 커질수록 값이 올라가 읽기 쉽기 때문입니다. L*와의 차이는 작으므로 과하게 해석하면 안 됩니다.");
  }

  // ===== 13. b* vs mmpy =====
  {
    const s = base("Each stage sits in its own b* and corrosion-rate band");
    const box = { x: MX, y: 1.3, w: 7.3, h: 5.4 };
    const xmin = -2, xmax = 16, ymin = 0, ymax = 2;
    const L = { x: 0.11, y: 0.04, w: 0.85, h: 0.82 };
    const pa = { x: box.x + L.x * box.w, y: box.y + L.y * box.h, w: L.w * box.w, h: L.h * box.h };
    const px = (v) => pa.x + ((v - xmin) / (xmax - xmin)) * pa.w, py = (v) => pa.y + (1 - (v - ymin) / (ymax - ymin)) * pa.h;
    // stage range boxes (Table 3)
    const rng = [[-0.2, 2.7, 0.193, 0.447], [6.6, 12.2, 0.524, 0.838], [12.1, 13.9, 1.176, 1.745]];
    rng.forEach(([b0, b1, r0, r1], k) => {
      const pad = 0.25, padY = 0.04;
      s.addShape(RR, { x: px(b0 - pad), y: py(r1 + padY), w: px(b1 + pad) - px(b0 - pad), h: py(r0 - padY) - py(r1 + padY), fill: { color: STAGE[k].light }, line: { color: STAGE[k].col, width: 1, dashType: "dash" }, rectRadius: 0.06 });
    });
    // scatter: one series per stage, markers only
    const xs = [], ser = STAGE.map((st) => ({ name: st.name, values: [] }));
    CYC.forEach((c, i) => { xs.push(Bs[i]); ser.forEach((sr, k) => sr.values.push(stageOf(c) === k ? CR[i] : null)); });
    s.addChart(pres.charts.SCATTER, [{ name: "b*", values: xs }].concat(ser), {
      x: box.x, y: box.y, w: box.w, h: box.h, layout: L,
      chartColors: STAGE.map((st) => st.col), lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 13, lineDataSymbolLineSize: 1.5, lineDataSymbolLineColor: C.white,
      showLegend: false, valAxisMinVal: ymin, valAxisMaxVal: ymax, valAxisMajorUnit: 0.5,
      catAxisMinVal: xmin, catAxisMaxVal: xmax, catAxisMajorUnit: 2, catAxisCrossesAt: xmin, catAxisLabelFormatCode: "0",
      catAxisLabelColor: C.muted, valAxisLabelColor: C.muted, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
      valGridLine: { color: "E6E8EB", size: 0.75 }, catGridLine: { style: "none" }, catAxisLineColor: C.faint, valAxisLineShow: false,
      showValAxisTitle: true, valAxisTitle: "Corrosion rate (mmpy)", valAxisTitleFontSize: 12, valAxisTitleColor: C.muted,
      showCatAxisTitle: true, catAxisTitle: "b*  (yellowness)", catAxisTitleFontSize: 12, catAxisTitleColor: C.muted,
    });
    // stage labels near boxes
    txt(s, "Incubation", { x: px(-0.2) - 0.05, y: py(0.447) - 0.42, w: 1.5, h: 0.3, fontSize: 13, bold: true, color: C.inc });
    txt(s, "Propagation", { x: px(6.6) - 0.05, y: py(0.838) - 0.42, w: 1.6, h: 0.3, fontSize: 13, bold: true, color: C.prop });
    txt(s, "Maturation", { x: px(12.1) - 1.6, y: py(1.745) - 0.05, w: 1.45, h: 0.3, fontSize: 13, bold: true, color: C.mat, align: "right" });
    // rho badge
    s.addShape(RR, { x: pa.x + 0.2, y: pa.y + 0.2, w: 2.1, h: 0.75, fill: { color: C.white }, line: { color: C.line }, rectRadius: 0.08 });
    txt(s, [{ text: "ρ = 0.883", options: { fontFace: HF, bold: true, fontSize: 20, color: C.ink, breakLine: true } }, { text: "p = 0.0016,  n = 9", options: { fontSize: 11, color: C.muted } }], { x: pa.x + 0.2, y: pa.y + 0.2, w: 2.1, h: 0.75, align: "center", valign: "middle" });
    // mechanism panel
    const mx = 8.35, mw = W - MX - mx;
    card(s, mx, 1.3, mw, 5.4, C.panel);
    txt(s, "WHY YELLOWNESS FOLLOWS RATE", { x: mx + 0.3, y: 1.5, w: mw - 0.5, h: 0.3, fontSize: 11, bold: true, color: C.rust, charSpacing: 1 });
    s.addShape(pres.shapes.OVAL, { x: mx + mw / 2 - 0.6, y: 2.0, w: 1.2, h: 1.2, fill: { color: "D38A2E" }, line: { color: C.white, width: 2 } });
    txt(s, "γ-FeOOH", { x: mx, y: 3.25, w: mw, h: 0.4, fontFace: HF, fontSize: 18, bold: true, align: "center" });
    const br = [["Orange-yellow", "b* ↑", C.amber], ["Porous", "mmpy ↑", C.rust]];
    br.forEach(([a, b, col], i) => {
      const x = mx + 0.3 + i * ((mw - 0.6) / 2 + 0.0), w = (mw - 0.8) / 2;
      const xx = x + i * 0.2;
      s.addShape(RR, { x: xx, y: 3.85, w, h: 1.5, fill: { color: C.white }, line: { color: C.line }, rectRadius: 0.08 });
      txt(s, a, { x: xx, y: 3.95, w, h: 0.4, fontSize: 14, color: C.muted, align: "center" });
      txt(s, b, { x: xx, y: 4.4, w, h: 0.7, fontFace: HF, fontSize: 26, bold: true, color: col, align: "center", valign: "middle" });
    });
    txt(s, "One phase moves both signals", { x: mx + 0.3, y: 5.6, w: mw - 0.6, h: 0.8, fontSize: 16, bold: true, align: "center", valign: "middle" });
    note(s,
      "Here every sampling point is plotted with b* on the x axis and corrosion rate on the y axis. Grouped by stage, the points fall into three separate boxes. The corrosion rate bands do not overlap: 0.19 to 0.45 in incubation, 0.52 to 0.84 in propagation, 1.18 to 1.75 in maturation. The b* bands almost separate. Propagation and maturation touch only at 12.1 to 12.2. Why should yellowness follow corrosion rate? Lepidocrocite is orange-yellow, and its layer is porous. As it accumulates, b* goes up and the layer lets more electrolyte through. One phase moves both signals. So we read the correlation as physically grounded, within these test conditions.",
      "모든 측정 시점을 x축 b*, y축 부식속도로 나타냈습니다. 단계별로 묶으면 점들이 세 개의 분리된 상자에 들어갑니다. 부식속도 구간은 겹치지 않습니다. 잠복기 0.19~0.45, 전파기 0.52~0.84, 성숙기 1.18~1.75 mmpy입니다. b* 구간도 거의 분리되며, 전파기와 성숙기가 12.1~12.2에서만 맞닿습니다. 황색도가 왜 부식속도를 따라갈까요? 레피도크로사이트는 주황빛 노란색이고 다공성입니다. 이 상이 쌓이면 b*가 오르고, 층은 전해질을 더 많이 통과시킵니다. 한 상이 두 신호를 함께 움직입니다. 그래서 이 상관관계를 이번 시험 조건 안에서 물리적 근거가 있는 관계로 봅니다.");
  }

  // ===== 14. Integrated three-stage model =====
  {
    const s = base("Every method changes at the same two boundaries");
    const norm = (a) => { const lo = Math.min(...a), hi = Math.max(...a); return a.map((v) => +((v - lo) / (hi - lo)).toFixed(3)); };
    const RA = RA1.map((v, i) => (v + RA2[i]) / 2);
    const ser = [
      { name: "Rust area", values: norm(RA), color: C.rust },
      { name: "b*", values: norm(Bs), color: C.amber },
      { name: "Sz", values: norm(SZ), color: C.teal },
      { name: "Corrosion rate", values: norm(CR), color: C.ink },
    ];
    const box = { x: MX, y: 1.25, w: 7.6, h: 4.0 };
    const m = xyChart(s, box, ser, { ymin: 0, ymax: 1.05, yunit: 0.25, ytitle: "Normalized (0 = min, 1 = max)", layout: { x: 0.11, y: 0.05, w: 0.85, h: 0.76 }, symbol: 6, lineSize: 2 });
    // O content has five points only: add as separate chart overlay is overkill — list as legend note instead
    const lg = ser.map((sr) => [sr.color, sr.name]);
    lg.forEach(([col, t], i) => {
      const y = 1.45 + i * 0.36;
      s.addShape(pres.shapes.LINE, { x: 8.5, y: y + 0.14, w: 0.35, h: 0, line: { color: col, width: 2.5 } });
      txt(s, t, { x: 8.95, y, w: 2.2, h: 0.28, fontSize: 12.5 });
    });
    txt(s, "All rise sharply in propagation", { x: 8.5, y: 2.95, w: 4.2, h: 0.35, fontSize: 14, bold: true, color: C.prop });
    txt(s, "Rust area and b* plateau in maturation; Sz and corrosion rate keep rising → layer thickens, not spreads", { x: 8.5, y: 3.35, w: 4.2, h: 0.95, fontSize: 12.5, color: C.muted });
    txt(s, "Sz 0.02 → 0.10 mm  (5×)", { x: 8.5, y: 4.4, w: 4.2, h: 0.35, fontSize: 14, bold: true, color: C.teal });
    // stage schematic
    const cw = (W - 2 * MX - 0.4) / 3, y0 = 5.5;
    const desc = ["Amorphous oxide nuclei at isolated sites", "γ-FeOOH + Fe₃O₄ crystallize and spread", "Layer thickens; transport-controlled"];
    STAGE.forEach((st, k) => {
      const x = MX + k * (cw + 0.2);
      card(s, x, y0, cw, 1.25, st.light);
      // mini cross-section
      const sx = x + 0.25, sw = 1.3, sy = y0 + 0.7;
      s.addShape(pres.shapes.RECTANGLE, { x: sx, y: sy, w: sw, h: 0.3, fill: { color: "8E949B" }, line: { color: "8E949B" } });
      if (k === 0) [0.15, 0.6, 1.0].forEach((d) => s.addShape(pres.shapes.OVAL, { x: sx + d, y: sy - 0.12, w: 0.18, h: 0.16, fill: { color: "5A3A28" }, line: { color: "5A3A28" } }));
      if (k === 1) [0, 0.45, 0.9].forEach((d) => s.addShape(pres.shapes.RECTANGLE, { x: sx + d + 0.03, y: sy - 0.2, w: 0.36, h: 0.2, fill: { color: "D38A2E" }, line: { color: "D38A2E" } }));
      if (k === 2) { s.addShape(pres.shapes.RECTANGLE, { x: sx, y: sy - 0.42, w: sw, h: 0.2, fill: { color: "7A3E1D" }, line: { color: "7A3E1D" } }); s.addShape(pres.shapes.RECTANGLE, { x: sx, y: sy - 0.22, w: sw, h: 0.22, fill: { color: "D38A2E" }, line: { color: "D38A2E" } }); }
      txt(s, `${st.name}  ${st.span}`, { x: x + 1.75, y: y0 + 0.14, w: cw - 1.9, h: 0.35, fontSize: 14, bold: true, color: st.col });
      txt(s, desc[k], { x: x + 1.75, y: y0 + 0.5, w: cw - 1.9, h: 0.65, fontSize: 12.5, color: C.ink });
    });
    note(s,
      "This chart puts four measurements on one normalized scale: rust area, b*, surface height Sz and corrosion rate. Every one of them changes at the same two boundaries, between cycles 5 and 10, and between 15 and 20. Oxygen content and XRD phases, shown earlier, follow the same boundaries. In incubation, amorphous oxide nuclei form at isolated sites. In propagation, lepidocrocite and magnetite crystallize and spread across the surface. In maturation the layer stops spreading but keeps thickening. Sz rose from 0.08 to 0.10 millimeters while coverage stayed near 63 percent. Corrosion is then controlled by transport through the layer.",
      "이 그래프는 녹 면적, b*, 표면 높이 Sz, 부식속도 네 가지를 하나의 정규화 척도에 나타낸 것입니다. 모두 같은 두 경계, 즉 5~10사이클 사이와 15~20사이클 사이에서 변합니다. 앞서 본 산소 함량과 XRD 상도 같은 경계를 따릅니다. 잠복기에는 고립된 지점에 비정질 산화물 핵이 생깁니다. 전파기에는 레피도크로사이트와 마그네타이트가 결정화되며 표면 전체로 퍼집니다. 성숙기에는 층이 더 퍼지지 않고 두꺼워지기만 합니다. 면적률이 63% 부근에 머무는 동안 Sz는 0.08에서 0.10 mm로 계속 증가했습니다. 이후 부식은 층을 통한 물질 이동이 지배합니다.");
    void m;
  }

  // ===== 15. Takeaways =====
  {
    const s = base("What the color signal tells us");
    const cards = [
      ["1 → 3", "Staging without cutting", "One color value placed bare steel in one of three corrosion-rate bands", I.camera],
      ["b*", "A phase readout", "Reports the phase make-up of the outer few µm — not total metal loss", I.search],
      ["?", "Testable prediction", "Equal thickness, different γ-FeOOH fraction → different b*", I.flask],
    ];
    const cw = (W - 2 * MX - 0.6) / 3;
    cards.forEach(([big, t, d, ic], i) => {
      const x = MX + i * (cw + 0.3), y = 1.4;
      card(s, x, y, cw, 3.85, i === 0 ? C.rustLight : C.panel);
      iconDot(s, ic, x + 0.35, y + 0.35, 0.75, i === 0 ? C.rust : C.ink);
      txt(s, big, { x: x + cw - 1.9, y: y + 0.3, w: 1.6, h: 0.85, fontFace: HF, fontSize: 40, bold: true, color: C.rust, align: "right", valign: "middle" });
      txt(s, t, { x: x + 0.35, y: y + 1.45, w: cw - 0.7, h: 0.5, fontFace: HF, fontSize: 20, bold: true });
      txt(s, d, { x: x + 0.35, y: y + 2.05, w: cw - 0.7, h: 1.5, fontSize: 15, color: C.muted });
    });
    // limits
    s.addShape(RR, { x: MX, y: 5.55, w: W - 2 * MX, h: 1.1, fill: { color: C.white }, line: { color: C.line, width: 1 }, rectRadius: 0.08 });
    txt(s, "LIMITS", { x: MX + 0.35, y: 5.55, w: 1.2, h: 1.1, fontSize: 12, bold: true, color: C.muted, charSpacing: 1, valign: "middle" });
    const lim = ["9 sampling points", "One chloride-free condition", "Marine / chloride transfer untested", "Fixed lighting"];
    const lw = (W - 2 * MX - 1.8) / 4;
    lim.forEach((l, i) => txt(s, l, { x: MX + 1.6 + i * lw, y: 5.55, w: lw - 0.1, h: 1.1, fontSize: 14, valign: "middle" }));
    note(s,
      "Three points to take away. First, one surface color value placed bare carbon steel into one of three corrosion rate bands. We did not cut a single specimen to do it. Second, b* reports the phase make-up of the outer few micrometers, not total metal loss. It is a readout of what the rust layer is made of. Third, this view makes a prediction. Two layers of equal thickness but different lepidocrocite fractions should read differently. That can be tested directly. The limits are clear. The relation rests on nine sampling points under one chloride-free condition with fixed lighting. Whether it transfers to marine, chloride-rich atmospheres is still untested.",
      "세 가지를 말씀드리겠습니다. 첫째, 표면 색 값 하나로 무도장 탄소강을 세 부식속도 구간 중 하나에 배정했습니다. 시편은 하나도 자르지 않았습니다. 둘째, b*는 총 금속 손실이 아니라 최외곽 수 마이크로미터의 상 구성을 반영합니다. 녹층이 무엇으로 이루어졌는지 읽어 내는 값입니다. 셋째, 이 해석은 예측을 하나 내놓습니다. 두께는 같고 레피도크로사이트 분율이 다른 두 층은 다른 값을 보여야 하며, 이는 직접 검증할 수 있습니다. 한계도 분명합니다. 이 관계는 염화물이 없는 단일 조건, 고정 조명에서 아홉 시점으로 얻었습니다. 염화물이 많은 해양 대기로 확장되는지는 아직 검증되지 않았습니다.");
  }

  // ===== 16. Thank you =====
  {
    const s = pres.addSlide(); pageNo += 1;
    s.background = { color: C.dark };
    txt(s, "Thank you", { x: MX, y: 1.9, w: 8, h: 1.1, fontFace: HF, fontSize: 54, bold: true, color: C.white, valign: "middle" });
    txt(s, "Questions are welcome.", { x: MX, y: 3.0, w: 8, h: 0.5, fontFace: HF, fontSize: 22, italic: true, color: C.rustSoft });
    txt(s, [
      { text: "Seonghun Woo  ·  zscfvgbb@naver.com", options: { bold: true, color: C.white, breakLine: true } },
      { text: "Department of Mechanical System Engineering, Gyeongsang National University", options: { color: "D5D7DA" } },
    ], { x: MX, y: 4.0, w: 9, h: 0.8, fontSize: 15 });
    // swatch motif: b* rising
    CYC.forEach((c, i) => s.addShape(pres.shapes.RECTANGLE, { x: 9.7 + (i % 3) * 0.95, y: 1.9 + Math.floor(i / 3) * 0.95, w: 0.85, h: 0.85, fill: { color: hex(...RGB[i]) }, line: { color: "3A3D42", width: 1 } }));
    txt(s, "Acknowledgment  This research was part of the project “Training Blue Tech Leaders for Eco-Friendly Ships” (No. RS-2025-02220459), funded by the Ministry of Oceans and Fisheries, Korea.", { x: MX, y: 6.4, w: W - 2 * MX, h: 0.5, fontSize: 11, color: C.faint });
    note(s,
      "This work was funded by the Ministry of Oceans and Fisheries of Korea. Thank you for listening. I am happy to take questions.",
      "이 연구는 해양수산부 지원으로 수행했습니다. 경청해 주셔서 감사합니다. 질문 받겠습니다.");
  }

  // Make pptxgenjs chart XML schema-valid; PowerPoint repairs (drops) charts otherwise.
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(await pres.write({ outputType: "nodebuffer" }));
  await cleanSlides(zip);
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    let fixed = xml
      // missing scatter points are written as empty <c:v></c:v>; drop them (idx gaps = blanks)
      .replace(/<c:pt idx="\d+">\s*<c:v>\s*<\/c:v>\s*<\/c:pt>/g, "")
      // per-point colours (<c:dPt>) must precede <c:dLbls> within a series
      .replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (ser) => {
        const dpts = ser.match(/<c:dPt>[\s\S]*?<\/c:dPt>/g);
        if (!dpts || !ser.includes("<c:dLbls>")) return ser;
        const rest = ser.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/g, "");
        return rest.replace("<c:dLbls>", dpts.join("") + "<c:dLbls>");
      })
      // category-axis-only elements are not allowed on a value axis (scatter X axis)
      .replace(/<c:valAx>[\s\S]*?<\/c:valAx>/g, (ax) => ax.replace(/<c:(auto|lblAlgn|noMultiLvlLbl) [^>]*\/>/g, ""));
    // 2-D charts reference exactly two axes; drop dangling third axId (no serAx present)
    if (!fixed.includes("<c:serAx>")) {
      fixed = fixed.replace(/(<c:(?:bar|line|area)Chart>[\s\S]*?)(<c:axId val="\d+"\/>\s*<c:axId val="\d+"\/>)\s*<c:axId val="\d+"\/>/g, "$1$2");
    }
    if (fixed !== xml) zip.file(name, fixed);
  }
  require("fs").writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT);
})();

// ---------- slide XML cleanup ----------
async function cleanSlides(zip) {
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    // pptxgenjs repeats <a:pPr> between runs of a multi-style paragraph (and after an
    // empty leading run); only the first is valid, PowerPoint ignores the rest.
    const fixed = xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (p, inner) => {
      const PPR = /<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g;
      const all = inner.match(PPR);
      if (!all || all.length < 2) return p;
      return `<a:p>${all[0]}${inner.replace(PPR, "")}</a:p>`;
    });
    if (fixed !== xml) zip.file(name, fixed);
  }
}

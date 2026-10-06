// Tex-Corps 제주 A-STREAM (2026-10-21) — 5-minute IR pitch
// Build: node ir.js  -> ../A-STREAM2026_IR_pitch.pptx
// Technology numbers come from the ICIMC 2026 manuscript data in build.js.
// Business items the team has not supplied yet are drawn as amber dashed
// "입력 필요" boxes so they are easy to find and replace.
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const path = require("path");

const IMG = (f) => path.join(__dirname, "img", f);
const OUT = process.argv[2] || path.join(__dirname, "..", "A-STREAM2026_IR_pitch.pptx");

// ---------- palette (same rust-lab colors as the ICIMC deck) ----------
const C = {
  dark: "1F2226", ink: "2B2D31", muted: "6B6F76", faint: "A3A7AD",
  line: "D9DCE0", panel: "F3F4F6", white: "FFFFFF", bg: "FFFFFF",
  rust: "B5501F", rustLight: "FBEDE4", rustSoft: "E8A87C",
  amber: "D4961C", teal: "2A8C8C",
  inc: "3A6EA5", prop: "3C8D5A", mat: "B23A2E",
  incL: "E6EEF7", propL: "E5F2EA", matL: "F8E4E2",
  todo: "9A6A00", todoL: "FFF6DB", todoLine: "E0B23C",
};
const F = "맑은 고딕";
const W = 13.333, H = 7.5, MX = 0.6;

// ---------- data (manuscript Tables 2–3) ----------
const CYC = [0, 1, 3, 5, 10, 15, 20, 25, 30];
const CR = [0.324, 0.193, 0.447, 0.406, 0.524, 0.838, 1.176, 1.414, 1.745];
const Bs = [-0.2, 2.1, 2.7, 2.3, 6.6, 12.2, 13.2, 13.9, 12.1];
const RGB = [[79,79,80],[72,71,68],[76,75,71],[72,71,68],[72,65,56],[78,63,47],[78,60,44],[78,60,43],[73,60,44]];
const hex = (r, g, b) => [r, g, b].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
const stageOf = (c) => (c <= 5 ? 0 : c <= 15 ? 1 : 2);
const STAGE = [
  { name: "초기", en: "Incubation", band: "0.19–0.45", col: C.inc, light: C.incL },
  { name: "진행", en: "Propagation", band: "0.52–0.84", col: C.prop, light: C.propL },
  { name: "성숙", en: "Maturation", band: "1.18–1.75", col: C.mat, light: C.matL },
];

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "A-STREAM 2026 IR Pitch — 녹 색상 기반 부식 수명 예측";
pres.author = "우성훈";

let pageNo = 0;
function base(title, kicker) {
  const s = pres.addSlide();
  pageNo += 1;
  s.background = { color: C.bg };
  if (kicker) txt(s, kicker, { x: MX, y: 0.32, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.rust, charSpacing: 1 });
  if (title) txt(s, title, { x: MX, y: 0.62, w: W - 2 * MX, h: 0.7, fontSize: 28, bold: true, color: C.ink, valign: "middle" });
  txt(s, "Tex-Corps  ·  제주 A-STREAM 2026", { x: MX, y: 7.05, w: 7, h: 0.3, fontSize: 9, color: C.faint });
  txt(s, String(pageNo), { x: W - MX - 1, y: 7.05, w: 1, h: 0.3, fontSize: 9, color: C.faint, align: "right" });
  return s;
}
function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: F, fontSize: 14, color: C.ink, margin: 0, isTextBox: true, valign: "top" }, o));
}
function card(s, x, y, w, h, fill, lineCol) {
  s.addShape("roundRect", { x, y, w, h, fill: { color: fill || C.panel }, line: { color: lineCol || fill || C.panel, width: lineCol ? 1 : 0.5 }, rectRadius: 0.08 });
}
// Placeholder the team must fill in before the event.
function todo(s, text, x, y, w, h, fs = 12.5) {
  s.addShape("roundRect", { x, y, w, h, fill: { color: C.todoL }, line: { color: C.todoLine, width: 1.25, dashType: "dash" }, rectRadius: 0.06 });
  txt(s, "✎  " + text, { x: x + 0.15, y, w: w - 0.3, h, fontSize: fs, color: C.todo, valign: "middle" });
}
function iconDot(s, data, x, y, d, bg) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const p = d * 0.26;
  s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
function pill(s, text, x, y, w, color, fs = 11, h = 0.32) {
  s.addShape("roundRect", { x, y, w, h, fill: { color }, line: { color }, rectRadius: h / 2 });
  txt(s, text, { x, y, w, h, fontSize: fs, bold: true, color: C.white, align: "center", valign: "middle" });
}
function note(s, ko) { s.addNotes(ko); }

(async () => {
  const I = {
    camera: await icon(fa.FaCamera, C.white), eye: await icon(fa.FaEye, C.white),
    flask: await icon(fa.FaFlask, C.white), search: await icon(fa.FaSearch, C.white),
    arrow: await icon(fa.FaArrowRight, C.faint), check: await icon(fa.FaCheck, C.white),
    ship: await icon(fa.FaShip, C.white), bridge: await icon(fa.FaArchway, C.white),
    anchor: await icon(fa.FaAnchor, C.white), industry: await icon(fa.FaIndustry, C.white),
    mobile: await icon(fa.FaMobileAlt, C.white), file: await icon(fa.FaFileAlt, C.white),
    key: await icon(fa.FaKey, C.white), user: await icon(fa.FaUser, C.white),
    users: await icon(fa.FaUsers, C.white), handshake: await icon(fa.FaHandshake, C.white),
    coins: await icon(fa.FaCoins, C.white), mail: await icon(fa.FaEnvelope, C.rustSoft),
    clock: await icon(fa.FaClock, C.white),
  };

  // ===== 1. Title =====
  {
    const s = pres.addSlide(); pageNo += 1;
    s.background = { color: C.dark };
    txt(s, "Tex-Corps  ·  제주 A-STREAM IR 피칭  ·  2026. 10. 21", { x: MX, y: 0.5, w: 9, h: 0.3, fontSize: 12, color: C.faint });
    todo(s, "팀명 / 로고", W - MX - 2.6, 0.38, 2.6, 0.55);
    txt(s, "사진 한 장으로\n강재의 부식 속도를 읽습니다", { x: MX, y: 1.15, w: 12, h: 1.75, fontSize: 40, bold: true, color: C.white, valign: "middle", lineSpacingMultiple: 1.05 });
    txt(s, "녹 색상 기반 비파괴 부식 수명 예측 기술", { x: MX, y: 3.0, w: 11, h: 0.45, fontSize: 20, color: C.rustSoft });
    txt(s, [
      { text: "우성훈", options: { bold: true, color: C.white } },
      { text: "   경상국립대학교 기계시스템공학과", options: { color: "D5D7DA" } },
    ], { x: MX, y: 3.6, w: 12, h: 0.35, fontSize: 15 });
    const n = 9, gap = 0.12, tw = (W - 2 * MX - gap * (n - 1)) / n;
    CYC.forEach((c, i) => {
      const x = MX + i * (tw + gap);
      s.addImage({ path: IMG(`photo_${c}.png`), x, y: 4.75, w: tw, h: tw });
      txt(s, `${c * 8} h`, { x, y: 4.75 + tw + 0.08, w: tw, h: 0.3, fontSize: 11, color: C.faint, align: "center" });
    });
    txt(s, "실제 강재 표면, 촉진부식 0 → 240시간", { x: MX, y: 4.38, w: 8, h: 0.3, fontSize: 11, color: C.faint });
    note(s,
      "[약 20초]\n안녕하십니까, 경상국립대학교 우성훈입니다. 아래 사진은 같은 강재가 240시간 동안 녹슬어 가는 모습입니다. 저희는 이 녹의 '색'만 보고, 그 아래 강재가 얼마나 빠르게 부식되고 있는지를 알아내는 기술을 만들었습니다.");
  }

  // ===== 2. Problem =====
  {
    const s = base("녹은 보이지만, 얼마나 빠른지는 모릅니다", "PROBLEM");
    txt(s, "3.4%", { x: MX, y: 1.55, w: 4.4, h: 1.2, fontSize: 72, bold: true, color: C.rust, valign: "middle" });
    txt(s, "전 세계 GDP 대비 연간 부식 비용\n약 2.5조 달러 (NACE IMPACT, 2016)", { x: MX, y: 2.8, w: 4.6, h: 0.7, fontSize: 14, color: C.muted });
    // two current options
    const opts = [
      [I.eye, "육안 점검", "빠르고 싸지만 사람마다 판단이 다르고, 녹의 '양'만 볼 뿐 부식 '속도'는 알 수 없음", C.ink],
      [I.flask, "실험실 분석 (XRD · SEM · 전기화학)", "정확하지만 시편을 잘라 실험실로 가져가야 하고, 고가 장비·전문 인력·수일이 필요", C.ink],
    ];
    opts.forEach(([ic, t, d, col], i) => {
      const y = 1.5 + i * 1.3, x = 5.6;
      card(s, x, y, W - MX - x, 1.12);
      iconDot(s, ic, x + 0.25, y + 0.2, 0.72, col);
      txt(s, t, { x: x + 1.2, y: y + 0.15, w: 5.9, h: 0.38, fontSize: 16, bold: true });
      txt(s, d, { x: x + 1.2, y: y + 0.53, w: 5.9, h: 0.5, fontSize: 12.5, color: C.muted });
    });
    // gap statement
    s.addShape("roundRect", { x: MX, y: 4.35, w: W - 2 * MX, h: 1.05, fill: { color: C.dark }, line: { color: C.dark }, rectRadius: 0.08 });
    txt(s, [
      { text: "현장에는 ", options: { color: "D5D7DA" } },
      { text: "'싸고 빠른데 정량적인'", options: { color: C.rustSoft, bold: true } },
      { text: " 부식 진단 수단이 없습니다", options: { color: C.white, bold: true } },
    ], { x: MX + 0.4, y: 4.35, w: W - 2 * MX - 0.8, h: 1.05, fontSize: 21, valign: "middle" });
    todo(s, "Tex-Corps 고객 인터뷰에서 들은 현장 목소리 한 줄  (예: \"…\" — ○○조선 품질팀 과장)", MX, 5.65, W - 2 * MX, 0.95, 14);
    note(s,
      "[약 35초]\n부식으로 인한 손실은 전 세계 GDP의 약 3.4%, 2.5조 달러에 이릅니다. 그런데 현장에서 쓸 수 있는 방법은 두 가지뿐입니다. 육안 점검은 빠르지만 사람마다 판단이 다르고, 녹이 얼마나 넓은지는 봐도 얼마나 빠르게 진행 중인지는 모릅니다. 실험실 분석은 정확하지만 시편을 잘라 가야 하고 비용과 시간이 듭니다. [인터뷰 인용을 여기서 읽어 주세요.] 즉, 싸고 빠르면서도 정량적인 진단 수단이 비어 있습니다.");
  }

  // ===== 3. Solution =====
  {
    const s = base("촬영 → 녹 추출 → 색 분석 → 부식 단계 판정", "SOLUTION");
    const sw = hex(...RGB[5]);
    const steps = [
      { t: "촬영", d: "일정한 조명에서\n카메라로 표면 촬영", img: "photo_15.png" },
      { t: "녹 영역 자동 추출", d: "녹 픽셀만 골라내\n맨 강재의 간섭 제거", img: "mask_15.png" },
      { t: "색 좌표 계산", d: "녹 영역 평균\nCIE L*a*b*", color: sw },
      { t: "부식 단계 · 속도", d: "b* 값으로\n3단계 속도 구간 판정", stage: true },
    ];
    const cw = 2.55, gap = (W - 2 * MX - 4 * cw) / 3, y = 1.6;
    steps.forEach((st, i) => {
      const x = MX + i * (cw + gap);
      if (st.img) s.addImage({ path: IMG(st.img), x, y, w: cw, h: cw });
      else if (st.color) {
        s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: cw, fill: { color: st.color }, line: { color: C.line } });
        txt(s, "L* 28.0\na* 4.4\nb* 12.2", { x, y, w: cw, h: cw, fontSize: 22, bold: true, color: C.white, align: "center", valign: "middle" });
      } else {
        card(s, x, y, cw, cw, C.panel);
        STAGE.forEach((sg, k) => {
          const yy = y + 0.3 + k * 0.72, hi = k === 1;
          s.addShape("roundRect", { x: x + 0.2, y: yy, w: cw - 0.4, h: 0.58, fill: { color: hi ? sg.col : sg.light }, line: { color: sg.col, width: 1 }, rectRadius: 0.06 });
          txt(s, [
            { text: sg.name + "  ", options: { bold: true, color: hi ? C.white : sg.col } },
            { text: sg.band + " mm/년", options: { fontSize: 11, color: hi ? C.white : C.muted } },
          ], { x: x + 0.3, y: yy, w: cw - 0.6, h: 0.58, fontSize: 14, valign: "middle", align: "center" });
        });
      }
      s.addShape(pres.shapes.OVAL, { x: x - 0.18, y: y - 0.18, w: 0.5, h: 0.5, fill: { color: C.rust }, line: { color: C.white, width: 2 } });
      txt(s, String(i + 1), { x: x - 0.18, y: y - 0.18, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
      txt(s, st.t, { x, y: y + cw + 0.15, w: cw + 0.2, h: 0.4, fontSize: 17, bold: true });
      txt(s, st.d, { x, y: y + cw + 0.58, w: cw + 0.2, h: 0.65, fontSize: 12.5, color: C.muted });
      if (i < 3) s.addImage({ data: I.arrow, x: x + cw + gap / 2 - 0.13, y: y + cw / 2 - 0.13, w: 0.26, h: 0.26 });
    });
    // value chips
    const chips = ["시편 절단 없음", "카메라 한 대", "넓은 면적을 한 번에", "정량 결과"];
    const chw = (W - 2 * MX - 0.45) / 4;
    chips.forEach((c, i) => pill(s, c, MX + i * (chw + 0.15), 6.15, chw, C.ink, 14, 0.5));
    note(s,
      "[약 30초]\n저희 방법은 네 단계입니다. 표면을 촬영하고, 녹이 있는 픽셀만 자동으로 골라낸 뒤, 그 녹의 색을 국제 표준 색 좌표로 바꿉니다. 그리고 그중 노란 정도를 나타내는 b* 값으로 지금 강재가 초기·진행·성숙 중 어느 단계에 있는지, 부식 속도가 어느 구간인지를 판정합니다. 시편을 자를 필요도, 고가 장비도 필요 없습니다.");
  }

  // ===== 4. Validation =====
  {
    const s = base("녹의 '노란 정도'가 부식 속도를 따라갑니다", "TECHNOLOGY  ·  실험실 검증");
    const box = { x: MX, y: 1.45, w: 7.1, h: 5.3 };
    const xmin = -2, xmax = 16, ymin = 0, ymax = 2;
    const L = { x: 0.11, y: 0.04, w: 0.85, h: 0.82 };
    const pa = { x: box.x + L.x * box.w, y: box.y + L.y * box.h, w: L.w * box.w, h: L.h * box.h };
    const px = (v) => pa.x + ((v - xmin) / (xmax - xmin)) * pa.w, py = (v) => pa.y + (1 - (v - ymin) / (ymax - ymin)) * pa.h;
    const rng = [[-0.2, 2.7, 0.193, 0.447], [6.6, 12.2, 0.524, 0.838], [12.1, 13.9, 1.176, 1.745]];
    rng.forEach(([b0, b1, r0, r1], k) => {
      const pad = 0.25, padY = 0.04;
      s.addShape("roundRect", { x: px(b0 - pad), y: py(r1 + padY), w: px(b1 + pad) - px(b0 - pad), h: py(r0 - padY) - py(r1 + padY), fill: { color: STAGE[k].light }, line: { color: STAGE[k].col, width: 1, dashType: "dash" }, rectRadius: 0.06 });
    });
    const xs = [], ser = STAGE.map((st) => ({ name: st.en, values: [] }));
    CYC.forEach((c, i) => { xs.push(Bs[i]); ser.forEach((sr, k) => sr.values.push(stageOf(c) === k ? CR[i] : null)); });
    s.addChart(pres.charts.SCATTER, [{ name: "b*", values: xs }].concat(ser), {
      x: box.x, y: box.y, w: box.w, h: box.h, layout: L,
      chartColors: STAGE.map((st) => st.col), lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 13, lineDataSymbolLineSize: 1.5, lineDataSymbolLineColor: C.white,
      showLegend: false, valAxisMinVal: ymin, valAxisMaxVal: ymax, valAxisMajorUnit: 0.5,
      catAxisMinVal: xmin, catAxisMaxVal: xmax, catAxisMajorUnit: 2, catAxisCrossesAt: xmin, catAxisLabelFormatCode: "0",
      catAxisLabelColor: C.muted, valAxisLabelColor: C.muted, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
      catAxisLabelFontFace: F, valAxisLabelFontFace: F,
      valGridLine: { color: "E6E8EB", size: 0.75 }, catGridLine: { style: "none" }, catAxisLineColor: C.faint, valAxisLineShow: false,
      showValAxisTitle: true, valAxisTitle: "부식 속도 (mm/년)", valAxisTitleFontSize: 12, valAxisTitleColor: C.muted, valAxisTitleFontFace: F,
      showCatAxisTitle: true, catAxisTitle: "b*  (노란 정도)", catAxisTitleFontSize: 12, catAxisTitleColor: C.muted, catAxisTitleFontFace: F,
    });
    txt(s, "초기", { x: px(-0.2) - 0.05, y: py(0.447) - 0.42, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.inc });
    txt(s, "진행", { x: px(6.6) - 0.05, y: py(0.838) - 0.42, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.prop });
    txt(s, "성숙", { x: px(12.1) - 1.3, y: py(1.745) - 0.05, w: 1.15, h: 0.3, fontSize: 13, bold: true, color: C.mat, align: "right" });
    // proof points
    const rx = 8.15, rw = W - MX - rx;
    const stats = [
      ["ρ = 0.88", "b*와 부식 속도의 순위 상관 (p < 0.002)"],
      ["3개 구간", "단계별 부식 속도 구간이 서로 겹치지 않음"],
      ["6가지 분석", "SEM/EDS · XRD · 전기화학 · 3D 표면 · 녹 면적 · 색으로 교차 검증"],
    ];
    stats.forEach(([v, d], i) => {
      const y = 1.55 + i * 1.3;
      card(s, rx, y, rw, 1.15, i === 0 ? C.rustLight : C.panel);
      txt(s, v, { x: rx + 0.3, y: y + 0.1, w: rw - 0.6, h: 0.55, fontSize: 24, bold: true, color: C.rust, valign: "middle" });
      txt(s, d, { x: rx + 0.3, y: y + 0.63, w: rw - 0.5, h: 0.45, fontSize: 12, color: C.muted });
    });
    txt(s, "탄소강 · 240시간 촉진부식 · 9개 시점 × 시편 2개", { x: rx, y: 5.55, w: rw, h: 0.3, fontSize: 11.5, color: C.muted });
    txt(s, "ICIMC 2026 국제학회(11월, 서울) 발표 예정", { x: rx, y: 5.9, w: rw, h: 0.3, fontSize: 11.5, color: C.muted });
    note(s,
      "[약 40초]\n이게 실제로 되는지 실험실에서 검증했습니다. 탄소강을 240시간 촉진부식 시키면서 아홉 시점에 측정했습니다. 가로축이 녹의 노란 정도 b*, 세로축이 전기화학으로 잰 실제 부식 속도입니다. 보시는 것처럼 세 단계가 서로 다른 구역에 깔끔하게 나뉩니다. b*와 부식 속도의 순위 상관계수는 0.88이었습니다. 이 결과를 전자현미경, X선 회절, 3D 표면 측정까지 여섯 가지 분석으로 교차 확인했습니다. 노란색 녹인 레피도크로사이트가 다공성이라 부식을 더 빠르게 만들기 때문에, 색과 속도가 함께 움직이는 물리적 근거도 있습니다.");
  }

  // ===== 5. Differentiation =====
  {
    const s = base("현장에서 쓸 수 있으면서, 숫자로 답합니다", "WHY US");
    const cols = ["육안 점검", "실험실 분석", "부착형 부식 센서", "우리 기술"];
    const rows = [
      ["비파괴 (시편 절단 없음)", ["○", "×", "○", "○"]],
      ["정량 결과 (부식 속도 구간)", ["×", "○", "○", "○"]],
      ["넓은 면적을 한 번에", ["○", "×", "×", "○"]],
      ["설치 · 장비 부담", ["낮음", "높음", "센서 설치 필요", "카메라"]],
      ["녹의 성분 정보", ["×", "○", "×", "△ (색으로 간접)"]],
    ];
    const x0 = MX, lw = 3.3, cw = (W - 2 * MX - lw) / 4, rh = 0.68, y0 = 1.6;
    // highlight our column
    s.addShape("roundRect", { x: x0 + lw + 3 * cw + 0.05, y: y0 - 0.05, w: cw - 0.1, h: rh * (rows.length + 1) + 0.1, fill: { color: C.rustLight }, line: { color: C.rust, width: 1.5 }, rectRadius: 0.08 });
    cols.forEach((c, j) => txt(s, c, { x: x0 + lw + j * cw, y: y0, w: cw, h: rh, fontSize: 15, bold: true, color: j === 3 ? C.rust : C.ink, align: "center", valign: "middle" }));
    rows.forEach(([label, vals], i) => {
      const y = y0 + (i + 1) * rh;
      s.addShape(pres.shapes.LINE, { x: x0, y, w: W - 2 * MX, h: 0, line: { color: C.line, width: 1 } });
      txt(s, label, { x: x0 + 0.1, y, w: lw - 0.1, h: rh, fontSize: 14, valign: "middle" });
      vals.forEach((v, j) => {
        const mark = v.length === 1;
        const col = v === "○" ? (j === 3 ? C.rust : C.prop) : v === "×" ? C.faint : j === 3 ? C.rust : C.muted;
        txt(s, v, { x: x0 + lw + j * cw, y, w: cw, h: rh, fontSize: mark ? 22 : 13, bold: j === 3, color: col, align: "center", valign: "middle" });
      });
    });
    todo(s, "특허 현황 (예: 녹 색상 기반 부식 단계 판정 방법 — 출원번호 / 출원일)", MX, 5.95, 7.4, 0.7);
    todo(s, "경쟁 업체 · 제품명 (인터뷰에서 들은 것)", MX + 7.6, 5.95, W - 2 * MX - 7.6, 0.7);
    note(s,
      "[약 30초]\n기존 방법과 비교하면, 육안 점검은 넓게 볼 수 있지만 숫자가 없고, 실험실 분석과 부착형 센서는 숫자는 주지만 시편을 자르거나 센서를 설치한 한 지점만 봅니다. 저희 기술은 카메라 한 대로 넓은 면적을 비파괴로 보면서 부식 속도 구간이라는 숫자로 답합니다. [특허 현황을 한 문장으로 언급해 주세요.]");
  }

  // ===== 6. Market =====
  {
    const s = base("첫 고객은 해양 강구조물 점검 현장입니다", "MARKET");
    // TAM/SAM/SOM rings
    const cx = 3.4, cy = 4.15;
    [[2.55, "F6E3D8"], [1.75, "EFC9B2"], [0.95, C.rust]].forEach(([r, col]) => s.addShape(pres.shapes.OVAL, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color: col }, line: { color: C.white, width: 2 } }));
    txt(s, "TAM", { x: cx - 0.6, y: cy - 2.35, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.rust, align: "center" });
    txt(s, "SAM", { x: cx - 0.6, y: cy - 1.55, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.rust, align: "center" });
    txt(s, "SOM", { x: cx - 0.6, y: cy - 0.35, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.white, align: "center" });
    txt(s, "₩ ?", { x: cx - 0.6, y: cy - 0.05, w: 1.2, h: 0.4, fontSize: 18, bold: true, color: C.white, align: "center" });
    // ring descriptions
    const rx = 6.6, rw = W - MX - rx;
    const ringTxt = [
      ["TAM  글로벌 부식 모니터링 시장", "보고서별 약 USD 1.1–2.9B (2025), 연 7–11% 성장 — 사용할 수치 1개를 골라 출처와 함께 확정"],
      ["SAM  국내 강구조물 부식 점검 시장", "대상(조선·교량·항만·플랜트) 범위와 금액"],
      ["SOM  3년 내 확보 가능한 시장", "첫 타깃 고객 수 × 연간 점검 횟수 × 단가"],
    ];
    ringTxt.forEach(([t, d], i) => {
      const y = 1.55 + i * 1.08;
      txt(s, t, { x: rx, y, w: rw, h: 0.35, fontSize: 15, bold: true });
      todo(s, d, rx, y + 0.38, rw, 0.62, 11);
    });
    // segments
    txt(s, "타깃 고객 후보", { x: rx, y: 4.85, w: 4, h: 0.3, fontSize: 13, bold: true, color: C.muted });
    const seg = [[I.ship, "조선·선박 수리"], [I.bridge, "교량·강구조물"], [I.anchor, "항만·해양 설비"], [I.industry, "노후 산업설비"]];
    const sw = (rw - 0.3) / 4;
    seg.forEach(([ic, t], i) => {
      const x = rx + i * (sw + 0.1);
      card(s, x, 5.2, sw, 1.45, i === 0 ? C.rustLight : C.panel);
      iconDot(s, ic, x + sw / 2 - 0.33, 5.35, 0.66, i === 0 ? C.rust : C.ink);
      txt(s, t, { x: x + 0.05, y: 6.1, w: sw - 0.1, h: 0.45, fontSize: 12, bold: true, align: "center", valign: "middle" });
    });
    note(s,
      "[약 30초]\n전체 시장은 글로벌 부식 모니터링 시장으로, [확정한 금액과 출처]입니다. 저희는 그중 국내 강구조물 부식 점검, 특히 경상국립대 통영캠퍼스가 가까운 조선·해양 현장을 첫 고객으로 봅니다. [SAM / SOM 금액과 근거를 한 문장씩.] Tex-Corps 인터뷰에서 [가장 반응이 좋았던 고객군]이 가장 큰 관심을 보였습니다.");
  }

  // ===== 7. Business model =====
  {
    const s = base("수익 모델", "BUSINESS MODEL");
    const bm = [
      [I.mobile, "점검 앱 구독", "현장 담당자가 스마트폰으로 촬영 → 부식 단계·속도 리포트", "월/연 구독료"],
      [I.file, "진단 서비스", "구조물 단위 촬영·분석 후 유지보수 우선순위 리포트 제공", "건당 진단비"],
      [I.key, "기술 라이선스", "드론·점검 로봇·안전진단 업체의 소프트웨어에 분석 엔진 탑재", "라이선스 · 로열티"],
    ];
    const cw = (W - 2 * MX - 0.6) / 3;
    bm.forEach(([ic, t, d, rev], i) => {
      const x = MX + i * (cw + 0.3), y = 1.6;
      card(s, x, y, cw, 3.55);
      iconDot(s, ic, x + 0.35, y + 0.35, 0.8, C.ink);
      txt(s, t, { x: x + 0.35, y: y + 1.35, w: cw - 0.7, h: 0.45, fontSize: 19, bold: true });
      txt(s, d, { x: x + 0.35, y: y + 1.85, w: cw - 0.7, h: 0.95, fontSize: 13, color: C.muted });
      pill(s, rev, x + 0.35, y + 2.9, cw - 0.7, C.rust, 12, 0.38);
    });
    todo(s, "실제 계획에 맞는 모델만 남기고 가격을 넣어 주세요  (예: 첫 해는 진단 서비스로 데이터 확보 → 이후 앱 구독 전환, 구조물 1기당 ○○만 원)", MX, 5.45, W - 2 * MX, 1.15, 14);
    note(s,
      "[약 25초]\n[확정한 모델 기준으로 수정하세요.] 처음에는 구조물 단위 진단 서비스로 시작해 현장 데이터를 모으고, 그 데이터로 정확도를 높인 뒤 현장 담당자가 직접 쓰는 앱 구독으로 확장하겠습니다. 장기적으로는 드론·점검 로봇 업체에 분석 엔진을 라이선스할 수 있습니다.");
  }

  // ===== 8. Traction & roadmap =====
  {
    const s = base("지금까지 한 것, 앞으로 할 것", "TRACTION  ·  ROADMAP");
    // done
    const lw = 5.3;
    txt(s, "지금까지", { x: MX, y: 1.5, w: lw, h: 0.35, fontSize: 16, bold: true, color: C.rust });
    const done = [
      ["실험실 검증 완료", "240시간 촉진부식, 6가지 분석 교차 검증"],
      ["국제학회 발표 확정", "ICIMC 2026 구두 발표 (2026.11, 서울)"],
    ];
    done.forEach(([t, d], i) => {
      const y = 2.0 + i * 1.0;
      iconDot(s, I.check, MX, y + 0.05, 0.5, C.prop);
      txt(s, t, { x: MX + 0.7, y, w: lw - 0.7, h: 0.35, fontSize: 15, bold: true });
      txt(s, d, { x: MX + 0.7, y: y + 0.38, w: lw - 0.7, h: 0.35, fontSize: 12.5, color: C.muted });
    });
    todo(s, "Tex-Corps 고객 인터뷰 ○○건 — 핵심 발견 1줄", MX, 4.05, lw, 0.75, 12.5);
    todo(s, "MOU · 실증 협의 · 수상 · 특허 출원 등", MX, 4.95, lw, 0.75, 12.5);
    // roadmap
    const rx = MX + lw + 0.6, rw = W - MX - rx;
    txt(s, "로드맵 (안)", { x: rx, y: 1.5, w: rw, h: 0.35, fontSize: 16, bold: true, color: C.rust });
    const steps = [
      ["2026 하반기", "해양(염화물) 환경 · 조명 변화 조건으로 검증 확대", false],
      ["2027 상반기", "현장 촬영 앱 MVP · 첫 현장 실증 (PoC)", false],
      ["2027 하반기", "법인 설립 · 첫 유료 고객", true],
      ["2028", "드론 · 점검 로봇 연계, 데이터 기반 수명 예측", true],
    ];
    const tx = rx + 0.2, ty0 = 2.05, step = 1.13;
    s.addShape(pres.shapes.LINE, { x: tx + 0.12, y: ty0 + 0.12, w: 0, h: step * 3, line: { color: C.line, width: 2 } });
    steps.forEach(([when, what, tbd], i) => {
      const y = ty0 + i * step;
      s.addShape(pres.shapes.OVAL, { x: tx, y, w: 0.24, h: 0.24, fill: { color: i === 0 ? C.rust : C.white }, line: { color: C.rust, width: 2 } });
      txt(s, when, { x: tx + 0.45, y: y - 0.06, w: 1.7, h: 0.35, fontSize: 14, bold: true, color: C.rust });
      if (tbd) todo(s, what + "  — 일정 확인", tx + 2.15, y - 0.12, rw - 2.35, 0.6, 12);
      else txt(s, what, { x: tx + 2.15, y: y - 0.06, w: rw - 2.35, h: 0.6, fontSize: 13.5 });
    });
    note(s,
      "[약 35초]\n지금까지 실험실 검증을 마쳤고, 다음 달 ICIMC 국제학회에서 결과를 발표합니다. Tex-Corps를 통해 [○○명]의 현장 담당자를 인터뷰했고, [핵심 발견]을 확인했습니다. 다음 과제는 분명합니다. 지금 결과는 염화물이 없는 실험실 조건이라, 올해 안에 해양 환경과 조명 변화 조건으로 검증을 넓히고, 내년 상반기에 현장 앱으로 첫 실증을 하겠습니다.");
  }

  // ===== 9. Team =====
  {
    const s = base("팀", "TEAM");
    const cw = (W - 2 * MX - 0.9) / 4, y = 1.65;
    const team = [["우성훈", "대표 · 기술 개발", "경상국립대학교 기계시스템공학과", false]].concat(
      [1, 2, 3].map((k) => [`팀원 ${k}`, "역할", "소속 · 핵심 경력 1줄", true]));
    team.forEach(([n, r, d, tbd], i) => {
      const x = MX + i * (cw + 0.3);
      card(s, x, y, cw, 3.3, i === 0 ? C.rustLight : C.panel);
      iconDot(s, I.user, x + cw / 2 - 0.5, y + 0.35, 1.0, i === 0 ? C.rust : C.faint);
      if (tbd) todo(s, `${n} 이름 / ${r}\n${d}`, x + 0.2, y + 1.6, cw - 0.4, 1.45, 12);
      else {
        txt(s, n, { x: x + 0.2, y: y + 1.55, w: cw - 0.4, h: 0.45, fontSize: 20, bold: true, align: "center" });
        txt(s, r, { x: x + 0.2, y: y + 2.05, w: cw - 0.4, h: 0.35, fontSize: 13.5, bold: true, color: C.rust, align: "center" });
        todo(s, "핵심 경력 · 역할 1줄", x + 0.2, y + 2.5, cw - 0.4, 0.6, 11.5);
      }
    });
    todo(s, "지도교수 · 멘토 · 자문 (예: 부식/재료 분야 지도교수, Tex-Corps 멘토)", MX, 5.25, W - 2 * MX, 0.65, 13);
    txt(s, "연구 기반: 해양수산부 「친환경 선박 블루테크 리더 양성」 사업 (RS-2025-02220459)", { x: MX, y: 6.15, w: W - 2 * MX, h: 0.35, fontSize: 12, color: C.muted });
    note(s,
      "[약 20초]\n저희 팀은 [팀 소개 — 각자 맡은 역할을 한 문장으로]. 부식·재료 분야 연구실의 측정 인프라와 [지도교수/멘토]의 지원을 받고 있습니다.");
  }

  // ===== 10. Ask & close =====
  {
    const s = pres.addSlide(); pageNo += 1;
    s.background = { color: C.dark };
    txt(s, "ASK", { x: MX, y: 0.55, w: 4, h: 0.3, fontSize: 12, bold: true, color: C.rustSoft, charSpacing: 2 });
    txt(s, "함께 현장을 찾고 있습니다", { x: MX, y: 0.9, w: 12, h: 0.8, fontSize: 34, bold: true, color: C.white, valign: "middle" });
    const asks = [
      [I.handshake, "실증 파트너", "조선소 · 지자체 · 안전진단 업체 등 현장 실증 대상"],
      [I.coins, "투자 · 지원", "금액과 사용처 (예: 현장 실증 ○○%, 앱 개발 ○○%)"],
      [I.users, "멘토링 · 네트워크", "필요한 도움 1줄"],
    ];
    const cw = (W - 2 * MX - 0.6) / 3;
    asks.forEach(([ic, t, d], i) => {
      const x = MX + i * (cw + 0.3), y = 2.0;
      s.addShape("roundRect", { x, y, w: cw, h: 2.55, fill: { color: "2B2F35" }, line: { color: "3A3D42" }, rectRadius: 0.08 });
      iconDot(s, ic, x + 0.3, y + 0.3, 0.75, C.rust);
      txt(s, t, { x: x + 1.25, y: y + 0.3, w: cw - 1.5, h: 0.75, fontSize: 19, bold: true, color: C.white, valign: "middle" });
      todo(s, d, x + 0.3, y + 1.3, cw - 0.6, 1.0, 12.5);
    });
    txt(s, "사진 한 장으로 강재의 부식 속도를 읽습니다", { x: MX, y: 5.0, w: 9, h: 0.55, fontSize: 22, bold: true, color: C.rustSoft });
    s.addImage({ data: I.mail, x: MX, y: 5.78, w: 0.3, h: 0.3 });
    txt(s, "우성훈  ·  zscfvgbb@naver.com", { x: MX + 0.45, y: 5.72, w: 8, h: 0.4, fontSize: 15, color: C.white, valign: "middle" });
    CYC.forEach((c, i) => s.addShape(pres.shapes.RECTANGLE, { x: 9.7 + (i % 3) * 0.95, y: 4.75 + Math.floor(i / 3) * 0.6, w: 0.85, h: 0.5, fill: { color: hex(...RGB[i]) }, line: { color: "3A3D42", width: 1 } }));
    txt(s, "감사합니다", { x: MX, y: 6.4, w: 6, h: 0.5, fontSize: 18, bold: true, color: C.white });
    note(s,
      "[약 25초]\n저희가 지금 가장 필요한 것은 첫 현장입니다. [실증 파트너 / 투자 금액과 용도 / 필요한 도움을 구체적으로.] 사진 한 장으로 강재의 부식 속도를 읽는 기술, 현장에서 함께 증명할 파트너를 찾고 있습니다. 감사합니다.");
  }

  // Make pptxgenjs chart XML schema-valid (same fixes as build.js).
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(await pres.write({ outputType: "nodebuffer" }));
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    const fixed = xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (p, inner) => {
      const PPR = /<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g;
      const all = inner.match(PPR);
      if (!all || all.length < 2) return p;
      return `<a:p>${all[0]}${inner.replace(PPR, "")}</a:p>`;
    });
    if (fixed !== xml) zip.file(name, fixed);
  }
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    const fixed = xml
      .replace(/<c:pt idx="\d+">\s*<c:v>\s*<\/c:v>\s*<\/c:pt>/g, "")
      .replace(/<c:valAx>[\s\S]*?<\/c:valAx>/g, (ax) => ax.replace(/<c:(auto|lblAlgn|noMultiLvlLbl) [^>]*\/>/g, ""));
    if (fixed !== xml) zip.file(name, fixed);
  }
  require("fs").writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT);
})();

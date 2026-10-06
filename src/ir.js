// Tex-Corps 제주 A-STREAM (2026-10-21) — 5-minute IR pitch for team MLC
// Build: node ir.js  -> ../A-STREAM2026_IR_pitch.pptx
// Technology numbers come from the ICIMC 2026 manuscript data in build.js.
// Customer findings come from the 35 U.S. customer-discovery interviews
// (2026-07-06 to 07-23) and the team's U.S. presentation (BMC, ecosystem,
// inspection levels, next steps). Items the team has not supplied yet are
// drawn as amber dashed "✎" boxes so they are easy to find and replace.
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
  amber: "D4961C", teal: "2A8C8C", sea: "1F5F7A", seaLight: "E3EFF4",
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
  { name: "낮음", en: "Incubation", band: "0.19–0.45", col: C.inc, light: C.incL },
  { name: "중간", en: "Propagation", band: "0.52–0.84", col: C.prop, light: C.propL },
  { name: "높음", en: "Maturation", band: "1.18–1.75", col: C.mat, light: C.matL },
];

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "MLC — A-STREAM 2026 IR Pitch";
pres.author = "MLC (우성훈, 박준범)";

let pageNo = 0;
function base(title, kicker) {
  const s = pres.addSlide();
  pageNo += 1;
  s.background = { color: C.bg };
  if (kicker) txt(s, kicker, { x: MX, y: 0.32, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.rust, charSpacing: 1 });
  if (title) txt(s, title, { x: MX, y: 0.62, w: W - 2 * MX, h: 0.7, fontSize: 28, bold: true, color: C.ink, valign: "middle" });
  txt(s, "MLC  ·  Tex-Corps 제주 A-STREAM 2026", { x: MX, y: 7.05, w: 7, h: 0.3, fontSize: 9, color: C.faint });
  txt(s, String(pageNo), { x: W - MX - 1, y: 7.05, w: 1, h: 0.3, fontSize: 9, color: C.faint, align: "right" });
  return s;
}
function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: F, fontSize: 14, color: C.ink, margin: 0, isTextBox: true, valign: "top" }, o));
}
function card(s, x, y, w, h, fill, lineCol) {
  s.addShape("roundRect", { x, y, w, h, fill: { color: fill || C.panel }, line: { color: lineCol || fill || C.panel, width: lineCol ? 1.25 : 0.5 }, rectRadius: 0.08 });
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
function pill(s, text, x, y, w, color, fs = 11, h = 0.32, fg = C.white) {
  s.addShape("roundRect", { x, y, w, h, fill: { color }, line: { color }, rectRadius: h / 2 });
  txt(s, text, { x, y, w, h, fontSize: fs, bold: true, color: fg, align: "center", valign: "middle" });
}
function note(s, ko) { s.addNotes(ko); }

(async () => {
  const I = {
    arrow: await icon(fa.FaArrowRight, C.faint), arrowR: await icon(fa.FaArrowRight, C.rust),
    check: await icon(fa.FaCheck, C.white), times: await icon(fa.FaTimes, C.white),
    route: await icon(fa.FaRoute, C.white), ship: await icon(fa.FaShip, C.white),
    coins: await icon(fa.FaCoins, C.white), helicopter: await icon(fa.FaHelicopter, C.white),
    handshake: await icon(fa.FaHandshake, C.white), users: await icon(fa.FaUsers, C.white),
    drone: await icon(fa.FaCamera, C.white), list: await icon(fa.FaListOl, C.white),
    crown: await icon(fa.FaUserTie, C.white), user: await icon(fa.FaHardHat, C.white),
    star: await icon(fa.FaUserCog, C.white), flask: await icon(fa.FaFlask, C.white),
    mail: await icon(fa.FaEnvelope, C.rustSoft), shield: await icon(fa.FaShieldAlt, C.white),
    wind: await icon(fa.FaWind, C.white), industry: await icon(fa.FaIndustry, C.white),
    ruler: await icon(fa.FaRulerVertical, C.white), map: await icon(fa.FaMapMarkerAlt, C.white),
    calc: await icon(fa.FaCalculator, C.white),
  };

  // ===== 1. Title =====
  {
    const s = pres.addSlide(); pageNo += 1;
    s.background = { color: C.dark };
    txt(s, "Tex-Corps  ·  제주 A-STREAM IR 피칭  ·  2026. 10. 21", { x: MX, y: 0.5, w: 9, h: 0.3, fontSize: 12, color: C.faint });
    txt(s, "MLC", { x: MX, y: 1.25, w: 3, h: 0.6, fontSize: 30, bold: true, color: C.rustSoft, charSpacing: 4 });
    txt(s, "어디부터 점검할지,\n녹의 색이 먼저\n알려드립니다", { x: MX, y: 1.95, w: 6.6, h: 2.5, fontSize: 38, bold: true, color: C.white, valign: "middle", lineSpacingMultiple: 1.05 });
    txt(s, "녹 색상 기반 AI 금속 열화 예측 플랫폼", { x: MX, y: 4.6, w: 6.6, h: 0.45, fontSize: 19, color: C.rustSoft });
    txt(s, [
      { text: "우성훈 · 박준범", options: { bold: true, color: C.white } },
      { text: "   경상국립대학교 기계시스템공학과", options: { color: "D5D7DA" } },
    ], { x: MX, y: 5.3, w: 7, h: 0.35, fontSize: 15 });
    // photo
    const pw = 5.6, ph = pw * 284 / 386, px = W - MX - pw, py = 1.3;
    s.addImage({ path: IMG("offshore_legs.png"), x: px, y: py, w: pw, h: ph });
    txt(s, "해양 플랫폼 하부 구조의 비말대(splash zone) 부식", { x: px, y: py + ph + 0.1, w: pw, h: 0.3, fontSize: 11, color: C.faint, align: "right" });
    note(s,
      "[약 20초]\n안녕하십니까, MLC 우성훈입니다. 해양 구조물을 점검할 때 가장 어려운 질문은 '어디부터 봐야 하나'입니다. 저희는 녹의 색을 읽어서 그 답을 먼저 드리는 기술을 만들고 있습니다.");
  }

  // ===== 2. Problem =====
  {
    const s = base("해양 구조물 점검, 가장 비싼 건 '접근'입니다", "PROBLEM");
    const iw = 4.55;
    s.addImage({ path: IMG("offshore_zones.png"), x: MX, y: 1.5, w: iw, h: iw * 678 / 681 });
    const rx = MX + iw + 0.45, rw = W - MX - rx;
    const pains = [
      [I.route, "모든 지점을 똑같이 순회", "상태와 관계없이 정해진 목록을 전부 방문", "Williams · 점검 관리자"],
      [I.helicopter, "비말대 · 해저 구간은 접근 자체가 어려움", "로프 액세스 · 잠수 · ROV를 미리 일정 잡아야만 확인 가능", "Deepwater Corrosion Services · CTO"],
      [I.ship, "비용은 '검사'가 아니라 '이동·대기'", "선박 용선일 · 기상 대기 · 리프트 · 비계 임대가 실제 비용", "Deepwater · Brown Corrosion · MISTRAS"],
    ];
    pains.forEach(([ic, t, d, who], i) => {
      const y = 1.5 + i * 1.22;
      card(s, rx, y, rw, 1.08);
      iconDot(s, ic, rx + 0.22, y + 0.18, 0.72, C.ink);
      txt(s, t, { x: rx + 1.15, y: y + 0.12, w: rw - 1.3, h: 0.36, fontSize: 15.5, bold: true });
      txt(s, d, { x: rx + 1.15, y: y + 0.48, w: rw - 1.3, h: 0.3, fontSize: 12.5, color: C.muted });
      txt(s, "— " + who, { x: rx + 1.15, y: y + 0.76, w: rw - 1.3, h: 0.25, fontSize: 10.5, italic: true, color: C.faint });
    });
    s.addShape("roundRect", { x: rx, y: 5.3, w: rw, h: 1.3, fill: { color: C.dark }, line: { color: C.dark }, rectRadius: 0.08 });
    txt(s, [
      { text: "점검 주기는 규제로 정해져 줄일 수 없습니다.", options: { color: "D5D7DA", breakLine: true } },
      { text: "문제는 그 주기 안에서 ", options: { color: C.white, bold: true } },
      { text: "어디에 사람과 배를 보낼지", options: { color: C.rustSoft, bold: true } },
      { text: "입니다.", options: { color: C.white, bold: true } },
    ], { x: rx + 0.35, y: 5.3, w: rw - 0.6, h: 1.3, fontSize: 17, valign: "middle", paraSpaceAfter: 4 });
    note(s,
      "[약 30초]\n해양 플랫폼의 하부, 특히 파도가 치는 비말대는 부식이 가장 빠른 곳인데, 로프 액세스나 잠수, ROV 없이는 볼 수가 없습니다. 미국 현장에서 들은 이야기는 일관됐습니다. 지금은 상태와 상관없이 모든 지점을 똑같이 돌고, 비용의 대부분은 검사 자체가 아니라 배를 빌리고 날씨를 기다리고 리프트를 빌리는 데서 나옵니다. 점검 주기는 규제로 정해져 있어서 줄일 수 없습니다. 결국 문제는 그 주기 안에서 어디에 사람과 배를 먼저 보낼지입니다.");
  }

  // ===== 3. Customer discovery =====
  {
    const s = base("미국 현장에서 35명을 만났습니다", "CUSTOMER DISCOVERY");
    txt(s, "35", { x: MX, y: 1.45, w: 2.6, h: 1.3, fontSize: 84, bold: true, color: C.rust, valign: "middle" });
    txt(s, "고객 인터뷰\n대면 24 · 화상 11\n2026. 7. 6 – 7. 23", { x: MX + 2.5, y: 1.62, w: 2.6, h: 1.05, fontSize: 14, color: C.muted });
    // segment bars
    const segs = [
      ["해양 석유·가스 플랫폼", 19, C.sea],
      ["미드스트림 가스 파이프라인", 5, C.sea],
      ["석유·가스 일반 (검사·재료)", 2, C.sea],
      ["자동차 · 배터리 · 반도체 등", 9, C.faint],
    ];
    const bx = MX + 3.1, bw = 3.0, by = 3.15;
    txt(s, "인터뷰 대상 산업", { x: MX, y: by - 0.1, w: 4, h: 0.3, fontSize: 12, bold: true, color: C.muted });
    segs.forEach(([l, n, col], i) => {
      const y = by + 0.35 + i * 0.55;
      txt(s, l, { x: MX, y, w: 3.0, h: 0.38, fontSize: 12.5, valign: "middle" });
      s.addShape("rect", { x: bx, y: y + 0.05, w: bw * n / 19, h: 0.28, fill: { color: col }, line: { color: col } });
      txt(s, String(n), { x: bx + bw * n / 19 + 0.1, y, w: 0.5, h: 0.38, fontSize: 13, bold: true, color: col === C.faint ? C.muted : C.sea, valign: "middle" });
    });
    txt(s, "스테인리스 위주의 자동차 · 배터리 공정에서는 부식이 현장 문제가 아니었음 → 타깃에서 제외", { x: MX, y: 5.75, w: 6.4, h: 0.6, fontSize: 12, color: C.muted });
    // orgs
    const rx = 7.35, rw = W - MX - rx;
    card(s, rx, 1.5, rw, 4.85, C.seaLight);
    txt(s, "만난 조직 (일부)", { x: rx + 0.3, y: 1.68, w: rw - 0.6, h: 0.3, fontSize: 12, bold: true, color: C.sea });
    const orgs = ["MISTRAS Group", "Williams", "Deepwater Corrosion Services", "Farwest Corrosion Control", "American Innovations", "Brown Corrosion Services", "HF Sinclair", "Tinker & Rasor"];
    let cx = rx + 0.3, cy = 2.12;
    orgs.forEach((o) => {
      const w = 0.3 + o.length * 0.088;
      if (cx + w > rx + rw - 0.25) { cx = rx + 0.3; cy += 0.46; }
      s.addShape("roundRect", { x: cx, y: cy, w, h: 0.38, fill: { color: C.white }, line: { color: C.line }, rectRadius: 0.19 });
      txt(s, o, { x: cx, y: cy, w, h: 0.38, fontSize: 11.5, align: "center", valign: "middle" });
      cx += w + 0.12;
    });
    const roles = [["5", "의사결정권자 (CEO · 대표)"], ["11", "현장 부식 엔지니어 · 검사원"], ["+", "방식(防蝕) 전문 업체 · 엔지니어링 매니저"]];
    roles.forEach(([n, l], i) => {
      const y = 4.6 + i * 0.55;
      txt(s, n, { x: rx + 0.3, y, w: 0.7, h: 0.45, fontSize: 22, bold: true, color: C.sea, valign: "middle" });
      txt(s, l, { x: rx + 1.0, y, w: rw - 1.2, h: 0.45, fontSize: 13, valign: "middle" });
    });
    note(s,
      "[약 30초]\n올해 7월, Tex-Corps 프로그램으로 미국에서 35명을 인터뷰했습니다. 24명은 직접 만났고 11명은 화상으로 만났습니다. 처음에는 자동차, 배터리, 반도체도 만났지만, 스테인리스를 쓰는 공정에서는 부식이 현장의 문제가 아니었습니다. 그래서 부식이 실제로 돈과 안전 문제가 되는 해양 석유·가스와 파이프라인으로 좁혔고, MISTRAS, Williams 같은 회사의 현장 검사원과 대표들을 만났습니다.");
  }

  // ===== 4. Pivot =====
  {
    const s = base("현장이 저희 가설을 바꿨습니다", "WHAT WE LEARNED");
    const cw = 5.55, y = 1.55, h = 2.75;
    // before
    card(s, MX, y, cw, h, C.panel);
    pill(s, "처음 가설", MX + 0.35, y + 0.3, 1.5, C.faint, 12);
    txt(s, "검사원마다 부식 판단이 달라\n'판단 보정' 도구가 필요하다", { x: MX + 0.35, y: y + 0.8, w: cw - 0.7, h: 0.9, fontSize: 18, bold: true, color: C.muted });
    iconDot(s, I.times, MX + 0.35, y + 1.9, 0.5, C.mat);
    txt(s, [{ text: "기각  ", options: { bold: true, color: C.mat } }, { text: "사내 매뉴얼과 인증 시험으로 이미 해결됨", options: { color: C.ink } }], { x: MX + 1.0, y: y + 1.9, w: cw - 1.3, h: 0.5, fontSize: 14, valign: "middle" });
    // arrow
    s.addImage({ data: I.arrowR, x: MX + cw + 0.17, y: y + h / 2 - 0.25, w: 0.5, h: 0.5 });
    // after
    const ax = MX + cw + 0.85, aw = W - MX - ax;
    card(s, ax, y, aw, h, C.rustLight, C.rust);
    pill(s, "현장이 원한 것", ax + 0.35, y + 0.3, 1.9, C.rust, 12);
    txt(s, "접근이 어려운 곳을 드론으로 먼저 훑어\n점검할 곳의 우선순위를 정해 달라", { x: ax + 0.35, y: y + 0.8, w: aw - 0.6, h: 0.9, fontSize: 18, bold: true, color: C.ink });
    iconDot(s, I.check, ax + 0.35, y + 1.9, 0.5, C.prop);
    txt(s, [{ text: "검증  ", options: { bold: true, color: C.prop } }, { text: "리프트 · 로프 액세스 · 선박 일수 절감", options: { color: C.ink } }], { x: ax + 1.0, y: y + 1.9, w: aw - 1.3, h: 0.5, fontSize: 14, valign: "middle" });
    // numbers
    const nums = [
      ["11 / 11", "현장 인력 모두 '판단 보정' 가설 기각", C.muted],
      ["10 / 11", "현장 인력이 '드론 사전 스캔' 필요성에 공감", C.rust],
      ["5 / 5", "의사결정권자 관심 · 파일럿 조건까지 제시", C.rust],
    ];
    const nw = (W - 2 * MX - 0.6) / 3;
    nums.forEach(([v, l, col], i) => {
      const x = MX + i * (nw + 0.3);
      txt(s, v, { x, y: 4.65, w: nw, h: 0.75, fontSize: 34, bold: true, color: col, valign: "middle" });
      txt(s, l, { x, y: 5.4, w: nw, h: 0.4, fontSize: 13.5, color: C.ink });
    });
    txt(s, "현장 인력 = MISTRAS · Williams 부식 엔지니어 · 검사원 11명 / 의사결정권자 = 해양·에너지 분야 CEO · 대표 5명", { x: MX, y: 6.15, w: W - 2 * MX, h: 0.3, fontSize: 11, color: C.faint });
    note(s,
      "[약 35초]\n저희의 처음 가설은 '검사원마다 판단이 달라서 보정 도구가 필요하다'였습니다. 그런데 현장 인력 11명 모두 아니라고 했습니다. 매뉴얼과 인증 시험으로 이미 해결된 문제였습니다. 대신 Williams의 한 관리자가 방향을 제시했습니다. '접근하기 어려운 곳을 드론으로 먼저 훑어서, 정말 가까이 봐야 할 곳만 알려 주면 시간과 비용이 줄어든다.' 이후 인터뷰에서 현장 인력 11명 중 10명이 여기에 공감했고, 의사결정권자 5명 모두 관심을 보이며 파일럿 조건까지 이야기했습니다.");
  }

  // ===== 5. Solution =====
  {
    const s = base("찍고, 녹 색으로 위험도를 매기고, 순서를 정합니다", "SOLUTION");
    const sw = hex(...RGB[5]);
    const steps = [
      { t: "드론 · 카메라 사전 스캔", d: "접근이 어려운 비말대 ·\n고소 구간을 먼저 촬영", img: "photo_15.png" },
      { t: "녹 영역 자동 추출", d: "녹 픽셀만 골라내\n도장면 · 맨 강재 간섭 제거", img: "mask_15.png" },
      { t: "색 분석 → 위험도", d: "녹의 노란 정도(b*)로\n부식 속도 구간 판정", color: sw },
      { t: "점검 우선순위 리포트", d: "가까이 볼 곳부터\n사람 · 배 · 리프트 배치", list: true },
    ];
    const cw = 2.55, gap = (W - 2 * MX - 4 * cw) / 3, y = 1.6;
    steps.forEach((st, i) => {
      const x = MX + i * (cw + gap);
      if (st.img) s.addImage({ path: IMG(st.img), x, y, w: cw, h: cw });
      else if (st.color) {
        s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: cw, fill: { color: st.color }, line: { color: C.line } });
        txt(s, "b* 12.2", { x, y: y + 0.55, w: cw, h: 0.7, fontSize: 30, bold: true, color: C.white, align: "center", valign: "middle" });
        pill(s, "위험도  중간", x + 0.45, y + 1.5, cw - 0.9, C.prop, 13, 0.42);
      } else {
        card(s, x, y, cw, cw, C.panel);
        [["A-07", "높음", "이번 주 점검", 2], ["C-12", "중간", "이번 주기 내", 1], ["B-03", "낮음", "다음 주기", 0]].forEach(([id, lv, act, k], j) => {
          const yy = y + 0.25 + j * 0.75, sg = STAGE[k];
          s.addShape("roundRect", { x: x + 0.18, y: yy, w: cw - 0.36, h: 0.62, fill: { color: j === 0 ? sg.col : C.white }, line: { color: sg.col, width: 1 }, rectRadius: 0.06 });
          txt(s, [
            { text: `${j + 1}. ${id}  `, options: { bold: true, color: j === 0 ? C.white : C.ink } },
            { text: lv, options: { bold: true, color: j === 0 ? C.white : sg.col, breakLine: true } },
            { text: act, options: { fontSize: 10.5, color: j === 0 ? C.white : C.muted } },
          ], { x: x + 0.32, y: yy, w: cw - 0.6, h: 0.62, fontSize: 13, valign: "middle" });
        });
      }
      s.addShape(pres.shapes.OVAL, { x: x - 0.18, y: y - 0.18, w: 0.5, h: 0.5, fill: { color: C.rust }, line: { color: C.white, width: 2 } });
      txt(s, String(i + 1), { x: x - 0.18, y: y - 0.18, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
      txt(s, st.t, { x, y: y + cw + 0.15, w: cw + 0.3, h: 0.4, fontSize: 16, bold: true });
      txt(s, st.d, { x, y: y + cw + 0.56, w: cw + 0.3, h: 0.65, fontSize: 12.5, color: C.muted });
      if (i < 3) s.addImage({ data: I.arrow, x: x + cw + gap / 2 - 0.13, y: y + cw / 2 - 0.13, w: 0.26, h: 0.26 });
    });
    const chips = ["점검 주기는 그대로, 자원 배치만 최적화", "로프 액세스 출동 여부 판단", "개인 판단이 아닌 데이터 리포트"];
    const chw = (W - 2 * MX - 0.3) / 3;
    chips.forEach((c, i) => pill(s, c, MX + i * (chw + 0.15), 6.2, chw, C.ink, 13, 0.48));
    txt(s, "예시 화면 (개념도)", { x: W - MX - 2.55, y: 1.15, w: 2.55, h: 0.25, fontSize: 10, italic: true, color: C.faint, align: "right" });
    note(s,
      "[약 30초]\n그래서 MLC는 이렇게 작동합니다. 드론이나 카메라로 접근이 어려운 구간을 먼저 찍습니다. 녹이 있는 부분만 자동으로 골라내고, 그 녹이 얼마나 노란지로 부식 속도 구간을 매깁니다. 마지막으로 어느 지점부터 가까이 봐야 하는지 순서를 정한 리포트를 드립니다. 규제 주기는 그대로 지키면서, 사람과 배와 리프트를 위험한 곳에 먼저 보내는 겁니다.");
  }

  // ===== 6. Validation =====
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
    txt(s, "위험도 낮음", { x: px(-0.2) - 0.05, y: py(0.447) - 0.42, w: 1.6, h: 0.3, fontSize: 13, bold: true, color: C.inc });
    txt(s, "중간", { x: px(6.6) - 0.05, y: py(0.838) - 0.42, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.prop });
    txt(s, "높음", { x: px(12.1) - 1.3, y: py(1.745) - 0.05, w: 1.15, h: 0.3, fontSize: 13, bold: true, color: C.mat, align: "right" });
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
      "[약 30초]\n핵심 기술은 실험실에서 검증했습니다. 탄소강을 240시간 촉진부식 시키며 아홉 시점에서 측정했는데, 녹의 노란 정도 b*가 높아질수록 실제 부식 속도도 올라갔고 순위 상관계수는 0.88이었습니다. 위험도 세 구간이 서로 겹치지 않게 나뉩니다. 노란 녹인 레피도크로사이트가 다공성이라 부식을 빠르게 만들기 때문에, 색과 속도가 함께 움직이는 물리적 근거도 있습니다.");
  }

  // ===== 7. Business model & ecosystem =====
  {
    const s = base("누가 쓰고, 누가 결정하고, 어떻게 돈을 버나", "BUSINESS MODEL");
    // ecosystem (left)
    const lx = MX, lw = 6.1;
    txt(s, "고객 생태계 · 해양 석유·가스 운영사", { x: lx, y: 1.45, w: lw, h: 0.3, fontSize: 13, bold: true, color: C.sea });
    const eco = [
      [I.crown, "결정", "CEO · 대표", "안전 · 비용 · 규제 대응 중 하나가 결정 동기", C.rust],
      [I.star, "추천", "부식 엔지니어링 매니저", "기술 검토 후 대표에게 추천", C.sea],
      [I.user, "사용", "현장 엔지니어 · 검사원", "리포트로 점검 동선과 출동 여부 결정", C.sea],
      [I.users, "영향", "방식(防蝕) 전문 업체 · 연구자", "기술 교차 검증 · 고객 소개", C.muted],
    ];
    eco.forEach(([ic, tag, who, d, col], i) => {
      const y = 1.9 + i * 1.08;
      card(s, lx, y, lw, 0.95);
      iconDot(s, ic, lx + 0.2, y + 0.16, 0.63, col);
      pill(s, tag, lx + 1.0, y + 0.3, 0.75, col, 11, 0.34);
      txt(s, who, { x: lx + 1.95, y: y + 0.1, w: lw - 2.1, h: 0.38, fontSize: 14.5, bold: true, valign: "middle" });
      txt(s, d, { x: lx + 1.95, y: y + 0.5, w: lw - 2.1, h: 0.35, fontSize: 12, color: C.muted });
    });
    // BM (right)
    const rx = lx + lw + 0.45, rw = W - MX - rx;
    const bm = [
      ["가치", "드론 사전 스캔으로 위험도 선별 → 규제 주기 안에서 자원 배치 최적화, 잔여 수명(RUL) 예측"],
      ["수익", "점검 건당 요금 또는 구독형 서비스 요금"],
      ["채널", "점검 프로그램 매니저 직접 영업 · 방식 전문 업체 추천"],
      ["파트너", "방식 전문 업체 (예: Farwest Corrosion Control) · 드론 하드웨어 업체"],
    ];
    bm.forEach(([k, v], i) => {
      const y = 1.45 + i * 0.95;
      txt(s, k, { x: rx, y, w: 0.95, h: 0.85, fontSize: 14, bold: true, color: C.rust, valign: "middle" });
      txt(s, v, { x: rx + 1.0, y, w: rw - 1.0, h: 0.85, fontSize: 13, valign: "middle" });
      s.addShape(pres.shapes.LINE, { x: rx, y: y + 0.9, w: rw, h: 0, line: { color: C.line, width: 1 } });
    });
    todo(s, "가격 가설 (예: 구조물 1기 스캔당 $○○ / 연 구독 $○○) — 수익 모델은 아직 검증 중", rx, 5.4, rw, 0.95, 12.5);
    note(s,
      "[약 25초]\n고객은 해양 석유·가스 운영사입니다. 현장 검사원이 리포트를 쓰고, 부식 엔지니어링 매니저가 추천하고, 대표가 결정합니다. 대표마다 결정 동기가 달랐는데, 안전, 비용, 규제와 보험 대응 세 가지였습니다. 수익은 점검 건당 요금이나 구독형으로 생각하고 있고, 방식 전문 업체와 드론 업체를 파트너이자 소개 채널로 봅니다. 가격은 다음 단계에서 검증합니다.");
  }

  // ===== 8. Market =====
  {
    const s = base("규제가 만드는 반복 수요", "MARKET");
    const lv = [
      ["Level I", "수상 육안 점검", "도장 · 구조 · 비말대 육안 스캔", "매년", C.rust],
      ["Level II", "수중 일반 점검", "잠수부 · ROV 조사", "3–5년", C.sea],
      ["Level III", "수중 정밀 점검", "지정 부위 세척 후 근접 육안 · 피트 계측", "6–10년", C.ink],
    ];
    const cw = (W - 2 * MX - 0.5) / 3;
    lv.forEach(([lvl, t, d, cyc, col], i) => {
      const x = MX + i * (cw + 0.25), y = 1.5;
      card(s, x, y, cw, 1.75);
      txt(s, lvl, { x: x + 0.3, y: y + 0.18, w: 2, h: 0.35, fontSize: 14, bold: true, color: col });
      txt(s, cyc, { x: x + cw - 1.8, y: y + 0.1, w: 1.5, h: 0.5, fontSize: 22, bold: true, color: col, align: "right" });
      txt(s, t, { x: x + 0.3, y: y + 0.65, w: cw - 0.6, h: 0.4, fontSize: 16, bold: true });
      txt(s, d, { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 0.5, fontSize: 12.5, color: C.muted });
    });
    const regs = ["BSEE", "PHMSA", "USCG", "ABS"];
    txt(s, "미국 해양 구조물 점검 체계 — 4개 기관이 겹쳐 규제", { x: MX, y: 3.42, w: 6, h: 0.3, fontSize: 12, color: C.muted });
    regs.forEach((r, i) => pill(s, r, MX + 5.3 + i * 1.05, 3.4, 0.95, C.sea, 11, 0.34));
    // TAM/SAM/SOM
    const ry = 4.1;
    const mk = [
      ["TAM", "글로벌 부식 모니터링 시장 — USD 1.1–2.9B (2025, 보고서별 상이): 1개 확정"],
      ["SAM", "미국 멕시코만 해양 플랫폼 · 미드스트림 외부 부식 점검 시장"],
      ["SOM", "첫 3년: 파일럿 고객 수 × 연간 스캔 구조물 수 × 단가"],
    ];
    mk.forEach(([k, d], i) => {
      const y = ry + i * 0.62;
      txt(s, k, { x: MX, y, w: 0.9, h: 0.52, fontSize: 16, bold: true, color: C.rust, valign: "middle" });
      todo(s, d, MX + 0.95, y, 7.2, 0.52, 11.5);
    });
    // domestic extension
    const dx = MX + 8.45, dw = W - MX - dx;
    card(s, dx, ry, dw, 1.76, C.seaLight);
    txt(s, "국내 확장 후보 (검증 예정)", { x: dx + 0.25, y: ry + 0.15, w: dw - 0.4, h: 0.3, fontSize: 12.5, bold: true, color: C.sea });
    [[I.wind, "해상풍력 하부 구조"], [I.industry, "해양 플랜트 · 조선 · 항만"]].forEach(([ic, t], i) => {
      const y = ry + 0.6 + i * 0.55;
      iconDot(s, ic, dx + 0.25, y, 0.42, C.sea);
      txt(s, t, { x: dx + 0.8, y, w: dw - 0.9, h: 0.42, fontSize: 13, valign: "middle" });
    });
    note(s,
      "[약 25초]\n이 시장의 특징은 규제가 반복 수요를 만든다는 점입니다. 미국 해양 플랫폼은 매년 수상 육안 점검, 3년에서 5년마다 수중 점검, 6년에서 10년마다 정밀 점검을 해야 하고, 네 개 기관이 이를 겹쳐서 규제합니다. [TAM·SAM·SOM 수치를 한 문장으로.] 미국에서 먼저 검증한 뒤, 국내 해상풍력과 해양 플랜트로 넓힐 계획입니다.");
  }

  // ===== 9. Traction & next =====
  {
    const s = base("지금까지, 그리고 다음 단계", "TRACTION  ·  ROADMAP");
    const lw = 4.6;
    txt(s, "지금까지", { x: MX, y: 1.5, w: lw, h: 0.35, fontSize: 16, bold: true, color: C.rust });
    const done = [
      ["실험실 기술 검증", "240시간 촉진부식, 6가지 분석 교차 검증"],
      ["미국 고객 인터뷰 35건", "Tex-Corps · 2026년 7월 · 타깃 확정"],
      ["국제학회 발표", "ICIMC 2026 구두 발표 (2026.11, 서울)"],
    ];
    done.forEach(([t, d], i) => {
      const y = 2.0 + i * 1.0;
      iconDot(s, I.check, MX, y + 0.05, 0.5, C.prop);
      txt(s, t, { x: MX + 0.7, y, w: lw - 0.7, h: 0.35, fontSize: 15, bold: true });
      txt(s, d, { x: MX + 0.7, y: y + 0.38, w: lw - 0.7, h: 0.55, fontSize: 12, color: C.muted });
    });
    todo(s, "MOU · 수상 · 특허 출원 등 추가 성과", MX, 4.95, lw, 0.6, 12);
    // next steps
    const rx = MX + lw + 0.5, rw = W - MX - rx;
    txt(s, "다음 단계", { x: rx, y: 1.5, w: rw, h: 0.35, fontSize: 16, bold: true, color: C.rust });
    const next = [
      [I.calc, "비용 절감 근거 정량화", "선박 일수 · 리프트/로프 액세스 임대 · 문서 작업 시간 절감액으로 ROI 산출"],
      [I.ruler, "두께 측정 데이터 결합", "사진이 못 보는 두께 감소를 보정해 예측 정확도 향상"],
      [I.map, "유망 현장 심화", "미드스트림 압축 기지 · 비말대, 라이저 지지 프레임이 있는 해양 플랫폼"],
      [I.flask, "단일 현장 파일럿", "범위 · 기간 한정, 성공 지표 사전 합의 (의사결정권자 요구 조건)"],
    ];
    next.forEach(([ic, t, d], i) => {
      const y = 1.98 + i * 0.98;
      card(s, rx, y, rw, 0.86, i === 3 ? C.rustLight : C.panel);
      iconDot(s, ic, rx + 0.18, y + 0.15, 0.56, i === 3 ? C.rust : C.ink);
      txt(s, t, { x: rx + 0.95, y: y + 0.08, w: rw - 1.1, h: 0.35, fontSize: 14.5, bold: true });
      txt(s, d, { x: rx + 0.95, y: y + 0.44, w: rw - 1.1, h: 0.35, fontSize: 12, color: C.muted });
    });
    todo(s, "각 단계 목표 시점 (예: 파일럿 2027 상반기)", rx, 5.95, rw, 0.55, 12);
    note(s,
      "[약 30초]\n지금까지 실험실 검증을 마쳤고, 미국에서 35명을 만나 타깃을 해양 석유·가스로 좁혔으며, 다음 달 국제학회에서 발표합니다. 다음 단계는 네 가지입니다. 선박 일수와 리프트 임대처럼 실제 절감액을 숫자로 만들고, 사진이 볼 수 없는 두께 감소를 보완하기 위해 두께 측정 데이터와 결합하고, 반응이 가장 좋았던 현장을 깊이 파고, 의사결정권자들이 말한 조건대로 범위와 기간을 정한 단일 현장 파일럿을 하겠습니다.");
  }

  // ===== 10. Team =====
  {
    const s = base("팀 MLC", "TEAM");
    const people = [
      ["team_woo.png", 413 / 531, "우성훈", "EL (Entrepreneurial Lead)", "녹 색상 기반 부식 예측 연구 · ICIMC 2026 발표자"],
      ["team_park.jpg", 689 / 886, "박준범", "EM", "부식 연구 공동 저자 · 고객 인터뷰 35건 공동 진행"],
    ];
    const cw = 5.85;
    people.forEach(([img, ar, n, r, d], i) => {
      const x = MX + i * (cw + 0.43), y = 1.55;
      card(s, x, y, cw, 3.3, i === 0 ? C.rustLight : C.panel);
      const ph = 2.7, pw = ph * ar;
      s.addImage({ path: IMG(img), x: x + 0.3, y: y + 0.3, w: pw, h: ph, sizing: { type: "cover", w: pw, h: ph } });
      txt(s, n, { x: x + pw + 0.6, y: y + 0.45, w: cw - pw - 0.8, h: 0.5, fontSize: 24, bold: true });
      txt(s, r, { x: x + pw + 0.6, y: y + 1.0, w: cw - pw - 0.8, h: 0.35, fontSize: 13.5, bold: true, color: C.rust });
      txt(s, "경상국립대학교\n기계시스템공학과", { x: x + pw + 0.6, y: y + 1.45, w: cw - pw - 0.8, h: 0.6, fontSize: 12.5, color: C.muted });
      txt(s, d, { x: x + pw + 0.6, y: y + 2.15, w: cw - pw - 0.8, h: 0.8, fontSize: 12.5 });
    });
    todo(s, "지도교수 · 멘토 · 자문 (예: 부식/재료 분야 지도교수, Tex-Corps 멘토)", MX, 5.15, W - 2 * MX, 0.6, 13);
    txt(s, "연구 기반: 해양수산부 「친환경 선박 블루테크 리더 양성」 사업 (RS-2025-02220459)", { x: MX, y: 6.05, w: W - 2 * MX, h: 0.35, fontSize: 12, color: C.muted });
    note(s,
      "[약 15초]\n팀 MLC는 경상국립대 기계시스템공학과의 우성훈, 박준범입니다. 저희 두 사람이 부식 연구를 직접 했고, 미국 인터뷰 35건도 함께 진행했습니다.");
  }

  // ===== 11. Ask & close =====
  {
    const s = pres.addSlide(); pageNo += 1;
    s.background = { color: C.dark };
    txt(s, "ASK", { x: MX, y: 0.55, w: 4, h: 0.3, fontSize: 12, bold: true, color: C.rustSoft, charSpacing: 2 });
    txt(s, "첫 파일럿 현장을 찾고 있습니다", { x: MX, y: 0.9, w: 12, h: 0.8, fontSize: 34, bold: true, color: C.white, valign: "middle" });
    const asks = [
      [I.handshake, "파일럿 파트너", "해양 구조물 · 플랜트 운영사", "단일 현장, 기간 한정, 성공 지표 사전 합의"],
      [I.helicopter, "드론 · 점검 파트너", "드론 운용사 · 방식 전문 업체", "현장 촬영 데이터 확보와 기술 교차 검증"],
      [I.coins, "투자 · 지원", null, "금액과 사용처 (예: 파일럿 ○○%, 두께 데이터 결합 개발 ○○%)"],
    ];
    const cw = (W - 2 * MX - 0.6) / 3;
    asks.forEach(([ic, t, who, d], i) => {
      const x = MX + i * (cw + 0.3), y = 2.0;
      s.addShape("roundRect", { x, y, w: cw, h: 2.6, fill: { color: "2B2F35" }, line: { color: "3A3D42" }, rectRadius: 0.08 });
      iconDot(s, ic, x + 0.3, y + 0.3, 0.75, C.rust);
      txt(s, t, { x: x + 1.25, y: y + 0.3, w: cw - 1.5, h: 0.75, fontSize: 19, bold: true, color: C.white, valign: "middle" });
      if (who) {
        txt(s, who, { x: x + 0.3, y: y + 1.3, w: cw - 0.6, h: 0.4, fontSize: 14, bold: true, color: C.rustSoft });
        txt(s, d, { x: x + 0.3, y: y + 1.75, w: cw - 0.6, h: 0.7, fontSize: 13, color: "D5D7DA" });
      } else todo(s, d, x + 0.3, y + 1.3, cw - 0.6, 1.05, 12.5);
    });
    txt(s, "어디부터 점검할지, 녹의 색이 먼저 알려드립니다", { x: MX, y: 5.0, w: 9, h: 0.55, fontSize: 22, bold: true, color: C.rustSoft });
    s.addImage({ data: I.mail, x: MX, y: 5.78, w: 0.3, h: 0.3 });
    txt(s, "MLC  ·  우성훈  ·  zscfvgbb@naver.com", { x: MX + 0.45, y: 5.72, w: 8, h: 0.4, fontSize: 15, color: C.white, valign: "middle" });
    CYC.forEach((c, i) => s.addShape(pres.shapes.RECTANGLE, { x: 9.7 + (i % 3) * 0.95, y: 4.95 + Math.floor(i / 3) * 0.55, w: 0.85, h: 0.45, fill: { color: hex(...RGB[i]) }, line: { color: "3A3D42", width: 1 } }));
    txt(s, "감사합니다", { x: MX, y: 6.4, w: 6, h: 0.5, fontSize: 18, bold: true, color: C.white });
    note(s,
      "[약 20초]\n저희에게 지금 가장 필요한 것은 첫 파일럿 현장입니다. 범위와 기간을 정하고 성공 지표를 미리 합의하는 작은 파일럿을 함께할 운영사, 그리고 드론 파트너를 찾고 있습니다. [투자·지원 요청을 한 문장으로.] 어디부터 점검할지, 녹의 색이 먼저 알려드리겠습니다. 감사합니다.");
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

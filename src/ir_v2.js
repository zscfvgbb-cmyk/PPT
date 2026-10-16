// MLC IR pitch v2 — same content as MLC_IR_A-STREAM2026.pptx, laid out after the
// structure of another A-STREAM team's deck: numbered section headers, a key-message
// banner at the foot of each slide, a process band for the solution, a partner →
// MLC → customer flow, a staged timeline, and a team + lab-infrastructure slide.
// Build: node ir_v2.js  -> ../MLC_IR_A-STREAM2026_v2.pptx
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const path = require("path");

const IMG = (f) => path.join(__dirname, "img", f);
const OUT = process.argv[2] || path.join(__dirname, "..", "MLC_IR_A-STREAM2026_v2.pptx");

const C = {
  navy: "1B3A5C", navyMid: "6F8FB5", navyLight: "B9CCE0", sea: "1F5F7A", seaLight: "E3EFF4",
  ink: "2B2D31", muted: "6B6F76", faint: "A3A7AD", line: "D9DCE0", panel: "F3F5F8", white: "FFFFFF", paper: "F7F9FB",
  rust: "B5501F", rustLight: "FBEDE4", rustSoft: "F2B48A", dark: "2E3238", band: "4A4F57",
  inc: "3A6EA5", prop: "3C8D5A", mat: "B23A2E", incL: "E6EEF7", propL: "E5F2EA", matL: "F8E4E2",
};
const F = "맑은 고딕";
const W = 13.333, H = 7.5, MX = 0.6;
const RGB = [[79,79,80],[72,71,68],[76,75,71],[72,71,68],[72,65,56],[78,63,47],[78,60,44],[78,60,43],[73,60,44]];
const hex = (a) => a.map((v) => Math.round(v).toString(16).padStart(2, "0")).join("").toUpperCase();
const avg = (ix) => [0, 1, 2].map((k) => ix.reduce((s, i) => s + RGB[i][k], 0) / ix.length);

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "MLC — A-STREAM 2026 IR Pitch";
pres.author = "MLC (우성훈, 박준범)";

function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: F, fontSize: 14, color: C.ink, margin: 0, isTextBox: true, valign: "top" }, o));
}
function rr(s, x, y, w, h, fill, line, r = 0.08) {
  s.addShape("roundRect", { x, y, w, h, fill: { color: fill }, line: line ? { color: line, width: 1 } : { color: fill }, rectRadius: r });
}
function dot(s, x, y, d, col) { s.addShape("ellipse", { x, y, w: d, h: d, fill: { color: col }, line: { color: col } }); }
function iconDot(s, data, x, y, d, bg) { dot(s, x, y, d, bg); const p = d * 0.26; s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p }); }
function pill(s, t, x, y, w, color, fs = 11, h = 0.32, fg = C.white) {
  s.addShape("roundRect", { x, y, w, h, fill: { color }, line: { color }, rectRadius: h / 2 });
  txt(s, t, { x, y, w, h, fontSize: fs, bold: true, color: fg, align: "center", valign: "middle" });
}
function dots(s, x, y, rev) {
  const cols = rev ? [C.navyLight, C.navyMid, C.navy] : [C.navy, C.navyMid, C.navyLight];
  cols.forEach((c, i) => dot(s, x + i * 0.16, y, 0.085, c));
}

let page = 0;
// Numbered section header: "01" (light/dark digits) + title, kicker above, page number at right.
function slide(num, kicker, title) {
  const s = pres.addSlide(); page += 1;
  s.background = { color: C.white };
  dots(s, MX, 0.42);
  txt(s, kicker, { x: MX + 0.6, y: 0.33, w: 6, h: 0.28, fontSize: 11, color: C.navy });
  txt(s, [{ text: num[0], options: { color: C.navyLight } }, { text: num[1], options: { color: C.navy } }], { x: MX - 0.05, y: 0.68, w: 1.0, h: 0.75, fontSize: 40, bold: true, valign: "middle" });
  txt(s, title, { x: MX + 1.0, y: 0.68, w: 10.4, h: 0.75, fontSize: 26, bold: true, valign: "middle" });
  dots(s, W - MX - 0.42, 0.42, true);
  s.addShape("line", { x: W - MX - 0.5, y: 0.86, w: 0, h: 0.34, line: { color: C.line, width: 1 } });
  txt(s, String(page), { x: W - MX - 0.42, y: 0.86, w: 0.42, h: 0.34, fontSize: 12, bold: true, color: C.navy, align: "right", valign: "middle" });
  return s;
}
// Key-message banner along the foot of the slide; runs marked hl are highlighted.
function banner(s, runs) {
  s.addShape("rect", { x: 0, y: 6.72, w: W, h: 0.78, fill: { color: C.navy }, line: { color: C.navy } });
  txt(s, runs.map(([t, hl]) => ({ text: t, options: { color: hl ? C.rustSoft : C.white, bold: !!hl } })), { x: MX, y: 6.72, w: W - 2 * MX, h: 0.78, fontSize: 19, align: "center", valign: "middle" });
}
function note(s, sec, body) { s.addNotes(`[약 ${sec}초]\n${body}`); }

const STG = [
  { name: "낮음", col: C.inc, img: "photo_3.png", time: "24 h", b: 1.7, cr: 0.34, rgb: avg([0, 1, 2, 3]) },
  { name: "중간", col: C.prop, img: "photo_15.png", time: "120 h", b: 9.4, cr: 0.68, rgb: avg([4, 5]) },
  { name: "높음", col: C.mat, img: "photo_30.png", time: "240 h", b: 13.1, cr: 1.45, rgb: avg([6, 7, 8]) },
];

(async () => {
  const ic = async (n, c = C.white) => icon(fa[n], c);
  const I = {
    coins: await ic("FaCoins"), smog: await ic("FaSmog"), steel: await ic("FaIndustry"),
    route: await ic("FaRoute"), heli: await ic("FaHelicopter"), ship: await ic("FaShip"),
    check: await ic("FaCheck"), times: await ic("FaTimes"), arrowW: await ic("FaArrowRight"),
    arrowR: await ic("FaArrowRight", C.rust), arrowN: await ic("FaArrowRight", C.navyMid),
    camera: await ic("FaCamera"), map: await ic("FaMapMarkedAlt"), shield: await ic("FaShieldAlt"),
    gavel: await ic("FaBalanceScale"), tools: await ic("FaTools"), file: await ic("FaFileAlt"),
    crown: await ic("FaUserTie"), star: await ic("FaUserCog"), hat: await ic("FaHardHat"), users: await ic("FaUsers"),
    fog: await ic("FaCloudShowersHeavy", C.navy), micro: await ic("FaMicroscope", C.navy), wave: await ic("FaWaveSquare", C.navy),
    bolt: await ic("FaBolt", C.navy), layers: await ic("FaLayerGroup", C.navy), cam: await ic("FaCamera", C.navy),
    mail: await ic("FaEnvelope", C.navy),
  };

  // ===== 1. Cover =====
  {
    const s = pres.addSlide(); page += 1;
    s.background = { color: C.paper };
    dots(s, MX + 0.3, 0.75);
    txt(s, "AI 기반 금속 열화 예측 플랫폼", { x: MX + 0.3, y: 1.55, w: 6.5, h: 0.45, fontSize: 19, bold: true, color: C.navy });
    txt(s, [{ text: "ML", options: { color: C.navy } }, { text: "C", options: { color: C.rust } }], { x: MX + 0.25, y: 2.0, w: 6, h: 1.3, fontSize: 88, bold: true, valign: "middle" });
    txt(s, "녹의 색으로 읽는\n해양 구조물 부식 위험도", { x: MX + 0.3, y: 3.85, w: 6.5, h: 0.95, fontSize: 21, color: C.ink, lineSpacingMultiple: 1.1 });
    txt(s, "우성훈, 창업팀 대표", { x: MX + 0.3, y: 4.9, w: 6, h: 0.35, fontSize: 14, color: C.muted });
    txt(s, "경상국립대학교  ·  Tex-Corps 제주 A-STREAM 2026", { x: MX + 0.3, y: 5.75, w: 7, h: 0.4, fontSize: 15, bold: true, color: C.navy });
    const pw = 5.3, ph = pw * 284 / 386, px = W - MX - pw - 0.2, py = 1.35;
    s.addImage({ path: IMG("offshore_legs.png"), x: px, y: py, w: pw, h: ph, rounding: false });
    const n = 9, sw = (pw - 0.08 * (n - 1)) / n;
    RGB.forEach((c, i) => s.addShape("rect", { x: px + i * (sw + 0.08), y: py + ph + 0.2, w: sw, h: 0.42, fill: { color: hex(c) }, line: { color: C.white, width: 0.75 } }));
    txt(s, "녹 색의 변화, 0 → 240시간", { x: px, y: py + ph + 0.68, w: pw, h: 0.28, fontSize: 10.5, color: C.muted, align: "right" });
    dots(s, W - MX - 0.62, 0.75, true);
    note(s, 15, "안녕하십니까, MLC 우성훈입니다. 해양 구조물을 점검할 때 가장 어려운 질문은 '어디부터 봐야 하나'입니다. 저희는 녹의 색을 읽어서 그 답을 먼저 드리는 기술을 만들고 있습니다.");
  }

  // ===== 2. Problem 1 =====
  {
    const s = slide("01", "문제점(Problem)", "부식이 남기는 경제적 · 환경적 비용");
    const cards = [
      ["경제적 손실", C.navy, "연간 전 세계 부식 비용", "US$ 2.5조", "세계 GDP의 약 3.4%", "NACE IMPACT (2016) · AMPP 2025 인용", I.coins],
      ["탄소 배출", C.sea, "부식 강재 교체용 제강", "1.6–3.4%", "전 세계 CO₂ 배출 중 비중\n2030년 4.1–9.1% 전망", "Iannuzzi & Frankel (2022)", I.smog],
      ["자원 낭비", C.rust, "부식 손실 보충", "25–33%", "매년 생산되는 철강 중\n부식 손실을 메우는 데 쓰이는 비율", "Iannuzzi & Frankel (2022)", I.steel],
    ];
    const cw = 3.55, gap = (W - 2 * MX - 3 * cw) / 2, y = 1.95, h = 4.4;
    cards.forEach(([tab, col, lab, big, d, src, icn], i) => {
      const x = MX + i * (cw + gap);
      rr(s, x, y, cw, h, C.white, C.line);
      pill(s, tab, x + 0.45, y - 0.24, cw - 0.9, col, 15, 0.48);
      rr(s, x + 0.25, y + 0.45, cw - 0.5, 0.48, C.panel, null, 0.06);
      rr(s, x + 0.25, y + 0.45, 0.48, 0.48, col, null, 0.06);
      txt(s, String(i + 1), { x: x + 0.25, y: y + 0.45, w: 0.48, h: 0.48, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
      txt(s, lab, { x: x + 0.85, y: y + 0.45, w: cw - 1.2, h: 0.48, fontSize: 14, bold: true, color: col, valign: "middle" });
      iconDot(s, icn, x + cw / 2 - 0.42, y + 1.15, 0.84, col);
      txt(s, big, { x: x + 0.2, y: y + 2.05, w: cw - 0.4, h: 0.8, fontSize: 36, bold: true, color: col, align: "center", valign: "middle" });
      txt(s, d, { x: x + 0.25, y: y + 2.9, w: cw - 0.5, h: 0.8, fontSize: 13, color: C.ink, align: "center" });
      txt(s, src, { x: x + 0.2, y: y + h - 0.42, w: cw - 0.4, h: 0.28, fontSize: 9.5, italic: true, color: C.faint, align: "center" });
      if (i < 2) s.addImage({ data: I.arrowN, x: x + cw + gap / 2 - 0.16, y: y + h / 2 - 0.16, w: 0.32, h: 0.32 });
    });
    banner(s, [["제때 찾아 관리하면 ", 0], ["부식 비용의 15–35%", 1], [" (연 US$ 3,750–8,750억)를 줄일 수 있습니다", 0]]);
    note(s, 30, "부식은 생각보다 큰 문제입니다. 전 세계가 매년 부식으로 잃는 비용은 약 2.5조 달러, 세계 GDP의 3.4%이고, 국제 부식 학회 AMPP가 2025년에도 공식 인용하는 수치입니다. 매년 생산되는 철강의 4분의 1에서 3분의 1이 부식 손실을 메우는 데 쓰이고, 이 강재를 다시 만드는 과정에서 전 세계 이산화탄소 배출의 1.6~3.4%가 나옵니다. 반대로 제때 찾아 관리하면 이 비용의 15~35%, 연간 3,750억에서 8,750억 달러를 줄일 수 있습니다.");
  }

  // ===== 3. Problem 2 =====
  {
    const s = slide("02", "문제점(Problem)", "해양 구조물 점검의 한계: 접근");
    const iw = 4.55;
    s.addImage({ path: IMG("offshore_zones.png"), x: MX, y: 1.7, w: iw, h: iw * 678 / 681 });
    const rx = MX + iw + 0.4, rw = W - MX - rx;
    const pains = [
      ["1", "모든 지점을 똑같이 순회", "상태와 관계없이 정해진 목록을 전부 방문", "Williams · 점검 관리자", I.route],
      ["2", "비말대 · 해저 구간은 접근 자체가 어려움", "로프 액세스 · 잠수 · ROV를 미리 일정 잡아야만 확인 가능", "Deepwater Corrosion Services · CTO", I.heli],
      ["3", "비용은 '검사'가 아니라 '이동 · 대기'", "선박 용선일 · 기상 대기 · 리프트 · 비계 임대가 실제 비용", "Deepwater · Brown Corrosion · MISTRAS", I.ship],
    ];
    pains.forEach(([n, t, d, who, icn], i) => {
      const y = 1.7 + i * 1.53;
      rr(s, rx, y, rw, 1.35, C.panel);
      iconDot(s, icn, rx + 0.3, y + 0.3, 0.75, C.navy);
      txt(s, t, { x: rx + 1.3, y: y + 0.17, w: rw - 1.5, h: 0.4, fontSize: 17, bold: true });
      txt(s, d, { x: rx + 1.3, y: y + 0.58, w: rw - 1.5, h: 0.35, fontSize: 13.5, color: C.muted });
      txt(s, "— " + who, { x: rx + 1.3, y: y + 0.95, w: rw - 1.5, h: 0.28, fontSize: 10.5, italic: true, color: C.faint });
    });
    banner(s, [["점검 주기는 규제로 고정 — 문제는 ", 0], ["어디에 사람과 배를 먼저 보내느냐", 1], ["입니다", 0]]);
    note(s, 28, "특히 해양 구조물이 그렇습니다. 파도가 치는 비말대는 부식이 가장 빠른 곳인데, 로프 액세스나 잠수, ROV 없이는 볼 수가 없습니다. 지금은 상태와 상관없이 모든 지점을 똑같이 돌고, 비용의 대부분은 검사 자체가 아니라 배를 빌리고 날씨를 기다리고 리프트를 빌리는 데서 나옵니다. 점검 주기는 규제로 정해져 있어 줄일 수 없습니다. 결국 문제는 그 주기 안에서 어디에 사람과 배를 먼저 보낼지입니다.");
  }

  // ===== 4. Customer discovery =====
  {
    const s = slide("03", "고객 검증(Customer Discovery)", "미국 현장 인터뷰 35건으로 찾은 진짜 수요");
    const cw = 5.55, y = 1.7, h = 2.55;
    rr(s, MX, y, cw, h, C.panel);
    pill(s, "처음 가설", MX + 0.35, y + 0.3, 1.5, C.faint, 12);
    txt(s, "검사원마다 부식 판단이 달라\n'판단 보정' 도구가 필요하다", { x: MX + 0.35, y: y + 0.8, w: cw - 0.7, h: 0.9, fontSize: 18, bold: true, color: C.muted });
    iconDot(s, I.times, MX + 0.35, y + 1.85, 0.48, C.mat);
    txt(s, [{ text: "기각  ", options: { bold: true, color: C.mat } }, { text: "사내 매뉴얼과 인증 시험으로 이미 해결됨" }], { x: MX + 1.0, y: y + 1.85, w: cw - 1.3, h: 0.48, fontSize: 14, valign: "middle" });
    s.addImage({ data: I.arrowR, x: MX + cw + 0.17, y: y + h / 2 - 0.25, w: 0.5, h: 0.5 });
    const ax = MX + cw + 0.85, aw = W - MX - ax;
    rr(s, ax, y, aw, h, C.rustLight, C.rust);
    pill(s, "현장이 원한 것", ax + 0.35, y + 0.3, 1.9, C.rust, 12);
    txt(s, "접근이 어려운 곳을 드론으로 먼저 훑어\n점검할 곳의 우선순위를 정해 달라", { x: ax + 0.35, y: y + 0.8, w: aw - 0.6, h: 0.9, fontSize: 18, bold: true });
    iconDot(s, I.check, ax + 0.35, y + 1.85, 0.48, C.prop);
    txt(s, [{ text: "검증  ", options: { bold: true, color: C.prop } }, { text: "리프트 · 로프 액세스 · 선박 일수 절감" }], { x: ax + 1.0, y: y + 1.85, w: aw - 1.3, h: 0.48, fontSize: 14, valign: "middle" });
    const nums = [
      ["35건", "미국 고객 인터뷰\n대면 24 · 화상 11 (2026.7)", C.navy],
      ["11 / 11", "현장 인력 모두\n'판단 보정' 가설 기각", C.muted],
      ["10 / 11", "현장 인력이 '드론 사전 스캔'\n필요성에 공감", C.rust],
      ["5 / 5", "의사결정권자 관심\n파일럿 조건까지 제시", C.rust],
    ];
    const nw = (W - 2 * MX - 0.6) / 4;
    nums.forEach(([v, l, col], i) => {
      const x = MX + i * (nw + 0.2);
      txt(s, v, { x, y: 4.55, w: nw, h: 0.75, fontSize: 32, bold: true, color: col, valign: "middle" });
      txt(s, l, { x, y: 5.3, w: nw, h: 0.65, fontSize: 13 });
    });
    txt(s, "현장 인력 = MISTRAS · Williams 부식 엔지니어 · 검사원 11명 / 의사결정권자 = 해양 · 에너지 분야 CEO · 대표 5명", { x: MX, y: 6.15, w: W - 2 * MX, h: 0.3, fontSize: 10.5, color: C.faint });
    banner(s, [["현장이 원한 것은 판단 보정이 아니라 ", 0], ["'어디부터 볼지'에 대한 답", 1], ["이었습니다", 0]]);
    note(s, 35, "저희는 올해 7월 Tex-Corps 프로그램으로 미국에서 35명을 인터뷰했습니다. 처음 가설은 '검사원마다 판단이 달라 보정 도구가 필요하다'였는데, 현장 인력 11명 모두 아니라고 했습니다. 매뉴얼과 인증 시험으로 이미 해결된 문제였습니다. 대신 접근하기 어려운 곳을 드론으로 먼저 훑어서 정말 가까이 봐야 할 곳만 알려 달라는 요구가 나왔고, 현장 인력 11명 중 10명이 공감했습니다. 의사결정권자 5명 모두 관심을 보이며 파일럿 조건까지 이야기했습니다.");
  }

  // ===== 5. Solution =====
  {
    const s = slide("04", "해결방안(Solution)", "녹의 색으로 점검 우선순위를 정하는 MLC");
    const by = 1.6, bh = 2.85;
    s.addShape("rect", { x: 0, y: by, w: W, h: bh, fill: { color: C.band }, line: { color: C.band } });
    const n = 4, iw = 1.75, cw = (W - 2 * MX) / n, iy = by + 0.22;
    const labels = ["드론 · 카메라 촬영", "녹 영역 자동 추출", "색(b*)으로 위험도 판정", "점검 우선순위 지도"];
    for (let i = 0; i < n; i++) {
      const cx = MX + i * cw + cw / 2, x = cx - iw / 2;
      if (i === 0) s.addImage({ path: IMG("photo_15.png"), x, y: iy, w: iw, h: iw });
      if (i === 1) s.addImage({ path: IMG("mask_15.png"), x, y: iy, w: iw, h: iw });
      if (i === 2) {
        s.addShape("rect", { x, y: iy, w: iw, h: iw, fill: { color: hex(RGB[5]) }, line: { color: C.white, width: 1 } });
        txt(s, "b* 12.2", { x, y: iy + 0.35, w: iw, h: 0.5, fontSize: 22, bold: true, color: C.white, align: "center", valign: "middle" });
        pill(s, "위험도 중간", x + 0.25, iy + 1.05, iw - 0.5, C.prop, 11.5, 0.36);
      }
      if (i === 3) {
        s.addImage({ path: IMG("offshore_legs.png"), x, y: iy, w: iw, h: iw, sizing: { type: "cover", w: iw, h: iw } });
        [[1, C.mat, 0.27, 0.64], [2, C.prop, 0.78, 0.42], [3, C.inc, 0.42, 0.16]].forEach(([k, col, fx, fy]) => {
          const d = 0.34, mx = x + fx * iw - d / 2, my = iy + fy * iw - d / 2;
          s.addShape("ellipse", { x: mx, y: my, w: d, h: d, fill: { color: col }, line: { color: C.white, width: 1.5 } });
          txt(s, String(k), { x: mx, y: my, w: d, h: d, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle" });
        });
      }
      if (i < n - 1) s.addImage({ data: I.arrowW, x: cx + cw / 2 - 0.17, y: iy + iw / 2 - 0.17, w: 0.34, h: 0.34 });
      const last = i === n - 1;
      s.addShape("homePlate", { x: MX + i * cw + 0.05, y: by + bh - 0.68, w: cw - 0.1, h: 0.5, fill: { color: last ? C.navy : C.white }, line: { color: last ? C.navy : C.white } });
      txt(s, labels[i], { x: MX + i * cw + 0.1, y: by + bh - 0.68, w: cw - 0.4, h: 0.5, fontSize: 14, bold: true, color: last ? C.white : C.ink, align: "center", valign: "middle" });
    }
    txt(s, "예시 화면 (개념도)", { x: W - MX - 2.5, y: 1.3, w: 2.5, h: 0.25, fontSize: 10, italic: true, color: C.faint, align: "right" });
    const diff = [
      ["차별성 1", "사진 한 장으로 비파괴 · 원격 판정", C.navy, [[I.camera, "시편 절단 · 실험실 분석 없이 판정"], [I.heli, "접근 어려운 비말대도 드론으로 촬영"]]],
      ["차별성 2", "규제 주기 안에서 자원 배치 최적화", C.band, [[I.map, "위험 지점부터 사람 · 배 · 리프트 투입"], [I.route, "로프 액세스 출동 여부 사전 판단"]]],
    ];
    const dw = (W - 2 * MX - 0.5) / 2, dy = 4.75, dh = 1.75;
    diff.forEach(([tag, head, col, rows], i) => {
      const x = MX + i * (dw + 0.5);
      rr(s, x + 0.35, dy + 0.2, dw - 0.35, dh - 0.2, C.white, C.line);
      s.addShape("rect", { x: x + 0.35, y: dy + 0.2, w: dw - 0.35, h: 0.48, fill: { color: col }, line: { color: col } });
      txt(s, head, { x: x + 1.5, y: dy + 0.2, w: dw - 1.3, h: 0.48, fontSize: 15, bold: true, color: C.white, valign: "middle" });
      s.addShape("ellipse", { x, y: dy, w: 1.05, h: 1.05, fill: { color: col }, line: { color: C.white, width: 3 } });
      txt(s, tag, { x, y: dy, w: 1.05, h: 1.05, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
      rows.forEach(([icn, t], k) => {
        const yy = dy + 0.85 + k * 0.45;
        iconDot(s, icn, x + 1.5, yy, 0.36, col);
        txt(s, t, { x: x + 1.98, y: yy, w: dw - 2.1, h: 0.36, fontSize: 13.5, valign: "middle" });
      });
    });
    banner(s, [["사진 한 장 → 위험도 → ", 0], ["점검 우선순위", 1], ["까지, 현장에서 바로", 0]]);
    note(s, 32, "MLC는 이렇게 작동합니다. 드론이나 카메라로 접근이 어려운 구간을 먼저 찍고, 녹이 있는 부분만 자동으로 골라낸 뒤, 그 녹의 색으로 위험도를 높음, 중간, 낮음으로 판정합니다. 마지막으로 구조물 사진 위에 위험한 지점부터 번호를 매긴 지도를 드립니다. 시편을 자르지 않고 사진 한 장으로 판정한다는 점, 그리고 규제 주기는 그대로 지키면서 사람과 배와 리프트를 위험한 곳에 먼저 보낼 수 있다는 점이 차별점입니다.");
  }

  // ===== 6. Technology =====
  {
    const s = slide("05", "핵심기술(Technology)", "색으로 추적하는 부식 속도");
    const lw = 1.55, x0 = MX + lw + 0.1, colW = (W - MX - x0) / 3, pw = 1.6;
    [["표면 사진", 2.85], ["녹의 색", 3.98], ["부식 속도", 5.4]].forEach(([t, yy]) => txt(s, t, { x: MX, y: yy - 0.2, w: lw, h: 0.4, fontSize: 14, bold: true, color: C.muted, valign: "middle" }));
    const base = 6.35, maxH = 1.35, maxCR = 1.45;
    STG.forEach((st, i) => {
      const cx = x0 + i * colW + colW / 2, px = cx - pw / 2;
      pill(s, "위험도  " + st.name, cx - 0.8, 1.6, 1.6, st.col, 12.5, 0.34);
      s.addImage({ path: IMG(st.img), x: px, y: 2.02, w: pw, h: pw });
      txt(s, st.time, { x: px, y: 2.02 + pw - 0.3, w: pw - 0.08, h: 0.26, fontSize: 10, bold: true, color: C.white, align: "right", valign: "middle" });
      s.addShape("rect", { x: px, y: 3.76, w: pw, h: 0.45, fill: { color: hex(st.rgb) }, line: { color: C.line } });
      txt(s, `b*  ${st.b}`, { x: px, y: 3.76, w: pw, h: 0.45, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
      const bh = maxH * st.cr / maxCR;
      s.addShape("rect", { x: cx - 0.5, y: base - bh, w: 1.0, h: bh, fill: { color: st.col }, line: { color: st.col } });
      txt(s, `${st.cr.toFixed(2)} mm/년`, { x: cx - 1, y: base - bh - 0.32, w: 2, h: 0.28, fontSize: 13, bold: true, color: st.col, align: "center" });
      if (i < 2) s.addImage({ data: I.arrowR, x: cx + colW / 2 - 0.16, y: 2.02 + pw / 2 - 0.16, w: 0.32, h: 0.32 });
    });
    s.addShape("line", { x: x0 + 0.2, y: base, w: W - MX - x0 - 0.4, h: 0, line: { color: C.faint, width: 1 } });
    txt(s, "약 4배", { x: W - MX - 1.3, y: base - maxH - 0.05, w: 1.3, h: 0.4, fontSize: 18, bold: true, color: C.mat, align: "right" });
    txt(s, "탄소강 · 240시간 촉진부식 · 단계별 평균값", { x: W - MX - 4.5, y: 1.3, w: 4.5, h: 0.25, fontSize: 10, italic: true, color: C.faint, align: "right" });
    banner(s, [["녹 색(b*)과 부식 속도 순위 상관 ", 0], ["ρ = 0.88", 1], ["  ·  6가지 분석으로 교차 검증", 0]]);
    note(s, 32, "핵심 기술은 실험실에서 검증했습니다. 탄소강을 240시간 촉진부식 시키며 측정했더니, 표면의 녹 색이 변하면서 색 좌표 b* 값이 1.7에서 13.1로 커지고, 같은 기간 실제 부식 속도는 연 0.34mm에서 1.45mm로 약 4배 빨라졌습니다. 색과 부식 속도의 순위 상관계수는 0.88이고, 위험도 세 구간의 부식 속도는 서로 겹치지 않습니다. 이 결과는 SEM, XRD, 전기화학 등 여섯 가지 분석으로 교차 검증했습니다.");
  }

  // ===== 7. Market =====
  {
    const s = slide("06", "시장(Market)", "규제가 만드는 반복 수요");
    const lv = [["Level I", "수상 육안 점검", "도장 · 구조 · 비말대 육안 스캔", "매년", C.rust], ["Level II", "수중 일반 점검", "잠수부 · ROV 조사", "3–5년", C.sea], ["Level III", "수중 정밀 점검", "지정 부위 세척 후 근접 육안 · 피트 계측", "6–10년", C.navy]];
    const cw = (W - 2 * MX - 0.5) / 3;
    lv.forEach(([l, t, d, cyc, col], i) => {
      const x = MX + i * (cw + 0.25), y = 1.6;
      rr(s, x, y, cw, 1.4, C.panel);
      txt(s, l, { x: x + 0.3, y: y + 0.15, w: 2, h: 0.32, fontSize: 13, bold: true, color: col });
      txt(s, cyc, { x: x + cw - 1.8, y: y + 0.08, w: 1.5, h: 0.45, fontSize: 21, bold: true, color: col, align: "right" });
      txt(s, t, { x: x + 0.3, y: y + 0.55, w: cw - 0.6, h: 0.35, fontSize: 15, bold: true });
      txt(s, d, { x: x + 0.3, y: y + 0.92, w: cw - 0.6, h: 0.35, fontSize: 12, color: C.muted });
    });
    txt(s, "미국 해양 구조물 점검 체계 — 4개 기관이 겹쳐 규제", { x: MX, y: 3.15, w: 6, h: 0.3, fontSize: 12, color: C.muted });
    ["BSEE", "PHMSA", "USCG", "ABS"].forEach((r, i) => pill(s, r, MX + 5.3 + i * 1.05, 3.13, 0.95, C.sea, 11, 0.34));
    const y = 3.65, h = 2.8, bw = (W - 2 * MX - 0.3) / 2;
    rr(s, MX, y, bw, h, C.panel);
    txt(s, "구매를 움직이는 3가지 동기", { x: MX + 0.3, y: y + 0.2, w: bw - 0.5, h: 0.32, fontSize: 13, bold: true, color: C.sea });
    [[I.shield, "안전", "놓친 부식 → 사고"], [I.coins, "비용", "선박 · 리프트 일수"], [I.gavel, "규제 · 보험", "기한 내 기록 제출"]].forEach(([icn, t, d], i) => {
      const yy = y + 0.72 + i * 0.66;
      iconDot(s, icn, MX + 0.3, yy, 0.48, C.navy);
      txt(s, [{ text: t + "   ", options: { bold: true, fontSize: 15 } }, { text: d, options: { color: C.muted, fontSize: 13 } }], { x: MX + 0.95, y: yy, w: bw - 1.15, h: 0.48, valign: "middle" });
    });
    const x2 = MX + bw + 0.3;
    rr(s, x2, y, bw, h, C.rustLight);
    txt(s, "MLC가 줄이는 비용", { x: x2 + 0.3, y: y + 0.2, w: bw - 0.5, h: 0.32, fontSize: 13, bold: true, color: C.rust });
    const items = [[I.ship, "선박 용선 · 기상 대기", "위험 지점부터 배 투입"], [I.tools, "리프트 · 비계 임대", "필요한 구간만 골라 임대"], [I.route, "로프 액세스 출동", "출동 여부 사전 판단"], [I.file, "점검 기록 · 문서 작업", "사진 리포트로 정리"]];
    const iw = (bw - 0.6) / 2;
    items.forEach(([icn, t, d], i) => {
      const xx = x2 + 0.3 + (i % 2) * iw, yy = y + 0.72 + Math.floor(i / 2) * 0.9;
      iconDot(s, icn, xx, yy + 0.02, 0.46, C.rust);
      txt(s, t, { x: xx + 0.58, y: yy, w: iw - 0.65, h: 0.3, fontSize: 13, bold: true });
      txt(s, d, { x: xx + 0.58, y: yy + 0.3, w: iw - 0.65, h: 0.3, fontSize: 11, color: C.muted });
    });
    txt(s, "미국 현장 인터뷰 (Deepwater · MISTRAS · Brown Corrosion 등)", { x: x2 + 0.3, y: y + h - 0.38, w: bw - 0.5, h: 0.25, fontSize: 10, italic: true, color: C.faint });
    banner(s, [["규제로 반복되는 점검 = ", 0], ["매년 돌아오는 수요", 1]]);
    note(s, 30, "이 시장은 규제가 반복 수요를 만듭니다. 미국 해양 플랫폼은 매년, 3~5년, 6~10년 주기로 점검이 의무이고, 네 개 기관이 겹쳐서 규제합니다. 인터뷰에서 확인한 구매 동기는 안전, 비용, 규제와 보험 대응 세 가지였습니다. MLC는 그중 비용에서 바로 효과를 냅니다. 위험한 지점부터 배를 보내 선박 용선과 기상 대기 일수를 줄이고, 필요한 구간에만 리프트와 비계를 빌리고, 로프 액세스 출동 여부를 미리 판단하고, 점검 기록도 사진 리포트로 정리합니다.");
  }

  // ===== 8. Business model =====
  {
    const s = slide("07", "비즈니스 모델(Business Model)", "파트너와 함께 만드는 점검 데이터 플랫폼");
    const by = 1.65, bh = 1.95;
    rr(s, MX, by, W - 2 * MX, bh, C.seaLight, null, 0.95);
    const d = 1.65, cy = by + (bh - d) / 2;
    const nodes = [["파트너", "드론 운용사 · 방식 전문 업체", C.sea, MX + 0.2], ["MLC", "녹 색 기반 위험도 분석", C.navy, W / 2 - d / 2], ["고객사", "해양 석유 · 가스 운영사", C.rust, W - MX - 0.2 - d]];
    nodes.forEach(([t, sub, col, x]) => {
      s.addShape("ellipse", { x, y: cy, w: d, h: d, fill: { color: col }, line: { color: C.white, width: 3 } });
      txt(s, t, { x, y: cy, w: d, h: d, fontSize: 21, bold: true, color: C.white, align: "center", valign: "middle" });
    });
    const a1x = MX + 0.2 + d + 0.15, a1w = W / 2 - d / 2 - 0.15 - a1x;
    const a2x = W / 2 + d / 2 + 0.15, a2w = W - MX - 0.2 - d - 0.15 - a2x;
    [[a1x, a1w, "현장 촬영 데이터 · 고객 소개", C.sea], [a2x, a2w, "점검 우선순위 리포트", C.navy]].forEach(([x, w, t, col]) => {
      s.addShape("line", { x, y: by + bh / 2, w, h: 0, line: { color: col, width: 2, endArrowType: "triangle" } });
      txt(s, t, { x, y: by + bh / 2 - 0.45, w, h: 0.35, fontSize: 14, bold: true, color: col, align: "center" });
    });
    txt(s, "점검 건당 · 구독형 요금", { x: a2x, y: by + bh / 2 + 0.1, w: a2w, h: 0.32, fontSize: 13, color: C.rust, bold: true, align: "center" });
    const subs = [[nodes[0][1], "Farwest Corrosion Control 등"], [nodes[1][1], "위험도 판정 · 우선순위 지도"], [nodes[2][1], "점검 프로그램 매니저"]];
    subs.forEach(([a, b], i) => {
      const cx = nodes[i][3] + d / 2, al = ["left", "center", "right"][i];
      const bx = i === 0 ? MX : i === 2 ? W - MX - 3.2 : cx - 1.6;
      txt(s, [{ text: a, options: { bold: true, breakLine: true } }, { text: b, options: { color: C.muted, fontSize: 11.5 } }], { x: bx, y: by + bh + 0.12, w: 3.2, h: 0.6, fontSize: 13, align: al });
    });
    txt(s, "고객사 안의 의사결정 구조", { x: MX, y: 4.55, w: 5, h: 0.3, fontSize: 13, bold: true, color: C.sea });
    const eco = [[I.crown, "결정", "CEO · 대표", "안전 · 비용 · 규제 대응", C.rust], [I.star, "추천", "부식 엔지니어링 매니저", "기술 검토 후 추천", C.sea], [I.hat, "사용", "현장 엔지니어 · 검사원", "점검 동선 · 출동 결정", C.sea], [I.users, "영향", "방식 전문 업체 · 연구자", "기술 교차 검증 · 소개", C.muted]];
    const ew = (W - 2 * MX - 0.45) / 4;
    eco.forEach(([icn, tag, who, dsc, col], i) => {
      const x = MX + i * (ew + 0.15), y = 4.95;
      rr(s, x, y, ew, 1.5, C.panel);
      iconDot(s, icn, x + 0.2, y + 0.2, 0.55, col);
      pill(s, tag, x + 0.9, y + 0.31, 0.75, col, 11, 0.32);
      txt(s, who, { x: x + 0.2, y: y + 0.85, w: ew - 0.3, h: 0.3, fontSize: 13.5, bold: true });
      txt(s, dsc, { x: x + 0.2, y: y + 1.13, w: ew - 0.3, h: 0.28, fontSize: 11.5, color: C.muted });
    });
    banner(s, [["점검 건당 · 구독형 요금으로 ", 0], ["매 점검 주기마다 반복 매출", 1]]);
    note(s, 25, "비즈니스 모델입니다. 드론 운용사와 방식 전문 업체가 현장 촬영 데이터를 가져오고 고객을 소개하면, MLC가 위험도를 분석해 해양 석유·가스 운영사에 점검 우선순위 리포트를 드립니다. 요금은 점검 건당 또는 구독형입니다. 고객사 안에서는 현장 검사원이 쓰고, 부식 엔지니어링 매니저가 추천하고, 대표가 결정합니다.");
  }

  // ===== 9. Traction & next =====
  {
    const s = slide("08", "성과 및 계획(Milestone)", "지금까지의 성과와 다음 단계");
    const items = [
      ["실험실\n기술 검증", "240시간 촉진부식\n6가지 분석 교차 검증", true],
      ["미국 고객\n인터뷰 35건", "Tex-Corps · 2026년 7월\n타깃: 해양 석유 · 가스", true],
      ["비용 절감\n근거 정량화", "선박 · 리프트 · 문서 작업 시간\n절감액으로 ROI 산출", false],
      ["두께 측정\n데이터 결합", "사진이 못 보는\n두께 감소를 보정", false],
      ["유망 현장\n심화", "미드스트림 압축 기지 · 비말대\n라이저 지지 프레임", false],
    ];
    const n = items.length, cw = (W - 2 * MX) / n, base = 4.25;
    s.addShape("rtTriangle", { x: MX + 0.3, y: 1.75, w: W - 2 * MX - 0.6, h: base - 1.75 - 0.1, flipH: true, fill: { color: C.seaLight }, line: { color: C.seaLight } });
    const cols = [C.navyMid, C.navyMid, C.sea, C.navy, C.navy];
    items.forEach(([t, d, done], i) => {
      const cx = MX + i * cw + cw / 2, dd = 1.4 + i * 0.2, y = base - dd - 0.05;
      s.addShape("ellipse", { x: cx - dd / 2, y, w: dd, h: dd, fill: { color: cols[i] }, line: { color: C.white, width: 3 } });
      txt(s, t, { x: cx - dd / 2, y, w: dd, h: dd, fontSize: 13 + i * 0.5, bold: true, color: C.white, align: "center", valign: "middle" });
      const chev = i === 0 ? "homePlate" : "chevron";
      const shade = done ? "8A9099" : ["4A4F57", "3A3F45", "2E3238"][i - 2];
      s.addShape(chev, { x: MX + i * cw, y: base + 0.15, w: cw + 0.1, h: 0.48, fill: { color: shade }, line: { color: shade } });
      txt(s, done ? "완료" : "다음 단계", { x: MX + i * cw + 0.2, y: base + 0.15, w: cw - 0.3, h: 0.48, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" });
      txt(s, d, { x: cx - cw / 2 + 0.05, y: base + 0.82, w: cw - 0.1, h: 0.75, fontSize: 12, color: C.ink, align: "center" });
    });
    banner(s, [["검증된 기술 + 확인된 수요 → ", 0], ["현장 데이터로 정확도를 넓혀 갑니다", 1]]);
    note(s, 25, "지금까지 실험실 검증을 마쳤고, 미국에서 35명을 만나 타깃을 해양 석유·가스로 좁혔습니다. 다음 단계는 세 가지입니다. 선박 일수와 리프트 임대처럼 실제 절감액을 숫자로 만들고, 사진이 볼 수 없는 두께 감소를 보완하기 위해 두께 측정 데이터와 결합하고, 반응이 가장 좋았던 현장을 더 깊이 검증하겠습니다.");
  }

  // ===== 10. Team & infrastructure =====
  {
    const s = slide("09", "팀 구성 및 인프라", "창업팀 현황 및 기술 개발 인프라");
    const people = [["team_woo.png", "(EL) 우성훈", "부식 예측 연구 · 고객 인터뷰"], ["team_park.jpg", "(EM) 박준범", "부식 연구 · 고객 인터뷰"]];
    const cw = 3.3, gap = 1.0, x0 = (W - 2 * cw - gap) / 2;
    people.forEach(([img, name, role], i) => {
      const x = x0 + i * (cw + gap), cx = x + cw / 2, dd = 1.75;
      s.addShape("diamond", { x: cx - dd / 2 - 0.12, y: 1.42, w: dd + 0.24, h: dd + 0.24, fill: { color: C.panel }, line: { color: C.navy, width: 2.25 } });
      s.addImage({ path: IMG(img), x: cx - 0.62, y: 1.68, w: 1.24, h: 1.24, rounding: true, sizing: { type: "cover", w: 1.24, h: 1.24 } });
      s.addShape("rect", { x, y: 3.15, w: cw, h: 0.4, fill: { color: C.navy }, line: { color: C.navy } });
      txt(s, name, { x, y: 3.15, w: cw, h: 0.4, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" });
      s.addShape("rect", { x, y: 3.55, w: cw, h: 0.82, fill: { color: C.panel }, line: { color: C.panel } });
      [["소속", "경상국립대학교"], ["전공", "기계시스템공학"], ["역할", role]].forEach(([k, v], j) => {
        txt(s, [{ text: "· " + k + "   ", options: { color: C.navy, bold: true } }, { text: v }], { x: x + 0.2, y: 3.6 + j * 0.25, w: cw - 0.3, h: 0.25, fontSize: 10.5, valign: "middle" });
      });
    });
    s.addShape("rect", { x: 0, y: 4.6, w: W, h: H - 4.6, fill: { color: C.dark }, line: { color: C.dark } });
    txt(s, "부식 · 분석 인프라 — 실험실 검증에 사용한 장비", { x: MX, y: 4.75, w: 8, h: 0.32, fontSize: 13, bold: true, color: C.navyLight });
    const eq = [[I.fog, "복합부식 시험기"], [I.micro, "SEM / EDS"], [I.wave, "XRD"], [I.bolt, "전위차계 (LPR)"], [I.layers, "공초점 현미경"], [I.cam, "촬영 · 이미지 분석"]];
    const ew = (W - 2 * MX) / eq.length;
    eq.forEach(([icn, t], i) => {
      const cx = MX + i * ew + ew / 2, d = 1.05;
      dot(s, cx - d / 2, 5.25, d, C.white);
      s.addImage({ data: icn, x: cx - 0.28, y: 5.25 + d / 2 - 0.28, w: 0.56, h: 0.56 });
      s.addShape("rect", { x: cx - ew / 2 + 0.06, y: 6.48, w: ew - 0.12, h: 0.48, fill: { color: "3A3F46" }, line: { color: "3A3F46" } });
      txt(s, t, { x: cx - ew / 2 + 0.06, y: 6.48, w: ew - 0.12, h: 0.48, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
    });
    note(s, 15, "팀 MLC는 경상국립대학교 기계시스템공학과의 우성훈, 박준범입니다. 두 사람이 부식 연구를 직접 했고 미국 고객 인터뷰 35건도 함께 진행했습니다. 복합부식 시험기, 전자현미경, X선 회절, 전기화학 장비 등 실험실 분석 인프라를 바탕으로 기술을 검증했습니다.");
  }

  // ===== 11. Closing =====
  {
    const s = pres.addSlide(); page += 1;
    s.background = { color: C.paper };
    dots(s, MX + 0.3, 0.75);
    txt(s, [
      { text: "S", options: { color: C.navy, bold: true } }, { text: "ee the rust,", options: { color: C.ink, breakLine: true } },
      { text: "S", options: { color: C.rust, bold: true } }, { text: "ave the structure.", options: { color: C.ink } },
    ], { x: MX + 0.3, y: 1.55, w: 7.2, h: 1.9, fontSize: 44, valign: "middle" });
    txt(s, "어디부터 점검할지, 녹의 색이 먼저 알려드립니다", { x: MX + 0.3, y: 3.55, w: 7.2, h: 0.45, fontSize: 18, bold: true, color: C.navy });
    const sw = 0.5;
    RGB.forEach((c, i) => s.addShape("rect", { x: MX + 0.3 + i * (sw + 0.06), y: 4.35, w: sw, h: sw, fill: { color: hex(c) }, line: { color: C.white, width: 0.75 } }));
    s.addImage({ data: I.mail, x: MX + 0.3, y: 5.55, w: 0.28, h: 0.28 });
    txt(s, "MLC  ·  우성훈 · 박준범  ·  zscfvgbb@naver.com", { x: MX + 0.7, y: 5.5, w: 7, h: 0.38, fontSize: 14, color: C.muted, valign: "middle" });
    txt(s, "경상국립대학교", { x: MX + 0.3, y: 6.0, w: 5, h: 0.38, fontSize: 15, bold: true, color: C.navy });
    const pw = 4.6, ph = pw * 284 / 386;
    s.addImage({ path: IMG("offshore_legs.png"), x: W - MX - pw - 0.2, y: 2.1, w: pw, h: ph });
    dots(s, W - MX - 0.62, 0.75, true);
    note(s, 12, "저희는 이 기술을 현장에서 함께 검증할 파트너를 찾고 있습니다. 어디부터 점검할지, 녹의 색이 먼저 알려드리겠습니다. 감사합니다.");
  }

  // Same XML clean-ups as build.js / ir.js.
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
  require("fs").writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT);
})();

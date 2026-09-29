// ICIMC 2026 talk — HTML deck (html-ppt skill, presenter-mode-reveal pattern)
// Build: node build-html.js  -> index.html
// Charts are inline SVG generated from the same data as ../src/build.js.
const fs = require("fs");
const path = require("path");

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
const stageOf = (c) => (c <= 5 ? 0 : c <= 15 ? 1 : 2);
const STAGE = [
  { name: "Incubation", span: "0–5", cls: "inc", x0: 0, x1: 7.5 },
  { name: "Propagation", span: "10–15", cls: "prop", x0: 7.5, x1: 17.5 },
  { name: "Maturation", span: "20–30", cls: "mat", x0: 17.5, x1: 30 },
];
// validated categorical order (dataviz validator, light surface): rust, teal, gold, purple
const S = { rust: "var(--c-rust)", teal: "var(--c-teal)", gold: "var(--c-gold)", purple: "var(--c-purple)", gray: "var(--c-gray)", red: "var(--c-red)" };
const f = (v, d = 1) => Number(v).toFixed(d);
const rgb = (a) => `rgb(${a.join(",")})`;

// ---------- SVG chart helpers ----------
// Line/scatter chart on a numeric x axis. Returns an SVG string.
function xyChart(o) {
  const W = o.w, H = o.h, m = Object.assign({ l: 84, r: 28, t: 20, b: 70 }, o.m || {});
  const pw = W - m.l - m.r, ph = H - m.t - m.b;
  const xmin = o.xmin ?? 0, xmax = o.xmax ?? 30;
  const X = (v) => m.l + ((v - xmin) / (xmax - xmin)) * pw;
  const Y = (v) => m.t + (1 - (v - o.ymin) / (o.ymax - o.ymin)) * ph;
  let s = `<svg class="chart ${o.cls || ""}" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="${o.aria}">`;
  if (o.bands !== false) STAGE.forEach((st) => {
    s += `<rect class="band ${st.cls}" x="${X(Math.max(st.x0, xmin))}" y="${m.t}" width="${X(Math.min(st.x1, xmax)) - X(Math.max(st.x0, xmin))}" height="${ph}"/>`;
  });
  (o.boxes || []).forEach((b) => {
    s += `<rect class="rng ${b.cls}" x="${X(b.x0)}" y="${Y(b.y1)}" width="${X(b.x1) - X(b.x0)}" height="${Y(b.y0) - Y(b.y1)}" rx="10"/>`;
  });
  o.yticks.forEach((v) => {
    s += `<line class="grid" x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}"/>`;
    s += `<text class="tick" x="${m.l - 14}" y="${Y(v) + 7}" text-anchor="end">${o.yfmt ? o.yfmt(v) : v}</text>`;
  });
  (o.xticks || [0, 5, 10, 15, 20, 25, 30]).forEach((v) => {
    s += `<text class="tick" x="${X(v)}" y="${H - m.b + 32}" text-anchor="middle">${v}</text>`;
  });
  s += `<line class="axis" x1="${m.l}" x2="${W - m.r}" y1="${m.t + ph}" y2="${m.t + ph}"/>`;
  s += `<text class="atitle" x="${m.l + pw / 2}" y="${H - 8}" text-anchor="middle">${o.xtitle || "CCT cycle"}</text>`;
  if (o.ytitle) s += `<text class="atitle" transform="translate(22 ${m.t + ph / 2}) rotate(-90)" text-anchor="middle">${o.ytitle}</text>`;
  o.series.forEach((sr) => {
    const xs = sr.xs || CYC;
    const pts = xs.map((x, i) => [x, sr.values[i]]).filter((p) => p[1] !== null);
    if (sr.line !== false) s += `<path class="line" pathLength="1" style="stroke:${sr.color}" d="M${pts.map((p) => `${X(p[0]).toFixed(1)},${Y(p[1]).toFixed(1)}`).join(" L")}"/>`;
    pts.forEach((p, i) => {
      const lab = sr.tip ? sr.tip(p, i) : `${sr.name} · cycle ${p[0]}: ${p[1]}`;
      s += `<g class="pt"><circle class="hit" cx="${X(p[0])}" cy="${Y(p[1])}" r="18"/><circle class="dot" cx="${X(p[0])}" cy="${Y(p[1])}" r="${sr.r || 7}" style="fill:${sr.color}"/><title>${lab}</title></g>`;
    });
  });
  (o.labels || []).forEach((l) => {
    s += `<text class="dlabel ${l.cls || ""}" x="${X(l.x) + (l.dx || 0)}" y="${Y(l.y) + (l.dy || -16)}" text-anchor="${l.anchor || "middle"}">${l.t}</text>`;
  });
  return s + "</svg>";
}

// Vertical bar chart on categories.
function barChart(o) {
  const W = o.w, H = o.h, m = Object.assign({ l: 70, r: 20, t: 40, b: 56 }, o.m || {});
  const pw = W - m.l - m.r, ph = H - m.t - m.b, n = o.cats.length, gs = o.groups.length;
  const Y = (v) => m.t + (1 - (v - o.ymin) / (o.ymax - o.ymin)) * ph;
  const slot = pw / n, bw = Math.min(o.barW || 90, (slot * 0.7) / gs);
  let s = `<svg class="chart" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="${o.aria}">`;
  o.yticks.forEach((v) => {
    s += `<line class="grid" x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}"/><text class="tick" x="${m.l - 12}" y="${Y(v) + 7}" text-anchor="end">${o.yfmt ? o.yfmt(v) : v}</text>`;
  });
  o.cats.forEach((c, i) => {
    const cx = m.l + slot * (i + 0.5);
    o.groups.forEach((g, k) => {
      const v = g.values[i], x = cx - (gs * bw + (gs - 1) * 4) / 2 + k * (bw + 4);
      const col = typeof g.color === "function" ? g.color(i) : g.color;
      const h = Y(o.ymin) - Y(v);
      s += `<g class="pt"><path class="bar" style="fill:${col}" d="M${x},${Y(o.ymin)} v${-(h - 4)} q0,-4 4,-4 h${bw - 8} q4,0 4,4 v${h - 4} z"/><title>${g.name ? g.name + " · " : ""}${c}: ${o.fmt(v)}</title></g>`;
      s += `<text class="vlabel" x="${x + bw / 2}" y="${Y(v) - 12}" text-anchor="middle">${o.fmt(v)}</text>`;
    });
    s += `<text class="tick cat" x="${cx}" y="${H - m.b + 38}" text-anchor="middle">${c}</text>`;
  });
  s += `<line class="axis" x1="${m.l}" x2="${W - m.r}" y1="${Y(o.ymin)}" y2="${Y(o.ymin)}"/>`;
  return s + "</svg>";
}

// Horizontal diverging bars (Spearman rho).
function rhoChart() {
  const W = 760, H = 360, m = { l: 80, r: 110, t: 10, b: 50 };
  const pw = W - m.l - m.r, X = (v) => m.l + ((v + 1) / 2) * pw;
  const rows = [["L*", -0.900, "0.0009", S.gray], ["b*", 0.883, "0.0016", S.gold], ["a*", 0.740, "0.0228", S.red]];
  const rh = (H - m.t - m.b) / 3;
  let s = `<svg class="chart" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="Spearman rho of L*, a*, b* with corrosion rate">`;
  [-1, -0.5, 0, 0.5, 1].forEach((v) => { s += `<line class="grid" x1="${X(v)}" x2="${X(v)}" y1="${m.t}" y2="${H - m.b}"/><text class="tick" x="${X(v)}" y="${H - 14}" text-anchor="middle">${v}</text>`; });
  rows.forEach(([k, v, p, col], i) => {
    const y = m.t + i * rh + rh * 0.2, bh = rh * 0.6, x0 = Math.min(X(0), X(v)), w = Math.abs(X(v) - X(0));
    s += `<text class="cat serif" x="${m.l - 20}" y="${y + bh / 2 + 10}" text-anchor="end">${k}</text>`;
    s += `<g class="pt"><rect class="bar" x="${x0}" y="${y}" width="${w}" height="${bh}" rx="4" style="fill:${col}"/><title>${k}: ρ = ${v.toFixed(3)}, p = ${p}</title></g>`;
    // negative bar: label inside the bar so it never collides with the category name
    const lx = v < 0 ? X(v) + 14 : X(v) + 12;
    s += `<text class="vlabel big ${v < 0 ? "inv" : ""}" x="${lx}" y="${y + bh / 2 + 10}" text-anchor="start">${v < 0 ? "−" : ""}${Math.abs(v).toFixed(3)}</text>`;
  });
  s += `<line class="axis" x1="${X(0)}" x2="${X(0)}" y1="${m.t}" y2="${H - m.b}"/>`;
  return s + "</svg>";
}

// Donut of one 8 h CCT cycle.
function donut() {
  const R = 150, r = 96, c = 170, seg = [["Fog", 2, "var(--c-inc)"], ["Dry", 4, S.gold], ["Humid", 2, "var(--c-prop)"]];
  let a0 = -Math.PI / 2, s = `<svg class="chart" viewBox="0 0 340 340" width="340" height="340" role="img" aria-label="One CCT cycle: fog 2 h, dry 4 h, humid 2 h">`;
  seg.forEach(([n, h, col]) => {
    const a1 = a0 + (h / 8) * 2 * Math.PI, g = 0.02;
    const p = (rad, ang) => `${c + rad * Math.cos(ang)},${c + rad * Math.sin(ang)}`;
    s += `<g class="pt"><path style="fill:${col}" d="M${p(R, a0 + g)} A${R},${R} 0 ${a1 - a0 > Math.PI ? 1 : 0} 1 ${p(R, a1 - g)} L${p(r, a1 - g)} A${r},${r} 0 ${a1 - a0 > Math.PI ? 1 : 0} 0 ${p(r, a0 + g)} Z"/><title>${n}: ${h} h</title></g>`;
    a0 = a1;
  });
  s += `<text class="donut-v" x="${c}" y="${c + 8}" text-anchor="middle">8 h</text><text class="donut-l" x="${c}" y="${c + 44}" text-anchor="middle">per cycle</text>`;
  return s + "</svg>";
}

// ---------- slide helpers ----------
let N = 0;
const slides = [];
function slide(title, body, notesEn, notesKo, opt = {}) {
  N += 1;
  slides.push(`
  <!-- ============ ${N}. ${opt.t || title} ============ -->
  <section class="slide ${opt.cls || ""}" data-title="${(opt.t || title).replace(/<[^>]+>/g, "")}">
    ${title ? `<h2 class="h2" data-anim="fade-up">${title}</h2>` : ""}
    <div class="body">${body}</div>
    <aside class="notes">
      ${notesEn.map((p) => `<p>${p}</p>`).join("\n      ")}
      <hr><p class="ko"><b>KO 참고</b> ${notesKo}</p>
    </aside>
  </section>`);
}
const stageLegend = () => `<div class="stage-legend">${STAGE.map((s) => `<span class="st ${s.cls}"><i></i>${s.name} <em>${s.span}</em></span>`).join("")}</div>`;

// ===== 1. Title =====
slide("", `
    <p class="kicker">ICIMC 2026 · Seoul Olympic Parktel · November 3–7, 2026</p>
    <h1 class="h1" data-anim="fade-up">Color-Based Life Prediction<br>of Carbon Steel</h1>
    <p class="subtitle">Linking CIE L*a*b* colorimetry to corrosion rate</p>
    <p class="authors"><b>Seonghun Woo</b>, Junbeom Park, Beomsoo Kim, Jaeseung Kwon, Jung-Pil Noh, Jeonghyeon Yang*</p>
    <p class="affil">Department of Mechanical System Engineering, Gyeongsang National University, Tongyeong, Korea · * Corresponding author</p>
    <div class="strip">${CYC.map((c) => `<figure><img src="img/photo_${c}.png" alt="Specimen surface at ${c} cycles"><figcaption>${c} cyc</figcaption></figure>`).join("")}</div>`,
  ["Hello everyone. I am <strong>Seonghun Woo</strong> from Gyeongsang National University.",
   "My talk asks <strong>one question</strong>: can the <em>color of rust</em> tell us how fast the steel underneath is corroding?",
   "The strip at the bottom is the <strong>actual specimen surface</strong>, cycle 0 to cycle 30.",
   "We tested this on <strong>bare carbon steel</strong> through <strong>30 cycles</strong> of accelerated corrosion."],
  "안녕하십니까. 경상국립대학교 우성훈입니다. 오늘 발표는 질문 하나에서 출발합니다. 녹의 색으로 그 아래 강재의 부식 속도를 알 수 있을까요? 아래 사진은 0~30사이클 실제 시편 표면입니다. 무도장 탄소강을 30사이클 촉진부식 시험으로 검증했습니다.",
  { cls: "cover dark", t: "Title" });

// ===== 2. Outline =====
slide("Outline", `
    <div class="outline">${[["01", "Why rust color?"], ["02", "How we measured"], ["03", "What changed"], ["04", "What color tells us"]].map(([n, t], i) => `<div class="ocard ${i === 3 ? "hl" : ""}"><span class="num">${n}</span><b>${t}</b></div>`).join('<span class="arrow">→</span>')}</div>`,
  ["Four parts.", "First, <strong>why rust color</strong> matters. Second, <strong>how we measured</strong> it.", "Third, <strong>what changed</strong> on the surface — six methods.", "Last, <strong>what the color signal tells us</strong>, and its limits."],
  "네 부분입니다. 녹 색이 왜 중요한지, 측정 방법, 여섯 가지 방법으로 본 변화, 마지막으로 색 신호가 알려 주는 것과 한계입니다.");

// ===== 3. Background =====
slide("Rust color carries phase information", `
    <div class="bg-grid">
      <div>
        <div class="hero-num">3–4<small>%</small></div>
        <p class="hero-cap">of global GDP lost to corrosion</p>
        <div class="card card-soft note-card"><b>Rust phases decide what happens next</b><span>Today: XRD / SEM — the specimen goes to the lab</span></div>
      </div>
      <div class="phases">${[["γ-FeOOH", "Lepidocrocite", "Orange-yellow", "#D38A2E"], ["α-FeOOH", "Goethite", "Reddish-brown", "#7A3E1D"], ["Fe₃O₄", "Magnetite", "Black", "#1B1B1B"]].map(([f, n, c, col]) => `<div class="phase"><i style="background:${col}"></i><b>${f}</b><span>${n}<br>${c}</span></div>`).join("")}
        <p class="fine">Swatch colors are illustrative.</p></div>
    </div>
    <div class="question"><span>Research question</span><p>Can one color value from a photograph report the corrosion rate of bare carbon steel?</p></div>`,
  ["Corrosion costs an estimated <strong>3 to 4 percent of global GDP</strong>.",
   "The <strong>rust layer</strong> matters because it controls what happens next — and its phases <em>differ in color</em>: lepidocrocite orange-yellow, goethite reddish-brown, magnetite black.",
   "Today we identify them with <strong>XRD or SEM</strong>. Both need the specimen in the lab.",
   "If color carries phase information, a <strong>camera</strong> could take over part of that work. That is what we tested."],
  "부식 비용은 세계 GDP의 3~4%입니다. 녹층이 이후 부식을 좌우하고, 녹의 상은 색이 다릅니다. 지금은 XRD·SEM으로 실험실에서 확인해야 합니다. 색에 상 정보가 있다면 카메라가 일부를 대신할 수 있고, 이를 검증했습니다.");

// ===== 4. Methods =====
slide("Six measurements on one specimen set", `
    <div class="methods">
      <div class="mgroup optical"><span class="glabel">Optical · non-destructive</span>
        <div class="mitem"><i>📷</i><b>Rust area</b><span>Image segmentation</span></div>
        <div class="mitem"><i>🎨</i><b>CIE L*a*b*</b><span>Color of rust pixels</span></div></div>
      <div class="marrow">→<span>checked against</span></div>
      <div class="mgroup ref"><span class="glabel">Reference · lab-based</span>
        <div class="mgrid">${[["🔬", "SEM / EDS", "Morphology, O content"], ["〰", "XRD", "Crystalline phases"], ["⚡", "LPR", "Rp, corrosion rate"], ["▤", "Confocal", "Surface height Sz"]].map(([i, b, s]) => `<div class="mitem"><i>${i}</i><b>${b}</b><span>${s}</span></div>`).join("")}</div></div>
    </div>
    <p class="facts"><b>9</b> sampling points · <b>0–30</b> cycles (240 h) · <b>2</b> specimens per point</p>`,
  ["Six measurements on the <strong>same specimens</strong>.",
   "Two are <strong>optical and non-destructive</strong>: rust area from segmentation, and color from CIE L*a*b*.",
   "Four are <strong>references</strong>: SEM/EDS for morphology and oxygen, XRD for phases, LPR for Rp and corrosion rate, confocal for surface height.",
   "The goal: <em>check the optical data against each reference.</em>"],
  "같은 시편에 여섯 가지 측정을 했습니다. 두 가지는 광학적 비파괴(녹 면적, L*a*b*), 네 가지는 기준 데이터(SEM/EDS, XRD, LPR, 공초점)입니다. 광학 데이터를 각 기준과 대조했습니다.");

// ===== 5. Specimen & CCT =====
slide("Specimens and accelerated exposure", `
    <div class="cct-grid">
      <div class="donut-wrap">${donut()}
        <ul class="dlegend"><li><i class="inc"></i><b>Fog 2 h</b><span>35 °C · UPW spray</span></li><li><i class="gold"></i><b>Dry 4 h</b><span>60 °C · RH &lt; 30%</span></li><li><i class="prop"></i><b>Humid 2 h</b><span>50 °C · RH &gt; 95%</span></li></ul>
        <p class="fine">Q-FOG CCT-600 · modified ISO 14993</p></div>
      <div class="stack">
        <div class="card card-soft"><span class="glabel">Specimen</span><b class="big">Cold-rolled carbon steel</b><p>30 × 70 × 1 mm · 0.12 wt.% C<br>SiC #220 ground · ethanol ultrasonic clean</p></div>
        <div class="card callout"><b>Ultrapure water, not 5% NaCl</b><p>Salt residue disturbed imaging → chloride-free atmospheric test</p></div>
      </div>
    </div>
    <div class="timeline"><div class="tl-line"></div>${CYC.map((c) => `<span class="tl-dot ${STAGE[stageOf(c)].cls}" style="left:${(c / 30) * 100}%"><i></i><em>${c}</em></span>`).join("")}</div>
    <p class="fine tl-cap">Sampling points (cycles) · 30 cycles = 240 h · 2 specimens per point · 18 specimens</p>`,
  ["Cold-rolled carbon steel, <strong>30 × 70 × 1 mm</strong>, 0.12 % carbon, one face ground with 220 grit.",
   "One CCT cycle is <strong>8 hours</strong>: 2 h fog at 35 °C, 4 h dry at 60 °C, 2 h humid above 95 % RH.",
   "We used <strong>ultrapure water instead of salt fog</strong> — salt deposits interfered with imaging. So this is <em>chloride-free atmospheric corrosion</em>, not marine.",
   "Specimens came out at <strong>nine points</strong> up to 30 cycles, in duplicate."],
  "냉연 탄소강 30×70×1 mm, 탄소 0.12%, 한 면 220번 연마. CCT 1사이클 8시간(분무 2 h, 건조 4 h, 습윤 2 h). 염 잔류물 때문에 초순수를 사용해 염화물 없는 대기부식 조건입니다. 아홉 시점에서 두 개씩 회수했습니다.");

// ===== 6. Pipeline =====
slide("Color is read only where rust is", `
    <div class="pipe">${[["Photograph", "Diffuse, fixed lighting · central 75%", `<img src="img/photo_15.png" alt="Specimen photo, cycle 15">`], ["HSV mask", "H 10–60°, S &gt; 24%, V &gt; 16%", `<img src="img/mask_15.png" alt="Rust mask, cycle 15">`], ["Common region", "Pixels rust in ≥ 5 of 9 intervals", `<img src="img/mask_30.png" alt="Common rust region">`], ["L*, a*, b*", "Mean over that fixed pixel set", `<div class="swatch" style="background:${rgb(RGB[5])}">L* 28.0<br>a* 4.4<br>b* 12.2</div>`]].map(([t, d, v], i) => `<div class="pstep"><span class="pn">${i + 1}</span>${v}<b>${t}</b><span>${d}</span></div>`).join('<span class="arrow">→</span>')}</div>
    <div class="chips"><b>Electrochemistry</b>${["LPR ±25 mV vs OCP", "0.167 mV/s", "3.5 wt.% NaCl, 1 cm²", "Ag/AgCl · Pt mesh", "8 scans / point"].map((c) => `<span class="pill">${c}</span>`).join("")}</div>`,
  ["Color is read <strong>only where rust is</strong>.",
   "Photo under diffuse light, central 75 %. Rust pixels masked in HSV — <strong>hue 10 to 60 degrees</strong> — small blobs removed.",
   "Then a <strong>common rust region</strong>: pixels classified as rust in at least five of nine intervals. We average L*, a*, b* over that <em>fixed</em> set every cycle — so a color change is not a coverage change.",
   "For LPR we used <strong>3.5 % NaCl</strong> as test electrolyte; ultrapure water conducts too poorly."],
  "색은 녹이 있는 곳에서만 읽습니다. 확산 조명, 중앙 75% 영역, HSV 색상각 10~60°로 녹 픽셀을 분리했습니다. 아홉 시점 중 5회 이상 녹인 공통 녹 영역에서 L*a*b*를 평균해 색 변화와 면적 변화를 분리했습니다. LPR은 3.5% NaCl에서 측정했습니다.");

// ===== 7. Rust area =====
const RA = RA1.map((v, i) => (v + RA2[i]) / 2);
slide("Coverage rose fastest between cycles 10 and 15", `
    <div class="two-col wide-right">
      <div class="masks">${[[5, 5.6], [10, 20.8], [15, 49.5], [30, 63.4]].map(([c, v]) => `<figure><img src="img/mask_${c}.png" alt="Rust mask ${c} cycles"><figcaption>${c} cyc <b class="${STAGE[stageOf(c)].cls}">${v}%</b></figcaption></figure>`).join("")}</div>
      <div>
        <div class="legend"><span><i style="background:${S.rust}"></i>Specimen 1</span><span><i style="background:${S.teal}"></i>Specimen 2</span></div>
        ${xyChart({ w: 1060, h: 520, ymin: 0, ymax: 70, yticks: [0, 10, 20, 30, 40, 50, 60, 70], ytitle: "Rust area (%)", aria: "Rust area vs cycle, two specimens",
          series: [{ name: "Specimen 1", values: RA1, color: S.rust }, { name: "Specimen 2", values: RA2, color: S.teal }] })}
        ${stageLegend()}
      </div>
    </div>
    <div class="stats3"><div><b class="inc">&lt; 6%</b><span>cycles 0–5</span></div><div><b class="prop">49.5%</b><span>cycle 15</span></div><div><b class="mat">≈ 63%</b><span>cycles 20–30</span></div></div>`,
  ["Left: rust masks. At cycle 5 — <strong>isolated spots</strong>. By cycle 15 — <strong>a network</strong>.",
   "Right: coverage for both specimens. <strong>Under 6 %</strong> through cycle 5, a jump to <strong>about 50 %</strong> between 10 and 15, then a <strong>plateau near 63 %</strong>.",
   "The specimens agree within 2 % from cycle 20.",
   "These breaks define the <strong>three stages</strong> — incubation, propagation, maturation. <em>Every chart from here uses the same shaded bands.</em>"],
  "왼쪽은 녹 마스크, 5사이클은 고립된 점, 15사이클은 망입니다. 면적률은 5사이클까지 6% 미만, 10~15사이클에 약 50%로 급증, 이후 63% 부근에서 평탄. 이 변곡점으로 세 단계를 정의했고 이후 모든 그래프에 같은 구간 색을 씁니다.");

// ===== 8. Rp & CR =====
slide("Polarization resistance fell as the layer grew", `
    <div class="two-col">
      <div><p class="ctitle">Polarization resistance, Rp (Ω·cm²)</p>${xyChart({ w: 820, h: 470, ymin: 0, ymax: 1000, yticks: [0, 200, 400, 600, 800, 1000], aria: "Rp vs cycle",
        series: [{ name: "Rp", values: RP, color: S.purple, tip: (p) => `Rp · cycle ${p[0]}: ${f(p[1])} Ω·cm²` }],
        labels: [{ x: 1, y: 852.2, t: "852", dx: 30, dy: 8, anchor: "start" }, { x: 10, y: 441.1, t: "441" }, { x: 15, y: 301.3, t: "301" }, { x: 30, y: 191.2, t: "191", dx: -8, anchor: "end" }] })}</div>
      <div><p class="ctitle">Corrosion rate (mmpy)</p>${xyChart({ w: 820, h: 470, ymin: 0, ymax: 2, yticks: [0, 0.5, 1, 1.5, 2], yfmt: (v) => f(v, 1), aria: "Corrosion rate vs cycle",
        series: [{ name: "Corrosion rate", values: CR, color: S.rust, tip: (p) => `CR · cycle ${p[0]}: ${f(p[1], 3)} mmpy` }],
        labels: [{ x: 3, y: 0.447, t: "0.45" }, { x: 30, y: 1.745, t: "1.75", dx: -14, anchor: "end" }] })}</div>
    </div>
    <div class="stats3 cards"><div><b>Cycle 1</b><span>Transient Rp peak — thin initial oxide film</span></div><div class="prop-bg"><b class="prop">441 → 301</b><span>cycles 10 → 15: steepest Rp drop</span></div><div class="rust-bg"><b class="rust">3.9×</b><span>corrosion rate, cycle 3 → 30</span></div></div>`,
  ["Rp and corrosion rate move in <strong>opposite directions</strong>.",
   "At cycle 1, Rp briefly rose to <strong>852</strong> — a thin initial oxide film.",
   "The <strong>steepest drop</strong> is cycles 10 to 15, <strong>441 to 301</strong> — the same window where coverage jumped.",
   "Corrosion rate went from <strong>0.45 to 1.75 mm per year</strong>, almost four times. After cycle 20 both slow down — <em>diffusion through a thicker layer.</em>"],
  "Rp와 부식속도는 반대로 움직입니다. 1사이클 Rp 852는 얇은 초기 산화막 때문입니다. 가장 가파른 감소는 10~15사이클(441→301)로 면적 급증 구간과 같습니다. 부식속도는 0.45→1.75 mmpy로 약 4배. 20사이클 이후 완만해져 확산 지배와 부합합니다.");

// ===== 9. SEM/EDS =====
slide("Surface oxygen climbed to 37.9 wt.%", `
    <div class="two-col sem-grid">
      <div class="sem"><span></span><b>SEM ×500</b><b>O map</b>${[0, 5, 15, 30].map((c) => `<em class="${STAGE[stageOf(c)].cls}">${c} cyc</em><img src="img/sem_${c}.jpg" alt="SEM ${c} cycles"><img src="img/omap_${c}.jpg" alt="O map ${c} cycles">`).join("")}</div>
      <div><p class="ctitle">EDS oxygen content (map sum, wt.%)</p>
        ${barChart({ w: 1020, h: 560, ymin: 0, ymax: 40, yticks: [0, 10, 20, 30, 40], cats: O_CYC.map((c) => `${c} cyc`), fmt: (v) => f(v), aria: "EDS oxygen content by cycle",
          groups: [{ values: O_VAL, color: (i) => `var(--c-${STAGE[stageOf(O_CYC[i])].cls})` }] })}
        <div class="callout-line"><b>+1.9 wt.% O</b> from cycle 20 → 30: composition near saturation</div></div>
    </div>`,
  ["SEM and EDS show the same trend at the <strong>microscale</strong>.",
   "Fresh surface: <strong>2.3 %</strong> oxygen, only grinding marks. Cycle 5: <strong>oxide nodules</strong>, 20 %. Cycle 15: a <strong>porous, cracked layer</strong>, 31.8 %.",
   "Cycle 30: <strong>no bare metal</strong>, 37.9 %.",
   "From 20 to 30 oxygen rose by only <strong>1.9 points</strong> — <em>near saturation.</em>"],
  "초기 표면 산소 2.3%, 5사이클 노듈과 20%, 15사이클 다공성 균열층과 31.8%, 30사이클 맨 금속이 사라지고 37.9%. 20→30사이클 증가폭은 1.9%p로 포화에 가깝습니다.");

// ===== 10. XRD =====
const X10 = (c) => (c / 30) * 100;
slide("γ-FeOOH appeared first, then Fe₃O₄", `
    <div class="two-col xrd-grid">
      <figure class="xrd"><img src="img/xrd.png" alt="XRD patterns No.1–9"><figcaption>No.1–9 = cycles 0, 1, 3, 5, 10, 15, 20, 25, 30</figcaption></figure>
      <div><p class="ctitle">Phase detected by XRD</p>
        <div class="gantt">${[["α-Fe", 0, "fe"], ["γ-FeOOH", 10, "lep"], ["Fe₃O₄", 15, "mag"], ["α-FeOOH", null, ""]].map(([n, from, cls]) => `<b>${n}</b><div class="track">${from !== null ? `<i class="${cls}" style="left:${X10(from)}%;width:${100 - X10(from)}%"></i>` : `<em>not resolved (no peak at ~21.2° / ~33.2°)</em>`}</div>`).join("")}
          <span></span><div class="gaxis">${[0, 5, 10, 15, 20, 25, 30].map((c) => `<span style="left:${X10(c)}%">${c}</span>`).join("")}</div></div>
        <div class="card callout"><b>No crystalline oxide through cycle 5</b><p class="rust">γ-FeOOH stayed dominant for all 240 h</p></div></div>
    </div>`,
  ["XRD found <strong>no crystalline oxide through cycle 5</strong> — early nodules were amorphous or too thin.",
   "<strong>Lepidocrocite at cycle 10</strong>, <strong>magnetite at cycle 15</strong>; both grew to cycle 30.",
   "No clear goethite peak at 21 or 33 degrees. A broad feature near 35 degrees may overlap goethite, so a small amount can't be ruled out.",
   "The layer stayed <strong>lepidocrocite-dominated</strong> for all 240 hours. <em>Keep this in mind for the color data.</em>"],
  "5사이클까지 결정성 산화물은 없었습니다. 레피도크로사이트는 10사이클, 마그네타이트는 15사이클에 나타나 계속 성장했습니다. 괴타이트 뚜렷한 피크는 없지만 35° 부근 겹침으로 소량 가능성은 배제 못 합니다. 240시간 내내 레피도크로사이트가 지배했습니다.");

// ===== 11. Color =====
slide("The surface turned yellow during propagation", `
    <div class="swatches">${CYC.map((c, i) => `<div><i style="background:${rgb(RGB[i])}" title="cycle ${c}: RGB ${RGB[i].join(", ")}"></i><span>${c} cyc</span></div>`).join("")}</div>
    <div class="two-col color-grid">
      <div><p class="ctitle">L* (lightness)</p>${xyChart({ w: 700, h: 430, ymin: 24, ymax: 36, yticks: [24, 28, 32, 36], aria: "L* vs cycle",
        series: [{ name: "L*", values: Ls, color: S.gray }], labels: [{ x: 0, y: 34, t: "34.0", dx: 14, dy: -10, anchor: "start" }, { x: 30, y: 26.4, t: "26.4", dx: -6, dy: 34, anchor: "end" }] })}</div>
      <div><div class="legend"><span><i style="background:${S.gold}"></i>b* (yellowness)</span><span><i style="background:${S.red}"></i>a* (redness)</span></div>${xyChart({ w: 960, h: 430, ymin: -2, ymax: 16, yticks: [0, 4, 8, 12, 16], aria: "a* and b* vs cycle",
        series: [{ name: "a*", values: As, color: S.red }, { name: "b*", values: Bs, color: S.gold }], labels: [{ x: 5, y: 2.3, t: "2.3", dy: -18 }, { x: 15, y: 12.2, t: "12.2", dx: -8, dy: -18, anchor: "end" }] })}</div>
    </div>
    <div class="takeaways"><span><b class="gray">L*</b> darkens in incubation 34.0 → 30.3</span><span><b class="gold">b*</b> jumps in propagation 2.3 → 12.2</span><span>Both level off in maturation</span></div>`,
  ["These swatches are the <strong>mean color of the common rust region</strong> at each cycle.",
   "In incubation the surface only <strong>darkened</strong>: L* from 34 to about 30, b* below 3.",
   "Between cycles 5 and 15, <strong>b* jumped from 2.3 to 12.2</strong> — the window where lepidocrocite and magnetite crystallized.",
   "In maturation L* and b* level off. <em>L* catches the early darkening; b* catches the yellowing that follows.</em>"],
  "사이클별 공통 녹 영역의 평균색입니다. 잠복기에는 어두워지기만 했고(L* 34→30), 5~15사이클에 b*가 2.3→12.2로 급증했습니다. 레피도크로사이트·마그네타이트 결정화 구간과 같습니다. 성숙기에는 평탄. L*는 초기 어두워짐, b*는 이후 황색화를 잡아냅니다.");

// ===== 12. Correlations =====
slide("b* tracks damage in the same direction", `
    <div class="two-col">
      <div><p class="ctitle">Linear fit, R²</p><div class="legend"><span><i style="background:${S.rust}"></i>vs rust area</span><span><i style="background:${S.teal}"></i>vs O content</span></div>
        ${barChart({ w: 840, h: 540, ymin: 0, ymax: 1, yticks: [0, 0.2, 0.4, 0.6, 0.8, 1], yfmt: (v) => f(v, 1), cats: ["L*", "a*", "b*"], fmt: (v) => f(v, 3), barW: 110, aria: "R squared of L*, a*, b* vs rust area and O content",
          groups: [{ name: "vs rust area", values: [0.783, 0.911, 0.970], color: S.rust }, { name: "vs O content", values: [0.894, 0.717, 0.894], color: S.teal }] })}</div>
      <div><p class="ctitle">Spearman ρ with corrosion rate</p>${rhoChart()}
        <p class="fine">p = 0.0009 (L*) · 0.0016 (b*) · 0.0228 (a*)</p>
        <div class="card callout"><b class="rust">b* chosen as indicator</b><p>Rises with damage; the |ρ| gap to L* is small</p></div></div>
    </div>`,
  ["Against <strong>rust area</strong>, b* had the highest R² — <strong>0.970</strong>. Against <strong>oxygen</strong>, L* and b* tie at 0.894.",
   "Against <strong>corrosion rate</strong> we used Spearman rank correlation — the relation need not be linear.",
   "L* had the largest magnitude, <strong>−0.900</strong>; b* followed at <strong>0.883</strong>. Both p &lt; 0.002.",
   "We chose <strong>b*</strong> because it <em>rises with damage</em>, which reads more easily. The gap to L* is small — don't over-read it."],
  "녹 면적과는 b* R² 0.970이 최고, 산소와는 L*·b*가 0.894로 동률. 부식속도와는 스피어만 상관으로 L* −0.900, b* 0.883(둘 다 p<0.002). 손상과 같은 방향으로 오르는 b*를 지표로 택했고, L*와의 차이는 작습니다.");

// ===== 13. b* vs mmpy =====
const RNG = [[-0.2, 2.7, 0.193, 0.447], [6.6, 12.2, 0.524, 0.838], [12.1, 13.9, 1.176, 1.745]];
slide("Each stage sits in its own b* and corrosion-rate band", `
    <div class="two-col wide-left">
      <div class="rel">${xyChart({ w: 1080, h: 720, xmin: -2, xmax: 16, ymin: 0, ymax: 2, xticks: [-2, 0, 2, 4, 6, 8, 10, 12, 14, 16], yticks: [0, 0.5, 1, 1.5, 2], yfmt: (v) => f(v, 1),
          xtitle: "b* (yellowness)", ytitle: "Corrosion rate (mmpy)", bands: false, aria: "Corrosion rate vs b*, grouped by stage",
          boxes: RNG.map(([b0, b1, r0, r1], k) => ({ x0: b0 - 0.3, x1: b1 + 0.3, y0: r0 - 0.05, y1: r1 + 0.05, cls: STAGE[k].cls })),
          series: STAGE.map((st, k) => { const idx = CYC.map((c, i) => i).filter((i) => stageOf(CYC[i]) === k); return { name: st.name, xs: idx.map((i) => Bs[i]), values: idx.map((i) => CR[i]), color: `var(--c-${st.cls})`, line: false, r: 11, tip: (p, j) => `cycle ${CYC[idx[j]]}: b* ${p[0]}, ${f(p[1], 3)} mmpy` }; }),
          labels: [{ x: -0.5, y: 0.447, t: "Incubation", dy: -22, anchor: "start", cls: "inc" }, { x: 6.3, y: 0.838, t: "Propagation", dy: -22, anchor: "start", cls: "prop" }, { x: 11.7, y: 1.745, t: "Maturation", dy: 8, anchor: "end", cls: "mat" }] })}
        <div class="rho-badge"><b>ρ = 0.883</b><span>p = 0.0016 · n = 9</span></div></div>
      <div class="card card-soft why"><span class="glabel rust">Why yellowness follows rate</span>
        <i class="phase-dot"></i><b class="big">γ-FeOOH</b>
        <div class="why2"><div><span>Orange-yellow</span><b class="gold">b* ↑</b></div><div><span>Porous</span><b class="rust">mmpy ↑</b></div></div>
        <p class="strong">One phase moves both signals</p></div>
    </div>`,
  ["Every sampling point: <strong>b* on x, corrosion rate on y</strong>.",
   "Grouped by stage, points fall into <strong>three separate boxes</strong>. Rate bands don't overlap: 0.19–0.45, 0.52–0.84, 1.18–1.75 mmpy. b* bands touch only at 12.1–12.2.",
   "<em>Why should yellowness follow rate?</em> Lepidocrocite is <strong>orange-yellow</strong> and <strong>porous</strong>. As it accumulates, b* goes up and the layer passes more electrolyte.",
   "<strong>One phase moves both signals</strong> — physically grounded, within these test conditions."],
  "모든 시점을 x축 b*, y축 부식속도로 그렸습니다. 단계별로 세 상자에 분리되고, 부식속도 구간은 겹치지 않으며 b*는 12.1~12.2에서만 맞닿습니다. 레피도크로사이트는 주황빛 노랑이면서 다공성이라 쌓일수록 b*와 부식속도가 함께 오릅니다. 한 상이 두 신호를 움직입니다.");

// ===== 14. Integrated model =====
const norm = (a) => { const lo = Math.min(...a), hi = Math.max(...a); return a.map((v) => +((v - lo) / (hi - lo)).toFixed(3)); };
const MS = [{ name: "Rust area", values: norm(RA), color: S.rust }, { name: "Sz", values: norm(SZ), color: S.teal }, { name: "b*", values: norm(Bs), color: S.gold }, { name: "Corrosion rate", values: norm(CR), color: S.purple }];
slide("Every method changes at the same two boundaries", `
    <div class="two-col wide-left">
      <div>${xyChart({ w: 1120, h: 600, ymin: 0, ymax: 1.05, yticks: [0, 0.25, 0.5, 0.75, 1], ytitle: "Normalized (0 = min, 1 = max)", aria: "Normalized rust area, Sz, b*, corrosion rate vs cycle",
        series: MS.map((m) => Object.assign({ r: 6, tip: (p) => `${m.name} · cycle ${p[0]}: ${p[1]} (normalized)` }, m)) })}</div>
      <div class="side"><div class="legend col">${MS.map((m) => `<span><i style="background:${m.color}"></i>${m.name}</span>`).join("")}</div>
        <p class="strong prop">All rise sharply in propagation</p>
        <p class="dim">Rust area and b* plateau in maturation; Sz and corrosion rate keep rising → the layer thickens, not spreads</p>
        <p class="strong teal">Sz 0.02 → 0.10 mm (5×)</p></div>
    </div>
    <div class="stages">${STAGE.map((st, k) => `<div class="stage ${st.cls}"><div class="xsec x${k}"><i></i><i></i><i></i></div><div><b>${st.name} ${st.span}</b><span>${["Amorphous oxide nuclei at isolated sites", "γ-FeOOH + Fe₃O₄ crystallize and spread", "Layer thickens; transport-controlled"][k]}</span></div></div>`).join("")}</div>`,
  ["Four measurements on <strong>one normalized scale</strong>: rust area, Sz, b*, corrosion rate.",
   "Each changes at the <strong>same two boundaries</strong> — between 5 and 10, and between 15 and 20. Oxygen and XRD, shown earlier, follow them too.",
   "Incubation: <strong>amorphous nuclei</strong>. Propagation: lepidocrocite and magnetite <strong>crystallize and spread</strong>. Maturation: the layer <strong>stops spreading but thickens</strong> — Sz 0.08 to 0.10 mm while coverage stays near 63 %.",
   "<em>Corrosion is then controlled by transport through the layer.</em>"],
  "녹 면적, Sz, b*, 부식속도를 정규화해 한 그래프에 그렸습니다. 모두 같은 두 경계(5~10, 15~20사이클)에서 변하며 산소·XRD도 같습니다. 잠복기 비정질 핵, 전파기 결정화·확산, 성숙기에는 퍼지지 않고 두꺼워져(Sz 0.08→0.10 mm) 층을 통한 물질 이동이 지배합니다.");

// ===== 15. Takeaways =====
slide("What the color signal tells us", `
    <div class="take3">${[["1 → 3", "Staging without cutting", "One color value placed bare steel in one of three corrosion-rate bands"], ["b*", "A phase readout", "Reports the phase make-up of the outer few µm — not total metal loss"], ["?", "Testable prediction", "Equal thickness, different γ-FeOOH fraction → different b*"]].map(([n, t, d], i) => `<div class="card ${i === 0 ? "callout" : "card-soft"}"><span class="tnum">${n}</span><b>${t}</b><p>${d}</p></div>`).join("")}</div>
    <div class="limits"><span class="glabel">Limits</span>${["9 sampling points", "One chloride-free condition", "Marine / chloride transfer untested", "Fixed lighting"].map((l) => `<span>${l}</span>`).join("")}</div>`,
  ["Three takeaways.",
   "One: a single surface color value placed bare steel into <strong>one of three corrosion-rate bands</strong> — <em>without cutting a specimen</em>.",
   "Two: b* reports the <strong>phase make-up of the outer few micrometers</strong>, not total metal loss.",
   "Three: a <strong>testable prediction</strong> — equal thickness, different lepidocrocite fraction, different b*.",
   "Limits: <strong>nine points, one chloride-free condition, fixed lighting</strong>. Transfer to marine atmospheres is untested."],
  "첫째, 색 값 하나로 시편을 자르지 않고 세 부식속도 구간에 배정했습니다. 둘째, b*는 총 손실이 아니라 최외곽 수 µm의 상 구성을 반영합니다. 셋째, 두께가 같고 레피도크로사이트 분율이 다르면 b*가 달라야 한다는 검증 가능한 예측입니다. 한계: 아홉 시점, 염화물 없는 단일 조건, 고정 조명이며 해양 환경 적용은 미검증입니다.");

// ===== 16. Thanks =====
slide("", `
    <div class="thanks">
      <div><h1 class="h1" data-anim="fade-up">Thank you</h1><p class="subtitle">Questions are welcome.</p>
        <p class="authors"><b>Seonghun Woo</b> · zscfvgbb@naver.com<br><span class="affil">Department of Mechanical System Engineering, Gyeongsang National University</span></p></div>
      <div class="tgrid">${RGB.map((c, i) => `<i style="background:${rgb(c)}" title="cycle ${CYC[i]}"></i>`).join("")}</div>
    </div>
    <p class="ack">Acknowledgment — This research was part of the project “Training Blue Tech Leaders for Eco-Friendly Ships” (No. RS-2025-02220459), funded by the Ministry of Oceans and Fisheries, Korea.</p>`,
  ["This work was funded by the <strong>Ministry of Oceans and Fisheries of Korea</strong>.", "Thank you for listening — <strong>happy to take questions</strong>."],
  "해양수산부 지원으로 수행했습니다. 경청해 주셔서 감사합니다. 질문 받겠습니다.",
  { cls: "cover dark", t: "Thank you" });

// ---------- page ----------
const html = `<!DOCTYPE html>
<html lang="en" data-themes="rust-lab">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Color-Based Life Prediction of Carbon Steel · ICIMC 2026</title>
<link rel="stylesheet" href="assets/fonts.css">
<link rel="stylesheet" href="assets/base.css">
<link rel="stylesheet" id="theme-link" href="assets/themes/rust-lab.css">
<link rel="stylesheet" href="assets/animations/animations.css">
<link rel="stylesheet" href="deck.css">
</head>
<body class="tpl-icimc">
<div class="deck">
${slides.join("\n")}
  <div class="deck-footer"><span>ICIMC 2026 · S. Woo et al., Gyeongsang National University</span><span class="slide-number" data-current="1" data-total="${N}"></span></div>
</div>
<script src="assets/runtime.js"></script>
<script src="fit.js"></script>
</body>
</html>
`;
fs.writeFileSync(path.join(__dirname, "index.html"), html);
console.log("wrote index.html,", N, "slides");

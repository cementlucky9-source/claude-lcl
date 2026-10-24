// Lucky Cement — Intercem Los Angeles 2026 (rebuild)
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const A = "/tmp/claude-0/build/assets/";
const OUT = process.argv[2] || "/tmp/claude-0/build/deck.pptx";

const THEME = {
  name: "Lucky Intercem 2026",
  headFontFace: "Anton",
  bodyFontFace: "Montserrat",
  colors: {
    dk1: "1F1F1F", lt1: "FFFFFF", dk2: "6E0E1A", lt2: "F4F3F3",
    accent1: "A6192E", accent2: "6E0E1A", accent3: "C8A2A8", accent4: "8A8A8A",
    accent5: "D9D6D6", accent6: "3A3A3A", hlink: "A6192E", folHlink: "6E0E1A",
  },
};
const HEX = { maroon: "6E0E1A", red: "A6192E", ink: "1F1F1F", grey: "6B6B6B", light: "F4F3F3", line: "E2DEDE", blush: "F6ECED" };
const HEAD = "Anton";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "Lucky Cement Limited";
pres.company = "Lucky Cement Limited";
pres.title = "From Quarry to Quay — Intercem Los Angeles 2026";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const SW = 13.333, SH = 7.5, M = 0.6;

// ---------- icons ----------
async function icon(Comp, color = "#FFFFFF", size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.42, w: 9.5, h: 0.32, fontFace: "Montserrat", fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, margin: 0, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.74, w: 12.1, h: 0.72, fontFace: HEAD, fontSize: 32, color: C.text2, margin: 0, valign: "top", align: "left" }, text: "" } },
    { image: { path: A + "logo.png", x: M, y: 6.93, w: 0.73, h: 0.34, altText: "Lucky Cement logo" } },
    { text: { text: "INTERCEM LOS ANGELES  ·  NOVEMBER 2026", options: { x: 1.5, y: 6.95, w: 5, h: 0.3, fontFace: "Montserrat", fontSize: 8, color: C.accent4, charSpacing: 2, margin: 0, valign: "middle" } } },
  ],
  slideNumber: { x: 12.33, y: 6.95, w: 0.4, h: 0.3, fontFace: "Montserrat", fontSize: 9, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({ title: "BLANK", background: { color: C.background1 }, objects: [] });

// ---------- helpers ----------
function kt(slide, kicker, title) {
  slide.addText(kicker, { placeholder: "kicker" });
  slide.addText(title, { placeholder: "title" });
}
function manualHead(slide, x, kicker, title, w = 5.6, titleSize = 32) {
  slide.addText(kicker, { x, y: 0.42, w, h: 0.32, fontFace: "Montserrat", fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, margin: 0, valign: "middle", isTextBox: true, objectName: "Kicker" });
  slide.addText(title, { x, y: 0.74, w, h: 1.2, fontFace: HEAD, fontSize: titleSize, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "Title" });
}
function footer(slide, num, dark = false) {
  slide.addImage({ path: A + "logo.png", x: M, y: 6.93, w: 0.73, h: 0.34, altText: "Lucky Cement logo" });
}
function stat(slide, x, y, w, num, label, o = {}) {
  const ns = o.size || 40, nh = o.nh || ns / 72 * 1.25;
  slide.addText(num, { x, y, w, h: nh, fontFace: HEAD, fontSize: ns, color: o.color || C.text2, margin: 0, valign: "bottom", isTextBox: true, objectName: "Stat " + num, fit: "none" });
  slide.addText(label, { x, y: y + nh + 0.06, w, h: o.lh || 0.5, fontFace: "Montserrat", fontSize: o.ls || 11, color: o.lcolor || C.text1, margin: 0, valign: "top", isTextBox: true, bold: !!o.lbold });
}
function pill(slide, x, y, w, text, o = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: o.h || 0.34, rectRadius: 0.17, fill: { color: o.fill || C.background1 }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 4, offset: 1, angle: 90 } });
  slide.addText(text, { x, y, w, h: o.h || 0.34, fontFace: "Montserrat", fontSize: o.size || 10, bold: true, color: o.color || C.text2, align: "center", valign: "middle", margin: 0, charSpacing: 1, isTextBox: true });
}
function source(slide, text, o = {}) {
  slide.addText(text, { x: o.x || 6.6, y: o.y || 6.95, w: o.w || 5.6, h: 0.3, fontFace: "Montserrat", fontSize: 7, color: C.accent4, align: "right", valign: "middle", margin: 0, isTextBox: true, objectName: "Source" });
}
function circleIcon(slide, x, y, d, data, fill = C.accent2) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  const p = d * 0.25;
  slide.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

(async () => {
  const I = {};
  const W = "#FFFFFF";
  for (const [k, comp] of Object.entries({
    chart: fa.FaChartLine, bank: fa.FaLandmark, users: fa.FaUsers, user: fa.FaUserTie, hand: fa.FaHandHoldingUsd,
    vault: fa.FaPiggyBank, star: fa.FaStar, anchor: fa.FaAnchor, ship: fa.FaShip, clock: fa.FaClock,
    plus: fa.FaPlusCircle, percent: fa.FaPercent, bolt: fa.FaBolt, city: fa.FaCity, robot: fa.FaRobot,
    calc: fa.FaCalculator, list: fa.FaListOl, store: fa.FaStore, comments: fa.FaComments, play: fa.FaPlay,
  })) I[k] = await icon(comp, W);
  I.anchorRed = await icon(fa.FaAnchor, "#A6192E");

  // ===== 1. TITLE =====
  {
    const s = pres.addSlide({ masterName: "BLANK" });
    s.addImage({ path: A + "title.jpg", x: 6.2, y: 0, w: 7.133, h: 7.5, altText: "Container and bulk terminal at Karachi at sunset" });
    s.addImage({ path: A + "logo.png", x: M, y: 0.6, w: 1.6, h: 0.74, altText: "Lucky Cement logo" });
    s.addText("INTERCEM  ·  LOS ANGELES  ·  NOVEMBER 2026", { x: M, y: 2.25, w: 5.5, h: 0.35, fontFace: "Montserrat", fontSize: 11, bold: true, color: C.accent1, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText("FROM QUARRY\nTO QUAY", { x: M, y: 2.65, w: 5.5, h: 2.25, fontFace: HEAD, fontSize: 70, color: C.text2, margin: 0, valign: "top", lineSpacingMultiple: 0.95, isTextBox: true, objectName: "Title" });
    s.addText("Pakistan’s ports, terminals and the cement that moves through them", { x: M, y: 4.95, w: 5.0, h: 0.75, fontFace: "Montserrat", fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    s.addText([
      { text: "NAUMAN AHMED", options: { bold: true, color: C.text2, fontSize: 13, charSpacing: 2, breakLine: true } },
      { text: "Head of Exports, Lucky Cement Limited", options: { color: C.accent4, fontSize: 11 } },
    ], { x: M, y: 6.15, w: 5.2, h: 0.65, fontFace: "Montserrat", margin: 0, valign: "top", isTextBox: true });
    s.addNotes("Welcome. Our title is 'From Quarry to Quay': how Pakistan's cement moves from the kiln, onto the quay at Karachi and Port Qasim, and out to the world. This year the ports and terminals are the focus.");
  }

  // ===== 2. INDEX =====
  {
    const s = pres.addSlide({ masterName: "BLANK" });
    s.addImage({ path: A + "index.jpg", x: 0, y: 0, w: 5.0, h: 7.5, altText: "Satellite view of Karachi Port" });
    pill(s, 0.4, 6.75, 2.5, "KARACHI PORT FROM SPACE", { size: 9 });
    manualHead(s, 5.6, "TODAY", "WHAT WE WILL COVER", 7.1);
    const items = [
      ["01", "Pakistan at a glance", "The economy in four numbers"],
      ["02", "Three gateways", "Karachi · Port Qasim · Gwadar"],
      ["03", "Karachi Port", "KGTL & KGTML terminals"],
      ["04", "Port Qasim", "PIBT bulk terminal"],
      ["05", "Gwadar", "The deep-water future"],
      ["06", "Cement industry", "Capacity, exports & news"],
      ["07", "Lucky Cement", "Projects & clean energy"],
      ["08", "Lucky Ustaad", "AI for builders"],
    ];
    items.forEach(([n, t, sub], i) => {
      const col = Math.floor(i / 4), row = i % 4;
      const x = 5.6 + col * 3.75, y = 1.75 + row * 1.2;
      const port = i >= 1 && i <= 4;
      s.addText(n, { x, y, w: 0.85, h: 0.8, fontFace: HEAD, fontSize: 36, color: port ? C.accent1 : C.accent5, margin: 0, valign: "top", isTextBox: true });
      s.addText([
        { text: t, options: { bold: true, fontSize: 15, color: C.text1, breakLine: true } },
        { text: sub, options: { fontSize: 11, color: C.accent4 } },
      ], { x: x + 0.95, y: y + 0.04, w: 2.75, h: 0.8, fontFace: "Montserrat", margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 2 });
    });
    s.addImage({ data: I.anchorRed, x: 5.6, y: 6.62, w: 0.26, h: 0.26 });
    s.addText("In red: this year’s focus, ports & terminals", { x: 6.0, y: 6.58, w: 6.0, h: 0.34, fontFace: "Montserrat", fontSize: 11, color: C.accent1, bold: true, margin: 0, valign: "middle", isTextBox: true });
    s.addNotes("Eight stops today. Numbers in red, 02 to 05, are the heart of this year's talk: the ports and terminals that move Pakistan's cement and clinker.");
  }

  // ===== 3. ECONOMY =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "01  ·  PAKISTAN AT A GLANCE", "A STABLE ECONOMY, BACK ON A GROWTH PATH");
    const big = [
      [I.chart, "3.7%", "GDP growth, FY26", "Fastest in four years"],
      [I.bank, "US$452bn", "Size of the economy", "GDP at current prices"],
      [I.users, "252m", "People", "5th most populous nation"],
      [I.user, "US$1,901", "Income per person", "Up from US$1,751"],
    ];
    const cw = 2.84, gap = 0.25;
    big.forEach(([ic, n, l, sub], i) => {
      const x = M + i * (cw + gap), y = 1.8;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 2.75, fill: { color: C.background2 }, line: { type: "none" } });
      circleIcon(s, x + 0.3, y + 0.3, 0.62, ic);
      s.addText(n, { x: x + 0.3, y: y + 1.0, w: cw - 0.5, h: 0.85, fontFace: HEAD, fontSize: 38, color: C.text2, margin: 0, valign: "bottom", isTextBox: true });
      s.addText([
        { text: l, options: { bold: true, fontSize: 13, color: C.text1, breakLine: true } },
        { text: sub, options: { fontSize: 10.5, color: C.accent4 } },
      ], { x: x + 0.3, y: y + 1.95, w: cw - 0.5, h: 0.65, fontFace: "Montserrat", margin: 0, valign: "top", isTextBox: true });
    });
    const small = [
      [I.hand, "US$41.6bn", "Record remittances"],
      [I.vault, "US$18.4bn", "FX reserves (SBP, Jun-26)"],
      [I.star, "‘B’", "S&P upgrade, Jul-26 — first in 9 years"],
    ];
    const sw = 3.87;
    small.forEach(([ic, n, l], i) => {
      const x = M + i * (sw + 0.25), y = 4.85;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: sw, h: 1.25, fill: { color: C.background1 }, line: { color: HEX.line, width: 1 } });
      circleIcon(s, x + 0.25, y + 0.33, 0.58, ic, C.accent1);
      s.addText(n, { x: x + 1.05, y: y + 0.18, w: sw - 1.2, h: 0.55, fontFace: HEAD, fontSize: 26, color: C.accent1, margin: 0, valign: "bottom", isTextBox: true });
      s.addText(l, { x: x + 1.05, y: y + 0.76, w: sw - 1.2, h: 0.35, fontFace: "Montserrat", fontSize: 10.5, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    source(s, "Source: Pakistan Economic Survey 2025-26; State Bank of Pakistan; S&P Global. FY = July–June.", { y: 6.35, x: M, w: 12.13 });
    s.addNotes("Briefly, the backdrop. FY2026 ended in June with 3.7% growth, the fastest in four years, on a US$452 billion economy. Remittances hit a record, reserves were rebuilt, and S&P upgraded Pakistan to B in July, the first upgrade in nine years. Now, to the ports.");
  }

  // ===== 4. THREE PORTS MAP =====
  {
    const s = pres.addSlide({ masterName: "BLANK" });
    s.addImage({ path: A + "coast.jpg", x: 0, y: 0, w: SW, h: SH, altText: "Satellite map of Pakistan's southern coastline" });
    const mk = (x, y) => {
      s.addShape(pres.shapes.OVAL, { x: x - 0.17, y: y - 0.17, w: 0.34, h: 0.34, fill: { color: C.accent1, transparency: 65 }, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x - 0.09, y: y - 0.09, w: 0.18, h: 0.18, fill: { color: C.accent1 }, line: { color: "FFFFFF", width: 1.5 } });
    };
    mk(2.09, 3.71); mk(10.7, 4.29); mk(11.35, 4.41);
    pill(s, 1.35, 3.98, 1.5, "GWADAR");
    pill(s, 8.6, 4.12, 1.9, "KARACHI PORT");
    pill(s, 11.0, 4.68, 1.7, "PORT QASIM");
    s.addText("02  ·  THREE GATEWAYS", { x: M, y: 4.85, w: 7, h: 0.32, fontFace: "Montserrat", fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, margin: 0, isTextBox: true });
    s.addText("ONE COASTLINE, THREE DEEP-SEA PORTS", { x: M, y: 5.15, w: 8.5, h: 0.65, fontFace: HEAD, fontSize: 32, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "Title" });
    const cards = [
      ["GWADAR", "Natural deep water: 14 m today, 24 m planned"],
      ["KARACHI PORT", "The busiest: record 2.65m TEU in FY26"],
      ["PORT QASIM", "18 berths · 89 Mtpa capacity"],
    ];
    cards.forEach(([h, t], i) => {
      const x = M + i * 4.1, y = 5.95;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.85, h: 0.95, fill: { color: "FFFFFF" }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.15, blur: 6, offset: 2, angle: 90 } });
      s.addText([
        { text: h, options: { fontFace: HEAD, fontSize: 17, color: C.accent1, breakLine: true } },
        { text: t, options: { fontFace: "Montserrat", fontSize: 11, color: C.text1 } },
      ], { x: x + 0.22, y: y + 0.1, w: 3.5, h: 0.78, margin: 0, valign: "middle", isTextBox: true });
    });
    s.addText("≈ 470 km", { x: 5.6, y: 4.0, w: 1.2, h: 0.3, fontFace: "Montserrat", fontSize: 10, italic: true, color: C.accent4, margin: 0, align: "center", isTextBox: true });
    s.addShape(pres.shapes.LINE, { x: 2.4, y: 4.45, w: 6.1, h: 0, line: { color: HEX.grey, width: 1, dashType: "dash" } });
    s.addText("Imagery: Esri, Maxar, Earthstar Geographics", { x: 9.0, y: 7.18, w: 4.15, h: 0.22, fontFace: "Montserrat", fontSize: 7, color: C.accent4, align: "right", margin: 0, isTextBox: true });
    s.addNotes("Here is the whole coastline. On the right, Karachi Port and Port Qasim, about 30 km apart, handle nearly all of Pakistan's sea trade, and all of our cement and clinker exports. About 470 km west along the Makran coast is Gwadar, a natural deep-water port at the mouth of the Gulf.");
  }

  // ===== 5. KPT: KGTL & KGTML =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "03  ·  KARACHI PORT", "KGTL & KGTML: KARACHI’S DEEP-WATER TERMINALS");
    s.addImage({ path: A + "kpt_sat.jpg", x: M, y: 1.6, w: 12.13, h: 3.75, altText: "Satellite view of KGTL and KGTML at Karachi Port with Lucky's loose-cement berth" });
    s.addShape(pres.shapes.RECTANGLE, { x: 10.55, y: 1.75, w: 2.03, h: 0.95, fill: { color: "FFFFFF" }, line: { type: "none" } });
    s.addText("OPERATED BY", { x: 10.65, y: 1.8, w: 1.85, h: 0.22, fontFace: "Montserrat", fontSize: 7, bold: true, color: C.accent4, charSpacing: 2, margin: 0, isTextBox: true });
    s.addImage({ path: A + "kgtl_logo.png", x: 10.75, y: 2.05, w: 1.6, h: 0.64, altText: "KGTL logo" });
    const st = [
      ["1,525 m", "multipurpose quay"],
      ["15 m", "draft alongside"],
      ["100,000 t", "post-Panamax bulk ships"],
      ["30,000 t", "loaded per day"],
      ["175,000 t", "silo storage"],
      ["59 ha", "yard & storage"],
    ];
    const w = 12.13 / 6;
    st.forEach(([n, l], i) => stat(s, M + i * w, 5.55, w - 0.15, n, l, { size: 28, ls: 11, lh: 0.35 }));
    source(s, "Source: Noatum Ports (AD Ports Group), KGTL & KGTML corporate presentation, 2026.");
    s.addNotes("Karachi Port's Gateway Terminals, run by AD Ports Group. KGTL handles containers; KGTML is Pakistan's only multipurpose terminal, for clinker, cement, grain, fertiliser and steel. 1,525 metres of quay, 15 metres of water, ships up to 100,000 tonnes and a loading rate of 30,000 tonnes a day. Lucky also has its own loose-cement berth and silos at KPT.");
  }

  // ===== 6. KPT: WHAT'S NEXT =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "03  ·  KARACHI PORT  ·  WHAT’S NEXT", "BIGGER SHIPS, FASTER LOADING");
    s.addImage({ path: A + "kgtml_plan.jpg", x: M, y: 1.6, w: 6.6, h: 3.15, altText: "KGTML master plan" });
    pill(s, M + 0.15, 1.75, 2.4, "KGTML MASTER PLAN", { size: 9 });
    const news = [
      ["JUN 2026", "Channel dredging done: post-Panamax era at Karachi"],
      ["2026", "AI-based bag counting live at KGTML"],
    ];
    news.forEach(([d, t], i) => {
      const x = M + i * 3.35, y = 4.98;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.25, h: 1.65, fill: { color: C.background2 }, line: { type: "none" } });
      s.addText("NEWS  ·  " + d, { x: x + 0.22, y: y + 0.18, w: 2.9, h: 0.28, fontFace: "Montserrat", fontSize: 9, bold: true, color: C.accent1, charSpacing: 2, margin: 0, isTextBox: true });
      s.addText(t, { x: x + 0.22, y: y + 0.5, w: 2.85, h: 1.0, fontFace: "Montserrat", fontSize: 13, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    const up = [
      ["15 m", "Deeper channel", "Ships up to 100,000 t"],
      ["120,000 t", "Covered clinker shed", "9 Mt clinker a year"],
      ["8 Mt", "Clean-bulk silos & bagging", "Capacity per year"],
      ["1m TEU", "KGTL container capacity", "After 200 m quay extension"],
      [">US$200m", "AD Ports Group investment", "In KGTL & KGTML"],
    ];
    up.forEach(([n, h, sub], i) => {
      const y = 1.6 + i * 1.02;
      s.addText(n, { x: 7.6, y, w: 2.3, h: 0.8, fontFace: HEAD, fontSize: 30, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
      s.addText([
        { text: h, options: { bold: true, fontSize: 13, color: C.text1, breakLine: true } },
        { text: sub, options: { fontSize: 11, color: C.accent4 } },
      ], { x: 10.0, y, w: 2.73, h: 0.8, fontFace: "Montserrat", margin: 0, valign: "middle", isTextBox: true });
      if (i < up.length - 1) s.addShape(pres.shapes.LINE, { x: 7.6, y: y + 0.91, w: 5.13, h: 0, line: { color: HEX.line, width: 0.75 } });
    });
    source(s, "Source: KGTL & KGTML corporate presentation (2026); Profit / Business Recorder (12-Jun-2026).");
    s.addNotes("What is coming at Karachi. The channel is now dredged, so post-Panamax bulk carriers of around 100,000 tonnes can come in. A 120,000-tonne covered clinker shed takes clinker capacity to 9 million tonnes a year, and a clean-bulk facility with silos and bagging adds 8 million. Bigger ships mean lower freight per tonne to distant markets like West Africa and the Americas. And AI-based bag counting is already running at KGTML.");
  }

  // ===== 7. PORT QASIM: PIBT =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "04  ·  PORT QASIM", "PIBT: SILO TO SHIP AT PORT QASIM");
    s.addImage({ path: A + "pibt_aerial.jpg", x: M, y: 1.6, w: 7.4, h: 4.6, altText: "Aerial view of Pakistan International Bulk Terminal" });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.45, y: 1.75, w: 1.4, h: 0.5, fill: { color: "FFFFFF" }, line: { type: "none" } });
    s.addImage({ path: A + "pibt_logo.png", x: 6.55, y: 1.81, w: 1.2, h: 0.38, altText: "PIBT logo" });
    s.addText("Port Qasim:  49 km channel  ·  18 berths  ·  89 Mtpa", { x: M, y: 6.3, w: 7.4, h: 0.35, fontFace: "Montserrat", fontSize: 11, bold: true, color: C.accent4, margin: 0, valign: "middle", isTextBox: true });
    const st = [
      ["1,200 t/h", "ship loader\n≈ 28,800 t a day"],
      ["13 m", "depth alongside\n2 berths, 565 m"],
      ["50,000 t", "cement in 5 silos"],
      ["250,000 t", "clinker yard"],
      ["4 Mt", "cement & clinker a year"],
      ["20 Mt", "total, with expansion"],
    ];
    st.forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      stat(s, 8.4 + col * 2.2, 1.6 + row * 1.6, 2.1, n, l, { size: 30, ls: 11, lh: 0.55 });
    });
    source(s, "Source: Pakistan International Bulk Terminal Ltd (PIBT), terminal presentation 2026; Port Qasim Authority.");
    s.addNotes("At Port Qasim, PIBT is Pakistan's first terminal built for coal, clinker and cement. A 1,200 tonne-an-hour ship loader, roughly 28,800 tonnes a day, two berths with 13 metres alongside, five cement silos holding 50,000 tonnes and a 250,000-tonne clinker yard. Cement and clinker capacity is 4 million tonnes a year, and the whole terminal can go to 20 million.");
  }

  // ===== 8. PORT QASIM: WHAT'S NEXT =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "04  ·  PORT QASIM  ·  WHAT’S NEXT", "CEMENT GETS PRIORITY");
    s.addImage({ path: A + "pibt_dusk.jpg", x: 7.333, y: 0, w: 6.0, h: 7.5, altText: "PIBT jetty and ship loader at dusk" });
    const rows = [
      [I.anchor, "Priority berthing", "Cement goes first at Marginal Wharf I & II"],
      [I.plus, "Two new berths", "Under construction to lift export capacity"],
      [I.percent, "Port charges cut 50%", "For cement exporters, via government taskforce"],
      [I.city, "New terminal & ‘Energy City’", "Planned beside a 15,474-acre industrial zone"],
    ];
    rows.forEach(([ic, h, t], i) => {
      const y = 1.85 + i * 1.18;
      circleIcon(s, M, y, 0.72, ic);
      s.addText([
        { text: h, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } },
        { text: t, options: { fontSize: 12, color: C.text1 } },
      ], { x: M + 0.95, y: y - 0.08, w: 5.4, h: 0.9, fontFace: "Montserrat", margin: 0, valign: "middle", isTextBox: true });
    });
    source(s, "Source: Port Qasim Authority; Ministry of Maritime Affairs (May-2026).", { x: 1.5, y: 6.62, w: 5.4 });
    s.addNotes("What is changing at Port Qasim for cement: priority berthing at Marginal Wharves I and II, two new multipurpose berths under construction, port charges halved for cement exporters, and plans for a new multipurpose terminal and an Energy City next to a large industrial zone.");
  }

  // ===== 9. GWADAR =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "05  ·  GWADAR", "THE DEEP-WATER FUTURE");
    s.addImage({ path: A + "gwadar.jpg", x: 6.933, y: 0, w: 6.4, h: 7.5, altText: "Satellite view of Gwadar Port" });
    pill(s, 7.25, 6.85, 2.6, "GWADAR PORT FROM SPACE", { size: 9 });
    const st = [
      ["14 → 24 m", "natural depth, today → future phases"],
      ["11,000 TEU", "Jan–Apr 2026, vs 8,300 in all of 2025"],
      ["50%", "of public-sector cargo routed to Gwadar"],
      ["620 km", "to the Strait of Hormuz"],
    ];
    st.forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      stat(s, M + col * 3.1, 1.75 + row * 1.85, 2.85, n, l, { size: 38, ls: 11, lh: 0.55 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 5.55, w: 5.95, h: 1.05, fill: { color: C.background2 }, line: { type: "none" } });
    s.addText([
      { text: "SCOPE  ", options: { bold: true, color: C.accent1, charSpacing: 2, fontSize: 10 } },
      { text: "A transshipment hub for the Gulf & East Africa, and the sea gateway for CPEC and Balochistan’s mining projects.", options: { color: C.text1, fontSize: 12 } },
    ], { x: M + 0.25, y: 5.6, w: 5.5, h: 0.95, fontFace: "Montserrat", margin: 0, valign: "middle", isTextBox: true });
    s.addText("Source: Gwadar Port Authority; Asia Times (May-2026); PID. Imagery: Esri, Maxar.", { x: M, y: 6.66, w: 5.9, h: 0.2, fontFace: "Montserrat", fontSize: 7, color: C.accent4, margin: 0, isTextBox: true });
    s.addNotes("Gwadar is the third gateway and the long-term story. It is naturally deep, 14 metres today and planned to 24. Traffic is finally building: about 11,000 containers by April 2026 against 8,300 in all of 2025, helped by disruption around Hormuz, and government policy now routes half of public-sector cargo through Gwadar. For our industry it is a future transshipment hub for the Gulf and East Africa.");
  }

  // ===== 10. CEMENT INDUSTRY =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "06  ·  CEMENT INDUSTRY", "CAPACITY IS BUILT. EXPORTS ARE THE OUTLET");
    s.addChart([
      { type: pres.charts.BAR, data: [
        { name: "Domestic", labels: ["FY22", "FY23", "FY24", "FY25", "FY26"], values: [47.6, 40.0, 38.2, 37.9, 41.5] },
        { name: "Exports", labels: ["FY22", "FY23", "FY24", "FY25", "FY26"], values: [5.3, 4.6, 7.1, 9.2, 9.0] },
      ], options: { barGrouping: "stacked" } },
    ], {
      x: M, y: 1.6, w: 6.4, h: 3.55, barDir: "col", barGapWidthPct: 60,
      chartColors: [HEX.maroon, "8A8A8A"], dataLabelFormatCode: "0.0",
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 10, dataLabelFontFace: "Montserrat",
      showLegend: true, legendPos: "t", legendFontFace: "Montserrat", legendFontSize: 10, legendColor: HEX.ink,
      showTitle: true, title: "Despatches, million tonnes", titleFontFace: "Montserrat", titleFontSize: 11, titleColor: HEX.ink, titleBold: true,
      catAxisLabelColor: HEX.grey, catAxisLabelFontFace: "Montserrat", catAxisLabelFontSize: 10,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      catAxisLineShow: false,
    });
    const st = [["84.6 Mt", "installed capacity"], ["50.5 Mt", "despatched in FY26, +7%"], ["60%", "capacity in use"], ["9.0 Mt", "exported; two-thirds clinker"]];
    st.forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      stat(s, 7.55 + col * 2.65, 1.7 + row * 1.75, 2.5, n, l, { size: 40, ls: 11, lh: 0.5 });
    });
    const news = [
      ["FEB 2026", "Fauji Cement & KAPCO take over Attock Cement"],
      ["JUL 2026", "Maple Leaf merges Pioneer: #3 producer, 13.1 Mt"],
      ["JUL–SEP 2026", "Sea-borne exports from the south up 19%"],
    ];
    news.forEach(([d, t], i) => {
      const x = M + i * 4.12, y = 5.4;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.89, h: 1.22, fill: { color: C.background2 }, line: { type: "none" } });
      s.addText("NEWS  ·  " + d, { x: x + 0.22, y: y + 0.14, w: 3.5, h: 0.26, fontFace: "Montserrat", fontSize: 9, bold: true, color: C.accent1, charSpacing: 2, margin: 0, isTextBox: true });
      s.addText(t, { x: x + 0.22, y: y + 0.44, w: 3.5, h: 0.7, fontFace: "Montserrat", fontSize: 12, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    source(s, "Source: APCMA; Lucky Cement Annual Report 2026; Global Cement; CemNet; Pakistan Observer (Oct-2026).");
    s.addNotes("The cement industry in one slide. 84.6 million tonnes of capacity, 50.5 million despatched last year, so about 60% in use. That idle capacity is why exports matter: 9 million tonnes went abroad, two-thirds as clinker, almost all through Karachi's ports. The industry is consolidating, with Fauji taking Attock and Maple Leaf merging Pioneer, and this year's first quarter shows sea-borne exports from the south up 19%.");
  }

  // ===== 11. LUCKY: CLEAN ENERGY =====
  {
    const s = pres.addSlide({ masterName: "BLANK" });
    s.addImage({ path: A + "renew.jpg", x: 0, y: 0, w: 6.6, h: 7.5, altText: "Lucky Cement wind turbines and solar park" });
    manualHead(s, 7.2, "07  ·  LUCKY CEMENT  ·  CLEAN ENERGY", "~55% OF OUR POWER IS CLEAN", 5.6);
    const st = [
      ["56 MW", "waste heat recovery"],
      ["89.3 MW", "solar, incl. +15 MW new"],
      ["28.8 MW", "wind farm"],
      ["22.7 MWh", "battery storage — largest in Pakistan"],
      ["274,614 t", "CO₂ avoided in FY26"],
      ["I-REC", "all projects certified"],
    ];
    st.forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      stat(s, 7.2 + col * 2.9, 1.75 + row * 1.6, 2.7, n, l, { size: 34, ls: 11, lh: 0.55 });
    });
    s.addImage({ path: A + "logo.png", x: 7.2, y: 6.93, w: 0.73, h: 0.34, altText: "Lucky Cement logo" });
    source(s, "Source: Lucky Cement decarbonisation fact sheet FY2026; Annual Report 2026.", { x: 7.2, w: 5.53 });
    s.addNotes("About 55% of the energy we use now comes from clean sources: 56 MW of waste heat recovery, 89 MW of solar including 15 MW just added at Karachi, a 28.8 MW wind farm, and 22.7 MWh of battery storage, the largest in Pakistan. Together they avoided about 275,000 tonnes of CO2 last year.");
  }

  // ===== 12. LUCKY: PROJECTS =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "07  ·  LUCKY CEMENT  ·  NEXT PROJECTS", "BUILDING A FUTURE-READY COMPANY");
    const cards = [
      ["ifrs.jpg", "IFRS S1 & S2", "FY26", "First year of global sustainability & climate reporting"],
      ["lineh.jpg", "LINE H EXPANSION", "300,000 t", "More cement-grinding capacity a year (MTPA)"],
      ["plant.jpg", "UTIS HYDROGEN", "4 of 4", "Karachi kilns burn with on-site hydrogen; ~10% less fuel"],
    ];
    const cw = 3.88;
    cards.forEach(([img, h, n, t], i) => {
      const x = M + i * (cw + 0.245), y = 1.6;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 5.1, fill: { color: C.background2 }, line: { type: "none" } });
      s.addImage({ path: A + img, x, y, w: cw, h: cw * 2.3 / 3.85, altText: h });
      const iy = y + cw * 2.3 / 3.85 + 0.25;
      s.addText(h, { x: x + 0.3, y: iy, w: cw - 0.6, h: 0.32, fontFace: "Montserrat", fontSize: 11, bold: true, color: C.accent1, charSpacing: 2, margin: 0, isTextBox: true });
      s.addText(n, { x: x + 0.3, y: iy + 0.36, w: cw - 0.6, h: 0.85, fontFace: HEAD, fontSize: 44, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
      s.addText(t, { x: x + 0.3, y: iy + 1.28, w: cw - 0.6, h: 0.75, fontFace: "Montserrat", fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 9.0, y: 1.75, w: 1.0, h: 0.62, fill: { color: "000000" }, line: { type: "none" } });
    s.addImage({ path: A + "utis.png", x: 9.19, y: 1.78, w: 0.56, h: 0.56, altText: "UTIS logo" });
    s.addNotes("Three projects that make us future-ready. First, IFRS S1 and S2: FY26 is our first year reporting sustainability and climate risk to the global standard. Second, Line H: expanding cement-grinding capacity by 300,000 tonnes a year. Third, UTIS hydrogen-assisted combustion, now on all four Karachi kiln lines, designed to cut fuel use by about 10%. [Presenter: confirm Line H timing and wording with the project team.]");
  }

  // ===== 13. LUCKY USTAAD =====
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    kt(s, "08  ·  LUCKY USTAAD", "AI FOR BUILDERS, IN THEIR POCKET");
    // phone mock-up: replace the screen with the screen recording
    const px = 1.35, py = 1.55, pw = 2.75, ph = 5.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, rectRadius: 0.35, fill: { color: "151515" }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.25, blur: 12, offset: 4, angle: 90 }, objectName: "Phone body" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px + 0.12, y: py + 0.12, w: pw - 0.24, h: ph - 0.24, rectRadius: 0.26, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Phone screen — place screen recording here" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px + pw / 2 - 0.4, y: py + 0.22, w: 0.8, h: 0.16, rectRadius: 0.08, fill: { color: "151515" }, line: { type: "none" } });
    s.addImage({ data: I.robot, x: px + pw / 2 - 0.42, y: py + 1.6, w: 0.84, h: 0.84 });
    s.addText([
      { text: "LUCKY USTAAD", options: { fontFace: HEAD, fontSize: 24, color: "FFFFFF", breakLine: true } },
      { text: "Your AI construction companion", options: { fontFace: "Montserrat", fontSize: 10, color: "F0D9DC" } },
    ], { x: px + 0.2, y: py + 2.6, w: pw - 0.4, h: 0.95, align: "center", margin: 0, valign: "top", isTextBox: true });
    s.addImage({ data: I.play, x: px + pw / 2 - 0.15, y: py + 4.1, w: 0.3, h: 0.3 });
    s.addText("ESTIMATE  ·  PLAN  ·  BUILD", { x: 5.0, y: 1.7, w: 7.7, h: 0.75, fontFace: HEAD, fontSize: 36, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
    const rows = [
      [I.comments, "Ask anything", "Construction & renovation answers, in plain words"],
      [I.calc, "Smart estimates", "Cement, materials and cost for the job"],
      [I.list, "Step-by-step guidance", "Through every stage of the build"],
      [I.store, "Nearest dealer", "Product advice and dealer access in one app"],
    ];
    rows.forEach(([ic, h, t], i) => {
      const y = 2.75 + i * 0.95;
      circleIcon(s, 5.0, y, 0.62, ic);
      s.addText([
        { text: h, options: { bold: true, fontSize: 15, color: C.text2, breakLine: true } },
        { text: t, options: { fontSize: 12, color: C.text1 } },
      ], { x: 5.85, y: y - 0.1, w: 6.8, h: 0.82, fontFace: "Montserrat", margin: 0, valign: "middle", isTextBox: true });
    });
    s.addText("From the quarry to the customer’s home: AI at every step.", { x: 5.0, y: 6.5, w: 7.7, h: 0.35, fontFace: "Montserrat", fontSize: 12, italic: true, bold: true, color: C.accent1, margin: 0, isTextBox: true });
    s.addNotes("Finally, something no one else in this room is doing. Lucky Ustaad is our AI construction assistant: a builder or homeowner asks a question in plain words, gets an estimate of cement and materials, step-by-step guidance and the nearest dealer. [Play the phone screen recording here.] TO ADD THE RECORDING: Insert > Video > This Device, then size it to the dark-red phone screen (about 2.5 in x 5.0 in) and drag it on top.");
  }

  // ===== 14. THANK YOU =====
  {
    const s = pres.addSlide({ masterName: "BLANK" });
    s.addImage({ path: A + "closing.jpg", x: 6.663, y: 0, w: 6.67, h: 7.5, altText: "Lucky Cement plant at night" });
    s.addImage({ path: A + "logo.png", x: M, y: 0.6, w: 1.6, h: 0.74, altText: "Lucky Cement logo" });
    s.addText("PORT-READY.\nEXPORT-READY.\nFUTURE-READY.", { x: M, y: 2.1, w: 5.6, h: 1.75, fontFace: "Montserrat", fontSize: 20, bold: true, color: C.accent1, charSpacing: 2, margin: 0, valign: "top", isTextBox: true });
    s.addText("THANK YOU", { x: M, y: 3.9, w: 5.8, h: 1.3, fontFace: HEAD, fontSize: 80, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "Title" });
    s.addText([
      { text: "NAUMAN AHMED", options: { bold: true, color: C.text2, fontSize: 13, charSpacing: 2, breakLine: true } },
      { text: "Head of Exports, Lucky Cement Limited  ·  lucky-cement.com", options: { color: C.accent4, fontSize: 11 } },
    ], { x: M, y: 6.15, w: 5.8, h: 0.65, fontFace: "Montserrat", margin: 0, valign: "top", isTextBox: true });
    s.addNotes("Thank you. Pakistan's ports are getting deeper and faster, our industry has the capacity, and Lucky is ready for the next ship. Happy to take questions.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();

// Lucky Cement — Intercem Los Angeles 2026 (v2: sections, ports deep dive)
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const S = "/tmp/claude-0/src/";
const OUT = process.argv[2] || "/tmp/claude-0/build/deck2.pptx";

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
const HEX = { maroon: "6E0E1A", red: "A6192E", ink: "1F1F1F", grey: "8A8A8A", line: "E2DEDE", rose: "C8A2A8" };
const HEAD = "Anton", BODY = "Montserrat";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "Lucky Cement Limited";
pres.company = "Lucky Cement Limited";
pres.title = "From Quarry to Quay — Intercem Los Angeles 2026";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const SW = 13.333, SH = 7.5, M = 0.6;

// ---------- images ----------
const meta = {};
async function size(src) { if (!meta[src]) meta[src] = await sharp(S + src).metadata(); return meta[src]; }
// crop a photo to exactly fill w x h (inches) around a focus point; never stretch
async function cropData(src, w, h, fx = 0.5, fy = 0.5, zoom = 1) {
  const m = await size(src);
  const r = w / h;
  let cw = m.width / zoom, ch = cw / r;
  if (ch > m.height / zoom) { ch = m.height / zoom; cw = ch * r; }
  const left = Math.round((m.width - cw) * fx), top = Math.round((m.height - ch) * fy);
  const outW = Math.min(Math.round(cw), Math.round(w * 220));
  const buf = await sharp(S + src).rotate().extract({ left, top, width: Math.round(cw), height: Math.round(ch) })
    .resize(outW, Math.round(outW / r)).jpeg({ quality: 88 }).toBuffer();
  return "image/jpeg;base64," + buf.toString("base64");
}
async function photo(slide, src, x, y, w, h, o = {}) {
  slide.addImage({ data: await cropData(src, w, h, o.fx ?? 0.5, o.fy ?? 0.5, o.zoom ?? 1), x, y, w, h, altText: o.alt || src, objectName: o.alt || src });
}
// place a logo inside a box, keeping its aspect ratio
async function logo(slide, src, x, y, w, h, align = "center") {
  const m = await size(src); const r = m.width / m.height;
  let lw = w, lh = w / r; if (lh > h) { lh = h; lw = h * r; }
  const lx = align === "left" ? x : x + (w - lw) / 2;
  slide.addImage({ path: S + src, x: lx, y: y + (h - lh) / 2, w: lw, h: lh, altText: src.replace(".png", "") + " logo" });
}
async function icon(Comp, color = "#FFFFFF", sz = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color, size: String(sz) }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.42, w: 9.5, h: 0.32, fontFace: BODY, fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, margin: 0, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.74, w: 12.1, h: 0.72, fontFace: HEAD, fontSize: 32, color: C.text2, margin: 0, valign: "top", align: "left" }, text: "" } },
    { image: { path: S + "logo.png", x: M, y: 6.93, w: 0.73, h: 0.34, altText: "Lucky Cement logo" } },
    { text: { text: "INTERCEM LOS ANGELES  ·  NOVEMBER 2026", options: { x: 1.5, y: 6.95, w: 5, h: 0.3, fontFace: BODY, fontSize: 8, color: C.accent4, charSpacing: 2, margin: 0, valign: "middle" } } },
  ],
  slideNumber: { x: 12.33, y: 6.95, w: 0.4, h: 0.3, fontFace: BODY, fontSize: 9, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({ title: "BLANK", background: { color: C.background1 }, objects: [] });

// ---------- helpers ----------
const T = (slide, text, o) => slide.addText(text, { margin: 0, isTextBox: true, fontFace: BODY, ...o });
function kt(slide, kicker, title) { slide.addText(kicker, { placeholder: "kicker" }); slide.addText(title, { placeholder: "title" }); }
function manualHead(slide, x, kicker, title, w = 5.6) {
  T(slide, kicker, { x, y: 0.42, w, h: 0.32, fontSize: 11, bold: true, color: C.accent1, charSpacing: 3, valign: "middle", objectName: "Kicker" });
  T(slide, title, { x, y: 0.74, w, h: 1.25, fontFace: HEAD, fontSize: 32, color: C.text2, valign: "top", objectName: "Title" });
}
function footLogo(slide, x = M) { slide.addImage({ path: S + "logo.png", x, y: 6.93, w: 0.73, h: 0.34, altText: "Lucky Cement logo" }); }
function stat(slide, x, y, w, num, label, o = {}) {
  const ns = o.size || 34, nh = o.nh || ns / 72 * 1.25;
  T(slide, num, { x, y, w, h: nh, fontFace: HEAD, fontSize: ns, color: o.color || C.text2, valign: "bottom", objectName: "Stat " + num });
  T(slide, label, { x, y: y + nh + 0.05, w, h: o.lh || 0.5, fontSize: o.ls || 11, color: o.lcolor || C.text1, valign: "top" });
}
function pill(slide, x, y, w, text, o = {}) {
  const h = o.h || 0.34;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: o.fill || C.background1 }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 4, offset: 1, angle: 90 } });
  T(slide, text, { x, y, w, h, fontSize: o.size || 9, bold: true, color: o.color || C.text2, align: "center", valign: "middle", charSpacing: 1 });
}
function source(slide, text, o = {}) {
  T(slide, text, { x: o.x ?? 6.6, y: o.y ?? 6.95, w: o.w ?? 5.6, h: 0.3, fontSize: 7, color: C.accent4, align: o.align || "right", valign: "middle", objectName: "Source" });
}
function circleIcon(slide, x, y, d, data, fill = C.accent2) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  const p = d * 0.25; slide.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
function rowText(slide, x, y, w, h, head, sub, o = {}) {
  T(slide, [
    { text: head, options: { bold: true, fontSize: o.hs || 15, color: o.hc || C.text2, breakLine: true } },
    { text: sub, options: { fontSize: o.ss || 12, color: C.text1 } },
  ], { x, y, w, h, valign: o.valign || "middle" });
}
function card(slide, x, y, w, h, fill = C.background2) { slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" } }); }
function marker(slide, x, y, big = false) {
  const r1 = big ? 0.22 : 0.17, r2 = big ? 0.11 : 0.09;
  slide.addShape(pres.shapes.OVAL, { x: x - r1, y: y - r1, w: 2 * r1, h: 2 * r1, fill: { color: C.accent1, transparency: 65 }, line: { type: "none" } });
  slide.addShape(pres.shapes.OVAL, { x: x - r2, y: y - r2, w: 2 * r2, h: 2 * r2, fill: { color: C.accent1 }, line: { color: "FFFFFF", width: 1.5 } });
}
function credit(slide, text, x, y, w, o = {}) {
  T(slide, text, { x, y, w, h: 0.2, fontSize: 6.5, color: o.color || "FFFFFF", align: o.align || "right", valign: "middle", objectName: "Photo credit" });
}

// section divider: photo on the right, number and title on white
async function divider(num, title, sub, items, img, o = {}) {
  pres.addSection({ title });
  const s = pres.addSlide({ masterName: "BLANK", sectionTitle: title });
  await photo(s, img, 5.333, 0, 8.0, 7.5, { fx: o.fx, fy: o.fy, zoom: o.zoom, alt: title + " section photo" });
  T(s, num, { x: M, y: 0.7, w: 4.4, h: 1.9, fontFace: HEAD, fontSize: 120, color: C.accent3, valign: "top" });
  T(s, title, { x: M, y: 2.75, w: 4.5, h: 1.6, fontFace: HEAD, fontSize: 44, color: C.text2, valign: "top", lineSpacingMultiple: 0.95 });
  T(s, sub, { x: M, y: 4.35, w: 4.3, h: 0.7, fontSize: 14, color: C.text1, valign: "top" });
  T(s, items.map((t, i) => ({ text: t, options: { breakLine: i < items.length - 1 } })), { x: M, y: 5.2, w: 4.4, h: 1.4, fontSize: 11, color: C.accent4, valign: "top", paraSpaceAfter: 4 });
  footLogo(s);
  if (o.credit) credit(s, o.credit, 8.33, 7.22, 4.9);
  s.addNotes(o.notes || `Section ${num}: ${title}.`);
  return s;
}

(async () => {
  const I = {};
  for (const [k, comp] of Object.entries({
    chart: fa.FaChartLine, bank: fa.FaLandmark, users: fa.FaUsers, user: fa.FaUserTie, hand: fa.FaHandHoldingUsd,
    vault: fa.FaPiggyBank, star: fa.FaStar, anchor: fa.FaAnchor, ship: fa.FaShip, plus: fa.FaPlusCircle, percent: fa.FaPercent,
    city: fa.FaCity, robot: fa.FaRobot, calc: fa.FaCalculator, list: fa.FaListOl, store: fa.FaStore, comments: fa.FaComments,
    play: fa.FaPlay, water: fa.FaWater, radar: fa.FaSatelliteDish, industry: fa.FaIndustry, arrow: fa.FaArrowRight,
  })) I[k] = await icon(comp);
  I.arrowRed = await icon(fa.FaArrowRight, "#A6192E");
  const CR_VG = "Photo: VileGecko / Wikimedia Commons, CC BY-SA 4.0";

  // =================== 1. TITLE ===================
  {
    pres.addSection({ title: "Opening" });
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: "Opening" });
    await photo(s, "idc.jpg", 6.2, 0, 7.133, 7.5, { fx: 0.5, fy: 0.5, alt: "Bulk carrier entering Port Qasim" });
    credit(s, CR_VG, 8.3, 7.22, 4.9);
    s.addImage({ path: S + "logo.png", x: M, y: 0.6, w: 1.6, h: 0.74, altText: "Lucky Cement logo" });
    T(s, "INTERCEM  ·  LOS ANGELES  ·  NOVEMBER 2026", { x: M, y: 2.25, w: 5.5, h: 0.35, fontSize: 11, bold: true, color: C.accent1, charSpacing: 2 });
    T(s, "FROM QUARRY\nTO QUAY", { x: M, y: 2.65, w: 5.5, h: 2.25, fontFace: HEAD, fontSize: 70, color: C.text2, valign: "top", lineSpacingMultiple: 0.95, objectName: "Title" });
    T(s, "Pakistan’s ports, dry bulk terminals and the cement that moves through them", { x: M, y: 4.95, w: 5.1, h: 0.75, fontSize: 15, color: C.text1, valign: "top" });
    T(s, [
      { text: "NAUMAN AHMED", options: { bold: true, color: C.text2, fontSize: 13, charSpacing: 2, breakLine: true } },
      { text: "Head of Exports, Lucky Cement Limited", options: { color: C.accent4, fontSize: 11 } },
    ], { x: M, y: 6.15, w: 5.2, h: 0.65, valign: "top" });
    s.addNotes("Welcome. 'From Quarry to Quay': how Pakistan's cement and clinker move from the kiln, across the quay and onto the ship. This year we go deep into the ports and their dry bulk terminals.");
  }

  // =================== 2. AGENDA ===================
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
    kt(s, "AGENDA", "FIVE STOPS, ONE VOYAGE");
    const items = [
      ["01", "Pakistan at a glance", "pk_city.jpg", 0.6],
      ["02", "Ports & terminals", "n36.jpg", 0.5],
      ["03", "Gwadar", "gw_aerial.jpg", 0.5],
      ["04", "Cement industry", "n16.jpg", 0.5],
      ["05", "Lucky Cement", "n03.jpg", 0.5],
    ];
    const cw = 2.25, gap = (12.13 - 5 * cw) / 4;
    for (let i = 0; i < items.length; i++) {
      const [n, t, img, fx] = items[i], x = M + i * (cw + gap);
      await photo(s, img, x, 1.7, cw, 3.55, { fx, alt: t });
      T(s, n, { x, y: 5.4, w: cw, h: 0.7, fontFace: HEAD, fontSize: 32, color: i === 1 || i === 2 ? C.accent1 : C.accent3, valign: "top" });
      T(s, t, { x, y: 6.05, w: cw, h: 0.6, fontSize: 14, bold: true, color: C.text1, valign: "top" });
    }
    s.addNotes("Five stops. We touch briefly on Pakistan, then spend most of our time on the ports and terminals, including Gwadar, before the cement industry and Lucky Cement.");
  }

  // =================== SECTION 01: PAKISTAN ===================
  await divider("01", "PAKISTAN AT A GLANCE", "A large, young economy back on a growth path.", ["The economy in numbers", "Trade, remittances & reserves"], "pk_city.jpg",
    { fx: 0.6, credit: "Photo: Col Mian Zahid / Wikimedia Commons, CC BY-SA 4.0" });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PAKISTAN AT A GLANCE" });
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
      card(s, x, y, cw, 2.75);
      circleIcon(s, x + 0.3, y + 0.3, 0.62, ic);
      T(s, n, { x: x + 0.3, y: y + 1.0, w: cw - 0.5, h: 0.85, fontFace: HEAD, fontSize: 38, color: C.text2, valign: "bottom" });
      rowText(s, x + 0.3, y + 1.95, cw - 0.5, 0.65, l, sub, { hs: 13, ss: 10.5, hc: C.text1, valign: "top" });
    });
    [[I.hand, "US$41.6bn", "Record remittances"], [I.vault, "US$18.4bn", "FX reserves (SBP, Jun-26)"], [I.star, "‘B’", "S&P upgrade, Jul-26: first in 9 years"]].forEach(([ic, n, l], i) => {
      const sw = 3.87, x = M + i * (sw + 0.26), y = 4.85;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: sw, h: 1.25, fill: { color: "FFFFFF" }, line: { color: HEX.line, width: 1 } });
      circleIcon(s, x + 0.25, y + 0.33, 0.58, ic, C.accent1);
      T(s, n, { x: x + 1.05, y: y + 0.18, w: sw - 1.2, h: 0.55, fontFace: HEAD, fontSize: 26, color: C.accent1, valign: "bottom" });
      T(s, l, { x: x + 1.05, y: y + 0.76, w: sw - 1.2, h: 0.35, fontSize: 10.5, color: C.text1, valign: "top" });
    });
    source(s, "Source: Pakistan Economic Survey 2025-26; State Bank of Pakistan; S&P Global. FY = July–June.", { y: 6.35, x: M, w: 12.13 });
    s.addNotes("The backdrop, briefly. FY2026 ended in June with 3.7% growth, the fastest in four years, on a US$452 billion economy and 252 million people. Remittances hit a record, reserves were rebuilt, and S&P upgraded Pakistan to B in July. Now, to the ports.");
  }

  // =================== SECTION 02: PORTS & TERMINALS ===================
  await divider("02", "PORTS & TERMINALS", "Where Pakistan’s cement and clinker meet the sea.", ["Karachi Port & KGTML", "Bagged and loose cement at the quay", "Port Qasim & PIBT", "Terminals side by side"], "inthira.jpg",
    { fx: 0.55, credit: CR_VG, notes: "The heart of today's talk: the ports and their dry bulk and breakbulk terminals." });

  // ---- map
  {
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: "PORTS & TERMINALS" });
    s.addImage({ path: S + "coast.jpg", x: 0, y: 0, w: SW, h: SH, altText: "Satellite map of Pakistan's southern coastline" });
    marker(s, 2.09, 3.71); marker(s, 10.7, 4.29); marker(s, 11.35, 4.41);
    pill(s, 1.35, 3.98, 1.5, "GWADAR", { size: 10 });
    pill(s, 8.6, 4.12, 1.9, "KARACHI PORT", { size: 10 });
    pill(s, 11.0, 4.68, 1.7, "PORT QASIM", { size: 10 });
    T(s, "≈ 470 km", { x: 5.6, y: 4.0, w: 1.2, h: 0.3, fontSize: 10, italic: true, color: C.accent4, align: "center" });
    s.addShape(pres.shapes.LINE, { x: 2.4, y: 4.45, w: 6.1, h: 0, line: { color: "8A8A8A", width: 1, dashType: "dash" } });
    T(s, "02  ·  PORTS & TERMINALS", { x: M, y: 4.85, w: 7, h: 0.32, fontSize: 11, bold: true, color: C.accent1, charSpacing: 3 });
    T(s, "ONE COASTLINE, THREE DEEP-SEA PORTS", { x: M, y: 5.15, w: 8.5, h: 0.65, fontFace: HEAD, fontSize: 32, color: C.text2, valign: "top", objectName: "Title" });
    [["GWADAR", "Natural deep water, 630 km from Hormuz"], ["KARACHI PORT", "Since 1887 · 54 Mt of cargo in FY25"], ["PORT QASIM", "18 berths · 89 Mt a year capacity"]].forEach(([h, t], i) => {
      const x = M + i * 4.1, y = 5.95;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.85, h: 0.95, fill: { color: "FFFFFF" }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.15, blur: 6, offset: 2, angle: 90 } });
      T(s, [{ text: h, options: { fontFace: HEAD, fontSize: 17, color: C.accent1, breakLine: true } }, { text: t, options: { fontSize: 11, color: C.text1 } }], { x: x + 0.22, y: y + 0.1, w: 3.5, h: 0.78, valign: "middle" });
    });
    T(s, "Imagery: Esri, Maxar, Earthstar Geographics", { x: 9.0, y: 7.18, w: 4.15, h: 0.22, fontSize: 7, color: C.accent4, align: "right" });
    s.addNotes("The whole coastline. On the right, Karachi Port and Port Qasim, about 30 km apart, handle nearly all of Pakistan's sea trade and all of our cement and clinker exports. About 470 km west is Gwadar, a natural deep-water port 630 km from the Strait of Hormuz.");
  }

  // ---- KPT vs PQA
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  PORTS & TERMINALS", "TWO PORTS, 30 KM APART");
    const cols = [
      ["KARACHI PORT", "n08.jpg", "Breakbulk vessel at KGTML, Karachi Port", [["1887", "in service since"], ["54 Mt", "cargo handled, FY25"], ["30", "dry-cargo berths"], ["15.5 m", "deepest berths, East Wharf"]]],
      ["PORT QASIM", "pqa_pibt.jpg", "PIBT jetty, Port Qasim", [["1980", "in service since"], ["89 Mt", "capacity a year"], ["18", "berths, 15 run privately"], ["12.5 m", "channel, going to 15 m"]]],
    ];
    for (let c = 0; c < 2; c++) {
      const [name, img, alt, st] = cols[c], x = M + c * 6.2;
      await photo(s, img, x, 1.6, 5.93, 2.85, { alt });
      pill(s, x + 0.15, 1.75, 1.9, name, { size: 10 });
      st.forEach(([n, l], i) => stat(s, x + i * 1.5, 4.7, 1.42, n, l, { size: 28, ls: 10.5, lh: 0.55 }));
    }
    s.addShape(pres.shapes.LINE, { x: 6.665, y: 1.6, w: 0, h: 4.3, line: { color: HEX.line, width: 1 } });
    source(s, "Source: Karachi Port Trust / Arab News (FY25 cargo); Port Qasim Authority development plan (Oct-2026).");
    s.addNotes("Karachi Port, since 1887, handled 54 million tonnes of cargo in FY25 across 30 dry-cargo berths, with East Wharf berths dredged to 15.5 metres. Port Qasim, since 1980, has 18 berths, 15 of them run by private operators, and 89 million tonnes of annual capacity. Its 12.5-metre channel is being deepened to 15 metres.");
  }

  // ---- KGTML
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  KARACHI PORT  ·  KGTML", "PAKISTAN’S ONLY MULTIPURPOSE TERMINAL");
    await photo(s, "kpt_sat.jpg", M, 1.6, 12.13, 3.75, { alt: "Satellite view of KGTL and KGTML at Karachi Port" });
    s.addShape(pres.shapes.RECTANGLE, { x: 10.55, y: 1.75, w: 2.03, h: 0.95, fill: { color: "FFFFFF" }, line: { type: "none" } });
    T(s, "OPERATED BY", { x: 10.65, y: 1.8, w: 1.85, h: 0.22, fontSize: 7, bold: true, color: C.accent4, charSpacing: 2 });
    await logo(s, "kgtl_logo.png", 10.75, 2.03, 1.6, 0.62);
    pill(s, M + 0.15, 4.85, 3.3, "KGTML: CLINKER, CEMENT & BREAKBULK", { size: 9 });
    const st = [["1,525 m", "of quay"], ["15 m", "draft alongside"], ["100,000 t", "post-Panamax bulk ships"], ["30,000 t", "loaded per day"], ["175,000 t", "silo storage"], ["13 Mt", "cargo a year"]];
    const w = 12.13 / 6;
    st.forEach(([n, l], i) => stat(s, M + i * w, 5.55, w - 0.15, n, l, { size: 28, ls: 11, lh: 0.35 }));
    source(s, "Source: Noatum Ports (AD Ports Group), KGTL & KGTML corporate presentation 2026; Arab News.");
    s.addNotes("KGTML at Karachi's East Wharf, run by AD Ports Group, is Pakistan's only multipurpose terminal: clinker, cement, grain, fertiliser, steel and project cargo. 1,525 metres of quay, 15 metres of water, bulk ships up to 100,000 tonnes and a loading rate of 30,000 tonnes a day. It handles more than 13 million tonnes a year.");
  }

  // ---- Breakbulk gallery
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  BREAKBULK AT THE QUAY", "BAGGED CEMENT: FROM QUAY TO HOLD");
    await photo(s, "n36.jpg", M, 1.6, 6.3, 5.05, { fx: 0.55, alt: "Jumbo bags of cement on the quay beside a geared bulk carrier" });
    await photo(s, "n34.jpg", 7.1, 1.6, 2.75, 2.45, { fx: 0.5, alt: "Ship's cranes lifting jumbo bags" });
    await photo(s, "n35.jpg", 9.98, 1.6, 2.75, 2.45, { fy: 0.6, alt: "Jumbo bags stowed in the hold" });
    await photo(s, "n28.jpg", 7.1, 4.2, 5.63, 2.45, { fy: 0.6, alt: "Geared bulk carrier" });
    pill(s, M + 0.15, 6.15, 1.9, "1  ·  ON THE QUAY");
    pill(s, 7.22, 3.58, 1.6, "2  ·  LIFTED");
    pill(s, 10.1, 3.58, 1.6, "3  ·  STOWED");
    pill(s, 7.22, 6.15, 2.5, "4  ·  GEARED BULKER SAILS");
    s.addNotes("Breakbulk in action. Bagged cement arrives at the quay in jumbo bags, is lifted aboard by the ship's own cranes, stowed tightly in the hold, and the geared bulker sails. This flexibility lets us serve smaller ports in Africa and the Indian Ocean that have no bulk-handling equipment.");
  }

  // ---- Loose cement
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  LOOSE CEMENT AT KARACHI PORT", "OUR OWN SILOS AT THE QUAY");
    await photo(s, "n46.jpg", M, 1.6, 7.2, 5.05, { fx: 0.3, alt: "Lucky Cement steel silos and a bulk cement carrier at Karachi Port" });
    await photo(s, "n45.jpg", 8.05, 1.6, 2.25, 2.45, { fx: 0.75, alt: "Night-time pneumatic loading" });
    await photo(s, "n47.jpg", 10.48, 1.6, 2.25, 2.45, { fx: 0.5, fy: 0.3, alt: "Pneumatic ship loader" });
    const st = [["4", "steel silos"], ["24,000 t", "storage at the berth"], ["10.5 m", "Berth 25, dredged"], ["110,000 t", "loose cement shipped, FY26"]];
    st.forEach(([n, l], i) => { const col = i % 2, row = Math.floor(i / 2); stat(s, 8.05 + col * 2.43, 4.3 + row * 1.2, 2.3, n, l, { size: 28, ls: 10.5, lh: 0.35 }); });
    pill(s, M + 0.15, 1.75, 4.6, "THE ONLY CEMENT COMPANY WITH ITS OWN SILOS AT KPT", { size: 8 });
    source(s, "Source: Lucky Cement port data; Lucky Cement export performance review FY26.");
    s.addNotes("Loose cement is shipped from our own terminal at Karachi Port: four steel silos holding 24,000 tonnes and pneumatic loading straight into the hold. Berth 25 has been dredged to 10.5 metres, and loose cement exports restarted in FY26 after two and a half years, with 110,000 tonnes shipped.");
  }

  // ---- KPT what's next
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  KARACHI PORT  ·  WHAT’S NEXT", "BIGGER SHIPS, FASTER LOADING");
    await photo(s, "n42.jpg", M, 1.6, 3.2, 3.05, { alt: "Dredging at Karachi Port" });
    await photo(s, "kgtml_plan.jpg", 3.95, 1.6, 3.25, 3.05, { fx: 0.55, alt: "KGTML master plan" });
    pill(s, M + 0.12, 4.15, 1.5, "DREDGING");
    pill(s, 4.07, 4.15, 1.9, "MASTER PLAN");
    [["JUN 2026", "Channel dredging done: the post-Panamax era starts at Karachi"], ["2026", "AI-based bag counting live at KGTML"]].forEach(([d, t], i) => {
      const x = M + i * 3.35, y = 4.9;
      card(s, x, y, 3.25, 1.75);
      T(s, "NEWS  ·  " + d, { x: x + 0.22, y: y + 0.18, w: 2.9, h: 0.28, fontSize: 9, bold: true, color: C.accent1, charSpacing: 2 });
      T(s, t, { x: x + 0.22, y: y + 0.5, w: 2.85, h: 1.1, fontSize: 13, bold: true, color: C.text1, valign: "top" });
    });
    const up = [["120,000 t", "Bulk ships after dredging", "Up from ~60,000 t"], ["120,000 t", "Covered clinker shed", "9 Mt clinker a year"], ["8 Mt", "Clean-bulk silos & bagging", "Capacity a year"], ["48,000 t", "Planned automated loading", "Per day, 2 shore units"], [">US$200m", "AD Ports Group investment", "In KGTL & KGTML"]];
    up.forEach(([n, h, sub], i) => {
      const y = 1.6 + i * 1.02;
      T(s, n, { x: 7.6, y, w: 2.3, h: 0.8, fontFace: HEAD, fontSize: 28, color: C.text2, valign: "middle" });
      rowText(s, 10.0, y, 2.73, 0.8, h, sub, { hs: 12.5, ss: 10.5, hc: C.text1 });
      if (i < up.length - 1) s.addShape(pres.shapes.LINE, { x: 7.6, y: y + 0.91, w: 5.13, h: 0, line: { color: HEX.line, width: 0.75 } });
    });
    source(s, "Source: KGTL & KGTML corporate presentation (2026); Arab News; Profit / Business Recorder (12-Jun-2026).");
    s.addNotes("What is coming at Karachi. Dredging is done, so bulk carriers of up to about 120,000 tonnes can now berth, roughly double the old limit. A 120,000-tonne covered clinker shed, a clean-bulk facility with silos and bagging, and planned automated loading of about 48,000 tonnes a day. Bigger ships mean lower freight per tonne to distant markets, including the Americas.");
  }

  // ---- Port Qasim terminals chart
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  PORT QASIM", "FIVE DRY BULK TERMINALS, ONE CHANNEL");
    await photo(s, "pqa_bulk.jpg", 8.2, 1.6, 4.53, 2.6, { alt: "Bulk carrier at Port Qasim" });
    await photo(s, "pqa_pibt.jpg", 8.2, 4.35, 4.53, 2.3, { alt: "PIBT jetty, Port Qasim" });
    s.addChart(pres.charts.BAR, [{ name: "Capacity", labels: ["PIBT: coal, clinker & cement", "PQEPC: coal", "Marginal Wharves: bulk & general", "FAP: grain & fertiliser", "IOCB: iron ore & coal"], values: [8.0, 4.5, 4.0, 4.0, 3.03] }], {
      x: M, y: 1.6, w: 7.2, h: 4.2, barDir: "bar", barGapWidthPct: 45, chartColors: [HEX.maroon, HEX.rose, HEX.rose, HEX.rose, HEX.rose],
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.ink, dataLabelFontSize: 12, dataLabelFontFace: BODY, dataLabelFontBold: true, dataLabelFormatCode: "0.0",
      showTitle: true, title: "Annual capacity, million tonnes", titleFontFace: BODY, titleFontSize: 11, titleColor: HEX.ink, titleBold: true,
      catAxisLabelColor: HEX.ink, catAxisLabelFontFace: BODY, catAxisLabelFontSize: 11, catAxisOrientation: "maxMin",
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLineShow: false, showLegend: false,
    });
    T(s, [{ text: "18 berths  ·  89 Mt a year  ·  ", options: { bold: true, color: C.text2 } }, { text: "Pakistan’s largest industrial zone, 14,590 acres", options: { color: C.text1 } }], { x: M, y: 6.0, w: 7.2, h: 0.4, fontSize: 12, valign: "middle" });
    source(s, "Source: Port Qasim Authority, development plan presentation (20-Oct-2026).", { x: M, y: 6.4, w: 7.2, align: "left" });
    s.addNotes("Port Qasim has five terminals that handle dry bulk. The largest for our industry is PIBT, built for coal, clinker and cement, at 8 million tonnes a year. Then a coal terminal for the power plant, the two Marginal Wharves for bulk and general cargo where cement gets priority berthing, the grain and fertiliser terminal, and the original iron-ore and coal berth.");
  }

  // ---- PIBT overview
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  PORT QASIM  ·  PIBT", "PAKISTAN’S FIRST DRY BULK TERMINAL");
    await photo(s, "pibt004.jpg", M, 1.6, 7.4, 4.6, { alt: "Aerial view of Pakistan International Bulk Terminal" });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.45, y: 1.75, w: 1.4, h: 0.5, fill: { color: "FFFFFF" }, line: { type: "none" } });
    await logo(s, "pibt_logo.png", 6.55, 1.8, 1.2, 0.4);
    T(s, "30-year BOT concession from Port Qasim Authority", { x: M, y: 6.3, w: 7.4, h: 0.35, fontSize: 11, bold: true, color: C.accent4, valign: "middle" });
    [["1,200 t/h", "ship loader\n≈ 28,800 t a day"], ["13 m", "depth alongside\n2 berths, 565 m"], ["50,000 t", "cement in 5 silos"], ["250,000 t", "clinker yard"], ["4 Mt", "cement & clinker a year"], ["20 Mt", "total, with expansion"]].forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2); stat(s, 8.4 + col * 2.2, 1.6 + row * 1.6, 2.1, n, l, { size: 30, ls: 11, lh: 0.55 });
    });
    source(s, "Source: Pakistan International Bulk Terminal Ltd (PIBT), terminal presentation 2026.");
    s.addNotes("PIBT is Pakistan's first terminal built for coal, clinker and cement, on a 30-year concession. A 1,200 tonne-an-hour ship loader, about 28,800 tonnes a day, two berths with 13 metres alongside, five cement silos holding 50,000 tonnes and a 250,000-tonne clinker yard.");
  }

  // ---- PIBT flow
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  PORT QASIM  ·  PIBT", "SILO TO SHIP IN FOUR STEPS");
    const steps = [["pibt016.jpg", "1", "Bulkers feed the silos", "4 feeding points per silo, 100 t/h each"], ["pibt020.jpg", "2", "Five silos, 50,000 t", "IBAU Germany, central-cone design"], ["pibt007.jpg", "3", "Conveyor to the jetty", "Belt runs at 1,200 t/h"], ["pibt006.jpg", "4", "Ship loader fills the hold", "Cement and clinker, 1,200 t/h"]];
    const cw = 2.8, gap = (12.13 - 4 * cw) / 3;
    for (let i = 0; i < 4; i++) {
      const [img, n, h, sub] = steps[i], x = M + i * (cw + gap);
      await photo(s, img, x, 1.6, cw, 2.9, { alt: h });
      s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 1.75, w: 0.5, h: 0.5, fill: { color: C.accent1 }, line: { color: "FFFFFF", width: 1.5 } });
      T(s, n, { x: x + 0.15, y: 1.75, w: 0.5, h: 0.5, fontFace: HEAD, fontSize: 18, color: "FFFFFF", align: "center", valign: "middle" });
      rowText(s, x, 4.65, cw, 0.85, h, sub, { hs: 13, ss: 10.5, valign: "top" });
      if (i < 3) s.addImage({ data: I.arrowRed, x: x + cw + gap / 2 - 0.12, y: 2.93, w: 0.24, h: 0.24 });
    }
    card(s, M, 5.65, 12.13, 0.9);
    T(s, [{ text: "CLINKER  ", options: { bold: true, color: C.accent1, charSpacing: 2, fontSize: 11 } }, { text: "Wheel loaders feed 7 hoppers onto belt BC04 to the jetty. The export yard holds about 250,000 t.", options: { color: C.text1, fontSize: 12 } }], { x: M + 0.3, y: 5.65, w: 11.5, h: 0.9, valign: "middle" });
    source(s, "Source: Pakistan International Bulk Terminal Ltd (PIBT), terminal presentation 2026.");
    s.addNotes("How cement leaves PIBT. Bulk trucks feed five German-built silos, a conveyor carries the cement to the jetty at 1,200 tonnes an hour, and the ship loader fills the hold. Clinker goes from the yard by wheel loader through seven hoppers onto the same jetty belt.");
  }

  // ---- PQA what's next
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  PORT QASIM  ·  WHAT’S NEXT", "A DEEPER, WIDER CHANNEL AND A NEW TERMINAL");
    await photo(s, "pqa_mpt.jpg", M, 1.6, 6.2, 3.3, { alt: "Planned multipurpose cargo terminal at Port Qasim" });
    pill(s, M + 0.15, 1.75, 3.6, "PLANNED MULTIPURPOSE TERMINAL", { size: 9 });
    [["16 Mt", "a year capacity"], ["75,000 DWT", "ships at 14 m"], ["US$360m", "30-year BOT"]].forEach(([n, l], i) => stat(s, M + i * 2.1, 5.1, 2.0, n, l, { size: 28, ls: 11, lh: 0.35 }));
    // channel before / after
    T(s, "NAVIGATION CHANNEL", { x: 7.3, y: 1.6, w: 5.4, h: 0.3, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 });
    const bar = (y, label, w, val, col) => {
      T(s, label, { x: 7.3, y, w: 1.3, h: 0.45, fontSize: 11, color: C.text1, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: 8.6, y: y + 0.07, w, h: 0.32, fill: { color: col }, line: { type: "none" } });
      T(s, val, { x: 8.6 + w + 0.1, y, w: 1.5, h: 0.45, fontFace: HEAD, fontSize: 18, color: C.text2, valign: "middle" });
    };
    bar(2.0, "Width today", 1.65, "150 m", HEX.rose); bar(2.5, "Width next", 2.75, "250 m", HEX.maroon);
    bar(3.1, "Depth today", 2.3, "12.5 m", HEX.rose); bar(3.6, "Depth next", 2.76, "15 m", HEX.maroon);
    T(s, "Capital dredging Phase I: May 2026 – May 2028", { x: 7.3, y: 4.15, w: 5.4, h: 0.3, fontSize: 10.5, italic: true, color: C.accent4 });
    const rows = [[I.anchor, "Priority berthing for cement", "At Marginal Wharves I & II"], [I.percent, "Port charges cut 50%", "For cement exporters"], [I.radar, "Vessel traffic system", "VTMS due May 2027"]];
    rows.forEach(([ic, h, t], i) => { const y = 4.6 + i * 0.7; circleIcon(s, 7.3, y, 0.5, ic); rowText(s, 7.95, y - 0.05, 4.8, 0.6, h, t, { hs: 12.5, ss: 10.5 }); });
    source(s, "Source: Port Qasim Authority, development plan presentation (20-Oct-2026).");
    s.addNotes("Port Qasim's plan. Capital dredging, started May 2026, takes the channel from 150 to 250 metres wide and from 12.5 to 15 metres deep. A new 16-million-tonne multipurpose cargo terminal, a US$360 million BOT project, will take 75,000-tonne ships at 14 metres. Cement already gets priority berthing and port charges were halved.");
  }

  // ---- Terminal comparison
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "PORTS & TERMINALS" });
    kt(s, "02  ·  TERMINALS SIDE BY SIDE", "WHICH QUAY FOR WHICH SHIP");
    const cols = [["TERMINAL", 2.6], ["DRAFT", 1.6], ["LARGEST SHIP", 2.2], ["LOADING RATE", 2.4], ["BEST FOR", 3.33]];
    let x = M; cols.forEach(([h, w]) => { T(s, h, { x, y: 1.65, w, h: 0.3, fontSize: 9, bold: true, color: C.accent4, charSpacing: 2 }); x += w; });
    const rows = [
      ["KGTML", "Karachi Port", "15 m", "120,000 t", "30,000 t/day", "Clinker, bagged cement, breakbulk"],
      ["LUCKY SILOS", "Karachi Port, Berth 25", "10.5 m", "Handysize", "Pneumatic", "Loose cement"],
      ["PIBT", "Port Qasim", "13 m", "~65,000 t", "28,800 t/day", "Clinker, bulk cement"],
      ["MARGINAL WHARVES", "Port Qasim", "10.5 m", "60,000 t", "Ship’s gear", "Bagged cement, priority berth"],
      ["GWADAR", "Gwadar Port", "14 m+", "50,000 t", "Ship’s gear", "Breakbulk, transshipment"],
    ];
    rows.forEach((r, i) => {
      const y = 2.05 + i * 0.9;
      if (i % 2 === 0) card(s, M, y, 12.13, 0.84);
      let cx = M + 0.15;
      T(s, [{ text: r[0], options: { fontFace: HEAD, fontSize: 17, color: C.text2, breakLine: true } }, { text: r[1], options: { fontSize: 9.5, color: C.accent4 } }], { x: cx, y, w: 2.45, h: 0.84, valign: "middle" });
      cx = M + 2.6;
      [[r[2], 1.6], [r[3], 2.2], [r[4], 2.4]].forEach(([v, w]) => { T(s, v, { x: cx, y, w, h: 0.84, fontFace: HEAD, fontSize: 20, color: C.accent1, valign: "middle" }); cx += w; });
      T(s, r[5], { x: cx, y, w: 3.2, h: 0.84, fontSize: 11, color: C.text1, valign: "middle" });
    });
    source(s, "Source: KGTL & KGTML; PIBT; Port Qasim Authority; Gwadar Port Authority; Lucky Cement. Ship sizes in deadweight tonnes, indicative.");
    s.addNotes("For buyers planning vessels: KGTML now takes bulk carriers up to about 120,000 tonnes and loads 30,000 tonnes a day. Our own silos at Berth 25 load loose cement pneumatically. PIBT at Port Qasim loads about 28,800 tonnes a day on ships up to around 65,000 tonnes. The Marginal Wharves give cement priority berthing for bagged cargo, and Gwadar offers natural deep water for breakbulk and transshipment.");
  }

  // =================== SECTION 03: GWADAR ===================
  await divider("03", "GWADAR", "A natural deep-water port at the mouth of the Gulf.", ["Strategic location", "Natural deep draft", "Growth and scope"], "gw_aerial.jpg",
    { fx: 0.5, fy: 0.55, credit: "Photo: Bjoertvedt / Wikimedia Commons, CC BY-SA 4.0" });
  {
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: "GWADAR" });
    s.addImage({ data: await cropData("region.jpg", SW, SH), x: 0, y: 0, w: SW, h: SH, altText: "Satellite map from the Strait of Hormuz to Karachi" });
    const H = [1.22, 2.06], G = [7.3, 3.68], K = [11.94, 3.98];
    s.addShape(pres.shapes.LINE, { x: H[0], y: H[1], w: G[0] - H[0], h: G[1] - H[1], line: { color: HEX.red, width: 2, dashType: "dash" } });
    s.addShape(pres.shapes.LINE, { x: G[0], y: G[1], w: K[0] - G[0], h: K[1] - G[1], line: { color: "8A8A8A", width: 1.25, dashType: "dash" } });
    marker(s, H[0], H[1]); marker(s, G[0], G[1], true); marker(s, K[0], K[1]);
    pill(s, 0.55, 1.45, 2.3, "STRAIT OF HORMUZ", { size: 10 });
    pill(s, 6.55, 4.0, 1.5, "GWADAR", { size: 11, fill: C.accent1, color: "FFFFFF" });
    pill(s, 11.1, 4.3, 1.75, "KARACHI", { size: 10 });
    T(s, "630 km", { x: 3.55, y: 2.45, w: 1.6, h: 0.45, fontFace: HEAD, fontSize: 24, color: C.accent1, align: "center", rotate: 15 });
    T(s, "470 km", { x: 8.9, y: 3.42, w: 1.4, h: 0.4, fontFace: HEAD, fontSize: 18, color: C.accent4, align: "center" });
    T(s, "03  ·  GWADAR", { x: 5.0, y: 5.0, w: 8, h: 0.32, fontSize: 11, bold: true, color: C.accent1, charSpacing: 3 });
    T(s, "AT THE MOUTH OF THE GULF", { x: 5.0, y: 5.3, w: 7.8, h: 0.65, fontFace: HEAD, fontSize: 32, color: C.text2, valign: "top", objectName: "Title" });
    T(s, "About one-fifth of the world’s oil passes through Hormuz. Gwadar is the closest deep-water port outside it, on open ocean.", { x: 5.0, y: 6.0, w: 7.7, h: 0.65, fontSize: 13, color: C.text1, valign: "top" });
    T(s, "Imagery: Esri, Maxar, Earthstar Geographics. Distances great-circle.", { x: 8.5, y: 7.18, w: 4.65, h: 0.22, fontSize: 7, color: C.accent4, align: "right" });
    s.addNotes("Why Gwadar matters. It sits about 630 km from the Strait of Hormuz, through which roughly a fifth of the world's oil passes, and it is on the open Arabian Sea, outside the Gulf. When Hormuz was disrupted in 2026, traffic diverted to Gwadar. Karachi is a further 470 km east.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "GWADAR" });
    kt(s, "03  ·  GWADAR", "NATURALLY DEEP, READY TO GROW");
    await photo(s, "gw_port.jpg", 6.933, 0, 6.4, 7.5, { fx: 0.5, alt: "Gwadar Port with the hammerhead headland" });
    credit(s, "Photo: Wikimedia Commons, CC BY-SA 3.0", 8.3, 7.22, 4.9);
    [["14 → 24 m", "natural depth, today → future phases"], ["602 m", "3 multipurpose berths"], ["11,000 TEU", "Jan–Apr 2026, vs 8,300 in all of 2025"], ["200,000 t", "transit breakbulk in 3 months (Aug-26)"]].forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2); stat(s, M + col * 3.1, 1.7 + row * 1.75, 2.85, n, l, { size: 36, ls: 11, lh: 0.55 });
    });
    card(s, M, 5.35, 5.95, 1.3);
    T(s, [{ text: "SCOPE  ", options: { bold: true, color: C.accent1, charSpacing: 2, fontSize: 10 } }, { text: "50% of public-sector cargo now routed via Gwadar. A future transshipment hub for the Gulf & East Africa, and the sea gateway for CPEC and Balochistan’s mines.", options: { color: C.text1, fontSize: 12 } }], { x: M + 0.25, y: 5.4, w: 5.5, h: 1.2, valign: "middle" });
    source(s, "Source: Gwadar Port Authority; Asia Times (May-2026); PID; Pakistan Logistics & Shipping Summit (Aug-2026).", { x: M, y: 6.66, w: 5.95, align: "left" });
    s.addNotes("Gwadar is naturally deep, 14 metres today and planned to 24 in future phases, with three multipurpose berths on a 602-metre quay. Traffic is finally building: about 11,000 containers by April 2026 versus 8,300 in all of 2025, and some 200,000 tonnes of transit breakbulk in three months. Half of public-sector cargo is now routed through Gwadar.");
  }

  // =================== SECTION 04: CEMENT INDUSTRY ===================
  await divider("04", "CEMENT INDUSTRY", "Capacity is built. Exports are the outlet.", ["Capacity & despatches", "Who’s who", "This year’s news"], "n16.jpg", { fx: 0.5 });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "CEMENT INDUSTRY" });
    kt(s, "04  ·  CEMENT INDUSTRY", "87.5 MT OF CAPACITY, LOOKING FOR SHIPS");
    s.addChart([{ type: pres.charts.BAR, data: [
      { name: "Domestic", labels: ["FY22", "FY23", "FY24", "FY25", "FY26"], values: [47.6, 40.0, 38.2, 37.9, 41.5] },
      { name: "Exports", labels: ["FY22", "FY23", "FY24", "FY25", "FY26"], values: [5.3, 4.6, 7.1, 9.2, 9.0] },
    ], options: { barGrouping: "stacked" } }], {
      x: M, y: 1.6, w: 6.8, h: 4.9, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.maroon, "8A8A8A"], dataLabelFormatCode: "0.0",
      showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 11, dataLabelFontFace: BODY,
      showLegend: true, legendPos: "t", legendFontFace: BODY, legendFontSize: 11, legendColor: HEX.ink,
      showTitle: true, title: "Despatches, million tonnes", titleFontFace: BODY, titleFontSize: 11, titleColor: HEX.ink, titleBold: true,
      catAxisLabelColor: HEX.ink, catAxisLabelFontFace: BODY, catAxisLabelFontSize: 11, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLineShow: false,
    });
    [["87.5 Mt", "installed capacity"], ["50.5 Mt", "despatched in FY26, +7%"], ["~58%", "of capacity in use"], ["9.0 Mt", "exported; two-thirds clinker"], ["91%", "of exports from the south, via Karachi"], ["+19%", "southern exports, Jul–Sep 2026"]].forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2); stat(s, 7.85 + col * 2.5, 1.6 + row * 1.65, 2.35, n, l, { size: 38, ls: 11, lh: 0.5 });
    });
    source(s, "Source: Lucky Cement (capacity); APCMA; Lucky Cement Annual Report 2026; CemNet; Pakistan Observer (Oct-2026).");
    s.addNotes("87.5 million tonnes of capacity, 50.5 million despatched last year: only about 58% in use. That is why exports matter. 9 million tonnes went abroad, two-thirds as clinker, and 91% of it from southern mills through Karachi's ports. This year has started well: southern, sea-borne exports up 19% in July to September.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "CEMENT INDUSTRY" });
    kt(s, "04  ·  CEMENT INDUSTRY", "WHO’S WHO: 87.5 MT ACROSS THE INDUSTRY");
    const rows = [["Lucky Cement", 15.6, "logo.png"], ["Bestway", 15.3, "bestway.png"], ["Maple Leaf + Pioneer", 13.6, "mlcf.png"], ["DG Khan", 7.1, "dgk.png"], ["Kohat", 5.3, null], ["Cherat", 4.5, null], ["Others, incl. Fauji + Attock", 26.0, null]];
    const max = 26, bx = 4.6, bw = 6.6;
    for (let i = 0; i < rows.length; i++) {
      const [n, v, lg] = rows[i], y = 1.75 + i * 0.72;
      if (lg) await logo(s, lg, M, y + 0.04, 0.9, 0.5, "center");
      T(s, n, { x: M + 1.05, y, w: 2.9, h: 0.58, fontSize: 12.5, bold: i === 0, color: C.text1, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: bx, y: y + 0.11, w: bw * v / max, h: 0.36, fill: { color: i === 0 ? HEX.red : i === 6 ? "D9D6D6" : HEX.rose }, line: { type: "none" } });
      T(s, v.toFixed(1) + " Mt", { x: bx + bw * v / max + 0.12, y, w: 1.3, h: 0.58, fontFace: HEAD, fontSize: 16, color: i === 0 ? C.accent1 : C.text2, valign: "middle" });
    }
    source(s, "Source: Lucky Cement industry data (Intercem Dubai 2026). Maple Leaf–Pioneer merger effective Jul-2026.");
    s.addNotes("Who is who. Lucky is the largest producer at 15.6 million tonnes, then Bestway, and now Maple Leaf, which merged with Pioneer this year. Fauji and KAPCO took control of Attock Cement in February.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "CEMENT INDUSTRY" });
    kt(s, "04  ·  CEMENT INDUSTRY", "THIS YEAR’S HEADLINES");
    const news = [["fauji.png", "FEB 2026", "Fauji Cement & KAPCO take over Attock Cement"], ["mlcf.png", "JUL 2026", "Maple Leaf merges Pioneer: #3 producer"], [null, "SINCE OCT 2025", "Afghan border closed: northern exports down 54%"], [null, "JUL–SEP 2026", "Sea-borne exports from the south up 19%"]];
    const cw = 2.88, gap = (12.13 - 4 * cw) / 3;
    for (let i = 0; i < 4; i++) {
      const [lg, d, t] = news[i], x = M + i * (cw + gap);
      card(s, x, 1.65, cw, 4.9);
      if (lg) await logo(s, lg, x + 0.6, 1.95, cw - 1.2, 1.6);
      else circleIcon(s, x + cw / 2 - 0.6, 2.15, 1.2, i === 2 ? I.industry : I.ship);
      T(s, d, { x: x + 0.25, y: 3.85, w: cw - 0.5, h: 0.3, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 });
      T(s, t, { x: x + 0.25, y: 4.2, w: cw - 0.5, h: 1.6, fontSize: 15, bold: true, color: C.text1, valign: "top" });
    }
    source(s, "Source: Global Cement (27-Feb-2026); CemNet; APCMA via Pakistan Observer (Oct-2026).");
    s.addNotes("Four headlines. Consolidation: Fauji and KAPCO took Attock, Maple Leaf merged Pioneer. Disruption: the Afghan border has been closed since October 2025, cutting northern exports by more than half. And momentum: southern, sea-borne exports grew 19% in the first quarter of this fiscal year.");
  }

  // =================== SECTION 05: LUCKY CEMENT ===================
  await divider("05", "LUCKY CEMENT", "Pakistan’s largest cement producer and exporter.", ["At a glance & beyond borders", "Clean energy", "Next projects", "Lucky Ustaad: AI for builders"], "n03.jpg", { fx: 0.5 });
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "LUCKY CEMENT" });
    kt(s, "05  ·  LUCKY CEMENT", "ONE IN THREE TONNES PAKISTAN EXPORTS IS OURS");
    await photo(s, "n12.jpg", 9.5, 1.6, 3.23, 4.45, { fy: 0.3, alt: "Top Exporter award presented by the Prime Minister" });
    T(s, "Top Exporter trophy 2024-25, from the Prime Minister", { x: 9.5, y: 6.12, w: 3.23, h: 0.5, fontSize: 10, italic: true, color: C.accent4, valign: "top" });
    [["15.6 Mt", "capacity, largest in Pakistan"], ["9.6 Mt", "sold in FY26"], ["3.1 Mt", "exported in FY26"], ["34.4%", "of Pakistan’s exports"], ["30", "markets on 4 continents"], ["#1", "cement exporter"]].forEach(([n, l], i) => {
      const col = i % 3, row = Math.floor(i / 3); stat(s, M + col * 2.95, 1.75 + row * 1.85, 2.7, n, l, { size: 44, ls: 12, lh: 0.5 });
    });
    T(s, "BEYOND BORDERS", { x: M, y: 5.45, w: 4, h: 0.3, fontSize: 10, bold: true, color: C.accent1, charSpacing: 2 });
    const subs = [["nas.png", "Iraq, Samawah"], ["amc.png", "Iraq, Basra"], ["cimko.png", "DR Congo → 2.9 Mt"]];
    for (let i = 0; i < 3; i++) { await logo(s, subs[i][0], M + i * 2.95, 5.8, 1.0, 0.75); T(s, subs[i][1], { x: M + i * 2.95 + 1.1, y: 5.8, w: 1.75, h: 0.75, fontSize: 10.5, color: C.text1, valign: "middle" }); }
    source(s, "Source: Lucky Cement Annual Report 2026.");
    s.addNotes("Lucky Cement: 15.6 million tonnes of capacity, the largest in Pakistan. We sold 9.6 million tonnes last year and exported 3.1 million, a third of all Pakistani exports, to 30 markets on four continents. We also run plants in Iraq and the DR Congo, where capacity is being expanded to 2.9 million tonnes.");
  }
  {
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: "LUCKY CEMENT" });
    await photo(s, "n02.jpg", 0, 0, 6.6, 7.5, { fx: 0.3, alt: "Lucky Cement wind turbines and solar park, Karachi" });
    manualHead(s, 7.2, "05  ·  LUCKY CEMENT  ·  CLEAN ENERGY", "~55% OF OUR ENERGY IS CLEAN");
    [["56 MW", "waste heat recovery"], ["89.3 MW", "solar, incl. +15 MW new"], ["28.8 MW", "wind, largest in the cement industry"], ["22.7 MWh", "battery storage, largest in Pakistan"], ["274,614 t", "CO₂ avoided in FY26"], ["I-REC", "all projects certified"]].forEach(([n, l], i) => {
      const col = i % 2, row = Math.floor(i / 2); stat(s, 7.2 + col * 2.9, 1.75 + row * 1.6, 2.7, n, l, { size: 34, ls: 11, lh: 0.55 });
    });
    footLogo(s, 7.2);
    source(s, "Source: Lucky Cement decarbonisation fact sheet FY2026.", { x: 7.2, w: 5.53 });
    s.addNotes("About 55% of our energy now comes from clean sources: 56 MW of waste heat recovery, 89 MW of solar, a 28.8 MW wind farm, and 22.7 MWh of battery storage, the largest in Pakistan. Together they avoided about 275,000 tonnes of CO2 last year.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "LUCKY CEMENT" });
    kt(s, "05  ·  LUCKY CEMENT  ·  NEXT PROJECTS", "BUILDING A FUTURE-READY COMPANY");
    const cards = [["n04.jpg", "IFRS S1 & S2", "FY26", "First year of global sustainability & climate reporting"], ["n30.jpg", "LINE H EXPANSION", "300,000 t", "More cement-grinding capacity a year (MTPA)"], ["n17.jpg", "UTIS HYDROGEN", "4 of 4", "Karachi kilns burn with on-site hydrogen; +0.3 Mt capacity"]];
    const cw = 3.88;
    for (let i = 0; i < 3; i++) {
      const [img, h, n, t] = cards[i], x = M + i * (cw + 0.245), y = 1.6, ih = 2.3;
      card(s, x, y, cw, 5.1);
      await photo(s, img, x, y, cw, ih, { alt: h });
      const iy = y + ih + 0.25;
      T(s, h, { x: x + 0.3, y: iy, w: cw - 0.6, h: 0.32, fontSize: 11, bold: true, color: C.accent1, charSpacing: 2 });
      T(s, n, { x: x + 0.3, y: iy + 0.36, w: cw - 0.6, h: 0.85, fontFace: HEAD, fontSize: 44, color: C.text2, valign: "middle" });
      T(s, t, { x: x + 0.3, y: iy + 1.28, w: cw - 0.6, h: 0.75, fontSize: 12, color: C.text1, valign: "top" });
    }
    s.addShape(pres.shapes.RECTANGLE, { x: 11.55, y: 1.75, w: 0.95, h: 0.95, fill: { color: "FFFFFF" }, line: { type: "none" } });
    await logo(s, "utis.png", 11.6, 1.8, 0.85, 0.85);
    s.addNotes("Three projects that make us future-ready. IFRS S1 and S2: FY26 is our first year of global sustainability and climate reporting. Line H: 300,000 tonnes a year more cement-grinding capacity. And UTIS hydrogen-assisted combustion, now on all four Karachi kiln lines. [Presenter: confirm Line H timing with the project team.]");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "LUCKY CEMENT" });
    kt(s, "05  ·  LUCKY USTAAD", "AI FOR BUILDERS, IN THEIR POCKET");
    const px = 1.35, py = 1.55, pw = 2.75, ph = 5.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, rectRadius: 0.35, fill: { color: "151515" }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.25, blur: 12, offset: 4, angle: 90 }, objectName: "Phone body" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px + 0.12, y: py + 0.12, w: pw - 0.24, h: ph - 0.24, rectRadius: 0.26, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Phone screen: place screen recording here" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px + pw / 2 - 0.4, y: py + 0.22, w: 0.8, h: 0.16, rectRadius: 0.08, fill: { color: "151515" }, line: { type: "none" } });
    s.addImage({ data: I.robot, x: px + pw / 2 - 0.42, y: py + 1.6, w: 0.84, h: 0.84 });
    T(s, [{ text: "LUCKY USTAAD", options: { fontFace: HEAD, fontSize: 24, color: "FFFFFF", breakLine: true } }, { text: "Your AI construction companion", options: { fontSize: 10, color: "F0D9DC" } }], { x: px + 0.2, y: py + 2.6, w: pw - 0.4, h: 0.95, align: "center", valign: "top" });
    s.addImage({ data: I.play, x: px + pw / 2 - 0.15, y: py + 4.1, w: 0.3, h: 0.3 });
    T(s, "ESTIMATE  ·  PLAN  ·  BUILD", { x: 5.0, y: 1.7, w: 7.7, h: 0.75, fontFace: HEAD, fontSize: 36, color: C.accent1, valign: "middle" });
    [[I.comments, "Ask anything", "Construction & renovation answers, in plain words"], [I.calc, "Smart estimates", "Cement, materials and cost for the job"], [I.list, "Step-by-step guidance", "Through every stage of the build"], [I.store, "Nearest dealer", "Product advice and dealer access in one app"]].forEach(([ic, h, t], i) => {
      const y = 2.75 + i * 0.95; circleIcon(s, 5.0, y, 0.62, ic); rowText(s, 5.85, y - 0.1, 6.8, 0.82, h, t);
    });
    T(s, "From the quarry to the customer’s home: AI at every step.", { x: 5.0, y: 6.5, w: 7.7, h: 0.35, fontSize: 12, italic: true, bold: true, color: C.accent1 });
    s.addNotes("Lucky Ustaad is our AI construction assistant: a builder or homeowner asks a question in plain words, gets an estimate of cement and materials, step-by-step guidance and the nearest dealer. [Play the phone screen recording here.] TO ADD THE RECORDING: Insert > Video > This Device, size it to the dark-red phone screen (about 2.5 in x 5.0 in) and drag it on top.");
  }

  // =================== CLOSING ===================
  {
    const s = pres.addSlide({ masterName: "BLANK", sectionTitle: "LUCKY CEMENT" });
    await photo(s, "n48.jpg", 6.663, 0, 6.67, 7.5, { fx: 0.45, alt: "Lucky Cement Pezu plant and solar park" });
    s.addImage({ path: S + "logo.png", x: M, y: 0.6, w: 1.6, h: 0.74, altText: "Lucky Cement logo" });
    T(s, "PORT-READY.\nEXPORT-READY.\nFUTURE-READY.", { x: M, y: 2.1, w: 5.6, h: 1.75, fontSize: 20, bold: true, color: C.accent1, charSpacing: 2, valign: "top" });
    T(s, "THANK YOU", { x: M, y: 3.9, w: 5.8, h: 1.3, fontFace: HEAD, fontSize: 80, color: C.text2, valign: "top", objectName: "Title" });
    T(s, [{ text: "NAUMAN AHMED", options: { bold: true, color: C.text2, fontSize: 13, charSpacing: 2, breakLine: true } }, { text: "Head of Exports, Lucky Cement Limited  ·  lucky-cement.com", options: { color: C.accent4, fontSize: 11 } }], { x: M, y: 6.15, w: 5.8, h: 0.65, valign: "top" });
    s.addNotes("Thank you. Pakistan's ports are getting deeper and faster, the industry has the capacity, and Lucky is ready for the next ship. Happy to take questions.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();

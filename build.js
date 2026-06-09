/* Paytm Velocity API — B2B Campaign Brief deck */
const PptxGenJS = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa6 = require("react-icons/fa6");
const fa = require("react-icons/fa");

// ---------- palette ----------
const NAVY  = "002970";   // Paytm navy
const NAVY2 = "001A4D";   // deeper navy (dark bgs)
const CYAN  = "00BAF2";   // Paytm cyan
const CYAND = "008FC4";   // darker cyan
const CLOUD = "EAF6FE";   // light wash
const CLOUD2= "F5FBFE";   // lighter wash
const WHITE = "FFFFFF";
const GREEN = "1FB271";   // success
const CORAL = "FF6B5E";   // pain accent
const INK   = "16263F";   // dark text
const SLATE = "3A4A63";   // body text
const MUTED = "6E7E96";   // muted
const LINE  = "D6E6F4";   // hairline

const HEAD = "Montserrat";
const BODY = "Calibri";

const W = 13.33, H = 7.5, M = 0.62;

const shadow = () => ({ type: "outer", color: "0A2540", blur: 9, offset: 3, angle: 90, opacity: 0.16 });
const softShadow = () => ({ type: "outer", color: "0A2540", blur: 7, offset: 2, angle: 90, opacity: 0.10 });

// ---------- icons ----------
const icCache = {};
function pick(name){ return fa6[name] || fa[name] || fa6.FaCircle; }
async function icon(name, colorHex, size = 256){
  const key = name + colorHex;
  if (icCache[key]) return icCache[key];
  const Comp = pick(name);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: colorHex, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  const data = "image/png;base64," + png.toString("base64");
  icCache[key] = data;
  return data;
}

(async () => {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "B2B Media Planning";
  pres.title = "Paytm Velocity API — Campaign Brief";

  // ---------- shared helpers ----------
  function footer(slide, n){
    slide.addText([
      { text: "Paytm Velocity API  ·  Mock B2B Campaign Brief", options: { color: MUTED } },
      { text: "      (fictional product — illustrative figures)", options: { color: MUTED, italic: true } }
    ], { x: M, y: H-0.42, w: 9, h: 0.3, fontFace: BODY, fontSize: 8.5, align: "left", margin: 0 });
    slide.addText(String(n).padStart(2,"0"), { x: W-1.0, y: H-0.42, w: 0.5, h: 0.3, fontFace: HEAD, fontSize: 9, color: MUTED, align: "right", margin: 0, bold: true });
  }

  async function sectionHeader(slide, num, title, iconName){
    // number chip
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 0.5, w: 0.82, h: 0.82, rectRadius: 0.12, fill: { color: NAVY }, shadow: softShadow() });
    slide.addText(num, { x: M, y: 0.5, w: 0.82, h: 0.82, fontFace: HEAD, fontSize: 22, bold: true, color: CYAN, align: "center", valign: "middle", margin: 0 });
    if (iconName){
      slide.addImage({ data: await icon(iconName, "#"+CYAND), x: W-M-0.62, y: 0.6, w: 0.62, h: 0.62 });
    }
    slide.addText(title, { x: M+1.02, y: 0.5, w: 9.8, h: 0.82, fontFace: HEAD, fontSize: 27, bold: true, color: NAVY, align: "left", valign: "middle", margin: 0 });
  }

  async function statCard(slide, x, y, w, h, value, label, accent, iconName){
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: softShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.1, fill: { color: accent } });
    if (iconName){
      slide.addImage({ data: await icon(iconName, "#"+accent), x: x+w-0.62, y: y+0.32, w: 0.42, h: 0.42 });
    }
    slide.addText(value, { x: x+0.24, y: y+0.26, w: w-0.55, h: 0.74, fontFace: HEAD, fontSize: 30, bold: true, color: NAVY, align: "left", valign: "top", margin: 0 });
    slide.addText(label, { x: x+0.24, y: y+1.04, w: w-0.42, h: h-1.14, fontFace: BODY, fontSize: 11.5, color: SLATE, align: "left", valign: "top", margin: 0 });
  }

  function pill(slide, x, y, w, h, text, fillColor, textColor, fs){
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h/2, fill: { color: fillColor } });
    slide.addText(text, { x, y, w, h, fontFace: HEAD, fontSize: fs||11, bold: true, color: textColor, align: "center", valign: "middle", margin: 0 });
  }

  // brand mini-logo lockup
  function logoLockup(slide, x, y, scale, onDark){
    const s = scale || 1;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.34*s, h: 0.34*s, rectRadius: 0.07*s, fill: { color: CYAN } });
    slide.addText("p", { x, y: y-0.02*s, w: 0.34*s, h: 0.34*s, fontFace: HEAD, fontSize: 16*s, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
    slide.addText("Paytm", { x: x+0.4*s, y: y-0.06*s, w: 1.6*s, h: 0.46*s, fontFace: HEAD, fontSize: 17*s, bold: true, color: onDark ? WHITE : NAVY, align: "left", valign: "middle", margin: 0 });
  }

  // =====================================================================
  // SLIDE 1 — TITLE
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY2 };
    // decorative shapes
    s.addShape(pres.shapes.OVAL, { x: 9.7, y: -2.2, w: 6.2, h: 6.2, fill: { color: NAVY }, line: { color: NAVY } });
    s.addShape(pres.shapes.OVAL, { x: 11.3, y: 3.6, w: 4.6, h: 4.6, fill: { color: CYAND, transparency: 72 }, line: { type: "none" } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: -0.8, y: 5.7, w: 5.2, h: 3.4, rectRadius: 0.4, fill: { color: NAVY }, line: { type:"none" } });
    s.addImage({ data: await icon("FaBolt", "#"+CYAN, 400), x: 10.55, y: 1.15, w: 2.7, h: 2.7, transparency: 12 });

    logoLockup(s, M, 0.62, 1.15, true);
    pill(s, W-3.55, 0.66, 2.95, 0.5, "MOCK CAMPAIGN · FICTIONAL", "12345A", CYAN, 10);

    s.addText("PAYTM  ·  GO-TO-MARKET CAMPAIGN BRIEF", { x: M, y: 2.35, w: 9, h: 0.4, fontFace: HEAD, fontSize: 13, bold: true, color: CYAN, charSpacing: 3, margin: 0 });
    s.addText([
      { text: "Velocity", options: { color: CYAN } },
      { text: " API", options: { color: WHITE } }
    ], { x: M-0.03, y: 2.78, w: 12, h: 1.5, fontFace: HEAD, fontSize: 70, bold: true, align: "left", margin: 0 });
    s.addText("UPI, at the speed of intent.", { x: M, y: 4.28, w: 11, h: 0.6, fontFace: HEAD, fontSize: 22, italic: true, color: CLOUD, margin: 0 });
    s.addText("An ultra-fast, drop-in UPI payments API for e-commerce platforms and online businesses.",
      { x: M, y: 4.92, w: 9.4, h: 0.5, fontFace: BODY, fontSize: 14, color: "AFC4E6", margin: 0 });

    // meta chips
    const chips = ["30-Day Flight", "₹30,00,000 Budget", "Meta · LinkedIn · Google Search"];
    let cx = M;
    for (const c of chips){
      const cw = 0.42 + c.length*0.115;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: 6.0, w: cw, h: 0.56, rectRadius: 0.28, fill: { color: "FFFFFF", transparency: 88 }, line: { color: CYAN, width: 1 } });
      s.addText(c, { x: cx, y: 6.0, w: cw, h: 0.56, fontFace: HEAD, fontSize: 11.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      cx += cw + 0.3;
    }
    s.addText("Recommended product name: Paytm Velocity API   (client working name: “Paytm Swift-UPI API”)",
      { x: M, y: 6.95, w: 12, h: 0.35, fontFace: BODY, fontSize: 10.5, italic: true, color: "8FA6CC", margin: 0 });
  }

  // =====================================================================
  // SLIDE 2 — CAMPAIGN SNAPSHOT
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "00", "Campaign at a Glance", "FaBullseye");

    // left: product + objective panel
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 1.7, w: 5.5, h: 5.05, fill: { color: NAVY }, shadow: shadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 1.7, w: 0.12, h: 5.05, fill: { color: CYAN } });
    s.addText("THE PRODUCT", { x: M+0.4, y: 1.98, w: 4.8, h: 0.35, fontFace: HEAD, fontSize: 12, bold: true, color: CYAN, charSpacing: 2, margin: 0 });
    s.addText("Paytm Velocity API", { x: M+0.4, y: 2.32, w: 4.8, h: 0.5, fontFace: HEAD, fontSize: 23, bold: true, color: WHITE, margin: 0 });
    s.addText("A developer-first UPI payments API that drops into any checkout in hours. One-tap collect, 0.8s confirmation, signed webhooks, 99.99% uptime, T+1 settlement.",
      { x: M+0.4, y: 2.86, w: 4.75, h: 1.2, fontFace: BODY, fontSize: 13, color: "D5E2F5", lineSpacingMultiple: 1.05, margin: 0 });

    s.addText("PRIMARY OBJECTIVE", { x: M+0.4, y: 4.18, w: 4.8, h: 0.35, fontFace: HEAD, fontSize: 12, bold: true, color: CYAN, charSpacing: 2, margin: 0 });
    s.addText("Drive qualified developer & merchant sign-ups to the sandbox, and a pipeline of integrated, live-processing merchants.",
      { x: M+0.4, y: 4.5, w: 4.75, h: 1.0, fontFace: BODY, fontSize: 13, color: "D5E2F5", lineSpacingMultiple: 1.05, margin: 0 });

    s.addText([
      { text: "Funnel goal:  ", options: { bold: true, color: WHITE } },
      { text: "Reach → Click → Sandbox sign-up → Integration → Go-live", options: { color: CYAN } }
    ], { x: M+0.4, y: 5.7, w: 4.8, h: 0.8, fontFace: BODY, fontSize: 12.5, margin: 0, valign: "top" });

    // right: 4 stat cards (2x2) + audience strip
    const gx = 6.55, gw = 3.0, gy = 1.7, gh = 1.72, gap = 0.26;
    await statCard(s, gx,        gy,        gw, gh, "₹30L", "Monthly media budget (aggressive)", CYAND, "FaSackDollar");
    await statCard(s, gx+gw+gap, gy,        gw, gh, "3", "Channels: Meta · LinkedIn · Google", CYAN, "FaLayerGroup");
    await statCard(s, gx,        gy+gh+gap, gw, gh, "30", "Day flight, full-funnel", NAVY, "FaCalendarDays");
    await statCard(s, gx+gw+gap, gy+gh+gap, gw, gh, "3.9M", "Projected reach (impressions)", GREEN, "FaChartLine");

    // audience strip
    const stripY = gy+2*gh+2*gap+0.08;
    s.addShape(pres.shapes.RECTANGLE, { x: gx, y: stripY, w: gw*2+gap, h: 1.08, fill: { color: CLOUD }, line: { color: LINE, width: 1 } });
    s.addImage({ data: await icon("FaUsers", "#"+NAVY), x: gx+0.28, y: stripY+0.28, w: 0.54, h: 0.54 });
    s.addText("Target: online business owners, e-commerce founders, freelance developers & tech leads who choose the payment stack — across 3 sharp segments (next).",
      { x: gx+1.05, y: stripY, w: gw*2+gap-1.3, h: 1.08, fontFace: BODY, fontSize: 12, color: SLATE, valign: "middle", margin: 0 });

    footer(s, 2);
  }

  // =====================================================================
  // SLIDE 3 — AUDIENCE: 3 SEGMENTS
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "01", "Audience Research — Three B2B Segments", "FaUsers");
    s.addText("The 18–60 band splits into three decision-maker personas with very different triggers, vocabularies and fears.",
      { x: M, y: 1.4, w: 12, h: 0.4, fontFace: BODY, fontSize: 13, italic: true, color: MUTED, margin: 0 });

    const segs = [
      { ic: "FaBolt", clr: CYAND, name: "The Conversion-Obsessed Founder", age: "Age 22–35 · D2C & app-first startups",
        who: "Gen-Z / millennial founders of D2C brands, marketplaces and subscription apps. Bootstrapped to Series-A, growth is everything.",
        pains: ["Redirect-heavy checkout abandons ~1 in 5 carts","High, opaque MDR eats thin D2C margins","Low payment success rates at peak traffic"],
        buy: "Buys on: checkout conversion + speed to integrate" },
      { ic: "FaCode", clr: NAVY, name: "The Pragmatic Tech Lead", age: "Age 28–45 · CTOs, tech leads, freelance devs",
        who: "Engineers who own or build the payment layer for their platform or for clients — growth-stage SaaS, agencies, marketplaces.",
        pains: ["Thin docs & flaky sandbox slow integration","Webhooks that miss / aren’t idempotent","Latency & downtime spikes; slow dev support"],
        buy: "Buys on: docs quality, uptime/SLA, latency" },
      { ic: "FaStore", clr: GREEN, name: "The Digitizing Retail Owner", age: "Age 40–60 · Traditional SMB going online",
        who: "Gen-X offline retailers, distributors and service businesses launching a website or online store for the first time.",
        pains: ["Intimidated by anything that looks ‘technical’","Fear of failed txns & messy reconciliation","T+2/T+3 settlement strains cash flow"],
        buy: "Buys on: ease, trust, T+1 settlement, support" },
    ];
    const cw = (W - 2*M - 2*0.4)/3, cy = 1.95, ch = 4.85;
    for (let i=0;i<3;i++){
      const x = M + i*(cw+0.4);
      s.addShape(pres.shapes.RECTANGLE, { x, y: cy, w: cw, h: ch, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: softShadow() });
      // header band
      s.addShape(pres.shapes.RECTANGLE, { x, y: cy, w: cw, h: 1.18, fill: { color: segs[i].clr } });
      s.addShape(pres.shapes.OVAL, { x: x+0.3, y: cy+0.27, w: 0.64, h: 0.64, fill: { color: "FFFFFF", transparency: 84 }, line: { type:"none" } });
      s.addImage({ data: await icon(segs[i].ic, "#"+WHITE), x: x+0.42, y: cy+0.39, w: 0.4, h: 0.4 });
      s.addText(segs[i].name, { x: x+1.06, y: cy+0.18, w: cw-1.2, h: 0.85, fontFace: HEAD, fontSize: 14.5, bold: true, color: WHITE, valign: "middle", margin: 0 });
      s.addText(segs[i].age, { x: x+0.3, y: cy+1.28, w: cw-0.5, h: 0.35, fontFace: HEAD, fontSize: 10.5, bold: true, color: segs[i].clr, margin: 0 });
      s.addText(segs[i].who, { x: x+0.3, y: cy+1.62, w: cw-0.55, h: 1.15, fontFace: BODY, fontSize: 11.5, color: SLATE, lineSpacingMultiple: 1.03, margin: 0 });
      s.addText("PAIN POINTS WITH TODAY'S GATEWAYS", { x: x+0.3, y: cy+2.78, w: cw-0.5, h: 0.3, fontFace: HEAD, fontSize: 9.5, bold: true, color: CORAL, charSpacing: 1, margin: 0 });
      s.addText(segs[i].pains.map((p,j)=>({ text: p, options: { bullet: { code: "2022", indent: 12 }, color: SLATE, breakLine: true, paraSpaceAfter: 4 } })),
        { x: x+0.32, y: cy+3.12, w: cw-0.58, h: 1.25, fontFace: BODY, fontSize: 11, margin: 0, valign: "top" });
      // buy-on footer
      s.addShape(pres.shapes.RECTANGLE, { x: x+0.0, y: cy+ch-0.62, w: cw, h: 0.62, fill: { color: CLOUD } });
      s.addText(segs[i].buy, { x: x+0.3, y: cy+ch-0.62, w: cw-0.5, h: 0.62, fontFace: HEAD, fontSize: 10.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    }
    footer(s, 3);
  }

  // =====================================================================
  // SLIDE 4 — PAIN → FIX MAP
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "01", "The Shared Enemy — and How Velocity Wins", "FaTriangleExclamation");

    const rows = [
      [ {text:"Segment",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"left"}},
        {text:"Sharpest pain with current online payment solutions",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"left"}},
        {text:"What Paytm Velocity fixes",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"left"}} ],
      [ {text:"Conversion-Obsessed Founder",options:{bold:true,color:NAVY}},
        "Redirect-heavy checkout kills ~20% of carts; flat MDR eats thin D2C margins; success rates dip at peak.",
        {text:"In-context one-tap UPI · 0.8s confirmation · higher auth success — checkout that converts.",options:{color:"0E5C3F"}} ],
      [ {text:"Pragmatic Tech Lead",options:{bold:true,color:NAVY}},
        "Thin docs, a flaky sandbox and webhooks that miss turn a ‘quick integration’ into a multi-sprint debug.",
        {text:"Clean docs + true sandbox · signed, idempotent webhooks with retries · 99.99% uptime, <120ms.",options:{color:"0E5C3F"}} ],
      [ {text:"Digitizing Retail Owner",options:{bold:true,color:NAVY}},
        "Too technical to set up; failed transactions and messy reconciliation; T+2/T+3 settlement squeezes cash flow.",
        {text:"1-click Shopify/Woo plugin · auto-reconciliation · T+1 settlement · regional-language support.",options:{color:"0E5C3F"}} ],
    ];
    s.addTable(rows, {
      x: M, y: 1.75, w: W-2*M, colW: [2.7, 5.0, 4.39],
      rowH: [0.45, 1.18, 1.18, 1.18],
      border: { type: "solid", color: LINE, pt: 1 },
      fill: { color: WHITE },
      fontFace: BODY, fontSize: 12, color: SLATE, valign: "middle",
      align: "left", margin: [4, 8, 4, 8],
    });
    // alt row shading via overlay is complex; keep clean white.

    // universal pains strip
    s.addText("FOUR UNIVERSAL PAINS THIS CAMPAIGN ATTACKS", { x: M, y: 5.98, w: 12, h: 0.32, fontFace: HEAD, fontSize: 11, bold: true, color: CORAL, charSpacing: 1.5, margin: 0 });
    const ups = [["FaArrowDownLong","Low success rates"],["FaStopwatch","Slow checkout"],["FaPercent","Opaque fees"],["FaCalendarDays","Delayed settlement"]];
    const uw = (W-2*M-3*0.3)/4;
    for (let i=0;i<4;i++){
      const x = M + i*(uw+0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 6.32, w: uw, h: 0.6, rectRadius: 0.1, fill: { color: WHITE }, line: { color: CORAL, width: 1 } });
      s.addImage({ data: await icon(ups[i][0], "#"+CORAL), x: x+0.18, y: 6.47, w: 0.3, h: 0.3 });
      s.addText(ups[i][1], { x: x+0.55, y: 6.32, w: uw-0.65, h: 0.6, fontFace: HEAD, fontSize: 11.5, bold: true, color: NAVY, valign: "middle", margin: 0 });
    }
    footer(s, 4);
  }

  // =====================================================================
  // SLIDE 5 — BRAND / CREATIVE SYSTEM
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "02", "Ad Creatives — The Brand Kit for Canva/Figma", "FaPalette");

    // palette swatches
    s.addText("COLOR PALETTE", { x: M, y: 1.55, w: 5, h: 0.32, fontFace: HEAD, fontSize: 12, bold: true, color: NAVY, charSpacing: 1.5, margin: 0 });
    const sw = [["Paytm Navy","002970",WHITE],["Paytm Cyan","00BAF2",NAVY],["Cloud Wash","EAF6FE",NAVY],["Success Green","1FB271",WHITE],["Ink / Body","16263F",WHITE]];
    const swW = 1.28, swH = 1.42;
    for (let i=0;i<sw.length;i++){
      const x = M + i*(swW+0.18);
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.95, w: swW, h: swH, fill: { color: sw[i][1] }, line: { color: LINE, width: 1 }, shadow: softShadow() });
      s.addText(sw[i][0], { x, y: 2.55, w: swW, h: 0.4, fontFace: HEAD, fontSize: 10, bold: true, color: sw[i][2], align: "center", margin: 0 });
      s.addText("#"+sw[i][1], { x, y: 2.92, w: swW, h: 0.35, fontFace: BODY, fontSize: 9.5, color: sw[i][2], align: "center", margin: 0 });
    }
    s.addText("Use Navy as the dominant ground (~60%), Cyan as the single sharp accent, Green only for success/confirmation moments.",
      { x: M, y: 3.5, w: 6.95, h: 0.6, fontFace: BODY, fontSize: 11.5, italic: true, color: MUTED, margin: 0 });

    // typography
    s.addText("TYPOGRAPHY", { x: M, y: 4.2, w: 5, h: 0.32, fontFace: HEAD, fontSize: 12, bold: true, color: NAVY, charSpacing: 1.5, margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 4.55, w: 6.95, h: 2.05, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: softShadow() });
    s.addText([
      { text: "Montserrat", options: { fontFace: HEAD, fontSize: 22, bold: true, color: NAVY, breakLine: true } },
      { text: "Headlines & numbers — bold, geometric, confident", options: { fontFace: BODY, fontSize: 11, color: MUTED, breakLine: true } },
    ], { x: M+0.3, y: 4.72, w: 6.4, h: 0.85, margin: 0, valign: "top" });
    s.addText([
      { text: "Inter / Calibri", options: { fontFace: BODY, fontSize: 16, bold: true, color: INK, breakLine: true } },
      { text: "Body & captions — clean, highly legible at small sizes", options: { fontFace: BODY, fontSize: 11, color: MUTED } },
    ], { x: M+0.3, y: 5.62, w: 6.4, h: 0.85, margin: 0, valign: "top" });

    // right: motif + do/don't
    const rx = 7.9, rw = W-M-rx;
    s.addShape(pres.shapes.RECTANGLE, { x: rx, y: 1.55, w: rw, h: 5.05, fill: { color: NAVY }, shadow: shadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: rx, y: 1.55, w: rw, h: 0.12, fill: { color: CYAN } });
    s.addText("VISUAL MOTIF & RULES", { x: rx+0.35, y: 1.85, w: rw-0.7, h: 0.35, fontFace: HEAD, fontSize: 12, bold: true, color: CYAN, charSpacing: 1.5, margin: 0 });
    const rules = [
      ["FaBolt","Speed motif — bolt / motion lines / a stopwatch reading 0.8s on every concept."],
      ["FaCircleCheck","Green success tick is the hero moment — show the ‘paid’ state, not the form."],
      ["FaMobileScreenButton","Real UI: show a phone or checkout with a one-tap UPI screen, never abstract clip-art."],
      ["FaRegSquare","Rounded corners (12–16px), generous whitespace, logo top-left, CTA pill bottom."],
    ];
    let ry = 2.35;
    for (const r of rules){
      s.addShape(pres.shapes.OVAL, { x: rx+0.35, y: ry, w: 0.5, h: 0.5, fill: { color: CYAND }, line:{type:"none"} });
      s.addImage({ data: await icon(r[0], "#"+WHITE), x: rx+0.46, y: ry+0.11, w: 0.28, h: 0.28 });
      s.addText(r[1], { x: rx+1.0, y: ry-0.06, w: rw-1.3, h: 0.7, fontFace: BODY, fontSize: 11.5, color: "DCE7F8", valign: "middle", margin: 0 });
      ry += 0.78;
    }
    s.addShape(pres.shapes.RECTANGLE, { x: rx, y: 5.78, w: rw, h: 0.82, fill: { color: NAVY2 } });
    s.addText([
      { text: "DON'T:  ", options: { bold: true, color: CORAL } },
      { text: "gradients muddier than navy→cyan, stocky ‘handshake’ photos, more than one accent colour, tiny body text.", options: { color: "C7D4EA" } }
    ], { x: rx+0.35, y: 5.78, w: rw-0.7, h: 0.82, fontFace: BODY, fontSize: 10.5, valign: "middle", margin: 0 });

    footer(s, 5);
  }

  // =====================================================================
  // CREATIVE DETAIL SLIDES (6,7,8)
  // =====================================================================
  async function creativeSlide(n, cfg, drawMock){
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "02", cfg.title, cfg.icon);
    // chips under header
    pill(s, M+1.02, 1.42, 2.1, 0.42, cfg.size, NAVY, WHITE, 10);
    pill(s, M+1.02+2.25, 1.42, cfg.segW, 0.42, cfg.seg, CLOUD, NAVY, 10);

    // left mock canvas
    const mx = M, my = 2.05, mh = 4.7;
    await drawMock(s, mx, my, mh);

    // right specs
    const sx = 6.55, sw2 = W-M-sx;
    const blocks = [
      ["LAYOUT", cfg.layout],
      ["PALETTE", cfg.palette],
      ["IMAGERY / GRAPHICS", cfg.imagery],
      ["TEXT PLACEMENT", cfg.textp],
    ];
    let by = 2.05;
    for (const b of blocks){
      const bh = b[0]==="LAYOUT" || b[0]==="IMAGERY / GRAPHICS" ? 1.28 : 1.06;
      s.addShape(pres.shapes.RECTANGLE, { x: sx, y: by, w: sw2, h: bh, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: softShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x: sx, y: by, w: 0.1, h: bh, fill: { color: CYAN } });
      s.addText(b[0], { x: sx+0.28, y: by+0.13, w: sw2-0.5, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: CYAND, charSpacing: 1.2, margin: 0 });
      s.addText(b[1], { x: sx+0.28, y: by+0.42, w: sw2-0.5, h: bh-0.5, fontFace: BODY, fontSize: 11.5, color: SLATE, lineSpacingMultiple: 1.02, valign: "top", margin: 0 });
      by += bh + 0.16;
    }
    footer(s, n);
  }

  // ---- Creative 1 mock: speed/conversion (square) ----
  await creativeSlide(6, {
    title: "Creative 1 — “Speed = Conversion”", icon: "FaBolt",
    size: "1080×1080 (Meta) + 1200×627 (LI)", seg: "Founder", segW: 1.5,
    layout: "Square. A diagonal split: bottom-left navy wedge (~40%), the rest a smooth cyan→cloud wash. A phone mockup sits centre-right showing a UPI ‘Payment successful’ screen. A large ‘0.8s’ stat with a stopwatch icon anchors the top-left.",
    palette: "Navy 002970 wedge · Cyan 00BAF2 wash · White copy · Green 1FB271 success tick.",
    imagery: "Smartphone with a one-tap UPI success UI (₹1,499 · paid), a stopwatch/bolt icon, faint speed/motion lines behind the phone, green check badge.",
    textp: "Top-left: ‘0.8s’ stat. Headline mid-left over the wedge. Primary line beneath it. CTA pill bottom-centre. Paytm logo top-left corner.",
  }, async (s, x, y, mh) => {
    const mw = 4.4, wedgeH = 1.55;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: mw, h: mh, fill: { color: CLOUD }, line: { color: LINE, width: 1 }, shadow: shadow() });
    // cyan wash block
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: mw, h: mh-wedgeH, fill: { color: CYAN, transparency: 74 } });
    // navy wedge bottom band
    s.addShape(pres.shapes.RECTANGLE, { x, y: y+mh-wedgeH, w: mw, h: wedgeH, fill: { color: NAVY } });
    logoLockup(s, x+0.25, y+0.24, 0.85, false);
    // stat (left)
    s.addText("0.8s", { x: x+0.25, y: y+0.95, w: 2.25, h: 0.85, fontFace: HEAD, fontSize: 44, bold: true, color: NAVY, margin: 0 });
    s.addText("checkout", { x: x+0.28, y: y+1.74, w: 2.2, h: 0.3, fontFace: HEAD, fontSize: 13, bold: true, color: CYAND, margin: 0 });
    // phone (centre-right, fully inside the cyan zone)
    const pw = 1.45, ph = 2.05, px = x+mw-pw-0.34, py = y+0.78;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, rectRadius: 0.16, fill: { color: WHITE }, line: { color: NAVY, width: 2 }, shadow: softShadow() });
    s.addShape(pres.shapes.OVAL, { x: px+pw/2-0.21, y: py+0.42, w: 0.42, h: 0.42, fill: { color: GREEN } });
    s.addImage({ data: await icon("FaCheck", "#"+WHITE), x: px+pw/2-0.12, y: py+0.51, w: 0.24, h: 0.24 });
    s.addText("Paid", { x: px, y: py+1.0, w: pw, h: 0.32, fontFace: HEAD, fontSize: 14, bold: true, color: NAVY, align: "center", margin: 0 });
    s.addText("₹1,499", { x: px, y: py+1.36, w: pw, h: 0.3, fontFace: BODY, fontSize: 11, color: MUTED, align: "center", margin: 0 });
    // headline + cta in the wedge
    s.addText("Checkout in 0.8 seconds.", { x: x+0.25, y: y+mh-wedgeH+0.16, w: mw-0.5, h: 0.5, fontFace: HEAD, fontSize: 18, bold: true, color: WHITE, margin: 0 });
    pill(s, x+0.25, y+mh-0.64, 2.0, 0.5, "Get API Keys ›", CYAN, NAVY, 12);
  });

  // ---- Creative 2 mock: developer (landscape) ----
  await creativeSlide(7, {
    title: "Creative 2 — “Ship UPI this weekend”", icon: "FaCode",
    size: "1200×627 (LinkedIn / Google)", seg: "Tech Lead", segW: 1.7,
    layout: "Landscape, full dark ‘IDE’ canvas. Left ~60%: a rounded code-editor card showing the ~5 lines needed to integrate. Right ~40%: a terminal chip reading ‘200 OK · payment.success’ plus three proof chips (99.99% uptime · <120ms · signed webhooks).",
    palette: "Navy2 001A4D / near-black ground · Cyan syntax highlights · Green status dot · White copy.",
    imagery: "Stylised code editor window with traffic-light dots, monospace code with cyan/green syntax, a terminal/status line, bracket motif in the corner.",
    textp: "Headline top-left (white). Code card centre-left. Proof chips stacked right. CTA pill bottom-right. ‘Paytm for Developers’ lockup top-left.",
  }, async (s, x, y, mh) => {
    const mw = 4.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: mw, h: mh, fill: { color: NAVY2 }, line: { color: NAVY, width: 1 }, shadow: shadow() });
    logoLockup(s, x+0.25, y+0.22, 0.78, true);
    s.addText("Ship UPI this weekend.", { x: x+0.25, y: y+0.74, w: mw-0.5, h: 0.5, fontFace: HEAD, fontSize: 17, bold: true, color: WHITE, margin: 0 });
    // code card
    const cx = x+0.25, cy = y+1.35, cwd = 2.5, chd = 2.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: cwd, h: chd, rectRadius: 0.1, fill: { color: "0B1730" }, line: { color: CYAND, width: 1 } });
    for (let i=0;i<3;i++) s.addShape(pres.shapes.OVAL, { x: cx+0.18+i*0.22, y: cy+0.16, w: 0.12, h: 0.12, fill: { color: i===0?"FF6B5E":i===1?"FFD166":GREEN }, line:{type:"none"} });
    s.addText([
      { text: "import ", options: { color: CYAN } }, { text: "Velocity\n", options: { color: WHITE, breakLine: true } },
      { text: "pay = Velocity(key)\n", options: { color: "9FE8FF", breakLine: true } },
      { text: "pay.collect(₹1499)\n", options: { color: "9FE8FF", breakLine: true } },
      { text: "# → 200 success", options: { color: GREEN } },
    ], { x: cx+0.2, y: cy+0.4, w: cwd-0.35, h: chd-0.5, fontFace: "Consolas", fontSize: 11, valign: "top", margin: 0, lineSpacingMultiple: 1.15 });
    // proof chips right
    const chips = ["200 OK · payment.success","99.99% uptime","<120ms latency","Signed webhooks + retries"];
    let yy = cy;
    for (let i=0;i<chips.length;i++){
      const isStatus = i===0;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+mw-1.55, y: yy, w: 1.4, h: 0.44, rectRadius: 0.08, fill: { color: isStatus?GREEN:"12345A" }, line: { color: isStatus?GREEN:CYAND, width: 1 } });
      s.addText(chips[i], { x: x+mw-1.55, y: yy, w: 1.4, h: 0.44, fontFace: BODY, fontSize: 8.5, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      yy += 0.56;
    }
    pill(s, x+mw-1.95, y+mh-0.7, 1.8, 0.5, "Read the Docs ›", CYAN, NAVY, 11.5);
  });

  // ---- Creative 3 mock: SMB trust (portrait) ----
  await creativeSlide(8, {
    title: "Creative 3 — “Get paid online, the easy way”", icon: "FaStore",
    size: "1080×1350 (Meta / Instagram)", seg: "Retail Owner", segW: 1.9,
    layout: "Portrait. Top ~58%: a bright, authentic photo of a mid-40s Indian business owner smiling at a laptop that shows the Paytm merchant dashboard. Bottom ~42%: a solid navy band carrying the headline, a row of trust badges, and the CTA.",
    palette: "Natural photo up top · Navy 002970 band below · Cyan accent word + CTA · White copy · Green badge ticks.",
    imagery: "Real Indian SMB owner at a laptop/phone, Paytm ‘Payments’ dashboard on screen, trust badges (RBI-regulated · PCI-DSS · 2.5cr+ merchants), a small ‘T+1 settlement’ icon.",
    textp: "Headline on the navy band (cyan accent on ‘easy way’). Trust-badge row beneath. CTA pill bottom-centre. Paytm logo top-left over the photo.",
  }, async (s, x, y, mh) => {
    const mw = 3.5;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: mw, h: mh, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: shadow() });
    // photo zone placeholder
    const ph = mh*0.56;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: mw, h: ph, fill: { color: CLOUD } });
    s.addImage({ data: await icon("FaUserTie", "#"+CYAND), x: x+mw/2-0.55, y: y+0.5, w: 1.1, h: 1.1 });
    s.addText("PHOTO AREA", { x, y: y+1.7, w: mw, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: CYAND, align: "center", charSpacing: 1, margin: 0 });
    s.addText("Indian SMB owner smiling at laptop\n(Paytm dashboard on screen)", { x: x+0.2, y: y+2.0, w: mw-0.4, h: 0.6, fontFace: BODY, fontSize: 9.5, italic: true, color: MUTED, align: "center", margin: 0 });
    logoLockup(s, x+0.22, y+0.2, 0.78, false);
    // navy band
    s.addShape(pres.shapes.RECTANGLE, { x, y: y+ph, w: mw, h: mh-ph, fill: { color: NAVY } });
    s.addText([
      { text: "Get paid online,\n", options: { color: WHITE, breakLine: true } },
      { text: "the easy way.", options: { color: CYAN } },
    ], { x: x+0.28, y: y+ph+0.14, w: mw-0.56, h: 0.85, fontFace: HEAD, fontSize: 18, bold: true, margin: 0, lineSpacingMultiple: 0.98 });
    // badges
    const bd = ["RBI-regulated","PCI-DSS","2.5cr+ merchants"];
    let bx = x+0.28;
    for (const b of bd){
      const bw = 0.35 + b.length*0.083;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y+ph+1.08, w: bw, h: 0.34, rectRadius: 0.17, fill: { color: "12345A" }, line:{type:"none"} });
      s.addText(b, { x: bx, y: y+ph+1.08, w: bw, h: 0.34, fontFace: BODY, fontSize: 8, bold: true, color: CLOUD, align: "center", valign: "middle", margin: 0 });
      bx += bw + 0.12;
    }
    pill(s, x+mw/2-0.95, y+mh-0.62, 1.9, 0.46, "Start Free ›", CYAN, NAVY, 12);
  });

  // =====================================================================
  // SLIDE 9 — AD COPY & CTAs
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "03", "Ad Copy & CTAs", "FaPenNib");
    s.addText("Short, punchy, conversion-first. Each line maps to one creative and one segment.",
      { x: M, y: 1.4, w: 12, h: 0.4, fontFace: BODY, fontSize: 13, italic: true, color: MUTED, margin: 0 });

    const hdr = (t)=>({text:t,options:{bold:true,color:WHITE,fill:{color:NAVY},align:"left",valign:"middle"}});
    const rows = [
      [ hdr("Creative"), hdr("Primary Text (1–2 lines)"), hdr("Headline (<10 words)"), hdr("CTA Button") ],
      [ {text:"1 · Speed = Conversion\n(Founder)",options:{bold:true,color:NAVY}},
        "Every extra second at checkout costs you a sale. Velocity collapses UPI into a single in-app tap — so more carts actually convert.",
        {text:"“Checkout in 0.8 seconds.”",options:{bold:true,color:NAVY}},
        {text:"Get API Keys",options:{bold:true,color:WHITE,fill:{color:CYAND},align:"center"}} ],
      [ {text:"2 · Ship it this weekend\n(Tech Lead)",options:{bold:true,color:NAVY}},
        "Clean docs, a real sandbox, and signed webhooks that fire on time. Integrate UPI in an afternoon — not a sprint.",
        {text:"“Ship UPI this weekend.”",options:{bold:true,color:NAVY}},
        {text:"Read the Docs",options:{bold:true,color:WHITE,fill:{color:CYAND},align:"center"}} ],
      [ {text:"3 · The easy way\n(Retail Owner)",options:{bold:true,color:NAVY}},
        "Accept UPI on your website without the tech headache. Plug in, get paid, and settle by tomorrow — backed by India’s most trusted payments brand.",
        {text:"“Get paid online, the easy way.”",options:{bold:true,color:NAVY}},
        {text:"Start Free",options:{bold:true,color:WHITE,fill:{color:CYAND},align:"center"}} ],
    ];
    s.addTable(rows, {
      x: M, y: 1.95, w: W-2*M, colW: [2.5, 5.25, 2.75, 1.59],
      rowH: [0.45, 1.32, 1.18, 1.32],
      border: { type: "solid", color: LINE, pt: 1 }, fill: { color: WHITE },
      fontFace: BODY, fontSize: 12, color: SLATE, valign: "middle", align: "left", margin: [5, 8, 5, 8],
    });

    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 6.55, w: W-2*M, h: 0.6, fill: { color: CLOUD }, line: { color: CYAN, width: 1 } });
    s.addText([
      { text: "CTA strategy:  ", options: { bold: true, color: NAVY } },
      { text: "“Read the Docs” / “Get API Keys” pull high-intent developers (low-friction, BOFU); “Start Free” lowers the bar for non-technical owners. Keep one CTA per ad — never stack two.", options: { color: SLATE } }
    ], { x: M+0.25, y: 6.55, w: W-2*M-0.5, h: 0.6, fontFace: BODY, fontSize: 11.5, valign: "middle", margin: 0 });
    footer(s, 9);
  }

  // =====================================================================
  // SLIDE 10 — TARGETING STRATEGY
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "04", "Targeting Strategy", "FaBullseye");

    const cols = [
      { ic:"FaLinkedin", clr:NAVY, name:"LinkedIn Ads", role:"ABM · decision-makers (BOFU)", split:"₹12.0L · 40%",
        lines:[["Job titles","Founder, Co-founder, CTO, VP Eng, Head of Payments / Product, Eng Manager, Lead Developer, Head of Growth / E-commerce"],
               ["Firmographics","Internet, E-commerce, Fintech, SaaS, Retail · 11–500 emp · India metros (BLR, MUM, DEL-NCR, PUN, HYD, CHN)"],
               ["Levers","Skills: Payment Gateways, APIs, Node/React · retarget site visitors · upload CRM/ABM lists"],
               ["Formats","Single-image, Document & Thought-leader ads, Conversation ads"]] },
      { ic:"FaMeta", clr:CYAND, name:"Meta Ads", role:"Reach + retargeting (TOFU/MOFU)", split:"₹10.5L · 35%",
        lines:[["Interests","Entrepreneurship, Shopify, WooCommerce, E-commerce, Razorpay, Stripe, Startup India, Digital marketing"],
               ["Behaviors","Small-business owners, Business-page admins, Engaged shoppers"],
               ["Lookalikes","1–3% LAL of Paytm for Business merchants & sign-up list"],
               ["Placements","FB/IG Feed, Reels, Stories · retarget visitors, video viewers, lead-form openers · Age 25–55, India"]] },
      { ic:"FaGoogle", clr:GREEN, name:"Google Search", role:"High-intent capture (BOFU)", split:"₹7.5L · 25%",
        lines:[["Keywords","“upi payment gateway api”, “payment gateway integration”, “razorpay alternative”, “upi api for website”, “collect upi payments api”"],
               ["Negatives","jobs, salary, login, customer care, refund, free, charges complaint"],
               ["Levers","Exact/phrase, match-type segmented · RLSA · Display/YouTube retargeting · weekly search-term mining"],
               ["Note","Long-tail intent keywords: ~40–60% cheaper, 2–3× higher conversion"]] },
    ];
    const cw = (W-2*M-2*0.35)/3, cy=1.7, ch=4.0;
    for (let i=0;i<3;i++){
      const x = M + i*(cw+0.35);
      s.addShape(pres.shapes.RECTANGLE, { x, y: cy, w: cw, h: ch, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: softShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: cy, w: cw, h: 0.95, fill: { color: cols[i].clr } });
      s.addImage({ data: await icon(cols[i].ic, "#"+WHITE), x: x+0.28, y: cy+0.26, w: 0.44, h: 0.44 });
      s.addText(cols[i].name, { x: x+0.85, y: cy+0.13, w: cw-1.0, h: 0.4, fontFace: HEAD, fontSize: 15, bold: true, color: WHITE, valign:"middle", margin: 0 });
      s.addText(cols[i].role, { x: x+0.85, y: cy+0.5, w: cw-1.0, h: 0.35, fontFace: BODY, fontSize: 10, color: CLOUD, margin: 0 });
      s.addText(cols[i].split, { x: x+0.0, y: cy+0.13, w: cw-0.2, h: 0.4, fontFace: HEAD, fontSize: 11, bold: true, color: WHITE, align:"right", valign:"middle", margin: 0 });
      let ly = cy+1.12;
      for (const ln of cols[i].lines){
        s.addText(ln[0].toUpperCase(), { x: x+0.28, y: ly, w: cw-0.5, h: 0.24, fontFace: HEAD, fontSize: 8.5, bold: true, color: cols[i].clr, charSpacing: 1, margin: 0 });
        s.addText(ln[1], { x: x+0.28, y: ly+0.23, w: cw-0.5, h: 0.5, fontFace: BODY, fontSize: 9.7, color: SLATE, lineSpacingMultiple: 0.98, valign:"top", margin: 0 });
        ly += 0.71;
      }
    }
    // funnel strip
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 5.95, w: W-2*M, h: 1.15, fill: { color: NAVY } });
    s.addText("FULL-FUNNEL FLOW", { x: M+0.3, y: 6.08, w: 3, h: 0.3, fontFace: HEAD, fontSize: 10, bold: true, color: CYAN, charSpacing: 1.5, margin: 0 });
    const fsteps = [["TOFU","Meta reach + LinkedIn thought-leader ads"],["MOFU","Retargeting · doc downloads · webinar · lead forms"],["BOFU","Google Search + LI Conversation ads → sandbox / demo"]];
    const fw = (W-2*M-0.6-2*0.3)/3;
    for (let i=0;i<3;i++){
      const x = M+0.3 + i*(fw+0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 6.42, w: fw, h: 0.55, rectRadius: 0.1, fill: { color: "12345A" }, line:{color:CYAND,width:1} });
      s.addText([{text:fsteps[i][0]+":  ",options:{bold:true,color:CYAN}},{text:fsteps[i][1],options:{color:"DCE7F8"}}],
        { x: x+0.18, y: 6.42, w: fw-0.3, h: 0.55, fontFace: BODY, fontSize: 9.5, valign:"middle", margin: 0 });
      if (i<2) s.addImage({ data: await icon("FaArrowRightLong","#"+CYAN), x: x+fw+0.02, y: 6.58, w: 0.24, h: 0.24 });
    }
    footer(s, 10);
  }

  // =====================================================================
  // SLIDE 11 — MOCK RESULTS: HEADLINE KPIs + FUNNEL
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    // header (light on dark variant)
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 0.5, w: 0.82, h: 0.82, rectRadius: 0.12, fill: { color: CYAN } });
    s.addText("05", { x: M, y: 0.5, w: 0.82, h: 0.82, fontFace: HEAD, fontSize: 22, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0 });
    s.addText("Mock Results — 30-Day Run", { x: M+1.02, y: 0.5, w: 7.4, h: 0.82, fontFace: HEAD, fontSize: 27, bold: true, color: WHITE, valign: "middle", margin: 0 });
    pill(s, W-3.95, 0.66, 3.35, 0.5, "₹30,00,000 SPEND · INDIA B2B", "12345A", CYAN, 10);

    // three hero KPI cards
    const heroes = [
      { v:"1.15%", l:"Blended Click-Through Rate", a:CYAN, ic:"FaArrowPointer" },
      { v:"₹67", l:"Blended Cost Per Click", a:WHITE, ic:"FaIndianRupeeSign" },
      { v:"4.2×", l:"ROAS — 12-mo projected, pipeline-weighted", a:GREEN, ic:"FaArrowTrendUp" },
    ];
    const hw = (W-2*M-2*0.4)/3, hy=1.65, hh=2.15;
    for (let i=0;i<3;i++){
      const x = M + i*(hw+0.4);
      s.addShape(pres.shapes.RECTANGLE, { x, y: hy, w: hw, h: hh, fill: { color: NAVY2 }, line: { color: "12345A", width: 1 }, shadow: shadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: hy, w: hw, h: 0.12, fill: { color: heroes[i].a } });
      s.addImage({ data: await icon(heroes[i].ic, "#"+heroes[i].a), x: x+hw-0.85, y: hy+0.32, w: 0.5, h: 0.5 });
      s.addText(heroes[i].v, { x: x+0.32, y: hy+0.45, w: hw-0.5, h: 1.0, fontFace: HEAD, fontSize: 50, bold: true, color: WHITE, margin: 0, valign:"middle" });
      s.addText(heroes[i].l, { x: x+0.32, y: hy+hh-0.62, w: hw-0.55, h: 0.55, fontFace: BODY, fontSize: 11.5, color: CLOUD, margin: 0, valign:"top" });
    }

    // funnel row
    s.addText("CONVERSION FUNNEL", { x: M, y: 4.15, w: 6, h: 0.3, fontFace: HEAD, fontSize: 11, bold: true, color: CYAN, charSpacing: 1.5, margin: 0 });
    const funnel = [["3.9M","Impressions"],["44,990","Clicks"],["2,700","Sandbox sign-ups"],["945","MQLs"],["280","SQLs / demos"],["58","Merchants live"]];
    const fw = (W-2*M-5*0.22)/6, fy=4.5, fh=1.0;
    for (let i=0;i<6;i++){
      const x = M + i*(fw+0.22);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: fy, w: fw, h: fh, rectRadius: 0.1, fill: { color: NAVY2 }, line: { color: CYAND, width: 1 } });
      s.addText(funnel[i][0], { x, y: fy+0.13, w: fw, h: 0.5, fontFace: HEAD, fontSize: 21, bold: true, color: CYAN, align:"center", margin: 0 });
      s.addText(funnel[i][1], { x: x+0.05, y: fy+0.6, w: fw-0.1, h: 0.35, fontFace: BODY, fontSize: 9, color: CLOUD, align:"center", margin: 0 });
      if (i<5) s.addImage({ data: await icon("FaArrowRightLong","#"+CYAN), x: x+fw-0.02, y: fy+fh/2-0.1, w: 0.2, h: 0.2 });
    }

    // secondary metrics chips
    const sec = ["Sign-up rate 6.0%","CAC ₹52,000 / merchant","LTV:CAC ≈ 2.7:1 (12-mo)","In-month realised ROAS ≈ 0.4× (fintech payback: 4–6 mo)"];
    let scx = M;
    for (const c of sec){
      const cw = 0.4 + c.length*0.087;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: scx, y: 5.95, w: cw, h: 0.5, rectRadius: 0.25, fill: { color: "FFFFFF", transparency: 90 }, line: { color: CYAN, width: 1 } });
      s.addText(c, { x: scx, y: 5.95, w: cw, h: 0.5, fontFace: HEAD, fontSize: 10, bold: true, color: WHITE, align:"center", valign:"middle", margin: 0 });
      scx += cw + 0.22;
    }
    s.addText("Illustrative figures, modelled on 2025–26 Indian B2B / fintech ad benchmarks. ROAS shown is projected (LTV-based), not in-month realised.",
      { x: M, y: 6.62, w: 12, h: 0.4, fontFace: BODY, fontSize: 9.5, italic: true, color: "9FB3D6", margin: 0 });
    footer(s, 11);
  }

  // =====================================================================
  // SLIDE 12 — CHANNEL BREAKDOWN + RATIONALE
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: CLOUD2 };
    await sectionHeader(s, "05", "Why These Numbers Make Sense", "FaChartLine");

    const hc = (t)=>({text:t,options:{bold:true,color:WHITE,fill:{color:NAVY},align:"center",valign:"middle"}});
    const r = (a)=>a.map((c,i)=>({text:c,options:{align: i===0?"left":"center",color: i===0?NAVY:SLATE, bold: i===0}}));
    const rows = [
      [hc("Channel"),hc("Spend"),hc("Impressions"),hc("Clicks"),hc("CTR"),hc("CPC"),hc("Sign-ups"),hc("Role")],
      r(["LinkedIn","₹12.0L","0.74M","4,290","0.58%","₹280","480","ABM · decision-makers"]),
      r(["Meta","₹10.5L","2.98M","32,800","1.10%","₹32","1,310","Reach + retargeting"]),
      r(["Google Search","₹7.5L","0.19M","7,900","4.20%","₹95","910","High-intent capture"]),
      [{text:"Blended",options:{bold:true,color:WHITE,fill:{color:CYAND},align:"left"}},
       ...["₹30.0L","3.91M","44,990","1.15%","₹67","2,700","Full-funnel"].map(c=>({text:c,options:{bold:true,color:WHITE,fill:{color:CYAND},align:"center"}}))],
    ];
    s.addTable(rows, {
      x: M, y: 1.7, w: W-2*M, colW: [1.85,1.15,1.6,1.35,1.0,1.0,1.3,2.84],
      rowH: [0.42,0.5,0.5,0.5,0.52],
      border: { type:"solid", color: LINE, pt: 1 }, fill: { color: WHITE },
      fontFace: BODY, fontSize: 11.5, valign: "middle", margin: [3,5,3,5],
    });

    // rationale cards
    const cards = [
      { ic:"FaArrowPointer", clr:CYAND, h:"CTR 1.15% — a realistic mix",
        b:"LinkedIn sponsored content runs 0.4–0.6% in India; Meta B2B ~1.1%; Google Search ~4.2% on high-intent terms. Weighted by spend, the blend lands at 1.15%." },
      { ic:"FaIndianRupeeSign", clr:NAVY, h:"CPC ₹67 — Indian B2B economics",
        b:"India LinkedIn CPC is ₹150–650 (fintech pushes the top end); Meta ₹10–40; Google payment-gateway keywords ₹60–180. The cheap-reach/expensive-intent mix averages to ₹67." },
      { ic:"FaArrowTrendUp", clr:GREEN, h:"ROAS 4.2× — LTV-driven, not in-month",
        b:"UPI P2M is zero-MDR, so revenue is platform fees + cross-sell to cards/EMI/VAS. 58 live + ~32 weighted-pipeline ≈ 90 merchants × ~₹1.4L first-year value ≈ ₹1.26Cr vs ₹30L. Payback is 4–6 months — month-1 realised ROAS is ~0.4× by design." },
    ];
    const cw = (W-2*M-2*0.35)/3, cy=4.5, ch=2.25;
    for (let i=0;i<3;i++){
      const x = M + i*(cw+0.35);
      s.addShape(pres.shapes.RECTANGLE, { x, y: cy, w: cw, h: ch, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: softShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: cy, w: cw, h: 0.1, fill: { color: cards[i].clr } });
      s.addShape(pres.shapes.OVAL, { x: x+0.28, y: cy+0.3, w: 0.55, h: 0.55, fill: { color: cards[i].clr } });
      s.addImage({ data: await icon(cards[i].ic, "#"+WHITE), x: x+0.41, y: cy+0.43, w: 0.3, h: 0.3 });
      s.addText(cards[i].h, { x: x+0.95, y: cy+0.3, w: cw-1.1, h: 0.6, fontFace: HEAD, fontSize: 12.5, bold: true, color: NAVY, valign:"middle", margin: 0 });
      s.addText(cards[i].b, { x: x+0.28, y: cy+1.0, w: cw-0.55, h: ch-1.1, fontFace: BODY, fontSize: 10.8, color: SLATE, lineSpacingMultiple: 1.02, valign:"top", margin: 0 });
    }
    footer(s, 12);
  }

  // =====================================================================
  // SLIDE 13 — NEXT STEPS / CLOSE
  // =====================================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY2 };
    s.addShape(pres.shapes.OVAL, { x: 10.2, y: 3.9, w: 5.2, h: 5.2, fill: { color: NAVY }, line:{type:"none"} });
    s.addShape(pres.shapes.OVAL, { x: -1.6, y: -1.8, w: 4.4, h: 4.4, fill: { color: CYAND, transparency: 80 }, line:{type:"none"} });
    logoLockup(s, M, 0.62, 1.0, true);
    pill(s, W-3.55, 0.66, 2.95, 0.5, "MOCK CAMPAIGN · FICTIONAL", "12345A", CYAN, 10);

    s.addText("Ready to build.", { x: M, y: 2.0, w: 11, h: 0.9, fontFace: HEAD, fontSize: 44, bold: true, color: WHITE, margin: 0 });
    s.addText("From this brief to live in two weeks.", { x: M, y: 2.95, w: 11, h: 0.5, fontFace: HEAD, fontSize: 18, italic: true, color: CYAN, margin: 0 });

    const steps = [
      ["FaPalette","Design the 3 creatives in Canva/Figma from the spec sheets (Slides 6–8)."],
      ["FaCode","Wire tracking: Meta Pixel + CAPI, LinkedIn Insight Tag, GA4; build sandbox sign-up & demo flows."],
      ["FaBolt","Launch a 2-week test (~₹6L) across all 3 channels; optimise to CPL → activation, not clicks."],
      ["FaArrowTrendUp","Scale winning ad-sets & keywords; shift budget to the lowest cost-per-activated-merchant."],
    ];
    let yy = 3.85;
    for (const st of steps){
      s.addShape(pres.shapes.OVAL, { x: M, y: yy, w: 0.5, h: 0.5, fill: { color: CYAN } });
      s.addImage({ data: await icon(st[0], "#"+NAVY), x: M+0.12, y: yy+0.12, w: 0.26, h: 0.26 });
      s.addText(st[1], { x: M+0.72, y: yy-0.05, w: 9.6, h: 0.6, fontFace: BODY, fontSize: 13.5, color: "DCE7F8", valign:"middle", margin: 0 });
      yy += 0.72;
    }
    s.addText("Figures illustrative, grounded in 2025–26 Indian B2B / fintech ad benchmarks (LinkedIn, Meta, Google). Paytm Velocity API is a fictional product for this exercise.",
      { x: M, y: 6.95, w: 12, h: 0.4, fontFace: BODY, fontSize: 9.5, italic: true, color: "8FA6CC", margin: 0 });
  }

  await pres.writeFile({ fileName: "Paytm_Velocity_API_Campaign_Brief.pptx" });
  console.log("DECK WRITTEN");
})().catch(e => { console.error(e); process.exit(1); });

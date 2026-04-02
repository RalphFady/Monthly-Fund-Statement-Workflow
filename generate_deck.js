/**
 * Monthly Fund Statement — PowerPoint Deck Generator
 * Cana Capital Senior Credit Fund I, L.P. | April 2026
 */
const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ── BRAND ────────────────────────────────────────────────
const NAVY      = "1B2A4A";
const NAVY_DARK = "0D1925";
const NAVY_MID  = "243454";
const GOLD      = "C9A84C";
const BG_LIGHT  = "F7F8FA";
const TEXT_DARK = "2C3E50";
const MUTED     = "8899AA";
const GREEN     = "1A6B3A";
const RED       = "B71C1C";
const MED_BLUE  = "3D6B9E";
const ALT_ROW   = "EEF2F7";
const WHITE     = "FFFFFF";

const FUND_NAME  = "Cana Capital Senior Credit Fund I, L.P.";
const FUND_SHORT = "Cana Capital Fund I";
const STRATEGY   = "Senior Secured Commercial Real Estate Debt";
const MANAGER    = "Cana Ventures";
const PERIOD     = "April 2026";
const CONTACT    = "ir@canaventures.com  |  www.canaventures.com";

// ── MOCK DATA ─────────────────────────────────────────────
const cur = {
  nav: 118.4, gav: 145.2, committed: 150.0, called: 112.5, uncalled: 37.5,
  nav_per_unit: 105.24, loan_book: 138.2, active_loans: 14,
  wtd_ltv: 62.5, wtd_dscr: 1.35, wtd_coupon: 8.85, avg_term: 18, occupancy: 91.3,
  monthly_dist: 0.095, ann_yield: 8.2, pref_return: 7.5, pref_current: true,
  cumulative_dist: 6.84, net_irr: 9.1, moic: 1.08, tvpi: 1.08,
  new_orig_count: 2, new_orig_total: 14.0,
  repay_count: 1,  repay_total: 9.2,
  pipeline_count: 3, pipeline_total: 22.0,
};
const prev = {
  nav: 116.5, loan_book: 136.0, active_loans: 13,
  wtd_ltv: 63.1, wtd_dscr: 1.33, wtd_coupon: 8.75, avg_term: 19, occupancy: 90.1,
  monthly_dist: 0.093, ann_yield: 8.0, called: 110.5, uncalled: 39.5,
};

const PROPERTY_TYPES = [
  ["Multifamily", 35, 48.4], ["Industrial", 28, 38.7],
  ["Mixed-Use", 18, 24.9],   ["Retail", 12, 16.6],  ["Office", 7, 9.7],
];
const GEOGRAPHY = [
  ["Texas",      5, 44.2, 32], ["Florida",    3, 33.2, 24],
  ["Southeast",  3, 26.3, 19], ["Southwest",  2, 20.7, 15], ["Other", 1, 13.8, 10],
];
const DEALS = [
  ["Riverside Commons", "Austin, TX", "$8.5M", "58.2%", "SOFR+350", "24 mo", "Multifamily"],
  ["Harbor Point Industrial", "Tampa, FL", "$5.5M", "63.8%", "SOFR+375", "18 mo", "Industrial"],
];
const DIST_HISTORY = [
  ["Nov 2025","$0.088","7.8%"],["Dec 2025","$0.090","7.9%"],
  ["Jan 2026","$0.091","8.0%"],["Feb 2026","$0.092","8.1%"],
  ["Mar 2026","$0.093","8.0%"],["Apr 2026","$0.095","8.2%"],
];

// ── HELPERS ───────────────────────────────────────────────
const fmtM   = v => `$${v.toFixed(1)}M`;
const fmtPct = v => `${v.toFixed(1)}%`;
const fmtX   = v => `${v.toFixed(2)}x`;
const delta  = (a,b) => { const d=(a-b)/b*100; return `${d>=0?"▲":"▼"} ${Math.abs(d).toFixed(1)}%`; };
const bps    = (a,b) => { const d=(a-b)*100; return `${d>=0?"▲":"▼"} ${Math.abs(d).toFixed(0)} bps`; };
const dColor = (a,b,posGood=true) => { const up=a>b; return (up===posGood)?GREEN:RED; };

// ── PPTX SETUP ────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";   // 13.33" × 7.5"
pres.author  = MANAGER;
pres.subject = `${FUND_SHORT} — ${PERIOD} Investor Report`;
pres.title   = `${FUND_SHORT} Monthly Investor Report`;

const TOTAL_SLIDES = 9;

// ── FOOTER HELPER ─────────────────────────────────────────
function addFooter(slide, slideNum, darkBg=false) {
  if (slideNum === 1) return;
  const col = darkBg ? "8899AA" : "8899AA";
  slide.addText("CONFIDENTIAL — FOR LIMITED PARTNERS ONLY", {
    x:0.35, y:7.18, w:5.5, h:0.2, fontSize:7, color:col, fontFace:"Calibri Light"
  });
  slide.addText(`${slideNum} of ${TOTAL_SLIDES}`, {
    x:6.3, y:7.18, w:0.8, h:0.2, fontSize:7, color:col, fontFace:"Calibri Light", align:"center"
  });
  slide.addText(`${FUND_SHORT}  |  ${PERIOD}`, {
    x:9.0, y:7.18, w:4.1, h:0.2, fontSize:7, color:col, fontFace:"Calibri Light", align:"right"
  });
}

// ── SLIDE 1: COVER ────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: NAVY };

  // Top bar
  slide.addShape(pres.ShapeType.rect, {
    x:0, y:0, w:"100%", h:0.7, fill:{color:NAVY_DARK}
  });
  slide.addText(MANAGER.toUpperCase(), {
    x:0.4, y:0.2, w:5, h:0.35, fontSize:13, bold:true, color:GOLD, fontFace:"Calibri"
  });
  slide.addText("www.canaventures.com", {
    x:8.5, y:0.2, w:4.5, h:0.35, fontSize:10, color:MUTED, fontFace:"Calibri Light", align:"right"
  });

  // Gold rule
  slide.addShape(pres.ShapeType.rect, { x:0.4, y:3.0, w:12.5, h:0.025, fill:{color:GOLD} });

  // Fund name
  slide.addText(FUND_NAME, {
    x:0.4, y:3.1, w:12.5, h:0.9, fontSize:36, bold:true, color:WHITE, fontFace:"Calibri"
  });
  slide.addText(STRATEGY, {
    x:0.4, y:3.95, w:12.5, h:0.4, fontSize:15, italic:true, color:GOLD, fontFace:"Calibri Light"
  });

  // Report type + period
  slide.addText("Monthly Investor Report", {
    x:0.4, y:4.65, w:12.5, h:0.35, fontSize:14, color:"AABDD0", fontFace:"Calibri Light"
  });
  slide.addText(PERIOD, {
    x:0.4, y:5.0, w:12.5, h:0.6, fontSize:32, bold:true, color:WHITE, fontFace:"Calibri"
  });

  // Bottom rule + confidentiality
  slide.addShape(pres.ShapeType.rect, { x:0, y:6.9, w:"100%", h:0.025, fill:{color:GOLD} });
  slide.addText("CONFIDENTIAL — FOR LIMITED PARTNERS ONLY", {
    x:0.4, y:6.95, w:7, h:0.3, fontSize:8, color:MUTED, fontFace:"Calibri Light"
  });
  slide.addText(`Prepared by ${MANAGER}  |  ${PERIOD}`, {
    x:6.5, y:6.95, w:6.5, h:0.3, fontSize:8, color:MUTED, fontFace:"Calibri Light", align:"right"
  });
}

// ── STAT BOX HELPER ───────────────────────────────────────
function addStatBox(slide, x, y, value, label, deltaText, deltaColor) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w:3.9, h:1.55, fill:{color:WHITE},
    line:{color:GOLD, width:0.75}, rectRadius:0.05
  });
  slide.addText(value, {
    x:x+0.1, y:y+0.1, w:3.7, h:0.85,
    fontSize:34, bold:true, color:GOLD, fontFace:"Calibri", align:"center"
  });
  slide.addText(label, {
    x:x+0.1, y:y+0.95, w:3.7, h:0.3,
    fontSize:9, color:TEXT_DARK, fontFace:"Calibri Light", align:"center"
  });
  slide.addText(deltaText, {
    x:x+0.1, y:y+1.23, w:3.7, h:0.25,
    fontSize:9, bold:true, color:deltaColor, fontFace:"Calibri", align:"center"
  });
}

// ── SLIDE 2: EXECUTIVE SUMMARY ────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color:BG_LIGHT };
  slide.addShape(pres.ShapeType.rect, {x:0,y:0,w:0.08,h:7.5,fill:{color:NAVY}});
  slide.addText("Executive Summary", {x:0.3,y:0.2,w:12,h:0.6,fontSize:26,bold:true,color:NAVY,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect, {x:0.3,y:0.82,w:12.7,h:0.022,fill:{color:GOLD}});

  // Stat boxes
  addStatBox(slide, 0.3,  0.95, fmtM(cur.nav),           "Total Fund NAV",        delta(cur.nav, prev.nav),           dColor(cur.nav,prev.nav));
  addStatBox(slide, 4.52, 0.95, `${fmtPct(cur.ann_yield)}`, "Distribution Yield",  bps(cur.ann_yield, prev.ann_yield), dColor(cur.ann_yield,prev.ann_yield));
  addStatBox(slide, 8.73, 0.95, String(cur.active_loans), "Active Loans",          "▲ +1 loan vs. prior month",        GREEN);

  // Narrative
  slide.addText(
    "The Fund delivered another month of stable performance in April 2026, with Net Asset Value growing to $118.4M — a 1.6% increase from March — driven by continued interest income and two new loan originations. The portfolio remains 100% performing with a weighted average DSCR of 1.35x across all 14 active loans.\n\nMonthly distributions of $0.095 per unit were paid on schedule, representing an annualized yield of 8.2% and maintaining the fund's preferred return of 7.5% in current status. The fund's defensive positioning in senior secured loans — with a weighted average LTV of 62.5% — continues to provide robust downside protection in the current rate environment.",
    {x:0.3, y:2.7, w:9.5, h:3.8, fontSize:11, color:TEXT_DARK, fontFace:"Calibri Light", valign:"top"}
  );

  // Health sidebar
  slide.addShape(pres.ShapeType.roundRect, {x:10.1,y:2.7,w:3.0,h:3.0,fill:{color:NAVY},rectRadius:0.05});
  slide.addText("Fund Health", {x:10.1,y:2.82,w:3.0,h:0.35,fontSize:11,bold:true,color:WHITE,fontFace:"Calibri",align:"center"});
  slide.addShape(pres.ShapeType.rect,{x:10.1,y:3.17,w:3.0,h:0.015,fill:{color:GOLD}});
  const health = [
    ["✓ Preferred Return: CURRENT", GREEN],
    ["✓ Portfolio: 100% Performing", GREEN],
    [`✓ Capital: ${Math.round(cur.called/cur.committed*100)}% Deployed`, GREEN],
    [`Net IRR: ${fmtPct(cur.net_irr)}`, WHITE],
    [`MOIC: ${fmtX(cur.moic)}`, GOLD],
  ];
  health.forEach(([txt,col],i) => {
    slide.addText(txt, {x:10.25, y:3.28+i*0.48, w:2.7, h:0.4, fontSize:9.5, color:col, fontFace:"Calibri Light"});
  });

  addFooter(slide, 2);
}

// ── SLIDE 3: FUND OVERVIEW (NAV & SIZE) ──────────────────
{
  const slide = pres.addSlide();
  slide.background = { color:BG_LIGHT };
  slide.addText(`Fund Overview — ${PERIOD}`, {x:0.35,y:0.15,w:12.5,h:0.55,fontSize:26,bold:true,color:NAVY,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect,{x:0.35,y:0.72,w:12.6,h:0.022,fill:{color:GOLD}});

  // Use a table-based 2×3 grid for reliable cross-platform rendering
  const metrics = [
    { val:fmtM(cur.nav),      lbl:"Total Fund NAV",         dlt:delta(cur.nav,prev.nav),              dcol:GREEN  },
    { val:fmtM(cur.gav),      lbl:"Gross Asset Value",       dlt:"▲ +$2.1M vs. prior month",           dcol:GREEN  },
    { val:fmtM(cur.called),   lbl:"Total Called Capital",    dlt:delta(cur.called,prev.called),        dcol:GREEN  },
    { val:fmtM(cur.uncalled), lbl:"Dry Powder / Uncalled",   dlt:`${fmtPct(cur.uncalled/cur.committed*100)} of committed`, dcol:MUTED },
    { val:fmtM(cur.committed),lbl:"Total Committed Capital", dlt:"Fund target",                        dcol:MUTED  },
    { val:`$${cur.nav_per_unit.toFixed(2)}`, lbl:"NAV per Unit", dlt:"▲ +$1.59 vs. prior month",      dcol:GREEN  },
  ];

  // Draw each stat box as explicit positioned elements in two rows
  const boxW = 4.05, boxH = 2.0;
  const row1Y = 0.92, row2Y = 3.12;
  const col1X = 0.35, col2X = 4.6, col3X = 8.85;
  const positions = [
    [col1X, row1Y], [col2X, row1Y], [col3X, row1Y],
    [col1X, row2Y], [col2X, row2Y], [col3X, row2Y],
  ];

  metrics.forEach(({val,lbl,dlt,dcol}, i) => {
    const [bx, by] = positions[i];
    // Box background (navy border on light bg)
    slide.addShape(pres.ShapeType.roundRect, {
      x:bx, y:by, w:boxW, h:boxH,
      fill:{color:WHITE}, line:{color:NAVY, width:1.0}, rectRadius:0.06
    });
    // Top accent bar inside box
    slide.addShape(pres.ShapeType.rect, {
      x:bx, y:by, w:boxW, h:0.06, fill:{color:NAVY}, line:{type:"none"}
    });
    // Value
    slide.addText(val, {
      x:bx+0.1, y:by+0.18, w:boxW-0.2, h:1.0,
      fontSize:38, bold:true, color:GOLD, fontFace:"Calibri", align:"center", valign:"middle"
    });
    // Label
    slide.addText(lbl, {
      x:bx+0.1, y:by+1.2, w:boxW-0.2, h:0.38,
      fontSize:10, color:TEXT_DARK, fontFace:"Calibri Light", align:"center"
    });
    // Delta
    slide.addText(dlt, {
      x:bx+0.1, y:by+1.6, w:boxW-0.2, h:0.3,
      fontSize:9, bold:true, color:dcol, fontFace:"Calibri", align:"center"
    });
  });

  addFooter(slide, 3);
}

// ── SLIDE 4: LOAN PORTFOLIO SNAPSHOT ─────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color:BG_LIGHT };
  slide.addText("Loan Portfolio Snapshot", {x:0.35,y:0.2,w:12.5,h:0.6,fontSize:26,bold:true,color:NAVY,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect,{x:0.35,y:0.82,w:12.6,h:0.022,fill:{color:GOLD}});

  // Metrics table (left)
  const metrics = [
    ["Total Loan Book Balance",      fmtM(cur.loan_book),     fmtM(prev.loan_book),    delta(cur.loan_book,prev.loan_book),       dColor(cur.loan_book,prev.loan_book)],
    ["Number of Active Loans",       "14",                     "13",                    "▲ +1 loan",                               GREEN],
    ["Wtd. Avg. LTV",                fmtPct(cur.wtd_ltv),     fmtPct(prev.wtd_ltv),    bps(cur.wtd_ltv, prev.wtd_ltv),           GREEN],
    ["Wtd. Avg. DSCR",               fmtX(cur.wtd_dscr),      fmtX(prev.wtd_dscr),     `▲ ${((cur.wtd_dscr-prev.wtd_dscr)).toFixed(2)}x`, GREEN],
    ["Wtd. Avg. Coupon",             fmtPct(cur.wtd_coupon),  fmtPct(prev.wtd_coupon), bps(cur.wtd_coupon,prev.wtd_coupon),      GREEN],
    ["Avg. Remaining Term",          "18 months",              "19 months",             "▼ 1 month",                               MUTED],
    ["Portfolio Occupancy",          fmtPct(cur.occupancy),   fmtPct(prev.occupancy),  bps(cur.occupancy,prev.occupancy),        GREEN],
    ["Performing Loans",             "14 / 100%",              "13 / 100%",             "—",                                      MUTED],
  ];

  const colW = [2.8, 1.05, 1.05, 1.1];
  const tableRows = [
    [{text:"Metric",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},
     {text:"Apr 2026",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Mar 2026",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Δ Change",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}}],
    ...metrics.map(([lbl,cv,pv,dt,dc],i) => [
      {text:lbl,    options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},
      {text:cv,     options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri",bold:true,color:GOLD,align:"right"}},
      {text:pv,     options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:MUTED,align:"right"}},
      {text:dt,     options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri",bold:true,color:dc,align:"right"}},
    ])
  ];

  slide.addTable(tableRows, {
    x:0.35, y:1.05, w:6.0, colW,
    border:{type:"solid",color:"DDDDDD",pt:0.3},
    rowH:0.38, fontFace:"Calibri",
  });

  // Property type breakdown (right) — bar chart style, no table overlap
  slide.addText("Property Type Mix", {x:6.8,y:1.0,w:6.2,h:0.35,fontSize:11,bold:true,color:NAVY,fontFace:"Calibri"});

  // Header row
  slide.addShape(pres.ShapeType.rect,{x:6.8,y:1.38,w:6.18,h:0.36,fill:{color:NAVY},line:{type:"none"}});
  slide.addText("Property Type",  {x:6.9, y:1.42,w:2.0,h:0.28,fontSize:9,bold:true,color:WHITE,fontFace:"Calibri"});
  slide.addText("Allocation",     {x:9.0, y:1.42,w:1.5,h:0.28,fontSize:9,bold:true,color:WHITE,fontFace:"Calibri",align:"center"});
  slide.addText("Balance",        {x:10.6,y:1.42,w:1.3,h:0.28,fontSize:9,bold:true,color:WHITE,fontFace:"Calibri",align:"right"});

  const propColors = [NAVY, GOLD, MED_BLUE, "5A8FA8", "8BADBF"];
  PROPERTY_TYPES.forEach(([pt,pct,bal],i) => {
    const ry = 1.74 + i*0.42;
    const rowFill = i%2===0 ? WHITE : ALT_ROW;
    slide.addShape(pres.ShapeType.rect,{x:6.8,y:ry,w:6.18,h:0.42,fill:{color:rowFill},line:{type:"none"}});
    // Property name
    slide.addText(pt, {x:6.9,y:ry+0.06,w:1.9,h:0.3,fontSize:9,color:TEXT_DARK,fontFace:"Calibri Light"});
    // Visual bar (proportional, max 35% = 1.8" wide)
    const barW = Math.max(0.08, (pct/35)*1.8);
    slide.addShape(pres.ShapeType.roundRect,{x:8.85,y:ry+0.1,w:barW,h:0.22,fill:{color:propColors[i]},line:{type:"none"},rectRadius:0.02});
    // Percentage label after bar
    slide.addText(`${pct}%`,{x:8.85+barW+0.08,y:ry+0.06,w:0.5,h:0.3,fontSize:9,bold:true,color:GOLD,fontFace:"Calibri"});
    // Balance
    slide.addText(`$${bal}M`,{x:10.6,y:ry+0.06,w:1.3,h:0.3,fontSize:9,color:TEXT_DARK,fontFace:"Calibri Light",align:"right"});
    // Row border
    slide.addShape(pres.ShapeType.rect,{x:6.8,y:ry+0.41,w:6.18,h:0.01,fill:{color:"DDDDDD"},line:{type:"none"}});
  });

  addFooter(slide, 4);
}

// ── SLIDE 5: DISTRIBUTIONS & RETURNS ─────────────────────
{
  const slide = pres.addSlide();

  // Dark left half
  slide.addShape(pres.ShapeType.rect,{x:0,y:0,w:6.5,h:7.5,fill:{color:NAVY}});
  // Light right half
  slide.addShape(pres.ShapeType.rect,{x:6.5,y:0,w:6.83,h:7.5,fill:{color:WHITE}});

  slide.addText("Distributions & Returns",{x:0.35,y:0.2,w:6.0,h:0.6,fontSize:22,bold:true,color:WHITE,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect,{x:0.35,y:0.82,w:5.8,h:0.022,fill:{color:GOLD}});

  // Stat stack (left)
  const stats = [
    [`$${cur.monthly_dist.toFixed(3)}`, "Monthly Distribution / Unit", bps(cur.monthly_dist,prev.monthly_dist).replace("bps","¢ vs. prior"), GREEN],
    [fmtPct(cur.ann_yield),             "Annualized Distribution Yield", bps(cur.ann_yield,prev.ann_yield), GREEN],
    [fmtPct(cur.pref_return) + " p.a.", "Preferred Return — CURRENT ✓", "In current status — no accrued shortfall", GREEN],
    [fmtM(cur.cumulative_dist),         "Cumulative Distributions Paid", "Since fund inception (2023)", MUTED],
  ];
  stats.forEach(([val,lbl,sub,col],i) => {
    slide.addText(val, {x:0.4,y:1.1+i*1.45,w:5.8,h:0.75,fontSize:30,bold:true,color:GOLD,fontFace:"Calibri",align:"center"});
    slide.addText(lbl, {x:0.4,y:1.83+i*1.45,w:5.8,h:0.28,fontSize:9.5,color:"AABDD0",fontFace:"Calibri Light",align:"center"});
    slide.addText(sub, {x:0.4,y:2.1+i*1.45,w:5.8,h:0.25,fontSize:8,bold:true,color:col,fontFace:"Calibri",align:"center"});
    if(i<3) slide.addShape(pres.ShapeType.rect,{x:0.5,y:2.35+i*1.45,w:5.6,h:0.012,fill:{color:"243454"}});
  });

  // Returns table (right)
  slide.addText("Returns Summary", {x:6.8,y:0.9,w:6.2,h:0.35,fontSize:11,bold:true,color:NAVY,fontFace:"Calibri"});
  const retRows = [
    [{text:"Return Metric",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},
     {text:"Net",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Gross",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}}],
    [{text:"Net IRR (Since Inception)",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:"9.1%",options:{fill:{color:WHITE},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}},{text:"10.8%",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
    [{text:"MOIC (Net)",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:"1.08x",options:{fill:{color:ALT_ROW},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}},{text:"N/A",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:MUTED,align:"right"}}],
    [{text:"TVPI (Net)",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:"1.08x",options:{fill:{color:WHITE},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}},{text:"N/A",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:MUTED,align:"right"}}],
    [{text:"Preferred Hurdle Rate",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:"7.50%",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}},{text:"—",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:MUTED,align:"right"}}],
  ];
  slide.addTable(retRows,{x:6.8,y:1.28,w:6.1,colW:[3.6,1.2,1.2],border:{type:"solid",color:"DDDDDD",pt:0.3},rowH:0.38,fontFace:"Calibri"});

  // Distribution history table
  slide.addText("Distribution History", {x:6.8,y:3.5,w:6.1,h:0.35,fontSize:11,bold:true,color:NAVY,fontFace:"Calibri"});
  const distRows = [
    [{text:"Period",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},
     {text:"$/Unit",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Ann. Yield",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}}],
    ...DIST_HISTORY.map(([p,d,y],i) => [
      {text:p,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:i===5?"Calibri":"Calibri Light",bold:i===5,color:i===5?NAVY:TEXT_DARK}},
      {text:d,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:i===5?"Calibri":"Calibri Light",bold:i===5,color:i===5?GOLD:TEXT_DARK,align:"right"}},
      {text:y,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:i===5?"Calibri":"Calibri Light",bold:i===5,color:i===5?GOLD:TEXT_DARK,align:"right"}},
    ])
  ];
  slide.addTable(distRows,{x:6.8,y:3.88,w:6.1,colW:[3.0,1.5,1.5],border:{type:"solid",color:"DDDDDD",pt:0.3},rowH:0.36,fontFace:"Calibri"});

  addFooter(slide, 5);
}

// ── SLIDE 6: DEAL ACTIVITY ────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color:WHITE };
  slide.addText(`Deal Activity — ${PERIOD}`,{x:0.35,y:0.2,w:12.5,h:0.6,fontSize:26,bold:true,color:NAVY,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect,{x:0.35,y:0.82,w:12.6,h:0.022,fill:{color:GOLD}});

  // 3-stat boxes
  const boxes = [
    [`${cur.new_orig_count} loans / ${fmtM(cur.new_orig_total)}`, "New Originations", NAVY],
    [`${cur.repay_count} loan / ${fmtM(cur.repay_total)}`, "Repayments / Exits", MED_BLUE],
    [`${cur.pipeline_count} loans / ${fmtM(cur.pipeline_total)}`, "Pipeline (Q2 2026)", "4A6741"],
  ];
  boxes.forEach(([val,lbl,col],i) => {
    slide.addShape(pres.ShapeType.roundRect,{x:0.35+i*4.32,y:1.1,w:4.1,h:1.3,fill:{color:col},rectRadius:0.05});
    slide.addText(val,{x:0.45+i*4.32,y:1.2,w:3.9,h:0.7,fontSize:20,bold:true,color:WHITE,fontFace:"Calibri",align:"center"});
    slide.addText(lbl,{x:0.45+i*4.32,y:1.88,w:3.9,h:0.3,fontSize:10,color:"AACCEE",fontFace:"Calibri Light",align:"center"});
  });

  // New deals table
  slide.addText("New Originations — April 2026",{x:0.35,y:2.65,w:12.6,h:0.35,fontSize:11,bold:true,color:NAVY,fontFace:"Calibri"});
  const dealRows = [
    [{text:"Property",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},
     {text:"Market",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},
     {text:"Loan Amt",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"LTV",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Coupon",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Term",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Type",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}}],
    ...DEALS.map(([prop,mkt,loan,ltv,cpn,trm,typ],i) => [
      {text:prop,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},
      {text:mkt, options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},
      {text:loan,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}},
      {text:ltv, options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}},
      {text:cpn, options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}},
      {text:trm, options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}},
      {text:typ, options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},
    ])
  ];
  slide.addTable(dealRows,{x:0.35,y:3.03,w:12.6,colW:[2.5,1.5,1.0,0.75,1.25,0.75,1.5],border:{type:"solid",color:"DDDDDD",pt:0.3},rowH:0.4,fontFace:"Calibri"});

  slide.addText(
    "The fund closed two new senior secured loans in April, deploying $14.0M at a blended weighted average LTV of 60.5% and coupon of SOFR+360 bps. Riverside Commons (Austin) represents the fund's fifth Texas exposure, reinforcing the Sunbelt concentration strategy. One $9.2M retail center loan repaid at par. The pipeline of $22M is expected to close in Q2 2026.",
    {x:0.35,y:4.85,w:12.6,h:1.2,fontSize:11,color:TEXT_DARK,fontFace:"Calibri Light"}
  );

  addFooter(slide, 6);
}

// ── SLIDE 7: PORTFOLIO COMPOSITION ────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color:BG_LIGHT };
  slide.addText("Portfolio Composition",{x:0.35,y:0.2,w:12.5,h:0.6,fontSize:26,bold:true,color:NAVY,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect,{x:0.35,y:0.82,w:12.6,h:0.022,fill:{color:GOLD}});

  // Geography table
  slide.addText("By Geography (MSA / State)",{x:0.35,y:1.05,w:6.0,h:0.35,fontSize:11,bold:true,color:NAVY,fontFace:"Calibri"});
  const geoRows = [
    [{text:"Market",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},
     {text:"# Loans",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"Deployed",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}},
     {text:"% Portfolio",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}}],
    ...GEOGRAPHY.map(([mkt,cnt,dep,pct],i) => [
      {text:mkt,    options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},
      {text:String(cnt),options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}},
      {text:`$${dep}M`,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}},
      {text:`${pct}%`,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}},
    ]),
    [{text:"Total",options:{fill:{color:BG_LIGHT},fontSize:9,bold:true,fontFace:"Calibri",color:NAVY}},
     {text:"14",options:{fill:{color:BG_LIGHT},fontSize:9,bold:true,fontFace:"Calibri",align:"right"}},
     {text:"$138.2M",options:{fill:{color:BG_LIGHT},fontSize:9,bold:true,fontFace:"Calibri",align:"right"}},
     {text:"100%",options:{fill:{color:BG_LIGHT},fontSize:9,bold:true,color:GOLD,fontFace:"Calibri",align:"right"}}],
  ];
  slide.addTable(geoRows,{x:0.35,y:1.43,w:6.3,colW:[1.9,0.9,1.4,1.5],border:{type:"solid",color:"DDDDDD",pt:0.3},rowH:0.39,fontFace:"Calibri"});

  // Vertical divider
  slide.addShape(pres.ShapeType.rect,{x:6.95,y:1.0,w:0.022,h:4.0,fill:{color:GOLD}});

  // Property type table
  slide.addText("By Property Type",{x:7.2,y:1.05,w:5.8,h:0.35,fontSize:11,bold:true,color:NAVY,fontFace:"Calibri"});
  const propRows2 = [
    [{text:"Property Type",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},
     {text:"% Allocation",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"center",fontFace:"Calibri",fontSize:9}},
     {text:"Balance",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}}],
    ...PROPERTY_TYPES.map(([pt,pct,bal],i) => [
      {text:pt,      options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},
      {text:`${pct}%`,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,bold:true,color:GOLD,align:"center",fontFace:"Calibri"}},
      {text:`$${bal}M`,options:{fill:{color:i%2===0?WHITE:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}},
    ])
  ];
  slide.addTable(propRows2,{x:7.2,y:1.43,w:5.8,colW:[2.3,1.6,1.6],border:{type:"solid",color:"DDDDDD",pt:0.3},rowH:0.39,fontFace:"Calibri"});

  // Loan size section
  slide.addShape(pres.ShapeType.rect,{x:0.35,y:5.2,w:12.6,h:0.022,fill:{color:GOLD}});
  slide.addText("Loan Size Concentration",{x:0.35,y:5.32,w:12.6,h:0.35,fontSize:11,bold:true,color:NAVY,fontFace:"Calibri"});
  const szBoxes = [
    ["< $5M", "2 loans", "5.7%"],["$5M–$10M","6 loans","35.0%"],
    ["$10M–$20M","5 loans","50.0%"],["> $20M","1 loan","9.3%"],
  ];
  szBoxes.forEach(([sz,cnt,pct],i) => {
    slide.addShape(pres.ShapeType.roundRect,{x:0.35+i*3.22,y:5.72,w:3.0,h:1.3,fill:{color:WHITE},line:{color:NAVY,width:0.5},rectRadius:0.04});
    slide.addText(pct,  {x:0.45+i*3.22,y:5.82,w:2.8,h:0.55,fontSize:26,bold:true,color:GOLD,fontFace:"Calibri",align:"center"});
    slide.addText(cnt,  {x:0.45+i*3.22,y:6.35,w:2.8,h:0.25,fontSize:9,color:TEXT_DARK,fontFace:"Calibri Light",align:"center"});
    slide.addText(sz,   {x:0.45+i*3.22,y:6.6, w:2.8,h:0.25,fontSize:8,bold:true,color:NAVY,fontFace:"Calibri",align:"center"});
  });

  addFooter(slide, 7);
}

// ── SLIDE 8: OUTLOOK ──────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color:NAVY };

  slide.addText("Outlook",{x:0.4,y:0.2,w:12.5,h:0.6,fontSize:32,bold:true,color:WHITE,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect,{x:0.4,y:0.82,w:12.5,h:0.022,fill:{color:GOLD}});

  // Pull quote box
  slide.addShape(pres.ShapeType.roundRect,{x:1.0,y:1.2,w:11.3,h:1.5,fill:{color:NAVY_MID},line:{type:"none"},rectRadius:0.05});
  slide.addShape(pres.ShapeType.rect,{x:1.0,y:1.2,w:0.07,h:1.5,fill:{color:GOLD}});
  slide.addText(
    "$22 million in pipeline loans expected to close in Q2 2026, increasing portfolio utilization to approximately 97% of called capital.",
    {x:1.2,y:1.3,w:10.9,h:1.3,fontSize:18,bold:true,color:WHITE,fontFace:"Calibri",valign:"middle"}
  );

  // Main narrative
  slide.addText(
    "The fund's pipeline of $22 million in near-term originations positions the portfolio for continued income growth through Q2 2026. Credit conditions across the Sunbelt markets where the fund is concentrated remain stable, supported by sustained population inflows and above-average rent growth in multifamily and industrial sectors. Management continues to maintain disciplined underwriting standards with maximum LTVs of 70% on new originations, providing meaningful downside protection across the portfolio.",
    {x:0.8,y:3.0,w:11.7,h:2.2,fontSize:14,color:WHITE,fontFace:"Calibri Light",valign:"top"}
  );

  // Market context box
  slide.addShape(pres.ShapeType.roundRect,{x:1.0,y:5.4,w:11.3,h:1.4,fill:{color:NAVY_DARK},line:{color:MUTED,width:0.5},rectRadius:0.04});
  slide.addText("Market Context", {x:1.2,y:5.5,w:2.0,h:0.3,fontSize:9,bold:true,color:GOLD,fontFace:"Calibri"});
  slide.addText(
    "SOFR has stabilized in the 5.30–5.35% range, keeping floating-rate coupon income near cycle highs. Sunbelt multifamily fundamentals remain strong with vacancy rates below 5% in the fund's key markets. Industrial demand continues to benefit from e-commerce tailwinds and near-shoring activity.",
    {x:1.2,y:5.82,w:10.9,h:0.85,fontSize:9.5,color:"AABDD0",fontFace:"Calibri Light"}
  );

  // White footer on dark
  slide.addText("CONFIDENTIAL — FOR LIMITED PARTNERS ONLY", {x:0.35,y:7.18,w:5.5,h:0.2,fontSize:7,color:MUTED,fontFace:"Calibri Light"});
  slide.addText(`8 of ${TOTAL_SLIDES}`,{x:6.3,y:7.18,w:0.8,h:0.2,fontSize:7,color:MUTED,fontFace:"Calibri Light",align:"center"});
  slide.addText(`${FUND_SHORT}  |  ${PERIOD}`,{x:9.0,y:7.18,w:4.1,h:0.2,fontSize:7,color:MUTED,fontFace:"Calibri Light",align:"right"});
}

// ── SLIDE 9: APPENDIX ─────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color:WHITE };
  slide.addText("Appendix — Complete Data",{x:0.35,y:0.2,w:12.5,h:0.6,fontSize:26,bold:true,color:NAVY,fontFace:"Calibri"});
  slide.addShape(pres.ShapeType.rect,{x:0.35,y:0.82,w:12.6,h:0.022,fill:{color:GOLD}});

  const appData = [
    [{text:"Section / Metric",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:9}},{text:"Value",options:{bold:true,color:WHITE,fill:{color:NAVY},align:"right",fontFace:"Calibri",fontSize:9}}],
    [{text:"FUND OVERVIEW",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:8}},{text:"",options:{fill:{color:NAVY}}}],
    [{text:"Total Fund NAV",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:fmtM(cur.nav),options:{fill:{color:WHITE},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}}],
    [{text:"Gross Asset Value",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:fmtM(cur.gav),options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
    [{text:"Called / Uncalled Capital",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:`${fmtM(cur.called)} / ${fmtM(cur.uncalled)}`,options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
    [{text:"LOAN PORTFOLIO",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:8}},{text:"",options:{fill:{color:NAVY}}}],
    [{text:"Total Loan Book Balance",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:fmtM(cur.loan_book),options:{fill:{color:WHITE},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}}],
    [{text:"Active Loans / Wtd. Avg. LTV / DSCR",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:`14 / ${fmtPct(cur.wtd_ltv)} / ${fmtX(cur.wtd_dscr)}`,options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
    [{text:"Wtd. Avg. Coupon / Avg. Term",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:`${fmtPct(cur.wtd_coupon)} / ${cur.avg_term} months`,options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
    [{text:"DISTRIBUTIONS & RETURNS",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:8}},{text:"",options:{fill:{color:NAVY}}}],
    [{text:"Monthly Dist. / Annualized Yield",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:`$${cur.monthly_dist.toFixed(3)} / ${fmtPct(cur.ann_yield)}`,options:{fill:{color:WHITE},fontSize:9,bold:true,color:GOLD,align:"right",fontFace:"Calibri"}}],
    [{text:"Preferred Return Status / Cumulative Paid",options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:`CURRENT / ${fmtM(cur.cumulative_dist)}`,options:{fill:{color:ALT_ROW},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
    [{text:"Net IRR / MOIC / TVPI",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:`${fmtPct(cur.net_irr)} / ${fmtX(cur.moic)} / ${fmtX(cur.tvpi)}`,options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
    [{text:"DEAL ACTIVITY",options:{bold:true,color:WHITE,fill:{color:NAVY},fontFace:"Calibri",fontSize:8}},{text:"",options:{fill:{color:NAVY}}}],
    [{text:"New Originations / Repayments / Pipeline",options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",color:TEXT_DARK}},{text:`2 / $14.0M  |  1 / $9.2M  |  3 / $22.0M`,options:{fill:{color:WHITE},fontSize:9,fontFace:"Calibri Light",align:"right"}}],
  ];

  slide.addTable(appData,{x:0.35,y:1.05,w:12.6,colW:[7.5,4.8],border:{type:"solid",color:"DDDDDD",pt:0.3},rowH:0.36,fontFace:"Calibri"});

  slide.addText(
    `Data source: Test mock data (demo run). Figures are unaudited and for illustrative purposes only.  |  ${MANAGER}  |  ${CONTACT}`,
    {x:0.35,y:6.95,w:12.6,h:0.25,fontSize:7,italic:true,color:MUTED,fontFace:"Calibri Light"}
  );

  addFooter(slide, 9);
}

// ── SAVE ─────────────────────────────────────────────────
const OUT = "/sessions/zen-trusting-carson/mnt/outputs/Cana_Capital_Fund_I_April_2026_Investor_Deck.pptx";
pres.writeFile({fileName: OUT}).then(() => {
  console.log(`PPTX saved: ${OUT}`);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});

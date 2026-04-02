# PowerPoint Deck Template — Monthly Fund Statement (Investment-Banking Grade)

## Brand Customization (Collected at Onboarding)

Before generating any slide, read brand config from `../brand/brand.json`.
If missing, use defaults.

| Asset | Variable | Default |
|-------|----------|---------|
| Primary color | PRIMARY_COLOR | `1B2A4A` (navy) |
| Accent color | ACCENT_COLOR | `C9A84C` (gold) |
| Light background | BG_LIGHT | `F7F8FA` |
| Body text color | TEXT_DARK | `2C3E50` |
| Header font | FONT_HEADER | Calibri Bold |
| Body font | FONT_BODY | Calibri Light |
| Logo path | LOGO_PATH | None |
| Fund manager name | FUND_MANAGER | From config |
| Fund name | FUND_NAME | From config |

---

## Slide Dimensions
- 16:9 widescreen: 13.33" × 7.5" (pptxgenjs default)
- Margin zones: 0.4" from all edges as safe zone for content

## Footer (on ALL slides except cover)
```
Left: "CONFIDENTIAL — FOR LIMITED PARTNERS ONLY"
Center: [Slide number] of [Total]
Right: [FUND_NAME] | [MONTH YEAR]
Font: 7pt, color: 8899AA
y-position: 7.2" (near bottom)
```

---

## Slide-by-Slide Specification

---

### Slide 1 — Cover (Dark / Full Bleed)

```
Background: Full PRIMARY_COLOR fill (#1B2A4A)
No footer on cover slide.

LOGO (if provided):
  Position: top-left  x:0.4", y:0.3", maxWidth:1.8", maxHeight:0.6"
  Color mode: White/reversed version if available (set options.color = 'FFFFFF')

ACCENT horizontal rule (full width):
  y: 2.4", height: 3pt, color: ACCENT_COLOR

FUND NAME:
  x:0.4", y:2.65", w:9", fontSize:40, bold:true, color:'FFFFFF'
  fontFace: FONT_HEADER

FUND SERIES / STRATEGY LINE:
  x:0.4", y:3.3", w:9", fontSize:16, color:ACCENT_COLOR, italic:true
  e.g., "Senior Secured Commercial Real Estate Credit"

SPACER

REPORT TYPE + PERIOD:
  x:0.4", y:4.1", "Monthly Investor Report", fontSize:14, color:'CCDDEE'
  x:0.4", y:4.5", "[Month Year]", fontSize:28, bold:true, color:'FFFFFF'

BOTTOM ACCENT rule (full width):
  y: 6.7", 2pt, ACCENT_COLOR

CONFIDENTIALITY LINE:
  x:0.4", y:6.85", fontSize:8, color:'8899AA'
  "Confidential — Prepared for Limited Partners Only"
  RIGHT SIDE: "Prepared by [FUND_MANAGER] | [Date]"
```

---

### Slide 2 — Executive Summary (Light)

```
Background: BG_LIGHT (F7F8FA)

[LEFT ACCENT BAR]
  x:0, y:0, w:0.08", h:full height, fill:PRIMARY_COLOR

SLIDE TITLE:
  x:0.5", y:0.15", "Executive Summary", fontSize:26, bold, PRIMARY_COLOR

ACCENT RULE under title:
  x:0.5", y:0.6", w:12", h:2pt, ACCENT_COLOR

[STAT CALLOUT ROW — 3 boxes across top]
  y-range: 0.8" – 2.1"
  Box 1 (x:0.5"):  Total NAV
  Box 2 (x:4.6"):  Distribution Yield
  Box 3 (x:8.7"):  Active Loans

  Each box:
    Shape: ROUNDED_RECT, w:3.7", h:1.2", fill:WHITE, line:ACCENT_COLOR (1pt), lineSize:0.5
    Metric value: fontSize:32, bold, color:ACCENT_COLOR, centered, y:0.1" inside box
    Label: fontSize:9, TEXT_DARK, centered, y:0.65" inside box
    MoM delta: fontSize:9, GREEN or RED, centered, y:0.9" inside box
      e.g., "▲ +1.2% vs prior month"

[NARRATIVE — below stat row]
  x:0.5", y:2.25", w:7.5", fontSize:11, color:TEXT_DARK, lineSpacingMultiple:1.4
  [Executive Summary narrative from Step 4]

[RIGHT SIDEBAR — Fund Health Box]
  x:10.0", y:2.25", w:3.1", h:4.5"
  Background: PRIMARY_COLOR, rounded corners 5pt
  Title: "Fund Health" fontSize:11, bold, white, centered, top
  Items (each with icon indicator):
    ✓ Preferred Return: CURRENT
    ✓ Portfolio: Performing
    ✓ Capital: [X]% Deployed
    [Green checkmarks if healthy, red ⚠️ if flagged]
  Footer of box: "As of [DATE]" in 7pt, muted
```

---

### Slide 3 — NAV & Fund Size (Light)

```
Background: WHITE

SLIDE TITLE: "Fund Overview — [Month Year]", PRIMARY_COLOR, 26pt bold

[2×3 GRID OF STAT CALLOUT BOXES — fills content area]
  Layout: 2 rows × 3 boxes per row, with equal spacing

  Row 1:
    Box 1: Total Fund NAV
    Box 2: Gross Asset Value
    Box 3: Total Called Capital

  Row 2:
    Box 4: Uncalled / Dry Powder
    Box 5: Total Committed Capital
    Box 6: NAV per Unit (if applicable) OR Net IRR

  Each box: w:3.9", h:2.5"
    Top section (h:1.5"): ACCENT_COLOR number 52pt bold, centered
    Bottom section (h:1.0"): metric label 10pt TEXT_DARK, ± delta 9pt GREEN/RED
    Box style: no fill (white), thin PRIMARY_COLOR border (0.75pt), 4pt corner radius

  Grid spacing: 0.2" between boxes
```

---

### Slide 4 — Loan Portfolio Snapshot (Light)

```
Background: BG_LIGHT

SLIDE TITLE: "Loan Portfolio Snapshot", PRIMARY_COLOR

[LEFT PANEL — 60% width, x:0.4"]
  TABLE — 8 data rows + header row
  Columns: Metric (40%) | Current (20%) | Prior (20%) | Δ (20%)
  Header row: PRIMARY_COLOR fill, white text, Calibri Bold 9pt
  Alt rows: WHITE / F0F4F8
  Data: Calibri Light 9pt, right-align all values, left-align Metric
  Border: no outer border, 0.5pt horizontal rules between rows in E0E0E0
  Delta column: GREEN ▲ or RED ▼, 9pt bold
  ⚠️ symbol on flagged rows (prepended to delta value)

  Rows:
    Total Loan Book Balance
    Number of Active Loans
    Wtd. Avg. LTV
    Wtd. Avg. DSCR
    Wtd. Avg. Coupon
    Avg. Remaining Term (mo.)
    Portfolio Occupancy
    Performing Loans (%)

[RIGHT PANEL — 37% width, x:8.3"]
  "Property Type Mix"    [label, 10pt bold PRIMARY]
  [Horizontal bar chart]
    pptxgen.ChartType.bar, horizontal orientation
    Categories: property types (Multifamily, Industrial, Office, Retail, Other)
    Values: % allocation
    Bar color: cycle through [PRIMARY, ACCENT, 3D6B9E, 5A8FA8, 8BADBF]
    Data labels: outside end, 9pt bold
    No gridlines visible; axis labels left side only
    Plot area: transparent background
    Legend: none (labels serve as legend)
```

---

### Slide 5 — Distributions & Returns (Split — Dark Left / Light Right)

```
LEFT HALF (0" to 6.5"): PRIMARY_COLOR background
RIGHT HALF (6.5" to 13.33"): WHITE background

[LEFT SIDE]
  SLIDE TITLE: "Distributions & Returns", WHITE, 26pt bold, x:0.4", y:0.2"
  ACCENT RULE: y:0.75", w:6", PRIMARY_COLOR side

  STAT STACK — 4 boxes, stacked vertically
    Box 1: Monthly Distribution per Unit
      $X.XX            32pt ACCENT_COLOR
      per unit/share   9pt WHITE, muted
    Box 2: Annualized Distribution Yield
      X.X%             32pt ACCENT_COLOR
    Box 3: Preferred Return Rate & Status
      X.X% — CURRENT ✓  [GREEN badge if current, RED if behind]
    Box 4: Cumulative Distributions Paid
      $X.X M to date   24pt ACCENT_COLOR

[RIGHT SIDE]
  "Returns Summary"      [10pt bold PRIMARY_COLOR, y:0.8"]
  TABLE — Net vs Gross returns
    Header: light PRIMARY fill, white text
    Net IRR / Gross IRR
    MOIC (net / gross)
    TVPI (net)
    Preferred Hurdle
    [Note: 7pt MUTED "Returns are unaudited."]

  "Distribution History"   [10pt bold PRIMARY_COLOR, below table]
  TABLE — Last 6 months
    Period | $/Unit | Yield%
    [Light alternating rows, right-align numbers]
    Most recent period bolded

  NARRATIVE:
    [Distribution commentary from Step 4, 10pt, Calibri Light, TEXT_DARK]
```

---

### Slide 6 — Deal Activity (Light)

```
Background: WHITE

SLIDE TITLE: "Deal Activity — [Month Year]", PRIMARY_COLOR

[TOP STAT ROW — 3 boxes]
  Box: New Originations    [PRIMARY fill, white text]
    X loans / $X.X M committed
  Box: Repayments / Exits  [3D6B9E fill, white text]
    X loans / $X.X M returned
  Box: Pipeline             [ACCENT fill, PRIMARY text]
    X loans / $X.X M in process

  Each box: w:4.1", h:1.3", rounded corners, large stat + smaller label

[NEW DEALS TABLE — if originations exist]
  Full width, below stat row, y:2.0"
  Columns: Property | Market | Loan $ | LTV | Coupon | Term | Type
  Header: PRIMARY fill, white
  Alt rows, right-align numbers
  If > 5 deals, truncate and add "See appendix for full list"

[NARRATIVE — Deal Activity commentary]
  x:0.4", below table or at y:5.0", fontSize:10, Calibri Light

[REPAYMENTS TABLE — if exits exist, right column]
  "Repayments / Exits" header in 3D6B9E
  Property | Loan $ | Payoff $ | Hold Period
```

---

### Slide 7 — Portfolio Composition (Light)

```
Background: BG_LIGHT

SLIDE TITLE: "Portfolio Composition", PRIMARY_COLOR

[LEFT — By Property Type, x:0.4", w:5.5"]
  "By Property Type"        [section label, 10pt bold PRIMARY]
  [Horizontal bar chart — same as Slide 4 right panel but larger]
  Include: % on label AND $ amount on second line of label
  Sort: largest to smallest

[CENTER DIVIDER — vertical rule, 0.5pt ACCENT_COLOR, x:6.2"]

[RIGHT — By Geography, x:6.5", w:6.4"]
  "By Geography (MSA / State)"   [section label, 10pt bold]
  TABLE:
    Market | # Loans | Deployed $ | % of Portfolio
    Sort by $ descending
    Top 5-8 markets; remaining as "Other"
    Total row: bold, PRIMARY_COLOR text

[BOTTOM ROW — Loan Size Concentration]
  y:5.8", full width
  ACCENT_COLOR rule above
  "Loan Size Concentration"    [10pt bold]
  [4 boxes across: <$5M | $5M–$10M | $10M–$20M | >$20M]
    Each box: count + total $ + % of portfolio
```

---

### Slide 8 — Outlook (Dark / Accent)

```
Background: PRIMARY_COLOR (full bleed)

SLIDE TITLE: "Outlook", WHITE, 32pt bold, x:0.4", y:0.2"
ACCENT RULE: y:0.8", w:12", 2pt ACCENT_COLOR

[PULL QUOTE BOX — centered, y:1.3"]
  w:11", h:1.5"
  Left border: 5pt ACCENT_COLOR
  Background: rgba of PRIMARY_COLOR + 20% lighter (approx #2C3F5C)
  Text: Key forward stat or headline
  fontSize:20, bold, WHITE, left-aligned with 0.3" left padding
  e.g., "$X.X M in pipeline loans expected to close in [next month]"

[MAIN NARRATIVE — y:3.0"]
  Outlook commentary from Step 4
  fontSize:14, Calibri Light, WHITE, lineSpacingMultiple:1.5
  x:1.0", w:11.3"

[MARKET CONTEXT BOX — y:5.2"]
  Background: 10% lighter navy (#2C3F5C-ish, use hex '243454')
  1pt MUTED border
  "Market Context" label: ACCENT, 9pt bold
  [1–2 lines of macro/market commentary if available in source data]
  fontSize:10, WHITE

[BOTTOM: NO FOOTER DARK OVERLAY — use white footer text on dark bg for this slide]
```

---

### Slide 9 — Appendix (Light)

```
Background: WHITE

SLIDE TITLE: "Appendix — Complete Data", PRIMARY_COLOR

[FULL-WIDTH TABLE — grouped by section]
  Column 1 (35%): Field name
  Column 2 (65%): Value

  Section headers (SPANNING full row):
    "FUND OVERVIEW"           [PRIMARY fill, white, 9pt bold]
    "LOAN PORTFOLIO"          [PRIMARY fill, white, 9pt bold]
    "DISTRIBUTIONS & RETURNS" [PRIMARY fill, white, 9pt bold]
    "DEAL ACTIVITY"           [PRIMARY fill, white, 9pt bold]

  Data rows: alternating white / BG_LIGHT, 9pt Calibri Light
  Right-align all values; bold any ⚠️ flagged items

[BOTTOM FOOTNOTES]
  y:6.9"
  "Data sourced from: [FILENAME], [FILE DATE]."
  "Figures unaudited. Prepared by [FUND_MANAGER]. All rights reserved."
  fontSize:7.5, MUTED, italic

[CONTACT BLOCK — right side, y:6.0"]
  [FUND_MANAGER]
  [Email] | [Website] | [Phone]
  fontSize:8, TEXT_DARK
```

---

## pptxgenjs Implementation Notes

### Install
```bash
npm install -g pptxgenjs
```

### Key Patterns

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

// Set slide dimensions (widescreen)
pres.layout = "LAYOUT_WIDE"; // 13.33" x 7.5"

// Global footer helper
function addFooter(slide, slideNum, total, fundName, period) {
  slide.addText(`CONFIDENTIAL — FOR LIMITED PARTNERS ONLY`, {
    x:0.4, y:7.15, w:5, h:0.2, fontSize:7, color:'8899AA'
  });
  slide.addText(`${slideNum} of ${total}`, {
    x:6.3, y:7.15, w:0.8, h:0.2, fontSize:7, color:'8899AA', align:'center'
  });
  slide.addText(`${fundName} | ${period}`, {
    x:9, y:7.15, w:4.1, h:0.2, fontSize:7, color:'8899AA', align:'right'
  });
}

// Stat callout box helper
function addStatBox(slide, x, y, value, label, delta, deltaPositive) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w:3.7, h:1.2,
    fill:{color:'FFFFFF'}, line:{color:'C9A84C', width:0.5},
    rectRadius:0.05
  });
  slide.addText(value, {x:x+0.1, y:y+0.05, w:3.5, h:0.7,
    fontSize:32, bold:true, color:'C9A84C', align:'center'});
  slide.addText(label, {x:x+0.1, y:y+0.73, w:3.5, h:0.2,
    fontSize:9, color:'2C3E50', align:'center'});
  const arrowChar = deltaPositive ? '▲' : '▼';
  const deltaColor = deltaPositive ? '1A6B3A' : 'B71C1C';
  slide.addText(`${arrowChar} ${delta}`, {x:x+0.1, y:y+0.94, w:3.5, h:0.2,
    fontSize:9, color:deltaColor, align:'center'});
}

// Table helper — standard IB table
function addTable(slide, headers, rows, x, y, w) {
  const tableData = [
    headers.map(h => ({text: h, options: {
      fill:{color:'1B2A4A'}, color:'FFFFFF', bold:true, fontSize:9
    }})),
    ...rows.map((row, i) => row.map(cell => ({text: String(cell), options: {
      fill:{color: i % 2 === 0 ? 'FFFFFF' : 'F0F4F8'},
      color:'2C3E50', fontSize:9
    }})))
  ];
  slide.addTable(tableData, {x, y, w, border:{type:'none'},
    rowH:0.28, fontFace:'Calibri'});
}

// Logo placement helper
function addLogo(slide, logoPath, x, y, maxW, maxH) {
  if (!logoPath || !fs.existsSync(logoPath)) return;
  slide.addImage({path:logoPath, x, y, w:maxW, h:maxH, sizing:{type:'contain'}});
}
```

### QA Checklist (run before declaring done)
```bash
# Convert to PDF for visual inspection
python scripts/office/soffice.py --headless --convert-to pdf deck.pptx
pdftoppm -jpeg -r 150 deck.pdf slide

# Check each slide image for:
# - Text overflow or clipping at edges
# - Overlapping elements
# - Placeholder text not replaced
# - Footer present on all non-cover slides
# - Logo correctly positioned
# - No blank/empty slides
```

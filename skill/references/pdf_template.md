# PDF Report Template — Monthly Fund Statement (Investment-Banking Grade)

## Brand Customization (Collected at Onboarding)

Before rendering any page, verify these brand assets are available. Read them from
`../brand/` folder if provided. If missing, use defaults below.

| Asset | Source | Default |
|-------|--------|---------|
| `brand.json` | Onboarding-provided config | Navy + Gold palette |
| `logo.png` | Fund manager logo | Text fallback |
| `PRIMARY_COLOR` | `brand.json` → primary | `#1B2A4A` (navy) |
| `ACCENT_COLOR` | `brand.json` → accent | `#C9A84C` (gold) |
| `FONT_BODY` | `brand.json` → font | Helvetica |
| `FUND_MANAGER` | `brand.json` → manager | From config |
| `DISCLAIMER` | `brand.json` → disclaimer | Standard legal text below |

---

## Standard Legal Disclaimer Text

Place on Page 2 (or inside cover page footer, 7pt):

> This report is confidential and prepared solely for the addressee(s). It is not to be copied, distributed, or used by any other person or for any other purpose without the prior written consent of [FUND_MANAGER]. This material does not constitute an offer to sell or solicitation of an offer to buy any securities. Past performance is not indicative of future results. Investing in private credit funds involves substantial risk, including the potential loss of principal. Forward-looking statements reflect management's current expectations and are subject to uncertainty. Figures presented are unaudited and subject to change. Returns are net of all fees and expenses unless otherwise noted.

---

## Design Specification

### Colors
| Name | Default Hex | Variable | Use |
|------|------------|----------|-----|
| Primary | `#1B2A4A` | PRIMARY_COLOR | Cover bg, header bars, table headers |
| Accent | `#C9A84C` | ACCENT_COLOR | Key numbers, rules, callout borders |
| Background | `#F7F8FA` | BG_LIGHT | Page tint, alt table rows |
| Body Text | `#2C3E50` | TEXT_DARK | All body copy |
| Positive | `#1A6B3A` | GREEN | MoM improvements, current preferred return |
| Negative | `#B71C1C` | RED | MoM deterioration, preferred return behind |
| Muted | `#8899AA` | TEXT_MUTED | Footers, captions, disclaimers |
| White | `#FFFFFF` | — | Text on dark backgrounds |

### Typography
| Use | Font | Size | Weight |
|-----|------|------|--------|
| Cover fund name | Helvetica-Bold | 42pt | Bold |
| Cover subtitle | Helvetica | 18pt | Regular |
| Section header bars | Helvetica-Bold | 13pt | Bold |
| Page title | Helvetica-Bold | 16pt | Bold |
| Body | Helvetica | 10pt | Regular |
| Table header | Helvetica-Bold | 9pt | Bold |
| Table body | Helvetica | 9pt | Regular |
| Stat callout number | Helvetica-Bold | 28pt | Bold |
| Stat callout label | Helvetica | 9pt | Regular |
| Footer | Helvetica | 7pt | Regular |
| Disclaimer | Helvetica | 7.5pt | Regular |

### Page Layout
- Paper: US Letter, 8.5" × 11"
- Margins: 0.75" all sides (54pt)
- Header bar: Horizontal band at top of each interior page: 22pt tall, PRIMARY_COLOR background
  - Left: Fund name in white, 9pt bold
  - Right: Month Year in white, 9pt regular
- Footer: 7pt, MUTED color, three columns:
  - Left: "CONFIDENTIAL — FOR LIMITED PARTNERS ONLY"
  - Center: Page number (e.g., "3 of 9")
  - Right: Report date
- Minimum margins from footer to content: 14pt
- Horizontal rule after header on interior pages: 0.5pt, ACCENT_COLOR

---

## Page-by-Page Specification

### Page 1 — Cover Page

Layout: Full PRIMARY_COLOR background. All text white or ACCENT_COLOR.

```
[TOP BAND — 20% of page height, PRIMARY_COLOR]
  [Logo — if provided, top-left, max 1.5" wide, white or color version]
  [ACCENT_COLOR horizontal rule, 2pt, full width, at 18%]

[CENTER BLOCK — vertically centered in middle 60% of page]
  [FUND_NAME]                  42pt, Helvetica-Bold, WHITE, left-aligned
  [FUND_VINTAGE / SERIES]      16pt, Helvetica, ACCENT_COLOR, left-aligned
                               e.g., "Senior Secured Real Estate Credit — Series I"
  [Spacer 12pt]
  [ACCENT_COLOR rule, 1pt, 40% of page width]
  [Spacer 12pt]
  Monthly Investor Report      14pt, Helvetica, WHITE, left-aligned
  [MONTH YEAR]                 24pt, Helvetica-Bold, WHITE, left-aligned

[BOTTOM BAND — 15% of page height]
  [ACCENT_COLOR rule, 1pt, full width]
  CONFIDENTIAL — FOR LIMITED PARTNERS ONLY    9pt, MUTED, left-aligned
  Prepared by [FUND_MANAGER]                  9pt, MUTED, left-aligned
  [Date of preparation]                       9pt, MUTED, right-aligned
```

No page number on cover. No header bar on cover.

---

### Page 2 — Legal Disclaimer & Contents

```
[Header bar]

[LEFT COLUMN — 70%]
  "Important Disclosures"     Heading, 14pt bold, PRIMARY_COLOR
  [DISCLAIMER text in 7.5pt grey, flowing paragraph]

[RIGHT COLUMN — 25%, offset 5%]
  "Contents"                  Heading, 11pt bold
  [Simple list — section names and page numbers, 9pt]
    Executive Summary ............. 3
    Portfolio Snapshot ............. 4
    Distributions & Returns ......... 5
    Deal Activity .................. 6
    Portfolio Composition ........... 7
    Outlook ........................ 8
    Appendix ....................... 9
```

---

### Page 3 — Executive Summary

```
[Header bar]
[Page title: "Executive Summary" in PRIMARY_COLOR, 16pt bold, left]
[ACCENT_COLOR horizontal rule, 0.5pt]

[THREE-COLUMN STAT CALLOUT ROW — top of content area]
  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────┐
  │  Total NAV       │  │ Distribution     │  │  Active Loans    │
  │  $X.X M          │  │ Yield X.X%       │  │  X               │
  │  ▲ X.X% MoM      │  │ ▲ X bps MoM      │  │  ▲ X MoM        │
  └──────────────────┘  └──────────────────┘  └──────────────────┘
  [Box: thin ACCENT border, 3pt corner radius, ACCENT callout number 28pt]
  [MoM delta: GREEN or RED depending on direction, 8pt with arrow symbol]

[TWO-COLUMN BODY — below stat row]
  LEFT (60%): Executive Summary narrative (3-4 paragraphs, 10pt)
  RIGHT (40%):
    [Preferred Return Status box]
      Status: CURRENT ✓   [GREEN background, white text]  — or —
      Status: BEHIND ⚠️  [RED, white text]
    [Horizontal data list:]
      Gross Asset Value:     $X.X M
      Called Capital:        $X.X M  (X.X% of committed)
      Uncalled / Dry Powder: $X.X M
      Net IRR (since incep): X.X%
```

---

### Page 4 — Loan Portfolio Snapshot

```
[Header bar]
[Page title: "Loan Portfolio Snapshot"]

[METRICS TABLE — full width]
Column widths: 40% / 20% / 20% / 20%
Header row: PRIMARY_COLOR bg, white text, 9pt bold
  Metric | Current Month | Prior Month | Change

Rows (alternating white / BG_LIGHT):
  Total Loan Book Balance
  Number of Active Loans
  Wtd. Avg. Loan-to-Value (LTV)
  Wtd. Avg. DSCR
  Wtd. Avg. Coupon / Interest Rate
  Avg. Remaining Term (months)
  Portfolio Occupancy Rate
  Performing Loans (count / %)

[Delta column formatting:]
  Positive: GREEN text, ▲ prefix   e.g., "▲ +2.3%"
  Negative: RED text, ▼ prefix     e.g., "▼ -4.1%"
  Flagged:  ⚠️ symbol prepended    e.g., "⚠️ ▼ -8.2%"
  Flat:     MUTED text, "—"

[SECTION DIVIDER in ACCENT_COLOR, 0.5pt]

[TWO-COLUMN BREAKDOWN — below table]
  LEFT: Property Type Breakdown (horizontal bar chart, bars in PRIMARY/ACCENT/MED_BLUE rotation)
  RIGHT: Geographic Breakdown (table: Market | # Loans | $ | % of Portfolio)
```

---

### Page 5 — Distributions & Returns

```
[Header bar]
[Page title: "Distributions & Returns"]

[FOUR-STAT CALLOUT ROW — full width]
  Monthly Distribution  |  Annualized Yield  |  Preferred Return  |  Cumulative Distributions
  $X.XX / unit          |  X.X%              |  X.X% (CURRENT)    |  $X.X M to date
  [ACCENT numbers, 28pt; labels 8pt MUTED]
  [Preferred Return: box bg GREEN if Current, RED if Behind]

[RETURNS TABLE — top-right quadrant]
  Metric | Net | Gross
  Net IRR (since inception) | X.X% | X.X%
  Cash-on-Cash Multiple (MOIC) | X.Xx | X.Xx
  Total Value to Paid-In (TVPI) | X.Xx | —
  Preferred Return Hurdle | X.X% | —
  [Note: "Returns are unaudited. Net returns are after all fees and expenses."]

[NARRATIVE — bottom-left, 10pt, 60% width]
  [Distribution & Returns commentary from Step 4]

[DISTRIBUTION HISTORY TABLE — bottom-right, 35% width]
  Period | Distribution | Yield
  [Last 6 months of data if available]
  [Alternating row shading, right-align numbers]
```

---

### Page 6 — Deal Activity

```
[Header bar]
[Page title: "Deal Activity — [MONTH YEAR]"]

[THREE-BOX SUMMARY ROW]
  ┌─────────────────────┐   ┌─────────────────────┐   ┌─────────────────────┐
  │ New Originations    │   │ Repayments / Exits   │   │ Pipeline            │
  │ X loans / $X.X M   │   │ X loans / $X.X M     │   │ X loans / $X.X M   │
  └─────────────────────┘   └─────────────────────┘   └─────────────────────┘
  [PRIMARY boxes with white text; numbers in ACCENT]

[NEW DEAL TABLE — if any new originations:]
  Property Name | Market | Loan Amt | LTV | Coupon | Term | Property Type
  [Full-width table, header PRIMARY, alternating rows]
  [Right-align all financial figures]

[REPAYMENT TABLE — if any payoffs:]
  Property Name | Original Balance | Payoff Amount | Hold Period

[NARRATIVE — Deal Activity commentary, 10pt]

[If no new originations: italic muted text "No new originations this period."]
```

---

### Page 7 — Portfolio Composition

```
[Header bar]
[Page title: "Portfolio Composition"]

[TWO-COLUMN LAYOUT]

LEFT (48%):
  "By Property Type"           [section label, PRIMARY, 10pt bold]
  [Horizontal bar chart]
    Bars ordered largest → smallest
    Bar color: PRIMARY, ACCENT, MED_BLUE, SLATE, cycling
    Value labels at end of each bar (% and $ amount)
  [Summary below: "Top concentration: [X]% in [Type]"]

RIGHT (48%):
  "By Geography (MSA / State)"  [section label]
  [Table:]
    Market | # Loans | Deployed $ | % of Portfolio
    [Sorted by $ descending]
    [No outer border; horizontal dividers only, 0.5pt MUTED]
    [Total row at bottom: bold]

[BOTTOM ROW — Loan Size Distribution]
  "Loan Size Distribution"     [section label]
  [Simple histogram or table: Size Bucket | Count | % of Loans | Total $]
  e.g.,   <$5M | $5-$10M | $10-$20M | >$20M
```

---

### Page 8 — Outlook

```
[Header bar]
[Page title: "Outlook"]

[PULL QUOTE BOX — centered, 60% width]
  [ACCENT border, 2pt left side]
  [Key forward-looking stat or quote, 14pt Helvetica-Bold, PRIMARY_COLOR]
  e.g., "$X.X M in pipeline expected to close in [next month]"

[NARRATIVE — Outlook commentary, 11pt, generous leading (14pt)]

[MARKET CONTEXT BOX — if data available]
  [BG_LIGHT background, 1pt MUTED border]
  "Market Context" label in PRIMARY, 10pt bold
  [1-2 sentences on relevant macro/CRE market conditions]
```

---

### Page 9 — Appendix

```
[Header bar]
[Page title: "Appendix — Complete Data Extract"]

[FULL DATA TABLE — two-column: Field | Value]
  [List ALL extracted metrics from Step 2 in grouped sections:]

  [FUND OVERVIEW]
    Fund NAV (current)
    Fund NAV (prior month)
    NAV Change ($)
    NAV Change (%)
    Gross Asset Value
    ...

  [LOAN PORTFOLIO]
    Total Loan Book
    Active Loan Count
    ...

  [DISTRIBUTIONS & RETURNS]
    Monthly Distribution Amount
    ...

  [DEAL ACTIVITY]
    New Originations (count)
    ...

[FOOTNOTE:]
  "Data sourced from: [FILENAME], dated [FILE DATE]."
  "Figures are unaudited and subject to revision by the fund administrator."
  "This report was prepared by automated extraction and reviewed by [FUND_MANAGER]."

[CONTACT BLOCK]
  [FUND_MANAGER name, address, website, email]
```

---

## Python / ReportLab Implementation Notes

### Install
```
pip install reportlab pdfplumber --break-system-packages
```

### Key Patterns

**Cover page** — use `canvas.Canvas` for full-bleed background rect, positioned text
```python
c = canvas.Canvas("report.pdf", pagesize=letter)
c.setFillColor(HexColor('#1B2A4A'))
c.rect(0, 0, letter[0], letter[1], fill=1, stroke=0)
# Logo
if logo_path:
    c.drawImage(logo_path, 54, letter[1]-90, width=108, preserveAspectRatio=True, mask='auto')
```

**Header bar** — at top of every interior page via `onFirstPage` / `onLaterPages` callbacks:
```python
def header(canvas, doc):
    canvas.setFillColor(HexColor('#1B2A4A'))
    canvas.rect(0, letter[1]-36, letter[0], 36, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawString(54, letter[1]-22, FUND_NAME)
    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(letter[0]-54, letter[1]-22, report_period)
```

**Stat callout boxes** — use `Table` with custom `TableStyle`:
```python
Table([["$12.4M", "NAV"]], colWidths=[120], rowHeights=[50, 20],
      style=[('TEXTCOLOR', (0,0), (-1,-1), HexColor('#C9A84C')),
             ('FONTSIZE', (0,0), (0,0), 28),
             ('BOX', (0,0), (-1,-1), 1, HexColor('#C9A84C')),
             ('ALIGN', (0,0), (-1,-1), 'CENTER')])
```

**Delta formatting**:
```python
def format_delta(value, is_positive_good=True):
    arrow = "▲" if value > 0 else "▼"
    color = GREEN if (value > 0) == is_positive_good else RED
    return f"{arrow} {abs(value):.1f}%", color
```

**Always embed fonts** — avoid font substitution issues in distributed PDFs:
```python
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# If using custom fonts from brand.json, register here
# pdfmetrics.registerFont(TTFont('CustomFont', 'path/to/font.ttf'))
```

**Final validation**:
```python
from pypdf import PdfReader
r = PdfReader("report.pdf")
assert len(r.pages) >= 7, "PDF page count too low — check layout"
print(f"PDF generated: {len(r.pages)} pages")
```

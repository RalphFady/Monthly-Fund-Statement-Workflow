---
name: monthly-fund-statement
description: >
  Generate a monthly investor update for a commercial real estate private credit fund.
  Use this skill whenever someone asks to: run the monthly fund report, generate investor
  statements, process end-of-month fund data, create a fund summary, produce the monthly
  investor update, or compare this month's fund performance to last month.
  This skill finds the latest raw statement file in Google Drive, extracts NAV, loan
  portfolio metrics, distributions & returns, and new deal/pipeline data, compares to the
  prior month, writes a narrative commentary, and outputs a polished PDF + PowerPoint deck
  ready for investor distribution.
---

# Monthly Fund Statement — Commercial Real Estate Private Credit

## What This Skill Does

Each month-end, this skill:
1. **Locates** the new raw statement file in the configured Google Drive folder
2. **Extracts** key fund metrics (NAV, loan portfolio, distributions, pipeline)
3. **Builds a Data Availability Map** — determines which sections actually have data
4. **Compares** to the prior month's data to compute deltas and trends
5. **Writes** a professional narrative commentary
6. **Produces** two investor-ready outputs: a PDF report and a PowerPoint deck — **only including sections where real data was found**

---

## Configuration (Set Once During Onboarding)

Before the FIRST run, collect all brand and fund configuration from the user.
Save everything to `brand/brand.json` in the skill folder. Future runs load from this file — don't ask again unless the user wants to update something.

### Step 0 — Check Google Drive Connection (Always First)

Before asking any configuration questions, verify Google Drive is connected:

1. Search the MCP registry for Google Drive and call `suggest_connectors` to surface the Connect button in chat.
2. Say to the user: *"Before we set up your fund, I need Google Drive connected — that's where I'll find your monthly statements. Please click Connect above and authorize access. Once you're done, let me know and we'll continue."*
3. Wait for the user to confirm they've connected it.
4. If they say Google Drive is already connected, proceed immediately.
5. Do not proceed to the configuration questions until this step is complete.

### Onboarding Checklist — Ask the User for Each Item

**1. Fund Identity**
- Full legal fund name (e.g., "Acme Capital Senior Credit Fund II, L.P.")
- Fund strategy / series label (e.g., "Senior Secured CRE Debt — Series II")
- Fund manager / GP name (e.g., "Acme Capital Management LLC")
- Fund vintage year (e.g., 2023)
- Reporting currency (typically USD)

**2. Google Drive**
- Name of the Google Drive folder where raw monthly statements are uploaded (e.g., "Monthly Statements")
- File naming pattern to expect (e.g., "Jan 2026 Statement.pdf", "2026-01 Fund Report.xlsx")
- Any password on the files? If so, how is it shared with you?

**3. Brand Colors**
Ask: "What are your brand colors? If you know the hex codes, share them. Otherwise describe the colors (e.g., 'dark navy and gold') and I'll use professional defaults."
- Primary color (hex or description) → used for header bars, covers, table headers
- Accent color (hex or description) → used for callout numbers, rules
- If unknown: use `#1B2A4A` (navy) and `#C9A84C` (gold) as institutional defaults

**4. Logo**
Ask: "Please upload your fund logo (PNG or SVG preferred, white or dark version both helpful). This goes on every page header and the cover slide."
- Save uploaded logo to `brand/logo.png`
- If not provided, use text-only fallback (fund manager name as text)

**5. Fonts**
Ask: "Do you have specific fonts you use in your fund materials? (e.g., Gotham, Avenir, Garamond). If not, we'll use Calibri / Helvetica which are professional standards."
- If custom fonts are provided, ask user to upload .ttf or .otf files → save to `brand/fonts/`
- Note: Custom fonts only work if the font files are available in the environment

**6. Investor Distribution**
- Who receives this report? (List or description)
- Is there a standard email template or subject line to use?
- Any specific disclaimers required by your legal team beyond the standard?

**7. Report Preferences**
- Include a cover letter page? (Yes/No — common for LP distribution)
- Include a loan-by-loan schedule in the appendix? (Yes if # loans < 20)
- Include performance benchmarks / index comparisons? If so, which?
- Preferred date format (e.g., "March 2026" vs "Mar-26" vs "2026-03")

### Save Configuration

After collecting answers, write `brand/brand.json`:
```json
{
  "fund_name": "...",
  "fund_strategy": "...",
  "fund_manager": "...",
  "vintage_year": "...",
  "currency": "USD",
  "gdrive_folder": "...",
  "file_pattern": "...",
  "primary_color": "#1B2A4A",
  "accent_color": "#C9A84C",
  "bg_light": "#F7F8FA",
  "text_dark": "#2C3E50",
  "font_header": "Calibri",
  "font_body": "Calibri Light",
  "logo_path": "brand/logo.png",
  "date_format": "Month YYYY",
  "include_cover_letter": false,
  "include_loan_schedule": true,
  "custom_disclaimer": "",
  "contact_email": "",
  "contact_website": ""
}
```

If any setting is missing, ask the user before proceeding — don't guess.

---

## Step-by-Step Workflow

### Step 1 — Find the New Statement in Google Drive

Use the Google Drive MCP to search for the most recently modified file in the configured folder.

```
Search Google Drive for files in folder: [GDRIVE_FOLDER_NAME]
Sort by: last modified, descending
Take the most recent file as: CURRENT_MONTH_FILE
Take the second most recent as: PRIOR_MONTH_FILE
```

If you cannot access Google Drive or find no files, stop and tell the user:
> "I couldn't find any files in the '[GDRIVE_FOLDER_NAME]' folder on Google Drive. Please make sure Google Drive is connected (see onboarding guide) and that the folder name is correct."

If there's only one file and no prior month exists, proceed with current month only and skip all comparison sections — note this in the report.

### Step 2 — Extract Key Data

Read both files (current and prior month). Extract the following fields. If a field is missing or unclear, note it as `null` (not "N/A" — use null so the DAM can detect it clearly).

#### Section A: NAV & Fund Size
- Total Fund NAV (net asset value)
- Gross Asset Value
- Total Committed Capital
- Total Called Capital
- Uncalled Capital / Dry Powder
- Fund NAV per unit/share (if applicable)

#### Section B: Loan Portfolio Metrics
- Number of active loans
- Total loan book balance (outstanding principal)
- Weighted Average Loan-to-Value (LTV)
- Weighted Average Debt Service Coverage Ratio (DSCR)
- Weighted Average Coupon / Interest Rate
- Average loan term remaining
- Portfolio occupancy (if disclosed)
- Geographic breakdown (by state or market) — as a list of {location, %, $}
- Property type breakdown (office, multifamily, industrial, retail, etc.) — as a list of {type, %, $}
- Individual loan schedule (loan name/ID, property, balance, maturity, LTV) — as a list of rows

#### Section C: Distributions & Returns
- Monthly distribution amount (per unit/share or total)
- Annualized distribution yield
- Preferred return rate and whether it is current
- Cumulative distributions paid to date
- Net IRR (if reported)
- Total Value to Paid-In (TVPI, if reported)

#### Section D: New Deals & Pipeline
- New loans originated this month (count, total $, avg size, list of deals if named)
- Loans repaid/matured this month (count, total $)
- Loans in pipeline (count, total $)
- Any notable deal highlights or property descriptions

### Step 3 — Build the Data Availability Map (DAM)

After extraction, construct a Data Availability Map. **This governs which sections are rendered in Steps 5 and 6 — never render a section whose DAM flag is `false`.**

```json
{
  "nav_fund_size": false,
  "loan_portfolio_metrics": false,
  "loan_portfolio_geo": false,
  "loan_portfolio_property_type": false,
  "loan_schedule": false,
  "distributions_returns": false,
  "deal_activity_new": false,
  "deal_activity_repaid": false,
  "deal_pipeline": false,
  "prior_month_comparison": false
}
```

Set each flag to `true` only if **at least one real numeric or text value** was extracted — not null, not an empty string.

**DAM rules:**
- `nav_fund_size`: true if any of NAV, GAV, Committed Capital, or Called Capital is non-null
- `loan_portfolio_metrics`: true if at least 3 of (# loans, loan book balance, LTV, DSCR, coupon, term remaining) are non-null
- `loan_portfolio_geo`: true if geographic breakdown list has ≥ 2 rows
- `loan_portfolio_property_type`: true if property type breakdown list has ≥ 2 rows
- `loan_schedule`: true if individual loan rows exist AND `include_loan_schedule` is true in brand.json
- `distributions_returns`: true if at least one of (distribution amount, yield, IRR, TVPI) is non-null
- `deal_activity_new`: true if new origination count or total $ is non-null
- `deal_activity_repaid`: true if repayment count or total $ is non-null
- `deal_pipeline`: true if pipeline count or total $ is non-null
- `prior_month_comparison`: true if prior month file loaded AND has ≥ 1 overlapping non-null metric

**Print the DAM to console before building outputs:**
```
DATA AVAILABILITY MAP:
  nav_fund_size                ✅ true
  loan_portfolio_metrics       ✅ true
  loan_portfolio_geo           ⬜ false  (no geographic data found in source)
  loan_portfolio_property_type ✅ true
  loan_schedule                ⬜ false  (include_loan_schedule = false in config)
  distributions_returns        ✅ true
  deal_activity_new            ✅ true
  deal_activity_repaid         ⬜ false  (not found in source)
  deal_pipeline                ✅ true
  prior_month_comparison       ✅ true

→ Building PDF: 7 sections | Building PPTX: 7 slides + cover + appendix
```

### Step 4 — Compute Month-over-Month Deltas

Only run this step if `dam.prior_month_comparison` is `true`.

For each numeric metric, compute:
- Absolute change: Current − Prior
- Percentage change: (Current − Prior) / Prior × 100%

Flag significant changes (>5% swing in NAV, loan book, or yield) with a ⚠️ symbol so the human reviewer notices.

If the prior month file was unavailable, skip this step.

### Step 5 — Write Narrative Commentary

Write a clear, professional narrative **only for sections where the DAM flag is `true`**. Omit sections entirely from the commentary if their DAM flag is `false` — do not mention "data was unavailable" in the investor-facing narrative (that is only for internal notes).

The tone should be:
- **Confident and factual** — no hedging or filler language
- **Investor-appropriate** — assume sophisticated readers, not retail
- **Concise** — each section 2–4 sentences; the whole commentary < 600 words
- **Forward-looking where data supports it** — mention pipeline, upcoming maturities, or market conditions if present in the source

Use this structure for the commentary (include only DAM-enabled sections):

```
EXECUTIVE SUMMARY (3–4 sentences)
  High-level: NAV movement (if dam.nav_fund_size), distribution status (if dam.distributions_returns), portfolio health (if dam.loan_portfolio_metrics)

PORTFOLIO UPDATE (2–4 sentences) — include only if dam.loan_portfolio_metrics = true
  Active loans, key metrics, any notable additions or payoffs

DISTRIBUTIONS & RETURNS (2–3 sentences) — include only if dam.distributions_returns = true
  This month's distribution, yield, preferred return status

DEAL ACTIVITY (2–3 sentences) — include only if dam.deal_activity_new OR dam.deal_activity_repaid OR dam.deal_pipeline = true
  New originations, repayments, pipeline outlook

OUTLOOK (1–2 sentences)
  Forward-looking note; keep measured and data-grounded
```

**Do not fabricate data.** If the Executive Summary cannot be written because nav_fund_size and distributions_returns are both false, write only: "Portfolio Update: [month] [year]" and move directly to whichever sections have data.

### Step 6 — Build the PDF Report

Use reportlab (Python) to produce a professional PDF. Follow the layout in `references/pdf_template.md`.

**Apply DAM-conditional rendering — only include pages/sections whose DAM flag is `true`:**

| PDF Section | DAM Gate |
|-------------|----------|
| Cover page | Always included |
| Executive Summary narrative | Always included (with available data only) |
| NAV & Fund Size stats table | `dam.nav_fund_size` |
| Loan Portfolio Metrics table (with MoM delta column) | `dam.loan_portfolio_metrics` |
| Property Type breakdown | `dam.loan_portfolio_property_type` |
| Geographic breakdown | `dam.loan_portfolio_geo` |
| Distributions & Returns stats | `dam.distributions_returns` |
| Deal Activity — New Originations | `dam.deal_activity_new` |
| Deal Activity — Repayments | `dam.deal_activity_repaid` |
| Pipeline | `dam.deal_pipeline` |
| Appendix: Loan Schedule | `dam.loan_schedule` |
| Appendix: Raw Data Table | Always included (with only available fields) |

Key design requirements:
- Cover page with fund name, report period, and "CONFIDENTIAL — FOR INVESTOR USE ONLY"
- Color palette: Deep navy (`#1B2A4A`) primary, light grey (`#F5F6F8`) background, gold accent (`#C9A84C`)
- Page numbers in footer; fund name in header

Save to: `/outputs/[FUND_NAME]_[PERIOD]_Investor_Report.pdf`

Read `references/pdf_template.md` for the full layout specification before coding.

### Step 7 — Build the PowerPoint Deck

Use pptxgenjs (Node.js) to produce a polished deck. Follow the layout in `references/pptx_template.md`.

**Apply DAM-conditional rendering — only build slides whose DAM flag is `true`:**

| Slide | Content | DAM Gate |
|-------|---------|----------|
| 1 | Cover — Fund name, period, "CONFIDENTIAL" | Always |
| 2 | Executive Summary — narrative + key stats | Always |
| 3 | NAV & Fund Size — large callout numbers | `dam.nav_fund_size` |
| 4 | Loan Portfolio Snapshot — table + property type bars | `dam.loan_portfolio_metrics` |
| 5 | Portfolio Composition — geo table + property type | `dam.loan_portfolio_geo` OR `dam.loan_portfolio_property_type` |
| 6 | Distributions & Returns — yield callout, pref return | `dam.distributions_returns` |
| 7 | Deal Activity — new deals, repayments, pipeline | `dam.deal_activity_new` OR `dam.deal_activity_repaid` OR `dam.deal_pipeline` |
| 8 | Outlook — forward-looking commentary | Always |
| 9 | Appendix — full data table | Always (with available fields only) |

**Important slide-level rules:**
- If a slide is gated by multiple flags (e.g., Slide 5), include it if **any** of those flags is `true`, but only render the sub-sections with data.
- Never include a chart or bar visual if its underlying data has fewer than 2 data points.
- Never include a table row with all-null values.
- Slide count will vary run-to-run based on available data — this is expected and correct.

Design guidelines (from pptx skill):
- Dark navy cover + conclusion slides, light background for content slides
- Large stat callouts (60–72pt numbers with 12pt labels)
- No plain bullet-only slides — every content slide has at least one visual element
- Font: Calibri headers, Calibri Light body
- No accent lines under titles

Save to: `/outputs/[FUND_NAME]_[PERIOD]_Investor_Deck.pptx`

Read `references/pptx_template.md` for detailed slide specs before coding.

### Step 8 — QA & Verification

Before declaring done:

1. **DAM check**: Confirm every section in both outputs corresponds to a `true` DAM flag. No section with `false` should appear anywhere.
2. **Data check**: Re-read the source file and spot-check 3 metrics against the PDF/PPTX. Confirm no transposition errors.
3. **PDF check**: Open the PDF and verify cover page, all sections present, no broken layout.
4. **PPTX check**: Convert to images and visually inspect each slide for overlaps, cut-off text, placeholder content.
5. **Commentary check**: Re-read the narrative. Flag any claim not supported by extracted data.
6. **Completeness check**: Confirm both files are saved to `/outputs/`.

If any check fails, fix and re-verify before presenting to the user.

### Step 9 — Present to Human for Review

After QA, present both output files to the user and give a brief summary:

> "Here are your [Month] [Year] investor reports for [Fund Name]:
> - [PDF link] — Investor Report PDF
> - [PPTX link] — Investor Deck
>
> Key numbers this month: NAV $X.X M ([delta]%), distribution yield X.X%, [X] active loans.
>
> **Sections included:** [list DAM-true sections]
> **Sections omitted** (not found in source file): [list DAM-false sections]
>
> Please review before distributing to investors — I've flagged [N] items that may need your attention: [list any ⚠️ items]."

---

## Reference Files

Read these before building outputs:

- `references/pdf_template.md` — Full PDF layout, color codes, section specs
- `references/pptx_template.md` — Full slide-by-slide PowerPoint spec

---

## Troubleshooting

| Problem | Action |
|---------|--------|
| Google Drive not connected | Tell user to connect Google Drive MCP (see onboarding guide) |
| File is password-protected PDF | Ask user to provide the password or an unlocked copy |
| File is a scanned image (no text layer) | Use OCR (pytesseract) — note this in report as "OCR-extracted, verify accuracy" |
| Prior month file not found | Set `dam.prior_month_comparison = false`; proceed without comparison |
| Metric not in source file | Set to `null`; DAM will gate the section automatically |
| All DAM flags are false | Stop and tell user: "No structured fund data was found in the source file. Please verify the file format and re-upload." |
| Fewer than 3 DAM flags are true | Warn user before generating: "Only partial data was found. The report will be abbreviated. Is that okay?" |

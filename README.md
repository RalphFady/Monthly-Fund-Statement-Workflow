# Monthly Fund Statement Workflow

Automated monthly investor reporting for a commercial real estate private credit fund. Each month-end, this workflow finds the latest fund statement in Google Drive, extracts key metrics, compares to the prior month, and produces a branded PDF report and PowerPoint deck ready for LP distribution.

## What's in this repo

```
monthly-fund-statement.skill   ← Install this in Claude desktop (click "Save skill")
brand.json                     ← Fund configuration (name, colors, Drive folder, preferences)
generate_pdf.py                ← Python/reportlab PDF generator
generate_deck.js               ← Node.js/pptxgenjs PowerPoint generator
skill/
  SKILL.md                     ← Full skill instructions + workflow
  references/
    pdf_template.md            ← PDF layout spec (colors, sections, reportlab patterns)
    pptx_template.md           ← PPTX slide-by-slide spec (pptxgenjs patterns)
outputs/
  *.pdf                        ← Sample investor report
  *.pptx                       ← Sample investor deck
```

## How to use

### 1. Install the skill
Open the Claude desktop app, then click **"Save skill"** on the `monthly-fund-statement.skill` file. This installs the workflow into Claude.

### 2. Connect Google Drive
On first run, Claude will prompt you to connect Google Drive. This is required — it's how Claude finds your monthly statement file automatically.

### 3. Run onboarding
Say `run the monthly fund statement` in Claude. On first run, Claude will ask for:
- Fund name
- Google Drive folder name
- Brand colors
- Logo (drag a PNG or SVG into the chat)
- Report preferences (loan schedule, schedule timing)

Claude saves your answers to `brand.json` — you're never asked again.

### 4. Monthly runs
After setup, the report runs automatically on the 1st of every month. Claude finds the latest file in your Drive folder, builds a Data Availability Map (only renders sections with real data), and outputs a PDF + PPTX to your outputs folder.

You can also trigger it manually at any time by saying `run the monthly fund statement`.

## Data Availability Map (DAM)

The workflow only builds sections where actual data was found in the source file. If geographic data is missing, that slide/page is skipped entirely — no blank charts. The DAM is printed to console on every run so you can see exactly what was included.

## Sample outputs

See the `outputs/` folder for a sample PDF report and PowerPoint deck generated from mock data.

## Requirements

- Claude desktop app with the skill installed
- Google Drive MCP connector (prompted during onboarding)
- Python 3 + `reportlab` (`pip install reportlab`)
- Node.js + `pptxgenjs` (`npm install pptxgenjs`)

"""
Monthly Fund Statement — PDF Generator
Cana Capital Senior Credit Fund I, L.P. | April 2026
"""

import json
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.pdfgen import canvas
from reportlab.lib import colors

# ── BRAND ─────────────────────────────────────────────
NAVY       = HexColor('#1B2A4A')
GOLD       = HexColor('#C9A84C')
BG_LIGHT   = HexColor('#F7F8FA')
TEXT_DARK  = HexColor('#2C3E50')
MUTED      = HexColor('#8899AA')
GREEN      = HexColor('#1A6B3A')
RED        = HexColor('#B71C1C')
WHITE      = white
MED_BLUE   = HexColor('#3D6B9E')
ALT_ROW    = HexColor('#EEF2F7')

FUND_NAME    = "Cana Capital Senior Credit Fund I, L.P."
FUND_SHORT   = "Cana Capital Fund I"
STRATEGY     = "Senior Secured Commercial Real Estate Debt"
MANAGER      = "Cana Ventures"
PERIOD       = "April 2026"
CONTACT_EMAIL = "ir@canaventures.com"
CONTACT_WEB  = "www.canaventures.com"

W, H = letter  # 612 x 792

# ── MOCK DATA ──────────────────────────────────────────
cur = {
    "nav":               118_400_000,
    "gav":               145_200_000,
    "committed":         150_000_000,
    "called":            112_500_000,
    "uncalled":           37_500_000,
    "nav_per_unit":          105.24,
    "loan_book":         138_200_000,
    "active_loans":              14,
    "wtd_ltv":                 62.5,
    "wtd_dscr":                1.35,
    "wtd_coupon":              8.85,
    "avg_term":                  18,
    "occupancy":               91.3,
    "performing_pct":          100.0,
    "monthly_dist":            0.095,
    "ann_yield":               8.2,
    "pref_return":             7.50,
    "pref_current":            True,
    "cumulative_dist":     6_840_000,
    "net_irr":                 9.1,
    "moic":                   1.08,
    "tvpi":                   1.08,
    "new_orig_count":             2,
    "new_orig_total":      14_000_000,
    "repay_count":                1,
    "repay_total":          9_200_000,
    "pipeline_count":             3,
    "pipeline_total":      22_000_000,
}
prev = {
    "nav":               116_500_000,
    "gav":               143_100_000,
    "called":            110_500_000,
    "uncalled":           39_500_000,
    "nav_per_unit":          103.65,
    "loan_book":         136_000_000,
    "active_loans":              13,
    "wtd_ltv":                 63.1,
    "wtd_dscr":                1.33,
    "wtd_coupon":              8.75,
    "avg_term":                  19,
    "occupancy":               90.1,
    "performing_pct":          100.0,
    "monthly_dist":            0.093,
    "ann_yield":               8.0,
    "cumulative_dist":     6_205_000,
}

def fmt_m(v):   return f"${v/1e6:,.1f}M"
def fmt_pct(v): return f"{v:.1f}%"
def fmt_x(v):   return f"{v:.2f}x"
def fmt_bps(a, b):
    d = (a - b) * 100
    arrow = "▲" if d > 0 else "▼"
    col = GREEN if d > 0 else RED
    return f"{arrow} {abs(d):.0f} bps", col
def fmt_delta_pct(a, b, pos_good=True):
    if b == 0: return "—", MUTED
    d = (a - b) / b * 100
    arrow = "▲" if d > 0 else "▼"
    col = (GREEN if d > 0 else RED) if pos_good else (RED if d > 0 else GREEN)
    return f"{arrow} {abs(d):.1f}%", col
def fmt_delta_abs(a, b, pos_good=True):
    d = a - b
    arrow = "▲" if d > 0 else "▼"
    col = (GREEN if d > 0 else RED) if pos_good else (RED if d > 0 else GREEN)
    return f"{arrow} {abs(d):.2f}", col

PROPERTY_TYPES = [
    ("Multifamily",  35, 48_370_000),
    ("Industrial",   28, 38_696_000),
    ("Mixed-Use",    18, 24_876_000),
    ("Retail",       12, 16_584_000),
    ("Office",        7,  9_674_000),
]
GEOGRAPHY = [
    ("Texas",          5,  44_224_000, 32.0),
    ("Florida",        3,  33_168_000, 24.0),
    ("Southeast",      3,  26_258_000, 19.0),
    ("Southwest",      2,  20_730_000, 15.0),
    ("Other",          1,  13_820_000, 10.0),
]
NEW_DEALS = [
    ("Riverside Commons", "Austin, TX",   "$8.5M", "58.2%", "SOFR+350", "24 mo", "Multifamily"),
    ("Harbor Point Ind.", "Tampa, FL",    "$5.5M", "63.8%", "SOFR+375", "18 mo", "Industrial"),
]
DIST_HISTORY = [
    ("Nov 2025", "$0.088", "7.8%"),
    ("Dec 2025", "$0.090", "7.9%"),
    ("Jan 2026", "$0.091", "8.0%"),
    ("Feb 2026", "$0.092", "8.1%"),
    ("Mar 2026", "$0.093", "8.0%"),
    ("Apr 2026", "$0.095", "8.2%"),
]

DISCLAIMER = (
    "This report is confidential and prepared solely for the addressee(s). It is not to be copied, "
    "distributed, or used by any other person or for any other purpose without the prior written "
    "consent of Cana Ventures. This material does not constitute an offer to sell or solicitation "
    "of an offer to buy any securities. Past performance is not indicative of future results. "
    "Investing in private credit funds involves substantial risk, including the potential loss of "
    "principal. Forward-looking statements reflect management's current expectations and are subject "
    "to uncertainty. Figures presented are unaudited and subject to change. Returns are net of all "
    "fees and expenses unless otherwise noted. This report is prepared for informational purposes "
    "only for current limited partners of Cana Capital Senior Credit Fund I, L.P."
)

COMMENTARY = {
    "exec": (
        "The Fund delivered another month of stable performance in April 2026, with Net Asset Value "
        "growing to $118.4 million — a 1.6% increase from March — driven by continued interest "
        "income and two new loan originations. The portfolio remains 100% performing with a weighted "
        "average DSCR of 1.35x, providing strong coverage across all 14 active loans. "
        "Monthly distributions of $0.095 per unit were paid on schedule, representing an annualized "
        "yield of 8.2% and maintaining the fund's preferred return of 7.5% in current status. "
        "The fund's defensive positioning in senior secured loans — with a weighted average LTV of "
        "62.5% — continues to provide robust downside protection in the current rate environment."
    ),
    "portfolio": (
        "The loan portfolio expanded to $138.2 million across 14 positions following the close of "
        "two new originations totaling $14.0 million: a $8.5 million multifamily bridge loan in "
        "Austin, TX and a $5.5 million industrial facility in Tampa, FL. One loan totaling $9.2 "
        "million repaid at par during the period. Portfolio credit metrics improved modestly, with "
        "weighted average DSCR increasing to 1.35x from 1.33x and weighted average LTV tightening "
        "to 62.5% from 63.1%. Multifamily (35%) and industrial (28%) remain the dominant property "
        "type exposures, consistent with the fund's thesis of targeting recession-resilient asset classes."
    ),
    "distributions": (
        "Monthly distributions of $0.095 per unit were paid on April 30, 2026, representing an "
        "annualized yield of 8.2% — above the fund's 7.5% preferred return hurdle. Cumulative "
        "distributions paid to limited partners since inception total $6.84 million. The net IRR "
        "since inception stands at 9.1% on a net basis, with a MOIC of 1.08x reflecting "
        "distributions paid plus remaining NAV. The preferred return of 7.5% per annum remains "
        "current with no accrued shortfall."
    ),
    "deals": (
        "The fund closed two new senior secured loans in April, deploying $14.0 million of capital "
        "at a blended weighted average LTV of 60.5% and coupon of SOFR+360 bps. The Riverside "
        "Commons multifamily loan in Austin represents the fund's fifth Texas exposure, reinforcing "
        "the fund's concentration in high-growth Sunbelt markets. One loan — a $9.2 million retail "
        "center in Atlanta — repaid at par following a successful refinancing by the sponsor. "
        "The fund has a pipeline of three additional loans totaling $22 million expected to close "
        "in Q2 2026, which would increase loan book utilization to approximately 97% of called capital."
    ),
    "outlook": (
        "The fund's pipeline of $22 million in near-term originations positions the portfolio for "
        "continued income growth through Q2 2026. Credit conditions across the Sunbelt markets "
        "where the fund is concentrated remain stable, supported by sustained population inflows "
        "and above-average rent growth in multifamily and industrial sectors. Management continues "
        "to maintain disciplined underwriting standards with maximum LTVs of 70% on new originations."
    ),
}

OUT_PATH = "/sessions/zen-trusting-carson/mnt/outputs/Cana_Capital_Fund_I_April_2026_Investor_Report.pdf"

# ══════════════════════════════════════════════════════
# STYLES
# ══════════════════════════════════════════════════════
styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, **kw)

s_body      = S('body',      fontName='Helvetica',      fontSize=10, leading=15, textColor=TEXT_DARK)
s_body_sm   = S('body_sm',   fontName='Helvetica',      fontSize=9,  leading=13, textColor=TEXT_DARK)
s_bold      = S('bold',      fontName='Helvetica-Bold', fontSize=10, leading=15, textColor=TEXT_DARK)
s_muted     = S('muted',     fontName='Helvetica',      fontSize=8,  leading=12, textColor=MUTED)
s_section   = S('section',   fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=WHITE, backColor=NAVY)
s_title     = S('title',     fontName='Helvetica-Bold', fontSize=17, leading=21, textColor=NAVY)
s_gold_big  = S('gold_big',  fontName='Helvetica-Bold', fontSize=26, leading=30, textColor=GOLD, alignment=TA_CENTER)
s_label     = S('label',     fontName='Helvetica',      fontSize=8,  leading=11, textColor=MUTED, alignment=TA_CENTER)
s_delta_g   = S('delta_g',   fontName='Helvetica-Bold', fontSize=8,  leading=11, textColor=GREEN,   alignment=TA_CENTER)
s_delta_r   = S('delta_r',   fontName='Helvetica-Bold', fontSize=8,  leading=11, textColor=RED,     alignment=TA_CENTER)
s_delta_m   = S('delta_m',   fontName='Helvetica',      fontSize=8,  leading=11, textColor=MUTED,   alignment=TA_CENTER)

# ══════════════════════════════════════════════════════
# COVER PAGE  (full canvas, no platypus)
# ══════════════════════════════════════════════════════
def draw_cover(c, doc):
    c.saveState()
    # full navy background
    c.setFillColor(NAVY)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # gold rule at 55% height
    rule_y = H * 0.57
    c.setStrokeColor(GOLD)
    c.setLineWidth(2)
    c.line(54, rule_y, W - 54, rule_y)

    # fund name
    c.setFillColor(WHITE)
    c.setFont('Helvetica-Bold', 34)
    c.drawString(54, rule_y + 50, FUND_NAME.replace(", L.P.", ""))
    c.setFont('Helvetica', 18)
    c.setFillColor(GOLD)
    c.drawString(54, rule_y + 24, "L.P.")
    c.setFillColor(WHITE)

    # strategy
    c.setFillColor(GOLD)
    c.setFont('Helvetica', 14)
    c.drawString(54, rule_y - 22, STRATEGY)

    # report type
    c.setFillColor(WHITE)
    c.setFont('Helvetica', 13)
    c.drawString(54, rule_y - 54, "Monthly Investor Report")
    c.setFont('Helvetica-Bold', 24)
    c.drawString(54, rule_y - 84, PERIOD)

    # bottom gold rule
    c.setStrokeColor(GOLD)
    c.setLineWidth(1.5)
    c.line(54, 80, W - 54, 80)

    # confidentiality
    c.setFillColor(MUTED)
    c.setFont('Helvetica', 8)
    c.drawString(54, 64, "CONFIDENTIAL — FOR LIMITED PARTNERS ONLY")
    c.drawRightString(W - 54, 64, f"Prepared by {MANAGER}  |  {PERIOD}")

    # top bar
    c.setFillColor(HexColor('#0D1925'))
    c.rect(0, H - 52, W, 52, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.setFont('Helvetica-Bold', 12)
    c.drawString(54, H - 32, MANAGER.upper())
    c.setFillColor(MUTED)
    c.setFont('Helvetica', 9)
    c.drawRightString(W - 54, H - 32, CONTACT_WEB)

    c.restoreState()

# ══════════════════════════════════════════════════════
# INTERIOR HEADER / FOOTER
# ══════════════════════════════════════════════════════
class FundTemplate:
    page_count = 0

def header_footer(c, doc):
    page = doc.page
    c.saveState()
    # header bar
    c.setFillColor(NAVY)
    c.rect(0, H - 36, W, 36, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont('Helvetica-Bold', 9)
    c.drawString(54, H - 22, FUND_SHORT)
    c.setFont('Helvetica', 9)
    c.drawRightString(W - 54, H - 22, PERIOD)

    # gold accent line under header
    c.setStrokeColor(GOLD)
    c.setLineWidth(1)
    c.line(0, H - 37, W, H - 37)

    # footer
    c.setFillColor(MUTED)
    c.setFont('Helvetica', 7)
    c.drawString(54, 22, "CONFIDENTIAL — FOR LIMITED PARTNERS ONLY")
    c.drawCentredString(W / 2, 22, f"Page {page} of 9")
    c.drawRightString(W - 54, 22, f"{PERIOD}  |  {CONTACT_WEB}")

    c.restoreState()

# ══════════════════════════════════════════════════════
# HELPER — section header bar
# ══════════════════════════════════════════════════════
def section_bar(title):
    t = Table([[Paragraph(f"<b>{title.upper()}</b>", S('sh', fontName='Helvetica-Bold', fontSize=10, leading=13, textColor=WHITE))]],
              colWidths=[W - 108])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), NAVY),
        ('TOPPADDING',    (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING',   (0, 0), (-1, -1), 10),
    ]))
    return t

def gold_rule():
    return HRFlowable(width="100%", thickness=0.5, color=GOLD, spaceAfter=8, spaceBefore=4)

def stat_box(value, label, delta_txt, delta_color):
    val_style   = S('sv', fontName='Helvetica-Bold', fontSize=22, leading=26, textColor=GOLD, alignment=TA_CENTER)
    label_style = S('sl', fontName='Helvetica',      fontSize=8,  leading=11, textColor=MUTED, alignment=TA_CENTER)
    delta_style = S('sd', fontName='Helvetica-Bold', fontSize=8,  leading=11, textColor=delta_color, alignment=TA_CENTER)
    data = [[Paragraph(value, val_style)],
            [Paragraph(label, label_style)],
            [Paragraph(delta_txt, delta_style)]]
    t = Table(data, colWidths=[142])
    t.setStyle(TableStyle([
        ('BOX',           (0,0), (-1,-1), 0.75, GOLD),
        ('TOPPADDING',    (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING',   (0,0), (-1,-1), 4),
        ('RIGHTPADDING',  (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS',(0,0), (-1,-1), [WHITE]),
    ]))
    return t

# ══════════════════════════════════════════════════════
# BUILD DOCUMENT
# ══════════════════════════════════════════════════════
doc = SimpleDocTemplate(
    OUT_PATH,
    pagesize=letter,
    topMargin=54, bottomMargin=44,
    leftMargin=54, rightMargin=54,
    title=f"{FUND_SHORT} — {PERIOD} Investor Report",
    author=MANAGER,
    subject="Monthly Investor Report — Confidential",
)

story = []

# ── PAGE 1: COVER  (blank platypus page, drawn via onFirstPage) ──
story.append(PageBreak())

# ── PAGE 2: DISCLAIMER & CONTENTS ──────────────────────
story.append(section_bar("Important Disclosures"))
story.append(Spacer(1, 10))
story.append(Paragraph(DISCLAIMER, S('dis', fontName='Helvetica', fontSize=8, leading=12, textColor=TEXT_DARK)))
story.append(Spacer(1, 18))

toc_data = [
    ["Section", "Page"],
    ["Executive Summary", "3"],
    ["Loan Portfolio Snapshot", "4"],
    ["Distributions & Returns", "5"],
    ["Deal Activity", "6"],
    ["Portfolio Composition", "7"],
    ["Outlook", "8"],
    ["Appendix — Full Data", "9"],
]
toc_style = TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('TEXTCOLOR',  (0,0), (-1,0), WHITE),
    ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
    ('FONTSIZE',   (0,0), (-1,-1), 9),
    ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, ALT_ROW]),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('LEFTPADDING', (0,0), (-1,-1), 10),
    ('RIGHTPADDING',(0,0), (-1,-1), 10),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
])
toc = Table(toc_data, colWidths=[350, 80])
toc.setStyle(toc_style)
story.append(Paragraph("<b>Table of Contents</b>", S('toc_h', fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=NAVY)))
story.append(Spacer(1, 8))
story.append(toc)
story.append(PageBreak())

# ── PAGE 3: EXECUTIVE SUMMARY ───────────────────────────
story.append(Paragraph("Executive Summary", S('pt', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)))
story.append(gold_rule())

# 3-stat callout row
d_nav, d_nav_c   = fmt_delta_pct(cur['nav'], prev['nav'])
d_yield, d_yield_c = fmt_bps(cur['ann_yield'], prev['ann_yield'])
d_loans = f"▲ +1 loan" if cur['active_loans'] > prev['active_loans'] else "—"

stat_row = Table([
    [stat_box(fmt_m(cur['nav']),        "Total Fund NAV",         d_nav,   d_nav_c),
     stat_box(fmt_pct(cur['ann_yield']),"Distribution Yield",     d_yield, d_yield_c),
     stat_box(str(cur['active_loans']), "Active Loans",           d_loans, GREEN)],
], colWidths=[154, 154, 154], hAlign='LEFT')
stat_row.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('LEFTPADDING',(0,0),(-1,-1),0), ('RIGHTPADDING',(0,0),(-1,-1),8)]))
story.append(stat_row)
story.append(Spacer(1, 14))

# Narrative + sidebar
pref_color = GREEN if cur['pref_current'] else RED
pref_txt   = "CURRENT ✓" if cur['pref_current'] else "BEHIND ⚠"
sidebar_data = [
    [Paragraph("<b>Fund Health</b>", S('fh', fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=WHITE, backColor=NAVY))],
    [Paragraph(f"Preferred Return: <font color='#90EE90'>{pref_txt}</font>", S('fhi', fontName='Helvetica', fontSize=9, leading=13, textColor=WHITE, backColor=NAVY))],
    [Paragraph("Portfolio: 100% Performing", S('fhi2', fontName='Helvetica', fontSize=9, leading=13, textColor=WHITE, backColor=NAVY))],
    [Paragraph(f"Capital: {cur['called']/cur['committed']*100:.0f}% Deployed", S('fhi3', fontName='Helvetica', fontSize=9, leading=13, textColor=WHITE, backColor=NAVY))],
    [Paragraph(f"Net IRR: {cur['net_irr']:.1f}%", S('fhi4', fontName='Helvetica', fontSize=9, leading=13, textColor=WHITE, backColor=NAVY))],
    [Paragraph(f"MOIC: {fmt_x(cur['moic'])}", S('fhi5', fontName='Helvetica', fontSize=9, leading=13, textColor=WHITE, backColor=NAVY))],
]
sidebar = Table(sidebar_data, colWidths=[130])
sidebar.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), NAVY),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 10),
    ('LINEBELOW', (0,0), (0,0), 0.5, GOLD),
]))

exec_table = Table([
    [Paragraph(COMMENTARY['exec'], s_body), sidebar]
], colWidths=[310, 140])
exec_table.setStyle(TableStyle([
    ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ('LEFTPADDING', (0,0), (-1,-1), 0),
    ('RIGHTPADDING',(0,0), (0,0), 14),
]))
story.append(exec_table)

# Sub-metrics row
story.append(Spacer(1, 14))
sub_data = [
    [Paragraph("<b>Gross Asset Value</b>", s_bold),  Paragraph(fmt_m(cur['gav']),       s_body),
     Paragraph("<b>Called Capital</b>",    s_bold),  Paragraph(fmt_m(cur['called']),     s_body),
     Paragraph("<b>Uncalled / Dry Powder</b>", s_bold), Paragraph(fmt_m(cur['uncalled']), s_body)],
]
sub_t = Table(sub_data, colWidths=[110, 65, 90, 65, 110, 65])
sub_t.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), BG_LIGHT),
    ('BOX',        (0,0), (-1,-1), 0.5, HexColor('#DDE3EC')),
    ('TOPPADDING',    (0,0), (-1,-1), 7),
    ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ('LEFTPADDING',   (0,0), (-1,-1), 10),
]))
story.append(sub_t)
story.append(PageBreak())

# ── PAGE 4: LOAN PORTFOLIO SNAPSHOT ────────────────────
story.append(Paragraph("Loan Portfolio Snapshot", S('pt2', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)))
story.append(gold_rule())

port_headers = [
    Paragraph("<b>Metric</b>",           S('th', fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=WHITE)),
    Paragraph("<b>April 2026</b>",        S('th', fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>March 2026</b>",        S('th', fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Change</b>",            S('th', fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=WHITE, alignment=TA_RIGHT)),
]

def prow(label, cur_val, prev_val, delta_txt, delta_col):
    return [
        Paragraph(label,    s_body_sm),
        Paragraph(cur_val,  S('rv', fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=TEXT_DARK, alignment=TA_RIGHT)),
        Paragraph(prev_val, S('pv', fontName='Helvetica',      fontSize=9, leading=12, textColor=MUTED,     alignment=TA_RIGHT)),
        Paragraph(delta_txt, S('dv', fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=delta_col, alignment=TA_RIGHT)),
    ]

dt_nav,   dc_nav   = fmt_delta_pct(cur['loan_book'],   prev['loan_book'])
dt_loans, dc_loans = (f"▲ +1", GREEN) if cur['active_loans'] > prev['active_loans'] else ("—", MUTED)
dt_ltv,   dc_ltv   = fmt_bps(cur['wtd_ltv'],  prev['wtd_ltv'],  )
dc_ltv = RED if cur['wtd_ltv'] > prev['wtd_ltv'] else GREEN   # lower LTV is better
dt_ltv = f"{'▲' if cur['wtd_ltv']>prev['wtd_ltv'] else '▼'} {abs(cur['wtd_ltv']-prev['wtd_ltv'])*100:.0f} bps"
dt_dscr,  dc_dscr  = fmt_delta_abs(cur['wtd_dscr'],  prev['wtd_dscr'])
dt_coup,  dc_coup  = fmt_bps(cur['wtd_coupon'], prev['wtd_coupon'])
dt_term,  dc_term  = (f"▼ 1 mo", MUTED)
dt_occ,   dc_occ   = fmt_bps(cur['occupancy'], prev['occupancy'])

port_rows = [
    prow("Total Loan Book Balance",         fmt_m(cur['loan_book']),     fmt_m(prev['loan_book']),     dt_nav,   dc_nav),
    prow("Number of Active Loans",          str(cur['active_loans']),    str(prev['active_loans']),    dt_loans, GREEN),
    prow("Wtd. Avg. Loan-to-Value (LTV)",   fmt_pct(cur['wtd_ltv']),    fmt_pct(prev['wtd_ltv']),    dt_ltv,   GREEN),
    prow("Wtd. Avg. DSCR",                  fmt_x(cur['wtd_dscr']),     fmt_x(prev['wtd_dscr']),     dt_dscr,  dc_dscr),
    prow("Wtd. Avg. Coupon / Rate",         fmt_pct(cur['wtd_coupon']), fmt_pct(prev['wtd_coupon']), dt_coup,  dc_coup),
    prow("Avg. Remaining Term",             f"{cur['avg_term']} months", f"{prev['avg_term']} months", dt_term, MUTED),
    prow("Portfolio Occupancy Rate",        fmt_pct(cur['occupancy']),  fmt_pct(prev['occupancy']),  dt_occ,   dc_occ),
    prow("Performing Loans",                "14 / 100%",                "13 / 100%",                 "—",      MUTED),
]

port_table = Table([port_headers] + port_rows, colWidths=[200, 85, 85, 85])
port_ts = TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, ALT_ROW]),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('TOPPADDING',    (0,0), (-1,-1), 7),
    ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
])
port_table.setStyle(port_ts)
story.append(port_table)

story.append(Spacer(1, 16))
story.append(gold_rule())

# Property type breakdown (text table since no chart library inline)
story.append(Paragraph("<b>Portfolio Composition by Property Type</b>", s_bold))
story.append(Spacer(1, 8))
prop_headers = [
    Paragraph("<b>Property Type</b>", S('pth', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b>% Allocation</b>",  S('pth', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Loan Balance</b>",  S('pth', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
]
prop_rows = [[
    Paragraph(pt, s_body_sm),
    Paragraph(f"{pct}%", S('pr', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)),
    Paragraph(fmt_m(bal), S('pr2', fontName='Helvetica', fontSize=9, textColor=TEXT_DARK, alignment=TA_RIGHT)),
] for pt, pct, bal in PROPERTY_TYPES]
prop_table = Table([prop_headers] + prop_rows, colWidths=[200, 100, 155])
prop_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, ALT_ROW]),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
]))
story.append(prop_table)
story.append(PageBreak())

# ── PAGE 5: DISTRIBUTIONS & RETURNS ────────────────────
story.append(Paragraph("Distributions & Returns", S('pt3', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)))
story.append(gold_rule())

# 4-stat row
d_dist, dc_dist = fmt_delta_abs(cur['monthly_dist'], prev['monthly_dist'])
d_yield2, dc_yield2 = fmt_bps(cur['ann_yield'], prev['ann_yield'])

pref_box_data = [[Paragraph(
    f"<b>Preferred Return: {fmt_pct(cur['pref_return'])} p.a.</b><br/>"
    + ("<font color='#1A6B3A'>● CURRENT ✓</font>" if cur['pref_current'] else "<font color='#B71C1C'>⚠ BEHIND</font>"),
    S('pb', fontName='Helvetica-Bold', fontSize=11, leading=16, textColor=WHITE, alignment=TA_CENTER)
)]]
pref_t = Table(pref_box_data, colWidths=[460])
pref_t.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), NAVY),
    ('TOPPADDING',    (0,0), (-1,-1), 10),
    ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ('BOX',       (0,0), (-1,-1), 1.5, GOLD),
]))
story.append(pref_t)
story.append(Spacer(1, 12))

dist_stat_row = Table([
    [stat_box(f"${cur['monthly_dist']:.3f}", "Monthly Distribution / Unit",  d_dist,   dc_dist),
     stat_box(fmt_pct(cur['ann_yield']),     "Annualized Yield",             d_yield2, dc_yield2),
     stat_box(fmt_m(cur['cumulative_dist']), "Cumulative Distributions",     "Since Inception", MUTED)],
], colWidths=[154, 154, 154])
dist_stat_row.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),8)]))
story.append(dist_stat_row)
story.append(Spacer(1, 14))

# Returns table + narrative
ret_headers = [
    Paragraph("<b>Return Metric</b>", S('rth', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b>Net</b>",   S('rth', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Gross</b>", S('rth', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
]
ret_rows = [
    [Paragraph("Net IRR (Since Inception)", s_body_sm), Paragraph("<b>9.1%</b>", S('rbv', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)), Paragraph("10.8%", S('rbg', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT))],
    [Paragraph("Cash-on-Cash Multiple (MOIC)", s_body_sm), Paragraph("<b>1.08x</b>", S('rbv', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)), Paragraph("N/A", S('rbg', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT, textColor=MUTED))],
    [Paragraph("TVPI (Net)", s_body_sm), Paragraph("<b>1.08x</b>", S('rbv', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)), Paragraph("N/A", S('rbg', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT, textColor=MUTED))],
    [Paragraph("Preferred Return Hurdle", s_body_sm), Paragraph("7.50%", S('rbg2', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)), Paragraph("—", S('rbg', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT, textColor=MUTED))],
]
ret_note = Paragraph("<i>Returns are unaudited. Net returns are after all management fees, carried interest, and fund expenses.</i>", s_muted)

ret_dist_col = Table(
    [ret_headers] + ret_rows + [[ret_note, "", ""]],
    colWidths=[115, 42, 38]
)
ret_dist_col.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-1,-2), [WHITE, ALT_ROW]),
    ('SPAN', (0,-1), (-1,-1)),
    ('GRID',       (0,0), (-1,-2), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
]))

# Distribution history table
dist_h_headers = [
    Paragraph("<b>Period</b>",        S('dh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b>$/Unit</b>",        S('dh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Ann. Yield</b>",    S('dh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
]
dist_h_rows = [[
    Paragraph(p, s_body_sm if i < 5 else S('db', fontName='Helvetica-Bold', fontSize=9, textColor=TEXT_DARK)),
    Paragraph(d, S('dr', fontName='Helvetica' if i < 5 else 'Helvetica-Bold', fontSize=9, alignment=TA_RIGHT, textColor=GOLD if i==5 else TEXT_DARK)),
    Paragraph(y, S('dy', fontName='Helvetica' if i < 5 else 'Helvetica-Bold', fontSize=9, alignment=TA_RIGHT, textColor=GOLD if i==5 else TEXT_DARK)),
] for i,(p,d,y) in enumerate(DIST_HISTORY)]
dist_h = Table([dist_h_headers] + dist_h_rows, colWidths=[62, 38, 44])
dist_h.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, ALT_ROW]),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('TOPPADDING',    (0,0), (-1,-1), 5),
    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ('LEFTPADDING',   (0,0), (-1,-1), 6),
    ('RIGHTPADDING',  (0,0), (-1,-1), 6),
]))

narrative_p = Paragraph(COMMENTARY['distributions'], s_body)

full_row = Table([[narrative_p, ret_dist_col, dist_h]], colWidths=[165, 195, 144])
full_row.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(0,0),12),('RIGHTPADDING',(1,0),(1,0),10)]))
story.append(full_row)
story.append(PageBreak())

# ── PAGE 6: DEAL ACTIVITY ───────────────────────────────
story.append(Paragraph("Deal Activity — April 2026", S('pt4', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)))
story.append(gold_rule())

box_labels = [
    (f"{cur['new_orig_count']} loans / {fmt_m(cur['new_orig_total'])}", "New Originations", NAVY),
    (f"{cur['repay_count']} loan / {fmt_m(cur['repay_total'])}", "Repayments / Exits", MED_BLUE),
    (f"{cur['pipeline_count']} loans / {fmt_m(cur['pipeline_total'])}", "Pipeline", HexColor('#4A6741')),
]
box_cells = []
for val, lbl, col in box_labels:
    t = Table([[Paragraph(f"<b>{val}</b>", S('bv', fontName='Helvetica-Bold', fontSize=14, leading=18, textColor=WHITE, alignment=TA_CENTER))],
               [Paragraph(lbl, S('bl', fontName='Helvetica', fontSize=9, leading=12, textColor=HexColor('#AACCEE'), alignment=TA_CENTER))]],
              colWidths=[152])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),col),('TOPPADDING',(0,0),(-1,-1),10),('BOTTOMPADDING',(0,0),(-1,-1),8),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4)]))
    box_cells.append(t)

box_row = Table([box_cells], colWidths=[154, 154, 154])
box_row.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),8)]))
story.append(box_row)
story.append(Spacer(1, 14))

# New deals table
story.append(Paragraph("<b>New Originations — April 2026</b>", s_bold))
story.append(Spacer(1, 6))
deal_headers = [
    Paragraph("<b>Property</b>", S('deh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b>Market</b>",   S('deh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b>Loan Amt</b>", S('deh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>LTV</b>",      S('deh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Coupon</b>",   S('deh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Term</b>",     S('deh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Type</b>",     S('deh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
]
deal_rows = [[
    Paragraph(prop, s_body_sm), Paragraph(mkt, s_body_sm),
    Paragraph(loan, S('der', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)),
    Paragraph(ltv,  S('der2', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
    Paragraph(cpn,  S('der2', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
    Paragraph(trm,  S('der2', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
    Paragraph(typ,  s_body_sm),
] for prop,mkt,loan,ltv,cpn,trm,typ in NEW_DEALS]
deal_table = Table([deal_headers] + deal_rows, colWidths=[100, 75, 58, 40, 66, 42, 75])
deal_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, ALT_ROW]),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 6),
    ('RIGHTPADDING',  (0,0), (-1,-1), 6),
]))
story.append(deal_table)
story.append(Spacer(1, 14))
story.append(Paragraph(COMMENTARY['deals'], s_body))
story.append(PageBreak())

# ── PAGE 7: PORTFOLIO COMPOSITION ──────────────────────
story.append(Paragraph("Portfolio Composition", S('pt5', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)))
story.append(gold_rule())

geo_headers = [
    Paragraph("<b>Market</b>",          S('gh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b># Loans</b>",         S('gh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Deployed ($)</b>",    S('gh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>% Portfolio</b>",     S('gh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
]
geo_rows = [[
    Paragraph(mkt, s_body_sm),
    Paragraph(str(cnt), S('gr', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
    Paragraph(fmt_m(dep), S('gr2', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
    Paragraph(f"{pct:.1f}%", S('gr3', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)),
] for mkt,cnt,dep,pct in GEOGRAPHY]
geo_total = [
    Paragraph("<b>Total</b>", S('gt', fontName='Helvetica-Bold', fontSize=9)),
    Paragraph("<b>14</b>",    S('gt', fontName='Helvetica-Bold', fontSize=9, alignment=TA_RIGHT)),
    Paragraph("<b>$138.2M</b>", S('gt', fontName='Helvetica-Bold', fontSize=9, alignment=TA_RIGHT)),
    Paragraph("<b>100%</b>",  S('gt', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)),
]
geo_table = Table([geo_headers] + geo_rows + [geo_total], colWidths=[95, 50, 72, 58])
geo_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-2,-1), [WHITE, ALT_ROW]),
    ('BACKGROUND', (0,-1), (-1,-1), BG_LIGHT),
    ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('LINEABOVE',  (0,-1), (-1,-1), 1, HexColor('#CCCCCC')),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
]))

prop_headers2 = [
    Paragraph("<b>Property Type</b>",   S('pth2', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b>% Alloc</b>",         S('pth2', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Balance</b>",         S('pth2', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
]
prop_rows2 = [[
    Paragraph(pt, s_body_sm),
    Paragraph(f"{pct}%", S('pr3', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)),
    Paragraph(fmt_m(bal), S('pr4', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
] for pt,pct,bal in PROPERTY_TYPES]
prop_table2 = Table([prop_headers2] + prop_rows2, colWidths=[90, 52, 78])
prop_table2.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, ALT_ROW]),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
]))

comp_row = Table([
    [Paragraph("<b>By Geography (MSA / State)</b>", s_bold), Paragraph("<b>By Property Type</b>", s_bold)],
    [geo_table, prop_table2],
], colWidths=[280, 220])
comp_row.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(0,-1),16),('TOPPADDING',(0,0),(-1,0),0),('BOTTOMPADDING',(0,0),(-1,0),8)]))
story.append(comp_row)

story.append(Spacer(1, 16))
# Loan size distribution
story.append(gold_rule())
story.append(Paragraph("<b>Loan Size Concentration</b>", s_bold))
story.append(Spacer(1, 8))
size_headers = [
    Paragraph("<b>Loan Size Bucket</b>", S('lsh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)),
    Paragraph("<b># Loans</b>",         S('lsh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>Total ($)</b>",        S('lsh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
    Paragraph("<b>% of Portfolio</b>",   S('lsh', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE, alignment=TA_RIGHT)),
]
size_rows = [
    ["< $5 million",         "2",  "$7.9M",   "5.7%"],
    ["$5M – $10 million",    "6",  "$48.4M",  "35.0%"],
    ["$10M – $20 million",   "5",  "$69.1M",  "50.0%"],
    ["> $20 million",        "1",  "$12.8M",  "9.3%"],
]
size_data = [[size_headers[0], size_headers[1], size_headers[2], size_headers[3]]] + [
    [Paragraph(r[0], s_body_sm),
     Paragraph(r[1], S('sr', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT)),
     Paragraph(r[2], S('sr2', fontName='Helvetica-Bold', fontSize=9, textColor=GOLD, alignment=TA_RIGHT)),
     Paragraph(r[3], S('sr3', fontName='Helvetica', fontSize=9, alignment=TA_RIGHT))] for r in size_rows]
size_t = Table(size_data, colWidths=[140, 60, 80, 90])
size_t.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, ALT_ROW]),
    ('GRID',       (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
    ('LINEBELOW',  (0,0), (-1,0), 1, GOLD),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
]))
story.append(size_t)
story.append(PageBreak())

# ── PAGE 8: OUTLOOK ─────────────────────────────────────
story.append(Paragraph("Outlook", S('pt6', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)))
story.append(gold_rule())

pull_data = [[Paragraph(
    "<b>$22 million in pipeline loans expected to close in Q2 2026, increasing portfolio utilization to ~97% of called capital.</b>",
    S('pull', fontName='Helvetica-Bold', fontSize=13, leading=18, textColor=NAVY)
)]]
pull_t = Table(pull_data, colWidths=[460])
pull_t.setStyle(TableStyle([
    ('LEFTPADDING',   (0,0), (-1,-1), 18),
    ('RIGHTPADDING',  (0,0), (-1,-1), 18),
    ('TOPPADDING',    (0,0), (-1,-1), 14),
    ('BOTTOMPADDING', (0,0), (-1,-1), 14),
    ('LINEBEFORE',    (0,0), (0,-1), 4, GOLD),
    ('BACKGROUND',    (0,0), (-1,-1), BG_LIGHT),
    ('BOX',           (0,0), (-1,-1), 0.5, HexColor('#DDE3EC')),
]))
story.append(pull_t)
story.append(Spacer(1, 16))
story.append(Paragraph(COMMENTARY['outlook'], S('out', fontName='Helvetica', fontSize=11, leading=17, textColor=TEXT_DARK)))
story.append(Spacer(1, 20))

mkt_data = [[Paragraph(
    "<b>Market Context:</b> SOFR has stabilized in the 5.30–5.35% range, keeping floating-rate "
    "coupon income near cycle highs. Sunbelt multifamily fundamentals remain strong with vacancy "
    "rates below 5% in the fund's key markets. Industrial demand continues to benefit from "
    "e-commerce tailwinds and near-shoring activity.",
    S('mkt', fontName='Helvetica', fontSize=9, leading=14, textColor=TEXT_DARK)
)]]
mkt_t = Table(mkt_data, colWidths=[460])
mkt_t.setStyle(TableStyle([
    ('BACKGROUND',    (0,0), (-1,-1), BG_LIGHT),
    ('BOX',           (0,0), (-1,-1), 0.5, HexColor('#DDE3EC')),
    ('LEFTPADDING',   (0,0), (-1,-1), 14),
    ('RIGHTPADDING',  (0,0), (-1,-1), 14),
    ('TOPPADDING',    (0,0), (-1,-1), 10),
    ('BOTTOMPADDING', (0,0), (-1,-1), 10),
]))
story.append(mkt_t)
story.append(PageBreak())

# ── PAGE 9: APPENDIX ────────────────────────────────────
story.append(Paragraph("Appendix — Complete Data Extract", S('pt7', fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=NAVY)))
story.append(gold_rule())

def section_hdr(txt):
    return [Paragraph(f"<b>{txt}</b>", S('ah', fontName='Helvetica-Bold', fontSize=9, textColor=WHITE)), ""]

def arow(label, value, bold_val=False):
    vs = S('av', fontName='Helvetica-Bold' if bold_val else 'Helvetica', fontSize=9, leading=12, textColor=GOLD if bold_val else TEXT_DARK, alignment=TA_RIGHT)
    return [Paragraph(label, s_body_sm), Paragraph(value, vs)]

app_data = [
    section_hdr("FUND OVERVIEW"),
    arow("Total Fund NAV",                fmt_m(cur['nav']),            True),
    arow("Fund NAV (Prior Month)",        fmt_m(prev['nav'])),
    arow("NAV Change ($)",               fmt_m(cur['nav']-prev['nav'])),
    arow("NAV Change (%)",               fmt_pct((cur['nav']-prev['nav'])/prev['nav']*100)),
    arow("Gross Asset Value",             fmt_m(cur['gav'])),
    arow("Total Committed Capital",       fmt_m(cur['committed'])),
    arow("Total Called Capital",          fmt_m(cur['called'])),
    arow("Uncalled / Dry Powder",         fmt_m(cur['uncalled'])),
    arow("NAV per Unit",                  f"${cur['nav_per_unit']:.2f}"),
    section_hdr("LOAN PORTFOLIO"),
    arow("Total Loan Book Balance",       fmt_m(cur['loan_book']),      True),
    arow("Number of Active Loans",        str(cur['active_loans'])),
    arow("Wtd. Avg. LTV",                fmt_pct(cur['wtd_ltv'])),
    arow("Wtd. Avg. DSCR",               fmt_x(cur['wtd_dscr'])),
    arow("Wtd. Avg. Coupon",             fmt_pct(cur['wtd_coupon'])),
    arow("Avg. Remaining Term",           f"{cur['avg_term']} months"),
    arow("Portfolio Occupancy",           fmt_pct(cur['occupancy'])),
    arow("Performing Loans",              "14 / 100%"),
    section_hdr("DISTRIBUTIONS & RETURNS"),
    arow("Monthly Distribution / Unit",  f"${cur['monthly_dist']:.3f}",  True),
    arow("Annualized Distribution Yield", fmt_pct(cur['ann_yield'])),
    arow("Preferred Return Hurdle",       fmt_pct(cur['pref_return'])),
    arow("Preferred Return Status",       "Current"),
    arow("Cumulative Distributions Paid", fmt_m(cur['cumulative_dist'])),
    arow("Net IRR (Since Inception)",     fmt_pct(cur['net_irr'])),
    arow("MOIC (Net)",                   fmt_x(cur['moic'])),
    arow("TVPI (Net)",                   fmt_x(cur['tvpi'])),
    section_hdr("DEAL ACTIVITY"),
    arow("New Originations (Count)",      str(cur['new_orig_count'])),
    arow("New Originations (Total $)",    fmt_m(cur['new_orig_total'])),
    arow("Repayments / Exits (Count)",    str(cur['repay_count'])),
    arow("Repayments / Exits (Total $)",  fmt_m(cur['repay_total'])),
    arow("Pipeline (Count)",              str(cur['pipeline_count'])),
    arow("Pipeline (Total $)",            fmt_m(cur['pipeline_total'])),
]

app_table = Table(app_data, colWidths=[280, 175])
app_ts = TableStyle([
    ('FONTSIZE',   (0,0), (-1,-1), 9),
    ('TOPPADDING',    (0,0), (-1,-1), 5),
    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ('RIGHTPADDING',  (0,0), (-1,-1), 8),
    ('LINEBELOW',  (0,0), (-1,-1), 0.3, HexColor('#DDDDDD')),
])
# Section header rows
hdr_idxs = [0, 10, 19, 28]
for idx in hdr_idxs:
    app_ts.add('BACKGROUND', (0,idx), (-1,idx), NAVY)
    app_ts.add('LINEBELOW',  (0,idx), (-1,idx), 1, GOLD)
    app_ts.add('SPAN',       (0,idx), (-1,idx))
# Alt rows
for i in range(len(app_data)):
    if i not in hdr_idxs:
        if (i % 2) == 0:
            app_ts.add('BACKGROUND', (0,i), (-1,i), ALT_ROW)
        else:
            app_ts.add('BACKGROUND', (0,i), (-1,i), WHITE)
app_table.setStyle(app_ts)
story.append(app_table)

story.append(Spacer(1, 16))
story.append(Paragraph(
    "<i>Data source: Test mock data (no source file — demo run). "
    "Figures are unaudited and for illustrative purposes only. "
    f"Report prepared by {MANAGER} | {CONTACT_EMAIL} | {CONTACT_WEB}</i>",
    s_muted
))

# ══════════════════════════════════════════════════════
# BUILD  (cover gets its own canvas)
# ══════════════════════════════════════════════════════
def on_first_page(c, doc):
    draw_cover(c, doc)

def on_later_pages(c, doc):
    if doc.page > 1:
        header_footer(c, doc)

doc.build(story, onFirstPage=on_first_page, onLaterPages=on_later_pages)

from pypdf import PdfReader
r = PdfReader(OUT_PATH)
print(f"PDF created: {len(r.pages)} pages → {OUT_PATH}")

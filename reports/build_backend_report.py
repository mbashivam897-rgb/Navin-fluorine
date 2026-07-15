"""
Build BACKEND/WORKING-PAPERS Word report: full research trail behind the
two-idea IC memo (Karur Vysya Bank, CCL Products) - Indian equities, Jul 2026.
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x2D, 0x50)
DARKGREY = RGBColor(0x33, 0x33, 0x33)
ACCENT = RGBColor(0x8A, 0x1F, 0x1F)
GREEN = RGBColor(0x1E, 0x6B, 0x3A)
GREY = RGBColor(0x66, 0x66, 0x66)

doc = Document()

normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = Pt(10)
normal.font.color.rgb = DARKGREY
normal.paragraph_format.space_after = Pt(5)
normal.paragraph_format.line_spacing = 1.12

for s in doc.sections:
    s.top_margin = Cm(1.6)
    s.bottom_margin = Cm(1.6)
    s.left_margin = Cm(1.9)
    s.right_margin = Cm(1.9)

def set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_heading(text, level=1, color=NAVY, size=None, space_before=14, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = True
    r.font.color.rgb = color
    if size:
        r.font.size = Pt(size)
    else:
        sizes = {1: 17, 2: 13, 3: 11.5}
        r.font.size = Pt(sizes.get(level, 11.5))
    return p

def add_body(text, bold=False, italic=False, size=10, color=DARKGREY, space_after=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p

def add_bullet(text, bold_lead=None, size=10):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    if bold_lead:
        r1 = p.add_run(bold_lead)
        r1.bold = True
        r1.font.size = Pt(size)
        r2 = p.add_run(text)
        r2.font.size = Pt(size)
    else:
        r = p.add_run(text)
        r.font.size = Pt(size)
    return p

def add_hr():
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '1F2D50')
    pbdr.append(bottom)
    pPr.append(pbdr)

def style_table(table, header_color='1F2D50'):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(2)
                for r in p.runs:
                    r.font.size = Pt(8.5)
            if i == 0:
                set_cell_shading(cell, header_color)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                        r.bold = True
            else:
                if i % 2 == 0:
                    set_cell_shading(cell, 'F2F2F2')

def verdict_tag(text, kind='reject'):
    """kind: reject, watch, interest, clear"""
    colors = {'reject': ACCENT, 'watch': RGBColor(0xB8, 0x86, 0x0B), 'interest': GREEN, 'clear': GREEN}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f"VERDICT: {text}")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = colors.get(kind, ACCENT)
    return p

def company_brief(name, ticker, sector, stats, findings, verdict_text, verdict_kind, sources):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    r = p.add_run(f"{name} ")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = NAVY
    r2 = p.add_run(f"({ticker}) \u2014 {sector}")
    r2.italic = True
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = GREY
    add_body(stats, italic=True, size=9, color=GREY, space_after=4)
    for lead, rest in findings:
        add_bullet(rest, bold_lead=lead, size=9.5)
    verdict_tag(verdict_text, verdict_kind)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(8)
    r = sp.add_run("Sources: ")
    r.bold = True
    r.italic = True
    r.font.size = Pt(8.5)
    r2 = sp.add_run(sources)
    r2.italic = True
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = GREY

# =========================================================
# COVER
# =========================================================
title = doc.add_paragraph()
title.paragraph_format.space_before = Pt(50)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run("RESEARCH WORKING PAPERS")
r.bold = True
r.font.size = Pt(24)
r.font.color.rgb = NAVY

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.paragraph_format.space_after = Pt(4)
r = sub.add_run("Full Research Trail \u2014 Screening, Diligence & Rejection Rationale")
r.font.size = Pt(14)
r.font.color.rgb = ACCENT
r.italic = True

sub2 = doc.add_paragraph()
sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub2.paragraph_format.space_after = Pt(30)
r = sub2.add_run("Supporting documentation for: \u201cTwo High-Conviction Long Ideas \u2014 Indian Equities\u201d")
r.font.size = Pt(11)
r.font.color.rgb = DARKGREY

meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = meta.add_run(
    "This document is the backend research record: every company screened, every data point "
    "gathered, every rejection reasoned through, before the final two names were selected.\n\n"
    "Coverage: 16 companies screened \u2192 6 rejected outright \u2192 10 ranked \u2192 2 taken to deep diligence "
    "\u2192 1 dropped after deep diligence \u2192 1 promoted as replacement \u2192 2 finalists underwritten.\n\n"
    "Date: July 2026"
)
r.font.size = Pt(10)
r.font.color.rgb = DARKGREY

doc.add_page_break()

# =========================================================
# TABLE OF CONTENTS (manual)
# =========================================================
add_heading("Contents", 1)
add_hr()
toc = [
    "1. Research Process Overview",
    "2. Initial Screening Universe \u2014 16 Candidates & Selection Rationale",
    "3. Screening Briefs \u2014 Batch A: Craftsman Automation, Karur Vysya Bank, Anand Rathi Wealth, Neuland Laboratories",
    "4. Screening Briefs \u2014 Batch B: Kaynes Technology, Metropolis Healthcare, CCL Products, Sansera Engineering",
    "5. Screening Briefs \u2014 Batch C: NOCIL, Praj Industries, Techno Electric & Engineering, Concord Biotech",
    "6. Screening Briefs \u2014 Batch D: Greenpanel Industries, Triveni Turbine, Global Health (Medanta), KIMS Hospitals",
    "7. Deep Diligence \u2014 Karur Vysya Bank (Finalist)",
    "8. Deep Diligence \u2014 Techno Electric & Engineering (Rejected After Deep Diligence)",
    "9. Deep Diligence \u2014 CCL Products India (Promoted Finalist)",
    "10. Valuation Data-Gathering Notes",
    "11. Final Ranking Methodology & Score Rationale",
    "12. Open Items / Unverified Claims Register",
]
for t in toc:
    add_bullet(t, size=10.5)

doc.add_page_break()

# =========================================================
# SECTION 1: PROCESS OVERVIEW
# =========================================================
add_heading("1. Research Process Overview", 1)
add_hr()
add_body(
    "Scope confirmed with the requester: Indian-listed equities, no market-cap floor specified beyond avoiding "
    "both mega-cap consensus names and obscure micro-caps, no sector exclusions, 3\u20135 year holding period, "
    "public-source research (no Bloomberg/FactSet terminal access) with explicit flagging of any claim resting "
    "on secondary/aggregator sources rather than primary filings."
)
add_body("Process followed:", bold=True, space_after=3)
steps = [
    ("Step 1 \u2014 Universe construction: ", "16 names selected across 10 sectors, deliberately spanning specialty "
     "chemicals, diagnostics/hospitals, EMS, capital goods/engineering, auto ancillary, building materials, "
     "wealth management, pharma CDMO, and mid-size banking, to avoid both popular large-caps and illiquid obscurities."),
    ("Step 2 \u2014 Parallel screening: ", "4 sub-agent research passes (4 companies each) against screener.in, company "
     "IR pages, Quartr/stockanalysis.com transcript summaries, and financial media, covering valuation, 3\u20135yr "
     "financial trends, debt, credit ratings, and explicit red-flag searches."),
    ("Step 3 \u2014 Screening-stage rejection: ", "6 of 16 names rejected outright on hard evidence (governance, "
     "accounting quality, or guidance misses) before any ranking was attempted."),
    ("Step 4 \u2014 Ranking: ", "Remaining 10 names scored and ranked on conviction, expected return, risk, quality, "
     "and valuation attractiveness."),
    ("Step 5 \u2014 Deep diligence: ", "Top 2 (Karur Vysya Bank, Techno Electric & Engineering) taken to deep "
     "primary-source diligence \u2014 multi-year cash flow analysis, management track record verification, credit "
     "rating rationale, related-party transaction checks."),
    ("Step 6 \u2014 Mid-course rejection: ", "Techno Electric & Engineering dropped after deep diligence revealed "
     "negative operating cash flow in 2 of the last 3 fiscal years despite rising reported profit, plus an "
     "unresolved counterparty exposure (Sankhya Financial Services NCDs) that could not be fully cleared "
     "against primary filings within the search session."),
    ("Step 7 \u2014 Replacement & re-diligence: ", "CCL Products (India) promoted from the ranked list and taken "
     "through the same deep-diligence bar."),
    ("Step 8 \u2014 Valuation build: ", "Live price/multiple data gathered for both finalists plus a peer bank "
     "comparison set, used to build bear/base/bull sensitivity tables."),
]
for lead, rest in steps:
    add_bullet(rest, bold_lead=lead, size=10)

add_body(
    "Evidentiary note carried through the entire process: research relied on public aggregator data "
    "(Screener.in, stockanalysis.com, Economic Times, Moneycontrol, Quartr transcript summaries, company IR pages) "
    "rather than direct terminal access. Multiple attempts to fetch primary PDFs (ICRA rating rationale documents, "
    "full annual report notes, NSE filing archives) returned inaccessible content, 404s, or non-primary mirrors; "
    "these gaps are logged explicitly in Section 12 rather than filled with assumed figures.",
    italic=True, size=9.5
)

doc.add_page_break()

# =========================================================
# SECTION 2: UNIVERSE
# =========================================================
add_heading("2. Initial Screening Universe \u2014 16 Candidates", 1)
add_hr()
universe = [
    ["#", "Company", "Sector", "Screening Rationale"],
    ["1", "Craftsman Automation", "Auto ancillary / industrial", "Diversified into aluminum/powertrain beyond autos; margin mix shift"],
    ["2", "Karur Vysya Bank", "Mid-size private bank", "Turnaround story, asset quality repair, trades below peers"],
    ["3", "Anand Rathi Wealth", "Wealth management", "Asset-light, high ROE, financialization-of-savings tailwind"],
    ["4", "Neuland Laboratories", "Pharma CDMO", "Capacity expansion, US FDA track record, mix shift to CDMO"],
    ["5", "Greenpanel Industries", "Building materials (MDF)", "Anti-dumping duty tailwind, import substitution"],
    ["6", "Triveni Turbine", "Capital goods (turbines)", "Export capacity expansion, aftermarket revenue mix"],
    ["7", "Global Health (Medanta)", "Hospitals", "Bed capacity expansion, occupancy ramp-up"],
    ["8", "KIMS Hospitals", "Hospitals", "South India expansion, brownfield ROCE"],
    ["9", "Kaynes Technology", "EMS / electronics", "Order book scaling, OSAT/semiconductor optionality"],
    ["10", "Metropolis Healthcare", "Diagnostics", "Margin recovery post-COVID normalization"],
    ["11", "CCL Products (India)", "Instant coffee (B2B export)", "Capacity expansion, brand business optionality"],
    ["12", "Sansera Engineering", "Auto ancillary / aerospace", "Diversification away from ICE, aerospace ramp"],
    ["13", "NOCIL", "Rubber/tire chemicals", "Import substitution, capacity utilization inflection"],
    ["14", "Praj Industries", "Ethanol/biofuel EPC", "Government biofuel blending policy tailwind"],
    ["15", "Techno Electric & Engineering", "T&D EPC + renewables", "Grid capex cycle, asset monetization"],
    ["16", "Concord Biotech", "Fermentation API/CDMO", "Niche molecule leadership, high ROCE"],
]
table = doc.add_table(rows=len(universe), cols=len(universe[0]))
table.style = 'Table Grid'
for i, row in enumerate(universe):
    for j, val in enumerate(row):
        table.cell(i, j).text = val
style_table(table)

doc.add_page_break()

# =========================================================
# SECTION 3: BATCH A
# =========================================================
add_heading("3. Screening Briefs \u2014 Batch A", 1)
add_hr()

company_brief(
    "Craftsman Automation Ltd", "NSE: CRAFTSMAN", "Auto ancillary / industrial",
    "Mkt cap \u20b924,158 Cr | Price \u20b99,230 | P/E 63.2x | ROCE 13.9% | ROE 12.5% (screener.in, 15 Jul 2026)",
    [
        ("Revenue/margin: ", "Revenue nearly doubled FY25\u2192FY26 (\u20b95,690 Cr \u2192 \u20b98,069 Cr) largely via consolidation of "
         "Sunbeam Lightweighting (acquired Oct 2024, \u20b9376 Cr) and DR Axion. OPM compressed from 20% (FY24) to 15% "
         "(FY25-26) as lower-margin acquired businesses diluted the mix."),
        ("ROCE trend: ", "Fell from 20-21% (FY22-23) to 12-14% (FY24-26) \u2014 capex/M&A-driven, not organic."),
        ("Debt: ", "Net debt/equity ~0.66; debt/equity ~111%; interest cost up from \u20b958 Cr (Q4FY25) to \u20b986 Cr (Q4FY26)."),
        ("Credit rating: ", "CRISIL AA-/Positive (29 Jun 2026); Sunbeam entity still on \u201cRating Watch \u2013 Positive\u201d pending merger into DR Axion."),
        ("Red flags: ", "Sunbeam remains loss-making (\u20b940 Cr loss on \u20b91,397 Cr revenue, FY26); promoter sold ~2.01% stake "
         "(~\u20b9484-486 Cr block deal, mid-2026) on top of a recent \u20b92,000 Cr QIP; promoter holding fell from 54.99% "
         "(Sep-23) to 42.41% (Jun-26); complex multi-entity group structure."),
    ],
    "Lean toward reject at current price. Growth is real but debt/acquisition-fueled, not margin-led. At 63x "
    "P/E for ~14% ROCE, valuation already assumes a Sunbeam turnaround not yet visible in numbers. Heavy "
    "promoter selling is a caution flag. Revisit if Sunbeam margins actually turn positive.",
    'reject',
    "Screener.in (15 Jul 2026); CRISIL rating action (29 Jun 2026); FY26 results filings; block deal news coverage, mid-2026."
)

company_brief(
    "Karur Vysya Bank Ltd", "NSE: KARURVYSYA", "Mid-size private bank",
    "Mkt cap \u20b929,488 Cr | Price \u20b9305 | P/E 11.9x | Book value \u20b9146 | ROE 19.1% (screener.in, 15 Jul 2026)",
    [
        ("PAT trend: ", "FY23 \u20b91,106 Cr \u2192 FY24 \u20b91,605 Cr \u2192 FY25 \u20b91,942 Cr \u2192 FY26 \u20b92,510 Cr (+29% YoY, record high)."),
        ("Returns trend: ", "ROE climbed steadily FY22 (9%) \u2192 FY26 (19%). NIM ~3.9-3.99%, within FY26 guidance. ROA hit "
         "1.93% FY26 (Q4 ROA 2.10%), meeting management's >1.85% target."),
        ("Asset quality: ", "GNPA 0.75%, NNPA 0.19%, PCR 96.56% as of Mar 2026 \u2014 among the cleanest in mid-cap private banking."),
        ("Recent quarters: ", "Q3 FY26 PAT \u20b9690 Cr (+25% YoY); Q4 FY26 PAT \u20b9725 Cr (+41% YoY). Advances crossed \u20b91 lakh "
         "crore in Q1 FY27 (+17.12% YoY)."),
        ("Credit rating: ", "ICRA reaffirmed issuer rating AA (Stable), 29 Jun 2026; A1+ for enhanced \u20b912,000 Cr CD programme. "
         "Capital adequacy 18.76%."),
        ("Red flags: ", "CASA ratio declined from ~48% (Jun-23) to ~39% (Mar-26), implying cost-of-funds pressure; "
         "contingent liabilities \u20b914,661 Cr; promoter holding structurally low (2.07%); agencies held AA/Stable "
         "rather than upgrading despite improving metrics."),
    ],
    "Most attractive risk-adjusted candidate of the batch. Genuine, credible compounding \u2014 improving ROA/NIM, "
    "best-in-class asset quality, reasonable valuation. Fewest red flags of the four screened in this batch.",
    'interest',
    "Screener.in (15 Jul 2026); ICRA rating action (29 Jun 2026); ET, Livemint, CNBC-TV18 coverage 2022-2026; Quartr FY26 summary."
)

company_brief(
    "Anand Rathi Wealth Ltd", "NSE: ANANDRATHI", "Wealth management",
    "Mkt cap \u20b934,555 Cr | Price \u20b92,082 | P/E 74.3x | ROCE 59.2% | ROE 47.3% (pre-dates Q1 FY27 margin miss)",
    [
        ("FY26 results: ", "Revenue \u20b91,253 Cr (+28%), PAT \u20b9397 Cr (+32%). ROCE rising: 48% (FY20)\u219259% (FY26). "
         "BUT full-year OPM already fell from 45% (FY25) to 39% (FY26) \u2014 margin compression predates Q1FY27. "
         "AUM \u20b91,06,300 Cr as of Jun-2026 (+21% YoY)."),
        ("Q1 FY27 shock: ", "PAT \u20b9163 Cr (+74% YoY, boosted by investment/other income), but EBITDA fell 15% YoY to "
         "\u20b9109 Cr and margin collapsed to 33.7% from 46.6% \u2014 driven by a 53% YoY jump in employee costs including "
         "a one-time ESOP charge. Motilal Oswal downgraded to Sell (target \u20b91,700, ~21% downside), citing a 27% "
         "EBITDA miss and cutting FY27 margin estimates by 563bps."),
        ("Red flags: ", "Trades at 34.6x book value; promoter stake fell from 48.7% (Sep-23) to 41.4% (Jun-26), "
         "including a \u20b9500 Cr open-market sale in Jul 2026; a promoter pledge disclosure error was reported and "
         "later \u201ccorrected\u201d involving a 5,60,000-share pledge, and pledge levels reportedly spiked to 11.42% "
         "in Mar-2026 from ~0% in Dec-2025; PAT growth increasingly reliant on \u201cother income\u201d (\u20b9205 Cr flagged), "
         "not core distribution economics."),
    ],
    "Reject at current price. Excellent underlying franchise but priced for perfection at 74x P/E while margins "
    "are now visibly cracking and a promoter pledge irregularity is unexplained. Market has not fully repriced "
    "margin/governance risk.",
    'reject',
    "Screener.in; Motilal Oswal research note (2026); Q1 FY27 results filing; pledge disclosure news coverage, 2026."
)

company_brief(
    "Neuland Laboratories Ltd", "NSE: NEULANDLAB", "Pharma CDMO",
    "Mkt cap \u20b924,525 Cr | Price \u20b919,114 | P/E 67.9x | Book value \u20b91,461 | ROCE 26.5% | ROE 21.2%",
    [
        ("Revenue/PAT: ", "Revenue: FY24 \u20b91,559 Cr \u2192 FY25 \u20b91,477 Cr (dip) \u2192 FY26 \u20b92,023 Cr (+37%). PAT: FY24 \u20b9300 Cr \u2192 "
         "FY25 \u20b9260 Cr \u2192 FY26 \u20b9364 Cr (+40%). EBITDA margin expanded materially FY26 to 29.4% from 22.9% FY25 \u2014 "
         "genuine full-year improvement, driven by CMS (custom manufacturing for innovator pharma) mix shift."),
        ("Quarterly volatility: ", "Severe. Q3 FY26 revenue fell 10.5% YoY/14.8% QoQ with PAT dropping to \u20b941 Cr from "
         "\u20b997 Cr prior quarter (\u201cprofit drops sharply\u201d headlines, Feb 2026). Q4 FY26 then surged: PAT \u20b9213 Cr, "
         "gross margin 62.1%, EBITDA margin ~40%."),
        ("Credit rating: ", "India Ratings affirmed IND A+/Positive and IND A1 (~13 Jul 2026); Fitch and CRISIL also "
         "improving through 2025-26."),
        ("Balance sheet: ", "Low interest burden (~\u20b924 Cr FY26); FCF turned negative (-\u20b950 Cr FY26) on heavy capex "
         "(CWIP jumped to \u20b9211 Cr); cash conversion cycle stretched (~214 days)."),
        ("Red flags: ", "Extreme quarter-to-quarter lumpiness tied to large CMS customer order timing (implicit "
         "customer concentration risk within CMS book); promoter holding down 3.5% over 3 years; trades at 13.1x "
         "book; negative FCF during expansion phase raises execution risk."),
    ],
    "Not a reject, but not high-conviction on current evidence. Most plausible genuine inflection story of the "
    "batch \u2014 margin expansion real at annual level, rating agencies validating improving credit quality \u2014 but "
    "extreme quarterly swings mean the market could be over-extrapolating one strong quarter (Q4FY26) onto a "
    "structurally lumpy order book. Worth deeper diligence on CMS pipeline/customer concentration.",
    'watch',
    "Screener.in; India Ratings affirmation (~13 Jul 2026); Q3/Q4 FY26 results coverage, Feb-May 2026."
)

doc.add_page_break()

# =========================================================
# SECTION 4: BATCH B
# =========================================================
add_heading("4. Screening Briefs \u2014 Batch B", 1)
add_hr()

company_brief(
    "Kaynes Technology India Ltd", "NSE: KAYNES", "EMS / electronics",
    "Mkt cap \u20b922,374 Cr | P/E 87.4x | ROCE 10.6% | ROE 7.2%",
    [
        ("FY26 results: ", "Revenue \u20b93,626 Cr (+33%), but Q4 FY26 PAT fell 22% YoY causing an 18-20% stock crash "
         "and a JPMorgan downgrade."),
        ("Governance red flag: ", "Company admitted an undisclosed related-party transaction with subsidiary "
         "Iskraemeco in standalone FY24 financials; Kotak flagged inconsistent RPT disclosures and unclear "
         "acquisition accounting (Dec 2025)."),
        ("Working capital blowout: ", "Smart-metering segment has a severe receivables problem (\u20b91,365 Cr "
         "receivables vs \u20b9971 Cr revenue; working capital days blew out from 267 to 486)."),
        ("Market reaction: ", "Stock down ~50% from Oct 2025 high."),
        ("Credit rating: ", "CRISIL A/Stable."),
    ],
    "Reject/avoid. Real governance and accounting-quality issues, not just sentiment.",
    'reject',
    "Kotak research note (Dec 2025); JPMorgan downgrade coverage; Q4 FY26 results filing; screener.in."
)

company_brief(
    "Metropolis Healthcare Ltd", "NSE: METROPOLIS", "Diagnostics",
    "Mkt cap \u20b911,858 Cr | P/E ~73x | ROCE 15.9% | ROE 12.4% | P/B 8.5x",
    [
        ("Growth quality: ", "Screener explicitly flags \u201cpoor sales growth of 10.8% over 5 years.\u201d"),
        ("Returns trend: ", "ROE structurally declined from 19% (10yr) to 12% (3yr/last year); ROCE collapsed from "
         "50% (FY17) to 13% (FY24) before recovering to 16% (FY26), coinciding with debt-funded M&A."),
        ("Acquisition dependence: ", "CORE Diagnostics acquired Dec 2024 for \u20b9247 Cr, currently margin-dilutive."),
        ("Recent improvement: ", "Q4 FY26 PAT +75% YoY \u2014 real, but acquisition-dependent."),
    ],
    "Cautious / lean reject. Margin recovery narrative is real but growth is acquisition-dependent and already "
    "priced at ~70x earnings.",
    'reject',
    "Screener.in; acquisition announcement filings (Dec 2024); Q4 FY26 results coverage."
)

company_brief(
    "CCL Products (India) Ltd", "NSE: CCL", "Instant coffee (B2B export)",
    "Mkt cap \u20b916,338 Cr | P/E 42.1x | ROCE 15.8% | ROE 18%",
    [
        ("FY26 results: ", "Revenue surged 43% to ~\u20b94,460 Cr, PAT +25% to \u20b9388 Cr, debt cut by ~\u20b9700 Cr."),
        ("Caveat noted at screening stage: ", "OPM% actually compressed from ~18% to ~16% during the revenue "
         "surge \u2014 growth isn't pure operating leverage on a headline basis (later resolved in deep diligence, "
         "Section 9, as a EBITDA/kg vs. EBITDA% optical effect)."),
        ("Governance: ", "No promoter pledging, RPT, or governance red flags found in search."),
        ("Customer concentration: ", "Structurally likely given private-label model but exact figures unverified "
         "at screening stage."),
    ],
    "Most credible candidate of the batch. Cleanest story, real growth + deleveraging, no red flags found, "
    "though margin compression needs explaining (see Section 9 for resolution).",
    'interest',
    "Screener.in; FY26 results filing; Q3/Q4 FY26 transcript summaries (Quartr, stockanalysis.com)."
)

company_brief(
    "Sansera Engineering Ltd", "NSE: SANSERA", "Auto ancillary / aerospace",
    "Mkt cap ~\u20b920,325 Cr (consolidated) | P/E 60.4x | ROCE 14.1% | ROE 11.5% (flagged \u201clow ROE over last 3 years\u201d)",
    [
        ("FY26 results: ", "Revenue +16%, PAT +51% (aided by low prior-year base and one-time Q3 charge)."),
        ("Deleveraging: ", "Real \u2014 interest cost nearly halved YoY."),
        ("Diversification: ", "Aerospace/Defense/Semiconductor (ADS) segment and FY30 \u20b910,000 Cr revenue target."),
        ("Red flag: ", "High customer concentration risk \u2014 supplies nearly every major Indian 2-wheeler OEM."),
        ("Valuation vs. fundamentals: ", "Stock up ~137% in 1 year; ROE/ROCE have NOT actually inflected despite "
         "the growth narrative."),
    ],
    "Priced for perfection on a good story, not an obvious mispricing yet.",
    'reject',
    "Screener.in; FY26 results filing; ADS segment investor presentation commentary."
)

doc.add_page_break()

# =========================================================
# SECTION 5: BATCH C
# =========================================================
add_heading("5. Screening Briefs \u2014 Batch C", 1)
add_hr()

company_brief(
    "NOCIL Ltd", "NSE: NOCIL", "Rubber/tire chemicals",
    "Mkt cap \u20b92,784 Cr | P/E 60.5x",
    [
        ("ROCE collapse: ", "26% \u2192 4% over ~8 years."),
        ("Revenue/margin: ", "Revenue declining (-6% TTM), OPM halved from 28% to 8%."),
        ("Earnings quality: ", "Earnings propped up by other income; margin erosion likely structural "
         "(import/pricing pressure)."),
    ],
    "Reject \u2014 value trap, no visible inflection.",
    'reject',
    "Screener.in; multi-year results filings."
)

company_brief(
    "Praj Industries Ltd", "NSE: PRAJIND", "Ethanol/biofuel EPC",
    "Mkt cap \u20b96,486 Cr | P/E ~335x (earnings near-zero)",
    [
        ("Guidance miss: ", "Management guided ~10% growth but delivered a revenue decline \u2014 confirmed guidance "
         "miss, citing domestic ethanol overcapacity as a structural headwind."),
        ("Profit collapse: ", "PAT down ~90% YoY in FY26, OPM crashed 11%\u21925%."),
    ],
    "Reject \u2014 guidance miss and margin collapse, not a mispricing.",
    'reject',
    "Screener.in; FY26 results filing; management commentary on ethanol overcapacity."
)

company_brief(
    "Techno Electric & Engineering Co Ltd", "NSE: TECHNOE", "T&D EPC + renewables",
    "Mkt cap \u20b912,201 Cr | P/E 27.2x",
    [
        ("Growth profile: ", "Best of the batch on headline growth \u2014 revenue 3yr CAGR 58%, PAT +19% in FY26. "
         "Large order book (\u20b99,567 Cr, ~2.9x revenue), debt-free."),
        ("Structural tailwind: ", "Data-center/smart-metering pivot is a real structural tailwind."),
        ("Unresolved red flag (screening stage): ", "A May 2026 disclosure about \u20b980 Cr received against "
         "\u201cSankhya Financial Services NCDs\u201d with \u201cresidual balance resolution under discussion\u201d \u2014 implies a "
         "stressed receivable parked in the \u201cinvestments\u201d line; nature (related-party or not) unconfirmed at "
         "this stage."),
        ("Working capital: ", "Days ballooned 222\u2192430."),
    ],
    "Tentative interest \u2014 best candidate of the batch on headline metrics, but the Sankhya NCD exposure needs "
    "direct annual-report/related-party-note verification before sizing a position. (See Section 8: rejected "
    "after deep diligence.)",
    'watch',
    "Screener.in; BSE/NSE corporate disclosure (~28 May 2026); FY26 results filing."
)

company_brief(
    "Concord Biotech Ltd", "NSE: CONCORDBIO", "Fermentation API/CDMO",
    "Mkt cap \u20b914,019 Cr | P/E 54-58x",
    [
        ("Balance sheet quality: ", "High-quality zero-debt fermentation-API franchise."),
        ("FY26 results: ", "Revenue -12%, PAT -30% YoY, driven by delayed customer procurement, tariff uncertainty, "
         "and injectables-facility ramp costs (adjusted EBITDA margin 38.8% ex-startup costs vs 24.6% reported, "
         "suggesting some of the hit is transitional)."),
        ("Working capital: ", "Debtor days elevated at 159."),
    ],
    "Watch, not yet buy. Good business, but valuation doesn't yet price in enough caution given consecutive "
    "quarters of decline.",
    'watch',
    "Screener.in; FY26 results filing; management commentary on tariff uncertainty and facility ramp."
)

doc.add_page_break()

# =========================================================
# SECTION 6: BATCH D
# =========================================================
add_heading("6. Screening Briefs \u2014 Batch D", 1)
add_hr()

company_brief(
    "Triveni Turbine Ltd", "NSE: TRITURBINE", "Capital goods (turbines)",
    "Price \u20b9607-639 | Mkt cap ~\u20b919,300-20,320 Cr | P/E ~55-62x | EV/EBITDA ~42x",
    [
        ("Balance sheet: ", "Near debt-free (net cash \u20b9545 Cr)."),
        ("FY26 results: ", "First PAT decline (-2-3% YoY) despite revenue +9%, with rising receivable days "
         "(49\u219284) and lumpy order booking."),
        ("Governance: ", "No pledging/RPT flags found. No verified credit rating found."),
    ],
    "Strong execution but priced as a flawless compounder; FY26 cracks (margin/PAT dip, working capital "
    "deterioration) not yet reflected in price \u2014 caution over conviction.",
    'reject',
    "Screener.in; FY26 results filing."
)

company_brief(
    "Greenpanel Industries Ltd", "NSE: GREENPANEL", "Building materials (MDF)",
    "Price \u20b9202 | Mkt cap \u20b92,479 Cr | EV/EBITDA ~14.9x (FY26 net loss year, PAT -\u20b929.1 Cr)",
    [
        ("Margin collapse: ", "EBITDA margin collapsed from 25-30% (FY21-23) to ~9-11% (FY25-26) on low utilization "
         "(~60%) and chemical-cost spikes."),
        ("Returns: ", "ROCE fell from ~27-30% to 12.9%."),
        ("Credit rating: ", "ICRA A+/Negative outlook (Jul 2026)."),
        ("Guidance: ", "Clear miss vs. FY26 targets."),
    ],
    "Classic cyclical-trough/operating-leverage story, not a reject, but bull case unproven \u2014 wait for 2 "
    "quarters of >70% utilization and >15% margin before treating as a real inflection.",
    'watch',
    "Screener.in; ICRA rating action (Jul 2026); FY26 results filing."
)

company_brief(
    "Global Health Ltd (Medanta)", "NSE: MEDANTA", "Hospitals",
    "Price \u20b91,338 | Mkt cap \u20b935,994 Cr | P/E 64.6x | P/B 9.08x",
    [
        ("Revenue growth: ", "Steady \u2014 \u20b92,167 Cr FY22 \u2192 \u20b94,410 Cr FY26 (+19.4% CAGR)."),
        ("Returns trend: ", "ROCE rolled over from 20% to 17% as Noida/Ranchi capex dilutes returns; Noida "
         "breakeven guided H2 FY27."),
        ("Balance sheet: ", "Net cash historically (~\u20b9812 Cr Mar-25), exact FY26 net debt unverified."),
        ("Governance overhang: ", "Legacy 2020 FIR/ED case against founder-chairman (status unverified)."),
    ],
    "Quality compounder priced for perfection, thin margin of safety at current multiple.",
    'reject',
    "Screener.in; FY26 results filing; news coverage of legacy FIR/ED matter (status unverified)."
)

company_brief(
    "KIMS Hospitals (Krishna Institute of Medical Sciences)", "NSE: KIMS", "Hospitals",
    "Price \u20b9808-810 | Mkt cap ~\u20b933,947 Cr | P/E ~137-141x | EV/EBITDA ~34x \u2014 richest of the batch",
    [
        ("Profit vs. revenue: ", "FY26 PAT fell ~38-42% YoY despite 28.6% revenue growth; margin compressed "
         "26.6%\u2192~21%."),
        ("Returns collapse: ", "ROCE now ~9.3% (down from 16.4% 10-yr avg ROE)."),
        ("Funding quality: ", "Net debt rose to ~\u20b93,100 Cr funded partly via new QIP (\u20b91,500 Cr) and promoter "
         "warrants (\u20b9600 Cr) rather than internal cash generation."),
        ("Red flags: ", "Promoter pledging reportedly rose 3.7%\u219210.3% (single-source, uncorroborated), promoter "
         "holding declining with further dilution underway, institutional dissent at Jul-26 EGM, a debt-disclosure "
         "correction after NSE query."),
    ],
    "Most red flags of the batch relative to its rich multiple \u2014 a \u201cstory stock\u201d not yet showing the promised "
    "inflection; warrants the most skepticism.",
    'reject',
    "Screener.in; NSE query/disclosure correction coverage; EGM coverage (Jul 2026); FY26 results filing."
)

add_body(
    "Cross-batch takeaway from initial screening: none of the sixteen showed a clean, already-realized "
    "mispricing purely on headline numbers \u2014 several bull cases depended on future inflections not yet visible "
    "in trailing results, while trading at premium multiples. Karur Vysya Bank and Techno Electric & Engineering "
    "screened as the two strongest risk-adjusted candidates and were carried into deep diligence (Sections 7-9).",
    italic=True, size=9.5
)

doc.add_page_break()

# =========================================================
# SECTION 7: DEEP DILIGENCE - KVB
# =========================================================
add_heading("7. Deep Diligence \u2014 Karur Vysya Bank (Finalist)", 1)
add_hr()
add_body("Note on dates: search results place current date context at mid-2026, so figures through Q3/Q4 FY26 (Dec 2025/Mar 2026) are cited as current.", italic=True, size=9)

add_heading("7.1 Management Track Record", 2)
add_body(
    "MD & CEO B. Ramesh Babu has run KVB through a multi-year NPA cleanup with notably consistent guidance "
    "discipline. In 2023 he stated the bank aims \u201cto meet every guidance given for next eight quarters\u201d "
    "(Economic Times, May 2023). In Oct 2022 he guided GNPA below 3% by March 2023 (The Hindu); actual GNPA "
    "came in at 2.27% for Q4FY23 (Livemint, May 2023) \u2014 a beat. NIM guidance tracked closely: 4% NIM called "
    "\u201cdoable\u201d for Q4FY24 (CNBC-TV18, Jan 2024) and delivered. For FY26, guided NIM 3.9\u20133.95% and ROA "
    "\u201cabove 1.85%\u201d mid-year; actual FY26 ROA came in at 1.93%, NIM ~4% (Quartr FY26 summary) \u2014 another modest "
    "beat. For FY27, guidance turned more conservative: NIM down to 3.75\u20133.8% (from ~4%), ROA to 1.7\u20131.8% "
    "(from 1.93%), citing rising deposit costs and competition (ET BFSI / CNBC-TV18, early 2026)."
)
add_body("Assessment: conservative-guide / modest-beat pattern, not aggressive over-promising; commentary has grown more candid about margin compression over time.", bold=True, size=9.5)

add_heading("7.2 Asset Quality Deep Dive", 2)
add_body(
    "Verified 5-year GNPA arc: FY21 ~7.85% \u2192 FY22 5.96% (ET, May 2022) \u2192 FY23 2.27%, PCR 94.85% (Livemint) \u2192 "
    "FY24 GNPA 1.40%/NNPA 0.40% (The Hindu, May 2024) \u2192 FY25 ~0.83-1.32% across quarters \u2192 FY26: GNPA 0.75%, "
    "NNPA 0.19%, PCR 96.56% (Q3), CRAR 18.76% (Quartr, scanx.trade)."
)
add_body(
    "One credible dissenting data point: Geojit (Oct 2024) flagged rising Portfolio-At-Risk (PAR 31-180/181-360) "
    "in MSME, suggesting NPA ratios \u201cbottomed out\u201d (Moneycontrol). Could NOT verify exact restructured-book/SMA "
    "percentages from primary filings in this search \u2014 needs direct pull from latest investor presentation PDF. "
    "No evidence found of evergreening; decline appears genuine (recoveries/write-offs, not restructuring-driven "
    "concealment)."
)

add_heading("7.3 CASA and Margin Sustainability", 2)
add_body(
    "CASA decline is a system-wide phenomenon, not KVB-specific: system CASA fell to 37.9% (Dec-2025) from 40.1% "
    "(Dec-2023) per RBI data (ET BFSI); another source cites system CASA falling ~48%\u2192~36% (2022-2025) as "
    "households shifted to higher-yield instruments (CRISIL-cited Substack, 2026). KVB's own reported CASA was "
    "31.16% (down from 31.8%) as of Dec-2024 (ET) \u2014 this conflicts with the ~39% figure referenced at the "
    "screening stage; flagged for reconciliation against the latest investor presentation (see Section 12)."
)
add_body(
    "FE/ET (March 2025) confirms KVB fixed deposits outgrowing CASA, pressuring cost of funds. Loan mix: "
    "RAM (Retail/Agri/MSME) ~86% of advances (FY25/FY26); gold loans 28% of advances as of March 2025 (up "
    "from 25%), capped below 35% by management (FE, May 2025; CNBC-TV18). Corporate book cut to 14% (from "
    "40% in 2019) but management now plans to raise it back to ~20% over 2-3 years via lease-rental-discounting/"
    "NCD \u201ccredit substitution\u201d (ET interview, 2026) \u2014 a strategic reversal worth flagging just as margins compress."
)

add_heading("7.4 Capital Allocation", 2)
add_body(
    "DPS rising steadily: \u20b90.50 (FY21) \u2192 \u20b91.60 (FY22) \u2192 \u20b92.00 (FY23) \u2192 \u20b92.40 (FY24) \u2192 \u20b92.60 (FY25) \u2192 "
    "\u20b92.60 proposed FY26 (dhan.co; scanx.trade AGM notice), payout ratio low (~15-20%). CRAR strong and rising: "
    "16.05% (Dec 2025) \u2192 18.76% (Mar 2026). 1:5 bonus issue executed Aug 2025 (cosmetic, not a raise). "
    "Management stated explicitly (early 2026) no need for external capital over the next two years (CNBC-TV18). "
    "No QIP/rights issue found FY22-FY26 \u2014 bank appears fully capital-generative from internal accruals."
)

add_heading("7.5 Credit Rating", 2)
add_body(
    "ICRA reaffirmed 'ICRA AA (Stable)' issuer rating and 'ICRA A1+' for CD program, raising program size "
    "\u20b910,000cr\u2192\u20b912,000cr (TradingView/Moneycontrol, 2026); A1+ for CDs reaffirmed again 29 June 2026 "
    "(investywise.com). CRISIL FAAA cited only via a secondary FD-aggregator source (angelone.in) \u2014 NOT "
    "verified against a primary CRISIL rationale document. Direct attempts to retrieve the qualitative "
    "rationale text (strengths/concerns) from icra.in's rationale archive were unsuccessful \u2014 only rating "
    "level/program-size changes confirmed (see Section 12: open item)."
)

add_heading("7.6 Valuation Context vs. Peers", 2)
add_body(
    "KVB: P/E ~11.6x, market cap ~\u20b929,500cr, ROE ~19.2% recently, FY27 consensus ROE ~17.5% \u2014 reportedly "
    "highest among mid-size peers (ET BFSI: \u201camong mid-sized private banks, Karur Vysya Bank stands out with "
    "a projected FY27 ROE of 17.5 per cent, the highest among its peers\u201d). City Union Bank ROE ~12-13.5% "
    "(sources diverge: shareprice-target.com cites 12.08%, ET BFSI cites >13%), P/E ~16.5x. RBL Bank P/E ~63x "
    "despite weaker CASA (~29.2%) \u2014 likely reflects a low earnings base, not superior quality."
)
add_body(
    "Directionally, KVB screens cheaper than City Union Bank despite comparable/superior ROE \u2014 a real "
    "relative-value gap, plausibly explained by smaller scale, Tamil Nadu geographic concentration, and "
    "lingering NPA-cleanup-history discount rather than hidden risk."
)

add_heading("7.7 Red Flags Checked", 2)
add_body(
    "Promoter holding is minimal (~2.06%) \u2014 KVB is professionally managed, not promoter-controlled. No RBI "
    "enforcement action or governance scandal found. One immaterial consumer-court matter (\u20b910,000 ATM dispute, "
    "penalty dating to 2017) surfaced (Indian Express, 2026) \u2014 indicates slow historical grievance redressal "
    "but not a material red flag. No auditor qualifications/adverse Key Audit Matters found in search snippets, "
    "but the full annual report audit section was not directly accessed \u2014 needs direct confirmation (Section 12). "
    "No large single-borrower concentration issues surfaced, consistent with the bank's deliberate 2019+ "
    "corporate de-risking."
)

add_heading("7.8 Growth Runway", 2)
add_body(
    "FY26 credit growth guidance \u201c1-2% above industry\u201d was delivered (industry 16.1%, KVB ~17%) per Ramesh "
    "Babu (ET interview, 2026). Branch count ~902 as of 2026. Confirmed new initiatives directly from the MD: "
    "credit cards (imminent launch, via 80-lakh BNPL/Amazon customer base), loans against mutual funds, "
    "microfinance scale-up (\u20b9180cr \u2192 target \u20b9500-700cr), affordable housing re-entry (\u20b920-50 lakh tickets "
    "via co-lending), and corporate book expansion 14%\u219220% via lease-rental-discounting/NCD substitution. "
    "These read as credible incremental diversification, though the corporate re-lever is the key underwritten "
    "risk for a 3-5 year thesis."
)

doc.add_page_break()

# =========================================================
# SECTION 8: DEEP DILIGENCE - TECHNO ELECTRIC (REJECTED)
# =========================================================
add_heading("8. Deep Diligence \u2014 Techno Electric & Engineering (Rejected After Deep Diligence)", 1)
add_hr()

add_heading("8.1 Priority Item: Sankhya Financial Services NCD Issue", 2)
add_body(
    "The actual disclosure (BSE/NSE corporate update, ~28 May 2026, coincident with FY26 board meeting on "
    "25 May): \u201cReceived Rs.80 crore on 25.05.2026 against Sankhya Financial Services NCDs; residual balance "
    "resolution under discussion.\u201d Source: Screener.in filings feed (mirrors BSE/NSE) and Groww market-news feed."
)
add_body(
    "Sankhya Financial Services Pvt Ltd = the NBFC arm of the Mumbai-based \u201cTrust Group\u201d (Trust Investment "
    "Advisors / Trust Mutual Fund, CEO Sandeep Bagla), a debt-capital-markets house founded 2005, active in "
    "repo/corporate bond trading (Cbonds issuer data). No evidence found linking Sankhya/Trust Group to Techno's "
    "promoters (P.P. Gupta, Chatterjee family) or to known group entities (Simran Wind Project, Techno Infra) \u2014 "
    "these appear to be two unrelated business houses (Kolkata industrial promoters vs. Mumbai debt-market "
    "promoters)."
)
add_body(
    "Techno is debt-free with a large treasury book (~\u20b92,600 Cr cash+investments; balance-sheet \u201cInvestments\u201d "
    "line grew \u20b91,142 Cr FY24 \u2192 \u20b92,836 Cr FY25 \u2192 \u20b92,258 Cr FY26). Sankhya NCDs are most plausibly one line "
    "within this treasury/yield portfolio, not a strategic or related-party investment \u2014 this reads as a "
    "counterparty-credit/treasury-management issue, not evidence of fund diversion."
)
add_body(
    "CRITICAL GAP: the FY26 Annual Report's actual \u201cInvestments\u201d note or \u201cRelated Party Transactions\u201d note "
    "could not be accessed (not yet indexed/available online at time of research) to (i) confirm original "
    "exposure size, (ii) confirm Sankhya is NOT classified as a related party, (iii) check if any impairment/"
    "provision was booked. Conclusion on this item alone was provisional: \u201cverify, then proceed\u201d \u2014 not a "
    "clear rejection trigger by itself.",
    bold=True
)

add_heading("8.2 The Decisive Issue: Operating Cash Flow Quality", 2)
add_body(
    "Independent of the Sankhya matter, the cash flow statement (goodreturns.in, cross-checked against "
    "Screener.in working-capital-days data) shows a pattern that, on its own, is sufficient grounds for "
    "rejection on a 3-5 year long thesis:"
)
cf_table = [
    ["\u20b9 Cr", "FY22", "FY23", "FY24", "FY25", "FY26"],
    ["PBT", "326.35", "288.12", "334.46", "545.63", "590.68"],
    ["Operating Cash Flow", "258.73", "32.55", "-337.42", "837.28", "-574.91"],
    ["Investing Cash Flow", "-197.39", "73.33", "403.29", "-1,978.80", "699.35"],
]
t = doc.add_table(rows=len(cf_table), cols=len(cf_table[0]))
t.style = 'Table Grid'
for i, row in enumerate(cf_table):
    for j, val in enumerate(row):
        t.cell(i, j).text = val
style_table(t)
add_body(
    "Operating cash flow was negative in two of the last three years, and FY26's operating cash outflow "
    "(-\u20b9575 Cr) is the worst in the five-year window, arriving in the same year reported profit hit a record "
    "high. The offsetting positive investing-activity cash flow in FY26 (+\u20b9699 Cr) suggests the company funded "
    "operations by liquidating part of its treasury/investment book (the same balance sheet line that holds "
    "the still-unresolved Sankhya exposure) rather than converting profit into cash.",
    bold=True
)
add_body(
    "This corroborates, rather than resolves, the working-capital-days blowout (222\u2192430) found during "
    "screening. Deep diligence further established that debtor days actually FELL over the same period "
    "(282\u2192108\u2192137, FY23-26) \u2014 meaning the deterioration is NOT simple government-receivables stretch as "
    "commonly assumed in Indian EPC/power names; it more likely traces to swings in the current-investments/"
    "treasury book. No management explanation was found on available transcripts."
)

add_heading("8.3 Other Findings", 2)
add_bullet("Management guidance FY26 (from FY24 concall: revenue \u20b93,500-3,600 Cr, EPS \u20b950) vs delivered (\u20b93,252 Cr standalone revenue, EPS \u20b946.60) \u2014 a modest miss after a string of in-line/beat years. FY27 guidance now \u20b94,000 Cr / EPS \u20b960.", size=9.5)
add_bullet("FY26 growth driven by broad EPC/transmission execution (~65% of \u20b99,566 Cr order book), not one-offs. Data centers and smart metering still small contributors (data center FY27 revenue guided only \u20b940-50 Cr; management explicitly walked back market expectations of \u20b9100 Cr).", size=9.5)
add_bullet("Order book \u20b99,566 Cr, ~65% transmission EPC, PSU/state-utility/Navratna CPSU-heavy (RailTel, DVC, state DISCOMs) \u2014 structurally normal for Indian T&D EPC.", size=9.5)
add_bullet("Debt-free balance sheet, FY26 dividend \u20b97/share, unmodified (clean) audit opinion for FY26, no SEBI action or qualified opinion found. Promoter holding declined 61.5%\u219256.9% over 3 years (moderate dilution flag, not pledging/fraud signal).", size=9.5)
add_bullet("Credit ratings (ICRA/CRISIL, \u201cAA stable\u201d per management commentary) rating rationale PDFs were not directly accessible \u2014 rationale text on related-party/receivables commentary remains unverified.", size=9.5)

verdict_tag(
    "REJECTED. Do not treat the Sankhya matter alone as a disqualifying related-party diversion \u2014 the "
    "counterparty appears unaffiliated. However, combined with two years of negative operating cash flow "
    "against rising reported profit, and no management explanation for the working-capital anomaly available "
    "in accessible transcripts, the bear case (poor cash conversion, opaque treasury exposure) is at least as "
    "strong as the bull case (T&D capex cycle, data center optionality). Rejected per the mandate to reject "
    "where evidence of a mispricing is weak.",
    'reject'
)
sp = doc.add_paragraph()
r = sp.add_run("Sources: ")
r.bold = True; r.italic = True; r.font.size = Pt(8.5)
r2 = sp.add_run("BSE/NSE corporate disclosure (~28 May 2026); goodreturns.in cash flow statement; screener.in working-capital data; Quartr Q4 25/26 and prior-quarter transcript summaries; FY26 results filings.")
r2.italic = True; r2.font.size = Pt(8.5); r2.font.color.rgb = GREY

doc.add_page_break()

# =========================================================
# SECTION 9: DEEP DILIGENCE - CCL PRODUCTS
# =========================================================
add_heading("9. Deep Diligence \u2014 CCL Products India (Promoted Finalist)", 1)
add_hr()

add_heading("9.1 Margin Compression Explanation", 2)
add_body(
    "OPM fell ~18%\u2192~16% in FY26, but management (per Tradebrains/Financial Express coverage of the Q4 FY26 "
    "call, May 2026) frames this as an optical/cost-plus pass-through effect \u2014 rising green coffee prices "
    "inflate revenue faster than absolute profit, mechanically depressing the margin percentage. EBITDA/kg "
    "actually rose \u20b9120\u2192\u20b9135-140 over FY26, driven by freeze-dried mix shift, small-pack growth, and capacity "
    "utilization. Assessed as largely transient, corroborated by unit economics \u2014 not structural margin loss."
)

add_heading("9.2 Capacity Expansion", 2)
add_body(
    "Combined India+Vietnam capacity ~77,000 MT (up from 38,500 TPA in FY22). Utilization currently ~65%, "
    "guided to reach 72-73% (FY27) and 80-85% (FY28). Management explicitly guided NO major capex for 2 years "
    "(only \u20b925-35 Cr/yr maintenance) \u2014 FY27-28 growth runway is fundable from existing capacity."
)

add_heading("9.3 Customer Concentration \u2014 Unresolved Gap", 2)
add_body(
    "Could NOT find hard top-5/top-10 revenue percentage in any source reviewed, including a direct web_fetch "
    "attempt against cclproducts.com/investors (returned no matching content). Directional evidence (Karnik "
    "Substack commentary) suggests broad diversification (100+ countries, \u201chundreds of customers,\u201d no single "
    "dominant client) versus smaller peer Vintage Coffee's concentrated Russia exposure, but this is not a "
    "substitute for hard disclosure. Flagged as a material verification gap requiring a direct annual report "
    "pull (see Section 12).",
    bold=True
)

add_heading("9.4 Branded Business Optionality", 2)
add_body(
    "Continental Coffee ~\u20b9430-440 Cr FY26 branded sales (~10% of \u20b94,466 Cr consolidated revenue), growing "
    "40-50% YoY, now India's #3 instant coffee brand. Real optionality/re-rating lever, but unproven at "
    "national scale beyond its South India stronghold."
)

add_heading("9.5 Capital Allocation", 2)
add_body(
    "Net debt cut from ~\u20b91,815-2,000 Cr to \u20b91,073 Cr (>\u20b9750 Cr reduction), net debt/EBITDA 3.1x\u21921.45x, "
    "FCF swung to +\u20b9788 Cr in FY26. Dividends maintained (20-33% historical payout). Promoter holding stable "
    "~46.1% (dhan.co shareholding pattern, Jun 2025\u2013Jun 2026), no pledge evidence found but NOT independently "
    "verified against a primary BSE filing (Section 12)."
)

add_heading("9.6 Management Track Record", 2)
add_body(
    "Generally credible \u2014 defended FY25 volume guidance through a weak Q3, and FY26 guidance was beaten "
    "significantly (though largely commodity-price driven, not pure execution alpha). FY27 guidance reset "
    "conservatively to 15% volume / 15% EBITDA growth, with no major capex planned (Q4 25/26 transcript, "
    "stockanalysis.com)."
)

add_heading("9.7 Industry Structure", 2)
add_body(
    "CCL growing 18-20% vs. ~5% global instant coffee market growth strongly implies share gains, though no "
    "direct Nestl\u00e9/Olam/Tata Coffee head-to-head share data was found \u2014 this is an inference, not a stated "
    "fact, and is flagged as such in the final pitch. Structural tailwinds (premiumization, India coffee "
    "culture shift, natural FX hedge) are real; robusta/arabica price volatility is a live risk (Arabica hit "
    "an all-time record October 2025)."
)

add_heading("9.8 Credit Rating \u2014 Resolved", 2)
add_body(
    "Initial diligence pass could not verify a credit rating letter grade despite multiple attempts (Screener.in "
    "confirmed Fitch/India Ratings as agency of record for historical updates, 2019-2026, but letter grade was "
    "not retrievable). Follow-up search resolved this: ET Markets explicitly lists CCL Products as \u201cNo Rating "
    "(NR)\u201d (economictimes.indiatimes.com, accessed Jul 2026). This is consistent with a company that has been "
    "actively delevering and likely runs primarily on bank working-capital lines rather than rated public debt "
    "instruments \u2014 not evidence of distress. The falling net debt/EBITDA trend (3.1x\u21921.45x) corroborates this "
    "benign reading."
)

add_heading("9.9 Red Flags Checked", 2)
add_body(
    "Inventory days normalized favorably (262\u2192131 days, FY21\u2192FY26) as commodity prices eased. No auditor "
    "qualification found, but this is an information gap rather than a clean-bill confirmation \u2014 the annual "
    "report auditor's report/RPT note should be pulled directly (Section 12). Stock has re-rated sharply "
    "(~\u20b9475\u2192~\u20b91,220 in ~1 year, now ~41-42x P/E vs 5yr median ~31x per etmoney.com and tradebrains.in) \u2014 "
    "valuation/de-rating risk if FY27's guided deceleration disappoints, especially since the FY26 beat was "
    "largely commodity-driven."
)

verdict_tag(
    "PROMOTED AS FINALIST #2. Margin compression resolved as transient/optical (EBITDA/kg confirms improving "
    "unit economics). Capacity, deleveraging, and branded-business optionality are real and well-evidenced. "
    "Customer concentration remains a genuine, unresolved disclosure gap and is carried into the final pitch "
    "as an explicit limitation rather than assumed away.",
    'interest'
)
sp = doc.add_paragraph()
r = sp.add_run("Sources: ")
r.bold = True; r.italic = True; r.font.size = Pt(8.5)
r2 = sp.add_run("cclproducts.com/investors (direct fetch attempt); Tradebrains, Financial Express (Q4 FY26 call coverage, May 2026); Q1-Q4 FY25/FY26 transcript summaries (stockanalysis.com, Quartr); dhan.co shareholding pattern; etmoney.com valuation snapshot; economictimes.indiatimes.com rating listing.")
r2.italic = True; r2.font.size = Pt(8.5); r2.font.color.rgb = GREY

doc.add_page_break()

# =========================================================
# SECTION 10: VALUATION DATA GATHERING
# =========================================================
add_heading("10. Valuation Data-Gathering Notes", 1)
add_hr()

add_heading("10.1 Karur Vysya Bank \u2014 Live Data Points (15 Jul 2026)", 2)
kvb_val = [
    ["Metric", "Value", "Source"],
    ["Price", "\u20b9305.40", "dhan.co, 15 Jul 2026, 3:45 PM"],
    ["Market cap", "\u20b929,517 Cr", "dhan.co"],
    ["P/E (TTM)", "11.60x", "dhan.co"],
    ["Book value / share", "~\u20b9146", "screener.in"],
    ["ROE", "19.1-19.2%", "screener.in; ET BFSI"],
    ["FY27 consensus ROE", "17.5% (highest among mid-size peers)", "ET BFSI, 2026"],
    ["City Union Bank P/E", "~16.5x", "moneycontrol.com peer comparison"],
    ["City Union Bank ROE", "12-13.5% (sources diverge)", "shareprice-target.com; ET BFSI"],
    ["RBL Bank P/E", "~63x (low earnings base)", "cross-source screening"],
]
t = doc.add_table(rows=len(kvb_val), cols=len(kvb_val[0]))
t.style = 'Table Grid'
for i, row in enumerate(kvb_val):
    for j, val in enumerate(row):
        t.cell(i, j).text = val
style_table(t)
add_body(
    "Note: multiple data points show minor divergence across aggregators (e.g., P/E cited as 11.6x-13.6x "
    "across dhan.co, cnbctv18.com, and other sources at different snapshot dates) \u2014 attributable to timing "
    "differences and standalone vs. consolidated presentation. The 11.6x figure (dhan.co, same-day as price "
    "used) was used as the base case input.",
    italic=True, size=9
)

add_heading("10.2 CCL Products \u2014 Live Data Points (15 Jul 2026)", 2)
ccl_val = [
    ["Metric", "Value", "Source"],
    ["Price", "~\u20b91,220 (range \u20b91,183-1,229 intraday)", "dhan.co, etmoney.com, 15 Jul 2026"],
    ["Market cap", "\u20b915,919 Cr", "etmoney.com"],
    ["P/E (TTM)", "41.02x", "etmoney.com"],
    ["Sector P/E", "58.80x", "etmoney.com"],
    ["ROE", "18%", "etmoney.com"],
    ["ROCE", "15.90%", "etmoney.com"],
    ["Book value / share", "\u20b9175.59", "etmoney.com"],
    ["P/B", "6.79x", "etmoney.com"],
    ["Debt/equity", "0.55", "etmoney.com"],
    ["EBITDA margin (TTM)", "16.63%", "etmoney.com"],
    ["52-week range", "\u20b9815.50 \u2013 \u20b91,236.20", "samco.in"],
    ["Enterprise value", "\u20b916,876 Cr (\u2248\u20b9168.76bn)", "stockanalysis.com"],
    ["Promoter holding", "46.09-46.14% (stable, Jun25-Jun26)", "dhan.co shareholding pattern"],
]
t = doc.add_table(rows=len(ccl_val), cols=len(ccl_val[0]))
t.style = 'Table Grid'
for i, row in enumerate(ccl_val):
    for j, val in enumerate(row):
        t.cell(i, j).text = val
style_table(t)

add_heading("10.3 Valuation Methods Selected & Rationale", 2)
add_bullet("Karur Vysya Bank: EPS compounding + terminal P/E re-rating, cross-checked via P/B-ROE. Chosen because bank equity is best valued on earnings power and return-on-equity trajectory rather than EV/EBITDA (not meaningful for a financial institution) or DCF (balance-sheet leverage makes unlevered FCF less informative than for an operating company).", size=9.5)
add_bullet("CCL Products: EV/EBITDA forward multiple applied to guided EBITDA growth. Chosen over pure P/E because the company is mid-deleveraging (net debt/EBITDA falling from 3.1x to 1.45x), so EV/EBITDA better isolates operating performance from the equity-value effect of debt paydown.", size=9.5)
add_bullet("Sensitivity ranges (bear/base/bull) were built directly from the range of management's own guidance statements (e.g., CCL's FY27 guidance of 15%/15%; KVB's FY27 NIM/ROA guidance) rather than arbitrary spreads, so that scenario inputs are traceable to primary commentary.", size=9.5)

doc.add_page_break()

# =========================================================
# SECTION 11: RANKING METHODOLOGY
# =========================================================
add_heading("11. Final Ranking Methodology & Score Rationale", 1)
add_hr()
add_body(
    "Each of the 10 non-rejected names was scored 0-100 based on four weighted inputs, assessed qualitatively "
    "against the evidence gathered rather than a rigid formula (appropriate given the mix of a bank, industrials, "
    "healthcare, and consumer names, which are not directly comparable on identical metrics):"
)
add_bullet("Evidence of genuine (not cosmetic) fundamental improvement \u2014 margin, ROCE/ROE, or cash-flow inflection corroborated by more than one data point (e.g., unit economics vs. headline margin, or multi-year trend vs. single quarter).", size=10)
add_bullet("Absence of disqualifying red flags \u2014 governance, related-party transactions, promoter pledging, accounting quality, or guidance misses.", size=10)
add_bullet("Valuation relative to demonstrated (not promised) quality \u2014 names pricing in a turnaround that has not yet shown up in trailing numbers scored lower on this dimension.", size=10)
add_bullet("Management credibility \u2014 track record of guidance accuracy and candor about deteriorating metrics, drawn from multi-year concall/interview commentary where available.", size=10)
add_body(
    "Scores in the range 70-85 (Karur Vysya Bank, CCL Products, and Techno Electric & Engineering pre-deep-"
    "diligence) reflect names where the initial screening evidence showed a genuine, multi-point-corroborated "
    "improvement with limited red flags. Scores in the 55-68 range (Neuland Labs, Greenpanel, Concord Biotech) "
    "reflect a real but less certain thesis \u2014 either high quarterly volatility (Neuland), an unproven cyclical "
    "inflection (Greenpanel), or a good business at a valuation that already assumes recovery (Concord). Scores "
    "below 50 reflect either a valuation that has run ahead of fundamentals (Metropolis, Triveni Turbine, "
    "Sansera) or a name where red flags actively outweighed the growth story (Global Health/Medanta's governance "
    "overhang; KIMS's pledging and disclosure issues, which placed it at the bottom despite superficially "
    "attractive growth headlines)."
)
add_body(
    "The Techno Electric score was revised down from 74 (screening) to 60 (post-deep-diligence, shown in the "
    "final memo's ranking table) rather than removed from the table entirely, so the ranking transparently shows "
    "the effect of deeper evidence on the initial assessment \u2014 consistent with the requirement to reject where "
    "the bear case turns out to be stronger than the bull case, even for a name that screened well initially."
)

doc.add_page_break()

# =========================================================
# SECTION 12: OPEN ITEMS REGISTER
# =========================================================
add_heading("12. Open Items / Unverified Claims Register", 1)
add_hr()
add_body(
    "The following items were identified during research as material to the thesis but could not be verified "
    "against a primary source within this research session. They are listed here so that anyone acting on this "
    "research knows exactly what remains to be confirmed before capital is committed.",
    italic=True
)

open_items = [
    ["Company", "Open Item", "Why It Matters", "Recommended Next Step"],
    ["Karur Vysya Bank", "CASA ratio discrepancy: 31.16% (ET, Dec-2024) vs. ~39% (other sources cited at screening stage)", "CASA is central to the margin-sustainability catalyst/risk; a 8pp discrepancy materially changes the read on cost-of-funds trajectory", "Pull the exact figure from KVB's latest (Q4 FY26 / Q1 FY27) investor presentation PDF directly from karurvysyabank.com"],
    ["Karur Vysya Bank", "Full ICRA/CRISIL rating rationale text (qualitative strengths/concerns), not just the rating level", "Rating agency's own risk framing may surface concerns not visible in ratio analysis", "Retrieve rationale PDF from icra.in/Rating/AllRatingRationales or KVB's credit-ratings IR page"],
    ["Karur Vysya Bank", "SMA-2 / restructured book percentages", "Could indicate asset-quality management via restructuring rather than genuine recovery, if elevated", "Pull from latest investor presentation, asset-quality slide"],
    ["Karur Vysya Bank", "Key Audit Matters in FY24/FY25 Annual Report audit section", "Any auditor-flagged concerns would be material to management-quality assessment", "Direct read of Annual Report audit section, not yet accessed"],
    ["CCL Products", "Top-5 / top-10 customer revenue concentration", "Private-label export model implies concentration risk not currently quantified anywhere in public sources reviewed", "Direct pull of FY26 Annual Report customer/segment disclosure notes"],
    ["CCL Products", "Promoter pledge status", "No pledge found in searches, but not verified against a primary BSE shareholding filing", "Cross-check BSE shareholding pattern filing directly"],
    ["CCL Products", "Credit rating confirmation", "ET Markets lists \u201cNo Rating (NR)\u201d; consistent with delevering/bank-funded profile but not verified against a primary company filing", "Confirm via cclproducts.com IR page or BSE filing"],
    ["Techno Electric & Engineering (rejected)", "Nature of Sankhya Financial Services NCD exposure \u2014 original size, related-party status, provisioning", "Central to whether this was a benign treasury holding or a governance concern; company was rejected on cash-flow grounds independent of this, but the item remains open", "Pull FY26 Annual Report Investments note and Related-Party-Transactions note directly once indexed/available"],
    ["Techno Electric & Engineering (rejected)", "Management explanation for working-capital-days blowout (222\u2192430) despite falling debtor days", "No explanation found in accessible transcripts; needed to distinguish a one-off treasury timing issue from a recurring cash-conversion problem", "Direct question at next concall / management commentary review"],
]
t = doc.add_table(rows=len(open_items), cols=len(open_items[0]))
t.style = 'Table Grid'
widths = [1.3, 2.0, 1.8, 1.8]
for i, row in enumerate(open_items):
    for j, val in enumerate(row):
        cell = t.cell(i, j)
        cell.text = val
        cell.width = Inches(widths[j])
style_table(t)

add_body(
    "General methodological caveat: this entire research effort relied on public web search and aggregator "
    "sites (Screener.in, stockanalysis.com, Economic Times, Moneycontrol, Quartr transcript summaries, dhan.co, "
    "etmoney.com) rather than a Bloomberg/FactSet terminal or guaranteed access to every primary company filing. "
    "Multiple direct-fetch attempts against primary sources (icra.in rationale PDFs, cclproducts.com investor "
    "pages, NSE filing archives) either returned inaccessible content, redirected to non-primary mirrors, or "
    "404'd. No figures were invented to fill these gaps; every unresolved item is listed above rather than "
    "silently assumed.",
    italic=True, size=9.5, space_after=0
)

doc.save('/projects/sandbox/Backend_Research_Working_Papers.docx')
print("Saved backend research working papers.")

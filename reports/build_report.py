"""
Build IC-style Word report: Two High-Conviction Long Ideas (Indian Equities)
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

doc = Document()

# ---------- base style ----------
normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = Pt(10.5)
normal.font.color.rgb = DARKGREY
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.15

sections = doc.sections
for s in sections:
    s.top_margin = Cm(1.8)
    s.bottom_margin = Cm(1.8)
    s.left_margin = Cm(2.0)
    s.right_margin = Cm(2.0)

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
        sizes = {1: 18, 2: 14, 3: 12}
        r.font.size = Pt(sizes.get(level, 12))
    return p

def add_body(text, bold=False, italic=False, size=10.5, color=DARKGREY, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p

def add_bullet(text, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(4)
    if bold_lead:
        r1 = p.add_run(bold_lead)
        r1.bold = True
        r1.font.size = Pt(10.5)
        r2 = p.add_run(text)
        r2.font.size = Pt(10.5)
    else:
        r = p.add_run(text)
        r.font.size = Pt(10.5)
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
                    r.font.size = Pt(9)
            if i == 0:
                set_cell_shading(cell, header_color)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                        r.bold = True
            else:
                if i % 2 == 0:
                    set_cell_shading(cell, 'F2F2F2')

# =========================================================
# COVER
# =========================================================
title = doc.add_paragraph()
title.paragraph_format.space_before = Pt(60)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run("TWO HIGH-CONVICTION LONG IDEAS")
r.bold = True
r.font.size = Pt(26)
r.font.color.rgb = NAVY

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.paragraph_format.space_after = Pt(4)
r = sub.add_run("Indian Equities \u2014 3\u20135 Year Investment Horizon")
r.font.size = Pt(15)
r.font.color.rgb = ACCENT
r.italic = True

sub2 = doc.add_paragraph()
sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub2.paragraph_format.space_after = Pt(40)
r = sub2.add_run("Investment Committee Memorandum")
r.font.size = Pt(12)
r.font.color.rgb = DARKGREY

meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = meta.add_run("Prepared for: Investment Committee\nCoverage: Indian mid/large-cap equities\nDate: July 2026\nPrepared by: Investment Research (Buy-Side)")
r.font.size = Pt(10.5)
r.font.color.rgb = DARKGREY

doc.add_page_break()

# =========================================================
# EXECUTIVE SUMMARY / METHODOLOGY
# =========================================================
add_heading("Executive Summary & Research Methodology", 1)
add_hr()
add_body(
    "This memo presents two long ideas for a 3\u20135 year holding period, selected from a screening universe of "
    "16 Indian mid/large-cap equities spanning auto ancillaries, banking, wealth management, pharma CDMO, "
    "building materials, capital goods, hospitals, diagnostics, F&B exports, specialty chemicals, and EPC. "
    "The universe was deliberately constructed to avoid both consensus large-cap names and obscure micro-caps."
)
add_body(
    "Research drew on company investor-relations disclosures, exchange filings, quarterly results, earnings-call "
    "transcript summaries, credit-rating actions, and financial media coverage, cross-referenced against "
    "peer filings and industry data. Six names were rejected outright at the screening stage on hard evidence "
    "(governance lapses, accounting-quality concerns, or guidance misses \u2014 see Section 2). Of the remaining ten, "
    "two were carried into deep primary-source diligence. One of the two initial finalists, Techno Electric & "
    "Engineering, was subsequently dropped after deep diligence revealed two years of negative operating cash "
    "flow in the last three despite rising reported profit, compounding an unresolved related-party-adjacent "
    "NCD counterparty exposure. Per the mandate to reject where the bear case is at least as strong as the "
    "bull case, it was replaced by CCL Products (India) Ltd, which cleared a comparable evidentiary bar."
)
add_body(
    "A note on evidentiary limits: this research relied on public aggregator data (Screener.in, stockanalysis.com, "
    "Economic Times, Moneycontrol, Quartr transcript summaries) cross-referenced against management commentary "
    "reported in financial media, rather than direct access to full annual report PDFs or complete ICRA/CRISIL "
    "rationale documents in every instance. Specific unresolved gaps are flagged explicitly within each pitch "
    "rather than filled with assumed figures.",
    italic=True, size=9.5
)

# =========================================================
# SECTION 1: TOP 10 RANKING
# =========================================================
add_heading("1. Top-10 Ranking (Post-Screening)", 1)
add_hr()

data = [
    ["Rank", "Company", "Score/100", "Conviction", "Exp. Ann. Return", "Risk", "Quality", "Valuation Attractiveness"],
    ["1", "Karur Vysya Bank", "81", "High", "17\u201320%", "Medium-Low", "High", "Attractive"],
    ["2", "CCL Products (India)", "72", "High", "15\u201318%", "Medium", "Good", "Fair-to-attractive"],
    ["3", "Techno Electric & Engineering", "60*", "Rejected", "n/a", "Medium-High", "Unresolved", "n/a \u2014 cash flow quality failed"],
    ["4", "Neuland Laboratories", "64", "Medium", "12\u201315%", "Medium-High", "Good", "Fair"],
    ["5", "Greenpanel Industries", "58", "Speculative", "15\u201325% if inflection real", "High", "Cyclical", "Cheap-on-trough, unproven"],
    ["6", "Concord Biotech", "55", "Low-Medium", "8\u201312%", "Medium", "Good", "Rich given 2 down quarters"],
    ["7", "Metropolis Healthcare", "48", "Low", "6\u20139%", "Medium", "Good franchise, weak growth", "Expensive"],
    ["8", "Triveni Turbine", "46", "Low", "6\u201310%", "Medium", "Good", "Expensive; first PAT decline"],
    ["9", "Sansera Engineering", "42", "Low", "5\u20138%", "Medium-High", "Improving", "Priced for perfection"],
    ["10", "Global Health (Medanta)", "40", "Low", "5\u20138%", "Medium (governance overhang)", "High", "Expensive; thin margin of safety"],
]
table = doc.add_table(rows=len(data), cols=len(data[0]))
table.style = 'Table Grid'
widths = [0.5, 1.7, 0.7, 0.9, 1.1, 1.0, 1.1, 1.5]
for i, row in enumerate(data):
    for j, val in enumerate(row):
        cell = table.cell(i, j)
        cell.text = val
        cell.width = Inches(widths[j])
style_table(table)
add_body("*Score reflects pre-deep-diligence screening view; rejected after primary-source cash-flow analysis (see Section 3).", italic=True, size=9)

add_body("Rejected before ranking (evidence-based):", bold=True, space_after=2)
rejects = [
    ("Kaynes Technology \u2014 ", "undisclosed related-party transaction with subsidiary Iskraemeco; smart-metering receivables blew out from 267 to 486 working-capital days."),
    ("Praj Industries \u2014 ", "guided ~10% growth, delivered a revenue decline; PAT down ~90% YoY."),
    ("NOCIL \u2014 ", "ROCE collapsed from 26% to 4% over eight years; structural margin erosion, classic value trap."),
    ("Anand Rathi Wealth \u2014 ", "Q1 FY27 EBITDA margin collapsed to 33.7% from 46.6%; unresolved promoter pledge disclosure spike from ~0% to 11.42%."),
    ("KIMS Hospitals \u2014 ", "promoter pledging spike, NSE-queried debt disclosure correction, dilutive capital raises without matching return improvement."),
    ("Craftsman Automation \u2014 ", "ROCE fell from ~20% to ~13% on debt/M&A-funded growth; heavy promoter stake sale on top of a recent QIP."),
]
for lead, rest in rejects:
    add_bullet(rest, bold_lead=lead)

doc.add_page_break()

# =========================================================
# PITCH TEMPLATE FUNCTION
# =========================================================
def add_pitch_header(name, ticker, subtitle):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run(f"{name}  ")
    r.bold = True
    r.font.size = Pt(17)
    r.font.color.rgb = NAVY
    r2 = p.add_run(f"({ticker})")
    r2.bold = False
    r2.italic = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = ACCENT
    add_body(subtitle, italic=True, size=10, space_after=10)
    add_hr()

def add_catalyst(num, title_text, what, why, impact, evidence):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"Catalyst {num}: {title_text}")
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = NAVY
    add_bullet(what, bold_lead="What changes: ")
    add_bullet(why, bold_lead="Why it matters: ")
    add_bullet(impact, bold_lead="Expected impact: ")
    ep = doc.add_paragraph()
    ep.paragraph_format.space_after = Pt(4)
    r1 = ep.add_run("Evidence: ")
    r1.bold = True
    r1.italic = True
    r1.font.size = Pt(9.5)
    r2 = ep.add_run(evidence)
    r2.italic = True
    r2.font.size = Pt(9.5)

def add_val_table(rows):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = 'Table Grid'
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            table.cell(i, j).text = val
    style_table(table)
    return table

# =========================================================
# PITCH 1: KARUR VYSYA BANK
# =========================================================
add_heading("2. Investment Pitch \u2014 Karur Vysya Bank Ltd", 1)
add_pitch_header("Karur Vysya Bank Ltd", "NSE: KARURVYSYA", "Recommendation: LONG  |  Core holding  |  3\u20135 year horizon")

add_heading("1. Investment Thesis", 2, size=12)
add_body(
    "Long Karur Vysya Bank. The market prices KVB at ~11.6x trailing P/E and ~2.1x P/B despite a five-year "
    "asset-quality turnaround that is now complete and structurally sound \u2014 GNPA fell from 7.85% (FY21) to "
    "0.75% (FY26), PCR is 96.6%, and ROE reached 19.1%, the highest among mid-size private banks (ET BFSI, 2026). "
    "This is not a re-leveraged or evergreened recovery: the bank has raised no external capital since 2019, is "
    "growing book purely from internal accruals, and CRAR stands at 18.76% (Mar 2026). The market continues to "
    "apply a legacy \u201cNPA-history\u201d discount and a small-scale / Tamil Nadu-concentration discount that the "
    "current risk profile no longer justifies relative to City Union Bank, which trades at a premium multiple "
    "for inferior ROE (~13.5% vs KVB\u2019s 19%)."
)

add_heading("2. Five Catalysts", 2, size=12)
add_catalyst(1, "Continued ROA/NIM outperformance vs. the bank\u2019s own conservative guidance",
    "management has repeatedly guided conservatively and beaten \u2014 GNPA guided <3% for FY23, delivered 2.27%; ROA guided >1.85% for FY26, delivered 1.93%.",
    "a pattern of under-promise/over-beat is a credibility signal the market has not yet fully priced into the multiple.",
    "continued earnings beats support gradual re-rating toward peer multiples.",
    "MD B. Ramesh Babu, on record (Economic Times, May 2023): the bank aims \u201cto meet every guidance given for next eight quarters.\u201d")

add_catalyst(2, "Corporate book re-lever from 14% to ~20% of advances via lease-rental-discounting / NCD substitution over 2\u20133 years",
    "after cutting corporate exposure from 40% (2019) to 14% for de-risking, management now plans a controlled expansion.",
    "diversifies the book beyond gold loans (28% of advances, capped below 35% by management) and RAM concentration (86% of book) \u2014 a growth lever currently absent from consensus models.",
    "incremental advances growth without proportionate risk-weighted-asset drag if executed at similar underwriting discipline.",
    "ET interview with management, 2026.")

add_catalyst(3, "New retail product lines (credit cards, loans against mutual funds, microfinance scale-up, affordable housing re-entry)",
    "KVB is monetizing an 80-lakh existing BNPL/digital customer base for an imminent credit card launch.",
    "fee income diversification and cross-sell economics on an existing low-cost customer base, not a new acquisition cost.",
    "incremental non-NII revenue and improved cost-to-income ratio over 3\u20135 years.",
    "MD commentary, ET interview 2026.")

add_catalyst(4, "Capital-generative compounding without dilution",
    "dividend per share has risen steadily (\u20b90.50 FY21 \u2192 \u20b92.60 FY26) at a low 15\u201320% payout, meaning book value compounds largely undiluted at ~17\u201319% ROE.",
    "for a 3\u20135 year holding period, undiluted BVPS compounding at high-teens ROE is the core driver of equity value, independent of multiple re-rating.",
    "BVPS could roughly double over five years on retained-earnings compounding alone.",
    "dividend history via AGM notices; capital adequacy trending up (16.05%\u219218.76%, Dec 2025\u2013Mar 2026).")

add_catalyst(5, "Peer-relative re-rating as the market digests that the asset-quality turnaround is durable, not cyclical",
    "rating agencies have not yet upgraded KVB beyond ICRA AA (Stable) despite metrics improving every year for five years.",
    "a future rating upgrade would itself be a catalyst, and typically follows sustained delivery rather than leading it.",
    "lower cost of funds on wholesale borrowing; incremental NIM support.",
    "ICRA reaffirmed AA (Stable)/A1+ for CD programme, enhancing programme size \u20b910,000cr\u2192\u20b912,000cr, 29 June 2026 (ICRA/Moneycontrol).")

add_heading("3. What Could Break the Thesis", 2, size=12)
add_body("Downside risks:", bold=True, space_after=2)
add_bullet("CASA ratio decline (system-wide, but KVB-specific figures diverge across sources \u2014 31% per ET Dec-2024 vs. 39% cited elsewhere; needs reconciliation against the FY26 investor presentation) raises cost of funds and could compress NIM faster than the FY27 guided 3.75\u20133.8%. Probability: Medium \u2014 partially pre-guided. Mitigating factor: RAM-heavy book (86%) reprices quickly, particularly gold loans.")
add_bullet("The corporate book re-lever (14%\u219220%) is the single largest underwritten execution risk \u2014 a return to lax corporate underwriting last seen pre-2019 would undo the asset-quality re-rating case. Probability: Low-Medium. Mitigating factor: management\u2019s explicit, repeated caution on underwriting discipline.")
add_bullet("One dissenting data point (Geojit, Oct 2024) flagged rising Portfolio-At-Risk in MSME, suggesting NPA ratios may have \u201cbottomed.\u201d Genuine, if early, signal to monitor.")
add_body("Upside risk: ", bold=True, space_after=2)
add_bullet("faster-than-guided credit card / fee income ramp, or an actual rating upgrade, would accelerate re-rating beyond the base case.")

add_heading("4. Valuation", 2, size=12)
add_body("Current: Price \u20b9305.40 (15 Jul 2026); market cap \u20b929,517 Cr; P/E ~11.6x (TTM EPS \u2248\u20b926.3); P/B \u22482.1x on BVPS \u2248\u20b9146; ROE 19.1%.")
add_body("Method: EPS compounding + terminal P/E re-rating (appropriate for a bank), cross-checked via P/B-ROE. Key assumptions: EPS CAGR driven by credit growth (~17% recently) moderating toward guided levels; gradual P/E re-rating toward, but still below, City Union Bank\u2019s premium.")
add_val_table([
    ["Scenario", "EPS CAGR", "Terminal P/E", "FY29 EPS", "Target Price", "Upside", "Implied CAGR"],
    ["Bear", "8%", "10x", "\u20b933.2", "\u20b9332", "+9%", "~3%"],
    ["Base", "15%", "13x", "\u20b940.1", "\u20b9521", "+71%", "~19%"],
    ["Bull", "18%", "14x", "\u20b943.3", "\u20b9606", "+98%", "~26%"],
])
add_body("Intrinsic value / target price: Base case \u20b9521 over three years implies ~19% annualized upside \u2014 attractive on a risk-adjusted basis for a bank with the sector\u2019s best asset quality and lowest capital-raise risk.", space_after=6)

add_heading("5. Investment Conclusion", 2, size=12)
add_body(
    "Karur Vysya Bank offers a rare combination for a mid-cap Indian bank: a completed, verifiably genuine "
    "asset-quality turnaround, best-in-class ROE among peers, full capital self-sufficiency, and a still-discounted "
    "multiple relative to demonstrably weaker peers. The primary risks \u2014 CASA erosion and corporate-book "
    "re-leverage \u2014 are visible, management-acknowledged, and partially pre-guided rather than hidden. Recommend "
    "initiating a long position, sized for a core holding, with the CASA-ratio reconciliation and corporate-book "
    "growth pace as the two variables to monitor each quarter."
)

doc.add_page_break()

# =========================================================
# PITCH 2: CCL PRODUCTS
# =========================================================
add_heading("3. Investment Pitch \u2014 CCL Products (India) Ltd", 1)
add_pitch_header("CCL Products (India) Ltd", "NSE: CCL", "Recommendation: LONG  |  Smaller initial size pending one disclosure gap  |  3\u20135 year horizon")

add_heading("1. Investment Thesis", 2, size=12)
add_body(
    "Long CCL Products. The market is treating FY26\u2019s operating margin compression (18%\u219216% OPM) as evidence "
    "of deteriorating unit economics during a 43.5% revenue surge, when unit-level data shows the opposite: "
    "EBITDA per kg rose from \u20b9120 to \u20b9135\u2013140 over FY26, and the margin-% decline is a mechanical artifact of "
    "green coffee price pass-through inflating the revenue denominator faster than absolute profit (Q4 FY26 "
    "concall, management commentary via Tradebrains/Financial Express, May 2026). Net debt was cut by over "
    "\u20b9750 Cr (to \u20b91,073 Cr; net debt/EBITDA 3.1x\u21921.45x) in the same year, and combined India+Vietnam capacity "
    "has roughly doubled to ~77,000 MT with utilization still at only ~65% \u2014 meaning FY27\u201328 growth is guided "
    "to require no major capex. The market is pricing a commodity exporter; the underlying trajectory is a "
    "deleveraging, capacity-optionality compounder with an underappreciated branded-business call option."
)

add_heading("2. Five Catalysts", 2, size=12)
add_catalyst(1, "Utilization ramp from ~65% to guided 80\u201385% (FY28) on existing capacity, with no major capex planned",
    "management explicitly guided only \u20b925\u201335 Cr/year of maintenance capex for the next two years.",
    "incremental volume drops through at high operating leverage since fixed costs are already sunk.",
    "EBITDA growth can outpace revenue growth as utilization climbs \u2014 a direct FCF inflection.",
    "Q4 FY26 transcript (stockanalysis.com summary, May 2026): \u201cFY27 guidance targets 15% volume and EBITDA growth, with no major CapEx planned.\u201d")

add_catalyst(2, "Structural deleveraging continuing, freeing cash for growth capex or shareholder returns",
    "net debt fell from ~\u20b91,815\u20132,000 Cr to \u20b91,073 Cr in one year; FCF swung to +\u20b9788 Cr in FY26.",
    "a debt-free-to-net-cash trajectory within 2\u20133 years would materially de-risk the equity and could support dividend increases beyond the current low 20\u201333% historical payout.",
    "lower interest expense directly adds to PAT; optionality for capital return.",
    "FY26 cash flow statement (goodreturns.in); Q1 FY26 transcript: \u201cNet debt reduced to INR 1,671 crore, with further reductions targeted.\u201d")

add_catalyst(3, "Branded domestic business (Continental Coffee) scaling from a low base",
    "branded sales ~\u20b9430\u2013440 Cr in FY26 (~10% of consolidated revenue), growing 40\u201350% YoY, now India\u2019s #3 instant coffee brand.",
    "branded FMCG businesses typically command materially higher margins and multiples than private-label export manufacturing; continued compounding becomes a re-rating catalyst independent of the core export business.",
    "mix-shift margin accretion over the medium term, and a potential SOTP re-rating argument.",
    "Q2\u2013Q3 FY26 transcripts: \u201cDomestic B2C grew 50% YTD\u201d; \u201cDomestic branded business and B2C segments are driving market share.\u201d")

add_catalyst(4, "Structural share gains in global private-label instant coffee",
    "CCL\u2019s volume growth (18\u201320%) is running well ahead of global instant coffee market growth (~5%, industry estimates).",
    "if this differential persists, it is a structural rather than cyclical growth driver.",
    "sustained double-digit volume growth supports the utilization-ramp thesis in Catalyst 1.",
    "Inference from company-reported volume growth vs. industry growth estimates \u2014 flagged explicitly as an inference, not a stated management or third-party fact, and should be treated with appropriate caution.")

add_catalyst(5, "Natural FX hedge and premiumization tailwind from an export-heavy revenue base",
    "as a dollar-revenue exporter with rupee costs, INR depreciation is a structural tailwind; global premiumization / private-label penetration trends in Europe and the US support secular demand.",
    "reduces one risk variable (currency) while adding a demand tailwind.",
    "margin support independent of operational execution.",
    "Industry-structure commentary cross-referenced against company transcripts citing \u201cinternational expansion\u201d as an FY27 focus area.")

add_heading("3. What Could Break the Thesis", 2, size=12)
add_body("Downside risks:", bold=True, space_after=2)
add_bullet("Robusta/Arabica price volatility is real and unhedged pass-through has a lag \u2014 Arabica hit an all-time record in October 2025; a sharp reversal in coffee prices could compress reported margins again before EBITDA/kg data confirms the underlying trend is intact. Probability: Medium \u2014 a genuine commodity-cycle risk, not a one-off. Mitigating factor: EBITDA/kg has trended up through the FY26 volatility.")
add_bullet("Customer concentration could not be verified from any public source accessed in this research \u2014 no top-5/top-10 customer disclosure was found in the annual report excerpts, concall transcripts, or aggregator data reviewed. This is a material unresolved gap, not a cleared item, and should be confirmed directly against the FY26 Annual Report\u2019s customer/segment notes before full position sizing.")
add_bullet("The stock has already re-rated sharply (~\u20b9475\u2192~\u20b91,220 over roughly one year), trading at ~41x TTM P/E versus a 5-year median closer to ~31x; a chunk of the FY26 EPS beat was commodity-price-driven rather than pure execution alpha, so a deceleration to FY27\u2019s guided (lower) 15%/15% targets could trigger a valuation de-rating even if the operating story remains intact.")
add_body("Upside risk: ", bold=True, space_after=2)
add_bullet("faster branded-business scaling or a durable coffee-price tailwind could support continued multiple expansion beyond the base case.")

add_heading("4. Valuation", 2, size=12)
add_body("Current: Price \u2248\u20b91,220 (15 Jul 2026); market cap \u2248\u20b915,900\u201316,000 Cr; TTM P/E \u224841x; EV/EBITDA \u224823x (EV \u2248\u20b916,990 Cr / TTM EBITDA \u2248\u20b9743 Cr); P/B \u22486.8x; ROE 18%; ROCE 15.9%; net debt/equity 0.55.")
add_body("Method: EV/EBITDA forward multiple applied to guided EBITDA growth, cross-checked for de-rating risk given the recent multiple expansion. Key assumptions: FY27 guidance of 15% EBITDA growth as the base case; multiple held flat-to-compressing given elevated starting point; continued net debt paydown.")
add_val_table([
    ["Scenario", "EBITDA CAGR", "EV/EBITDA", "FY29 EV", "Equity Value", "Target Price", "Upside (CAGR)"],
    ["Bear", "10%", "18x", "\u20b917,800 Cr", "\u20b916,900 Cr", "\u20b91,295", "+6% (~2%)"],
    ["Base", "15%", "22x", "\u20b924,840 Cr", "\u20b924,340 Cr", "\u20b91,865", "+53% (~15%)"],
    ["Bull", "18%", "23x", "\u20b928,080 Cr", "\u20b927,580 Cr", "\u20b92,113", "+73% (~20%)"],
])
add_body("Intrinsic value / target price: Base case \u2248\u20b91,865 over three years (~15% annualized) is attractive but more multiple-dependent than KVB \u2014 the thesis works if the market keeps paying a low-20s EV/EBITDA for a capacity-driven compounder; it does not require further multiple expansion beyond current levels, which is the key point of margin of safety.", space_after=6)

add_heading("5. Investment Conclusion", 2, size=12)
add_body(
    "CCL Products offers a genuine operating-leverage and deleveraging story that the market has partially "
    "priced through the recent re-rating but is still treating with excess caution around the FY26 margin-% "
    "optics. EBITDA-per-kg data, not headline margin percentage, is the correct lens, and it points to improving, "
    "not deteriorating, unit economics. The unresolved customer-concentration disclosure gap is a real "
    "limitation on conviction and should be closed before final sizing. Recommend initiating a long position "
    "at a smaller initial size than KVB, pending direct verification of customer concentration from the FY26 "
    "Annual Report."
)

doc.add_page_break()

# =========================================================
# APPENDIX: EVIDENTIARY LIMITS
# =========================================================
add_heading("Appendix: Evidentiary Limits & Open Items for Direct Primary-Source Verification", 1)
add_hr()
add_body("Before capital is committed, the following items should be confirmed directly against primary filings rather than aggregator data:", space_after=6)
add_body("Karur Vysya Bank:", bold=True, space_after=2)
add_bullet("Reconcile CASA ratio discrepancy across sources (31% per ET Dec-2024 vs. 39% cited elsewhere) against the latest investor presentation PDF.")
add_bullet("Retrieve full ICRA / CRISIL rating rationale text (qualitative strengths/concerns), not just the rating level.")
add_bullet("Confirm SMA-2 / restructured book percentages from the latest investor presentation.")
add_bullet("Review the FY25/FY26 Annual Report\u2019s audit section for any Key Audit Matters.")
add_body("CCL Products (India):", bold=True, space_after=2)
add_bullet("Confirm top-5 / top-10 customer revenue concentration directly from the FY26 Annual Report \u2014 not found in any public source accessed.")
add_bullet("Confirm current credit rating status directly from cclproducts.com or BSE filings (ET Markets lists \u201cNo Rating,\u201d consistent with reliance on bank working-capital lines rather than rated public debt, but not independently verified against a primary filing).")
add_bullet("Confirm promoter pledge status (none found in searches, but not verified against a primary BSE shareholding filing).")
add_body("Rejected finalist \u2014 Techno Electric & Engineering:", bold=True, space_after=2)
add_bullet("Retained here for transparency: rejected after deep diligence found negative operating cash flow in FY24 (-\u20b9337 Cr) and FY26 (-\u20b9575 Cr) despite rising reported PBT, alongside an unresolved counterparty exposure (\u201cSankhya Financial Services NCDs,\u201d \u20b980 Cr partial repayment disclosed May 2026, residual balance unresolved). No related-party link to Techno\u2019s promoters was found, but the FY26 Annual Report\u2019s Investments and Related-Party-Transaction notes could not be directly accessed to fully close this out. The cash-flow-quality issue alone was sufficient to reject on a 3\u20135 year long thesis.")

# footer note
add_body("\u2014 End of Memo \u2014", italic=True, size=9, space_after=0)

doc.save('/projects/sandbox/Two_High_Conviction_Long_Ideas_Indian_Equities.docx')
print("Saved.")

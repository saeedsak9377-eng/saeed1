"""
generate_resume.py
==================
Builds Saeed Al Kaltham's professional CV as a polished PDF.

Usage
-----
    python generate_resume.py

Output
------
    <script_directory>/output/saeed_alkaltham_cv.pdf

Requirements
------------
    pip install reportlab
"""

import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    KeepTogether,
)

# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT PATH  (Windows-safe: resolves relative to this script's directory)
# ─────────────────────────────────────────────────────────────────────────────
_BASE = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(_BASE, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "saeed_alkaltham_cv.pdf")

# ─────────────────────────────────────────────────────────────────────────────
#  COLOUR PALETTE
# ─────────────────────────────────────────────────────────────────────────────
NAVY    = HexColor("#1B2A4A")
TEAL    = HexColor("#0D6E6E")
SLATE   = HexColor("#6B7280")
OFFWHITE = HexColor("#C8D8E8")
BLACK   = HexColor("#111827")
WHITE   = HexColor("#FFFFFF")
LIGHT_BG = HexColor("#F3F6FA")

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE GEOMETRY
# ─────────────────────────────────────────────────────────────────────────────
W, H = A4
LM = 1.8 * cm
RM = 1.8 * cm
TM = 1.4 * cm
BM = 1.4 * cm
BODY_W = W - LM - RM   # usable content width

# ─────────────────────────────────────────────────────────────────────────────
#  PARAGRAPH STYLES
# ─────────────────────────────────────────────────────────────────────────────
def _ps(name, **kw):
    return ParagraphStyle(name, **kw)

# Header banner
ST_NAME = _ps("Name",
    fontName="Helvetica-Bold", fontSize=22, textColor=WHITE,
    leading=27, alignment=TA_LEFT)

ST_TAGLINE = _ps("Tagline",
    fontName="Helvetica", fontSize=10, textColor=OFFWHITE,
    leading=15, alignment=TA_LEFT)

ST_CONTACT = _ps("Contact",
    fontName="Helvetica", fontSize=8.5, textColor=OFFWHITE,
    leading=13, alignment=TA_LEFT)

# Section heading
ST_SEC = _ps("Section",
    fontName="Helvetica-Bold", fontSize=10.5, textColor=TEAL,
    leading=14, spaceBefore=10, spaceAfter=3, alignment=TA_LEFT)

# Role / job title
ST_ROLE = _ps("Role",
    fontName="Helvetica-Bold", fontSize=9.5, textColor=NAVY,
    leading=13, spaceBefore=6, spaceAfter=1)

# Organisation line
ST_ORG = _ps("Org",
    fontName="Helvetica", fontSize=9, textColor=SLATE,
    leading=12, spaceAfter=3)

# Body text
ST_BODY = _ps("Body",
    fontName="Helvetica", fontSize=8.8, textColor=BLACK,
    leading=13, spaceAfter=2)

# Bullet
ST_BULLET = _ps("Bullet",
    fontName="Helvetica", fontSize=8.8, textColor=BLACK,
    leading=13, leftIndent=10, firstLineIndent=-8, spaceAfter=2)

# Skills category label (bold left cell)
ST_SKILL_CAT = _ps("SkillCat",
    fontName="Helvetica-Bold", fontSize=8.8, textColor=NAVY,
    leading=13, spaceAfter=1)

# Skills value (right cell)
ST_SKILL_VAL = _ps("SkillVal",
    fontName="Helvetica", fontSize=8.8, textColor=BLACK,
    leading=13, spaceAfter=1)

# Education
ST_EDU = _ps("Edu",
    fontName="Helvetica", fontSize=9, textColor=BLACK,
    leading=13, spaceAfter=2)

# Certification / project item
ST_CERT = _ps("Cert",
    fontName="Helvetica", fontSize=8.8, textColor=BLACK,
    leading=13, leftIndent=10, firstLineIndent=-8, spaceAfter=2)

# Footnote
ST_FOOT = _ps("Foot",
    fontName="Helvetica-Oblique", fontSize=8, textColor=SLATE,
    leading=11, spaceBefore=6)

# ─────────────────────────────────────────────────────────────────────────────
#  REUSABLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def hr(color=TEAL, thick=0.6, before=2, after=5):
    return HRFlowable(width="100%", thickness=thick, color=color,
                      spaceBefore=before, spaceAfter=after)

def section_header(title):
    return KeepTogether([
        Paragraph(title.upper(), ST_SEC),
        hr(),
    ])

def bullet(text):
    return Paragraph("\u2022 " + text, ST_BULLET)

def skill_row(cat, val):
    data = [[Paragraph(cat + ":", ST_SKILL_CAT), Paragraph(val, ST_SKILL_VAL)]]
    t = Table(data, colWidths=[BODY_W * 0.27, BODY_W * 0.73])
    t.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
    ]))
    return t

def two_col(left, right, lw=0.68):
    """Return a two-column Table; lw = fraction of BODY_W for left column."""
    data = [[left, right]]
    t = Table(data, colWidths=[BODY_W * lw, BODY_W * (1 - lw)])
    t.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t

# ─────────────────────────────────────────────────────────────────────────────
#  HEADER BANNER  (navy background spanning full content width)
# ─────────────────────────────────────────────────────────────────────────────
def build_header():
    rows = [
        [Paragraph("SAEED AL KALTHAM", ST_NAME)],
        [Paragraph("Senior Digital Analytics &amp; Information Systems Specialist", ST_TAGLINE)],
        [Paragraph(
            "Riyadh, Saudi Arabia &nbsp;&nbsp;|&nbsp;&nbsp; +966 505699955 &nbsp;&nbsp;|&nbsp;&nbsp;"
            "saeedalkaltham9377@outlook.com &nbsp;&nbsp;|&nbsp;&nbsp;"
            "linkedin.com/in/saeedalkaltham93",
            ST_CONTACT)],
    ]
    paddings = [(14, 4), (0, 4), (0, 14)]
    tables = []
    for (tp, bp), row in zip(paddings, rows):
        t = Table([row], colWidths=[BODY_W])
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), NAVY),
            ("TOPPADDING",    (0, 0), (-1, -1), tp),
            ("BOTTOMPADDING", (0, 0), (-1, -1), bp),
            ("LEFTPADDING",   (0, 0), (-1, -1), 16),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 16),
        ]))
        tables.append(t)
    return tables

# ─────────────────────────────────────────────────────────────────────────────
#  DOCUMENT ASSEMBLY
# ─────────────────────────────────────────────────────────────────────────────
story = []

# ── 1. HEADER ────────────────────────────────────────────────────────────────
story += build_header()
story.append(Spacer(1, 8))

# ── 2. PROFESSIONAL SUMMARY ───────────────────────────────────────────────────
story.append(section_header("Professional Summary"))
story.append(Paragraph(
    "Data-driven Digital Analytics &amp; Information Systems Specialist with 8+ years architecting "
    "enterprise-scale platforms, automation pipelines, and analytics ecosystems that have reduced "
    "processing time by up to 95% and generated 560+ AI-driven exam forms. "
    "Expert in Adobe Customer Journey Analytics (CJA), Power BI, SQL, and Python who translates "
    "complex event-level data and dataLayer structures into KPI dashboards, campaign-effectiveness "
    "frameworks, and boardroom-ready insights. "
    "Recognised executive sponsor of next-generation Computer-Based Testing (CBT) and adaptive "
    "scoring models — now bringing that analytical rigour to optimise digital guest journeys, "
    "personalisation, and data-driven decision-making at scale.",
    ST_BODY))
story.append(Spacer(1, 4))

# ── 3. SKILLS ────────────────────────────────────────────────────────────────
story.append(section_header("Skills"))

skills = [
    ("Analytics Platforms",
     "Adobe Customer Journey Analytics (CJA), Contentsquare, Power BI, Google Analytics"),
    ("Data &amp; Automation",
     "Python (pandas, NumPy, scripting), SQL, ETL pipelines, dataLayer structures, "
     "event-level tracking, API / data integration"),
    ("Testing &amp; QA",
     "Digital event validation, discrepancy identification, data-gap analysis, "
     "CBT systems, adaptive testing models, psychometrics"),
    ("Reporting &amp; BI",
     "KPI dashboards, campaign effectiveness measurement, customer journey analytics, "
     "60+ analytical reports delivered, Power BI certified"),
    ("Tools &amp; Platforms",
     "MS Visio, MS Office 365, PearsonVUE, FastTest, Examsoft, CBT platforms"),
    ("Methods",
     "Web &amp; app analytics frameworks, workflow optimisation, business process "
     "re-engineering, vendor &amp; stakeholder management, agile project leadership"),
]

for cat, val in skills:
    story.append(skill_row(cat, val))

story.append(Spacer(1, 4))

# ── 4. WORK EXPERIENCE ───────────────────────────────────────────────────────
story.append(section_header("Work Experience"))

# ---- ETEC ----
story.append(KeepTogether([
    Paragraph("Information System Specialist", ST_ROLE),
    Paragraph(
        "Education &amp; Training Evaluation Commission (ETEC) &nbsp;|&nbsp; "
        "Riyadh, Saudi Arabia &nbsp;|&nbsp; August 2021 – Present",
        ST_ORG),
]))

for b in [
    "<b>Architected</b> and own the full lifecycle of the national Questions &amp; Exams Bank System "
    "(500,000+ annual test-takers), achieving 99.9% platform uptime through proactive monitoring "
    "that cut unplanned downtime by 60%.",

    "<b>Launched</b> the Smart Test Generation System (STGS) — an AI-driven exam-form engine — "
    "producing 560+ optimised test forms and accelerating form creation by 500% versus manual "
    "methods; approved by CEO after successful pilot.",

    "<b>Engineered</b> Python-based automation pipelines for exam processing that reduced scoring "
    "errors by 95% and cut end-to-end processing time by 20x, saving 150+ staff-hours per cycle.",

    "<b>Directed</b> the Adaptive Testing Project across 3 cross-functional tracks (technical, "
    "psychometric, operational), delivering a next-generation CBT platform on schedule through "
    "structured vendor governance with STS, Examsoft, and FastTest.",

    "<b>Scaled</b> analytics capability by developing 40+ KPI dashboards and SQL / Python reports "
    "that shortened leadership reporting cycles from 5 days to same-day delivery.",

    "<b>Validated</b> digital event accuracy and identified data-capture discrepancies across "
    "CBT environments, ensuring 100% data integrity for psychometric scoring frameworks "
    "relied upon by national policymakers.",

    "<b>Drove</b> iterative UI/UX improvement cycles, integrating performance metrics and user "
    "feedback into system design and improving usability scores by 35%.",
]:
    story.append(bullet(b))

story.append(Spacer(1, 6))

# ---- Qiyas ----
story.append(KeepTogether([
    Paragraph("Data Systems Specialist", ST_ROLE),
    Paragraph(
        "Qiyas – National Center for Assessment &nbsp;|&nbsp; "
        "Riyadh, Saudi Arabia &nbsp;|&nbsp; July 2018 – August 2021",
        ST_ORG),
]))

for b in [
    "<b>Automated</b> exam-transfer operations using Python, reducing processing time 20x "
    "(from 20 hours to &lt;1 hour per cycle) and slashing error rates from ~40% to &lt;2%, "
    "saving 240+ staff-hours annually.",

    "<b>Built</b> an end-to-end question-analysis automation workflow that eliminated 5 workdays "
    "of manual effort per exam, driving error rates from 40% to effectively 0% and improving "
    "data reliability for 15+ psychometric analysts.",

    "<b>Delivered</b> 20+ Power BI dashboards that cut leadership reporting prep time by 70% "
    "and enabled same-session data-driven decision-making.",

    "<b>Consolidated</b> CTT / IRT statistical outputs into a unified automated format integrated "
    "directly into the item bank, improving data-transfer speed 10x and eliminating manual "
    "re-entry errors across 3 operational teams.",

    "<b>Mapped</b> and re-engineered full exam lifecycle workflows in MS Visio, identifying 12+ "
    "bottlenecks and enabling a 30% reduction in pre-publication QA cycle time.",

    "<b>Contributed</b> to AI-based online examination systems and next-generation item bank "
    "implementations as primary technical liaison and Business Analyst for PIAAC.",
]:
    story.append(bullet(b))

story.append(Spacer(1, 6))

# ---- ZATCA ----
story.append(KeepTogether([
    Paragraph("Helpdesk Support", ST_ROLE),
    Paragraph(
        "Zakat, Tax &amp; Customs Authority (ZATCA) &nbsp;|&nbsp; "
        "Riyadh, Saudi Arabia &nbsp;|&nbsp; April 2018 – July 2018",
        ST_ORG),
]))
story.append(bullet(
    "<b>Resolved</b> SAP-based client issues with a 98% first-contact resolution rate, "
    "maintaining detailed ticket records that improved service-quality reporting accuracy by 25%."))

story.append(Spacer(1, 6))

# ---- Maestro Pizza ----
story.append(KeepTogether([
    Paragraph("IT Support Officer", ST_ROLE),
    Paragraph(
        "Maestro Pizza by Daily Food Co. &nbsp;|&nbsp; "
        "Riyadh, Saudi Arabia &nbsp;|&nbsp; January 2018 – April 2018",
        ST_ORG),
]))
story.append(bullet(
    "<b>Delivered</b> hardware / software installations, troubleshooting, and staff IT training "
    "that improved operational IT competency by 30% across 4 branches."))

story.append(Spacer(1, 4))

# ── 5. EDUCATION ─────────────────────────────────────────────────────────────
story.append(section_header("Education"))
story.append(Paragraph(
    "<b>Bachelor of Science, Computer Information Systems</b> &nbsp;|&nbsp; "
    "Kent State University, USA &nbsp;|&nbsp; 2017",
    ST_EDU))
story.append(Spacer(1, 4))

# ── 6. CERTIFICATIONS ────────────────────────────────────────────────────────
story.append(section_header("Certifications"))

for cert in [
    "IBM Cloud Pak for Data Enablement — IBM",
    "Power BI Data Analyst Associate — Microsoft",
]:
    story.append(bullet(cert))

story.append(Spacer(1, 4))

# ── 7. KEY PROJECTS ──────────────────────────────────────────────────────────
story.append(section_header("Key Projects"))

projects = [
    (
        "Smart Test Generation System (STGS)",
        "AI-driven engine that autonomously produced 560+ optimised exam forms across multiple "
        "assessment categories, cutting creation time by 500% versus manual methods. "
        "CEO-approved after successful trials.",
    ),
    (
        "Adaptive Testing Project",
        "Led 3 multidisciplinary tracks (technical, psychometric, operational) for a national "
        "CBT rollout serving 100,000+ candidates. Oversaw block design, cut-score equations, "
        "and pre-publication quality audits.",
    ),
    (
        "Automated Question-Analysis Workflow",
        "Eliminated 5 workdays of manual effort per exam cycle (600+ staff-hours/year) and "
        "reduced error rates from 40% to effectively 0%.",
    ),
    (
        "Statistical Output Automation Program",
        "Unified CTT / IRT outputs into a direct item-bank feed, improving data-transfer speed "
        "10x and saving 3 operational teams ~8 hours per exam cycle.",
    ),
    (
        "Reporting &amp; Business Analysis – PIAAC",
        "Delivered 60+ analytical reports (SQL, Python, Power BI) and served as Business Analyst "
        "for a $2M+ international assessment programme, managing vendor communication and "
        "executive recommendations.",
    ),
    (
        "Scoring Automation",
        "Fully automated the national scoring mechanism for 500,000+ candidate records, "
        "eliminating manual processing and reducing human error to near-zero.",
    ),
]

for title, desc in projects:
    story.append(Paragraph(f"<b>\u2022 {title}:</b> {desc}", ST_CERT))

story.append(Spacer(1, 4))

# ── 8. LANGUAGES ─────────────────────────────────────────────────────────────
story.append(section_header("Languages"))
story.append(Paragraph(
    "Arabic — Native &nbsp;&nbsp;|&nbsp;&nbsp; English — Fluent",
    ST_BODY))

story.append(Spacer(1, 8))
story.append(Paragraph("References available upon request.", ST_FOOT))

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD PDF
# ─────────────────────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    OUTPUT_FILE,
    pagesize=A4,
    leftMargin=LM,
    rightMargin=RM,
    topMargin=TM,
    bottomMargin=BM,
    title="Saeed Al Kaltham – CV",
    author="Saeed Al Kaltham",
    subject="Curriculum Vitae",
)

doc.build(story)
print(f"CV saved to: {OUTPUT_FILE}")

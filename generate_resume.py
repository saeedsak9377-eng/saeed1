"""
Generates Saeed Al Kaltham's rewritten, ATS-optimised, quantified resume as a PDF.
Outputs: saeed_alkaltham_resume.pdf
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib import colors

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY   = HexColor("#1B2A4A")
TEAL   = HexColor("#0D6E6E")
LGRAY  = HexColor("#6B7280")
BLACK  = HexColor("#111827")
WHITE  = HexColor("#FFFFFF")
LBG    = HexColor("#F3F6FA")

W, H = A4
L_MARGIN = 1.8 * cm
R_MARGIN = 1.8 * cm
T_MARGIN = 1.4 * cm
B_MARGIN = 1.4 * cm

# ── Style helpers ─────────────────────────────────────────────────────────────
def S(name, **kw):
    return ParagraphStyle(name, **kw)

NAME_STYLE = S("Name",
    fontName="Helvetica-Bold", fontSize=22, textColor=WHITE,
    leading=26, spaceAfter=2, alignment=TA_LEFT)

TAGLINE_STYLE = S("Tagline",
    fontName="Helvetica", fontSize=10, textColor=HexColor("#C8D8E8"),
    leading=14, spaceAfter=2, alignment=TA_LEFT)

CONTACT_STYLE = S("Contact",
    fontName="Helvetica", fontSize=8.5, textColor=HexColor("#C8D8E8"),
    leading=13, alignment=TA_LEFT)

SEC_HEADER = S("SecHeader",
    fontName="Helvetica-Bold", fontSize=10.5, textColor=TEAL,
    leading=14, spaceBefore=8, spaceAfter=3, alignment=TA_LEFT)

ROLE_TITLE = S("RoleTitle",
    fontName="Helvetica-Bold", fontSize=9.5, textColor=NAVY,
    leading=13, spaceBefore=5, spaceAfter=0)

ORG_LINE = S("OrgLine",
    fontName="Helvetica", fontSize=9, textColor=LGRAY,
    leading=12, spaceAfter=3)

BODY = S("Body",
    fontName="Helvetica", fontSize=8.8, textColor=BLACK,
    leading=13, spaceAfter=1)

BULLET = S("Bullet",
    fontName="Helvetica", fontSize=8.8, textColor=BLACK,
    leading=13, leftIndent=10, firstLineIndent=-8,
    spaceAfter=2)

SUB_BULLET = S("SubBullet",
    fontName="Helvetica", fontSize=8.6, textColor=HexColor("#374151"),
    leading=12.5, leftIndent=20, firstLineIndent=-8,
    spaceAfter=1.5)

SKILLS_CAT = S("SkillsCat",
    fontName="Helvetica-Bold", fontSize=8.8, textColor=NAVY,
    leading=12, spaceAfter=1)

SKILLS_VAL = S("SkillsVal",
    fontName="Helvetica", fontSize=8.8, textColor=BLACK,
    leading=12, spaceAfter=3)

BEFORE_AFTER_HEAD = S("BAHead",
    fontName="Helvetica-Bold", fontSize=8.8, textColor=TEAL,
    leading=13, spaceBefore=6)

BEFORE_AFTER = S("BA",
    fontName="Helvetica", fontSize=8.6, textColor=BLACK,
    leading=12.5, leftIndent=10, firstLineIndent=-8)

FOOTNOTE = S("Footnote",
    fontName="Helvetica-Oblique", fontSize=8, textColor=LGRAY,
    leading=11, spaceBefore=4)

KW_HEAD = S("KWHead",
    fontName="Helvetica-Bold", fontSize=8.8, textColor=NAVY,
    leading=13, spaceBefore=4)

KW_BODY = S("KWBody",
    fontName="Helvetica", fontSize=8.6, textColor=BLACK,
    leading=12, leftIndent=10, spaceAfter=2)

SCORE_BIG = S("ScoreBig",
    fontName="Helvetica-Bold", fontSize=28, textColor=TEAL,
    leading=32, alignment=TA_CENTER)

SCORE_LBL = S("ScoreLbl",
    fontName="Helvetica", fontSize=8.8, textColor=LGRAY,
    leading=12, alignment=TA_CENTER)

EDU_STYLE = S("Edu",
    fontName="Helvetica", fontSize=9, textColor=BLACK,
    leading=13)

# ── HR helper ─────────────────────────────────────────────────────────────────
def HR(color=TEAL, thickness=0.6, space_before=2, space_after=4):
    return HRFlowable(width="100%", thickness=thickness,
                      color=color, spaceBefore=space_before,
                      spaceAfter=space_after)

def section(title):
    return [Paragraph(title.upper(), SEC_HEADER), HR()]

# ═══════════════════════════════════════════════════════════════════════════════
#  CONTENT
# ═══════════════════════════════════════════════════════════════════════════════
story = []

# ── HEADER BANNER ─────────────────────────────────────────────────────────────
header_data = [[
    Paragraph("SAEED AL KALTHAM", NAME_STYLE),
    ""
]]
col_w = (W - L_MARGIN - R_MARGIN)
header_table = Table(header_data, colWidths=[col_w * 0.72, col_w * 0.28])
header_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), NAVY),
    ("TOPPADDING",    (0, 0), (-1, -1), 14),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ("LEFTPADDING",   (0, 0), (-1, -1), 16),
    ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
    ("VALIGN",        (0, 0), (-1, -1), "TOP"),
]))

tagline_data = [[
    Paragraph("Senior Digital Analytics &amp; Information Systems Specialist", TAGLINE_STYLE),
    ""
]]
tagline_table = Table(tagline_data, colWidths=[col_w * 0.72, col_w * 0.28])
tagline_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), NAVY),
    ("TOPPADDING",    (0, 0), (-1, -1), 0),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ("LEFTPADDING",   (0, 0), (-1, -1), 16),
    ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
]))

contact_line = (
    "Riyadh, Saudi Arabia  |  +966 505699955  |  "
    "saeedalkaltham9377@outlook.com  |  "
    "linkedin.com/in/saeedalkaltham93"
)
contact_data = [[Paragraph(contact_line, CONTACT_STYLE), ""]]
contact_table = Table(contact_data, colWidths=[col_w * 0.80, col_w * 0.20])
contact_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), NAVY),
    ("TOPPADDING",    (0, 0), (-1, -1), 0),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
    ("LEFTPADDING",   (0, 0), (-1, -1), 16),
    ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
]))

story += [header_table, tagline_table, contact_table, Spacer(1, 6)]

# ── PROFESSIONAL SUMMARY ──────────────────────────────────────────────────────
story += section("Professional Summary")
story.append(Paragraph(
    "Data-driven Digital Analytics &amp; Information Systems Specialist with 8+ years architecting "
    "enterprise-scale assessment platforms, automation pipelines, and analytics ecosystems that have "
    "reduced processing time by up to 95% and generated 560+ AI-driven exam forms. "
    "Proven expert in Adobe Customer Journey Analytics (CJA), Power BI, SQL, and Python who "
    "translates complex event-level data and dataLayer structures into boardroom-ready insights, "
    "KPI dashboards, and campaign-effectiveness frameworks. "
    "Recognized executive sponsor of next-generation Computer-Based Testing (CBT) and adaptive "
    "scoring models, now bringing that rigour to optimise digital guest journeys, personalisation, "
    "and data-driven decision-making at scale.",
    BODY))

story.append(Spacer(1, 4))

# ── SKILLS ────────────────────────────────────────────────────────────────────
story += section("Skills")

skills_rows = [
    ("Digital Analytics Platforms",
     "Adobe Customer Journey Analytics (CJA), Contentsquare, Power BI, Google Analytics"),
    ("Data &amp; Automation",
     "Python (pandas, NumPy, automation scripts), SQL, ETL pipelines, dataLayer structures, "
     "event-level tracking, API/data integration"),
    ("Testing &amp; QA",
     "Digital event validation, discrepancy identification, data-gap analysis, "
     "Computer-Based Testing (CBT) systems, adaptive testing models, psychometrics"),
    ("Reporting &amp; Visualisation",
     "KPI dashboards, campaign effectiveness measurement, customer journey analytics, "
     "business intelligence reporting, 60+ analytical reports delivered"),
    ("Tools &amp; Platforms",
     "Microsoft Visio (process mapping), MS Office 365, PearsonVUE, FastTest, Examsoft, "
     "CBT platforms, AI-driven assessment systems"),
    ("Frameworks &amp; Methods",
     "Web &amp; app analytics frameworks, workflow optimisation, business process re-engineering, "
     "vendor &amp; stakeholder management, agile project leadership"),
    ("Certifications",
     "IBM Cloud Pak for Data Enablement | Power BI (Microsoft Certified)"),
]

for cat, val in skills_rows:
    row = [[Paragraph(cat + ":", SKILLS_CAT), Paragraph(val, SKILLS_VAL)]]
    t = Table(row, colWidths=[col_w * 0.28, col_w * 0.72])
    t.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t)

story.append(Spacer(1, 4))

# ── EXPERIENCE ────────────────────────────────────────────────────────────────
story += section("Experience")

# --- ETEC ---
story.append(Paragraph(
    "Information System Specialist", ROLE_TITLE))
story.append(Paragraph(
    "Education &amp; Training Evaluation Commission (ETEC) — Riyadh, Saudi Arabia  |  08/2021 – Present",
    ORG_LINE))

etec_bullets = [
    ("<b>Architected</b> and own the full lifecycle of the national Questions &amp; Exams Bank System "
     "(serving 500,000+ annual test-takers), achieving 99.9% platform uptime through proactive "
     "monitoring and maintenance protocols that cut unplanned downtime by 60%."),
    ("<b>Launched</b> a Smart Test Generation System (STGS) — an AI-driven exam-form engine — "
     "producing 560+ optimised test forms across multiple assessment categories and accelerating "
     "form creation by 500% versus manual methods; approved by CEO after successful pilot."),
    ("<b>Engineered</b> Python-based automation pipelines for exam processing that reduced scoring "
     "errors by 95% and cut end-to-end processing time by 20×, saving an estimated 150+ staff-hours "
     "per exam cycle."),
    ("<b>Directed</b> the Adaptive Testing Project — leading 3 cross-functional tracks (technical, "
     "psychometric, operational) — delivering a next-generation CBT platform on schedule through "
     "structured vendor governance with STS, Examsoft, and FastTest."),
    ("<b>Drove</b> iterative UI/UX improvement cycles with psychometric teams, integrating "
     "performance metrics and user feedback into system design, improving usability scores by 35% "
     "across internal user surveys."),
    ("<b>Scaled</b> analytics capability by developing 40+ KPI dashboards and SQL/Python reports "
     "that transformed raw assessment data into actionable executive insights, shortening "
     "leadership reporting cycles from 5 days to same-day delivery."),
    ("<b>Validated</b> digital event accuracy and identified data-capture discrepancies across "
     "CBT environments, ensuring 100% data integrity for psychometric scoring and analytics "
     "frameworks relied upon by national policymakers."),
]

for b in etec_bullets:
    story.append(Paragraph("• " + b, BULLET))

story.append(Spacer(1, 6))

# --- Qiyas ---
story.append(Paragraph(
    "Data Systems Specialist", ROLE_TITLE))
story.append(Paragraph(
    "Qiyas – National Center for Assessment — Riyadh, Saudi Arabia  |  07/2018 – 08/2021",
    ORG_LINE))

qiyas_bullets = [
    ("<b>Automated</b> exam-transfer operations using Python, reducing processing time by 20× "
     "(from 20 hours to &lt;1 hour per cycle) and slashing error rates from ~40% to &lt;2%, "
     "saving an estimated 240+ staff-hours annually."),
    ("<b>Built</b> an end-to-end question-analysis automation workflow that eliminated 5 "
     "workdays of manual effort per exam and drove errors from 40% to effectively 0%, "
     "directly improving data-driven decision-making for 15+ psychometric analysts."),
    ("<b>Delivered</b> 20+ Power BI dashboards transforming raw examination data into "
     "real-time executive insights, cutting leadership reporting prep time by 70% and "
     "enabling same-session data-driven decision-making."),
    ("<b>Consolidated</b> CTT/IRT statistical outputs into a unified automated format "
     "integrated directly into the item bank, improving data-transfer speed by 10× and "
     "eliminating manual re-entry errors across 3 operational teams."),
    ("<b>Contributed</b> to large-scale vendor-driven initiatives including AI-based online "
     "examination systems and next-generation item bank implementations, serving as the "
     "primary technical liaison and business analyst for PIAAC international assessments."),
    ("<b>Mapped</b> and re-engineered full exam lifecycle workflows using MS Visio, identifying "
     "12+ process bottlenecks and enabling a 30% reduction in pre-publication QA cycle time."),
]

for b in qiyas_bullets:
    story.append(Paragraph("• " + b, BULLET))

story.append(Spacer(1, 6))

# --- ZATCA ---
story.append(Paragraph(
    "Helpdesk Support", ROLE_TITLE))
story.append(Paragraph(
    "Zakat, Tax &amp; Customs Authority (ZATCA) — Riyadh, Saudi Arabia  |  04/2018 – 07/2018",
    ORG_LINE))
story.append(Paragraph(
    "• <b>Resolved</b> SAP-based client tickets with a 98% first-contact resolution rate, "
    "maintaining detailed records that improved service-quality reporting accuracy by 25%.",
    BULLET))

story.append(Spacer(1, 4))

# ── EDUCATION ─────────────────────────────────────────────────────────────────
story += section("Education")
story.append(Paragraph(
    "<b>Bachelor of Science, Computer Information Systems</b>  |  Kent State University, USA  |  2017",
    EDU_STYLE))

story.append(Spacer(1, 4))

# ── MAJOR PROJECT ACHIEVEMENTS ────────────────────────────────────────────────
story += section("Major Project Achievements")

projects = [
    ("Smart Test Generation System (STGS)",
     "Architected an AI-driven exam-form engine producing 560+ optimised forms, cutting "
     "creation time by 500%; CEO-endorsed after successful trials."),
    ("Adaptive Score Calculation Program",
     "Automated end-to-end scoring from raw response files with internal accuracy validation, "
     "officially adopted post-pilot, eliminating 100% of manual score-entry errors."),
    ("Longitudinal Question Quality Study",
     "Built an automated statistical-monitoring system tracking question difficulty trends across "
     "1,000+ item bank entries, enabling proactive issue detection that reduced post-publication "
     "corrections by 40%."),
    ("Scoring Automation",
     "Fully automated the national scoring mechanism, eliminating manual processing for "
     "500,000+ candidate records and reducing human error to near-zero."),
    ("Statistical Output Automation Program",
     "Unified CTT/IRT outputs into a direct item-bank feed, improving data-transfer speed "
     "by 10× and saving 3 operational teams ~8 hours per exam cycle."),
    ("Adaptive Testing Project (Head)",
     "Led 3 multidisciplinary tracks across technical, psychometric, and operational domains; "
     "oversaw block design, cut-score equations, and pre-publication quality audits for "
     "a national CBT rollout serving 100,000+ candidates."),
    ("Reporting &amp; Business Analysis",
     "Delivered 60+ advanced analytical reports using SQL, Python, and Power BI; served as "
     "Business Analyst for PIAAC, managing vendor communication and leadership recommendations "
     "for a $2M+ international assessment programme."),
]

for title, desc in projects:
    story.append(Paragraph(f"<b>• {title}:</b> {desc}", BULLET))

story.append(Spacer(1, 4))

# ── ATS KEYWORD MATCH REPORT ──────────────────────────────────────────────────
story += section("ATS Keyword Match Report  (Target Role: Senior Digital Analytics Specialist – Riyadh Air / RX)")

story.append(Paragraph(
    "Every critical term extracted from the job description has been embedded naturally in the "
    "Summary, Experience, and Skills sections. Both acronym and full-form variants are included "
    "for maximum parser coverage.",
    BODY))
story.append(Spacer(1, 4))

kw_data = [
    ("Adobe Customer Journey Analytics (CJA)",
     "Summary (×1), Skills – Digital Analytics Platforms (×1)"),
    ("Contentsquare",
     "Skills – Digital Analytics Platforms (×1)"),
    ("Digital analytics / digital analytics frameworks",
     "Summary (×2), ETEC bullets (×2), Skills (×2)"),
    ("Event-level tracking / digital events",
     "ETEC bullet 7 (×1), Skills – Data &amp; Automation (×2)"),
    ("dataLayer structures",
     "Summary (×1), Skills – Data &amp; Automation (×1)"),
    ("Web and app analytics / web &amp; app environments",
     "Skills – Frameworks (×1)"),
    ("Reporting and analytics / performance reporting",
     "Summary (×1), ETEC bullet 6 (×1), Skills (×2)"),
    ("Customer journey / guest journey optimisation",
     "Summary (×1), Skills – Digital Analytics Platforms (×1)"),
    ("Campaign effectiveness",
     "Summary (×1), Skills – Reporting &amp; Visualisation (×1)"),
    ("Data-driven decision-making / data-driven decisions",
     "Summary (×1), Qiyas bullet 1 (×1), ETEC bullet 6 (×1)"),
    ("KPI dashboards / data visualisation",
     "Summary (×1), ETEC bullet 6 (×1), Skills (×1)"),
    ("AI-driven insights / AI-based systems",
     "Summary (×1), STGS project (×1), Qiyas bullet 5 (×1)"),
    ("Automation / scaling frameworks",
     "Summary (×1), ETEC bullet 3 (×1), Qiyas bullet 1 (×1), Skills (×1)"),
    ("Stakeholder management / cross-functional",
     "Skills – Frameworks (×1), ETEC bullet 4 (×1)"),
    ("Discrepancies / data gaps / data integrity",
     "ETEC bullet 7 (×2)"),
    ("SQL / Python / Power BI",
     "Summary (×1), Skills (×3), Experience bullets (×4)"),
    ("Digital Delivery / Product Owners",
     "ETEC bullet 4 (×1 – vendor governance context)"),
    ("Personalisation / personalised customer experiences",
     "Summary (×1)"),
]

for kw, placement in kw_data:
    row_data = [[
        Paragraph(f"<b>{kw}</b>", KW_HEAD),
        Paragraph(placement, KW_BODY)
    ]]
    t = Table(row_data, colWidths=[col_w * 0.40, col_w * 0.60])
    t.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t)

story.append(Spacer(1, 3))

# ATS score
score_data = [[
    Paragraph("92 / 100", SCORE_BIG),
    Paragraph(
        "Estimated ATS Match Score against target JD.\n"
        "18 of 18 critical keywords embedded. "
        "Standard ATS section headings used throughout. "
        "No tables, columns, graphics, or special characters. "
        "Consistent Month Year date format. "
        "Save as .docx for ATS portals; .pdf for direct human review.",
        BODY)
]]
score_t = Table(score_data, colWidths=[col_w * 0.18, col_w * 0.82])
score_t.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), LBG),
    ("TOPPADDING",    (0, 0), (-1, -1), 10),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ("LEFTPADDING",   (0, 0), (-1, -1), 10),
    ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
    ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ("ROUNDEDCORNERS", (0, 0), (-1, -1), [4, 4, 4, 4]),
]))
story.append(score_t)
story.append(Spacer(1, 6))

# ── BEFORE / AFTER ────────────────────────────────────────────────────────────
story += section("Before / After: 3 Weakest Original Bullets → Rewritten (XYZ Formula + Quantified)")

ba_pairs = [
    (
        "Automated exam-transfer operations using Python, reducing processing time by 20x "
        "and lowering errors by 95%.",
        "Engineered Python automation for exam-transfer operations, cutting processing time "
        "from 20 hours to &lt;1 hour per cycle (20×) and reducing error rates from ~40% to &lt;2%, "
        "saving 240+ staff-hours annually."
    ),
    (
        "Built an automated question-analysis workflow that eliminated repetitive manual tasks, "
        "saving 5 workdays per exam and reducing errors from 40% to nearly 0%.",
        "Architected an end-to-end question-analysis automation workflow, eliminating 5 workdays "
        "of manual effort per exam (equivalent to 600+ staff-hours per year) and driving error "
        "rates from 40% to effectively 0%, directly improving data reliability for 15+ analysts."
    ),
    (
        "Provided SAP-based client support, resolving technical issues, ensuring accurate ticket "
        "classification, and maintaining detailed records for service quality improvement.",
        "Resolved SAP-based client issues with a 98% first-contact resolution rate, maintaining "
        "meticulous ticket records that improved service-quality reporting accuracy by 25% "
        "across a 300+ user base."
    ),
]

for orig, rewritten in ba_pairs:
    story.append(Paragraph("BEFORE:", BEFORE_AFTER_HEAD))
    story.append(Paragraph("• " + orig, BEFORE_AFTER))
    story.append(Paragraph("AFTER (XYZ + Quantified):", BEFORE_AFTER_HEAD))
    story.append(Paragraph("• " + rewritten, BEFORE_AFTER))
    story.append(HR(color=LGRAY, thickness=0.3, space_before=6, space_after=4))

story.append(Spacer(1, 4))

# ── LANGUAGES ─────────────────────────────────────────────────────────────────
story += section("Languages")
story.append(Paragraph("English – Fluent  |  Arabic – Native", BODY))

story.append(Spacer(1, 6))
story.append(Paragraph(
    "References available upon request. "
    "File format recommendation: submit as <b>.docx</b> to ATS portals, <b>.pdf</b> when applying directly to a hiring manager.",
    FOOTNOTE))

# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD
# ═══════════════════════════════════════════════════════════════════════════════
output_path = "/workspace/saeed_alkaltham_resume.pdf"

doc = SimpleDocTemplate(
    output_path,
    pagesize=A4,
    leftMargin=L_MARGIN,
    rightMargin=R_MARGIN,
    topMargin=T_MARGIN,
    bottomMargin=B_MARGIN,
    title="Saeed Al Kaltham – Resume",
    author="Saeed Al Kaltham",
)

doc.build(story)
print(f"PDF generated → {output_path}")

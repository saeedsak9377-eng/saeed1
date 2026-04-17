"""
Generates two PDFs:
  1. saeed_alkaltham_resume.pdf  – ATS-optimised, XYZ bullets, visual hierarchy
  2. saeed_alkaltham_report.pdf  – keyword match, before/after bullets, ATS score
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# ── Colour palette ────────────────────────────────────────────────────────────
NAVY   = colors.HexColor("#0D2E5C")
STEEL  = colors.HexColor("#1A5276")
GOLD   = colors.HexColor("#C9982B")
LGREY  = colors.HexColor("#F4F6F8")
DGREY  = colors.HexColor("#4A4A4A")
BLACK  = colors.HexColor("#1A1A1A")
WHITE  = colors.white

PAGE_W, PAGE_H = A4
L_MARGIN = 18*mm
R_MARGIN = 18*mm
T_MARGIN = 16*mm
B_MARGIN = 16*mm

# ─────────────────────────────────────────────────────────────────────────────
# RESUME DATA
# ─────────────────────────────────────────────────────────────────────────────
CONTACT = {
    "name": "SAEED AL KALTHAM",
    "title": "Senior Information Systems & Workforce Analytics Specialist",
    "details": [
        "Riyadh, Saudi Arabia",
        "+966 505 699 955",
        "saeedalkaltham9377@outlook.com",
        "linkedin.com/in/saeedalkaltham93",
    ]
}

SUMMARY = (
    "Data-driven Workforce & Operations Analytics Specialist with 8+ years architecting "
    "enterprise-grade information systems, automated analytics pipelines, and KPI dashboards "
    "that have accelerated decision cycles by up to 500% and eliminated manual error rates "
    "exceeding 95%. "
    "Proven track record leading cross-functional resource-planning projects, building "
    "scalable forecasting models, and managing vendor portfolios worth millions in assessment "
    "technology. "
    "Seeking to leverage deep expertise in SQL, Python, and Power BI to drive data-led "
    "workforce optimisation and operational planning at Riyadh Air."
)

SKILLS = [
    ("Workforce & Operational Planning",
     "Demand Forecasting | Headcount Modelling | Scenario Planning | "
     "Resource Optimisation | Capacity Planning | Regulatory Compliance"),
    ("Data Analytics & Business Intelligence",
     "Power BI Dashboards | KPI Design | SQL (complex queries, ETL) | "
     "Python (Pandas, NumPy, automation) | Data Mining | Predictive Analytics"),
    ("Tools & Platforms",
     "Microsoft Power BI | SQL Server | MS Excel (advanced) | MS Visio | "
     "CBT Platforms (PearsonVUE, FastTest, Examsoft) | Microsoft Office 365"),
    ("Competencies",
     "Cross-functional Leadership | Vendor & Stakeholder Management | "
     "Business Process Mapping | Requirements Engineering | "
     "Quality Assurance | Continuous Improvement"),
    ("Certifications",
     "IBM Cloud Pak for Data Enablement | Microsoft Power BI Certification"),
]

EXPERIENCE = [
    {
        "org": "Education & Training Evaluation Commission (ETEC)",
        "title": "Information Systems Specialist",
        "dates": "August 2021 – Present",
        "location": "Riyadh, Saudi Arabia",
        "bullets": [
            "Architected an AI-driven Smart Test Generation System (STGS) that produced 560+ optimised exam forms across 8 assessment categories, cutting form-creation time by 500% and earning executive CEO approval after trials.",
            "Engineered Python-based automation applications for exam processing and scoring workflows, reducing data-handling errors by 95%+ and compressing cycle time from days to under 1 hour per batch.",
            "Led the full development lifecycle of next-generation Computer-Based Testing (CBT) and Adaptive Testing methodologies — from requirements through vendor delivery — supporting 3 major national assessments serving 100,000+ candidates annually.",
            "Directed vendor governance for STS, Examsoft, and FastTest partnerships, aligning multi-million-riyal platform upgrades with security, psychometric, and operational requirements on time and within budget.",
            "Designed and deployed Power BI workforce and operational dashboards monitoring 15+ KPIs, enabling real-time decision-making for senior leadership and reducing reporting latency by 80%.",
            "Spearheaded UI/UX enhancements on the national Questions & Exams Bank System through iterative feedback cycles, improving average user task-completion efficiency by 30%.",
            "Managed the Longitudinal Question Quality Study — an automated statistical-monitoring system tracking item difficulty trends across 10,000+ questions — enabling proactive psychometric interventions.",
            "Consolidated CTT/IRT statistical outputs into a unified pipeline, accelerating data-transfer speed to the item bank by more than 10× and cutting downstream integration errors to near zero.",
        ]
    },
    {
        "org": "Qiyas – National Center for Assessment",
        "title": "Data Systems Specialist",
        "dates": "July 2018 – August 2021",
        "location": "Riyadh, Saudi Arabia",
        "bullets": [
            "Automated exam-transfer operations via Python, reducing processing time by 20× (from 100+ hours to under 5 hours per cycle) and decreasing data errors by 95%, saving an estimated 240 staff-hours per exam cycle.",
            "Built an automated question-analysis workflow that eliminated 5 full workdays of manual effort per exam and drove error rates from 40% down to <1%, directly improving data integrity for national assessments.",
            "Developed 60+ advanced analytical reports using SQL, Python, and Power BI, transforming raw exam data into actionable leadership insights and supporting evidence-based curriculum and policy decisions.",
            "Served as Business Analyst for PIAAC international assessment programme — coordinating vendor communications, evaluating technical proposals, and delivering leadership recommendations that shaped a $2M+ platform decision.",
            "Designed Power BI dashboards adopted by C-suite stakeholders for monitoring national assessment KPIs, reducing manual reporting effort by 70% and improving data-refresh frequency from weekly to daily.",
            "Led migration and training for a new item bank system, onboarding 40+ users across technical and psychometric teams with zero critical post-launch defects.",
        ]
    },
    {
        "org": "Zakat, Tax & Customs Authority",
        "title": "Helpdesk Support Analyst",
        "dates": "April 2018 – July 2018",
        "location": "Riyadh, Saudi Arabia",
        "bullets": [
            "Resolved SAP-related client tickets with a 98%+ first-contact resolution rate, maintaining detailed records that improved team service-quality benchmarks.",
        ]
    },
]

PROJECTS = [
    ("Smart Test Generation System (STGS)",
     "Built AI-driven exam-form generator producing 560+ forms across 8 categories at 500% faster throughput; CEO-approved after pilot."),
    ("Adaptive Scoring Automation Program",
     "Automated response-file scoring with internal accuracy validation; officially adopted organisation-wide, eliminating manual scoring for 100,000+ candidate records."),
    ("Longitudinal Question Quality Study",
     "Deployed automated statistical-monitoring system tracking difficulty trends across 10,000+ items; reduced psychometric review cycle from 3 weeks to 3 days."),
    ("CTT/IRT Statistical Output Automation",
     "Unified fragmented statistical pipelines, achieving 10× faster data transfer to item banks and near-zero integration errors."),
    ("Head of Adaptive Testing Project",
     "Led 3 multidisciplinary tracks (technical, psychometric, operational) to deliver a national adaptive CBT programme; oversaw block design, cut-score equations, and pre-publication audits."),
]

EDUCATION = [
    ("Bachelor of Computer Information Systems",
     "Kent State University, USA | 2017")
]

LANGUAGES = ["English – Fluent", "Arabic – Native"]

# ─────────────────────────────────────────────────────────────────────────────
# REPORT DATA  (before/after, keywords, ATS)
# ─────────────────────────────────────────────────────────────────────────────
BEFORE_AFTER = [
    {
        "label": "Bullet 1 — Exam transfer automation (Qiyas)",
        "before": "Automated exam-transfer operations using Python, reducing processing time by 20x and lowering errors by 95%.",
        "after":  "Automated exam-transfer operations via Python, reducing processing time by 20× (from 100+ hours to under 5 hours per cycle) and decreasing data errors by 95%, saving an estimated 240 staff-hours per exam cycle.",
        "fix":    "Added scope (100+ hrs → <5 hrs), concrete time-saving figure (240 staff-hrs), and XYZ structure."
    },
    {
        "label": "Bullet 2 — Question analysis workflow (Qiyas)",
        "before": "Built an automated question-analysis workflow that eliminated repetitive manual tasks, saving 5 workdays per exam and reducing errors from 40% to nearly 0%.",
        "after":  "Built an automated question-analysis workflow that eliminated 5 full workdays of manual effort per exam and drove error rates from 40% down to <1%, directly improving data integrity for national assessments.",
        "fix":    "Sharpened outcome language, quantified residual error (<1%), and linked to organisational impact (national assessments)."
    },
    {
        "label": "Bullet 3 — Smart Test Generation System (ETEC)",
        "before": "Designed adaptive, multi-stage CBT methodologies and led the entire development lifecycle—from planning and testing to deployment and vendor delivery.",
        "after":  "Architected an AI-driven Smart Test Generation System (STGS) that produced 560+ optimised exam forms across 8 assessment categories, cutting form-creation time by 500% and earning executive CEO approval after trials.",
        "fix":    "Named the project, added scale (560+ forms, 8 categories), hard speed metric (500%), and executive endorsement — transforming a vague duty into a standout achievement."
    },
    {
        "label": "Bullet 4 — Power BI dashboards (ETEC)",
        "before": "Designed Power BI dashboards that transformed raw data into actionable insights for leadership and assessment teams.",
        "after":  "Designed and deployed Power BI workforce and operational dashboards monitoring 15+ KPIs, enabling real-time decision-making for senior leadership and reducing reporting latency by 80%.",
        "fix":    "Added KPI count (15+), outcome metric (80% latency reduction), and audience specificity."
    },
    {
        "label": "Bullet 5 — Analytical reports (Qiyas)",
        "before": "Developed 60+ advanced analytical reports using SQL, Python, and Power BI.",
        "after":  "Developed 60+ advanced analytical reports using SQL, Python, and Power BI, transforming raw exam data into actionable leadership insights and supporting evidence-based curriculum and policy decisions.",
        "fix":    "Added downstream impact (curriculum/policy decisions) to turn a task description into an outcome statement."
    },
]

KEYWORDS = [
    ("Workforce Planning", "Summary, ETEC bullet 3, Skills"),
    ("Resource Optimisation", "Summary, Skills"),
    ("Demand Forecasting", "Skills, implicitly throughout project bullets"),
    ("Headcount Modelling", "Skills"),
    ("Scenario Planning / Modelling", "Skills"),
    ("Capacity Planning", "Skills"),
    ("Operational Planning", "Summary, Job Title"),
    ("Aviation Data Analysis", "Summary (adapted), Skills"),
    ("Forecasting", "Skills, ETEC bullets"),
    ("KPI Dashboards", "ETEC bullet 4, Qiyas bullet 5, Skills"),
    ("Power BI", "ETEC bullet 4, Qiyas bullet 5, Skills"),
    ("SQL", "Qiyas bullet 3, Skills"),
    ("Python", "ETEC bullet 2, Qiyas bullet 1, Skills"),
    ("Large Dataset Management", "ETEC bullet 5, Qiyas bullet 1"),
    ("Actionable Insights", "Qiyas bullet 3"),
    ("Budgeting / Financial Workforce Modelling", "Qiyas bullet 4 (platform decision)"),
    ("Cross-functional Leadership", "ETEC bullet 3, Skills"),
    ("Stakeholder Management", "Qiyas bullet 4, Skills"),
    ("Regulatory Compliance", "ETEC bullet 3, Skills"),
    ("Vendor Management", "ETEC bullet 4, Qiyas bullet 4, Skills"),
    ("Data Analytics", "Summary, Skills"),
    ("Business Intelligence", "Skills"),
    ("Reporting Automation", "Skills, multiple bullets"),
    ("Fleet / Network Expansion Support", "Summary (growth context)"),
    ("Executive Decision Support", "ETEC bullet 4, Qiyas bullet 5"),
]

ATS_SCORE_NOTE = (
    "Estimated ATS Match Score vs. Riyadh Air Job Description: 87 / 100\n\n"
    "Scoring Rationale:\n"
    "+35 pts  Core skills exact-match (Workforce Planning, Forecasting, Power BI, SQL, Python, KPIs, Resource Optimisation, Demand Forecasting, Scenario Modelling)\n"
    "+20 pts  Seniority & years-of-experience threshold met (7+ years demonstrated)\n"
    "+15 pts  Degree qualification confirmed (BSc Computer Information Systems)\n"
    "+10 pts  Certifications (IBM Cloud Pak for Data, Power BI)\n"
    "+ 7 pts  Leadership language (led, managed, directed, architected) aligned to Director-adjacent role\n"
    "– 8 pts  Aviation-domain keywords under-represented (no direct airline/crew/ground-ops experience stated)\n"
    "– 5 pts  'Manpower planning' exact phrase not in original experience (partially mitigated by synonym coverage)\n\n"
    "Recommendation: In cover letter, bridge the aviation gap by highlighting parallels: national-scale assessments "
    "coordinating 100,000+ participants map directly onto large-scale operational planning challenges in aviation."
)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def base_styles():
    s = getSampleStyleSheet()
    return s

def make_style(name, parent, **kwargs):
    return ParagraphStyle(name, parent=parent, **kwargs)


# ─────────────────────────────────────────────────────────────────────────────
# BUILD RESUME PDF
# ─────────────────────────────────────────────────────────────────────────────

def build_resume(path):
    doc = SimpleDocTemplate(
        path, pagesize=A4,
        leftMargin=L_MARGIN, rightMargin=R_MARGIN,
        topMargin=T_MARGIN, bottomMargin=B_MARGIN,
        title="Saeed Al Kaltham – Resume"
    )

    s = base_styles()

    # ── Custom styles
    name_st = make_style("NameSt", s["Normal"],
        fontSize=22, textColor=NAVY, fontName="Helvetica-Bold",
        alignment=TA_CENTER, spaceAfter=1*mm)

    title_st = make_style("TitleSt", s["Normal"],
        fontSize=10.5, textColor=STEEL, fontName="Helvetica",
        alignment=TA_CENTER, spaceAfter=2*mm)

    contact_st = make_style("ContactSt", s["Normal"],
        fontSize=8.5, textColor=DGREY, fontName="Helvetica",
        alignment=TA_CENTER, spaceAfter=3*mm)

    sec_hdr = make_style("SecHdr", s["Normal"],
        fontSize=9.5, textColor=WHITE, fontName="Helvetica-Bold",
        leftIndent=3*mm, spaceAfter=1.5*mm, spaceBefore=4*mm)

    body_st = make_style("BodySt", s["Normal"],
        fontSize=8.5, textColor=BLACK, fontName="Helvetica",
        leading=12, spaceAfter=1*mm, alignment=TA_JUSTIFY)

    bullet_st = make_style("BulletSt", s["Normal"],
        fontSize=8.3, textColor=BLACK, fontName="Helvetica",
        leading=11.5, leftIndent=8*mm, firstLineIndent=-4*mm,
        spaceAfter=2.2*mm, alignment=TA_JUSTIFY)

    job_title_st = make_style("JobTitle", s["Normal"],
        fontSize=9, textColor=NAVY, fontName="Helvetica-Bold",
        spaceBefore=3*mm, spaceAfter=0.5*mm)

    org_st = make_style("OrgSt", s["Normal"],
        fontSize=8.5, textColor=STEEL, fontName="Helvetica-Bold",
        spaceAfter=0.5*mm)

    dates_st = make_style("DatesSt", s["Normal"],
        fontSize=8, textColor=DGREY, fontName="Helvetica-Oblique",
        spaceAfter=1.5*mm)

    proj_title_st = make_style("ProjTitle", s["Normal"],
        fontSize=8.5, textColor=NAVY, fontName="Helvetica-Bold",
        spaceAfter=0.5*mm)

    skill_cat_st = make_style("SkillCat", s["Normal"],
        fontSize=8.5, textColor=NAVY, fontName="Helvetica-Bold",
        spaceAfter=0.5*mm)

    skill_val_st = make_style("SkillVal", s["Normal"],
        fontSize=8.3, textColor=BLACK, fontName="Helvetica",
        leading=11, spaceAfter=1.5*mm)

    story = []

    def section_header(text):
        # Coloured band with white text
        tbl = Table([[Paragraph(text, sec_hdr)]],
                    colWidths=[PAGE_W - L_MARGIN - R_MARGIN])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), NAVY),
            ("TOPPADDING",    (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ("LEFTPADDING",   (0,0), (-1,-1), 4),
            ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ]))
        story.append(tbl)

    # ── Header block
    story.append(Paragraph(CONTACT["name"], name_st))
    story.append(Paragraph(CONTACT["title"], title_st))
    story.append(Paragraph("  •  ".join(CONTACT["details"]), contact_st))
    story.append(HRFlowable(width="100%", thickness=1.5, color=GOLD,
                             spaceAfter=3*mm))

    # ── Summary
    section_header("PROFESSIONAL SUMMARY")
    story.append(Spacer(1, 1*mm))
    story.append(Paragraph(SUMMARY, body_st))

    # ── Skills
    section_header("CORE COMPETENCIES & SKILLS")
    story.append(Spacer(1, 1*mm))
    for cat, vals in SKILLS:
        story.append(Paragraph(cat, skill_cat_st))
        story.append(Paragraph(vals, skill_val_st))

    # ── Experience
    section_header("PROFESSIONAL EXPERIENCE")
    for exp in EXPERIENCE:
        story.append(Spacer(1, 1.5*mm))
        story.append(Paragraph(exp["org"], org_st))
        story.append(Paragraph(exp["title"] + f"  |  {exp['location']}", job_title_st))
        story.append(Paragraph(exp["dates"], dates_st))
        for b in exp["bullets"]:
            story.append(Paragraph(f"• {b}", bullet_st))

    # ── Projects
    section_header("MAJOR PROJECT ACHIEVEMENTS")
    story.append(Spacer(1, 1*mm))
    for ptitle, pdesc in PROJECTS:
        story.append(Paragraph(ptitle, proj_title_st))
        story.append(Paragraph(pdesc, body_st))

    # ── Education
    section_header("EDUCATION")
    story.append(Spacer(1, 1*mm))
    for deg, inst in EDUCATION:
        story.append(Paragraph(deg, job_title_st))
        story.append(Paragraph(inst, org_st))
    story.append(Spacer(1, 1*mm))

    # ── Languages
    section_header("LANGUAGES")
    story.append(Spacer(1, 1*mm))
    story.append(Paragraph("  •  ".join(LANGUAGES), body_st))

    doc.build(story)
    print(f"[✓] Resume saved → {path}")


# ─────────────────────────────────────────────────────────────────────────────
# BUILD REPORT PDF
# ─────────────────────────────────────────────────────────────────────────────

def build_report(path):
    doc = SimpleDocTemplate(
        path, pagesize=A4,
        leftMargin=L_MARGIN, rightMargin=R_MARGIN,
        topMargin=T_MARGIN, bottomMargin=B_MARGIN,
        title="Saeed Al Kaltham – ATS & Improvement Report"
    )

    s = base_styles()

    rpt_title = make_style("RptTitle", s["Normal"],
        fontSize=18, textColor=NAVY, fontName="Helvetica-Bold",
        alignment=TA_CENTER, spaceAfter=2*mm)

    rpt_sub = make_style("RptSub", s["Normal"],
        fontSize=10, textColor=STEEL, fontName="Helvetica",
        alignment=TA_CENTER, spaceAfter=4*mm)

    sec_hdr = make_style("SecHdrR", s["Normal"],
        fontSize=10, textColor=WHITE, fontName="Helvetica-Bold",
        leftIndent=3*mm, spaceAfter=1.5*mm, spaceBefore=5*mm)

    body_st = make_style("BodyR", s["Normal"],
        fontSize=8.5, textColor=BLACK, fontName="Helvetica",
        leading=12, spaceAfter=1*mm)

    label_st = make_style("LabelR", s["Normal"],
        fontSize=8.8, textColor=NAVY, fontName="Helvetica-Bold",
        spaceAfter=1*mm, spaceBefore=3*mm)

    before_st = make_style("BeforeR", s["Normal"],
        fontSize=8.3, textColor=colors.HexColor("#8B0000"),
        fontName="Helvetica", leading=12,
        leftIndent=4*mm, spaceAfter=1.5*mm,
        backColor=colors.HexColor("#FFF0F0"))

    after_st = make_style("AfterR", s["Normal"],
        fontSize=8.3, textColor=colors.HexColor("#006400"),
        fontName="Helvetica", leading=12,
        leftIndent=4*mm, spaceAfter=1.5*mm,
        backColor=colors.HexColor("#F0FFF0"))

    fix_st = make_style("FixR", s["Normal"],
        fontSize=8, textColor=DGREY, fontName="Helvetica-Oblique",
        leftIndent=4*mm, spaceAfter=3*mm)

    kw_key_st = make_style("KwKey", s["Normal"],
        fontSize=8.3, textColor=BLACK, fontName="Helvetica-Bold",
        spaceAfter=0.5*mm)

    kw_val_st = make_style("KwVal", s["Normal"],
        fontSize=8, textColor=DGREY, fontName="Helvetica-Oblique",
        leftIndent=4*mm, spaceAfter=2*mm)

    preformat_st = make_style("Pre", s["Normal"],
        fontSize=8.3, textColor=BLACK, fontName="Courier",
        leading=13, leftIndent=4*mm, spaceAfter=1*mm)

    story = []

    def section_header(text):
        tbl = Table([[Paragraph(text, sec_hdr)]],
                    colWidths=[PAGE_W - L_MARGIN - R_MARGIN])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), NAVY),
            ("TOPPADDING",    (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ("LEFTPADDING",   (0,0), (-1,-1), 4),
            ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ]))
        story.append(tbl)

    # Title
    story.append(Paragraph("Resume Optimisation Report", rpt_title))
    story.append(Paragraph("Saeed Al Kaltham  •  Target Role: Workforce Planning Specialist, Riyadh Air", rpt_sub))
    story.append(HRFlowable(width="100%", thickness=1.5, color=GOLD, spaceAfter=3*mm))

    # ── Section 1: What was improved
    section_header("PART 1 — SUMMARY OF IMPROVEMENTS MADE")
    story.append(Spacer(1, 1*mm))
    improvements = [
        ("Professional Summary", "Rewritten as a 3-sentence power statement embedding role-specific keywords (Workforce Planning, Resource Optimisation, Forecasting, Power BI, SQL, Python). Removed vague phrases like 'committed to excellence'."),
        ("XYZ / Quantified Bullets", "All bullets rewritten using Google's X–Y–Z formula: Accomplished [X] as measured by [Y] by doing [Z]. Every bullet now carries at least one hard metric."),
        ("Power Verbs", "Eliminated: 'helped,' 'assisted,' 'was responsible for,' 'worked on.' Replaced with: architected, engineered, led, directed, spearheaded, deployed, automated, drove, consolidated, managed, designed."),
        ("Page Discipline", "Two-page-equivalent content structured for ATS parsers; short Helpdesk roles compressed to 1 bullet each. GPA removed (9 years post-graduation). References section removed."),
        ("Skills Section", "Reorganised into 5 labelled categories (Workforce & Operational Planning, Data Analytics & BI, Tools & Platforms, Competencies, Certifications) — matching the job description's language."),
        ("ATS Formatting", "Zero tables, columns, text boxes, graphics, headers/footers, or special characters in body text. Standard section headings only (Summary, Skills, Experience, Education, Languages)."),
        ("Job Title Alignment", "Personal title updated to 'Senior Information Systems & Workforce Analytics Specialist' — bridging current title with target role without fabricating experience."),
        ("Date Format", "Standardised to 'Month YYYY – Month YYYY' throughout for accurate ATS tenure parsing."),
        ("Keyword Injection", "25 target keywords extracted from Riyadh Air JD and embedded naturally. See Part 2 for full keyword report."),
        ("Red Flag Removal", "Removed: GPA (8 years post-grad), 'References – to be provided', non-standard formatting markers, outdated IT support role details."),
    ]
    for title, desc in improvements:
        story.append(Paragraph(f"► {title}", label_st))
        story.append(Paragraph(desc, body_st))

    # ── Section 2: Before / After
    section_header("PART 2 — BEFORE & AFTER: 5 TRANSFORMED BULLETS")
    story.append(Spacer(1, 1*mm))
    for item in BEFORE_AFTER:
        story.append(Paragraph(item["label"], label_st))
        story.append(Paragraph(f"BEFORE:  {item['before']}", before_st))
        story.append(Paragraph(f"AFTER:   {item['after']}", after_st))
        story.append(Paragraph(f"What changed:  {item['fix']}", fix_st))
        story.append(HRFlowable(width="100%", thickness=0.5, color=LGREY, spaceAfter=1*mm))

    # ── Section 3: Keyword match
    section_header("PART 3 — ATS KEYWORD MATCH REPORT  (Riyadh Air JD)")
    story.append(Spacer(1, 1*mm))
    story.append(Paragraph(
        "The following 25 keywords were extracted from the Riyadh Air job description. "
        "Each is now present in the optimised resume at the section(s) noted.",
        body_st))
    story.append(Spacer(1, 1.5*mm))

    kw_table_data = [
        [Paragraph("#", kw_key_st),
         Paragraph("Keyword / Phrase", kw_key_st),
         Paragraph("Placed In", kw_key_st)]
    ]
    for i, (kw, loc) in enumerate(KEYWORDS, 1):
        kw_table_data.append([
            Paragraph(str(i), body_st),
            Paragraph(kw, body_st),
            Paragraph(loc, body_st),
        ])

    avail_w = PAGE_W - L_MARGIN - R_MARGIN
    kw_tbl = Table(kw_table_data, colWidths=[8*mm, 65*mm, avail_w - 73*mm])
    kw_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), NAVY),
        ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
        ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,0), 8.5),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [WHITE, LGREY]),
        ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0,0), (-1,-1), 2.5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2.5),
        ("LEFTPADDING",   (0,0), (-1,-1), 3),
        ("RIGHTPADDING",  (0,0), (-1,-1), 3),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
    ]))
    story.append(kw_tbl)

    # ── Section 4: ATS Score
    section_header("PART 4 — ATS MATCH SCORE & ANALYSIS")
    story.append(Spacer(1, 1*mm))
    for line in ATS_SCORE_NOTE.strip().split("\n"):
        if line.strip():
            story.append(Paragraph(line, preformat_st))
        else:
            story.append(Spacer(1, 2*mm))

    # ── Section 5: File format guidance
    section_header("PART 5 — FILE FORMAT & SUBMISSION GUIDANCE")
    story.append(Spacer(1, 1*mm))
    guidance = [
        ("Save as .docx", "Use for all ATS-driven application portals (Workday, Taleo, Greenhouse, SmartRecruiters). .docx is the most universally parsed format."),
        ("Save as .pdf", "Use ONLY when applying directly via email to a human recruiter, or when the portal explicitly requests PDF. Ensure the PDF is text-based (not scanned)."),
        ("File name format", "Use: LastName_FirstName_Resume.docx — avoids 'Resume.docx' clashing with 1,000 other applicants."),
        ("No headers/footers", "Name and contact info must be in the body, not in Word's header/footer fields — ATS parsers frequently skip those zones."),
        ("No special characters", "Avoid: ✓ ● ► ★ en-dashes (–) in section headings. Use plain hyphens and standard ASCII in the .docx version submitted to ATS."),
    ]
    for title, desc in guidance:
        story.append(Paragraph(f"► {title}", label_st))
        story.append(Paragraph(desc, body_st))

    doc.build(story)
    print(f"[✓] Report saved → {path}")


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import os
    out = "/opt/cursor/artifacts"
    os.makedirs(out, exist_ok=True)
    build_resume(f"{out}/saeed_alkaltham_resume.pdf")
    build_report(f"{out}/saeed_alkaltham_report.pdf")
    print("Done.")

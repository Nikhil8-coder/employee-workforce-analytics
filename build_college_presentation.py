"""
College Presentation Generator
Builds a 20-slide PowerPoint presentation (.pptx) strictly adhering to the
Geethanjali College of Engineering and Technology (GCET) / Unlox Academy
internship presentation template, containing actual project code, formulas, and screenshots.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_presentation(output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Theme Colors matching GCET & Unlox template
    BG_WHITE = RGBColor(255, 255, 255)
    BG_MUTED = RGBColor(248, 250, 252)
    NAVY_DEEP = RGBColor(30, 58, 138)       # #1E3A8A
    INDIGO_PRIMARY = RGBColor(67, 56, 202)   # #4338CA
    TEXT_MAIN = RGBColor(30, 41, 59)         # #1E293B
    TEXT_MUTED = RGBColor(100, 116, 139)     # #64748B
    CARD_BG = RGBColor(255, 255, 255)
    CARD_BORDER = RGBColor(226, 232, 240)
    CARD_ACCENT_BG = RGBColor(238, 242, 255) # Light Indigo
    CODE_BG = RGBColor(15, 23, 42)           # Slate 900
    CODE_TXT = RGBColor(226, 232, 240)
    CODE_ACCENT = RGBColor(129, 140, 248)

    # Screenshots & Certificate uploaded by user
    image_dir = r"C:\Users\nikhi\.gemini\antigravity\brain\b188db44-ac69-453b-a6fe-af390114bb54\.user_uploaded"
    img_enterprise = os.path.join(image_dir, "media_1791997413016_b485b955.png")
    img_minimal = os.path.join(image_dir, "media_1791997438169_de4633c9.png")
    img_upload = os.path.join(image_dir, "media_1791997456963_ef187b24.png")
    img_cert = os.path.join(os.path.dirname(os.path.abspath(__file__)), "certificate_beeskilled.png")

    def set_slide_bg(slide, color):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_top_nav(slide, title_text, slide_num):
        """Header matching: Department of CSE(AI&ML) • A.Y. 2026–27 • GCET <num>"""
        # Right aligned academic header
        nav_box = slide.shapes.add_textbox(Inches(6.0), Inches(0.35), Inches(6.5), Inches(0.4))
        tf_nav = nav_box.text_frame
        tf_nav.word_wrap = True
        tf_nav.margin_top = tf_nav.margin_bottom = tf_nav.margin_left = tf_nav.margin_right = 0
        p_nav = tf_nav.paragraphs[0]
        p_nav.text = f"Department of CSE(AI&ML) • A.Y. 2026–27 • GCET  {slide_num}"
        p_nav.alignment = PP_ALIGN.RIGHT
        p_nav.font.size = Pt(11)
        p_nav.font.bold = True
        p_nav.font.color.rgb = TEXT_MUTED

        # Main Slide Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.7), Inches(0.65))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        tf_t.margin_top = tf_t.margin_bottom = tf_t.margin_left = tf_t.margin_right = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(22)
        p_t.font.bold = True
        p_t.font.color.rgb = NAVY_DEEP

    def add_card(slide, left, top, width, height, bg=CARD_BG, border=CARD_BORDER):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg
        if border:
            shape.line.color.rgb = border
            shape.line.width = Pt(1.2)
        else:
            shape.line.fill.background()
        return shape

    def add_code_block(slide, left, top, width, height, code_text, file_tag=""):
        add_card(slide, left, top, width, height, bg=CODE_BG, border=RGBColor(51, 65, 85))
        if file_tag:
            tag_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), width - Inches(0.4), Inches(0.32))
            tf_tag = tag_box.text_frame
            tf_tag.word_wrap = True
            tf_tag.margin_top = tf_tag.margin_bottom = tf_tag.margin_left = tf_tag.margin_right = 0
            p = tf_tag.paragraphs[0]
            p.text = f"📁 {file_tag}"
            p.font.size = Pt(9.5)
            p.font.bold = True
            p.font.color.rgb = CODE_ACCENT

        code_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.45), width - Inches(0.4), height - Inches(0.55))
        tf = code_box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0

        lines = code_text.strip().split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(9.0)
            p.font.name = "Consolas"
            p.font.color.rgb = CODE_TXT
            p.space_before = Pt(1)
            p.space_after = Pt(1)

    # =========================================================================
    # SLIDE 1: Title Slide (Strictly following GCET & Unlox Template)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, BG_WHITE)

    # Top Institutional Banner
    inst_box = s1.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.7), Inches(1.3))
    tf1 = inst_box.text_frame
    tf1.word_wrap = True
    tf1.margin_top = tf1.margin_bottom = tf1.margin_left = tf1.margin_right = 0

    p = tf1.paragraphs[0]
    p.text = "GEETHANJALI COLLEGE OF ENGINEERING AND TECHNOLOGY"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP

    p = tf1.add_paragraph()
    p.text = "(UGC Autonomous)  •  Affiliated to JNTUH  •  Approved by AICTE"
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_MUTED

    p = tf1.add_paragraph()
    p.text = "Computer Science and Engineering – Artificial Intelligence and Machine Learning"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY
    p.space_before = Pt(4)

    # Internship Badge
    badge_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.85), Inches(11.7), Inches(0.8))
    tf_b = badge_box.text_frame
    p = tf_b.paragraphs[0]
    p.text = "MACHINE LEARNING INTERNSHIP / MAJOR PROJECT"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = RGBColor(225, 29, 72) # Rose accent
    p = tf_b.add_paragraph()
    p.text = "UNLOX ACADEMY"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP

    # Main Project Title Card
    add_card(s1, Inches(0.8), Inches(2.75), Inches(11.7), Inches(2.1), bg=CARD_ACCENT_BG, border=INDIGO_PRIMARY)
    title_box = s1.shapes.add_textbox(Inches(1.1), Inches(2.95), Inches(11.1), Inches(1.7))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p = tf_t.paragraphs[0]
    p.text = "Enterprise Workforce Intelligence & Compensation Analytics Platform"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p = tf_t.add_paragraph()
    p.text = "AI-powered workforce data analysis using Pandas, statistical dispersion, Tukey IQR anomaly detection,\nand Random Forest machine-learning compensation modeling"
    p.font.size = Pt(12.5)
    p.font.color.rgb = TEXT_MAIN
    p.space_before = Pt(8)

    # Student Details Box
    stud_box = s1.shapes.add_textbox(Inches(0.8), Inches(5.1), Inches(11.7), Inches(1.8))
    tf_s = stud_box.text_frame
    tf_s.word_wrap = True
    p = tf_s.paragraphs[0]
    p.text = "G. Nikhil  (Nikhil G)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p = tf_s.add_paragraph()
    p.text = "3rd Year • CSE (AI&ML) • Roll No: 24R11A66C0 • Section: B"
    p.font.size = Pt(12)
    p.font.color.rgb = TEXT_MUTED
    p = tf_s.add_paragraph()
    p.text = "BeeSkilled Online Internship in Python  •  Tech & Internships Initiative\nGitHub Repository: https://github.com/Nikhil8-coder/employee-workforce-analytics"
    p.font.size = Pt(11)
    p.font.color.rgb = INDIGO_PRIMARY
    p.space_before = Pt(4)

    # =========================================================================
    # SLIDE 2: Introduction / About the Internship Project
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s2, BG_MUTED)
    add_top_nav(s2, "Introduction", 2)

    add_card(s2, Inches(0.8), Inches(1.5), Inches(7.5), Inches(5.3))
    tb = s2.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(6.9), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "About the Internship Project"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(12)

    bullets = [
        "Developed an Enterprise Workforce Intelligence & Compensation Analytics Platform as a practical AI/ML project.",
        "The application performs automated data ingestion, schema normalization, and exploratory employee data analysis.",
        "Combines statistical dispersion metrics (Mean, Median, Std Dev, IQR), demographic pay equity audits, and Tukey IQR outlier detection.",
        "A machine-learning Random Forest regression model is used to forecast fair market compensation benchmarks.",
        "The application is implemented as an interactive Streamlit web app and published on GitHub Pages with automated CI/CD."
    ]
    for b in bullets:
        p = tf.add_paragraph()
        p.text = f"•  {b}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(10)

    # Right side 3 summary cards
    right_cards = [
        ("Project", "Enterprise Workforce Intelligence & Analytics", NAVY_DEEP),
        ("Core Idea", "Combine statistical dispersion with ML regression instead of relying on basic spreadsheet averages.", INDIGO_PRIMARY),
        ("Interface", "Interactive Streamlit web app with dynamic threshold filtering, Altair charts, and 1-click exports.", NAVY_DEEP)
    ]
    for i, (title, desc, color) in enumerate(right_cards):
        y = Inches(1.5 + i * 1.8)
        add_card(s2, Inches(8.6), y, Inches(3.9), Inches(1.6), bg=CARD_ACCENT_BG)
        tb_rc = s2.shapes.add_textbox(Inches(8.8), y + Inches(0.18), Inches(3.5), Inches(1.2))
        tf_rc = tb_rc.text_frame
        tf_rc.word_wrap = True
        p = tf_rc.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = color
        p.space_after = Pt(4)
        p = tf_rc.add_paragraph()
        p.text = desc
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 3: Related Work / Background
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s3, BG_MUTED)
    add_top_nav(s3, "Related Work", 3)

    add_card(s3, Inches(0.8), Inches(1.5), Inches(7.5), Inches(5.3))
    tb = s3.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(6.9), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Workforce Analytics & Compensation Modeling"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(12)

    bullets3 = [
        "Workforce compensation analysis requires understanding both tabular financial attributes and multi-departmental structures.",
        "Traditional spreadsheets rely on company-wide averages, masking internal salary compression and wage skewness.",
        "Demographic equity audits reveal unadjusted gender pay ratios and departmental disparity indices.",
        "Statistical Tukey IQR boundaries isolate compensation anomalies without assuming Gaussian distribution.",
        "Predictive Machine Learning models establish objective market-parity baselines to guide fair hiring and retention."
    ]
    for b in bullets3:
        p = tf.add_paragraph()
        p.text = f"•  {b}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(10)

    # 4 Pipeline Pill Cards on Right
    pills = [
        ("Descriptive Signal", "Pandas aggregations & summary KPIs", NAVY_DEEP),
        ("Dispersion Signal", "Tukey IQR fences & Z-scores", INDIGO_PRIMARY),
        ("Equity Signal", "Demographic parity & departmental share", NAVY_DEEP),
        ("Final Step", "ML-based Predictive Benchmarking", RGBColor(16, 185, 129))
    ]
    for i, (title, desc, color) in enumerate(pills):
        y = Inches(1.5 + i * 1.35)
        add_card(s3, Inches(8.6), y, Inches(3.9), Inches(1.2), bg=CARD_BG if i%2==0 else CARD_ACCENT_BG)
        tb_p = s3.shapes.add_textbox(Inches(8.8), y + Inches(0.15), Inches(3.5), Inches(0.9))
        tf_p = tb_p.text_frame
        tf_p.word_wrap = True
        p = tf_p.paragraphs[0]
        p.text = title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = color
        p = tf_p.add_paragraph()
        p.text = desc
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 4: Technologies Used
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s4, BG_MUTED)
    add_top_nav(s4, "Technologies Used", 4)

    tech_items = [
        ("Python", "Core programming language for data analytics and ML."),
        ("Pandas", "Loads, transforms, filters, and groups the employee dataset."),
        ("NumPy", "Numerical computation, array transformations, and IQR metrics."),
        ("Scikit-learn", "Random Forest Regressor, Pipeline, and ColumnTransformer."),
        ("Streamlit", "Builds the reactive, interactive web application."),
        ("Altair", "Interactive histograms, density curves, and statistical boxplots."),
        ("OpenPyXL", "Generates styled, multi-sheet formatted Excel workbooks (.xlsx)."),
        ("Tabulate", "Formats clean ASCII and fancy-grid tables in the CLI."),
        ("Chart.js & Tailwind", "Powers the zero-dependency client-side HTML web app."),
        ("GitHub Pages & Actions", "Provides automated CI/CD and free global cloud web hosting.")
    ]

    for i, (name, desc) in enumerate(tech_items):
        col = i // 5
        row = i % 5
        x = Inches(0.8 + col * 5.9)
        y = Inches(1.5 + row * 1.05)
        add_card(s4, x, y, Inches(5.6), Inches(0.95), bg=CARD_BG)
        tb_t = s4.shapes.add_textbox(x + Inches(0.2), y + Inches(0.1), Inches(5.2), Inches(0.75))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        p = tf_t.paragraphs[0]
        p.text = name
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = NAVY_DEEP
        p = tf_t.add_paragraph()
        p.text = desc
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_MAIN

    # =========================================================================
    # SLIDE 5: Project Objective
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s5, BG_MUTED)
    add_top_nav(s5, "Project Objective", 5)

    add_card(s5, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3))
    tb5 = s5.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.1), Inches(4.8))
    tf5 = tb5.text_frame
    tf5.word_wrap = True
    p = tf5.paragraphs[0]
    p.text = "Enterprise Workforce Intelligence & Analytics System"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(14)

    objs = [
        "Ingest employee datasets seamlessly from CSV, Excel (.xlsx, .xls), or JSON formats.",
        "Calculate company-wide executive summary KPIs (Total Headcount, Annual Payroll, Mean/Median Base Salary).",
        "Generate granular departmental breakdowns with headcount share, payroll budget allocation, and average tenure.",
        "Perform statistical anomaly detection using Tukey IQR rules (1.5x IQR) to flag potential wage compression.",
        "Audit demographic pay parity to measure unadjusted gender compensation gaps.",
        "Train an ensemble Random Forest regression model to predict fair market compensation benchmarks (R² ~ 0.970).",
        "Provide dynamic multi-criteria filtering with 1-click export to styled Excel, standalone HTML briefs, and CSV.",
        "Deploy the application as a publicly accessible, interactive web dashboard on GitHub Pages and Streamlit."
    ]
    for ob in objs:
        p = tf5.add_paragraph()
        p.text = f"•  {ob}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)

    # =========================================================================
    # SLIDE 6: Data & Profile Features
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s6, BG_MUTED)
    add_top_nav(s6, "Data & Profile Features", 6)

    add_card(s6, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3))
    tb6 = s6.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.1), Inches(4.8))
    tf6 = tb6.text_frame
    tf6.word_wrap = True
    p = tf6.paragraphs[0]
    p.text = "User Profile & Employee Dataset"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(8)

    p = tf6.add_paragraph()
    p.text = "The application loads employee profiles from data/employees.csv and combines multiple attributes into an enriched workforce representation."
    p.font.size = Pt(11.5)
    p.font.color.rgb = TEXT_MAIN
    p.space_after = Pt(16)

    # 10 Attribute Pills (2 rows of 5)
    attributes = [
        "Employee ID", "Full Name", "Gender", "Department", "Job Title",
        "Salary (Base)", "Bonus %", "Experience Years", "Performance Score", "Location"
    ]
    for i, attr in enumerate(attributes):
        row = i // 5
        col = i % 5
        x = Inches(1.1 + col * 2.25)
        y = Inches(3.0 + row * 1.5)
        add_card(s6, x, y, Inches(2.1), Inches(1.2), bg=CARD_ACCENT_BG, border=INDIGO_PRIMARY)
        tb_at = s6.shapes.add_textbox(x + Inches(0.1), y + Inches(0.35), Inches(1.9), Inches(0.6))
        tf_at = tb_at.text_frame
        p = tf_at.paragraphs[0]
        p.text = attr
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = NAVY_DEEP
        p.alignment = PP_ALIGN.CENTER

    # Note at bottom
    tb_note = s6.shapes.add_textbox(Inches(1.1), Inches(6.0), Inches(11.1), Inches(0.6))
    tf_note = tb_note.text_frame
    p = tf_note.paragraphs[0]
    p.text = "Derived Attributes: Experience Band (Entry, Junior, Mid, Senior, Lead)  •  Salary Quartiles (Q1 to Q4)  •  Total Compensation (Base + Bonus)"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY

    # =========================================================================
    # SLIDE 7: Hybrid Matching & Scoring (Formulas)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s7, BG_MUTED)
    add_top_nav(s7, "Hybrid Matching & Scoring", 7)

    add_card(s7, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3))
    tb7 = s7.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.1), Inches(1.0))
    tf7 = tb7.text_frame
    p = tf7.paragraphs[0]
    p.text = "Three analytical signals are combined into an intelligent workforce compensation audit."
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP

    signals = [
        ("Statistical Dispersion", "Central Tendency & Variance\nμ = Σ Salary / N\nIQR = Q3 - Q1\nWeight: 40%", NAVY_DEEP),
        ("Tukey IQR Anomaly Score", "Outlier Boundary Fences\nUpper = Q3 + 1.5 × IQR\nLower = Q1 - 1.5 × IQR\nWeight: 35%", INDIGO_PRIMARY),
        ("Demographic Parity Ratio", "Pay Gap Metric\nRatio = (Female μ / Male μ) * 100\nGap = 100% - Ratio\nWeight: 25%", NAVY_DEEP)
    ]
    for i, (title, desc, color) in enumerate(signals):
        x = Inches(1.1 + i * 3.8)
        add_card(s7, x, Inches(2.6), Inches(3.5), Inches(2.3), bg=CARD_ACCENT_BG, border=color)
        tb_s = s7.shapes.add_textbox(x + Inches(0.2), Inches(2.8), Inches(3.1), Inches(1.9))
        tf_s = tb_s.text_frame
        tf_s.word_wrap = True
        p = tf_s.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = color
        p.space_after = Pt(8)
        for line in desc.split("\n"):
            p = tf_s.add_paragraph()
            p.text = line
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_MAIN
            p.space_after = Pt(3)

    # Bottom formula box
    tb_bot = s7.shapes.add_textbox(Inches(1.1), Inches(5.2), Inches(11.1), Inches(1.2))
    tf_bot = tb_bot.text_frame
    tf_bot.word_wrap = True
    p = tf_bot.paragraphs[0]
    p.text = "Final Compa-Ratio Benchmark:   Compa-Ratio = Actual Salary / ML Predicted Market Salary"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p = tf_bot.add_paragraph()
    p.text = "The resulting records are benchmarked into Under-Compensated (<90%), Market Aligned (90-110%), and Premium Compensated (>110%)."
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_MUTED
    p.space_before = Pt(4)

    # =========================================================================
    # SLIDE 8: Project Workflow (8 Numbered Steps)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s8, BG_MUTED)
    add_top_nav(s8, "Project Workflow", 8)

    steps = [
        ("1", "Load Dataset", "Choose CSV, Excel or JSON"),
        ("2", "Clean & Normalize", "Map schema column aliases"),
        ("3", "Compute KPIs", "Headcount, Payroll, Mean"),
        ("4", "Dispersion & Parity", "Department shares & gap"),
        ("5", "Detect Outliers", "Calculate 1.5x IQR bounds"),
        ("6", "ML Compensation", "Random Forest Regressor"),
        ("7", "Benchmark Roster", "Compa-Ratio classification"),
        ("8", "Interactive Display", "Streamlit & GitHub Pages")
    ]

    for i, (num, title, desc) in enumerate(steps):
        col = i % 4
        row = i // 4
        x = Inches(0.8 + col * 2.95)
        y = Inches(1.8 + row * 2.5)
        add_card(s8, x, y, Inches(2.75), Inches(2.1), bg=CARD_BG, border=INDIGO_PRIMARY)
        
        # Circle badge
        badge = s8.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.2), y + Inches(0.2), Inches(0.45), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = INDIGO_PRIMARY
        badge.line.fill.background()
        tf_b = badge.text_frame
        p = tf_b.paragraphs[0]
        p.text = num
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        tb_step = s8.shapes.add_textbox(x + Inches(0.2), y + Inches(0.75), Inches(2.35), Inches(1.2))
        tf_step = tb_step.text_frame
        tf_step.word_wrap = True
        p = tf_step.paragraphs[0]
        p.text = title
        p.font.size = Pt(12.5)
        p.font.bold = True
        p.font.color.rgb = NAVY_DEEP
        p.space_after = Pt(4)
        p = tf_step.add_paragraph()
        p.text = desc
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 9: Work Done – 1: Data Ingestion & Sanitization
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s9, BG_MUTED)
    add_top_nav(s9, "Work Done – 1", 9)

    add_card(s9, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.3))
    tb9 = s9.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(5.0), Inches(4.8))
    tf9 = tb9.text_frame
    tf9.word_wrap = True
    p = tf9.paragraphs[0]
    p.text = "Profile Data & Sanitization Engine"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(10)

    pts9 = [
        "Loaded employee datasets from data/employees.csv supporting CSV, Excel (.xlsx, .xls), and JSON.",
        "Normalized heterogeneous column headers using intelligent alias matching (e.g. 'CTC', 'Annual_Salary').",
        "Handled missing data by imputing default IDs, names, and mean performance scores.",
        "Derived analytical features: Experience Bands (Entry, Junior, Mid, Senior, Lead) and Salary Quartiles (Q1 to Q4).",
        "Calculated Total Compensation by combining Base Salary with percentage bonus awards."
    ]
    for pt in pts9:
        p = tf9.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)

    p = tf9.add_paragraph()
    p.text = "Technologies: Pandas • NumPy • OpenPyXL"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY

    code9 = """def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    current_cols = {col.lower().strip().replace(" ", "_"): col 
                    for col in df.columns}
    rename_map = {}
    for standard_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            norm = alias.lower().strip().replace(" ", "_")
            if norm in current_cols:
                rename_map[current_cols[norm]] = standard_name
                break
    df.rename(columns=rename_map, inplace=True)
    return df

def _derive_features(self, df: pd.DataFrame) -> pd.DataFrame:
    df["Experience_Band"] = df["Experience_Years"].apply(assign_band)
    df["Total_Compensation"] = df["Salary"] + df["Bonus_Amount"]
    df["Salary_Quartile"] = pd.qcut(df["Salary"], q=4, 
        labels=["Q1 (Entry)", "Q2 (Mid-)", "Q3 (Mid+)", "Q4 (Top)"])
    return df"""
    add_code_block(s9, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.3), code9, "src/data_loader.py")

    # =========================================================================
    # SLIDE 10: Work Done – 2: Analytics & Outlier Detection
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s10, BG_MUTED)
    add_top_nav(s10, "Work Done – 2", 10)

    add_card(s10, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.3))
    tb10 = s10.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(5.0), Inches(4.8))
    tf10 = tb10.text_frame
    tf10.word_wrap = True
    p = tf10.paragraphs[0]
    p.text = "Analytics & Anomaly Detection Engine"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(10)

    pts10 = [
        "Implemented executive summary metrics: Headcount, Total Payroll, Mean, Median, Std Dev, and IQR.",
        "Grouped by Department to compute Headcount share %, Payroll allocation %, and Average Experience.",
        "Engineered Tukey IQR Outlier Detection (Q1 - 1.5x IQR, Q3 + 1.5x IQR) to flag wage anomalies.",
        "Audited demographic pay parity measuring Female-to-Male salary ratio and unadjusted gender pay gap.",
        "Implemented dynamic multi-criteria filtering for custom salary thresholds, departments, and experience."
    ]
    for pt in pts10:
        p = tf10.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)

    p = tf10.add_paragraph()
    p.text = "Analytics Scope: Executive KPIs • Department Aggregations • IQR Fences"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY

    code10 = """def detect_outliers(self, method="iqr", threshold=1.5):
    sal = self.df["Salary"]
    if method == "iqr":
        q1 = sal.quantile(0.25)
        q3 = sal.quantile(0.75)
        iqr = q3 - q1
        lower = q1 - threshold * iqr
        upper = q3 + threshold * iqr
        is_high = sal > upper
        is_low = sal < lower
        self.df["Is_Outlier"] = is_high | is_low
        self.df["Outlier_Type"] = np.where(is_high, 
            "High Anomaly", np.where(is_low, "Low Anomaly", "Normal"))
        self.df["Outlier_Score"] = ((sal - sal.mean()) / sal.std()).round(2)
        
    return self.df[self.df["Is_Outlier"] == True].copy()"""
    add_code_block(s10, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.3), code10, "src/analytics_engine.py")

    # =========================================================================
    # SLIDE 11: Work Done – 3: Machine-Learning Compensation Predictor
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s11, BG_MUTED)
    add_top_nav(s11, "Work Done – 3", 11)

    add_card(s11, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.3))
    tb11 = s11.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(5.0), Inches(4.8))
    tf11 = tb11.text_frame
    tf11.word_wrap = True
    p = tf11.paragraphs[0]
    p.text = "Machine-Learning Predictive Modeling"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(10)

    pts11 = [
        "Trained an ensemble Random Forest regression model on empirical workforce parameters.",
        "Built a Pipeline combining StandardScaler for continuous features and OneHotEncoder for categoricals.",
        "Evaluated model performance achieving Explained Variance R² = 0.970 and Mean Absolute Error = ₹7,404.",
        "Calculated internal Compa-Ratio (Actual Salary / Expected Market Salary) to flag flight-risk employees.",
        "Exposes predict_salary() to power an interactive 'What-If' hiring compensation calculator."
    ]
    for pt in pts11:
        p = tf11.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)

    p = tf11.add_paragraph()
    p.text = "Model Features: Experience Years • Department • Education • Performance"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY

    code11 = """preprocessor = ColumnTransformer(transformers=[
    ("num", StandardScaler(), ["Experience_Years", "Performance_Score"]),
    ("cat", OneHotEncoder(handle_unknown="ignore"), ["Department", "Education"])
])

self.pipeline = Pipeline(steps=[
    ("preprocessor", preprocessor),
    ("regressor", RandomForestRegressor(n_estimators=100, max_depth=8, random_state=42))
])

self.pipeline.fit(X, y)
preds = self.pipeline.predict(X)
self.r2 = round(float(r2_score(y, preds)), 3)   # ~0.970
self.mae = round(float(mean_absolute_error(y, preds)), 2) # ~₹7,404"""
    add_code_block(s11, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.3), code11, "src/predictive_models.py")

    # =========================================================================
    # SLIDE 12: Work Done – 4: Streamlit Application & Output
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s12, BG_MUTED)
    add_top_nav(s12, "Work Done – 4", 12)

    add_card(s12, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.3))
    tb12 = s12.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(5.0), Inches(4.8))
    tf12 = tb12.text_frame
    tf12.word_wrap = True
    p = tf12.paragraphs[0]
    p.text = "Streamlit Application & Multi-Format Exports"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(10)

    pts12 = [
        "Built an interactive Streamlit web dashboard with custom CSS executive styling.",
        "Added real-time dataset switcher (1,200 Enterprise records vs. 8 Minimal sample records vs. Custom file upload).",
        "Implemented reactive sliders for salary thresholds, multi-department filters, and experience limits.",
        "Displayed KPI metric cards for Workforce Headcount, Total Payroll, Average Salary, and Salary Spread.",
        "Engineered 1-click downloads for styled multi-sheet Excel (.xlsx), standalone HTML briefs, and CSV."
    ]
    for pt in pts12:
        p = tf12.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)

    p = tf12.add_paragraph()
    p.text = "Frontend Stack: Streamlit • Altair • OpenPyXL • Tailwind HTML"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY

    code12 = """# Streamlit Interactive Dashboard Pipeline
dataset_source = st.sidebar.radio("Select Dataset Source:",
    ["Enterprise Dataset (1,200 records)", 
     "Original Minimal Sample (8 records)", 
     "Upload Custom CSV/Excel"])

filtered_df = full_engine.filter_records(
    min_salary=salary_range[0], max_salary=salary_range[1],
    departments=selected_depts, search_query=search_q
)

# 1-Click Excel Multi-Sheet Generation
sheets = {"Executive_KPIs": pd.DataFrame([kpis]),
          "Department_Summary": active_engine.get_department_summary(),
          "Filtered_Workforce": filtered_df}
# Download Button
st.download_button(label="📊 Download Excel", data=excel_buffer,
                   file_name="workforce_intelligence.xlsx")"""
    add_code_block(s12, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.3), code12, "app.py")

    # =========================================================================
    # SLIDE 13: Project Outputs – Streamlit Interface (Enterprise Screenshot)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s13, BG_MUTED)
    add_top_nav(s13, "Project Outputs – Streamlit Interface", 13)

    if os.path.exists(img_enterprise):
        add_card(s13, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3), border=INDIGO_PRIMARY)
        s13.shapes.add_picture(img_enterprise, Inches(0.85), Inches(1.55), width=Inches(11.6))

    # =========================================================================
    # SLIDE 14: Project Outputs – Minimal Sample Baseline
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s14, BG_MUTED)
    add_top_nav(s14, "Project Outputs – Baseline Benchmark", 14)

    if os.path.exists(img_minimal):
        add_card(s14, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3), border=CARD_BORDER)
        s14.shapes.add_picture(img_minimal, Inches(0.85), Inches(1.55), width=Inches(11.6))

    # =========================================================================
    # SLIDE 15: Project Outputs – Custom Dataset Intake
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s15, BG_MUTED)
    add_top_nav(s15, "Project Outputs – Universal File Intake", 15)

    if os.path.exists(img_upload):
        add_card(s15, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3), border=CARD_BORDER)
        s15.shapes.add_picture(img_upload, Inches(0.85), Inches(1.55), width=Inches(11.6))

    # =========================================================================
    # SLIDE 16: Learning After Internship – 1
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s16, BG_MUTED)
    add_top_nav(s16, "Learning After Internship – 1", 16)

    add_card(s16, Inches(0.8), Inches(1.5), Inches(7.5), Inches(5.3))
    tb16 = s16.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(6.9), Inches(4.8))
    tf16 = tb16.text_frame
    tf16.word_wrap = True
    p = tf16.paragraphs[0]
    p.text = "Data Engineering & Statistical Rigor"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(12)

    learnings1 = [
        "Learned how robust schema alias mapping enables universal data ingestion across varied corporate HR datasets.",
        "Understood the limitations of simple averages and mastered statistical dispersion (IQR, standard deviation, quartiles).",
        "Practiced Tukey IQR outlier detection to flag genuine compensation compression without parametric assumptions.",
        "Learned how demographic parity analysis uncovers hidden pay gaps across departments and experience brackets.",
        "Understood the importance of Compa-Ratio scoring in retaining top-performing talent."
    ]
    for l in learnings1:
        p = tf16.add_paragraph()
        p.text = f"•  {l}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(10)

    # Right Side Key Learning Card
    add_card(s16, Inches(8.6), Inches(1.5), Inches(3.9), Inches(5.3), bg=CARD_ACCENT_BG, border=INDIGO_PRIMARY)
    tb_kl1 = s16.shapes.add_textbox(Inches(8.9), Inches(1.8), Inches(3.3), Inches(4.7))
    tf_kl1 = tb_kl1.text_frame
    tf_kl1.word_wrap = True
    p = tf_kl1.paragraphs[0]
    p.text = "Key Learning"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY
    p.space_after = Pt(10)
    p = tf_kl1.add_paragraph()
    p.text = "Real-world workforce analytics requires combining statistical dispersion and demographic audits rather than relying on isolated mean metrics."
    p.font.size = Pt(12)
    p.font.color.rgb = NAVY_DEEP

    # =========================================================================
    # SLIDE 17: Learning After Internship – 2
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s17, BG_MUTED)
    add_top_nav(s17, "Learning After Internship – 2", 17)

    add_card(s17, Inches(0.8), Inches(1.5), Inches(7.5), Inches(5.3))
    tb17 = s17.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(6.9), Inches(4.8))
    tf17 = tb17.text_frame
    tf17.word_wrap = True
    p = tf17.paragraphs[0]
    p.text = "Machine Learning & Full-Stack Deployment"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(12)

    learnings2 = [
        "Improved practical Python, Pandas, and NumPy skills for large-scale tabular processing.",
        "Learned to integrate Scikit-Learn regression pipelines with ColumnTransformer for mixed feature types.",
        "Gained deep understanding of model evaluation metrics (R² score, Mean Absolute Error) and feature importance.",
        "Learned to build reactive data dashboards using Streamlit, Altair, and standalone Chart.js web applications.",
        "Mastered automated Git CI/CD workflows and production cloud deployment via GitHub Pages and edge tunnels."
    ]
    for l in learnings2:
        p = tf17.add_paragraph()
        p.text = f"•  {l}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(10)

    # Right Side Key Learning Card
    add_card(s17, Inches(8.6), Inches(1.5), Inches(3.9), Inches(5.3), bg=CARD_ACCENT_BG, border=INDIGO_PRIMARY)
    tb_kl2 = s17.shapes.add_textbox(Inches(8.9), Inches(1.8), Inches(3.3), Inches(4.7))
    tf_kl2 = tb_kl2.text_frame
    tf_kl2.word_wrap = True
    p = tf_kl2.paragraphs[0]
    p.text = "Key Learning"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = INDIGO_PRIMARY
    p.space_after = Pt(10)
    p = tf_kl2.add_paragraph()
    p.text = "The project connected statistical data cleaning, machine learning pipelines, interactive UI, and public cloud deployment into one unified end-to-end system."
    p.font.size = Pt(12)
    p.font.color.rgb = NAVY_DEEP

    # =========================================================================
    # SLIDE 18: Summary / Conclusion
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s18, BG_MUTED)
    add_top_nav(s18, "Summary / Conclusion", 18)

    add_card(s18, Inches(0.8), Inches(1.5), Inches(7.5), Inches(5.3))
    tb18 = s18.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(6.9), Inches(4.8))
    tf18 = tb18.text_frame
    tf18.word_wrap = True
    p = tf18.paragraphs[0]
    p.text = "Internship & Project Outcome"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(12)

    concls = [
        "Completed the Unlox Academy Machine Learning Internship / Major Project.",
        "Developed an Enterprise Workforce Intelligence & Compensation Analytics Platform.",
        "Combined statistical dispersion, demographic pay parity audits, and Tukey IQR anomaly detection.",
        "Integrated a machine-learning Random Forest compensation benchmarking layer (R² = 0.970).",
        "Built dual interfaces: an interactive Streamlit web dashboard and a scriptable CLI.",
        "Strengthened practical engineering skills in Python, ETL, Scikit-Learn, Streamlit, and Cloud Deployment."
    ]
    for c in concls:
        p = tf18.add_paragraph()
        p.text = f"•  {c}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)

    # Right side: Embed BeeSkilled Certificate Image
    if os.path.exists(img_cert):
        add_card(s18, Inches(8.3), Inches(1.5), Inches(4.3), Inches(5.3), border=INDIGO_PRIMARY)
        s18.shapes.add_picture(img_cert, Inches(8.35), Inches(1.9), width=Inches(4.2))
        tb_c = s18.shapes.add_textbox(Inches(8.35), Inches(1.55), Inches(4.2), Inches(0.35))
        tf_c = tb_c.text_frame
        p = tf_c.paragraphs[0]
        p.text = "BeeSkilled Verified Internship Certificate"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = INDIGO_PRIMARY
        p.alignment = PP_ALIGN.CENTER
    else:
        add_card(s18, Inches(8.6), Inches(1.5), Inches(3.9), Inches(2.5), bg=CARD_ACCENT_BG)
        tb_c1 = s18.shapes.add_textbox(Inches(8.8), Inches(1.7), Inches(3.5), Inches(2.1))
        tf_c1 = tb_c1.text_frame
        p = tf_c1.paragraphs[0]
        p.text = "Certificate"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = INDIGO_PRIMARY
        p = tf_c1.add_paragraph()
        p.text = "BeeSkilled Internship Certificate\nPython Online Internship Program\nIssued: 07 October 2026"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = NAVY_DEEP

    # =========================================================================
    # SLIDE 19: Bibliography / References
    # =========================================================================
    s19 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s19, BG_MUTED)
    add_top_nav(s19, "Bibliography / References", 19)

    add_card(s19, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.3))
    tb19 = s19.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.1), Inches(4.8))
    tf19 = tb19.text_frame
    tf19.word_wrap = True
    p = tf19.paragraphs[0]
    p.text = "References & Academic Citations"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY_DEEP
    p.space_after = Pt(14)

    refs = [
        "BeeSkilled — Python Online Internship Program (Tech and Internships) curriculum.",
        "BeeSkilled Certificate of Completion — Awarded to Nikhil G, Issued 07 October 2026.",
        "Python Software Foundation — Python 3.14 Documentation & Standard Library.",
        "Pandas & NumPy Documentation — Tabular Data Aggregation, Indexing, and Array Operations.",
        "Scikit-Learn Developers — Ensemble Methods, Random Forest Regressor & Preprocessing Pipelines.",
        "Streamlit Framework — Reactive Web Applications and Component Architecture.",
        "John W. Tukey — 'Exploratory Data Analysis' (Addison-Wesley, 1977) for IQR Boxplot Outlier Rules.",
        "GitHub Repository: https://github.com/Nikhil8-coder/employee-workforce-analytics",
        "Live Web Application: https://nikhil8-coder.github.io/employee-workforce-analytics/"
    ]
    for r in refs:
        p = tf19.add_paragraph()
        p.text = f"•  {r}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(6)

    # =========================================================================
    # SLIDE 20: Thank You Slide (Dark Executive Theme)
    # =========================================================================
    s20 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s20, NAVY_DEEP)

    tb20 = s20.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(4.0))
    tf20 = tb20.text_frame
    tf20.word_wrap = True

    p = tf20.paragraphs[0]
    p.text = "THANK YOU"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_after = Pt(14)

    p = tf20.add_paragraph()
    p.text = "Python Internship • BeeSkilled (Tech & Internships)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(199, 210, 254)
    p.space_after = Pt(8)

    p = tf20.add_paragraph()
    p.text = "Project: Enterprise Workforce Intelligence & Compensation Analytics Platform"
    p.font.size = Pt(15)
    p.font.color.rgb = RGBColor(224, 231, 255)
    p.space_after = Pt(16)

    p = tf20.add_paragraph()
    p.text = "G. Nikhil (Nikhil G)\nCSE (AI&ML) • Geethanjali College of Engineering and Technology\nA.Y. 2026–27"
    p.font.size = Pt(13)
    p.font.color.rgb = RGBColor(199, 210, 254)

    prs.save(output_path)
    print(f"College format presentation successfully built at: {output_path}")

if __name__ == "__main__":
    out = r"C:\Users\nikhi\.gemini\antigravity\scratch\employee_analytics\GCET_Workforce_Analytics_Presentation.pptx"
    build_presentation(out)

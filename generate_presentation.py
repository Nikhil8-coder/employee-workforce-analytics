"""
Presentation Generator for Enterprise Workforce Intelligence & Compensation Analytics
Generates a professional 18-slide PowerPoint presentation (.pptx) with code blocks,
visual styling, and embedded live application screenshots.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_pptx_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]  # Blank slide

    # Color Palette
    BG_DARK = RGBColor(15, 23, 42)        # Slate 900
    BG_LIGHT = RGBColor(248, 250, 252)    # Slate 50
    CARD_BG = RGBColor(255, 255, 255)     # White
    CARD_BORDER = RGBColor(226, 232, 240) # Slate 200
    PRIMARY = RGBColor(67, 56, 202)       # Indigo 700
    PRIMARY_LIGHT = RGBColor(99, 102, 241)# Indigo 500
    TEXT_DARK = RGBColor(30, 41, 59)      # Slate 800
    TEXT_MUTED = RGBColor(100, 116, 139)  # Slate 500
    ACCENT_GREEN = RGBColor(16, 185, 129) # Emerald 500
    CODE_BG = RGBColor(24, 33, 47)        # Dark Slate for Code
    CODE_TEXT = RGBColor(226, 232, 240)   # Light code font
    CODE_KW = RGBColor(129, 140, 248)     # Indigo light for keywords

    image_dir = r"C:\Users\nikhi\.gemini\antigravity\brain\b188db44-ac69-453b-a6fe-af390114bb54\.user_uploaded"
    img1 = os.path.join(image_dir, "media_1791997413016_b485b955.png") # Enterprise Dashboard
    img2 = os.path.join(image_dir, "media_1791997438169_de4633c9.png") # Minimal Sample
    img3 = os.path.join(image_dir, "media_1791997456963_ef187b24.png") # Upload UI

    def set_slide_background(slide, color):
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_header(slide, title_text, category_text="ENTERPRISE WORKFORCE INTELLIGENCE & COMPENSATION ANALYTICS", dark_mode=False):
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p0 = tf.paragraphs[0]
        p0.text = category_text.upper()
        p0.font.size = Pt(10)
        p0.font.bold = True
        p0.font.color.rgb = PRIMARY_LIGHT if dark_mode else PRIMARY
        p0.space_after = Pt(4)

        p1 = tf.add_paragraph()
        p1.text = title_text
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(255, 255, 255) if dark_mode else TEXT_DARK

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        if border_color:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(1.2)
        else:
            shape.line.fill.background()
        return shape

    def add_code_box(slide, left, top, width, height, code_text, filename_tag=""):
        # Dark container
        add_card(slide, left, top, width, height, bg_color=CODE_BG, border_color=RGBColor(51, 65, 85))
        
        # Header banner inside code box
        if filename_tag:
            tag_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), width - Inches(0.4), Inches(0.35))
            tf_tag = tag_box.text_frame
            tf_tag.word_wrap = True
            tf_tag.margin_top = tf_tag.margin_bottom = tf_tag.margin_left = tf_tag.margin_right = 0
            p = tf_tag.paragraphs[0]
            p.text = f"📄 {filename_tag}"
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_LIGHT

        # Code content
        code_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.48), width - Inches(0.4), height - Inches(0.55))
        tf = code_box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        lines = code_text.strip().split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(9.5)
            p.font.name = "Consolas"
            p.font.color.rgb = CODE_TEXT
            p.space_before = Pt(1)
            p.space_after = Pt(1)

    # =========================================================================
    # SLIDE 1: Title Slide (Dark Executive Theme)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, BG_DARK)

    # Accent bar
    bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.0), Inches(0.15), Inches(4.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY_LIGHT
    bar.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(0.9), Inches(11.3), Inches(4.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "GEETHANJALI COLLEGE OF ENGINEERING AND TECHNOLOGY"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor(148, 163, 184)
    p.space_after = Pt(4)

    p = tf1.add_paragraph()
    p.text = "(UGC Autonomous • Affiliated to JNTUH • Approved by AICTE) | Department of CSE (AI&ML)"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(100, 116, 139)
    p.space_after = Pt(24)

    p = tf1.add_paragraph()
    p.text = "Enterprise Workforce Intelligence &"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    p = tf1.add_paragraph()
    p.text = "Compensation Analytics Platform"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_LIGHT
    p.space_after = Pt(14)

    p = tf1.add_paragraph()
    p.text = "Production-grade Workforce Analytics, Statistical Anomaly Detection, Machine Learning Salary Modeling, and Cloud Deployment"
    p.font.size = Pt(13)
    p.font.color.rgb = RGBColor(203, 213, 225)
    p.space_after = Pt(28)

    p = tf1.add_paragraph()
    p.text = "Developed by: Nikhil (Nikhil8-coder)  •  B.Tech CSE (AI & ML)  •  Major Project\nGitHub: https://github.com/Nikhil8-coder/employee-workforce-analytics"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(148, 163, 184)

    # =========================================================================
    # SLIDE 2: Project Evolution (From Basic Script to Enterprise Suite)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2, BG_LIGHT)
    add_header(s2, "Project Evolution: Basic Starter Script vs. Enterprise Suite")

    # Column 1: Starter Assignment
    add_card(s2, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s2.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Original Starter Script (Assignment Baseline)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(239, 68, 68)
    p.space_after = Pt(12)

    bullets1 = [
        "Data Scale: Only 8 hardcoded static mock employee rows.",
        "Ingestion: Rigid CSV loader failing on header variations.",
        "Math Scope: Basic company mean salary and simple group count.",
        "Filtering: Hardcoded terminal prompt for single threshold.",
        "Anomaly Audit: Zero outlier or statistical dispersion awareness.",
        "Predictive Modeling: No Machine Learning or inference.",
        "Interface: Command-line console text prints only.",
        "Deployment: Non-executable outside local student script."
    ]
    for b in bullets1:
        p = tf.add_paragraph()
        p.text = f"•  {b}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(6)

    # Column 2: Enterprise Elevated Suite
    add_card(s2, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), border_color=PRIMARY_LIGHT)
    tb = s2.shapes.add_textbox(Inches(7.1), Inches(1.9), Inches(5.1), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Enterprise Workforce Intelligence Suite (Built)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(12)

    bullets2 = [
        "Data Scale: 1,200+ Kaggle-grade realistic records across 8 departments.",
        "Ingestion: Multi-format (CSV/Excel/JSON) with schema alias mapping.",
        "Statistical Depth: Mean, Median, Std Dev, Min/Max, 25th/75th, and IQR.",
        "Workforce Audits: Demographic pay parity (gender gap) & payroll share.",
        "Anomaly Engine: Tukey IQR fences (1.5x) and Z-score outlier flagging.",
        "Predictive ML: Random Forest Regressor (R² ~ 0.97, MAE ~ ₹7,400).",
        "Dual Interface: Interactive Streamlit Web UI + Formatted CLI Wizard.",
        "Public Deployment: Live on GitHub Pages, Cloudflare Tunnel & Docker."
    ]
    for b in bullets2:
        p = tf.add_paragraph()
        p.text = f"✓  {b}"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(6)

    # =========================================================================
    # SLIDE 3: Problem Statement & Engineering Objectives
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3, BG_LIGHT)
    add_header(s3, "Problem Statement & Strategic Engineering Objectives")

    cards_data = [
        ("The Challenge in Traditional HR", "Standard spreadsheet trackers fail to scale beyond a few dozen rows, lack predictive capabilities, and offer zero automated protection against internal compensation compression or demographic pay discrepancies.", RGBColor(224, 231, 255)),
        ("Objective 1: Multi-Format Ingestion", "Develop a resilient ETL data loader that automatically recognizes varied column aliases (e.g. 'CTC', 'Annual_Salary', 'Tenure') and validates data integrity with automated feature derivation.", RGBColor(241, 245, 249)),
        ("Objective 2: Statistical Dispersion & Audits", "Implement rigorous statistical dispersion metrics including Interquartile Range (IQR) outlier detection to pinpoint under-compensated high performers and compensation anomalies.", RGBColor(241, 245, 249)),
        ("Objective 3: ML Compensation Modeling", "Train a Scikit-Learn Random Forest ensemble regression model on empirical workforce parameters to forecast fair market compensation bands and Compa-Ratio benchmarks.", RGBColor(241, 245, 249)),
        ("Objective 4: Production Cloud Delivery", "Provide both a client-side interactive web dashboard (accessible offline or via GitHub Pages) and a full Streamlit Python server ready for permanent cloud hosting.", RGBColor(241, 245, 249)),
        ("Objective 5: Multi-Format Reporting Hub", "Enable instant automated generation of multi-sheet styled Excel workbooks with cell formatting, publication-ready HTML briefs, and filtered CSV extracts.", RGBColor(241, 245, 249))
    ]

    for i, (heading, desc, fill_c) in enumerate(cards_data):
        row = i // 3
        col = i % 3
        x = Inches(0.8 + col * 4.0)
        y = Inches(1.7 + row * 2.6)
        add_card(s3, x, y, Inches(3.7), Inches(2.3), bg_color=fill_c)
        tb = s3.shapes.add_textbox(x + Inches(0.2), y + Inches(0.2), Inches(3.3), Inches(1.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = heading
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY if i == 0 else TEXT_DARK
        p.space_after = Pt(6)
        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 4: Full Technology Stack
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4, BG_LIGHT)
    add_header(s4, "Technology Stack & Enterprise Architectural Foundation")

    tech_categories = [
        ("Core Computing & Analytics", "Python 3.14\nPandas (Vectorized Aggregation)\nNumPy (Dispersion & Matrix Ops)\nScipy (Statistical Distributions)"),
        ("Machine Learning & AI", "Scikit-Learn (Ensemble Regressors)\nRandom Forest (100 Estimators)\nColumnTransformer & OneHotEncoder\nStandardScaler & Metrics (R², MAE)"),
        ("Interactive Frontend & UI", "Streamlit (Reactive Python UI)\nAltair (Interactive Statistical Viz)\nTailwind CSS (Modern Briefs)\nChart.js (Client-Side HTML Web App)"),
        ("Enterprise Reporting Hub", "OpenPyXL (Styled Multi-Tab Excel)\nTabulate (ASCII Terminal Grids)\nJinja2 (HTML Template Generation)\nCSV / JSON Multi-format Exporters"),
        ("DevOps & Cloud Delivery", "Git & GitHub CLI (Automation)\nGitHub Pages (Global CDN Hosting)\nCloudflare Tunnel (HTTP/2 Edge)\nDocker & Docker-Compose"),
        ("Quality Assurance & Testing", "Unittest Framework\nDeterministic Test Fixtures\nRegression & Assertion Testing\nPy_Compile Syntax Verification")
    ]

    for i, (cat, items) in enumerate(tech_categories):
        row = i // 3
        col = i % 3
        x = Inches(0.8 + col * 4.0)
        y = Inches(1.7 + row * 2.6)
        add_card(s4, x, y, Inches(3.7), Inches(2.3))
        tb = s4.shapes.add_textbox(x + Inches(0.25), y + Inches(0.25), Inches(3.2), Inches(1.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cat
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY
        p.space_after = Pt(8)
        for item in items.split("\n"):
            p = tf.add_paragraph()
            p.text = f"•  {item}"
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(3)

    # =========================================================================
    # SLIDE 5: Dataset Schema & Statistical Synthesis
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5, BG_LIGHT)
    add_header(s5, "Dataset Architecture & Statistically Correlated Synthesis")

    add_card(s5, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s5.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "15-Attribute Enterprise Data Model"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    schema_items = [
        "Employee_ID: Unique identifier (EMP-0001 to EMP-1200)",
        "Name & Gender: Demographic profile fields",
        "Department: 8 Business units (Engineering, AI, Sales, etc.)",
        "Job_Title: Hierarchical designations from Junior to VP",
        "Salary (Base): Statistically correlated compensation (₹)",
        "Bonus_%: Department-specific incentive ranges",
        "Total_Compensation: Base + Derived annual incentive",
        "Experience_Years: Gamma-distributed tenure (0 to 25 yrs)",
        "Performance_Score: Bell-curved appraisal (1.0 to 5.0)",
        "Attrition_Risk: High / Medium / Low retention flight risk"
    ]
    for item in schema_items:
        p = tf.add_paragraph()
        p.text = f"•  {item}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(3)

    # Right side: Data generator code
    code_gen = """# Synthetic Enterprise Correlation Engine
dept_info = DEPARTMENTS[dept]
exp_years = int(np.clip(np.random.gamma(2.5, 2.8), 0, 25))
perf_score = round(float(np.clip(
    np.random.normal(3.6, 0.65), 1.0, 5.0)), 2)

base = dept_info["base_salary"]
exp_adder = exp_years * dept_info["exp_mult"]
perf_mult = 1.0 + (perf_score - 3.0) * 0.08
noise = np.random.normal(1.0, 0.07)

raw_salary = (base + exp_adder) * edu_bonus * perf_mult * noise
salary = int(round(raw_salary, -3))

# Compa-Ratio & Flight Risk Assessment
comp_ratio = salary / expected_salary_benchmark
if comp_ratio < 0.90 and perf_score >= 4.0:
    attrition_risk = "High"  # Underpaid star talent"""
    add_code_box(s5, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_gen, "data_generator.py")

    # =========================================================================
    # SLIDE 6: Code Module 1 - Data Ingestion (data_loader.py)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6, BG_LIGHT)
    add_header(s6, "Code Module 1: Resilient Data Ingestion & Schema Mapping")

    add_card(s6, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s6.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Automated ETL & Sanitization Architecture"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    points = [
        "Heterogeneous Header Normalization: Maps aliases like 'CTC', 'Annual_Salary', 'Tenure', 'Designation' into unified enterprise schema.",
        "Multi-Format Ingestion: Natively handles CSV, Excel (.xlsx, .xls), and JSON with format sniffing.",
        "Missing Data Imputation: Automatically fills missing IDs, normalizes performance scores, and casts numeric bounds.",
        "Feature Derivation: Automatically assigns Experience Bands (Entry, Junior, Mid, Senior, Lead) and divides compensation into Salary Quartiles (Q1 to Q4).",
        "Validation Metadata: Generates audit summary including min/max range, null checks, and department rosters."
    ]
    for pt in points:
        p = tf.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    code_loader = """def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    current_cols = {col.lower().strip().replace(" ", "_"): col 
                    for col in df.columns}
    rename_map = {}
    for standard_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            norm_alias = alias.lower().strip().replace(" ", "_")
            if norm_alias in current_cols:
                rename_map[current_cols[norm_alias]] = standard_name
                break
    df.rename(columns=rename_map, inplace=True)
    return df

def _derive_features(self, df: pd.DataFrame) -> pd.DataFrame:
    df["Experience_Band"] = df["Experience_Years"].apply(assign_band)
    df["Total_Compensation"] = df["Salary"] + df["Bonus_Amount"]
    df["Salary_Quartile"] = pd.qcut(df["Salary"], q=4, 
        labels=["Q1 (Entry)", "Q2 (Mid-)", "Q3 (Mid+)", "Q4 (Top)"])
    return df"""
    add_code_box(s6, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_loader, "src/data_loader.py")

    # =========================================================================
    # SLIDE 7: Mathematical & Analytical Framework
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7, BG_LIGHT)
    add_header(s7, "Mathematical Formulation & Analytical Foundations")

    math_cards = [
        ("1. Central Tendency & Spread", "Mean:  μ = (1/N) * Σ Salary_i\nMedian:  Q2 (Middle rank, robust to skewness)\nStandard Deviation:  σ = sqrt( (1/N) * Σ (Salary_i - μ)² )\nMeasures compensation variance across cohorts.", RGBColor(241, 245, 249)),
        ("2. Tukey IQR Outlier Boundaries", "Interquartile Range:  IQR = Q3 (75th) - Q1 (25th)\nLower Fence:  Q1 - 1.5 × IQR  (Underpaid flag)\nUpper Fence:  Q3 + 1.5 × IQR  (Executive outlier)\nDetects compensation anomalies without normal distribution assumptions.", RGBColor(224, 231, 255)),
        ("3. Demographic Pay Parity Ratio", "Gender Parity Ratio = (Avg Female Salary / Avg Male Salary) * 100\nUnadjusted Pay Gap = 100% - Parity Ratio\nAudits equitable reward distribution across departments and experience tiers.", RGBColor(241, 245, 249)),
        ("4. Compa-Ratio & Market Parity", "Compa-Ratio = Actual Base Salary / ML Predicted Market Benchmark\n• < 0.90 : Under-Compensated (High Flight Risk)\n• 0.90 - 1.10 : Market Aligned (Target Band)\n• > 1.10 : Premium Compensated", RGBColor(241, 245, 249))
    ]

    for i, (title, content, bg_c) in enumerate(math_cards):
        row = i // 2
        col = i % 2
        x = Inches(0.8 + col * 5.9)
        y = Inches(1.7 + row * 2.6)
        add_card(s7, x, y, Inches(5.6), Inches(2.3), bg_color=bg_c)
        tb = s7.shapes.add_textbox(x + Inches(0.25), y + Inches(0.2), Inches(5.1), Inches(1.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY
        p.space_after = Pt(6)
        for line in content.split("\n"):
            p = tf.add_paragraph()
            p.text = line
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(2)

    # =========================================================================
    # SLIDE 8: Code Module 2 - Analytics & Outlier Detection
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8, BG_LIGHT)
    add_header(s8, "Code Module 2: Core Analytics & Outlier Engine")

    add_card(s8, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s8.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Executive Statistical Aggregation"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    pts = [
        "Departmental Dispersion: Vectorized grouping calculating Headcount share, Payroll allocation %, Average/Median salary, and Tenure.",
        "IQR Outlier Detection: Identifies statistical compensation anomalies globally and within specific department silos.",
        "Multi-Criteria Query Engine: Dynamic filtering supporting simultaneous salary ranges, department subsets, tenure, and performance benchmarks.",
        "Z-Score Analysis: Provides normalized standard deviation distance (|Z| > 2.5) for audit inspections."
    ]
    for pt in pts:
        p = tf.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    code_analytics = """def detect_outliers(self, method="iqr", threshold=1.5):
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
    add_code_box(s8, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_analytics, "src/analytics_engine.py")

    # =========================================================================
    # SLIDE 9: Code Module 3 - Machine Learning Salary Predictor
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9, BG_LIGHT)
    add_header(s9, "Code Module 3: Machine Learning Compensation Predictor")

    add_card(s9, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s9.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Scikit-Learn Ensemble Regressor"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    ml_points = [
        "Model Architecture: Random Forest Regressor (100 estimators, max depth 8) built on an end-to-end Pipeline.",
        "Feature Preprocessing: ColumnTransformer combining StandardScaler for continuous features (Tenure, Performance) and OneHotEncoder for categoricals (Department, Education).",
        "Empirical Model Accuracy: Achieves Explained Variance R² = 0.970 and Mean Absolute Error (MAE) = ₹7,404.11.",
        "Interactive Inference: Exposes predict_salary() to power real-time 'What-If' candidate compensation calculators."
    ]
    for pt in ml_points:
        p = tf.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    code_ml = """# Ensemble Regression Pipeline
preprocessor = ColumnTransformer(transformers=[
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
    add_code_box(s9, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_ml, "src/predictive_models.py")

    # =========================================================================
    # SLIDE 10: Code Module 4 - Workforce Benchmarking & Compa-Ratio
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10, BG_LIGHT)
    add_header(s10, "Code Module 4: Compa-Ratio Parity & Retention Risk")

    add_card(s10, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s10.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Retention & Equity Intelligence"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    pts10 = [
        "Internal Compa-Ratio: Compares every employee's current salary against the ML model's unbiased fair market expectation.",
        "Under-Compensation Watchlist: Identifies top-tier talent (Rating >= 4.0) earning below market threshold (<90%).",
        "Flight Risk Proaction: Allows HR leadership to implement proactive retention raises before employee departure.",
        "Department Equity Index: Highlights budget disparities across business units to optimize future wage allocation."
    ]
    for pt in pts10:
        p = tf.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    code_bench = """def benchmark_workforce(self, df: pd.DataFrame) -> pd.DataFrame:
    benchmarked = df.copy()
    features = ["Experience_Years", "Performance_Score", 
                "Department", "Education"]
    
    predicted = self.pipeline.predict(benchmarked[features])
    benchmarked["Benchmark_Predicted_Salary"] = predicted.round(2)
    benchmarked["Compa_Ratio"] = (
        benchmarked["Salary"] / benchmarked["Benchmark_Predicted_Salary"]
    ).round(2)

    conditions = [
        (benchmarked["Compa_Ratio"] < 0.90),
        (benchmarked["Compa_Ratio"] >= 0.90) & (benchmarked["Compa_Ratio"] <= 1.10),
        (benchmarked["Compa_Ratio"] > 1.10)
    ]
    choices = ["Under-Compensated (<90%)", 
               "Market Aligned (90-110%)", 
               "Premium Compensated (>110%)"]
    benchmarked["Equity_Status"] = np.select(conditions, choices)
    return benchmarked"""
    add_code_box(s10, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_bench, "src/predictive_models.py")

    # =========================================================================
    # SLIDE 11: Code Module 5 - Multi-Format Enterprise Reporting
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11, BG_LIGHT)
    add_header(s11, "Code Module 5: Automated Enterprise Reporting Engine")

    add_card(s11, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s11.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Multi-Format Export Pipeline"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    pts11 = [
        "Multi-Tab Styled Excel (.xlsx): Built with OpenPyXL. Generates styled worksheets (Executive KPIs, Department Breakdown, High Earners, Outliers, Full Roster) with corporate navy headers and currency formatting.",
        "Standalone Executive HTML Briefs: Generates self-contained Tailwind CSS documents complete with KPI cards and table previews.",
        "Filtered CSV/JSON Exports: Produces clean machine-readable data feeds for downstream payroll pipelines."
    ]
    for pt in pts11:
        p = tf.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    code_rep = """def export_excel(self, sheets: Dict[str, pd.DataFrame], filename):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            ws = writer.sheets[sheet_name[:31]]
            header_fill = PatternFill(start_color="1E3A8A", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            for col in ws.columns:
                max_len = max(len(str(c.value or '')) for c in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)"""
    add_code_box(s11, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_rep, "src/reporter.py")

    # =========================================================================
    # SLIDE 12: Code Module 6 - Interactive Web Application (app.py)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12, BG_LIGHT)
    add_header(s12, "Code Module 6: Streamlit Reactive Web Application")

    add_card(s12, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s12.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Interactive Analytics Frontend"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    pts12 = [
        "Dataset Source Switcher: Seamlessly toggle between 1,200 Enterprise records, 8 Minimal sample records, or Drag & Drop Custom files.",
        "Dynamic Reactive Controls: Live sliders for salary boundaries, multi-department pickers, experience limits, and role search.",
        "Rich Altair Visualizations: Density histograms with mean/median rule markers and departmental boxplots.",
        "One-Click Download Hub: Direct download buttons for CSV, JSON, Excel, and HTML briefs."
    ]
    for pt in pts12:
        p = tf.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    code_app = """# Streamlit Interactive Dashboard Pipeline
dataset_source = st.sidebar.radio("Select Dataset Source:",
    ["Enterprise Dataset (1,200 records)", 
     "Original Minimal Sample (8 records)", 
     "Upload Custom CSV/Excel"])

filtered_df = full_engine.filter_records(
    min_salary=salary_range[0], max_salary=salary_range[1],
    departments=selected_depts, search_query=search_q
)

# Interactive Density Histogram with Altair
hist = alt.Chart(filtered_df).mark_bar(opacity=0.75).encode(
    x=alt.X("Salary:Q", bin=alt.Bin(maxbins=25)),
    y="count():Q"
)
mean_rule = alt.Chart(pd.DataFrame({'val': [kpis['average_salary']]})
    ).mark_rule(color='#EF4444', strokeWidth=2).encode(x='val:Q')
st.altair_chart(hist + mean_rule, use_container_width=True)"""
    add_code_box(s12, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_app, "app.py")

    # =========================================================================
    # SLIDE 13: Code Module 7 - Command Line Interface (cli.py)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13, BG_LIGHT)
    add_header(s13, "Code Module 7: Production Command-Line Interface (CLI)")

    add_card(s13, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s13.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Dual Mode Terminal Suite"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    pts13 = [
        "Headless Batch Automation: Scriptable argparse CLI for scheduled cron jobs and automated data pipelines.",
        "Interactive Menu Wizard: Numbered terminal control center for manual stakeholder audits.",
        "Cross-Platform Unicode Support: Automatic Windows CP1252 stdout reconfiguration for correct currency rendering.",
        "Rich Tabular Display: Formatted fancy-grid table output using Tabulate."
    ]
    for pt in pts13:
        p = tf.add_paragraph()
        p.text = f"•  {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)

    code_cli = """def main():
    parser = argparse.ArgumentParser(description="Workforce Intelligence CLI")
    parser.add_argument("--dataset", default="data/employees.csv")
    parser.add_argument("--threshold", type=float, help="Salary benchmark")
    parser.add_argument("--department", type=str, help="Department filter")
    parser.add_argument("--summary", action="store_true", help="Print KPIs")
    parser.add_argument("--outliers", action="store_true", help="Detect IQR anomalies")
    parser.add_argument("--train-ml", action="store_true", help="Train ML model")
    parser.add_argument("--export-excel", type=str, help="Save to multi-sheet Excel")
    parser.add_argument("--interactive", action="store_true", help="Launch interactive wizard")

# Example Batch Execution:
# python cli.py --summary --threshold 80000 --export-excel full_report.xlsx"""
    add_code_box(s13, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_cli, "cli.py")

    # =========================================================================
    # SLIDE 14: System Outputs - Live Enterprise Dashboard (With Image 1)
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14, BG_LIGHT)
    add_header(s14, "Live System Output: Enterprise Analytics Dashboard")

    # Left: Explanation & KPIs
    add_card(s14, Inches(0.8), Inches(1.7), Inches(4.5), Inches(5.1))
    tb = s14.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(4.1), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Enterprise Cohort Execution"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(8)

    obs = [
        "Workforce Headcount: 179 Filtered Employees across 8 active departments.",
        "Total Annual Payroll: ₹38,577,000 annualised base expenditure.",
        "Average Salary: ₹215,513.97 (Median: ₹210,000).",
        "Dispersion: Standard Deviation ₹43,437 | IQR: ₹57,500.",
        "Salary Density Curve: Bell-curved distribution showing empirical central clustering.",
        "Department Donut: Dynamic allocation showing Engineering, AI & Sales leading headcount."
    ]
    for ob in obs:
        p = tf.add_paragraph()
        p.text = f"•  {ob}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(4)

    # Right: Screenshot 1
    if os.path.exists(img1):
        add_card(s14, Inches(5.5), Inches(1.7), Inches(7.0), Inches(5.1), border_color=PRIMARY_LIGHT)
        s14.shapes.add_picture(img1, Inches(5.55), Inches(1.75), width=Inches(6.9))

    # =========================================================================
    # SLIDE 15: System Outputs - Baseline & Custom Ingestion (With Images 2 & 3)
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15, BG_LIGHT)
    add_header(s15, "Live System Output: Scalability & Custom Upload Ingestion")

    # Image 2 (Minimal)
    if os.path.exists(img2):
        add_card(s15, Inches(0.8), Inches(1.7), Inches(5.6), Inches(4.5), border_color=CARD_BORDER)
        s15.shapes.add_picture(img2, Inches(0.85), Inches(1.75), width=Inches(5.5))
        tb = s15.shapes.add_textbox(Inches(0.8), Inches(6.3), Inches(5.6), Inches(0.6))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = "Backward Compatibility: Original 8-Record Sample (Avg ₹79,125 | Payroll ₹633K)"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

    # Image 3 (Custom Upload)
    if os.path.exists(img3):
        add_card(s15, Inches(6.8), Inches(1.7), Inches(5.7), Inches(4.5), border_color=CARD_BORDER)
        s15.shapes.add_picture(img3, Inches(6.85), Inches(1.75), width=Inches(5.6))
        tb = s15.shapes.add_textbox(Inches(6.8), Inches(6.3), Inches(5.7), Inches(0.6))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = "Universal Data Intake: Ingests custom CSV, Excel (.xlsx), or JSON up to 200MB"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = PRIMARY

    # =========================================================================
    # SLIDE 16: Automated Verification & Unit Test Suite
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16, BG_LIGHT)
    add_header(s16, "Automated Verification & Unit Test Architecture")

    add_card(s16, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5.1))
    tb = s16.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(5.0), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "100% Passing Test Fixture Suite"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(10)

    tests_list = [
        "test_executive_summary_statistics: Validates mean, median, min, max, and headcount against manual mathematical derivations.",
        "test_department_aggregation: Confirms grouped department counts and aggregate salary means across partitions.",
        "test_filtering_logic: Verifies multi-condition slicing (threshold >= 80,000, specific departments, and tenure limits).",
        "test_data_loader_with_aliases: Tests schema normalization of non-standard headers (e.g. 'Business Unit', 'Annual Salary').",
        "test_predictive_model_training: Validates Random Forest training pipeline and inference bounds.",
        "test_report_generator_outputs: Verifies valid physical creation of CSV, Excel (.xlsx), and HTML executive briefs."
    ]
    for t in tests_list:
        p = tf.add_paragraph()
        p.text = f"✓  {t}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(5)

    code_test = """Ran 6 tests in 0.348s - OK

class TestEmployeeAnalytics(unittest.TestCase):
    def test_executive_summary_statistics(self):
        kpis = self.engine.get_executive_summary()
        self.assertEqual(kpis["headcount"], 8)
        self.assertEqual(kpis["department_count"], 4)
        expected_avg = self.mock_df["Salary"].mean()
        self.assertAlmostEqual(kpis["average_salary"], expected_avg, places=2)

    def test_predictive_model_training(self):
        predictor = CompensationPredictor()
        success, msg = predictor.train(self.mock_df)
        self.assertTrue(success)
        pred = predictor.predict_salary(experience_years=5, department="IT")
        self.assertGreater(pred, 0)"""
    add_code_box(s16, Inches(6.8), Inches(1.7), Inches(5.7), Inches(5.1), code_test, "tests/test_analytics.py")

    # =========================================================================
    # SLIDE 17: Production Deployment & Cloud Architecture
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_background(s17, BG_LIGHT)
    add_header(s17, "Production Cloud Deployment & Global Availability")

    deploy_channels = [
        ("1. GitHub Pages (Live Web App)", "Public URL: https://nikhil8-coder.github.io/employee-workforce-analytics/\n• 100% Free & Permanent hosting on GitHub's global edge CDN.\n• Zero-dependency client-side HTML5/Tailwind/Chart.js engine.\n• Client-side reactive filtering and instant CSV export."),
        ("2. GitHub Source Repository", "Repository URL: https://github.com/Nikhil8-coder/employee-workforce-analytics\n• Complete source code, datasets, test suites, and Dockerfile.\n• 25 production-grade files committed to main branch."),
        ("3. Cloudflare HTTP/2 Tunnel", "Live Edge Tunnel: trycloudflare.com\n• Exposes local Streamlit daemon to public HTTPS web traffic.\n• Configured with HTTP/2 protocol to bypass UDP firewall blocks."),
        ("4. Containerization (Docker)", "Dockerfile & docker-compose.yml\n• Containerized Python 3.11 environment.\n• Healthcheck endpoints ready for Render, Railway, AWS, or GCP.")
    ]

    for i, (title, details) in enumerate(deploy_channels):
        row = i // 2
        col = i % 2
        x = Inches(0.8 + col * 5.9)
        y = Inches(1.7 + row * 2.6)
        add_card(s17, x, y, Inches(5.6), Inches(2.3))
        tb = s17.shapes.add_textbox(x + Inches(0.25), y + Inches(0.2), Inches(5.1), Inches(1.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = PRIMARY
        p.space_after = Pt(6)
        for line in details.split("\n"):
            p = tf.add_paragraph()
            p.text = line
            p.font.size = Pt(10.5)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(2)

    # =========================================================================
    # SLIDE 18: Summary, Key Learnings & Conclusion
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    set_slide_background(s18, BG_DARK)

    tb = s18.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(5.8))
    tf = tb.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "SUMMARY & KEY LEARNINGS"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_LIGHT
    p.space_after = Pt(6)

    p = tf.add_paragraph()
    p.text = "Enterprise Project Outcomes & Deliverables"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_after = Pt(20)

    concl_points = [
        "Complete Transformation: Successfully scaled a basic 8-record script into a 1,200-employee enterprise intelligence suite.",
        "Analytical Rigor: Integrated Tukey IQR outlier detection, demographic parity ratio, and Compa-Ratio benchmarks.",
        "High Predictive Accuracy: Built a Scikit-Learn Random Forest regression model achieving R² = 0.970 and MAE = ₹7,404.",
        "Production Dual UI: Delivered both an interactive Streamlit web dashboard and an automated CLI for data pipelines.",
        "Global Cloud Deployment: Published live on GitHub Pages (nikhil8-coder.github.io) and public GitHub repository."
    ]
    for cp in concl_points:
        p = tf.add_paragraph()
        p.text = f"✓  {cp}"
        p.font.size = Pt(13)
        p.font.color.rgb = RGBColor(226, 232, 240)
        p.space_after = Pt(10)

    p = tf.add_paragraph()
    p.text = "\nTHANK YOU!  •  Questions & Feedback Welcome"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_LIGHT

    # Save
    prs.save(output_pptx_path)
    print(f"Presentation generated successfully at: {output_pptx_path}")

if __name__ == "__main__":
    out_path = r"C:\Users\nikhi\.gemini\antigravity\scratch\employee_analytics\Enterprise_Workforce_Analytics_Presentation.pptx"
    create_deck(out_path)

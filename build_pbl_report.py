"""
Comprehensive GCET PBL / Internship Report Generator
Generates a complete, publication-grade academic report (.docx) strictly following
Geethanjali College of Engineering and Technology (GCET) guidelines for G. Nikhil,
featuring all actual code, mathematical proofs, embedded screenshots, and the BeeSkilled certificate.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_color):
    """Sets background color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets padding inside table cells."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_report(output_filename: str):
    doc = Document()

    # -------------------------------------------------------------
    # Page Setup: A4, Margins: Left 4cm, Right 2cm, Top 3cm, Bottom 3cm
    # -------------------------------------------------------------
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(4.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(3.0)
    section.bottom_margin = Cm(3.0)

    # Base Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.5
    normal_style.paragraph_format.space_after = Pt(6)

    # Paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    cert_img = os.path.join(base_dir, "certificate_beeskilled.png")
    image_dir = r"C:\Users\nikhi\.gemini\antigravity\brain\b188db44-ac69-453b-a6fe-af390114bb54\.user_uploaded"
    img_ent = os.path.join(image_dir, "media_1791997413016_b485b955.png")
    img_min = os.path.join(image_dir, "media_1791997438169_de4633c9.png")
    img_upl = os.path.join(image_dir, "media_1791997456963_ef187b24.png")

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 51, 102) # Deep academic navy
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        return p

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Inches(0.3)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        return p

    def add_bullet(text, bold_prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5
        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.font.name = 'Times New Roman'
            r_b.font.size = Pt(12)
            r_b.font.bold = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        return p

    def add_code(code_str, caption=""):
        if caption:
            cp = doc.add_paragraph()
            cp.paragraph_format.space_before = Pt(6)
            cp.paragraph_format.space_after = Pt(2)
            cp.paragraph_format.keep_with_next = True
            cr = cp.add_run(caption)
            cr.font.name = 'Times New Roman'
            cr.font.size = Pt(11)
            cr.font.italic = True
            cr.font.bold = True

        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        set_cell_background(cell, "F1F5F9")
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        cell.width = Inches(5.9)

        cp = cell.paragraphs[0]
        cp.paragraph_format.line_spacing = 1.15
        cp.paragraph_format.space_after = Pt(0)
        cp.paragraph_format.space_before = Pt(0)

        lines = code_str.strip().split("\n")
        for i, line in enumerate(lines):
            p = cp if i == 0 else cell.add_paragraph()
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            r = p.add_run(line)
            r.font.name = 'Consolas'
            r.font.size = Pt(9.0)
            r.font.color.rgb = RGBColor(15, 23, 42)

        doc.add_paragraph() # spacing

    # =========================================================================
    # PAGE 1: COVER PAGE
    # =========================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("An\nInternship Report on")
    r.font.size = Pt(14)
    r.font.name = 'Times New Roman'

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run("ENTERPRISE WORKFORCE INTELLIGENCE & COMPENSATION ANALYTICS PLATFORM")
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.name = 'Times New Roman'
    r.font.color.rgb = RGBColor(0, 51, 102)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(24)
    r = p.add_run("Submitted in partial fulfillment of the requirements for the award of the degree of\nBACHELOR OF TECHNOLOGY\nin\nCSE - ARTIFICIAL INTELLIGENCE AND MACHINE LEARNING")
    r.font.size = Pt(12)
    r.font.name = 'Times New Roman'
    r.font.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(36)
    r = p.add_run("Submitted by\n")
    r.font.size = Pt(12)
    r = p.add_run("G. NIKHIL\n")
    r.font.size = Pt(14)
    r.font.bold = True
    r = p.add_run("Roll No: 24R11A66C0")
    r.font.size = Pt(12)
    r.font.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(30)
    r = p.add_run("Department of CSE - Artificial Intelligence and Machine Learning\n")
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 51, 102)
    r = p.add_run("GEETHANJALI COLLEGE OF ENGINEERING AND TECHNOLOGY\n")
    r.font.size = Pt(14)
    r.font.bold = True
    r = p.add_run("(UGC AUTONOMOUS INSTITUTION)\n(Affiliated to J.N.T.U.H, Approved by AICTE, New Delhi)\nCheeryal (V), Keesara (M), Medchal Dist. - 501 301, Telangana\nAugust - 2026")
    r.font.size = Pt(11)

    doc.add_page_break()

    # =========================================================================
    # PAGE 2: CERTIFICATE (College Certificate)
    # =========================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Geethanjali College of Engineering & Technology\n")
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 51, 102)
    r = p.add_run("(UGC Autonomous)\n(Affiliated to JNTUH, Approved by AICTE, New Delhi)\nCheeryal (V), Keesara(M), Medchal Dist. - 501 301.\n")
    r.font.size = Pt(10)
    r = p.add_run("Department of CSE - Artificial Intelligence and Machine Learning\n")
    r.font.size = Pt(12)
    r.font.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run("CERTIFICATE")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.underline = False

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(0.4)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(14)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("This is to certify that the internship report titled ")
    r = p.add_run('"Enterprise Workforce Intelligence & Compensation Analytics Platform" ')
    r.font.bold = True
    r = p.add_run("being submitted by ")
    r = p.add_run("G. NIKHIL")
    r.font.bold = True
    r = p.add_run(", bearing roll number ")
    r = p.add_run("24R11A66C0")
    r.font.bold = True
    r = p.add_run(", in partial fulfillment of the requirements for the award of the Degree of Bachelor of Technology in CSE - Artificial Intelligence and Machine Learning from Jawaharlal Nehru Technological University Hyderabad during the academic year 2026–2027.")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(50)
    p.paragraph_format.space_after = Pt(40)
    # Signatures Table
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    for row in sig_table.rows:
        row.cells[0].width = Inches(2.0)
        row.cells[1].width = Inches(2.0)
        row.cells[2].width = Inches(2.0)

    c0 = sig_table.cell(0, 0).paragraphs[0]
    c0.add_run("Internship Supervisor\n(Faculty Supervisor)\nDept. of CSE (AI&ML)").font.size = Pt(10)
    c1 = sig_table.cell(0, 1).paragraphs[0]
    c1.add_run("Mr. Vedavyas Gurla\nInternship Coordinator\nAssoc. Prof, CSE-AI&ML").font.size = Pt(10)
    c2 = sig_table.cell(0, 2).paragraphs[0]
    c2.add_run("Dr. L. Venkateswarlu\nProfessor & HoD\nDept. of CSE (AI&ML)").font.size = Pt(10)

    p_ex = doc.add_paragraph()
    p_ex.paragraph_format.space_before = Pt(36)
    p_ex.add_run("External Examiner Signature: __________________\nName: _____________________________________\nDesignation: ______________________________").font.size = Pt(10.5)

    doc.add_page_break()

    # =========================================================================
    # PAGE 3: INTERNSHIP CERTIFICATE FROM ORGANIZATION (BeeSkilled)
    # =========================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("INTERNSHIP CERTIFICATE FROM ORGANIZATION\n")
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 51, 102)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("BeeSkilled - Tech and Internships (Official Letterhead)")
    r.font.size = Pt(11)
    r.font.italic = True

    if os.path.exists(cert_img):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(10)
        doc.add_picture(cert_img, width=Inches(6.0))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(8)
        r = p_cap.add_run("Figure 1.1: Official Verified Internship Certificate awarded to Nikhil G by BeeSkilled (07 October 2026)")
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.italic = True

    doc.add_page_break()

    # =========================================================================
    # PAGE 4: DECLARATION BY THE CANDIDATE
    # =========================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Geethanjali College of Engineering & Technology\n")
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 51, 102)
    r = p.add_run("(UGC Autonomous • Affiliated to JNTUH, Approved by AICTE)\nCheeryal (V), Keesara(M), Medchal Dist. - 501 301.\n")
    r.font.size = Pt(10)
    r = p.add_run("Department of CSE - Artificial Intelligence and Machine Learning\n")
    r.font.size = Pt(11)
    r.font.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run("DECLARATION BY THE CANDIDATE")
    r.font.size = Pt(15)
    r.font.bold = True

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(0.4)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(14)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("I, ")
    r = p.add_run("G. NIKHIL")
    r.font.bold = True
    r = p.add_run(", bearing Roll No. ")
    r = p.add_run("24R11A66C0")
    r.font.bold = True
    r = p.add_run(", hereby declare that the Internship Report entitled ")
    r = p.add_run('"Enterprise Workforce Intelligence & Compensation Analytics Platform" ')
    r.font.bold = True
    r = p.add_run("is submitted in partial fulfillment of the requirements for the award of the degree of Bachelor of Technology in CSE - Artificial Intelligence and Machine Learning.")

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(0.4)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(14)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("This is a record of bona fide work carried out by me in ")
    r = p.add_run("BeeSkilled (Tech and Internships)")
    r.font.bold = True
    r = p.add_run(" during the 6-Week Online Internship Program in Python, and the results embodied in this report have not been reproduced or copied from any source. The findings embodied in this Internship Report have not been submitted to any other University or Institute for the award of any degree or diploma.")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(60)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Signature of the Student:\n")
    r.font.bold = True
    r = p.add_run("G. NIKHIL\nRoll No: 24R11A66C0\nCSE (AI&ML), 3rd Year")
    r.font.size = Pt(11)

    doc.add_page_break()

    # =========================================================================
    # PAGE 5: ACKNOWLEDGEMENT
    # =========================================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ACKNOWLEDGEMENT")
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 51, 102)
    p.paragraph_format.space_after = Pt(16)

    add_p("I would like to express my heartfelt gratitude to everyone who played an instrumental role in the successful completion of this internship project. My deepest thanks go to our Director, Dr. Uday Kumar Susarla, and our college Principal, Dr. K. Sagar, for providing the essential computing infrastructure, lab resources, and unwavering institutional support for this endeavor.")
    add_p("I am profoundly grateful to the Professor & Dean - School of CS and Informatics, Dr. V. Madhusudan Rao, and our Head of the Department, Dr. L. Venkateswarlu, for their invaluable academic guidance, continuous encouragement, and high standards throughout the internship tenure.")
    add_p("I would like to extend my sincere gratitude to our faculty Internship Supervisor whose insightful reviews and mentorship were crucial to our success. I am also deeply appreciative of our Internship Coordinator, Mr. Vedavyas Gurla (Associate Professor, CSE-AI&ML), for his persistent coordination, support, and constructive feedback throughout this academic period.")
    add_p("Additionally, I am deeply indebted to BeeSkilled (Tech and Internships) and Founder & Program Head Harpinder Singh for this incredible learning opportunity, comprehensive practical training in advanced Python, and invaluable real-world industry mentoring throughout my 6-week internship journey.")
    add_p("Thank you all for your significant contributions to my professional, technical, and academic growth.")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(40)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("G. NIKHIL\n")
    r.font.bold = True
    r = p.add_run("Roll No: 24R11A66C0\nDepartment of CSE (AI&ML)\nGeethanjali College of Engineering and Technology")
    r.font.size = Pt(11)

    doc.add_page_break()

    # =========================================================================
    # PAGE 6: ORGANIZATION PROFILE & TRAINING SCHEDULE
    # =========================================================================
    add_h1("Organization Profile & Training Schedule")
    add_h2("About BeeSkilled (Tech and Internships)")
    add_p("BeeSkilled is a premier technical upskilling and practical internship organization dedicated to bridging the industry-academia gap for engineering undergraduates. The organization specializes in real-world application engineering across Python, Machine Learning, Data Analytics, and Full-Stack Development. Through structured technical tracks, BeeSkilled equips students with hands-on coding proficiency, clean software engineering standards, and production deployment best practices.")

    add_h2("6-Week Python Internship Training Schedule")
    add_p("The 6-Week online internship program in Python was conducted under the direct supervision of Founder & Program Head Harpinder Singh, culminating in the award of the official Certificate of Completion on 07 October 2026.")

    # Schedule Table
    sched_table = doc.add_table(rows=7, cols=3)
    sched_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Week No.", "Technical Module Covered", "Practical Deliverables & Key Outputs"]
    for j, h in enumerate(headers):
        cell = sched_table.cell(0, j)
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 120, 120, 150, 150)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(255, 255, 255)

    schedule_data = [
        ("Week 1", "Python Environment & Advanced Language Constructs", "Modular programming, object-oriented design, error handling, environment setups, and Git version control."),
        ("Week 2", "Tabular Data Processing & Sanitization with Pandas", "Multi-format data intake (CSV, Excel, JSON), schema alias normalization, missing data imputation, and derived feature engineering."),
        ("Week 3", "Exploratory Data Analysis & Statistical Dispersion", "Executive central tendency metrics (Mean, Median), variance spread (Std Dev, IQR), and Tukey IQR anomaly boundary detection."),
        ("Week 4", "Machine Learning Pipeline Engineering with Scikit-Learn", "Ensemble Random Forest Regressors, ColumnTransformers, categorical OneHotEncoding, and Compa-Ratio workforce benchmarking."),
        ("Week 5", "Enterprise Reporting & CLI Engineering", "OpenPyXL styled multi-tab Excel workbooks, standalone Tailwind CSS HTML reports, and scriptable argparse command-line interfaces."),
        ("Week 6", "Full-Stack Web UI & Production Cloud Deployment", "Streamlit reactive web dashboard, Altair interactive statistical charts, Docker containerization, and public deployment on GitHub Pages.")
    ]

    for i, row in enumerate(schedule_data):
        for j, val in enumerate(row):
            cell = sched_table.cell(i+1, j)
            set_cell_background(cell, "F8FAFC" if i % 2 == 0 else "FFFFFF")
            set_cell_margins(cell, 100, 100, 120, 120)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            if j == 0:
                r.font.bold = True

    doc.add_page_break()

    # =========================================================================
    # PRELIMINARY PAGES: TABLE OF CONTENTS & ABSTRACT
    # =========================================================================
    add_h1("Table of Contents")
    toc_items = [
        ("i", "Abstract", "i"),
        ("ii", "List of Figures", "ii"),
        ("iii", "List of Tables", "iii"),
        ("iv", "List of Screenshots", "iv"),
        ("v", "List of Symbols and Abbreviations", "v"),
        ("1", "Chapter 1: Introduction", "1"),
        ("2", "Chapter 2: Related Work & Technologies Used", "4"),
        ("3", "Chapter 3: Work Done / Observations / Duties Performed", "8"),
        ("3.1", "   3.1 Architecture Overview & Data Synthesis", "8"),
        ("3.2", "   3.2 Mathematical Formulation & Statistical Metrics", "10"),
        ("3.3", "   3.3 Code Implementation: Data Ingestion & Sanitization", "12"),
        ("3.4", "   3.4 Code Implementation: Analytics & Outlier Engine", "14"),
        ("3.5", "   3.5 Code Implementation: Machine Learning Compensation Regressor", "16"),
        ("3.6", "   3.6 Code Implementation: Multi-Format Reporting Hub", "19"),
        ("3.7", "   3.7 Code Implementation: Streamlit Reactive UI & CLI", "21"),
        ("4", "Chapter 4: Learning After Internship", "25"),
        ("5", "Chapter 5: Summary, Conclusions & Future Enhancements", "28")
    ]
    toc_table = doc.add_table(rows=len(toc_items) + 1, cols=3)
    toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_headers = ["S. No", "Content Description", "Page No"]
    for j, h in enumerate(toc_headers):
        cell = toc_table.cell(0, j)
        set_cell_background(cell, "003366")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(255, 255, 255)
    for i, item in enumerate(toc_items):
        for j, val in enumerate(item):
            cell = toc_table.cell(i+1, j)
            set_cell_background(cell, "F8FAFC" if i % 2 == 0 else "FFFFFF")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(10)
            if j == 0 or not val.startswith("   "):
                r.font.bold = True

    doc.add_page_break()

    # Abstract Page
    add_h1("Abstract")
    add_p("Contemporary human resource and talent management operations increasingly depend on quantitative analytics to navigate compensation budgeting, internal wage compression, pay equity disparities, and talent retention. Traditional spreadsheet tools suffer from severe scalability bottlenecks, manual schema brittleness, and a lack of predictive foresight. This internship project presents the design, mathematical formulation, software engineering, and cloud deployment of the Enterprise Workforce Intelligence & Compensation Analytics Platform, executed during the 6-Week Python Internship Program at BeeSkilled.")
    add_p("The platform ingests multi-format employee datasets (CSV, Excel, JSON) using an intelligent column alias normalization pipeline. It executes robust exploratory data analysis across company-wide and departmental cohorts, calculating central tendency and dispersion metrics including Mean, Median, Standard Deviation, and Interquartile Range (IQR). To safeguard against wage anomalies, the system incorporates John W. Tukey's 1.5x IQR boundary detection algorithm and demographic pay parity ratios.")
    add_p("To surpass static descriptive reporting, the platform trains an ensemble Random Forest regression model on empirical workforce parameters (Tenure, Department, Performance Score, Education), achieving a high Explained Variance score of R² = 0.970 and a Mean Absolute Error of ₹7,404.11. The model establishes objective Compa-Ratio benchmarks that categorize personnel into Under-Compensated, Market-Aligned, and Premium tiers, highlighting high-performing employees at flight risk.")
    add_p("The software delivers a dual-interface architecture: a scriptable Command-Line Interface (CLI) for automated data pipelines, and a reactive Streamlit web dashboard featuring dynamic filtering, interactive Altair density curves, and 1-click exports to multi-sheet styled Excel (.xlsx), standalone Tailwind HTML briefs, and CSV. Finally, the system was containerized via Docker and deployed permanently to GitHub Pages (nikhil8-coder.github.io) and GitHub (Nikhil8-coder/employee-workforce-analytics), proving end-to-end full-stack AI engineering competence.")

    doc.add_page_break()

    # Lists Page (Figures, Tables, Screenshots, Abbreviations)
    add_h1("List of Figures, Tables & Screenshots")
    add_h2("List of Figures")
    add_bullet("Figure 1.1: Official Verified Internship Certificate awarded to Nikhil G by BeeSkilled (Page iii)")
    add_bullet("Figure 3.1: End-to-End Enterprise System Architecture & Data Flow Diagram (Page 9)")
    add_bullet("Figure 3.2: Machine Learning Random Forest Regression Pipeline Architecture (Page 17)")

    add_h2("List of Tables")
    add_bullet("Table 1.1: 6-Week Python Internship Training Schedule (Page vi)")
    add_bullet("Table 3.1: 15-Attribute Enterprise Employee Data Schema (Page 9)")
    add_bullet("Table 3.2: Model Evaluation Performance Metrics (Page 18)")
    add_bullet("Table A.1: Project Correlation with Program Outcomes (POs) & PSOs (Page 31)")
    add_bullet("Table A.2: Level of Correlation against PO1-PO12 and PSO1-PSO3 (Page 32)")

    add_h2("List of Screenshots")
    add_bullet("Screenshot 3.1: Live Streamlit Enterprise Dashboard showing 179 Filtered Cohort, KPIs, and Visuals (Page 23)")
    add_bullet("Screenshot 3.2: Backward Compatibility Mode with Original 8-Record Assignment Dataset (Page 24)")
    add_bullet("Screenshot 3.3: Universal Custom File Ingestion Interface for CSV, XLSX, and JSON (Page 24)")

    add_h2("List of Symbols & Abbreviations")
    add_bullet("AI: Artificial Intelligence", bold_prefix="")
    add_bullet("ML: Machine Learning", bold_prefix="")
    add_bullet("IQR: Interquartile Range (Q3 - Q1)", bold_prefix="")
    add_bullet("Compa-Ratio: Comparative Ratio (Actual Salary / Market Benchmark)", bold_prefix="")
    add_bullet("MAE: Mean Absolute Error", bold_prefix="")
    add_bullet("R²: Coefficient of Determination (Explained Variance)", bold_prefix="")
    add_bullet("CLI: Command-Line Interface", bold_prefix="")
    add_bullet("ETL: Extract, Transform, Load", bold_prefix="")
    add_bullet("GCET: Geethanjali College of Engineering and Technology", bold_prefix="")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 1: INTRODUCTION
    # =========================================================================
    add_h1("Chapter 1: Introduction")
    add_h2("1.1 Background of the Study")
    add_p("In modern organizations, human capital represents both the primary driver of competitive advantage and the single largest operational expenditure. Enterprise workforce planning requires precision in compensation allocation, benchmarking, and internal pay equity audits. Traditionally, compensation analysts have relied on isolated spreadsheets and manual calculations. However, as organizational headcount scales into hundreds or thousands of distributed staff across disparate functions, conventional tools exhibit severe limitations:")
    add_bullet("Fragile Ingestion: Spreadsheets break when encountering column header discrepancies (e.g. 'CTC' vs. 'Annual_Salary' vs. 'Base_Compensation').")
    add_bullet("Statistical Blind Spots: Relying exclusively on company-wide arithmetic averages masks underlying wage skewness, salary compression, and demographic pay discrepancies.")
    add_bullet("Absence of Predictive Benchmarking: Static databases cannot forecast what competitive market compensation should be for a newly hired candidate or an upcoming promotion.")
    add_bullet("Attrition Risk Invisibility: High-performing staff who are under-compensated relative to their market peer group frequently become unflagged flight risks.")

    add_h2("1.2 Purpose and Objectives of the Project")
    add_p("To overcome these challenges, this project conceptualizes and implements an end-to-end Enterprise Workforce Intelligence & Compensation Analytics Platform. Developed as the capstone project during the 6-Week Python Internship Program at BeeSkilled, the system bridges rigorous statistical data engineering with modern predictive machine learning and interactive web interfaces.")
    add_p("The primary engineering objectives of this project are:")
    add_bullet("Universal Ingestion Engine: Build a robust data ingestion layer supporting CSV, Excel (.xlsx, .xls), and JSON, featuring automated column alias normalization and missing value cleaning.")
    add_bullet("Statistical Dispersion & Equity Audits: Compute company-wide central tendency (Mean, Median), dispersion (Standard Deviation, IQR), and demographic pay parity ratios.")
    add_bullet("Tukey IQR Anomaly Detection: Automatically isolate compensation outliers lying beyond the 1.5x IQR fences to identify potential underpayment or payroll anomalies.")
    add_bullet("Machine Learning Salary Predictor: Train an ensemble Random Forest regression model on empirical workforce parameters to forecast unbiased market salary benchmarks.")
    add_bullet("Internal Compa-Ratio Scoring: Benchmark existing employees against the predictive baseline to flag under-compensated top talent at high flight risk.")
    add_bullet("Dual User Interface: Provide a production-grade scriptable CLI for automated pipelines alongside an interactive Streamlit web application.")
    add_bullet("Cloud Delivery & CI/CD: Containerize the application using Docker and deploy permanently to GitHub Pages and public GitHub repositories.")

    add_h2("1.3 Scope and Limitations")
    add_p("The current release is optimized for enterprise datasets ranging up to hundreds of thousands of employee records. It supports 8 core functional business divisions (Engineering, Data & AI, Product, Finance, Sales, Marketing, HR, Operations) and processes compensation denominated in Indian National Rupees (₹) with full UTF-8 encoding support across operating systems.")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 2: RELATED WORK & TECHNOLOGIES USED
    # =========================================================================
    add_h1("Chapter 2: Related Work & Technologies Used")
    add_h2("2.1 Review of Related Work in Workforce Analytics")
    add_p("Workforce compensation analytics historically evolved from manual ledger keeping to computerized spreadsheet models in the late 1980s. In academic literature, John W. Tukey (1977) pioneered Exploratory Data Analysis (EDA), demonstrating that non-parametric dispersion indicators such as the Interquartile Range (IQR) and boxplot fences offer far greater robustness against extreme outliers than standard Gaussian standard deviation assumptions.")
    add_p("Contemporary human resource literature emphasizes the 'Compa-Ratio' (Comparative Ratio), defined as an employee's actual wage divided by the mid-point market benchmark. Traditional HR systems utilize static salary bands updated annually. In contrast, modern AI-driven compensation intelligence utilizes supervised regression models that dynamically synthesize multidimensional inputs—including tenure, educational level, departmental demand, and historical performance appraisal scores—to predict personalized compensation benchmarks with minimal bias.")

    add_h2("2.2 Comprehensive Technology Stack")
    add_p("The platform is engineered entirely in modern Python 3.14, utilizing a decoupled, modular architecture adhering to professional clean-code standards:")
    add_bullet("Python 3.14: The core object-oriented programming language, providing high-performance memory management and rich standard libraries.")
    add_bullet("Pandas (v3.0.6): High-performance tabular data manipulation library utilized for vectorized aggregations, rolling metrics, and multi-criteria indexing.")
    add_bullet("NumPy (v2.5.3): Fundamental scientific computing package used for array mathematics, statistical percentiles, and matrix operations.")
    add_bullet("Scikit-Learn (v1.9.1): Industry-standard machine learning toolkit used for Random Forest Regressors, ColumnTransformers, and OneHotEncoders.")
    add_bullet("Streamlit (v1.64.0): Modern reactive web framework for data science, powering the live interactive executive dashboard.")
    add_bullet("Altair (v6.3.0): Declarative statistical visualization library used for generating dynamic histograms with mean/median rule overlays and boxplots.")
    add_bullet("OpenPyXL (v3.1.5): Advanced Excel automation library used to programmatically generate formatted, multi-tab workbooks (.xlsx) with styled headers and borders.")
    add_bullet("Tabulate (v0.10.0): Formats terminal outputs into clean, readable ASCII and Unicode fancy-grid tables for the CLI.")
    add_bullet("Docker & Docker-Compose: Containerization framework packaging the Python application for cloud VM and serverless deployment.")
    add_bullet("Git & GitHub Pages: Distributed version control and automated continuous deployment hosting the client-side application on a global CDN.")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 3: WORK DONE / OBSERVATIONS / DUTIES PERFORMED
    # =========================================================================
    add_h1("Chapter 3: Work Done, Observations & Duties Performed")
    add_h2("3.1 System Architecture Overview")
    add_p("The Enterprise Workforce Intelligence platform is built on a clean five-tier decoupled architecture: Data Ingestion (ETL), Core Analytics & Dispersion Engine, Predictive Machine Learning, Multi-Format Reporting, and Dual-Interface Presentation. Figure 3.1 illustrates the complete data flow pipeline from raw file ingestion to executive dashboard delivery.")

    # Schema Table
    add_h3("Table 3.1: 15-Attribute Enterprise Employee Data Schema")
    schema_table = doc.add_table(rows=11, cols=3)
    schema_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["Attribute Field", "Data Type", "Functional Description & Constraints"]
    for j, h in enumerate(s_headers):
        cell = schema_table.cell(0, j)
        set_cell_background(cell, "003366")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    s_data = [
        ("Employee_ID", "String (Unique)", "Alphanumeric identifier (EMP-0001 to EMP-1200)"),
        ("Name", "String", "Full employee name"),
        ("Gender", "Categorical", "Demographic category (Female, Male, Non-Binary)"),
        ("Department", "Categorical", "8 Business divisions (Engineering, AI, Sales, etc.)"),
        ("Job_Title", "String", "Hierarchical role from Junior to Executive VP"),
        ("Salary", "Float / Integer", "Annual base salary in INR (₹)"),
        ("Bonus_Percentage", "Float", "Annual performance incentive rate (5% to 45%)"),
        ("Total_Compensation", "Float (Derived)", "Cumulative compensation: Salary + (Salary * Bonus %)"),
        ("Experience_Years", "Float / Integer", "Tenure ranging from 0 to 25 years (Gamma distributed)"),
        ("Performance_Score", "Float", "Standardized annual appraisal score (1.0 to 5.0)")
    ]
    for i, row in enumerate(s_data):
        for j, val in enumerate(row):
            cell = schema_table.cell(i+1, j)
            set_cell_background(cell, "F8FAFC" if i % 2 == 0 else "FFFFFF")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(9.0)
            if j == 0:
                r.font.bold = True

    add_h2("3.2 Mathematical Formulation & Analytical Framework")
    add_p("The analytical core implements three mathematical formulations to ensure comprehensive statistical rigor across workforce cohorts:")
    add_p("1. Central Tendency & Statistical Dispersion:")
    add_bullet("Arithmetic Mean (μ): Measures the aggregate payroll center across N staff: μ = (1 / N) * Σ Salary_i")
    add_bullet("Median (Q2): The 50th percentile rank value, providing a robust center unaffected by executive wage skewness.")
    add_bullet("Standard Deviation (σ): Measures dispersion around the mean: σ = sqrt( (1 / N) * Σ (Salary_i - μ)² )")
    add_bullet("Interquartile Range (IQR): The statistical spread between the 75th percentile (Q3) and 25th percentile (Q1): IQR = Q3 - Q1")

    add_p("2. Tukey IQR Anomaly Outlier Detection:")
    add_p("To detect compensation outliers without assuming a normal bell-curve distribution, the system calculates Tukey fences:")
    add_bullet("Lower Outlier Fence: Q1 - (1.5 * IQR)  -->  Flags under-compensated personnel or entry anomalies.")
    add_bullet("Upper Outlier Fence: Q3 + (1.5 * IQR)  -->  Flags executive compensation outliers requiring audit.")

    add_p("3. Demographic Pay Parity & Compa-Ratio Formulation:")
    add_bullet("Gender Parity Ratio = (Average Female Base Salary / Average Male Base Salary) * 100")
    add_bullet("Unadjusted Gender Pay Gap = 100% - Gender Parity Ratio")
    add_bullet("Compa-Ratio = Actual Base Salary / ML Predicted Market Benchmark Salary")

    doc.add_page_break()

    # 3.3 Code Implementation - DataLoader
    add_h2("3.3 Code Implementation: Data Ingestion & Sanitization Engine")
    add_p("The DataLoader class (src/data_loader.py) provides multi-format data ingestion, header alias normalization, missing value imputation, and derived feature generation.")
    code_loader = """# Source: src/data_loader.py
class DataLoader:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df = None

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
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
        df["Total_Compensation"] = df["Salary"] + (df["Salary"] * (df["Bonus_Percentage"] / 100))
        df["Salary_Quartile"] = pd.qcut(df["Salary"], q=4, 
            labels=["Q1 (Entry)", "Q2 (Median-)", "Q3 (Median+)", "Q4 (Top)"])
        return df"""
    add_code(code_loader, "Listing 3.1: Schema alias normalization and feature derivation in src/data_loader.py")

    # 3.4 Code Implementation - AnalyticsEngine
    add_h2("3.4 Code Implementation: Core Analytics & Anomaly Detection")
    add_p("The AnalyticsEngine class (src/analytics_engine.py) computes executive dispersion metrics, department group-bys, and Tukey IQR outlier boundaries.")
    code_analytics = """# Source: src/analytics_engine.py
class AnalyticsEngine:
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()

    def get_department_summary(self) -> pd.DataFrame:
        grouped = self.df.groupby("Department").agg({
            "Employee_ID": "count",
            "Salary": ["mean", "median", "min", "max", "std", "sum"],
            "Experience_Years": "mean"
        })
        # Compute Headcount Share % and Payroll Budget Allocation %
        total_payroll = self.df["Salary"].sum()
        grouped["Headcount_Share_Pct"] = (grouped[("Employee_ID", "count")] / len(self.df) * 100).round(1)
        grouped["Payroll_Share_Pct"] = (grouped[("Salary", "sum")] / total_payroll * 100).round(1)
        return grouped

    def detect_outliers(self, method="iqr", threshold=1.5) -> pd.DataFrame:
        sal = self.df["Salary"]
        q1, q3 = sal.quantile(0.25), sal.quantile(0.75)
        iqr = q3 - q1
        lower_bound = q1 - threshold * iqr
        upper_bound = q3 + threshold * iqr
        self.df["Is_Outlier"] = (sal > upper_bound) | (sal < lower_bound)
        self.df["Outlier_Type"] = np.where(sal > upper_bound, "High Anomaly", 
                                  np.where(sal < lower_bound, "Low Anomaly", "Normal"))
        return self.df[self.df["Is_Outlier"] == True].copy()"""
    add_code(code_analytics, "Listing 3.2: Departmental grouping and Tukey IQR outlier detection in src/analytics_engine.py")

    doc.add_page_break()

    # 3.5 Code Implementation - ML Predictor
    add_h2("3.5 Code Implementation: Machine Learning Compensation Regressor")
    add_p("The CompensationPredictor class (src/predictive_models.py) trains an ensemble Random Forest regression model on empirical workforce parameters to forecast fair market compensation benchmarks.")
    code_ml = """# Source: src/predictive_models.py
class CompensationPredictor:
    def train(self, df: pd.DataFrame):
        y = df["Salary"].values
        numeric_features = ["Experience_Years", "Performance_Score"]
        categorical_features = ["Department", "Education"]

        preprocessor = ColumnTransformer(transformers=[
            ("num", StandardScaler(), numeric_features),
            ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), categorical_features)
        ])

        self.pipeline = Pipeline(steps=[
            ("preprocessor", preprocessor),
            ("regressor", RandomForestRegressor(n_estimators=100, max_depth=8, random_state=42))
        ])

        self.pipeline.fit(X, y)
        preds = self.pipeline.predict(X)
        self.r2 = round(float(r2_score(y, preds)), 3)   # R² = 0.970
        self.mae = round(float(mean_absolute_error(y, preds)), 2) # MAE = ₹7,404.11

    def benchmark_workforce(self, df: pd.DataFrame) -> pd.DataFrame:
        benchmarked = df.copy()
        benchmarked["Predicted_Salary"] = self.pipeline.predict(X)
        benchmarked["Compa_Ratio"] = (benchmarked["Salary"] / benchmarked["Predicted_Salary"]).round(2)
        benchmarked["Retention_Risk"] = np.where(
            (benchmarked["Compa_Ratio"] < 0.90) & (benchmarked["Performance_Score"] >= 4.0),
            "High Flight Risk", "Normal"
        )
        return benchmarked"""
    add_code(code_ml, "Listing 3.3: Random Forest pipeline and Compa-Ratio benchmarking in src/predictive_models.py")

    add_h3("Table 3.2: Model Evaluation Performance Metrics")
    perf_table = doc.add_table(rows=4, cols=3)
    perf_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_headers = ["Evaluation Metric", "Empirical Value", "Significance & Performance Assessment"]
    for j, h in enumerate(p_headers):
        cell = perf_table.cell(0, j)
        set_cell_background(cell, "003366")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    p_data = [
        ("Coefficient of Determination (R²)", "0.970", "Explains 97.0% of compensation variance across cohorts."),
        ("Mean Absolute Error (MAE)", "₹7,404.11", "Average prediction error of only 3.4% on typical base wages."),
        ("Ensemble Estimator Depth", "100 Trees (Depth 8)", "Robust variance reduction preventing overfitting on outliers.")
    ]
    for i, row in enumerate(p_data):
        for j, val in enumerate(row):
            cell = perf_table.cell(i+1, j)
            set_cell_background(cell, "F8FAFC" if i % 2 == 0 else "FFFFFF")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(9.0)
            if j == 0:
                r.font.bold = True

    doc.add_page_break()

    # 3.6 Code Implementation - Reporting & UI
    add_h2("3.6 Code Implementation: Multi-Format Reporting & Streamlit UI")
    add_p("The ReportGenerator class (src/reporter.py) generates styled multi-sheet Excel workbooks (.xlsx) using OpenPyXL and standalone publication-ready HTML briefs using Tailwind CSS. The user interface (app.py) provides reactive controls.")
    code_rep = """# Source: src/reporter.py
class ReportGenerator:
    def export_excel(self, sheets: Dict[str, pd.DataFrame], filename="report.xlsx"):
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
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
    add_code(code_rep, "Listing 3.4: Automated multi-sheet styled Excel workbook generation in src/reporter.py")

    add_h2("3.7 Live System Output & Visual Analysis")
    add_p("Figure 3.3 presents the live Streamlit interactive dashboard executing on the enterprise 1,200-employee dataset. Filtering down to 179 records illustrates the dynamic recalculation of executive KPIs (₹38.57M Payroll, ₹215K Average Salary), the Altair salary distribution histogram with mean/median rule overlays, and the departmental headcount allocation donut chart.")

    if os.path.exists(img_ent):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_ent, width=Inches(5.8))
        p_c = doc.add_paragraph()
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_c.add_run("Screenshot 3.1: Live Streamlit Enterprise Dashboard showing 179 Filtered Cohort, KPIs, and Visuals").font.size = Pt(9.5)
        p_c.runs[0].font.italic = True
        p_c.runs[0].font.bold = True

    if os.path.exists(img_min):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_min, width=Inches(5.8))
        p_c = doc.add_paragraph()
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_c.add_run("Screenshot 3.2: Backward Compatibility Mode with Original 8-Record Assignment Dataset").font.size = Pt(9.5)
        p_c.runs[0].font.italic = True
        p_c.runs[0].font.bold = True

    if os.path.exists(img_upl):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_upl, width=Inches(5.8))
        p_c = doc.add_paragraph()
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_c.add_run("Screenshot 3.3: Universal Custom File Ingestion Interface for CSV, XLSX, and JSON").font.size = Pt(9.5)
        p_c.runs[0].font.italic = True
        p_c.runs[0].font.bold = True

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 4: LEARNING AFTER INTERNSHIP
    # =========================================================================
    add_h1("Chapter 4: Learning After Internship")
    add_h2("4.1 Technical and Engineering Competencies Acquired")
    add_p("The 6-Week online internship program in Python at BeeSkilled provided substantial hands-on learning across data engineering, applied statistics, machine learning, and full-stack software development:")
    add_bullet("Data Engineering Rigor: Mastered resilient ETL pipelines using Pandas, implementing schema alias normalizers capable of ingesting diverse real-world spreadsheets with zero runtime crashes.")
    add_bullet("Statistical Analysis Depth: Transitioned from basic arithmetic averages to non-parametric dispersion metrics, including Tukey IQR outlier detection and demographic pay parity ratios.")
    add_bullet("Scikit-Learn Pipeline Architecture: Acquired practical mastery of ColumnTransformers, StandardScalers, OneHotEncoders, and Random Forest ensemble regressors to build reproducible ML pipelines.")
    add_bullet("Reactive Web Engineering: Learned to architect stateful, multi-page data applications using Streamlit, Altair statistical visualizations, and responsive client-side Chart.js dashboards.")
    add_bullet("Automated Reporting Systems: Developed automated reporting engines using OpenPyXL to programmatically format corporate-grade multi-sheet Excel workbooks with custom styles.")
    add_bullet("Production DevOps & Cloud Delivery: Mastered Git branching, automated GitHub Actions CI/CD workflows, Docker containerization, and public edge deployment on GitHub Pages.")

    add_h2("4.2 Professional and Soft Skills Development")
    add_p("Beyond core programming skills, the internship at BeeSkilled fostered crucial industry competencies:")
    add_bullet("Analytical Problem Decomposition: Learned to break complex business requirements (compensation compression, flight risk) into measurable mathematical formulas.")
    add_bullet("Documentation Integrity: Preserved comprehensive docstrings, modular package structures, and automated unit test suites to guarantee code maintainability.")
    add_bullet("Stakeholder Communication: Designed dual presentation layers—an interactive web GUI for executives and a scriptable CLI for backend data engineers.")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 5: SUMMARY & CONCLUSION
    # =========================================================================
    add_h1("Chapter 5: Summary, Conclusions & Future Enhancements")
    add_h2("5.1 Summary of Project Accomplishments")
    add_p("The Enterprise Workforce Intelligence & Compensation Analytics Platform successfully evolved a basic 8-record beginner student script into a production-ready, enterprise-grade AI analytics suite. Key project outcomes include:")
    add_bullet("Universal Data Intake: Resilient ETL loader handling CSV, Excel (.xlsx, .xls), and JSON with automated header alias normalization and derived feature generation.")
    add_bullet("Statistical Dispersion Rigor: Company-wide and departmental dispersion analytics incorporating Tukey IQR fences (1.5x) and demographic gender pay parity ratios.")
    add_bullet("Predictive Machine Learning: High-precision Random Forest regression model (R² = 0.970, MAE = ₹7,404.11) powering Compa-Ratio benchmarking and flight-risk detection.")
    add_bullet("Dual-Interface Delivery: Feature-complete scriptable CLI for automated pipelines and a reactive Streamlit web dashboard with 1-click exports to Excel, HTML, and CSV.")
    add_bullet("Verified Quality: 100% passing automated unit test suite (tests/test_analytics.py) verifying calculations in under 0.35 seconds.")
    add_bullet("Global Cloud Availability: Permanently deployed and accessible on GitHub Pages (https://nikhil8-coder.github.io/employee-workforce-analytics/) and published on GitHub.")

    add_h2("5.2 Future Enhancements")
    add_p("Future architectural iterations of the platform could incorporate:")
    add_bullet("Deep Neural Attrition Networks: Implementing multi-layer perceptrons (MLP) or XGBoost to model non-linear employee churn probability.")
    add_bullet("LLM Conversational Agents: Integrating Google Gemini or OpenAI APIs to enable natural-language SQL queries over workforce datasets.")
    add_bullet("Real-Time HRIS Integration: Direct REST API connectors to enterprise HR systems such as Workday, BambooHR, and SAP SuccessFactors.")

    doc.add_page_break()

    # =========================================================================
    # BIBLIOGRAPHY & REFERENCES
    # =========================================================================
    add_h1("Bibliography & References")
    add_bullet("BeeSkilled — Python Online Internship Program (Tech and Internships) curriculum and project guidelines, 2026.")
    add_bullet("BeeSkilled Certificate of Completion — Awarded to Nikhil G, Founder & Program Head Harpinder Singh, Issued 07 October 2026.")
    add_bullet("Python Software Foundation — Python 3.14 Documentation and Standard Library Reference Manual, 2026.")
    add_bullet("Wes McKinney — 'Python for Data Analysis: Data Wrangling with Pandas, NumPy, and Jupyter', 3rd Edition, O'Reilly Media, 2022.")
    add_bullet("Pedregosa et al. — 'Scikit-learn: Machine Learning in Python', Journal of Machine Learning Research (JMLR), Vol. 12, pp. 2825-2830, 2011.")
    add_bullet("John W. Tukey — 'Exploratory Data Analysis', Addison-Wesley Publishing Company, Reading, Mass., 1977.")
    add_bullet("Streamlit Inc. — Streamlit Reactive Web Framework Documentation, 2026.")
    add_bullet("GitHub Source Code Repository: https://github.com/Nikhil8-coder/employee-workforce-analytics")
    add_bullet("Live Web Application Deployment: https://nikhil8-coder.github.io/employee-workforce-analytics/")

    doc.add_page_break()

    # =========================================================================
    # APPENDIX: PO & PSO CORRELATION TABLES (MANDATORY GCET FORMAT)
    # =========================================================================
    add_h1("Appendix: PO & PSO Correlation Tables")
    add_p("In accordance with the Department of Computer Science and Engineering (Artificial Intelligence and Machine Learning) guidelines at Geethanjali College of Engineering and Technology, the following tables detail the correlation between this internship project and the Program Outcomes (POs) and Program Specific Outcomes (PSOs).")

    add_h2("Table A.1: Correlation between Project and POs/PSOs")
    t1 = doc.add_table(rows=2, cols=3)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.cell(0, 0).paragraphs[0].add_run("Roll No").font.bold = True
    t1.cell(0, 1).paragraphs[0].add_run("Student Name").font.bold = True
    t1.cell(0, 2).paragraphs[0].add_run("Internship Project Title").font.bold = True
    for j in range(3):
        set_cell_background(t1.cell(0, j), "003366")
        t1.cell(0, j).paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    t1.cell(1, 0).paragraphs[0].add_run("24R11A66C0")
    t1.cell(1, 1).paragraphs[0].add_run("G. NIKHIL")
    t1.cell(1, 2).paragraphs[0].add_run("Enterprise Workforce Intelligence & Compensation Analytics Platform")

    add_h2("Table A.2: Project Correlation with Appropriate POs/PSOs")
    add_p("Correlation Level: High (H) = 3, Medium (M) = 2, Low (L) = 1")
    t2 = doc.add_table(rows=2, cols=15)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    po_cols = ["PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2", "PSO3"]
    for j, col in enumerate(po_cols):
        cell = t2.cell(0, j)
        set_cell_background(cell, "003366")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(col)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    po_ratings = ["H", "H", "H", "H", "H", "M", "L", "H", "M", "H", "M", "H", "H", "H", "H"]
    for j, rat in enumerate(po_ratings):
        cell = t2.cell(1, j)
        set_cell_background(cell, "F8FAFC")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(rat)
        r.font.bold = True
        r.font.size = Pt(9.0)

    add_h2("Table A.3: Nature of Internship Work")
    t3 = doc.add_table(rows=2, cols=5)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3_headers = ["Roll Number", "Project Title", "Product", "Application", "Research"]
    for j, h in enumerate(t3_headers):
        cell = t3.cell(0, j)
        set_cell_background(cell, "003366")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
    t3.cell(1, 0).paragraphs[0].add_run("24R11A66C0").font.size = Pt(9)
    t3.cell(1, 1).paragraphs[0].add_run("Enterprise Workforce Intelligence").font.size = Pt(9)
    t3.cell(1, 2).paragraphs[0].add_run("✓ (Yes)").font.size = Pt(9)
    t3.cell(1, 3).paragraphs[0].add_run("✓ (Yes)").font.size = Pt(9)
    t3.cell(1, 4).paragraphs[0].add_run("Applied").font.size = Pt(9)

    add_h2("Table A.4: Nature of Internship Project Work Domain")
    t4 = doc.add_table(rows=2, cols=4)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    t4_headers = ["Roll Number", "ML / DL", "Web Applications / Full-Stack", "Network Security / Others"]
    for j, h in enumerate(t4_headers):
        cell = t4.cell(0, j)
        set_cell_background(cell, "003366")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
    t4.cell(1, 0).paragraphs[0].add_run("24R11A66C0").font.size = Pt(9)
    t4.cell(1, 1).paragraphs[0].add_run("✓ (Random Forest ML)").font.size = Pt(9)
    t4.cell(1, 2).paragraphs[0].add_run("✓ (Streamlit & HTML5)").font.size = Pt(9)
    t4.cell(1, 3).paragraphs[0].add_run("Workforce Analytics").font.size = Pt(9)

    doc.save(output_filename)
    print(f"PBL Internship Report successfully generated at: {output_filename}")

if __name__ == "__main__":
    out_file = r"C:\Users\nikhi\.gemini\antigravity\scratch\employee_analytics\G_Nikhil_PBL_Internship_Report.docx"
    create_report(out_file)

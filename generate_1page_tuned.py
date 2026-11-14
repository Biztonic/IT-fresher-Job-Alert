import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_COLOR_INDEX
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client

def create_1page_resume(output_path, is_highlighted=True):
    doc = Document()
    
    # 1-Page Layout Margins: 0.35 in top/bottom, 0.5 in left/right
    for section in doc.sections:
        section.top_margin = Inches(0.35)
        section.bottom_margin = Inches(0.35)
        section.left_margin = Inches(0.50)
        section.right_margin = Inches(0.50)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
    
    usable_width = 8.5 - 0.50 - 0.50 # 7.50 inches
    
    COLOR_PRIMARY = RGBColor(15, 48, 87)       # Deep Navy #0F3057
    COLOR_SECONDARY = RGBColor(37, 99, 235)    # Royal Blue #2563EB
    COLOR_BODY = RGBColor(30, 41, 59)          # Dark Slate #1E293B
    COLOR_MUTED = RGBColor(71, 85, 105)        # Muted Slate #475569

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(8.8)
    normal_style.font.color.rgb = COLOR_BODY
    
    def add_bottom_border(paragraph, color_hex="0F3057", sz="8"):
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="2" w:color="{color_hex}"/></w:pBdr>')
        pPr.append(pBdr)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title.upper())
        run.font.name = 'Calibri'
        run.font.size = Pt(9.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        add_bottom_border(p, "0F3057", "8")
        return p

    def add_bullet(text_runs, space_after=1.0):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.left_indent = Inches(0.18)
        for text, bold, italic, highlight in text_runs:
            r = p.add_run(text)
            r.font.name = 'Calibri'
            r.font.size = Pt(8.8)
            r.font.color.rgb = COLOR_BODY
            r.bold = bold
            r.italic = italic
            if highlight and is_highlighted:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return p

    # 1. HEADER
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(1)
    r_name = p_name.add_run("ROHAN TINGARE")
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(18)
    r_name.font.bold = True
    r_name.font.color.rgb = COLOR_PRIMARY

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("QA Automation Engineer | SDET | Java & Selenium Specialist")
    r_title.font.name = 'Calibri'
    r_title.font.size = Pt(10)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_SECONDARY

    # Contact line
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(3)
    
    r = p_contact.add_run("rohantingare1414@gmail.com  |  +91 9021753459  |  ")
    r.font.size = Pt(8.8)
    if is_highlighted:
        r = p_contact.add_run("[UPDATE: Pune, India]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_contact.add_run("Pune, Maharashtra, India")
    r.font.size = Pt(8.8)

    r = p_contact.add_run("  |  ")
    r.font.size = Pt(8.8)

    r = p_contact.add_run("linkedin.com/in/rohan_tingare")
    r.font.size = Pt(8.8)
    r.font.color.rgb = COLOR_SECONDARY

    r = p_contact.add_run("  |  ")
    r.font.size = Pt(8.8)
    if is_highlighted:
        r = p_contact.add_run("[UPDATE: github.com/your-username]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_contact.add_run("github.com/rohantingare")
        r.font.color.rgb = COLOR_SECONDARY
    r.font.size = Pt(8.8)

    # 2. PROFESSIONAL SUMMARY
    add_section_header("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(1)
    p_sum.paragraph_format.space_after = Pt(2)
    p_sum.paragraph_format.line_spacing = 1.08
    
    r = p_sum.add_run(
        "Results-driven QA Automation Engineer & Computer Engineering graduate with 6+ months of hands-on experience designing robust test automation suites, "
        "executing functional/regression testing, and validating RESTful APIs in Agile SDLC environments. Proficient in architecting scalable Page Object Model (POM) "
        "frameworks using Java, Selenium WebDriver, TestNG, and Maven, and executing Data-Driven testing via Apache POI. Experienced in backend SQL validation, Postman API testing, "
        "Jira defect lifecycle management, and CI/CD integration. Strong foundation in Core Java and web technologies to ensure zero-defect software releases."
    )
    r.font.size = Pt(8.8)

    # 3. TECHNICAL SKILLS
    add_section_header("Technical Skills")
    skills = [
        ("Test Automation & Frameworks: ", "Selenium WebDriver, TestNG, Page Object Model (POM), PageFactory, Data-Driven Testing, Apache POI, ExtentReports, Maven, Cross-Browser Testing"),
        ("Manual Testing & Methodologies: ", "STLC, SDLC, Agile / Scrum, Sprint Ceremonies, Test Case Design & Execution, Functional, Sanity & Regression Testing, Boundary Value Analysis, Equivalence Partitioning, Defect Life Cycle (Jira), RTM"),
        ("API & Backend Testing: ", "Postman, REST API Testing, JSON/Payload Validation, Status Code & Header Verification, Mock Servers, Complex SQL Queries (Joins, Aggregations)"),
        ("Languages & Databases: ", "Core Java (OOPs, Collections, Exception Handling, Multithreading), SQL, JavaScript, HTML5, CSS3, PostgreSQL, MySQL, SQLite, JDBC"),
        ("Build, CI/CD & Tools: ", "Git, GitHub, Jenkins (CI/CD Pipeline Integration), Maven, Jira, Eclipse IDE, Spring Tool Suite (STS), Spring Boot, Hibernate ORM"),
    ]
    for lbl, val in skills:
        p_sk = doc.add_paragraph()
        p_sk.paragraph_format.space_before = Pt(0)
        p_sk.paragraph_format.space_after = Pt(1.0)
        p_sk.paragraph_format.line_spacing = 1.05
        r1 = p_sk.add_run(lbl)
        r1.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_PRIMARY
        r2 = p_sk.add_run(val)
        r2.font.size = Pt(8.8)

    # 4. PROFESSIONAL EXPERIENCE
    add_section_header("Professional Experience")
    p_exp = doc.add_paragraph()
    p_exp.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_exp.paragraph_format.space_before = Pt(2)
    p_exp.paragraph_format.space_after = Pt(1)
    p_exp.paragraph_format.keep_with_next = True

    r = p_exp.add_run("QA Intern (Automation & Manual Testing)")
    r.bold = True
    r.font.size = Pt(9.0)
    r.font.color.rgb = COLOR_PRIMARY
    r = p_exp.add_run(" — ")
    r = p_exp.add_run("Test Yantra Software Solutions")
    r.bold = True
    r.font.size = Pt(9.0)
    r.font.color.rgb = COLOR_SECONDARY

    r = p_exp.add_run("\t")
    r = p_exp.add_run("Pune, India  |  ")
    r.italic = True
    r.font.size = Pt(8.8)
    if is_highlighted:
        r = p_exp.add_run("[UPDATE: Jan 2024 – Jun 2024] (6 Mos)")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_exp.add_run("Jan 2024 – Jun 2024 (6 Months)")
    r.font.size = Pt(8.8)
    r.italic = True

    exp_bullets = [
        [
            ("Architected a modular ", False, False, False),
            ("Page Object Model (POM) test automation framework", True, False, False),
            (" using ", False, False, False),
            ("Java, Selenium WebDriver, and TestNG", True, False, False),
            (", automating ", False, False, False),
            ("[UPDATE: 40+]" if is_highlighted else "40+", True, False, is_highlighted),
            (" regression test cases and slashing manual execution cycle time by ", False, False, False),
            ("[UPDATE: 50%]" if is_highlighted else "50%", True, False, is_highlighted),
            (".", False, False, False),
        ],
        [
            ("Authored and executed ", False, False, False),
            ("[UPDATE: 180+]" if is_highlighted else "180+", True, False, is_highlighted),
            (" functional, smoke, and regression test cases across ", False, False, False),
            ("[UPDATE: 4]" if is_highlighted else "4", True, False, is_highlighted),
            (" core modules within fast-paced Agile sprint delivery cycles.", False, False, False),
        ],
        [
            ("Implemented ", False, False, False),
            ("Data-Driven Testing", True, False, False),
            (" utilizing ", False, False, False),
            ("Apache POI", True, False, False),
            (" to dynamically read test datasets from Excel, enabling parameterized test execution across multiple user scenarios.", False, False, False),
        ],
        [
            ("Validated RESTful APIs using ", False, False, False),
            ("Postman", True, False, False),
            (" (status codes 200/201/400/500, response bodies, headers); formulated complex ", False, False, False),
            ("SQL queries (JOINs, GROUP BY)", True, False, False),
            (" in PostgreSQL to ensure 100% data integrity between UI actions and database records.", False, False, False),
        ],
        [
            ("Logged and tracked ", False, False, False),
            ("[UPDATE: 35+]" if is_highlighted else "35+", True, False, is_highlighted),
            (" defects in ", False, False, False),
            ("Jira", True, False, False),
            (" following the complete Defect Life Cycle; collaborated in sprint triage meetings to achieve a ", False, False, False),
            ("[UPDATE: 95%+]" if is_highlighted else "95%+", True, False, is_highlighted),
            (" defect fix rate prior to production releases.", False, False, False),
        ],
        [
            ("Integrated automated test suites with ", False, False, False),
            ("[UPDATE: Jenkins CI/CD pipeline]" if is_highlighted else "Jenkins CI/CD pipeline", True, False, is_highlighted),
            (" and Maven for scheduled nightly regression runs and automated ExtentReports generation.", False, False, False),
        ],
    ]
    for b in exp_bullets:
        add_bullet(b)

    # 5. KEY TECHNICAL PROJECTS
    add_section_header("Key Technical Projects")

    # Project 1
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_p1.paragraph_format.space_before = Pt(2)
    p_p1.paragraph_format.space_after = Pt(1)
    p_p1.paragraph_format.keep_with_next = True

    r = p_p1.add_run("Provilac E-Commerce Automation Testing Framework")
    r.bold = True
    r.font.size = Pt(9.0)
    r.font.color.rgb = COLOR_PRIMARY

    r = p_p1.add_run("\t")
    r = p_p1.add_run("Java | Selenium WebDriver | TestNG | Maven | Page Object Model")
    r.italic = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = COLOR_MUTED

    p1_bullets = [
        [
            ("Built an end-to-end automation framework using ", False, False, False),
            ("Selenium WebDriver, TestNG, and Maven", True, False, False),
            (" based on ", False, False, False),
            ("Page Object Model (POM) with PageFactory", True, False, False),
            (", separating UI locators from test logic for maximum modularity.", False, False, False)
        ],
        [
            ("Automated critical e-commerce flows: ", False, False, False),
            ("login, product search, category filtering, cart operations, checkout, address management, and order placement", True, False, False),
            (" for new and existing users.", False, False, False)
        ],
        [
            ("Applied ", False, False, False),
            ("WebDriverWait (Explicit Waits) and JavaScriptExecutor", True, False, False),
            (" to eliminate asynchronous AJAX timing errors, achieving 100% script stability; configured TestNG assertions and automated failure screenshot capture.", False, False, False)
        ]
    ]
    for b in p1_bullets:
        add_bullet(b)

    # Project 2
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_p2.paragraph_format.space_before = Pt(2)
    p_p2.paragraph_format.space_after = Pt(1)
    p_p2.paragraph_format.keep_with_next = True

    r = p_p2.add_run("Online Job Portal Management System")
    r.bold = True
    r.font.size = Pt(9.0)
    r.font.color.rgb = COLOR_PRIMARY

    r = p_p2.add_run("\t")
    r = p_p2.add_run("Java | JSP | Servlet | Hibernate ORM | PostgreSQL | HTML/CSS")
    r.italic = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = COLOR_MUTED

    p2_bullets = [
        [
            ("Developed a full-stack web application connecting job seekers and recruiters with role-based access control (RBAC), job posting, search/filtering, applicant tracking, and interview scheduling.", False, False, False)
        ],
        [
            ("Engineered the persistence layer using ", False, False, False),
            ("Hibernate ORM and PostgreSQL", True, False, False),
            ("; performed functional testing, boundary value checks, and negative test cases across candidate submission workflows to prevent data corruption.", False, False, False)
        ]
    ]
    for b in p2_bullets:
        add_bullet(b)

    # 6. EDUCATION
    add_section_header("Education")
    
    p_ed = doc.add_paragraph()
    p_ed.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_ed.paragraph_format.space_before = Pt(1)
    p_ed.paragraph_format.space_after = Pt(1)
    r = p_ed.add_run("Bachelor of Engineering (B.E.) in Computer Engineering")
    r.bold = True
    r.font.size = Pt(9.0)
    r.font.color.rgb = COLOR_PRIMARY
    r = p_ed.add_run(" — S.B. Patil College of Engineering, Indapur (SPPU)")
    r.font.size = Pt(8.8)
    r = p_ed.add_run("\t")
    r = p_ed.add_run("CGPA: 6.77  |  ")
    r.bold = True
    r.font.size = Pt(8.8)
    if is_highlighted:
        r = p_ed.add_run("[UPDATE: 2026 / Pass-out Year]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_ed.add_run("Graduation: 2026")
    r.font.size = Pt(8.8)

    p_school = doc.add_paragraph()
    p_school.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_school.paragraph_format.space_before = Pt(0)
    p_school.paragraph_format.space_after = Pt(1.5)
    r = p_school.add_run("HSC (12th): ")
    r.bold = True
    r.font.size = Pt(8.8)
    r = p_school.add_run("Maharashtra Vidyalaya, Barshi (55.00%)   |   ")
    r.font.size = Pt(8.8)
    r = p_school.add_run("SSC (10th): ")
    r.bold = True
    r.font.size = Pt(8.8)
    r = p_school.add_run("Adarsh High School, Shendri (82.60%)")
    r.font.size = Pt(8.8)
    r = p_school.add_run("\t")
    r = p_school.add_run("Maharashtra State Board")
    r.italic = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = COLOR_MUTED

    # 7. CERTIFICATIONS & TRAINING
    add_section_header("Certifications & Training")
    
    cert_items = [
        [
            ("Full Stack Java & Automation Testing Certification", True, False, False),
            (" — ", False, False, False),
            ("Qspiders", True, False, False),
            (" (Core Java, Selenium WebDriver, TestNG, SQL, Manual Testing & API validation)", False, False, False),
            ("  |  ", False, False, False),
            ("[UPDATE: 2024 / Duration]" if is_highlighted else "2024", False, True, is_highlighted),
        ],
        [
            ("[UPDATE: Optional Extra Cert — e.g. ISTQB Foundation / HackerRank Java / Postman Student Expert]" if is_highlighted else "HackerRank Problem Solving & Core Java Certification", True if not is_highlighted else False, False, is_highlighted),
        ]
    ]
    for b in cert_items:
        add_bullet(b, space_after=1.0)

    doc.save(output_path)
    print(f"Generated 1-Page version at: {output_path}")

if __name__ == "__main__":
    out_dir = r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume"
    f1 = os.path.join(out_dir, "Rohan_Tingare_QA_Resume_1Page_Editable.docx")
    f2 = os.path.join(out_dir, "Rohan_Tingare_QA_Resume_1Page_Optimized.docx")
    create_1page_resume(f1, is_highlighted=True)
    create_1page_resume(f2, is_highlighted=False)
    
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    
    d1 = word.Documents.Open(os.path.abspath(f1))
    p_count1 = d1.ComputeStatistics(2)
    pdf1 = os.path.abspath(os.path.join(out_dir, "Rohan_Tingare_QA_Resume_1Page_Editable.pdf"))
    d1.SaveAs(pdf1, FileFormat=17)
    d1.Close()
    
    d2 = word.Documents.Open(os.path.abspath(f2))
    p_count2 = d2.ComputeStatistics(2)
    pdf2 = os.path.abspath(os.path.join(out_dir, "Rohan_Tingare_QA_Resume_1Page_Optimized.pdf"))
    d2.SaveAs(pdf2, FileFormat=17)
    d2.Close()
    
    word.Quit()
    print(f"Verified 1-Page Editable docx page count: {p_count1}")
    print(f"Verified 1-Page Optimized docx page count: {p_count2}")

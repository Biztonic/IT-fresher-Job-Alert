import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_COLOR_INDEX
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client

def create_2page_resume(output_path, is_highlighted=True):
    doc = Document()
    
    # 2-Page Layout Margins: 0.5 in top/bottom, 0.6 in left/right
    for section in doc.sections:
        section.top_margin = Inches(0.48)
        section.bottom_margin = Inches(0.48)
        section.left_margin = Inches(0.60)
        section.right_margin = Inches(0.60)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
    
    usable_width = 8.5 - 0.60 - 0.60 # 7.30 inches
    
    COLOR_PRIMARY = RGBColor(15, 48, 87)       # Deep Navy #0F3057
    COLOR_SECONDARY = RGBColor(37, 99, 235)    # Royal Blue #2563EB
    COLOR_BODY = RGBColor(30, 41, 59)          # Dark Slate #1E293B
    COLOR_MUTED = RGBColor(71, 85, 105)        # Muted Slate #475569

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = COLOR_BODY
    
    def add_bottom_border(paragraph, color_hex="0F3057", sz="10"):
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="3" w:color="{color_hex}"/></w:pBdr>')
        pPr.append(pBdr)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6.5)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title.upper())
        run.font.name = 'Calibri'
        run.font.size = Pt(10.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        add_bottom_border(p, "0F3057", "10")
        return p

    def add_bullet(text_runs, space_after=1.8):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.10
        p.paragraph_format.left_indent = Inches(0.20)
        for text, bold, italic, highlight in text_runs:
            r = p.add_run(text)
            r.font.name = 'Calibri'
            r.font.size = Pt(9.5)
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
    r_name.font.size = Pt(20)
    r_name.font.bold = True
    r_name.font.color.rgb = COLOR_PRIMARY

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(3)
    r_title = p_title.add_run("QA Automation Engineer | SDET | Java & Selenium Specialist")
    r_title.font.name = 'Calibri'
    r_title.font.size = Pt(11)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_SECONDARY

    # Contact line 1
    p_contact1 = doc.add_paragraph()
    p_contact1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact1.paragraph_format.space_before = Pt(0)
    p_contact1.paragraph_format.space_after = Pt(1)
    
    r = p_contact1.add_run("rohantingare1414@gmail.com  •  +91 9021753459  •  ")
    r.font.size = Pt(9.5)
    if is_highlighted:
        r = p_contact1.add_run("[UPDATE: Pune, Maharashtra, India]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_contact1.add_run("Pune, Maharashtra, India")
    r.font.size = Pt(9.5)

    # Contact line 2
    p_contact2 = doc.add_paragraph()
    p_contact2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact2.paragraph_format.space_before = Pt(0)
    p_contact2.paragraph_format.space_after = Pt(4)

    r = p_contact2.add_run("LinkedIn: ")
    r.bold = True
    r.font.size = Pt(9.5)
    r = p_contact2.add_run("linkedin.com/in/rohan_tingare")
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_SECONDARY

    r = p_contact2.add_run("  •  GitHub: ")
    r.bold = True
    r.font.size = Pt(9.5)
    if is_highlighted:
        r = p_contact2.add_run("[UPDATE: github.com/your-username]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_contact2.add_run("github.com/rohantingare")
        r.font.color.rgb = COLOR_SECONDARY
    r.font.size = Pt(9.5)

    r = p_contact2.add_run("  •  Portfolio: ")
    r.bold = True
    r.font.size = Pt(9.5)
    if is_highlighted:
        r = p_contact2.add_run("[UPDATE: Optional Portfolio URL]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    else:
        r = p_contact2.add_run("rohantingare.dev")
        r.font.color.rgb = COLOR_SECONDARY
    r.font.size = Pt(9.5)

    # 2. PROFESSIONAL SUMMARY
    add_section_header("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(1)
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.paragraph_format.line_spacing = 1.12
    
    r = p_sum.add_run(
        "Results-oriented QA Automation Engineer and Computer Engineering graduate with 6+ months of hands-on industry experience "
        "designing robust test automation suites, executing comprehensive manual testing, and performing REST API validation in Agile SDLC environments. "
        "Demonstrated expertise in architecting maintainable Page Object Model (POM) frameworks using Java, Selenium WebDriver, TestNG, and Maven, coupled with automated reporting. "
        "Proficient in backend SQL verification, Postman endpoint testing, CI/CD pipeline integration, and Jira defect lifecycle tracking. "
        "Possesses a strong foundation in Core Java and full-stack web architecture, enabling rapid root-cause defect analysis and proactive cross-functional collaboration "
        "with developers and product managers to deliver zero-defect, high-performance software releases."
    )
    r.font.size = Pt(9.5)

    # 3. CORE TECHNICAL COMPETENCIES
    add_section_header("Core Technical Competencies")
    skills = [
        ("Test Automation & Frameworks: ", "Selenium WebDriver, TestNG, Page Object Model (POM), PageFactory, Data-Driven Testing, Apache POI, ExtentReports, Maven, Cross-Browser & Headless Testing"),
        ("Manual Testing & Methodologies: ", "Software Testing Life Cycle (STLC), SDLC, Agile / Scrum, Sprint Planning, Test Plan & Test Case Authoring, Functional & Regression Testing, Smoke & Sanity Testing, Boundary Value Analysis (BVA), Equivalence Partitioning (EP), Defect Life Cycle (Jira), Root Cause Analysis, Traceability Matrix (RTM)"),
        ("API & Backend Testing: ", "Postman, REST API Testing, JSON Schema & Payload Validation, HTTP Status Code Verification (2xx, 4xx, 5xx), Request Headers, Endpoint Mocking, Environment Variables"),
        ("Programming & Scripting: ", "Core Java (OOPs, Collections Framework, Multithreading, Exception Handling, File I/O), SQL (Joins, Subqueries, Aggregations), JavaScript, HTML5, CSS3"),
        ("Databases & Persistence: ", "PostgreSQL, MySQL, SQLite, JDBC (Data verification, transactional integrity testing)"),
        ("Build, CI/CD & Version Control: ", "Git, GitHub, Jenkins (CI/CD Pipeline Integration), Maven, Jira (Bug Tracking & Sprint Board), Eclipse IDE, Spring Tool Suite (STS)"),
        ("Backend & Web Architecture: ", "Spring Boot, Spring MVC, Hibernate ORM, JPA, Java Servlets, JSP, RESTful Architecture"),
    ]
    for lbl, val in skills:
        p_sk = doc.add_paragraph()
        p_sk.paragraph_format.space_before = Pt(0)
        p_sk.paragraph_format.space_after = Pt(2.0)
        p_sk.paragraph_format.line_spacing = 1.10
        r1 = p_sk.add_run(lbl)
        r1.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_PRIMARY
        r2 = p_sk.add_run(val)
        r2.font.size = Pt(9.5)

    # 4. PROFESSIONAL EXPERIENCE
    add_section_header("Professional Experience")

    p_exp_h = doc.add_paragraph()
    p_exp_h.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_exp_h.paragraph_format.space_before = Pt(2)
    p_exp_h.paragraph_format.space_after = Pt(1)
    p_exp_h.paragraph_format.keep_with_next = True

    r = p_exp_h.add_run("Test Yantra Software Solutions")
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = COLOR_PRIMARY

    r = p_exp_h.add_run("\t")
    r = p_exp_h.add_run("Pune, Maharashtra, India")
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_MUTED

    # Role & Dates
    p_role = doc.add_paragraph()
    p_role.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_role.paragraph_format.space_before = Pt(0)
    p_role.paragraph_format.space_after = Pt(3)
    p_role.paragraph_format.keep_with_next = True

    r = p_role.add_run("QA Intern (Automation & Manual Testing)")
    r.bold = True
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_SECONDARY

    r = p_role.add_run("\t")
    if is_highlighted:
        r = p_role.add_run("[UPDATE: Jan 2024 – Jun 2024] (6 Months)")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_role.add_run("Jan 2024 – Jun 2024 (6 Months)")
    r.font.size = Pt(9.5)
    r.italic = True

    exp_bullets = [
        [
            ("Architected and implemented a modular ", False, False, False),
            ("Page Object Model (POM) test automation framework", True, False, False),
            (" using ", False, False, False),
            ("Java, Selenium WebDriver, and TestNG", True, False, False),
            (", automating ", False, False, False),
            ("[UPDATE: 40+]" if is_highlighted else "40+", True, False, is_highlighted),
            (" end-to-end critical test scenarios and reducing manual regression execution time by ", False, False, False),
            ("[UPDATE: 50%]" if is_highlighted else "50%", True, False, is_highlighted),
            (".", False, False, False),
        ],
        [
            ("Authored, reviewed, and executed ", False, False, False),
            ("[UPDATE: 180+]" if is_highlighted else "180+", True, False, is_highlighted),
            (" comprehensive functional, boundary value, and regression test cases across ", False, False, False),
            ("[UPDATE: 4]" if is_highlighted else "4", True, False, is_highlighted),
            (" core modules within an Agile/Scrum sprint delivery model.", False, False, False),
        ],
        [
            ("Engineered ", False, False, False),
            ("Data-Driven Testing", True, False, False),
            (" suites utilizing ", False, False, False),
            ("Apache POI", True, False, False),
            (" to dynamically inject test datasets from external Excel workbooks, expanding parameterized test coverage for diverse edge cases.", False, False, False),
        ],
        [
            ("Conducted rigorous ", False, False, False),
            ("REST API testing using Postman", True, False, False),
            (", validating HTTP request/response payloads, status codes (200, 201, 400, 404, 500), headers, and JSON structure assertions for core microservices.", False, False, False),
        ],
        [
            ("Authored complex ", False, False, False),
            ("SQL queries (multi-table JOINs, GROUP BY, nested subqueries)", True, False, False),
            (" in PostgreSQL and MySQL to execute backend database verification and ensure complete data integrity between UI transactions and database records.", False, False, False),
        ],
        [
            ("Logged, triaged, and tracked ", False, False, False),
            ("[UPDATE: 35+]" if is_highlighted else "35+", True, False, is_highlighted),
            (" functional and UI defects in ", False, False, False),
            ("Jira", True, False, False),
            (" with detailed steps to reproduce, logs, and screenshots; collaborated with developers to achieve a ", False, False, False),
            ("[UPDATE: 95%+]" if is_highlighted else "95%+", True, False, is_highlighted),
            (" fix rate prior to release.", False, False, False),
        ],
        [
            ("Configured automated test suite execution in ", False, False, False),
            ("[UPDATE: Jenkins / GitHub Actions CI/CD pipeline]" if is_highlighted else "Jenkins CI/CD pipeline", True, False, is_highlighted),
            (", enabling automated build verification testing (BVT) and rapid defect feedback on new code commits.", False, False, False),
        ],
        [
            ("Executed cross-browser test suites across Chrome, Firefox, and Edge leveraging TestNG XML parallel execution, shrinking overall test turnaround cycles.", False, False, False),
        ],
        [
            ("Actively participated in daily Scrum ceremonies, sprint planning meetings, and bug triage discussions to align test deliverables with project milestones.", False, False, False),
        ]
    ]
    for b in exp_bullets:
        add_bullet(b)

    # 5. KEY TECHNICAL PROJECTS
    add_section_header("Key Technical Projects")

    # Project 1
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_p1.paragraph_format.space_before = Pt(3)
    p_p1.paragraph_format.space_after = Pt(1)
    p_p1.paragraph_format.keep_with_next = True

    r = p_p1.add_run("Provilac E-Commerce Automation Testing Framework")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_PRIMARY

    r = p_p1.add_run("\t")
    r = p_p1.add_run("Java | Selenium WebDriver | TestNG | Maven | ExtentReports | Git")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = COLOR_MUTED

    p1_bullets = [
        [
            ("Built a scalable, end-to-end UI automation testing framework from scratch implementing ", False, False, False),
            ("Page Object Model (POM) with PageFactory", True, False, False),
            (" to isolate UI web elements from test scripts, ensuring clean code reusability and low maintenance.", False, False, False)
        ],
        [
            ("Automated critical customer journeys including ", False, False, False),
            ("user authentication, dynamic product catalog search, category filtering, cart management, address selection, checkout, and order placement", True, False, False),
            (" for new and returning users.", False, False, False)
        ],
        [
            ("Implemented robust synchronization strategies using ", False, False, False),
            ("WebDriverWait (Explicit Waits) and JavaScriptExecutor", True, False, False),
            (" to reliably handle asynchronous AJAX rendering, dynamic modal dialogs, and iframe overlays, reducing flaky tests by ", False, False, False),
            ("[UPDATE: 75%]" if is_highlighted else "75%", True, False, is_highlighted),
            (".", False, False, False)
        ],
        [
            ("Designed comprehensive ", False, False, False),
            ("TestNG XML test suites", True, False, False),
            (" featuring Soft and Hard Assertions, parameterized testing, failure retry analyzers, and automated screenshot capture on test failure integrated with ExtentReports.", False, False, False)
        ],
        [
            ("Managed source code versioning in Git/GitHub with structured branch workflows, ensuring seamless collaboration and modular utility code distribution.", False, False, False)
        ]
    ]
    for b in p1_bullets:
        add_bullet(b)

    # Project 2
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_p2.paragraph_format.space_before = Pt(3)
    p_p2.paragraph_format.space_after = Pt(1)
    p_p2.paragraph_format.keep_with_next = True

    r = p_p2.add_run("Online Job Portal Management System")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_PRIMARY

    r = p_p2.add_run("\t")
    r = p_p2.add_run("Java | JSP | Servlet | Hibernate ORM | PostgreSQL | HTML/CSS")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = COLOR_MUTED

    p2_bullets = [
        [
            ("Developed a full-stack job application and recruiter management portal supporting dual user roles (Job Seekers and Recruiters) with secure session-based authentication and role-based access control (RBAC).", False, False, False)
        ],
        [
            ("Engineered recruiter workflows (job requisition posting, applicant resume filtering, shortlisting, interview scheduling) and candidate workflows (profile creation, search, real-time application tracking).", False, False, False)
        ],
        [
            ("Designed the relational database schema and back-end persistence layer using ", False, False, False),
            ("Hibernate ORM and PostgreSQL", True, False, False),
            (", writing optimized HQL queries to guarantee ACID compliance and sub-second transaction queries.", False, False, False)
        ],
        [
            ("Performed thorough manual functional testing, negative testing, and boundary value checks across all application forms, preventing data loss and unauthorized privilege escalation.", False, False, False)
        ]
    ]
    for b in p2_bullets:
        add_bullet(b)

    # Project 3 (Added project highlighting API & Backend testing depth)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_p3.paragraph_format.space_before = Pt(3)
    p_p3.paragraph_format.space_after = Pt(1)
    p_p3.paragraph_format.keep_with_next = True

    r = p_p3.add_run("REST API Test Automation & Mock Server Suite")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_PRIMARY

    r = p_p3.add_run("\t")
    r = p_p3.add_run("Postman | JavaScript | JSON Schema | RESTful APIs | Newman")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = COLOR_MUTED

    p3_bullets = [
        [
            ("Authored and structured comprehensive Postman API test collections covering CRUD operations across authentication, user profile, and transaction endpoints.", False, False, False)
        ],
        [
            ("Wrote JavaScript test scripts inside Postman to validate status codes, response times (<500ms), schema structures, and dynamic token extraction for chained requests.", False, False, False)
        ],
        [
            ("Configured Newman command-line collection runners to execute API regression test suites headlessly and output HTML test execution summaries.", False, False, False)
        ]
    ]
    for b in p3_bullets:
        add_bullet(b)

    # 6. EDUCATION
    add_section_header("Education")

    p_edu1 = doc.add_paragraph()
    p_edu1.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_edu1.paragraph_format.space_before = Pt(2)
    p_edu1.paragraph_format.space_after = Pt(1)

    r = p_edu1.add_run("Bachelor of Engineering (B.E.) in Computer Engineering")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_PRIMARY

    r = p_edu1.add_run("\t")
    if is_highlighted:
        r = p_edu1.add_run("[UPDATE: Expected 2026 / Pass-out Year]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_edu1.add_run("Graduation: 2026")
    r.font.size = Pt(9.5)
    r.italic = True

    p_edu1_sub = doc.add_paragraph()
    p_edu1_sub.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_edu1_sub.paragraph_format.space_before = Pt(0)
    p_edu1_sub.paragraph_format.space_after = Pt(2)

    r = p_edu1_sub.add_run("S.B. Patil College of Engineering, Indapur (Savitribai Phule Pune University)")
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_MUTED

    r = p_edu1_sub.add_run("\t")
    r = p_edu1_sub.add_run("CGPA: 6.77 / 10.0")
    r.bold = True
    r.font.size = Pt(9.5)

    p_hsc = doc.add_paragraph()
    p_hsc.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_hsc.paragraph_format.space_before = Pt(0)
    p_hsc.paragraph_format.space_after = Pt(1)

    r = p_hsc.add_run("Higher Secondary Certificate (HSC – 12th)  |  ")
    r.bold = True
    r.font.size = Pt(9.5)
    r = p_hsc.add_run("Maharashtra Vidyalaya, Barshi  (Maharashtra State Board)")
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_MUTED

    r = p_hsc.add_run("\t")
    r = p_hsc.add_run("Percentage: 55.00%")
    r.font.size = Pt(9.5)

    p_ssc = doc.add_paragraph()
    p_ssc.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    p_ssc.paragraph_format.space_before = Pt(0)
    p_ssc.paragraph_format.space_after = Pt(4)

    r = p_ssc.add_run("Secondary School Certificate (SSC – 10th)  |  ")
    r.bold = True
    r.font.size = Pt(9.5)
    r = p_ssc.add_run("Adarsh High School, Shendri  (Maharashtra State Board)")
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_MUTED

    r = p_ssc.add_run("\t")
    r = p_ssc.add_run("Percentage: 82.60%")
    r.font.size = Pt(9.5)
    r.bold = True

    # 7. CERTIFICATIONS & PROFESSIONAL TRAINING
    add_section_header("Certifications & Professional Training")

    cert_bullets = [
        [
            ("Full Stack Java & Automation Testing Specialist", True, False, False),
            (" — ", False, False, False),
            ("Qspiders", True, False, False),
            ("  |  ", False, False, False),
            ("[UPDATE: 2024 / Duration]" if is_highlighted else "2024", False, True, is_highlighted),
            (" — Intensive program in Core Java, Selenium WebDriver, TestNG Framework, SQL Database Development, Manual Testing Standards, and API testing principles.", False, False, False)
        ],
        [
            ("[UPDATE: Optional Additional Certification — e.g., ISTQB Foundation / HackerRank Java / Postman Student Expert]" if is_highlighted else "HackerRank Problem Solving & Core Java Certification", True if not is_highlighted else False, False, is_highlighted),
        ]
    ]
    for b in cert_bullets:
        add_bullet(b)

    # 8. KEY STRENGTHS & TESTING ATTRIBUTES
    add_section_header("Key Strengths & Methodological Highlights")
    strengths = [
        ("Defect Prevention Mindset: ", "Skilled in analyzing functional requirements early during sprint refinement to identify edge-case ambiguities before code implementation."),
        ("Test Automation Architecture: ", "Proficient in structuring modular Page Object Model codebases with reusable common utility libraries, explicit wait synchronization, and data-driven fixtures."),
        ("Cross-Functional Communication: ", "Proven ability to articulate defect severity, reproduction steps, and root-cause analysis clearly to developers, QA leads, and product owners."),
    ]
    for lbl, val in strengths:
        p_str = doc.add_paragraph()
        p_str.paragraph_format.space_before = Pt(0)
        p_str.paragraph_format.space_after = Pt(2.0)
        p_str.paragraph_format.line_spacing = 1.10
        r1 = p_str.add_run(lbl)
        r1.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_PRIMARY
        r2 = p_str.add_run(val)
        r2.font.size = Pt(9.5)

    doc.save(output_path)
    print(f"Generated 2-Page version at: {output_path}")

if __name__ == "__main__":
    out_dir = r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume"
    f1 = os.path.join(out_dir, "Rohan_Tingare_QA_Resume_2Page_Editable.docx")
    f2 = os.path.join(out_dir, "Rohan_Tingare_QA_Resume_2Page_Optimized.docx")
    create_2page_resume(f1, is_highlighted=True)
    create_2page_resume(f2, is_highlighted=False)
    
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    
    d1 = word.Documents.Open(os.path.abspath(f1))
    p_count1 = d1.ComputeStatistics(2)
    pdf1 = os.path.abspath(os.path.join(out_dir, "Rohan_Tingare_QA_Resume_2Page_Editable.pdf"))
    d1.SaveAs(pdf1, FileFormat=17)
    
    last_p1 = d1.Paragraphs(d1.Paragraphs.Count)
    vert_pos1 = last_p1.Range.Information(6)
    d1.Close()
    
    d2 = word.Documents.Open(os.path.abspath(f2))
    p_count2 = d2.ComputeStatistics(2)
    pdf2 = os.path.abspath(os.path.join(out_dir, "Rohan_Tingare_QA_Resume_2Page_Optimized.pdf"))
    d2.SaveAs(pdf2, FileFormat=17)
    
    last_p2 = d2.Paragraphs(d2.Paragraphs.Count)
    vert_pos2 = last_p2.Range.Information(6)
    d2.Close()
    
    word.Quit()
    print(f"Verified 2-Page Editable docx page count: {p_count1}, last paragraph vertical pos: {vert_pos1:.1f} pt")
    print(f"Verified 2-Page Optimized docx page count: {p_count2}, last paragraph vertical pos: {vert_pos2:.1f} pt")

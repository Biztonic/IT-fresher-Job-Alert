import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_COLOR_INDEX
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_resume(output_path, is_highlighted=True):
    doc = Document()
    
    # Page setup - Margins (0.65 in left/right, 0.6 in top/bottom)
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
    
    usable_width = 8.5 - 0.65 - 0.65 # 7.20 inches
    
    # Palette
    COLOR_PRIMARY = RGBColor(15, 48, 87)       # Deep Navy #0F3057
    COLOR_SECONDARY = RGBColor(43, 82, 152)    # Slate Blue #2B5298
    COLOR_BODY = RGBColor(30, 41, 59)          # Dark Slate #1E293B
    COLOR_MUTED = RGBColor(71, 85, 105)        # Muted Gray-Slate #475569
    COLOR_HIGHLIGHT = WD_COLOR_INDEX.YELLOW if is_highlighted else None

    # Base style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = COLOR_BODY
    
    def add_bottom_border(paragraph, color_hex="0F3057", sz="12"):
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="3" w:color="{color_hex}"/></w:pBdr>')
        pPr.append(pBdr)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title.upper())
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        add_bottom_border(p, "0F3057", "10")
        return p

    def add_bullet(text_runs, space_after=2.0):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.12
        p.paragraph_format.left_indent = Inches(0.22)
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

    # ---------------------------------------------------------
    # 1. HEADER
    # ---------------------------------------------------------
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(1)
    r_name = p_name.add_run("ROHAN TINGARE")
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(21)
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
    
    r = p_contact1.add_run("rohantingare1414@gmail.com")
    r.font.size = Pt(9.5)
    r = p_contact1.add_run("  •  +91 9021753459  •  ")
    r.font.size = Pt(9.5)
    
    if is_highlighted:
        r = p_contact1.add_run("[UPDATE: Pune, Maharashtra, India]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True
    else:
        r = p_contact1.add_run("Pune, Maharashtra, India")
    r.font.size = Pt(9.5)

    # Contact line 2 (Links)
    p_contact2 = doc.add_paragraph()
    p_contact2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact2.paragraph_format.space_before = Pt(0)
    p_contact2.paragraph_format.space_after = Pt(6)

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
        r = p_contact2.add_run("[UPDATE: Optional Link]")
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    else:
        r = p_contact2.add_run("rohantingare.dev")
        r.font.color.rgb = COLOR_SECONDARY
    r.font.size = Pt(9.5)

    # ---------------------------------------------------------
    # 2. PROFESSIONAL SUMMARY
    # ---------------------------------------------------------
    add_section_header("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(5)
    p_sum.paragraph_format.line_spacing = 1.15
    
    r = p_sum.add_run(
        "Results-oriented QA Automation Engineer and Computer Engineering graduate with 6+ months of hands-on industry experience "
        "designing robust test automation suites, executing comprehensive manual testing, and performing REST API validation in Agile SDLC environments. "
        "Skilled in architecting maintainable Page Object Model (POM) frameworks using Java, Selenium WebDriver, TestNG, and Maven, coupled with automated reporting. "
        "Proficient in backend SQL verification, Postman endpoint testing, CI/CD pipeline integration, and Jira defect lifecycle tracking. "
        "Possesses a strong foundation in Core Java and full-stack web architecture, enabling rapid root-cause defect analysis and proactive cross-functional collaboration "
        "to deliver zero-defect, high-performance software releases."
    )
    r.font.size = Pt(9.5)

    # ---------------------------------------------------------
    # 3. CORE COMPETENCIES & TECHNICAL SKILLS
    # ---------------------------------------------------------
    add_section_header("Technical Skills")

    skills_data = [
        ("Test Automation & Frameworks: ", [
            ("Selenium WebDriver, TestNG, Page Object Model (POM), PageFactory, Data-Driven Testing, Apache POI, ExtentReports, Maven, Cross-Browser & Headless Testing", False, False, False)
        ]),
        ("Manual Testing & Methodologies: ", [
            ("Software Testing Life Cycle (STLC), SDLC, Agile / Scrum, Sprint Ceremonies, Test Plan & Test Case Design, Functional & Regression Testing, Smoke & Sanity Testing, Boundary Value Analysis (BVA), Equivalence Partitioning (EP), Defect Life Cycle, Jira, Root Cause Analysis, Traceability Matrix (RTM)", False, False, False)
        ]),
        ("API & Backend Testing: ", [
            ("Postman, REST API Testing, JSON Schema & Payload Validation, HTTP Status Code Verification, Endpoint Mocking, Environment Variables", False, False, False)
        ]),
        ("Programming & Databases: ", [
            ("Core Java (OOPs, Collections, Multithreading, Exception Handling), SQL (Joins, Subqueries, Aggregations), JavaScript, HTML5, CSS3, PostgreSQL, MySQL, SQLite, JDBC", False, False, False)
        ]),
        ("CI/CD, Build & Tools: ", [
            ("Git, GitHub, Jenkins (CI/CD Automated Test Execution), Maven, Jira, Eclipse, Spring Tool Suite (STS)", False, False, False)
        ]),
        ("Backend & Framework Knowledge: ", [
            ("Spring Boot, Spring MVC, Hibernate ORM, JPA, Servlets, JSP, RESTful Architecture", False, False, False)
        ]),
    ]

    for label, runs in skills_data:
        p_skill = doc.add_paragraph()
        p_skill.paragraph_format.space_before = Pt(0)
        p_skill.paragraph_format.space_after = Pt(2)
        p_skill.paragraph_format.line_spacing = 1.12
        
        r_lbl = p_skill.add_run(label)
        r_lbl.bold = True
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.color.rgb = COLOR_PRIMARY
        
        for text, bold, italic, hl in runs:
            r = p_skill.add_run(text)
            r.font.size = Pt(9.5)
            r.bold = bold
            r.italic = italic
            if hl and is_highlighted:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW

    # ---------------------------------------------------------
    # 4. PROFESSIONAL EXPERIENCE
    # ---------------------------------------------------------
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

    # Experience Bullets
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
        ]
    ]

    for b in exp_bullets:
        add_bullet(b)

    # ---------------------------------------------------------
    # 5. KEY TECHNICAL PROJECTS
    # ---------------------------------------------------------
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
    r = p_p1.add_run("Java | Selenium WebDriver | TestNG | Maven | ExtentReports")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = COLOR_MUTED

    proj1_bullets = [
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
        ]
    ]
    for b in proj1_bullets:
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

    proj2_bullets = [
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
    for b in proj2_bullets:
        add_bullet(b)

    # ---------------------------------------------------------
    # 6. EDUCATION
    # ---------------------------------------------------------
    add_section_header("Education")

    # College
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

    # 12th & 10th combined clean single lines
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

    # ---------------------------------------------------------
    # 7. CERTIFICATIONS & TRAINING
    # ---------------------------------------------------------
    add_section_header("Certifications & Professional Training")

    cert_bullets = [
        [
            ("Full Stack Java & Automation Testing Specialist", True, False, False),
            (" — ", False, False, False),
            ("Qspiders", True, False, False),
            (" | ", False, False, False),
            ("[UPDATE: 2024 / Duration]" if is_highlighted else "2024", False, True, is_highlighted),
            ("\n• Rigorous professional certification program focused on Core Java, Object-Oriented Programming, Selenium WebDriver, TestNG Framework, SQL Database Development, Manual Testing Standards, and API testing principles.", False, False, False)
        ],
        [
            ("[UPDATE: Optional Additional Certification — e.g., ISTQB Foundation / Agile Scrum / HackerRank Java (Basic)]" if is_highlighted else "HackerRank Problem Solving & Java Skills Certification", True if not is_highlighted else False, False, is_highlighted),
        ]
    ]

    for b in cert_bullets:
        add_bullet(b, space_after=3.0)

    # Save
    doc.save(output_path)
    print(f"Successfully generated: {output_path}")

if __name__ == "__main__":
    out_dir = r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume"
    p1 = os.path.join(out_dir, "Rohan_Tingare_Resume_Template_Editable.docx")
    p2 = os.path.join(out_dir, "Rohan_Tingare_Resume_Ready_To_Use.docx")
    create_resume(p1, is_highlighted=True)
    create_resume(p2, is_highlighted=False)

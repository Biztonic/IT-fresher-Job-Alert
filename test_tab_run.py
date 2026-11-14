import docx
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_TAB_ALIGNMENT

doc = Document()
section = doc.sections[0]
section.left_margin = Inches(0.65)
section.right_margin = Inches(0.65)
usable_width = 8.5 - 0.65 - 0.65 # 7.2 inches

p = doc.add_paragraph()
p.paragraph_format.tab_stops.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
r1 = p.add_run("QA Intern (Automation & Manual Testing) — Test Yantra")
r1.bold = True
r_tab = p.add_run("\t")
r2 = p.add_run("Pune, India | June 2024 – Nov 2024")
r2.italic = True

doc.save(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\test_tab.docx")
print("Saved test_tab.docx")

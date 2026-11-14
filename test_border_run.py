import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

doc = Document()
p = doc.add_paragraph()
run = p.add_run("PROFESSIONAL EXPERIENCE")
run.font.name = "Calibri"
run.font.size = Pt(12)
run.font.bold = True
run.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)

pPr = p._p.get_or_add_pPr()
pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="3" w:color="1A365D"/></w:pBdr>')
pPr.append(pBdr)

p2 = doc.add_paragraph()
r_norm = p2.add_run("Standard text and ")
r_highlight = p2.add_run("[UPDATE: Missing Date Here]")
r_highlight.font.highlight_color = WD_COLOR_INDEX.YELLOW
r_highlight.font.bold = True

doc.save(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\test_border.docx")
print("Saved test_border.docx successfully")

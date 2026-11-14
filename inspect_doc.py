import docx

doc = docx.Document(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\Rohan_Tingare_Resume_Template_Editable.docx")
print("Total paragraphs:", len(doc.paragraphs))
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"P{i:02d}: {p.text[:75]}...")

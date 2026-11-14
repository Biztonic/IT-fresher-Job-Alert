import win32com.client
import os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

docx_path = os.path.abspath(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\Rohan_Tingare_QA_Resume_1Page_Editable.docx")
doc = word.Documents.Open(docx_path)

for i, p in enumerate(doc.Paragraphs):
    page_num = p.Range.Information(3)
    if page_num == 2:
        print(f"Overflow paragraph {i} on Page 2: {p.Range.Text.strip()}")

doc.Close()
word.Quit()

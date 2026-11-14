import win32com.client
import os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

docx_path = os.path.abspath(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\Rohan_Tingare_Resume_Template_Editable.docx")
doc = word.Documents.Open(docx_path)

print("Total pages:", doc.ComputeStatistics(2))

# Let's inspect where page breaks or what is at the end
# wdActiveEndPageNumber = 3
for i, p in enumerate(doc.Paragraphs):
    page_num = p.Range.Information(3) # 3 = wdActiveEndPageNumber
    if i == 0 or page_num != prev_page:
        print(f"--- Page {page_num} starts around Paragraph {i}: {p.Range.Text[:50]}...")
    prev_page = page_num

print(f"Last paragraph is on Page: {page_num}")
doc.Close()
word.Quit()

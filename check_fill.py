import win32com.client
import os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

docx_path = os.path.abspath(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\Rohan_Tingare_Resume_Template_Editable.docx")
doc = word.Documents.Open(docx_path)

last_p = doc.Paragraphs(doc.Paragraphs.Count)
print("Last paragraph text:", last_p.Range.Text)
print("Last paragraph page:", last_p.Range.Information(3)) # wdActiveEndPageNumber
# wdVerticalPositionRelativeToPage = 6
vert_pos_points = last_p.Range.Information(6)
print(f"Vertical position on page: {vert_pos_points:.1f} pt (Page height is 792 pt, bottom margin at {792 - 0.55*72:.1f} pt)")
print(f"Percentage of page 2 used: {(vert_pos_points / (792 - 0.55*72)) * 100:.1f}%")

doc.Close()
word.Quit()

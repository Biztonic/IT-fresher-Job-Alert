import win32com.client
import os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

docx_path = os.path.abspath(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\Rohan_Tingare_Resume_Template_Editable.docx")
pdf_path = os.path.abspath(r"C:\Users\Administrator\.gemini\antigravity\scratch\rohan_tingare_resume\Rohan_Tingare_Resume_Template_Editable.pdf")

doc = word.Documents.Open(docx_path)
page_count = doc.ComputeStatistics(2) # 2 = wdStatisticPages
print(f"Page count for Template: {page_count}")
doc.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
doc.Close()
word.Quit()
print("Converted to PDF successfully!")

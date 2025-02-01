import _3_1_3
import merged
import _3_1_2
from docx import Document
import os


file = "output.docx"
merged.CO_Table()
_3_1_2.process_docx_files()
_3_1_3.main()


doc = Document(file)
print(f"Data successfully saved in {file}")
os.startfile(file)
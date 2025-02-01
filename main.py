import _3_1
import _3_1_3
import _3_1_1
import _3_1_2
from docx import Document
import os

file = "output.docx"

_3_1.PO_PSO()
_3_1_1.CO_Table()
_3_1_2.process_docx_files()
_3_1_3.main()

# Adjust margins
doc = Document(file)

# Save the changes to the document
doc.save(file)

print(f"Data successfully saved in {file}")
os.startfile(file)

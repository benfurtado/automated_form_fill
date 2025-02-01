import _3_1_3
import merged
from docx import Document

merged.CO_Table()
doc = Document("output.docx")
doc.add_paragraph("\n" + "-" * 50 + "\n")
doc.save("output1.docx")

_3_1_3.main()
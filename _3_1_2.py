from docx import Document
import os
import re

def extract_course_info(paragraphs):
    course_info = {
        "course_code": "Not Found",
        "class_semester": "Not Found"
    }
    
    for para in paragraphs:
        text = para.text.strip()
        if "Course Code and Name:" in text:
            course_info["course_code"] = text.split(":", 1)[1].strip().replace("\u2013", "-")
        elif "Class and Semester:" in text:
            course_info["class_semester"] = text.split(":", 1)[1].strip()
    return course_info

def extract_revised_co_po_table(doc):
    """Extracts the table following the 'Revised CO-PO Mapping' heading."""
    table_data = []
    
    # Find the paragraph that mentions "Revised CO-PO Mapping"
    target_para_found = False
    for para in doc.paragraphs:
        if "revised co-po mapping" in para.text.lower():
            target_para_found = True
            break
    
    if not target_para_found:
        return table_data
    
    # Find the first table in the document
    if doc.tables:
        table = doc.tables[6]  # Assuming the 7th table is the required one
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if cells and any(x in cells[0].lower() for x in ["total", "average"]):
                continue
            if any(cell.strip() for cell in cells):
                table_data.append(cells)
    
    return table_data

def process_docx_files():
    """Processes all .docx files in 'temp_files' and creates a combined report."""
    output_doc = Document('output.docx')
    output_doc.add_heading('3.1.2 CO-PO and CO-PSO matrices', 0)
    
    temp_dir = "temp_files"
    files = sorted([f for f in os.listdir(temp_dir) if f.endswith(".docx") and f.startswith("18")],
                   key=lambda x: re.findall(r'\d{4}-\d{2}', x))
    
    for filename in files:
        file_path = os.path.join(temp_dir, filename)
        doc = Document(file_path)
        
        # Extract course info
        course_info = extract_course_info(doc.paragraphs)
        
        # Extract CO-PO table
        co_po_table = extract_revised_co_po_table(doc)
        
        # Format filename with year
        year_match = re.findall(r'\d{4}-\d{2}', filename)
        year = year_match[0] if year_match else "Unknown"
        heading = f"{filename} ({year})"
        
        # Add content to the document
        output_doc.add_heading(heading, level=1)
        output_doc.add_paragraph(f"Course Code and Name: {course_info['course_code']}")
        output_doc.add_paragraph(f"Class and Semester: {course_info['class_semester']}")
        
        if co_po_table:
            table = output_doc.add_table(rows=len(co_po_table)-1, cols=len(co_po_table[0]))
            table.style = "Table Grid"
            
            # Data rows (skip the first row)
            for row_idx, row in enumerate(co_po_table[1:]):
                new_row = table.rows[row_idx].cells
                for i, cell in enumerate(row):
                    new_row[i].text = cell
        else:
            output_doc.add_paragraph("Revised CO-PO Mapping table not found").italic = True
        
    
    # Save and open the final document
    output_file = "output.docx"
    output_doc.save(output_file)
    print(f"\nReport saved: {output_file}")


import os
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt
from connection import drive, list_files_in_folder

# Function to process DOCX files
def process_docx_file(file):
    try:
        print(f"Processing file: {file['title']}")

        # Download the file content as binary data
        file.GetContentFile(file['title'])

        # Open the .docx file
        doc = Document(file['title'])

        # Extract data
        course_code = None
        subject = None
        table_data = []

        for para in doc.paragraphs:
            if "Course Code and Name:" in para.text:
                line = para.text.split("Course Code and Name:")[1].strip()
                if "-" in line:
                    course_code, subject = line.split("-", 1)
                    course_code = course_code.strip()
                    subject = subject.strip()

            if "Revised CO-PO Mapping:" in para.text:
                # Assuming the table follows immediately after the paragraph
                table_index = doc.paragraphs.index(para) + 1
                if table_index < len(doc.tables):
                    table = doc.tables[table_index]
                    for row in table.rows:
                        second_column = row.cells[1].text.strip()  # Second column
                        last_column = row.cells[-1].text.strip()  # Last column
                        table_data.append((second_column, last_column))

        # Return the extracted data
        return {"course_code": course_code, "subject": subject, "table_data": table_data}
    except Exception as e:
        print(f"Error processing DOCX file {file['title']}: {e}")
        return None

# Function to write data directly into output.docx
def write_to_output_docx(extracted_data, doc):
    try:
        # Add course code and subject
        doc.add_heading('Course Code and Subject:', level=1)
        doc.add_paragraph(f"Course Code: {extracted_data['course_code']}")
        doc.add_paragraph(f"Subject: {extracted_data['subject']}")

        # Add CO-PO Mapping Table
        doc.add_heading('Revised CO-PO Mapping Table:', level=1)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'

        # Adding headers to the table
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Second Column'
        hdr_cells[1].text = 'Last Column'

        # Add table data
        for row in extracted_data['table_data']:
            row_cells = table.add_row().cells
            row_cells[0].text = row[0]
            row_cells[1].text = row[1]
    except Exception as e:
        print(f"Error writing data to output DOCX: {e}")

# Function to recursively process all folders and files
def process_folders(folder_id, doc):
    try:
        file_list = list_files_in_folder(folder_id)

        for file in file_list:
            if file['mimeType'] == 'application/vnd.google-apps.folder':  # Folder
                print(f"Entering folder: {file['title']}")
                process_folders(file['id'], doc)  # Recursive call for nested folders
            elif file['title'].startswith("18") and file['title'].endswith(".docx"):  # DOCX files starting with '18'
                extracted_data = process_docx_file(file)
                if extracted_data:
                    write_to_output_docx(extracted_data, doc)
    except Exception as e:
        print(f"Error processing folder ID {folder_id}: {e}")

# Function to write Excel data to output.docx
def process_excel_and_write_to_docx(excel_file_path, doc):
    try:
        workbook = load_workbook(excel_file_path)
        sheet = workbook.active

        # Add heading for Excel data
        doc.add_heading('Excel Data:', level=1)

        # Create a table in Word for the Excel data
        table = doc.add_table(rows=1, cols=sheet.max_column)
        table.style = 'Table Grid'

        # Add header row
        header_cells = table.rows[0].cells
        for col_index, column in enumerate(sheet.iter_cols(1, sheet.max_column, 1, 1), start=1):
            header_cells[col_index - 1].text = str(column[0].value)

        # Add Excel data rows
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_cells = table.add_row().cells
            for col_index, cell_value in enumerate(row):
                row_cells[col_index].text = str(cell_value) if cell_value is not None else ""
    except Exception as e:
        print(f"Error processing Excel file {excel_file_path}: {e}")

# Main function
def main():
    try:
        # Replace 'your-folder-id' with the actual folder ID from Google Drive
        root_folder_id = '1Bbp_TRb2dt-oRcKo3C7vHK7AHf0Hy90p'

        # Create a new DOCX file to save the extracted data
        output_file_path = 'output.docx'
        doc = Document()
        doc.add_heading('Extracted Information', 0)

        # Process folders and files
        process_folders(root_folder_id, doc)

        # Process Excel file and write data
        excel_file_path = '14 15 PO and PSO.xlsx'
        process_excel_and_write_to_docx(excel_file_path, doc)

        # Save the output file
        doc.save(output_file_path)
        print(f"Data successfully saved in {output_file_path}")
        os.startfile(output_file_path)
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

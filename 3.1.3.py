from docx import Document
import os
from connection import drive, list_files_in_folder

# Create a temporary directory for storing intermediate files
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

# Function to process DOCX files
# Function to process DOCX files
def process_docx_file(file):
    try:
        print(f"Processing file: {file['title']}")

        # Download the file content as binary data to the temp folder
        temp_file_path = os.path.join(TEMP_DIR, file['title'])
        file.GetContentFile(temp_file_path)

        # Open the .docx file
        doc = Document(temp_file_path)

        # Initialize variables to store extracted data
        course_code = None
        subject = None
        extracted_table = None

        # Extract course code and subject
        for para in doc.paragraphs:
            if "Course Code and Name:" in para.text:
                line = para.text.split("Course Code and Name:")[1].strip()
                if "–" in line:
                    course_code, subject = line.split("–", 1)
                    course_code = course_code.strip()
                    subject = subject.strip()
                break

        # Locate "Revised CO-PO Mapping:" and extract the following table
        table_found = False
        for para_idx, para in enumerate(doc.paragraphs):
            if "Revised CO-PO Mapping:" in para.text:
                table_found = True
                print(f"Found 'Revised CO-PO Mapping:' at line {para_idx + 1}")

                # Locate the table closest to this paragraph
                for table_idx, table in enumerate(doc.tables):
                    # Check the previous sibling of the table to find the heading
                    previous_paragraph = table._element.getprevious()
                    if previous_paragraph is not None and "Revised CO-PO Mapping:" in previous_paragraph.text:
                        extracted_table = table
                        print(f"Found table after 'Revised CO-PO Mapping:' at index {table_idx}")
                        break
                break

        if not table_found:
            print("No 'Revised CO-PO Mapping:' section found.")

        return {"course_code": course_code, "subject": subject, "table_data": extracted_table}
    except Exception as e:
        print(f"Error processing DOCX file {file['title']}: {e}")
        return None



# Function to write data to output.docx
def write_to_output_docx(extracted_data, doc):
    try:
        doc.add_heading("Course Code and Subject", level=1)
        doc.add_paragraph(f"Course Code: {extracted_data['course_code']}")
        doc.add_paragraph(f"Subject: {extracted_data['subject']}")

        doc.add_heading("Revised CO-PO Mapping", level=1)

        if extracted_data["table_data"] is not None:
            original_table = extracted_data["table_data"]
            table = doc.add_table(rows=len(original_table.rows), cols=len(original_table.columns))
            table.style = "Table Grid"

            for row_idx, row in enumerate(original_table.rows):
                for col_idx, cell in enumerate(row.cells):
                    table.cell(row_idx, col_idx).text = cell.text.strip()
        else:
            doc.add_paragraph("No table found below 'Revised CO-PO Mapping:'.")
    except Exception as e:
        print(f"Error writing data to output DOCX: {e}")

# Function to process folders and files recursively
def process_folders(folder_id, doc):
    try:
        file_list = list_files_in_folder(folder_id)
        for file in file_list:
            if file["mimeType"] == "application/vnd.google-apps.folder":  # Folder
                print(f"Entering folder: {file['title']}")
                process_folders(file["id"], doc)
            elif file["title"].startswith("18") and file["title"].endswith(".docx"):  # DOCX files starting with '18'
                extracted_data = process_docx_file(file)
                if extracted_data:
                    write_to_output_docx(extracted_data, doc)
    except Exception as e:
        print(f"Error processing folder ID {folder_id}: {e}")

# Main function
def main():
    try:
        root_folder_id = "1Bbp_TRb2dt-oRcKo3C7vHK7AHf0Hy90p"
        output_file_path = "output.docx"

        doc = Document()
        doc.add_heading("Extracted Information", 0)

        process_folders(root_folder_id, doc)
        doc.save(output_file_path)
        print(f"Data successfully saved in {output_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

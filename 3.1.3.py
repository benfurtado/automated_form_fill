from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from docx import Document
import openpyxl
import os
import sys

# Path to your service account JSON file
SERVICE_ACCOUNT_FILE = r'C:\Users\LENOVO\OneDrive\Desktop\PYQs\form filling\automated_form_fill\councellingapp-445207-98a66903ca2d.json'

# Folder ID for Google Drive
FOLDER_ID = '1Bbp_TRb2dt-oRcKo3C7vHK7AHf0Hy90p'

# Define the required scopes
SCOPES = ['https://www.googleapis.com/auth/drive']

try:
    # Authenticate using ServiceAccountCredentials
    credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
    gauth = GoogleAuth()
    gauth.credentials = credentials
    drive = GoogleDrive(gauth)
    print("Authenticated successfully using Service Account!")
except Exception as e:
    print(f"Error during authentication: {e}")
    sys.exit(1)  # Exit the script on error

# Query all files in the folder
query = f"'{FOLDER_ID}' in parents and trashed=false"
try:
    file_list = drive.ListFile({'q': query}).GetList()
    print("Files retrieved successfully!")
except Exception as e:
    print(f"Error retrieving files: {e}")
    sys.exit(1)

# Create a new Word document to save the combined extracted rows
output_doc = Document()

# Process each file starting with "26"
for file in file_list:
    if file['title'].startswith("26"):  # Only process files starting with "26"
        print(f"Processing file: {file['title']}")

        file_content = BytesIO()
        file.GetContentFile(file_content)
        file_content.seek(0)

        if file['title'].endswith(".docx"):
            # Process .docx files
            doc = Document(file_content)

            for table in doc.tables:
                if len(table.rows) < 2:
                    continue

                header_row = table.rows[1]
                average_row = table.rows[-1]

                num_columns = len(header_row.cells)
                new_table = output_doc.add_table(rows=0, cols=num_columns)
                new_table.style = 'Table Grid'

                header_cells = new_table.add_row().cells
                for col_idx, cell in enumerate(header_row.cells):
                    header_cells[col_idx].text = cell.text

                average_cells = new_table.add_row().cells
                for col_idx, cell in enumerate(average_row.cells):
                    average_cells[col_idx].text = cell.text

        elif file['title'].endswith(".xlsx"):
            # Process .xlsx files
            wb = openpyxl.load_workbook(file_content)
            for sheet in wb.sheetnames:
                ws = wb[sheet]

                if ws.max_row >= 2:
                    header_row = ws[2]
                    average_row = ws[ws.max_row]

                    num_columns = len(header_row)
                    new_table = output_doc.add_table(rows=0, cols=num_columns)
                    new_table.style = 'Table Grid'

                    header_cells = new_table.add_row().cells
                    for col_idx, cell in enumerate(header_row):
                        header_cells[col_idx].text = str(cell.value)

                    average_cells = new_table.add_row().cells
                    for col_idx, cell in enumerate(average_row):
                        average_cells[col_idx].text = str(cell.value)

        else:
            print(f"Skipping unsupported file format: {file['title']}")

# Save the output document
output_file_path = 'Combined_Extracted_Content.docx'
output_doc.save(output_file_path)

print(f"Filtered content saved to {output_file_path}.")

from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials

# Path to your service account JSON file
SERVICE_ACCOUNT_FILE = r'form-filling-448211-63002aeb7b96.json'
# Define required scopes for Google Drive access
SCOPES = ['https://www.googleapis.com/auth/drive']

# Global drive object
drive = None

def authenticate_drive():
    """
    Authenticates the Google Drive API using service account credentials
    and initializes the global 'drive' object.
    """
    global drive
    try:
        # Authenticate using Service Account credentials
        credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
        gauth = GoogleAuth()
        gauth.credentials = credentials
        drive = GoogleDrive(gauth)
        print("Authenticated successfully!")
    except Exception as e:
        print(f"Authentication failed: {e}")
        exit()

def list_files_in_folder(folder_id):
    """
    Lists all files and subfolders in the specified Google Drive folder.

    Args:
        folder_id (str): The ID of the Google Drive folder to process.

    Returns:
        list: A list of dictionaries containing file or folder metadata.
    """
    try:
        query = f"'{folder_id}' in parents and trashed=false"
        items = drive.ListFile({'q': query}).GetList()
        return items
    except Exception as e:
        print(f"Error listing files in folder ID {folder_id}: {e}")
        return []

def list_docx_files(folder_id):
    """
    Recursively lists all .docx files in a Google Drive folder and its subfolders.

    Args:
        folder_id (str): The ID of the Google Drive folder to process.
    """
    try:
        # Get all items (files and folders) in the current folder
        items = list_files_in_folder(folder_id)
        
        for item in items:
            if item['mimeType'] == 'application/vnd.google-apps.folder':  # If it's a folder
                print(f"Folder: {item['title']} (ID: {item['id']})")
                # Recursively call this function for the subfolder
                list_docx_files(item['id'])
            elif item['title'].endswith(".docx"):  # If it's a .docx file
                print(f"File: {item['title']} (ID: {item['id']})")
    except Exception as e:
        print(f"Error processing folder ID {folder_id}: {e}")

# Authenticate when the module is imported
authenticate_drive()

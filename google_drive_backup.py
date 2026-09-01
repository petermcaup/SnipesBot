from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from datetime import datetime, timedelta
import json
import os
import sys

# Use BASE_DIR for consistent path handling across Windows and Pi
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.dirname(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TOKEN_FILE = os.path.join(BASE_DIR, 'private', 'google_drive_token.json')
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'private', 'credentials.json')
BACKUP_FOLDER_NAME = 'SnipesBot Backups'

def load_credentials():
    """Load saved Google Drive credentials."""
    if not os.path.exists(TOKEN_FILE):
        return None

    with open(TOKEN_FILE, 'r') as f:
        creds_data = json.load(f)

    creds = Credentials(
        token=creds_data['token'],
        refresh_token=creds_data['refresh_token'],
        token_uri=creds_data['token_uri'],
        client_id=creds_data['client_id'],
        client_secret=creds_data['client_secret'],
        scopes=creds_data['scopes']
    )

    # Refresh if needed and save back to file
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        save_credentials(creds)

    return creds

def save_credentials(creds):
    """Save refreshed credentials back to the token file."""
    creds_data = {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': creds.scopes
    }
    with open(TOKEN_FILE, 'w') as f:
        json.dump(creds_data, f, indent=4)

def get_drive_service():
    """Get authenticated Google Drive service."""
    creds = load_credentials()
    if not creds:
        raise Exception("Google Drive credentials not found. Run auth_google_drive.py first.")
    return build('drive', 'v3', credentials=creds)

def get_or_create_backup_folder():
    """Get or create the backup folder in Google Drive."""
    service = get_drive_service()
    
    # Search for existing folder
    results = service.files().list(
        q=f"name='{BACKUP_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        spaces='drive',
        pageSize=1,
        fields='files(id, name)'
    ).execute()
    
    files = results.get('files', [])
    if files:
        return files[0]['id']
    
    # Create folder if it doesn't exist
    file_metadata = {
        'name': BACKUP_FOLDER_NAME,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = service.files().create(body=file_metadata, fields='id').execute()
    return folder['id']

def upload_backup(excel_file_path):
    """Upload Excel file to Google Drive and cleanup old backups."""
    service = get_drive_service()
    folder_id = get_or_create_backup_folder()
    
    # Create filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"SNIPESSTATS_backup_{timestamp}.xlsx"
    
    # Upload file
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaFileUpload(excel_file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, createdTime'
    ).execute()
    
    print(f"✅ Uploaded backup: {filename}")
    
    # Delete backups older than 5 hours
    cleanup_old_backups(service, folder_id)
    
    return file['id']

def cleanup_old_backups(service, folder_id):
    """Delete backups older than 5 hours."""
    cutoff_time = datetime.utcnow() - timedelta(hours=5)
    cutoff_iso = cutoff_time.isoformat() + 'Z'
    
    # List files older than 5 hours
    results = service.files().list(
        q=f"'{folder_id}' in parents and createdTime < '{cutoff_iso}' and trashed=false",
        spaces='drive',
        pageSize=10,
        fields='files(id, name, createdTime)'
    ).execute()
    
    files = results.get('files', [])
    for file in files:
        service.files().delete(fileId=file['id']).execute()
        print(f"🗑️ Deleted old backup: {file['name']}")

if __name__ == '__main__':
    # Test upload
    test_file = 'SNIPESSTATS.xlsx'
    if os.path.exists(test_file):
        upload_backup(test_file)
    else:
        print("No SNIPESSTATS.xlsx found for testing")

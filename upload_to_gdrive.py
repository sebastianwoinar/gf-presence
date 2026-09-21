"""
Upload the latest GF_Praesenzplanung_*.xlsx to a Google Drive shared folder.

Setup (one-time):
  pip install google-auth google-auth-oauthlib google-api-python-client
  1. Create a project in Google Cloud Console
  2. Enable the Google Drive API
  3. Create OAuth 2.0 credentials (Desktop app) → download as credentials.json
  4. Place credentials.json in the same directory as this script
  5. Set GDRIVE_FOLDER_ID in .env (the shared folder's ID from its URL)

Usage:
  python3 upload_to_gdrive.py                        # uploads latest xlsx
  python3 upload_to_gdrive.py GF_Praesenzplanung_2026-08.xlsx  # specific file
"""

import sys
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

HERE = Path(__file__).parent
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
TOKEN_FILE = HERE / "token.json"
CREDS_FILE = HERE / "credentials.json"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def load_env() -> dict:
    env = {}
    env_path = HERE / ".env"
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def get_credentials():
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDS_FILE.exists():
                raise FileNotFoundError(
                    f"credentials.json not found at {CREDS_FILE}\n"
                    "Download it from Google Cloud Console → APIs & Services → Credentials."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
    return creds


def find_latest_xlsx() -> Path:
    files = sorted(HERE.glob("GF_Praesenzplanung_????-??.xlsx"))
    if not files:
        raise FileNotFoundError("No GF_Praesenzplanung_YYYY-MM.xlsx found in script directory.")
    return files[-1]


def upload(file_path: Path, folder_id: str) -> str:
    creds   = get_credentials()
    service = build("drive", "v3", credentials=creds)

    # Check if a file with the same name already exists in the folder → update it
    # supportsAllDrives=True is required for Shared Drives (Team Drives)
    existing = service.files().list(
        q=f"name='{file_path.name}' and '{folder_id}' in parents and trashed=false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute().get("files", [])

    media = MediaFileUpload(file_path, mimetype=MIME_XLSX, resumable=True)

    if existing:
        file_id = existing[0]["id"]
        service.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True,
        ).execute()
        print(f"✓ Aktualisiert: {file_path.name} (id: {file_id})")
        return file_id
    else:
        meta = {"name": file_path.name, "parents": [folder_id]}
        result = service.files().create(
            body=meta,
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()
        file_id = result["id"]
        print(f"✓ Hochgeladen: {file_path.name} (id: {file_id})")
        return file_id


if __name__ == "__main__":
    env       = load_env()
    folder_id = env.get("GDRIVE_FOLDER_ID", "").strip()
    if not folder_id:
        raise ValueError(
            "GDRIVE_FOLDER_ID not set in .env\n"
            "Add:  GDRIVE_FOLDER_ID=<your-folder-id>"
        )

    if len(sys.argv) > 1:
        file_path = HERE / sys.argv[1]
    else:
        file_path = find_latest_xlsx()

    print(f"Lade hoch: {file_path.name} → Ordner {folder_id}")
    upload(file_path, folder_id)

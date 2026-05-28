"""
SharePoint/OneDrive reader using Microsoft Graph API
Uses Azure AD app credentials from .env file
"""
import os, io, tempfile, requests
from pathlib import Path

class SharePointReader:
    GRAPH_URL = "https://graph.microsoft.com/v1.0"
    
    def __init__(self, client_id, tenant_id, client_secret):
        self.client_id     = client_id
        self.tenant_id     = tenant_id
        self.client_secret = client_secret
        self.token         = None
        self._authenticate()

    def _authenticate(self):
        url  = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type":    "client_credentials",
            "client_id":     self.client_id,
            "client_secret": self.client_secret,
            "scope":         "https://graph.microsoft.com/.default"
        }
        resp = requests.post(url, data=data)
        resp.raise_for_status()
        self.token = resp.json()["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def _get_drive_id(self, user_email):
        """Get the OneDrive drive ID for a user."""
        url  = f"{self.GRAPH_URL}/users/{user_email}/drive"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        return resp.json()["id"]

    def list_folder(self, user_email, folder_path):
        """List files and subfolders in a SharePoint folder."""
        drive_id = self._get_drive_id(user_email)
        # Encode path
        path_enc = folder_path.strip("/")
        url  = f"{self.GRAPH_URL}/drives/{drive_id}/root:/{path_enc}:/children"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        return resp.json().get("value", [])

    def download_file(self, user_email, file_path):
        """Download a file and return as bytes."""
        drive_id = self._get_drive_id(user_email)
        path_enc = file_path.strip("/")
        url  = f"{self.GRAPH_URL}/drives/{drive_id}/root:/{path_enc}:/content"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        return resp.content

    def download_to_temp(self, user_email, file_path):
        """Download file to temp directory and return local path."""
        content  = self.download_file(user_email, file_path)
        filename = Path(file_path).name
        tmp_path = os.path.join(tempfile.gettempdir(), f"nen_sp_{filename}")
        with open(tmp_path, "wb") as f:
            f.write(content)
        return tmp_path

    def find_file_in_folder(self, user_email, folder_path, keyword):
        """Find first file in folder whose name contains keyword."""
        items = self.list_folder(user_email, folder_path)
        for item in items:
            if keyword.lower() in item["name"].lower() and "file" in item:
                return folder_path.rstrip("/") + "/" + item["name"]
        return None

    def list_subfolders(self, user_email, folder_path):
        """List only subfolders (venture folders)."""
        items = self.list_folder(user_email, folder_path)
        return [i["name"] for i in items if "folder" in i]

    def list_files(self, user_email, folder_path):
        """List only files in a folder."""
        items = self.list_folder(user_email, folder_path)
        return [i["name"] for i in items if "file" in i]

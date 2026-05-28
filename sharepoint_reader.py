"""
Personal OneDrive reader using Microsoft Graph API
Correctly accesses user's personal OneDrive drive
"""
import os, tempfile, requests
from pathlib import Path

class SharePointReader:
    GRAPH_URL  = "https://graph.microsoft.com/v1.0"
    USER_UPN   = "meenakshi.singh@wadhwanifoundation.org"

    def __init__(self, client_id, tenant_id, client_secret):
        self.client_id     = client_id
        self.tenant_id     = tenant_id
        self.client_secret = client_secret
        self.token         = None
        self.drive_id      = None
        self._authenticate()
        self._get_personal_drive()

    def _authenticate(self):
        url  = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type":    "client_credentials",
            "client_id":     self.client_id,
            "client_secret": self.client_secret,
            "scope":         "https://graph.microsoft.com/.default"
        }
        resp = requests.post(url, data=data, timeout=30)
        if not resp.ok:
            raise Exception(f"Auth failed ({resp.status_code}): {resp.json()}")
        self.token = resp.json()["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def _get_personal_drive(self):
        """Get personal OneDrive drive for user — uses /users/{upn}/drive endpoint."""
        url  = f"{self.GRAPH_URL}/users/{self.USER_UPN}/drive"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        if not resp.ok:
            raise Exception(f"Could not get personal drive ({resp.status_code}): {resp.json()}")
        drive = resp.json()
        self.drive_id = drive["id"]

    def list_folder(self, folder_path):
        """List files and subfolders."""
        if not folder_path or folder_path.strip() in ["", "/"]:
            url = f"{self.GRAPH_URL}/drives/{self.drive_id}/root/children"
        else:
            # Use path-based addressing
            path = folder_path.strip("/")
            url  = f"{self.GRAPH_URL}/drives/{self.drive_id}/root:/{path}:/children"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        if not resp.ok:
            raise Exception(f"{folder_path}: {resp.status_code} {resp.text[:200]}")
        return resp.json().get("value", [])

    def download_file(self, file_path):
        """Download file and return bytes."""
        path = file_path.strip("/")
        url  = f"{self.GRAPH_URL}/drives/{self.drive_id}/root:/{path}:/content"
        resp = requests.get(url, headers=self._headers(), timeout=60, allow_redirects=True)
        if not resp.ok:
            raise Exception(f"Download failed ({resp.status_code}): {file_path}")
        return resp.content

    def download_to_temp(self, file_path):
        content  = self.download_file(file_path)
        filename = Path(file_path).name
        tmp_path = os.path.join(tempfile.gettempdir(), f"nen_sp_{filename}")
        with open(tmp_path, "wb") as f:
            f.write(content)
        return tmp_path

    def list_subfolders(self, folder_path):
        items = self.list_folder(folder_path)
        return [i["name"] for i in items if "folder" in i]

    def list_files(self, folder_path):
        items = self.list_folder(folder_path)
        return [i["name"] for i in items if "file" in i]

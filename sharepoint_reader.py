"""
SharePoint reader using Microsoft Graph API - Sites approach
More reliable than user drive approach for app-only auth
"""
import os, tempfile, requests
from pathlib import Path

class SharePointReader:
    GRAPH_URL  = "https://graph.microsoft.com/v1.0"
    SP_HOST    = "wadhwanifoundation-my.sharepoint.com"
    SP_SITE    = "/personal/meenakshi_singh_wadhwanifoundation_org"

    def __init__(self, client_id, tenant_id, client_secret):
        self.client_id     = client_id
        self.tenant_id     = tenant_id
        self.client_secret = client_secret
        self.token         = None
        self.drive_id      = None
        self._authenticate()
        self._get_drive_id()

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
            error_detail = resp.json() if resp.content else resp.status_code
            raise Exception(f"Auth failed ({resp.status_code}): {error_detail}")
        resp.raise_for_status()
        self.token = resp.json()["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def _get_drive_id(self):
        """Get drive ID via Sites API — works with app-only auth."""
        # Get site ID first
        url  = f"{self.GRAPH_URL}/sites/{self.SP_HOST}:{self.SP_SITE}"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        site_id = resp.json()["id"]

        # Get default drive
        url  = f"{self.GRAPH_URL}/sites/{site_id}/drive"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        self.drive_id = resp.json()["id"]
        self.site_id  = site_id

    def list_folder(self, folder_path):
        """List files and subfolders."""
        path_enc = folder_path.strip("/")
        url  = f"{self.GRAPH_URL}/drives/{self.drive_id}/root:/{path_enc}:/children"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        return resp.json().get("value", [])

    def download_file(self, file_path):
        """Download file and return bytes."""
        path_enc = file_path.strip("/")
        url  = f"{self.GRAPH_URL}/drives/{self.drive_id}/root:/{path_enc}:/content"
        resp = requests.get(url, headers=self._headers(), timeout=60)
        resp.raise_for_status()
        return resp.content

    def download_to_temp(self, file_path):
        """Download file to temp and return local path."""
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

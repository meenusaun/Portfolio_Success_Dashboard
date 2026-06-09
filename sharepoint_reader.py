"""
OneDrive reader using Microsoft Graph API
Uses SharePoint personal site path directly
"""
import os, tempfile, requests
from pathlib import Path

class SharePointReader:
    GRAPH_URL  = "https://graph.microsoft.com/v1.0"
    # Personal OneDrive site
    SP_HOSTNAME = "wadhwanifoundation-my.sharepoint.com"
    SP_SITEPATH = "/personal/meenakshi_singh_wadhwanifoundation_org"

    def __init__(self, client_id, tenant_id, client_secret):
        self.client_id     = client_id
        self.tenant_id     = tenant_id
        self.client_secret = client_secret
        self.token         = None
        self.drive_id      = None
        self._authenticate()
        self._get_drive()

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

    def _get_drive(self):
        """Get drive via SharePoint personal site."""
        # Get site ID using hostname:sitepath format
        url  = f"{self.GRAPH_URL}/sites/{self.SP_HOSTNAME}:{self.SP_SITEPATH}"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        if not resp.ok:
            raise Exception(f"Site lookup failed ({resp.status_code}): {resp.json()}")
        site_id = resp.json()["id"]

        # Get the default document library drive
        url2 = f"{self.GRAPH_URL}/sites/{site_id}/drive"
        resp2 = requests.get(url2, headers=self._headers(), timeout=30)
        if not resp2.ok:
            raise Exception(f"Drive lookup failed ({resp2.status_code}): {resp2.json()}")
        self.drive_id = resp2.json()["id"]

    def list_folder(self, folder_path):
        """List files and subfolders."""
        if not folder_path or folder_path.strip() in ["", "/"]:
            url = f"{self.GRAPH_URL}/drives/{self.drive_id}/root/children"
        else:
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
        resp = requests.get(url, headers=self._headers(), timeout=120, allow_redirects=True)
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

    def upload_file(self, file_path, content_bytes):
        """Upload bytes to SharePoint at file_path."""
        path = file_path.strip("/")
        url  = f"{self.GRAPH_URL}/drives/{self.drive_id}/root:/{path}:/content"
        resp = requests.put(url, headers={**self._headers(),
                            "Content-Type": "application/octet-stream"},
                            data=content_bytes, timeout=60)
        if not resp.ok:
            raise Exception(f"Upload failed ({resp.status_code}): {file_path}")
        return resp.json()

    def upload_json(self, file_path, data):
        """Upload a Python dict as JSON to SharePoint."""
        import json
        content = json.dumps(data, indent=2, default=str).encode("utf-8")
        return self.upload_file(file_path, content)

    def download_json(self, file_path):
        """Download and parse a JSON file from SharePoint."""
        import json
        content = self.download_file(file_path)
        return json.loads(content.decode("utf-8"))

    def file_exists(self, file_path):
        """Check if a file exists."""
        try:
            path = file_path.strip("/")
            url  = f"{self.GRAPH_URL}/drives/{self.drive_id}/root:/{path}"
            resp = requests.get(url, headers=self._headers(), timeout=15)
            return resp.ok
        except:
            return False

    def get_file_modified_time(self, file_path):
        """Get last modified datetime string for a file."""
        try:
            path = file_path.strip("/")
            url  = f"{self.GRAPH_URL}/drives/{self.drive_id}/root:/{path}"
            resp = requests.get(url, headers=self._headers(), timeout=15)
            if resp.ok:
                return resp.json().get("lastModifiedDateTime","")
        except:
            pass
        return ""

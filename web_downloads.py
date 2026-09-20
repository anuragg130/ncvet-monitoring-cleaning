"""Restrict hosted workbook downloads to Google HTTPS endpoints, including redirects."""
from urllib.parse import urlparse
from requests.adapters import HTTPAdapter


def validate_url(url):
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()
    allowed = host in {"drive.google.com", "docs.google.com", "accounts.google.com", "drive.usercontent.google.com"} or any(
        host == domain or host.endswith("." + domain)
        for domain in ("googleusercontent.com", "googleapis.com")
    )
    if parsed.scheme != "https" or not allowed or parsed.username or parsed.password or parsed.port not in (None, 443):
        raise ValueError("The web app accepts Google Drive or Google Sheets HTTPS links only.")


class GoogleDownloadAdapter(HTTPAdapter):
    def send(self, request, **kwargs):
        validate_url(request.url)
        kwargs["verify"] = True
        return super().send(request, **kwargs)

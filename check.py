from google.oauth2 import service_account
from google.auth.transport.requests import Request

creds = service_account.Credentials.from_service_account_file(
    "service_account.json",
    scopes=["https://www.googleapis.com/auth/drive.readonly"]
)

try:
    creds.refresh(Request())
    print("SUCCESS: Google authentication working!")
except Exception as e:
    print("AUTHENTICATION FAILED:", e)
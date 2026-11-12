# Student Portfolio Photo Checker

A Streamlit dashboard that checks Google Slides presentations in a Drive folder for student photos (Weeks 1–5) and a self-reflection video ("The journey in their own words"), then writes a clear report to a Google Sheet.

## Features
- No file downloads — uses the Google Slides API
- Checks Weeks 1–5 with precise coordinate boxes
- Detects video elements on the Journey / self-reflection slide
- Clean Streamlit interface (enter Folder ID + Sheet ID)
- Secure credential handling via Streamlit secrets

### 2. Credentials (recommended: Streamlit secrets)
Create `.streamlit/secrets.toml` (or add the secrets in Streamlit Cloud):

```toml
[google_auth]
type = "service_account"
project_id = "your-project-id"
private_key_id = "xxxxxxxx"
private_key = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email = "your-service-account@your-project-id.iam.gserviceaccount.com"
client_id = "123456789012345678901"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/your-service-account%40your-project-id.iam.gserviceaccount.com"
universe_domain = "googleapis.com"
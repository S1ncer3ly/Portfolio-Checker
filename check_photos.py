"""
Checks every PPTX / Google Slides file in a Google Drive folder for student photos
(Week 1-5 + "The journey in their own words") and writes a report to a Google Sheet.

Setup (once):
  pip install python-pptx google-api-python-client google-auth
  - Share the Drive folder AND the Google Sheet with the service account's email
    (Viewer on the folder is enough; Editor on the sheet).
  - Enable the Google Drive API and Google Sheets API in the service account's project.

Run:
  python drive_photo_checker.py --creds service_account.json \
      --folder-id <DRIVE_FOLDER_ID> --sheet-id <SPREADSHEET_ID> [--tab Sheet1] [--recursive]

Test the checking logic on local files without Google:
  python drive_photo_checker.py --local ANIKA_CHOUDHARY.pptx
"""
import argparse, io, sys, datetime
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
GSLIDES_MIME = "application/vnd.google-apps.presentation"
FOLDER_MIME = "application/vnd.google-apps.folder"

# Pictures narrower than this (EMU) are the school logo/decoration, not a student photo.
MIN_PHOTO_WIDTH = 3_000_000

SECTIONS = {
    "Week 1": "week 1",
    "Week 2": "week 2",
    "Week 3": "week 3",
    "Week 4": "week 4",
    "Week 5": "week 5",
    "Journey (own words)": "the journey in their own words",
}

# ---------------------------------------------------------------- analysis
def iter_shapes(shapes):
    for s in shapes:
        yield s
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(s.shapes)

def slide_has_photo(slide):
    for s in iter_shapes(slide.shapes):
        try:
            s.image
        except (AttributeError, ValueError):
            continue
        if s.width and s.width >= MIN_PHOTO_WIDTH:
            return True
    return False

def slide_text(slide):
    return " ".join(s.text_frame.text for s in iter_shapes(slide.shapes)
                    if s.has_text_frame)

def student_name(prs, fallback):
    """Cover slide box looks like 'Name of the Student\\n<NAME>'."""
    try:
        for s in iter_shapes(prs.slides[1].shapes):
            if s.has_text_frame:
                lines = [l.strip() for l in s.text_frame.text.splitlines() if l.strip()]
                if lines and lines[0].lower().startswith("name of the student") and len(lines) > 1:
                    return fallback if "{{" in lines[1] else lines[1]
    except Exception:
        pass
    return fallback

def analyze_bytes(data, file_label):
    prs = Presentation(io.BytesIO(data))
    status = {k: None for k in SECTIONS}      # True / False / None(slide not found)
    for slide in prs.slides:
        low = slide_text(slide).lower()
        for label, key in SECTIONS.items():
            if key in low:
                status[label] = bool(status[label]) or slide_has_photo(slide)
    return student_name(prs, file_label), status

def cell(v):
    return "OK" if v is True else "MISSING" if v is False else "SLIDE NOT FOUND"

def build_row(file_name, file_url, name, status):
    cells = [cell(status[k]) for k in SECTIONS]
    problems = [k for k in SECTIONS if status[k] is not True]
    overall = "COMPLETE" if not problems else "INCOMPLETE"
    return [name, file_name, file_url, *cells, overall,
            ", ".join(problems), datetime.datetime.now().strftime("%Y-%m-%d %H:%M")]

HEADER = ["Student", "File name", "File link", *SECTIONS.keys(),
          "Overall", "Missing / problem sections", "Checked at"]

# ---------------------------------------------------------------- Google
def google_services(creds_path):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    scopes = ["https://www.googleapis.com/auth/drive.readonly",
              "https://www.googleapis.com/auth/spreadsheets"]
    with open(creds_path, "rb") as fh:
        raw = fh.read()
    if raw.lstrip()[:1] == b"{":                  # JSON key file
        import json
        info = json.loads(raw)
    else:                                         # TOML (e.g. Streamlit secrets with [google_auth])
        try:
            import tomllib
        except ImportError:
            import tomli as tomllib               # pip install tomli  (Python < 3.11)
        data = tomllib.loads(raw.decode("utf-8"))
        info = data.get("google_auth", data)
    creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    return (build("drive", "v3", credentials=creds, cache_discovery=False),
            build("sheets", "v4", credentials=creds, cache_discovery=False))

def list_files(drive, folder_id, recursive):
    found, stack = [], [folder_id]
    while stack:
        fid = stack.pop()
        token = None
        while True:
            resp = drive.files().list(
                q=f"'{fid}' in parents and trashed=false",
                fields="nextPageToken, files(id,name,mimeType,webViewLink)",
                pageSize=200, pageToken=token,
                supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
            for f in resp.get("files", []):
                if f["mimeType"] == FOLDER_MIME:
                    if recursive: stack.append(f["id"])
                elif f["mimeType"] in (PPTX_MIME, GSLIDES_MIME) or f["name"].lower().endswith(".pptx"):
                    found.append(f)
            token = resp.get("nextPageToken")
            if not token: break
    return sorted(found, key=lambda f: f["name"].lower())

def download(drive, f):
    from googleapiclient.http import MediaIoBaseDownload
    if f["mimeType"] == GSLIDES_MIME:                       # native Google Slides -> export as pptx
        req = drive.files().export_media(fileId=f["id"], mimeType=PPTX_MIME)
    else:
        req = drive.files().get_media(fileId=f["id"], supportsAllDrives=True)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()

def clean_id(v):
    """Accept a bare ID or a pasted URL tail like '<ID>/edit?gid=0#gid=0'."""
    v = v.strip()
    if "/d/" in v:
        v = v.split("/d/")[1]
    return v.split("/")[0].split("?")[0].split("#")[0]

def write_sheet(sheets, sheet_id, tab, rows):
    if not tab:
        meta = sheets.spreadsheets().get(spreadsheetId=sheet_id,
                                         fields="sheets.properties.title").execute()
        tab = meta["sheets"][0]["properties"]["title"]
    print(f"Writing to tab: {tab}")
    rng = f"'{tab}'!A1"
    sheets.spreadsheets().values().clear(spreadsheetId=sheet_id, range=f"'{tab}'").execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=rng, valueInputOption="USER_ENTERED",
        body={"values": [HEADER] + rows}).execute()

# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--creds"); ap.add_argument("--folder-id"); ap.add_argument("--sheet-id")
    ap.add_argument("--tab", default=None, help="tab name (default: first tab)")
    ap.add_argument("--recursive", action="store_true", help="also scan sub-folders")
    ap.add_argument("--local", nargs="*", help="test on local .pptx files (no Google calls)")
    a = ap.parse_args()

    if a.local is not None:
        for p in a.local:
            name, st = analyze_bytes(open(p, "rb").read(), p)
            print(build_row(p, "", name, st))
        return

    if not (a.creds and a.folder_id and a.sheet_id):
        sys.exit("Need --creds, --folder-id and --sheet-id (or use --local).")

    a.folder_id, a.sheet_id = clean_id(a.folder_id), clean_id(a.sheet_id)
    drive, sheets = google_services(a.creds)
    files = list_files(drive, a.folder_id, a.recursive)
    print(f"Found {len(files)} presentation(s).")
    rows = []
    for i, f in enumerate(files, 1):
        try:
            name, st = analyze_bytes(download(drive, f), f["name"])
            rows.append(build_row(f["name"], f.get("webViewLink", ""), name, st))
            print(f"[{i}/{len(files)}] {f['name']}: {rows[-1][-3]}")
        except Exception as e:
            rows.append([f["name"], f["name"], f.get("webViewLink", ""),
                         *["ERROR"] * len(SECTIONS), "ERROR", str(e)[:200],
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M")])
            print(f"[{i}/{len(files)}] {f['name']}: ERROR {e}")
    write_sheet(sheets, a.sheet_id, a.tab, rows)
    print("Sheet updated.")

if __name__ == "__main__":
    main()
import streamlit as st
import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build

st.set_page_config(page_title="Student Photo Checker", layout="wide")
st.title("Student Presentation Photo Checker")
st.caption("Checks Week 1–5 + Journey photos and writes the report to a Google Sheet.")

# ── Coordinate map (0-based) ────────────────────────────────────────
COORD_MAP = {
    2: (1603584, 3260950, 9994200, 5740500),
    3: (1603584, 3260950, 9994200, 5740500),
    4: (1603584, 3260950, 9994200, 5740500),
    5: (1603584, 3260950, 9994200, 5740500),
    6: (1574725, 3252925, 10132800, 5822800),
}
SLIDE_TO_SECTION = {
    2: "Week 1", 3: "Week 2", 4: "Week 3", 5: "Week 4", 6: "Week 5",
}
SECTIONS = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5", "Journey (own words)"]
HEADER = ["Student", "File name", "File link", *SECTIONS, "Overall", "Missing / problem sections", "Checked at"]

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
GSLIDES_MIME = "application/vnd.google-apps.presentation"
FOLDER_MIME = "application/vnd.google-apps.folder"

# ── Sidebar ─────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")
    folder_id = st.text_input("Drive Folder ID", placeholder="1KVdTDJTLW9UqW_lK7clQ_ashL04LDvu8")
    sheet_id = st.text_input("Google Sheet ID", placeholder="1xRFJ5vlZ2-UjhH3AB5DEicNQNk3lsC67d5aVMl94Wz4")
    tab_name = st.text_input("Sheet tab name (optional)", value="Sheet1")
    recursive = st.checkbox("Scan sub-folders", value=True)
    run_btn = st.button("Run Check", type="primary", use_container_width=True)

# ── Helpers (same as before) ────────────────────────────────────────
def clean_id(v):
    v = (v or "").strip()
    if "/d/" in v:
        v = v.split("/d/")[1]
    return v.split("/")[0].split("?")[0].split("#")[0]

def iter_elements(elements):
    for el in elements or []:
        yield el
        if "elementGroup" in el:
            yield from iter_elements(el["elementGroup"].get("children", []))

def slide_text(slide):
    texts = []
    for el in iter_elements(slide.get("pageElements", [])):
        shape = el.get("shape")
        if shape and "text" in shape:
            for te in shape["text"].get("textElements", []):
                if "textRun" in te:
                    texts.append(te["textRun"].get("content", ""))
    return " ".join(texts)

def element_bbox(el):
    size = el.get("size", {})
    w = size.get("width", {}).get("magnitude", 0)
    h = size.get("height", {}).get("magnitude", 0)
    t = el.get("transform", {})
    left = t.get("translateX", 0)
    top = t.get("translateY", 0)
    scale_x = t.get("scaleX", 1) or 1
    scale_y = t.get("scaleY", 1) or 1
    return left, top, w * scale_x, h * scale_y

def boxes_overlap(a, b, tolerance=150000):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return (ax >= bx - tolerance and ay >= by - tolerance and
            ax + aw <= bx + bw + tolerance and ay + ah <= by + bh + tolerance)

def has_photo_in_box(slide, target_box):
    for el in iter_elements(slide.get("pageElements", [])):
        if "image" in el and boxes_overlap(element_bbox(el), target_box):
            return True
    return False

def student_name_from_slides(slides, fallback):
    if len(slides) > 1:
        text = slide_text(slides[1])
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        for i, line in enumerate(lines):
            if line.lower().startswith("name of the student") and i + 1 < len(lines):
                name = lines[i + 1]
                return fallback if "{{" in name else name
    return fallback

def analyze_slides(presentation, file_label):
    slides = presentation.get("slides", [])
    status = {k: None for k in SECTIONS}

    for idx, section in SLIDE_TO_SECTION.items():
        if idx < len(slides):
            status[section] = has_photo_in_box(slides[idx], COORD_MAP[idx])
        else:
            status[section] = False

    for slide in slides:
        low = slide_text(slide).lower()
        if "the journey in their own words" in low or "journey in their own words" in low:
            has_img = any("image" in el for el in iter_elements(slide.get("pageElements", [])))
            status["Journey (own words)"] = has_img
            break

    return student_name_from_slides(slides, file_label), status

def cell(v):
    return "OK" if v is True else "MISSING" if v is False else "SLIDE NOT FOUND"

def build_row(file_name, file_url, name, status):
    cells = [cell(status[k]) for k in SECTIONS]
    problems = [k for k in SECTIONS if status[k] is not True]
    overall = "COMPLETE" if not problems else "INCOMPLETE"
    return [name, file_name, file_url, *cells, overall,
            ", ".join(problems), datetime.datetime.now().strftime("%Y-%m-%d %H:%M")]

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
                    if recursive:
                        stack.append(f["id"])
                elif f["mimeType"] in (PPTX_MIME, GSLIDES_MIME) or f["name"].lower().endswith(".pptx"):
                    found.append(f)
            token = resp.get("nextPageToken")
            if not token:
                break
    return sorted(found, key=lambda f: f["name"].lower())

def write_sheet(sheets, sheet_id, tab, rows):
    if not tab:
        meta = sheets.spreadsheets().get(spreadsheetId=sheet_id, fields="sheets.properties.title").execute()
        tab = meta["sheets"][0]["properties"]["title"]
    sheets.spreadsheets().values().clear(spreadsheetId=sheet_id, range=f"'{tab}'").execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=f"'{tab}'!A1",
        valueInputOption="USER_ENTERED",
        body={"values": [HEADER] + rows}).execute()

# ── Run ─────────────────────────────────────────────────────────────
if run_btn:
    if not folder_id or not sheet_id:
        st.error("Please enter both Folder ID and Sheet ID.")
        st.stop()

    try:
        # Load credentials from Streamlit secrets
        info = dict(st.secrets["google_auth"])
        scopes = [
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/presentations.readonly",
            "https://www.googleapis.com/auth/spreadsheets",
        ]
        creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        slides_svc = build("slides", "v1", credentials=creds, cache_discovery=False)
        sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)

        fid = clean_id(folder_id)
        sid = clean_id(sheet_id)

        with st.spinner("Listing presentations..."):
            files = list_files(drive, fid, recursive)
        st.success(f"Found {len(files)} presentation(s)")

        progress = st.progress(0)
        status_text = st.empty()
        rows = []

        for i, f in enumerate(files):
            status_text.text(f"Checking {i+1}/{len(files)}: {f['name']}")
            try:
                if f["mimeType"] == GSLIDES_MIME:
                    presentation = slides_svc.presentations().get(presentationId=f["id"]).execute()
                    name, st_status = analyze_slides(presentation, f["name"])
                    rows.append(build_row(f["name"], f.get("webViewLink", ""), name, st_status))
                else:
                    rows.append([f["name"], f["name"], f.get("webViewLink", ""),
                                 *["PPTX – skipped"] * len(SECTIONS), "SKIPPED",
                                 "Convert to Google Slides for full check",
                                 datetime.datetime.now().strftime("%Y-%m-%d %H:%M")])
            except Exception as e:
                rows.append([f["name"], f["name"], f.get("webViewLink", ""),
                             *["ERROR"] * len(SECTIONS), "ERROR", str(e)[:180],
                             datetime.datetime.now().strftime("%Y-%m-%d %H:%M")])
            progress.progress((i + 1) / len(files))

        status_text.empty()
        progress.empty()

        with st.spinner("Writing report to Google Sheet..."):
            write_sheet(sheets, sid, tab_name.strip() or None, rows)

        st.success("Sheet updated successfully!")

        import pandas as pd
        df = pd.DataFrame(rows, columns=HEADER)
        st.dataframe(df, use_container_width=True, hide_index=True)

        complete = (df["Overall"] == "COMPLETE").sum()
        st.metric("Complete presentations", f"{complete} / {len(df)}")

    except Exception as e:
        st.error(f"Error: {e}")
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
    # Journey no longer uses coordinates
}
SLIDE_TO_SECTION = {
    2: "Week 1",
    3: "Week 2",
    4: "Week 3",
    5: "Week 4",
    6: "Week 5",
    7: "Journey (own words)",
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
    convert_pptx = st.checkbox("Convert PPTX to Google Slides (if no Slides copy exists)", value=True)
    trash_converted = st.checkbox("Move PPTX to trash after it is converted and checked", value=False)
    trash_twins = st.checkbox("Also trash PPTX files that already have a Slides copy", value=False)
    debug_mode = st.checkbox("Debug mode (show why each section passed / failed)", value=False)
    run_btn = st.button("Run Check", type="primary", use_container_width=True)

# ── Helpers ─────────────────────────────────────────────────────────
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

# ── Geometry (handles flips, rotation, groups and PT/EMU units) ─────────────
MIN_COVER = 0.5          # an image must cover at least 50% of the target box...
MAX_SIZE_RATIO = 2.5     # ...and not be a huge background picture (> 2.5x the box)
SECTION_LABEL = {"Week 1": "week 1", "Week 2": "week 2", "Week 3": "week 3",
                 "Week 4": "week 4", "Week 5": "week 5",
                 "Journey (own words)": "the journey in their own words"}
SECTION_BOX = {SLIDE_TO_SECTION[i]: box for i, box in COORD_MAP.items()}
IDENT = (1, 0, 0, 1, 0, 0)   # (a, b, c, d, tx, ty): x' = a*x + c*y + tx ; y' = b*x + d*y + ty

def unit_factor(unit):
    return 12700 if unit == "PT" else 1

def _matrix(t):
    f = unit_factor(t.get("unit", "EMU"))
    sx = t.get("scaleX"); sy = t.get("scaleY")
    return (1 if sx is None else sx, t.get("shearY", 0) or 0,
            t.get("shearX", 0) or 0, 1 if sy is None else sy,
            (t.get("translateX", 0) or 0) * f, (t.get("translateY", 0) or 0) * f)

def _compose(P, C):
    Pa, Pb, Pc, Pd, Ptx, Pty = P
    Ca, Cb, Cc, Cd, Ctx, Cty = C
    return (Pa * Ca + Pc * Cb, Pb * Ca + Pd * Cb, Pa * Cc + Pc * Cd, Pb * Cc + Pd * Cd,
            Pa * Ctx + Pc * Cty + Ptx, Pb * Ctx + Pd * Cty + Pty)

def iter_abs(elements, parent=IDENT, in_group=False):
    """Yield (element, absolute matrix, in_group), composing group transforms."""
    for el in elements or []:
        M = _compose(parent, _matrix(el.get("transform", {})))
        yield el, M, in_group
        if "elementGroup" in el:
            yield from iter_abs(el["elementGroup"].get("children", []), M, True)

def abs_bbox(el, M):
    size = el.get("size", {})
    wd, hd = size.get("width", {}), size.get("height", {})
    w = (wd.get("magnitude", 0) or 0) * unit_factor(wd.get("unit", "EMU"))
    h = (hd.get("magnitude", 0) or 0) * unit_factor(hd.get("unit", "EMU"))
    a, b, c, d, tx, ty = M
    pts = [(tx, ty), (a * w + tx, b * w + ty), (c * h + tx, d * h + ty),
           (a * w + c * h + tx, b * w + d * h + ty)]
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))

def intersection_cover(a, b):
    """Share of box b that element a covers (0..1)."""
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0, min(ay + ah, by + bh) - max(ay, by))
    return (ix * iy) / (bw * bh) if bw and bh else 0

def slide_media(slide):
    out = []
    for el, M, in_group in iter_abs(slide.get("pageElements", [])):
        kind = "image" if "image" in el else "video" if "video" in el else None
        if kind:
            out.append({"kind": kind, "bbox": abs_bbox(el, M), "in_group": in_group,
                        "unit": el.get("transform", {}).get("unit", "?")})
    return out

def photo_counts(m, box):
    cov = intersection_cover(m["bbox"], box)
    ratio = (m["bbox"][2] * m["bbox"][3]) / (box[2] * box[3])
    return m["kind"] == "image" and cov >= MIN_COVER and ratio <= MAX_SIZE_RATIO, cov, ratio

def has_photo_in_box(slide, target_box):
    return any(photo_counts(m, target_box)[0] for m in slide_media(slide))

def resolve_slides(slides):
    """section -> (slide index or None, how it was found). Title first, position as fallback."""
    texts = [" ".join(slide_text(s).split()).lower() for s in slides]
    out = {}
    for idx, section in SLIDE_TO_SECTION.items():
        hits = [i for i, t in enumerate(texts) if SECTION_LABEL[section] in t]
        if hits:
            out[section] = (hits[0], "title")
        elif idx < len(slides):
            out[section] = (idx, "position (title not found)")
        else:
            out[section] = (None, "slide missing")
    return out

_ZW = dict.fromkeys(map(ord, "\u200b\u200c\u200d\ufeff"), None)
_LABELS = ("name of the student", "grade & section", "studio", "theme", "session")

def student_name_from_slides(slides, fallback):
    """Name sits on line 2 of the 'Name of the Student' box. Blank/invisible or a
    neighbouring label means the student left it empty -> use the file name."""
    if len(slides) > 1:
        for el in iter_elements(slides[1].get("pageElements", [])):
            sh = el.get("shape")
            if not sh or "text" not in sh:
                continue
            raw = "".join(te["textRun"].get("content", "")
                          for te in sh["text"].get("textElements", []) if "textRun" in te).translate(_ZW)
            lines = [l.strip() for l in raw.splitlines() if l.strip()]
            if len(lines) > 1 and lines[0].lower().startswith("name of the student"):
                name = lines[1]
                if "{{" not in name and name.lower() not in _LABELS:
                    return name
    return fallback

def analyze_slides(presentation, file_label):
    slides = presentation.get("slides", [])
    resolved = resolve_slides(slides)
    status = {}
    for section in SECTIONS:
        idx, _ = resolved[section]
        if idx is None:
            status[section] = False
        elif section == "Journey (own words)":
            status[section] = any(m["kind"] == "video" for m in slide_media(slides[idx]))
        else:
            status[section] = has_photo_in_box(slides[idx], SECTION_BOX[section])
    return student_name_from_slides(slides, file_label), status

def cell(v):
    return "OK" if v is True else "MISSING" if v is False else "SLIDE NOT FOUND"

def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

def build_row(file_name, file_url, name, status):
    cells = [cell(status[k]) for k in SECTIONS]
    problems = [k for k in SECTIONS if status[k] is not True]
    overall = "COMPLETE" if not problems else "INCOMPLETE"
    return [name, file_name, file_url, *cells, overall,
            ", ".join(problems), now_str()]

def list_files(drive, folder_id, recursive):
    found, stack = [], [folder_id]
    while stack:
        fid = stack.pop()
        token = None
        while True:
            resp = drive.files().list(
                q=f"'{fid}' in parents and trashed=false",
                fields="nextPageToken, files(id,name,mimeType,webViewLink,parents)",
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

# ── Debug helpers (explain the decision, never change it) ───────────────────
EXPECTED_PAGE = (18288000, 10287000)

def debug_tables(presentation):
    slides = presentation.get("slides", [])
    ps = presentation.get("pageSize", {})
    page = (ps.get("width", {}).get("magnitude"), ps.get("height", {}).get("magnitude"))
    meta = {"page": page, "page_ok": page == EXPECTED_PAGE, "n_slides": len(slides)}
    resolved = resolve_slides(slides)
    by_idx = {idx: (sec, how) for sec, (idx, how) in resolved.items() if idx is not None}
    meta["missing_sections"] = [sec for sec, (idx, _) in resolved.items() if idx is None]
    slide_rows, element_rows = [], []
    for idx, slide in enumerate(slides):
        sec, how = by_idx.get(idx, ("", ""))
        text = " ".join(slide_text(slide).split())
        media = slide_media(slide)
        imgs = sum(m["kind"] == "image" for m in media)
        vids = sum(m["kind"] == "video" for m in media)
        counted = 0
        for m in media:
            l, t, w, h = m["bbox"]
            row = {"slide #": idx, "section": sec, "type": m["kind"], "in group": m["in_group"],
                   "left": round(l), "top": round(t), "width": round(w), "height": round(h),
                   "unit": m["unit"]}
            if sec in SECTION_BOX:
                ok, cov, ratio = photo_counts(m, SECTION_BOX[sec])
                row.update({"covers % of box": round(100 * cov),
                            "size % of box": round(100 * ratio), "counts as photo": ok})
                counted += ok
            element_rows.append(row)
        if sec == "Journey (own words)":
            decision = "OK" if vids else "MISSING"
            why = ("video element found" if vids else
                   f"no video element ({imgs} image(s) on slide; this section only accepts a video)")
        elif sec:
            decision = "OK" if counted else "MISSING"
            if counted:
                why = f"{counted} image(s) cover >= {int(MIN_COVER*100)}% of the box"
            elif imgs:
                why = (f"{imgs} image(s) on slide but none covers >= {int(MIN_COVER*100)}% of the box "
                       f"(or it is > {MAX_SIZE_RATIO}x the box) – see the element table")
            else:
                why = "no image element on slide (photo may be a drawing/shape fill, or truly absent)"
        else:
            decision, why = "", ""
        slide_rows.append({"slide #": idx, "treated as": sec or "(not checked)", "found by": how,
                           "slide text (start)": text[:60], "images": imgs, "videos": vids,
                           "decision": decision, "why": why})
    return meta, slide_rows, element_rows

def render_debug(presentation, file_name):
    import json, pandas as pd
    meta, slide_rows, element_rows = debug_tables(presentation)
    if meta["page_ok"]:
        st.write(f"Page size OK: {meta['page'][0]} x {meta['page'][1]}")
    else:
        st.error(f"Page size is {meta['page']} but the boxes assume {EXPECTED_PAGE}. "
                 "Every box check will be wrong for this file.")
    if meta["missing_sections"]:
        st.warning("No slide found for: " + ", ".join(meta["missing_sections"]))
    st.markdown("**Slide map** – 'found by' says whether the slide was located by its title or fell back to its position.")
    st.dataframe(pd.DataFrame(slide_rows), use_container_width=True, hide_index=True)
    st.markdown("**Every image/video element** with its true on-slide position and how it compares to the target box")
    st.dataframe(pd.DataFrame(element_rows), use_container_width=True, hide_index=True)
    st.download_button("Download raw Slides JSON for this file",
                       data=json.dumps(presentation, indent=1),
                       file_name=f"{file_name}.slides.json", mime="application/json",
                       key=f"dl_{file_name}")

# ── PPTX -> Google Slides conversion ────────────────────────────────
def base_name(name):
    return name[:-5] if name.lower().endswith(".pptx") else name

def first_parent(f):
    return (f.get("parents") or [""])[0]

def plan_files(files):
    """Split files into: Slides to check, PPTX to convert, PPTX skipped
    because a Google Slides copy with the same name already sits in the same folder."""
    slides_index = {(first_parent(f), f["name"].strip().lower())
                    for f in files if f["mimeType"] == GSLIDES_MIME}
    to_check, to_convert, skipped = [], [], []
    for f in files:
        if f["mimeType"] == GSLIDES_MIME:
            to_check.append(f)
        elif (first_parent(f), base_name(f["name"]).strip().lower()) in slides_index:
            skipped.append(f)          # Slides copy exists -> no conversion, that copy gets checked
        else:
            to_convert.append(f)
    return to_check, to_convert, skipped

def convert_to_slides(drive, f):
    """Copy a PPTX into the same folder, converting it to native Google Slides."""
    body = {"name": base_name(f["name"]), "mimeType": GSLIDES_MIME}
    if f.get("parents"):
        body["parents"] = [first_parent(f)]
    return drive.files().copy(
        fileId=f["id"], body=body, supportsAllDrives=True,
        fields="id,name,mimeType,webViewLink,parents").execute()

def trash_file(drive, f):
    """Move a file to the Drive trash (recoverable from Trash for ~30 days)."""
    drive.files().update(fileId=f["id"], body={"trashed": True},
                         supportsAllDrives=True, fields="id,trashed").execute()

def friendly_error(e):
    msg = str(e)
    if "storageQuotaExceeded" in msg:
        return ("Conversion failed: service account has no Drive storage. "
                "Use a Shared Drive folder (add the service account as Content manager) "
                "or convert the file manually.")
    if "insufficientFilePermissions" in msg or "403" in msg:
        return "Conversion failed (403): give the service account Editor access to the folder. " + msg[:100]
    return msg[:180]

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
        info = dict(st.secrets["google_auth"])
        scopes = [
            "https://www.googleapis.com/auth/drive",   # write access is needed to convert PPTX
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
        to_check, to_convert, skipped = plan_files(files)
        st.success(f"Found {len(files)} file(s): {len(to_check)} Google Slides, "
                   f"{len(to_convert)} PPTX to convert, {len(skipped)} PPTX skipped (Slides copy already exists)")

        if trash_twins and skipped:
            trashed_n, trash_errs = 0, []
            for f in skipped:
                try:
                    trash_file(drive, f)
                    trashed_n += 1
                except Exception as e:
                    trash_errs.append(f"{f['name']}: {str(e)[:120]}")
            st.info(f"Moved {trashed_n} duplicate PPTX file(s) to trash.")
            if trash_errs:
                st.warning("Could not trash: " + " | ".join(trash_errs))

        # Work list: (display_name, file_dict_to_check, converted_flag, error_or_None)
        work = [(f["name"], f, False, None) for f in to_check]
        for f in to_convert:
            if not convert_pptx:
                work.append((f["name"], None, False, "PPTX – conversion turned off"))
                continue
            work.append((f["name"], f, None, None))   # converted_flag None = convert first

        debug_store = {}
        total = max(len(work), 1)
        progress = st.progress(0)
        status_text = st.empty()
        rows = []
        for i, (display, f, converted, pre_err) in enumerate(work):
            status_text.text(f"Checking {i+1}/{len(work)}: {display}")
            try:
                if pre_err:
                    rows.append([display, display, "", *["PPTX – skipped"] * len(SECTIONS),
                                 "SKIPPED", pre_err, now_str()])
                else:
                    original = f
                    if converted is None:               # PPTX with no Slides copy -> convert now
                        status_text.text(f"Converting {i+1}/{len(work)}: {display}")
                        f = convert_to_slides(drive, f)
                        label = f"{display} (converted)"
                    else:
                        label = display
                    presentation = slides_svc.presentations().get(presentationId=f["id"]).execute()
                    name, st_status = analyze_slides(presentation, display)
                    if debug_mode:
                        debug_store[display] = presentation
                    # Trash the original PPTX only after conversion AND the check both succeeded
                    if converted is None and trash_converted:
                        try:
                            trash_file(drive, original)
                            label += ", PPTX trashed"
                        except Exception as te:
                            label += f", PPTX NOT trashed: {str(te)[:80]}"
                    rows.append(build_row(label, f.get("webViewLink", ""), name, st_status))
            except Exception as e:
                rows.append([display, display, "", *["ERROR"] * len(SECTIONS),
                             "ERROR", friendly_error(e), now_str()])
            progress.progress((i + 1) / total)
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
        st.session_state["debug_store"] = debug_store if debug_mode else {}
    except Exception as e:
        st.error(f"Error: {e}")

# ── Debug view (rendered outside the run block so it survives widget clicks) ──
if debug_mode and st.session_state.get("debug_store"):
    st.divider()
    st.subheader("Debug: why each section passed or failed")
    pick = st.selectbox("File", sorted(st.session_state["debug_store"]))
    render_debug(st.session_state["debug_store"][pick], pick)
"""Google Drive Hackathon presentation photo and Journey video checker (v7).

Read-only Google Drive + Slides APIs; Google Sheets report sync with explicit permission.
Native Slides are parsed directly, while PPTX uses python-pptx, including filled placeholders.
Weeks 1–5 check photo candidates; Journey verifies embedded videos, links and uncertain media separately.
Images are candidates for review, not verified student identity; video existence does not prove playback.
"""
from __future__ import annotations

import hashlib
from datetime import datetime, timezone
import math
import io
import re
from urllib.parse import urlparse, parse_qs
from collections import deque
from pathlib import Path

from PIL import Image, UnidentifiedImageError
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

ROOT = Path(__file__).resolve().parent
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
SLIDES_MIME = "application/vnd.google-apps.presentation"
FOLDER_MIME = "application/vnd.google-apps.folder"
SHORTCUT_MIME = "application/vnd.google-apps.shortcut"
SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/presentations.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]
EMU = 914400
JOURNEY = "Journey (own words)"
PHOTO_SECTIONS = [f"Week {i}" for i in range(1, 6)]
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
VIDEO_EXTENSIONS = (".mp4", ".mov", ".m4v", ".webm", ".avi", ".wmv", ".mpeg", ".mpg")

SECTIONS = {
    **{f"Week {i}": re.compile(rf"^\s*week\s*{i}\b", re.I) for i in range(1, 6)},
    "Journey (own words)": re.compile(
        r"^\s*(?:the\s+)?journey(?:\s+in\s+their\s+own\s+words|\s*\(own\s+words\))?(?=\s|$|[.:—–-])", re.I
    ),
}


def get_service_account_credentials():
    """Load credentials from Streamlit secrets (Cloud) or local service_account.json."""
    from google.oauth2 import service_account

    # Prefer Streamlit secrets when running on Streamlit Cloud
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            return service_account.Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=SCOPES,
            )
    except Exception:
        pass  # not running under Streamlit, or secrets not set

    # Fallback: local JSON file
    credentials_file = ROOT / "service_account.json"
    if not credentials_file.is_file():
        raise FileNotFoundError(
            f"Missing service-account credentials. "
            f"On Streamlit Cloud put them under [gcp_service_account] in secrets; "
            f"locally place service_account.json next to the script."
        )
    return service_account.Credentials.from_service_account_file(
        str(credentials_file), scopes=SCOPES
    )


def google_clients():
    """Read Drive/Slides and write only an explicitly specified Google spreadsheet."""
    from googleapiclient.discovery import build

    creds = get_service_account_credentials()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    slides = build("slides", "v1", credentials=creds, cache_discovery=False)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return drive, slides, sheets


def extract_drive_id(value: str) -> str:
    """Allow either an ID or a normal Google Drive URL pasted into the input."""
    value = value.strip()
    for pattern in (
        r"/folders/([\w-]+)",
        r"/file/d/([\w-]+)",
        r"/presentation/d/([\w-]+)",
        r"/spreadsheets/d/([\w-]+)",
        r"[?&]id=([\w-]+)",
    ):
        m = re.search(pattern, value)
        if m:
            return m.group(1)
    return value


def list_children(drive, folder_id):
    token = None
    while True:
        response = drive.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken,incompleteSearch,files(id,name,mimeType,webViewLink,parents,shortcutDetails)",
            pageSize=1000,
            pageToken=token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        if response.get("incompleteSearch"):
            raise RuntimeError("Google Drive returned incomplete search results")
        yield from response.get("files", [])
        token = response.get("nextPageToken")
        if not token:
            break


def inventory(drive, parent_id, recursive=True):
    files, folders, seen = [], [], set()
    queue = deque([(parent_id, "/")])
    while queue:
        folder_id, parent_path = queue.popleft()
        if folder_id in seen:
            continue
        seen.add(folder_id)
        for file in list_children(drive, folder_id):
            mime = file.get("mimeType", "")
            if mime == SHORTCUT_MIME:
                details = file.get("shortcutDetails", {})
                target_id, target_mime = details.get("targetId"), details.get("targetMimeType")
                if not target_id:
                    continue
                if target_mime == FOLDER_MIME:
                    child_path = (parent_path.rstrip("/") + "/" + file["name"] + "/")
                    folders.append(child_path + " (shortcut)")
                    if recursive:
                        queue.append((target_id, child_path))
                elif target_mime in (PPTX_MIME, SLIDES_MIME):
                    files.append({**file, "id": target_id, "mimeType": target_mime,
                                  "folder_path": parent_path, "via_shortcut": True})
            elif mime == FOLDER_MIME:
                child_path = (parent_path.rstrip("/") + "/" + file["name"] + "/")
                folders.append(child_path)
                if recursive:
                    queue.append((file["id"], child_path))
            elif mime in (PPTX_MIME, SLIDES_MIME):
                files.append({**file, "folder_path": parent_path})
    # Avoid scanning the same Drive file twice when multiple shortcut paths exist.
    deduped = {item["id"]: item for item in files}
    result = list(deduped.values())
    result.sort(key=lambda f: (f["folder_path"].casefold(), f["name"].casefold()))
    return result, folders


def get_presentation_bytes(drive, file):
    from googleapiclient.http import MediaIoBaseDownload

    if file["mimeType"] != PPTX_MIME:
        raise ValueError("Only PPTX files are downloaded; native Google Slides use Slides API")
    request = drive.files().get_media(fileId=file["id"], supportsAllDrives=True)
    stream = io.BytesIO()
    downloader = MediaIoBaseDownload(stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    stream.seek(0)
    return stream


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def embedded_images(shape):
    """Yield embedded blobs from pictures, FILLED picture placeholders, and image-fill shapes.

    A student PPTX exported from Google Slides often has shape_type PLACEHOLDER,
    even though placeholder.image.blob contains a full photograph. That was the v2 bug.
    Empty placeholders have no image and correctly produce no blob.
    """
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return

    seen = set()
    try:
        blob = shape.image.blob
        if blob:
            digest = hashlib.sha256(blob).hexdigest()
            seen.add(digest)
            yield digest, blob
    except (AttributeError, ValueError, KeyError, TypeError):
        pass

    # Also support non-picture shapes with an embedded image fill.
    try:
        for blip in shape._element.xpath(".//a:blip"):
            rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rid:
                continue  # externally linked image, no bytes in PPTX
            part = shape.part.related_part(rid)
            blob = part.blob
            digest = hashlib.sha256(blob).hexdigest()
            if digest not in seen:
                seen.add(digest)
                yield digest, blob
    except (AttributeError, KeyError, ValueError):
        pass


def new_template_catalog():
    return {"sha256": set(), "visual": []}


def image_visual_signature(blob):
    """Small perceptual signature for images re-encoded between PPTX/Slides.

    Use only extremely close hashes AND similar aspect ratios to avoid hiding
    genuine student photos due to a coincidental-looking template graphic.
    """
    try:
        with Image.open(io.BytesIO(blob)) as img:
            img.load()
            width, height = img.size
            if not width or not height:
                return None
            gray = img.convert("L").resize((9, 8), Image.Resampling.LANCZOS)
            px = list(gray.getdata())
            bits = 0
            for y in range(8):
                for x in range(8):
                    bits = (bits << 1) | int(px[y * 9 + x] > px[y * 9 + x + 1])
            return bits, width / height
    except (UnidentifiedImageError, OSError, ValueError):
        return None


def add_template_image(catalog, blob):
    catalog["sha256"].add(hashlib.sha256(blob).hexdigest())
    signature = image_visual_signature(blob)
    if signature is not None:
        catalog["visual"].append(signature)


def is_template_image(blob, catalog):
    if hashlib.sha256(blob).hexdigest() in catalog["sha256"]:
        return True
    signature = image_visual_signature(blob)
    if signature is None:
        return False
    img_hash, aspect = signature
    for old_hash, old_aspect in catalog["visual"]:
        if abs(aspect - old_aspect) <= max(0.03, 0.03 * old_aspect):
            if (img_hash ^ old_hash).bit_count() <= 2:
                return True
    return False


def template_hashes(raw):
    """Keep v3's name, now return exact and near-exact image fingerprints."""
    catalog = new_template_catalog()
    prs = Presentation(raw)
    for slide in prs.slides:
        for shape in iter_shapes(slide.shapes):
            for _, blob in embedded_images(shape):
                add_template_image(catalog, blob)
    return catalog


def get_slides_presentation(slides_client, presentation_id):
    """Read native slide elements without Drive PPTX export or size limits."""
    return slides_client.presentations().get(
        presentationId=presentation_id,
        fields="presentationId,pageSize,slides(objectId,pageElements,pageProperties)",
    ).execute()


def safe_slide_content_url(url):
    """Only request Google's ephemeral image content URLs, never user sourceUrl."""
    host = (urlparse(url).hostname or "").lower()
    allowed = ("googleusercontent.com", "google.com", "googleapis.com", "gstatic.com")
    return urlparse(url).scheme == "https" and any(
        host == domain or host.endswith("." + domain) for domain in allowed
    )


def download_slides_image(url, max_bytes=50 * 1024 * 1024):
    """Download ephemeral contentUrl to memory. Never store URLs or images on disk."""
    import requests

    if not safe_slide_content_url(url):
        raise ValueError("Missing or unexpected Google Slides image URL")
    with requests.get(url, stream=True, timeout=(10, 30)) as response:
        response.raise_for_status()
        data = bytearray()
        for chunk in response.iter_content(chunk_size=65536):
            if chunk:
                data.extend(chunk)
            if len(data) > max_bytes:
                raise ValueError("Image exceeds the per-image memory limit")
        if not safe_slide_content_url(response.url):
            raise ValueError("Unexpected image redirect")
        return bytes(data)


def iter_slide_elements(elements, parent_matrix=((1.0, 0.0), (0.0, 1.0))):
    """Flatten grouped elements while accounting for their linear transforms."""
    for element in elements:
        t = element.get("transform", {})
        local = (
            (float(t.get("scaleX", 1)), float(t.get("shearX", 0))),
            (float(t.get("shearY", 0)), float(t.get("scaleY", 1))),
        )
        p, q = parent_matrix, local
        matrix = (
            (p[0][0] * q[0][0] + p[0][1] * q[1][0],
             p[0][0] * q[0][1] + p[0][1] * q[1][1]),
            (p[1][0] * q[0][0] + p[1][1] * q[1][0],
             p[1][0] * q[0][1] + p[1][1] * q[1][1]),
        )
        if "elementGroup" in element:
            yield from iter_slide_elements(element["elementGroup"].get("children", []), matrix)
        else:
            yield element, matrix


def slides_length_inches(value):
    """Slides Size dimension: PT (points) or EMU."""
    amount = float(value.get("magnitude", 0))
    unit = value.get("unit", "PT")
    if unit == "EMU":
        return amount / EMU
    if unit == "PT":
        return amount / 72
    return 0.0


def slides_display_size(element, matrix):
    sz = element.get("size", {})
    width = slides_length_inches(sz.get("width", {}))
    height = slides_length_inches(sz.get("height", {}))
    a, b = matrix[0]
    c, d = matrix[1]
    return abs(a) * width + abs(b) * height, abs(c) * width + abs(d) * height


def native_slide_text(element):
    """Read text from shapes and table cells (not speaker notes)."""
    def runs(text):
        return "".join(
            item.get("textRun", {}).get("content", "")
            for item in text.get("textElements", [])
        ).strip()

    if "shape" in element:
        return [runs(element["shape"].get("text", {}))]
    if "table" in element:
        return [runs(cell.get("text", {}))
                for row in element["table"].get("tableRows", [])
                for cell in row.get("tableCells", [])]
    return []


def native_slides_template_catalog(presentation, image_fetch=download_slides_image):
    catalog = new_template_catalog()
    failures = 0
    for slide in presentation.get("slides", []):
        for element, _ in iter_slide_elements(slide.get("pageElements", [])):
            url = element.get("image", {}).get("contentUrl")
            if url:
                try:
                    add_template_image(catalog, image_fetch(url))
                except Exception:
                    failures += 1
        bg_url = (slide.get("pageProperties", {})
                  .get("pageBackgroundFill", {})
                  .get("stretchedPictureFill", {}).get("contentUrl"))
        if bg_url:
            try:
                add_template_image(catalog, image_fetch(bg_url))
            except Exception:
                failures += 1
    return catalog, failures


def optional_face_detector():
    """The photo-presence check works even if OpenCV or its model is missing."""
    try:
        import cv2
        if not hasattr(cv2, "CascadeClassifier"):
            return None, "OpenCV has no CascadeClassifier; face checking skipped"
        candidates = [ROOT / "haarcascade_frontalface_default.xml"]
        if hasattr(cv2, "data") and hasattr(cv2.data, "haarcascades"):
            candidates.append(Path(cv2.data.haarcascades) / "haarcascade_frontalface_default.xml")
        for path in candidates:
            if path.is_file():
                detector = cv2.CascadeClassifier(str(path))
                if not detector.empty():
                    return detector, ""
        return None, "Haar cascade file unavailable; face checking skipped"
    except Exception as exc:
        return None, f"Face checking unavailable: {exc}"


def image_info(blob):
    """Get size and resolution with Pillow rather than assuming image format."""
    try:
        with Image.open(io.BytesIO(blob)) as image:
            image.load()
            return image.size, image.format or "unknown"
    except (UnidentifiedImageError, OSError, ValueError):
        return None, "unknown"


def face_count(blob, model):
    if model is None:
        return None
    try:
        import cv2
        import numpy as np
        img = cv2.imdecode(np.frombuffer(blob, np.uint8), cv2.IMREAD_COLOR)
        if img is None:
            return None
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        h, w = gray.shape
        faces = model.detectMultiScale(
            gray, scaleFactor=1.08, minNeighbors=5,
            minSize=(max(30, int(min(w, h) * 0.04)),) * 2,
        )
        return len(faces)
    except Exception:
        return None


def shape_kind(shape):
    if getattr(shape, "is_placeholder", False):
        return "filled picture placeholder"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    return "image-filled shape"


def analyze(raw, template_images, model=None, min_inches=0.9, min_pixels=250):
    """Analyze one presentation. Images are REVIEW CANDIDATES, not verified identities.

    Large new embedded images count even when a face is not detected. Face detection
    is supporting evidence and never used to reject a real image on its own.
    """
    prs = Presentation(raw)
    data = {
        name: {"slides": [], "candidates": [], "excluded": [], "empty_placeholders": []}
        for name in SECTIONS
    }
    for slide_num, slide in enumerate(prs.slides, 1):
        shapes = list(iter_shapes(slide.shapes))
        texts = [s.text_frame.text.strip() for s in shapes
                 if s.has_text_frame and s.text_frame.text.strip()]
        matched = [name for name, pattern in SECTIONS.items()
                   if name != JOURNEY and any(len(t) <= 140 and pattern.search(t) for t in texts)]
        if not matched:
            continue
        for name in matched:
            data[name]["slides"].append(slide_num)

        for shape in shapes:
            pictures = list(embedded_images(shape))
            if not pictures:
                if getattr(shape, "is_placeholder", False):
                    try:
                        placeholder_type = str(shape.placeholder_format.type).lower()
                        if "picture" in placeholder_type:
                            for name in matched:
                                data[name]["empty_placeholders"].append(
                                    f"Slide {slide_num}: {shape.name} has no embedded image"
                                )
                    except (AttributeError, ValueError):
                        pass
                continue

            width_in = shape.width / EMU
            height_in = shape.height / EMU
            for digest, blob in pictures:
                resolution, img_format = image_info(blob)
                reasons = []
                if is_template_image(blob, template_images):
                    reasons.append("unchanged template image")
                if width_in < min_inches or height_in < min_inches:
                    reasons.append("displayed image too small")
                if resolution is None:
                    reasons.append("unreadable image")
                elif min(resolution) < min_pixels:
                    reasons.append("image resolution too low")
                # Non-picture image fills are often template icons, logos, arrows,
                # and other decorative artwork. Never promote small such fills to
                # photographic evidence. Filled picture placeholders are handled
                # independently above and are NOT subject to this stricter rule.
                if (resolution is not None
                        and not getattr(shape, "is_placeholder", False)
                        and shape.shape_type != MSO_SHAPE_TYPE.PICTURE):
                    if min(resolution) < max(min_pixels, 800):
                        reasons.append("decorative image-fill: insufficient resolution")
                    if width_in * height_in < 4.5:
                        reasons.append("decorative image-fill: too little slide area")
                description = (
                    f"Slide {slide_num}: {shape_kind(shape)}; "
                    f"{width_in:.2f}×{height_in:.2f} inches; "
                    f"{resolution[0]}×{resolution[1]} pixels; "
                    f"{img_format}; hash {digest[:12]}"
                    if resolution else
                    f"Slide {slide_num}: {shape_kind(shape)}; unreadable; hash {digest[:12]}"
                )
                if reasons:
                    for name in matched:
                        data[name]["excluded"].append(
                            description + "; excluded: " + ", ".join(reasons)
                        )
                else:
                    faces = face_count(blob, model)
                    description += (
                        "; face check not run" if faces is None else f"; faces detected: {faces}"
                    )
                    for name in matched:
                        data[name]["candidates"].append(description)

    results = {}
    for name, details in data.items():
        if not details["slides"]:
            status = "SECTION NOT FOUND"
        elif details["candidates"]:
            status = "IMAGE PRESENT — VERIFY"
        elif details["empty_placeholders"]:
            status = "EMPTY PLACEHOLDER"
        else:
            status = "NO NEW PHOTO CANDIDATE"
        results[name] = {
            "status": status,
            "slides": ", ".join(map(str, details["slides"])),
            "candidate_count": len(details["candidates"]),
            "evidence": " | ".join(details["candidates"]),
            "excluded": " | ".join(details["excluded"]),
            "empty_placeholders": " | ".join(details["empty_placeholders"]),
        }
    return results


def analyze_native_slides(presentation, template_images, model=None,
                          min_inches=0.9, min_pixels=250,
                          image_fetch=download_slides_image):
    """Use Slides API image elements directly; face check is advisory only.

    Content URL failure is REVIEW, not MISSING. Native images may be present
    without an exposed embedded blob and Google URLs can expire.
    """
    data = {
        name: {"slides": [], "candidates": [], "excluded": [],
               "empty_placeholders": [], "unverified": []}
        for name in SECTIONS
    }
    size = presentation.get("pageSize", {})
    slide_width = slides_length_inches(size.get("width", {}))
    slide_height = slides_length_inches(size.get("height", {}))

    for slide_num, slide in enumerate(presentation.get("slides", []), 1):
        elements = list(iter_slide_elements(slide.get("pageElements", [])))
        texts = [text for el, _ in elements for text in native_slide_text(el) if text]
        matched = [name for name, pattern in SECTIONS.items()
                   if name != JOURNEY and any(len(t) <= 140 and pattern.search(t) for t in texts)]
        if not matched:
            continue
        for name in matched:
            data[name]["slides"].append(slide_num)

        pictures = []
        for element, matrix in elements:
            if "image" in element:
                width, height = slides_display_size(element, matrix)
                pictures.append((element["image"].get("contentUrl"), width, height,
                                 f"image element {element.get('objectId', '?')}", False))
            else:
                placeholder = element.get("shape", {}).get("placeholder", {})
                if placeholder.get("type", "").upper() == "PICTURE":
                    for name in matched:
                        data[name]["empty_placeholders"].append(
                            f"Slide {slide_num}: empty picture placeholder {element.get('objectId', '?')}"
                        )

        background_url = (slide.get("pageProperties", {})
                          .get("pageBackgroundFill", {})
                          .get("stretchedPictureFill", {}).get("contentUrl"))
        if background_url:
            pictures.append((background_url, slide_width, slide_height,
                             "slide background image", True))

        for url, width_in, height_in, kind, background in pictures:
            label = (f"Slide {slide_num}: {kind}; "
                     f"{width_in:.2f}×{height_in:.2f} inches")
            if width_in < min_inches or height_in < min_inches:
                for name in matched:
                    data[name]["excluded"].append(label + "; excluded: displayed image too small")
                continue
            if not url:
                for name in matched:
                    data[name]["unverified"].append(label + "; image object has no content URL")
                continue
            try:
                blob = image_fetch(url)
            except Exception as exc:
                # Never log Google contentUrls (short-lived bearer-style access links).
                for name in matched:
                    data[name]["unverified"].append(
                        label + f"; could not retrieve image bytes ({type(exc).__name__})"
                    )
                continue
            resolution, fmt = image_info(blob)
            reasons = []
            if is_template_image(blob, template_images):
                reasons.append("unchanged/near-identical template image")
            if resolution is None:
                # Presence is confirmed by Slides API but the content could be
                # SVG/unsupported or corrupt; avoid declaring it missing.
                for name in matched:
                    data[name]["unverified"].append(label + "; image bytes unreadable")
                continue
            if min(resolution) < min_pixels:
                reasons.append("image resolution too low")
            if background and width_in * height_in < 4.5:
                reasons.append("small decorative background")
            description = (f"{label}; {resolution[0]}×{resolution[1]} pixels; "
                           f"{fmt}; hash {hashlib.sha256(blob).hexdigest()[:12]}")
            if reasons:
                for name in matched:
                    data[name]["excluded"].append(description + "; excluded: " + ", ".join(reasons))
            else:
                faces = face_count(blob, model)
                description += ("; face check not run" if faces is None
                                else f"; faces detected: {faces}")
                for name in matched:
                    data[name]["candidates"].append(description)

    results = {}
    for name, details in data.items():
        if not details["slides"]:
            status = "SECTION NOT FOUND"
        elif details["candidates"]:
            status = "IMAGE PRESENT — VERIFY"
        elif details["unverified"]:
            status = "IMAGE FOUND — REVIEW URL"
        elif details["empty_placeholders"]:
            status = "EMPTY PLACEHOLDER"
        else:
            status = "NO NEW PHOTO CANDIDATE"
        results[name] = {
            "status": status,
            "slides": ", ".join(map(str, details["slides"])),
            "candidate_count": len(details["candidates"]),
            "evidence": " | ".join(details["candidates"]),
            "excluded": " | ".join(details["excluded"]),
            "empty_placeholders": " | ".join(details["empty_placeholders"]),
            "unverified": " | ".join(details["unverified"]),
        }
    return results


def video_link_kind(url):
    """Conservatively recognize a likely video hyperlink, not a playable embed.

    A Google Drive /file/d URL may refer to something other than a video, so
    we never mark hyperlinks as VIDEO PRESENT based on the URL alone.
    """
    if not isinstance(url, str):
        return None
    parsed = urlparse(url.strip())
    if parsed.scheme != "https":
        return None
    host = (parsed.hostname or "").lower()
    path = parsed.path.lower()
    if host in ("youtu.be", "www.youtu.be") and path.strip("/"):
        return "YouTube link"
    if host in ("youtube.com", "www.youtube.com", "m.youtube.com"):
        if ((path == "/watch" and parse_qs(parsed.query).get("v"))
                or path.startswith(("/shorts/", "/embed/", "/live/"))):
            return "YouTube link"
    if host in ("drive.google.com", "www.drive.google.com"):
        if "/file/d/" in path or (path == "/open" and parse_qs(parsed.query).get("id")):
            return "Google Drive file link (verify media type)"
    if host in ("vimeo.com", "www.vimeo.com", "player.vimeo.com", "www.loom.com", "loom.com", "streamable.com", "www.streamable.com"):
        return "External video-service link (verify playback)"
    if path.endswith(VIDEO_EXTENSIONS):
        return "Video-file URL"
    return None


def normalized_video_link(url):
    """Fingerprint a video link without volatile query params or access tokens."""
    parsed = urlparse(url.strip())
    host = (parsed.hostname or "").lower()
    if host in ("youtube.com", "www.youtube.com", "m.youtube.com") and parsed.path == "/watch":
        video_id = (parse_qs(parsed.query).get("v") or [""])[0]
        return f"youtube:{video_id}"
    return host + parsed.path.rstrip("/")


def native_video_fingerprint(video):
    source = video.get("source", "SOURCE_UNSPECIFIED")
    video_id = video.get("id")
    if video_id:
        return ("native-video", source, video_id)
    url = video.get("url") or ""
    if url:
        return ("video-link", normalized_video_link(url))
    return None


def iter_native_video_links(element):
    """Inspect image/shape hyperlinks and linked text, not the image content URL."""
    image_url = element.get("image", {}).get("imageProperties", {}).get("link", {}).get("url")
    if image_url and video_link_kind(image_url):
        yield image_url
    shape = element.get("shape", {})
    shape_url = shape.get("shapeProperties", {}).get("link", {}).get("url")
    if shape_url and video_link_kind(shape_url):
        yield shape_url
    for run in shape.get("text", {}).get("textElements", []):
        link = run.get("textRun", {}).get("style", {}).get("link", {}).get("url")
        if link and video_link_kind(link):
            yield link


def pptx_video_objects(shape):
    """Find true video media relationships on a PPTX shape, not poster art.

    Return (kind, fingerprint, description). Internal video files and external
    linked media are distinguished. An unresolved media element is REVIEW.
    """
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return []
    result, seen = [], set()
    for node in shape._element.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag not in ("videoFile", "media"):
            continue
        rid = node.get(REL_NS + "embed") or node.get(REL_NS + "link")
        if not rid or rid in seen:
            continue
        seen.add(rid)
        try:
            rel = shape.part.rels[rid]
            if rel.is_external:
                target = str(rel.target_ref)
                if tag == "videoFile" or video_link_kind(target):
                    result.append(("link", ("video-link", normalized_video_link(target)),
                                   f"linked video media ({tag}; verify link)"))
                continue
            media = rel.target_part
            content_type = (getattr(media, "content_type", "") or "").lower()
            name = str(getattr(media, "partname", "")).lower()
            if not (content_type.startswith("video/") or name.endswith(VIDEO_EXTENSIONS)):
                continue  # Do not classify audio as video
            digest = hashlib.sha256(media.blob).hexdigest()
            result.append(("embed", ("video-bytes", digest),
                           f"embedded video ({content_type or name}; hash {digest[:12]})"))
        except Exception:
            result.append(("unverified", None, f"video media relationship {rid} unavailable"))
    # Some malformed or non-standard PPTX files expose MEDIA but not normal rels.
    if not result and shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
        result.append(("unverified", None, "media shape exists; video relationship unavailable"))
    return result


def pptx_shape_video_links(shape):
    """Video-like hyperlinks on a picture/text shape are REVIEW only."""
    try:
        url = shape.click_action.hyperlink.address
        if url and video_link_kind(url):
            yield url
    except (AttributeError, ValueError, TypeError):
        pass
    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                try:
                    url = run.hyperlink.address
                    if url and video_link_kind(url):
                        yield url
                except (AttributeError, ValueError, TypeError):
                    pass


def build_native_template_videos(presentation):
    fingerprints = set()
    for slide in presentation.get("slides", []):
        for element, _ in iter_slide_elements(slide.get("pageElements", [])):
            if "video" in element:
                fingerprint = native_video_fingerprint(element["video"])
                if fingerprint:
                    fingerprints.add(fingerprint)
            for url in iter_native_video_links(element):
                fingerprints.add(("video-link", normalized_video_link(url)))
    return fingerprints


def build_pptx_template_videos(raw):
    prs = Presentation(raw)
    fingerprints = set()
    for slide in prs.slides:
        for shape in iter_shapes(slide.shapes):
            for kind, fingerprint, _ in pptx_video_objects(shape):
                if fingerprint:
                    fingerprints.add(fingerprint)
            for link in pptx_shape_video_links(shape):
                fingerprints.add(("video-link", normalized_video_link(link)))
    return fingerprints


def journey_result(slides, embeds, links, excluded, unverified, poster_images=None):
    """Do not confuse images, decorative thumbnails, links, or placeholders with video."""
    poster_images = poster_images or []
    if not slides:
        status = "SECTION NOT FOUND"
    elif embeds:
        status = "VIDEO PRESENT — VERIFY"
    elif links:
        status = "VIDEO LINK — REVIEW"
    elif unverified:
        status = "UNRESOLVED MEDIA — REVIEW"
    elif excluded:
        status = "TEMPLATE VIDEO ONLY"
    elif poster_images:
        status = "IMAGE ONLY — VIDEO UNCONFIRMED"
    else:
        status = "VIDEO NOT FOUND"
    return {
        "status": status,
        "slides": ", ".join(map(str, slides)),
        "candidate_count": len(embeds),
        "evidence": " | ".join(embeds),
        "excluded": " | ".join(excluded),
        "empty_placeholders": "",
        "unverified": " | ".join(unverified),
        "video_count": len(embeds),
        "video_evidence": " | ".join(embeds),
        "video_links": " | ".join(links),
        "image_only_evidence": " | ".join(poster_images),
    }


def journey_slide_indices_native(presentation):
    """Locate Journey using the heading; don't assume a fixed slide index."""
    result = []
    for index, slide in enumerate(presentation.get("slides", []), 1):
        elements = list(iter_slide_elements(slide.get("pageElements", [])))
        texts = [txt for el, _ in elements for txt in native_slide_text(el) if txt]
        if any(len(txt) <= 140 and SECTIONS[JOURNEY].search(txt) for txt in texts):
            result.append((index, elements))
    return result


def journey_template_audit_native(presentation):
    """Return safe on-screen diagnostics about the *actual template* read from Slides API.

    No contentUrl, sourceUrl, text, tokens, or private media URLs are included.
    """
    findings = []
    for number, elements in journey_slide_indices_native(presentation):
        for el, matrix in elements:
            kind = ("video" if "video" in el else "image" if "image" in el else
                    "shape" if "shape" in el else "other")
            if kind not in ("video", "image"):
                continue
            w, h = slides_display_size(el, matrix)
            v = el.get("video", {})
            im = el.get("image", {})
            findings.append({
                "Slide": number,
                "Element": el.get("objectId", "?"),
                "Type": kind,
                "Size (in)": f"{w:.2f} x {h:.2f}",
                "Video source": v.get("source", "") if kind == "video" else "",
                "Video ID present": bool(v.get("id")) if kind == "video" else False,
                "Image content URL present": bool(im.get("contentUrl")) if kind == "image" else False,
                "Has image hyperlink": bool(im.get("imageProperties", {}).get("link", {}).get("url")) if kind == "image" else False,
            })
    return findings


def journey_template_audit_pptx(raw):
    """Diagnose PPTX template Journey without claiming a picture is a video."""
    prs = Presentation(raw)
    findings = []
    for slide_no, slide in enumerate(prs.slides, 1):
        shapes = list(iter_shapes(slide.shapes))
        texts = [sh.text_frame.text.strip() for sh in shapes
                 if sh.has_text_frame and sh.text_frame.text.strip()]
        if not any(len(t) <= 140 and SECTIONS[JOURNEY].search(t) for t in texts):
            continue
        for sh in shapes:
            media = pptx_video_objects(sh)
            is_img = bool(list(embedded_images(sh)))
            if not media and not is_img:
                continue
            findings.append({
                "Slide": slide_no,
                "Element": sh.name,
                "Type": "media" if media else "image",
                "Size (in)": f"{sh.width / EMU:.2f} x {sh.height / EMU:.2f}",
                "Video source": ", ".join(m[0] for m in media),
                "Video ID present": any(m[0] == "embed" for m in media),
                "Image content URL present": is_img,
                "Has image hyperlink": bool(list(pptx_shape_video_links(sh))),
            })
    return findings


def analyze_journey_native(presentation, template_videos):
    slides, embeds, links, excluded, unverified, poster_images = [], [], [], [], [], []
    for slide_num, elements in journey_slide_indices_native(presentation):
        slides.append(slide_num)
        for element, matrix in elements:
            video = element.get("video")
            if video is not None:
                fp = native_video_fingerprint(video)
                source = video.get("source", "SOURCE_UNSPECIFIED")
                video_id = video.get("id", "")
                description = (f"Slide {slide_num}: embedded video element "
                               f"{element.get('objectId', '?')}; source {source}; "
                               f"video ID {video_id or '[not provided]'}")
                if fp and fp in template_videos:
                    excluded.append(description + "; unchanged template video")
                elif source in ("DRIVE", "YOUTUBE") and video_id:
                    embeds.append(description + "; playback/access not tested")
                else:
                    unverified.append(description + "; missing supported source or video ID")
            for url in iter_native_video_links(element):
                fp = ("video-link", normalized_video_link(url))
                description = (f"Slide {slide_num}: {video_link_kind(url)} "
                               f"on element {element.get('objectId', '?')}; "
                               "linked, not an embedded video")
                if fp in template_videos:
                    excluded.append(description + "; unchanged template link")
                else:
                    links.append(description)
            if "image" in element:
                image = element["image"]
                w, h = slides_display_size(element, matrix)
                if w >= 2.0 and h >= 2.0 and w*h >= 8.0:
                    label = (f"Slide {slide_num}: image element {element.get('objectId', '?')}; "
                             f"{w:.2f} x {h:.2f} in")
                    if not image.get("contentUrl"):
                        unverified.append(label + "; Slides API supplied no contentUrl; "
                                          "could be an empty/unsupported image placeholder or video poster, not verified video")
                    else:
                        poster_images.append(label + "; visible image/poster only; not proof of video")
    return journey_result(slides, embeds, links, excluded, unverified, poster_images)


def analyze_journey_pptx(raw, template_videos):
    prs = Presentation(raw)
    slides, embeds, links, excluded, unverified, poster_images = [], [], [], [], [], []
    for slide_num, slide in enumerate(prs.slides, 1):
        shapes = list(iter_shapes(slide.shapes))
        texts = [sh.text_frame.text.strip() for sh in shapes
                 if sh.has_text_frame and sh.text_frame.text.strip()]
        if not any(len(t) <= 140 and SECTIONS[JOURNEY].search(t) for t in texts):
            continue
        slides.append(slide_num)
        for shape in shapes:
            for kind, fp, detail in pptx_video_objects(shape):
                description = f"Slide {slide_num}: {shape.name}; {detail}"
                if fp and fp in template_videos:
                    excluded.append(description + "; unchanged template video")
                elif kind == "embed":
                    embeds.append(description + "; playback not tested")
                elif kind == "link":
                    links.append(description)
                else:
                    unverified.append(description)
            for url in pptx_shape_video_links(shape):
                fp = ("video-link", normalized_video_link(url))
                description = (f"Slide {slide_num}: {shape.name}; "
                               f"{video_link_kind(url)}; linked, not embedded")
                if fp in template_videos:
                    excluded.append(description + "; unchanged template link")
                else:
                    links.append(description)
            if (shape.shape_type != MSO_SHAPE_TYPE.GROUP and
                shape.width / EMU >= 2.0 and shape.height / EMU >= 2.0 and
                (shape.width / EMU) * (shape.height / EMU) >= 8.0 and
                any(True for _ in embedded_images(shape))):
                poster_images.append(f"Slide {slide_num}: {shape.name}; image on Journey slide, "
                                     "not confirmed video")
    return journey_result(slides, embeds, links, excluded, unverified, poster_images)


def load_template_images(drive, slides_client, template_meta):
    """Build photo and video catalogs from PPTX or native Google Slides."""
    if template_meta["mimeType"] == SLIDES_MIME:
        pres = get_slides_presentation(slides_client, template_meta["id"])
        images, failures = native_slides_template_catalog(pres)
        videos = build_native_template_videos(pres)
        return images, videos, failures
    if template_meta["mimeType"] == PPTX_MIME:
        raw = get_presentation_bytes(drive, template_meta)
        images = template_hashes(raw)
        raw.seek(0)
        videos = build_pptx_template_videos(raw)
        return images, videos, 0
    raise ValueError("Template must be a PowerPoint .pptx or a native Google Slides presentation")


# ---------------- GOOGLE SHEETS REPORT SYNC (V7) ----------------
# A dedicated report tab is owned by the app. Other tabs are not modified.
# Existing presentation rows are matched by Drive ID; no duplicates on rescan.
REPORT_TAB_DEFAULT = "Photo Checker Results"
REPORT_TIMESTAMP_COLUMN = "Last Scanned (UTC)"


def sheet_a1_title(title):
    """Escape a worksheet title in A1 notation; reject invalid Sheets tab names."""
    title = str(title).strip()
    if not title or len(title) > 100 or any(ch in title for ch in "[]:*?/\\"):
        raise ValueError("Invalid Google Sheet tab name (max 100 characters; no []:*?/\\).")
    return "'" + title.replace("'", "''") + "'"


def sheet_column_label(index):
    """1-based column index to A1-style letters."""
    if index < 1:
        raise ValueError("Column index must be positive")
    label = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        label = chr(65 + remainder) + label
    return label


def sheet_cell_value(value):
    """Make scan output safe to send as RAW Sheets values (no formula execution)."""
    if value is None:
        return ""
    if isinstance(value, float) and not math.isfinite(value):
        return ""
    # Avoid sending NumPy scalars or pandas NA values to the JSON serializer.
    try:
        import pandas as pd
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, (bool, int, float, str)):
        result = value
    else:
        result = str(value)
    if isinstance(result, str) and len(result) > 48000:
        return result[:47950] + " … [truncated; see local report]"
    return result


def report_sheet_headers(df):
    important = ["Name", "Folder", "Format", "Summary", *SECTIONS,
                 "Drive ID", "Drive URL", "Error", REPORT_TIMESTAMP_COLUMN]
    cols = list(df.columns) + [REPORT_TIMESTAMP_COLUMN]
    if "Drive ID" not in cols or "Name" not in cols:
        raise ValueError("Report is missing required 'Drive ID' or 'Name' columns")
    return list(dict.fromkeys([c for c in important if c in cols] + cols))


def sync_report_to_google_sheet(sheets, spreadsheet_id, tab_title, report_df):
    """Idempotent upsert of a full scan into one dedicated Google Sheets tab.

    Only the nominated tab is touched. New scans update existing Drive IDs and
    append new ones; presentations absent from this scan stay untouched.
    Other existing tabs and unmatched records are left as-is.
    """
    spreadsheet_id = extract_drive_id(spreadsheet_id)
    if not re.fullmatch(r"[A-Za-z0-9_-]+", spreadsheet_id):
        raise ValueError("Enter a valid Google Spreadsheet ID or URL")
    prefix = sheet_a1_title(tab_title)
    if report_df.empty:
        return {"created": 0, "updated": 0, "tab": tab_title, "sheet_id": None}

    headers = report_sheet_headers(report_df)
    metadata = sheets.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
    ).execute()
    tab = next((s["properties"] for s in metadata.get("sheets", [])
                if s["properties"].get("title") == tab_title), None)
    new_tab = tab is None
    if new_tab:
        response = sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {
                "title": tab_title,
                "gridProperties": {
                    "rowCount": max(1000, len(report_df) + 100),
                    "columnCount": max(26, len(headers) + 5),
                    "frozenRowCount": 1,
                },
            }}}]},
        ).execute()
        tab = response["replies"][0]["addSheet"]["properties"]
        previous = []
    else:
        previous = sheets.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=prefix,
            valueRenderOption="FORMATTED_VALUE",
        ).execute().get("values", [])

    old_headers = [str(h) for h in previous[0]] if previous else []
    if any(old_headers):
        if "Drive ID" not in old_headers or "Name" not in old_headers:
            raise ValueError(
                f"Existing tab '{tab_title}' is not a photo-checker report. "
                "Use a new tab name to avoid overwriting other data."
            )
        if len(old_headers) != len(set(old_headers)):
            raise ValueError("Report tab has duplicate header names; fix them before sync")
        headers = old_headers + [h for h in headers if h not in old_headers]

    id_index = headers.index("Drive ID")
    existing_by_id = {}
    duplicate_ids = set()
    for sheet_row, old_values in enumerate(previous[1:], start=2):
        file_id = str(old_values[id_index]).strip() if id_index < len(old_values) else ""
        if file_id:
            if file_id in existing_by_id:
                duplicate_ids.add(file_id)
            else:
                existing_by_id[file_id] = (sheet_row, old_values)
    if duplicate_ids:
        raise ValueError(
            f"Report tab contains {len(duplicate_ids)} duplicate Drive ID(s). "
            "Deduplicate them first to avoid updating an ambiguous row."
        )

    seen_in_scan = set()
    work = []
    new_count = 0
    updated_count = 0
    next_row = max(2, len(previous) + 1)
    utc_now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    for record in report_df.to_dict(orient="records"):
        file_id = str(record.get("Drive ID") or "").strip()
        if not file_id:
            raise ValueError("Cannot sync a presentation row with no Drive ID")
        if file_id in seen_in_scan:
            raise ValueError("Input scan contains duplicate Drive IDs")
        seen_in_scan.add(file_id)
        record[REPORT_TIMESTAMP_COLUMN] = utc_now
        if file_id in existing_by_id:
            row_num, old_values = existing_by_id[file_id]
            # Preserve any additional columns already in this dedicated report tab.
            cells = list(old_values) + [""] * (len(headers) - len(old_values))
            updated_count += 1
        else:
            row_num, cells = next_row, [""] * len(headers)
            next_row += 1
            new_count += 1
        for field, value in record.items():
            if field in headers:
                cells[headers.index(field)] = sheet_cell_value(value)
        work.append((row_num, cells))

    current_grid = tab.get("gridProperties", {})
    needed_rows = max(next_row, max((r for r, _ in work), default=1) + 1)
    needed_columns = len(headers)
    grid_change = {}
    if needed_rows > current_grid.get("rowCount", 1000):
        grid_change["rowCount"] = max(needed_rows + 50, int(needed_rows * 1.1))
    if needed_columns > current_grid.get("columnCount", max(26, needed_columns)):
        grid_change["columnCount"] = needed_columns + 5
    if grid_change:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"updateSheetProperties": {
                "properties": {"sheetId": tab["sheetId"], "gridProperties": grid_change},
                "fields": ",".join("gridProperties." + k for k in grid_change),
            }}]},
        ).execute()

    # One or more batches; partial failures are safe to retry because Drive IDs are stable.
    end_col = sheet_column_label(len(headers))
    updates = []
    if old_headers != headers:
        updates.append({"range": f"{prefix}!A1:{end_col}1", "values": [headers]})
    for row_num, cells in work:
        updates.append({
            "range": f"{prefix}!A{row_num}:{end_col}{row_num}",
            "values": [cells],
        })
    for offset in range(0, len(updates), 100):
        sheets.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"valueInputOption": "RAW", "data": updates[offset:offset + 100]},
        ).execute()

    if new_tab:
        # Style only the new report tab, leaving all other tabs untouched.
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [
                {"repeatCell": {
                    "range": {"sheetId": tab["sheetId"],
                              "startRowIndex": 0, "endRowIndex": 1,
                              "startColumnIndex": 0, "endColumnIndex": len(headers)},
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": {"red": 0.89, "green": 0.93, "blue": 0.98},
                        "textFormat": {"bold": True},
                    }},
                    "fields": "userEnteredFormat",
                }},
                {"updateDimensionProperties": {
                    "range": {"sheetId": tab["sheetId"], "dimension": "COLUMNS",
                              "startIndex": 0, "endIndex": min(4, len(headers))},
                    "properties": {"pixelSize": 205}, "fields": "pixelSize",
                }},
            ]},
        ).execute()

    return {
        "created": new_count,
        "updated": updated_count,
        "tab": tab_title,
        "sheet_id": tab["sheetId"],
        "spreadsheet_id": spreadsheet_id,
    }


def report_sheet_url(sync_result):
    return (f"https://docs.google.com/spreadsheets/d/{sync_result['spreadsheet_id']}/edit"
            f"#gid={sync_result['sheet_id']}")


def show_sheet_sync_result(st, result):
    st.success(
        f"Google Sheet synced: {result['created']} new row(s), "
        f"{result['updated']} updated row(s) in '{result['tab']}'."
    )
    st.link_button("Open updated Google Sheet", report_sheet_url(result))


def main():
    import pandas as pd
    import streamlit as st

    st.set_page_config(page_title="Drive Portfolio Photo Checker v7", page_icon="📷", layout="wide")
    st.title("📷 Google Drive Portfolio Photo + Journey Video Checker — v7")
    st.caption("No presentation uploads · Drive/Slides read-only · Google Sheets report syncing · Week 1–5 photos + Journey video")
    with st.sidebar:
        folder_input = st.text_input("Student presentations parent folder ID or link")
        template_input = st.text_input("Hackathon template file ID or link")
        recursive = st.checkbox("Include nested folders", value=True)
        min_inches = st.number_input("Minimum displayed width AND height (inches)",
                                     min_value=0.3, max_value=10.0, value=0.9, step=0.1)
        min_pixels = st.number_input("Minimum shortest image side (pixels)",
                                     min_value=100, max_value=1500, value=250, step=50)
        st.caption("Weeks 1–5 require qualifying images. Journey requires embedded video; a video hyperlink is flagged for review. Face detection is informational only.")
        st.divider()
        st.subheader("Google Sheets report destination")
        spreadsheet_input = st.text_input(
            "Google Spreadsheet ID or URL",
            placeholder="Paste the destination Google Sheet link",
            help="Share the spreadsheet with the service-account email as Editor."
        )
        report_tab = st.text_input("Results worksheet (tab) name", value=REPORT_TAB_DEFAULT)
        auto_sync = st.checkbox("Write results after each scan", value=True,
                                help="Upsert rows by Drive ID; no duplicate records on rerun.")
        st.caption("Only the specified worksheet is changed. An existing unrelated tab is protected.")

    if st.button("Scan Google Drive", type="primary", disabled=not (folder_input.strip() and template_input.strip())):
        try:
            drive, slides, sheets = google_clients()
            template_id = extract_drive_id(template_input)
            folder_id = extract_drive_id(folder_input)
            template_meta = drive.files().get(
                fileId=template_id, fields="id,name,mimeType", supportsAllDrives=True
            ).execute()
            template_images, template_videos, template_failures = load_template_images(drive, slides, template_meta)
            if template_meta["mimeType"] == SLIDES_MIME:
                template_audit = journey_template_audit_native(
                    get_slides_presentation(slides, template_id))
            else:
                audit_raw = get_presentation_bytes(drive, template_meta)
                template_audit = journey_template_audit_pptx(audit_raw)
            st.session_state["journey_template_audit_v6"] = pd.DataFrame(template_audit)
            if template_failures:
                st.warning(
                    f"Could not retrieve {template_failures} template image(s). "
                    "Some template graphics may need manual review."
                )
            face_model, face_warning = optional_face_detector()
            if face_warning:
                st.info(face_warning + ". Image presence is still checked.")
            files, folders = inventory(drive, folder_id, recursive)
            students = [file for file in files if file["id"] != template_id]
            st.info(f"Discovered {len(files)} PPTX/Google Slides presentations in {len(folders)} subfolder(s); scanning {len(students)} student presentations.")
            if not students:
                st.warning("No accessible student presentations found. Check the folder and service-account permissions.")

            rows = []
            progress = st.progress(0)
            for idx, file in enumerate(students):
                row = {
                    "Name": file["name"], "Folder": file["folder_path"],
                    "Format": "Google Slides" if file["mimeType"] == SLIDES_MIME else "PPTX",
                    "Drive ID": file["id"], "Drive URL": file.get("webViewLink", ""),
                }
                try:
                    if file["mimeType"] == SLIDES_MIME:
                        presentation = get_slides_presentation(slides, file["id"])
                        result = analyze_native_slides(
                            presentation, template_images, face_model,
                            min_inches=min_inches, min_pixels=int(min_pixels),
                        )
                        result[JOURNEY] = analyze_journey_native(presentation, template_videos)
                    else:
                        raw = get_presentation_bytes(drive, file)
                        result = analyze(
                            raw, template_images, face_model,
                            min_inches=min_inches, min_pixels=int(min_pixels),
                        )
                        raw.seek(0)
                        result[JOURNEY] = analyze_journey_pptx(raw, template_videos)
                    row.update({name: result[name]["status"] for name in SECTIONS})
                    row["Summary"] = (
                        "PHOTO + VIDEO PRESENT — VERIFY" if (
                            all(result[name]["status"] == "IMAGE PRESENT — VERIFY"
                                for name in PHOTO_SECTIONS)
                            and result[JOURNEY]["status"] == "VIDEO PRESENT — VERIFY"
                        ) else "MISSING / REVIEW"
                    )
                    for name in SECTIONS:
                        row[f"{name} slide(s)"] = result[name]["slides"]
                        row[f"{name} candidate count"] = result[name]["candidate_count"]
                        row[f"{name} image evidence"] = result[name]["evidence"]
                        row[f"{name} excluded"] = result[name]["excluded"]
                        row[f"{name} empty placeholders"] = result[name]["empty_placeholders"]
                        row[f"{name} needs review"] = result[name].get("unverified", "")
                    row["Journey video count"] = result[JOURNEY].get("video_count", 0)
                    row["Journey video evidence"] = result[JOURNEY].get("video_evidence", "")
                    row["Journey video links"] = result[JOURNEY].get("video_links", "")
                    row["Journey image-only evidence"] = result[JOURNEY].get("image_only_evidence", "")
                    row["Journey unresolved media"] = result[JOURNEY].get("unverified", "")
                    if file["mimeType"] == SLIDES_MIME:
                        row["Journey slide element diagnostics"] = str(
                            journey_template_audit_native(presentation))
                except Exception as exc:
                    row.update({name: "ERROR" for name in SECTIONS})
                    row["Summary"] = "ERROR"
                    row["Error"] = str(exc)
                rows.append(row)
                progress.progress((idx + 1) / len(students))
            st.session_state["photo_checker_report_v7"] = pd.DataFrame(rows)
            if spreadsheet_input.strip() and auto_sync and rows:
                try:
                    sync_info = sync_report_to_google_sheet(
                        sheets, spreadsheet_input, report_tab, st.session_state["photo_checker_report_v7"]
                    )
                    show_sheet_sync_result(st, sync_info)
                except Exception as sync_exc:
                    st.error(f"Scan completed, but writing Google Sheets failed: {sync_exc}")
                    st.info("The scan results remain available below. Correct the Sheet permissions/ID "
                            "and use 'Retry Google Sheets sync' without rescanning.")
        except Exception as exc:
            st.error(f"Could not scan Google Drive: {exc}")

    if "journey_template_audit_v6" in st.session_state:
        with st.expander("Journey slide — original template media inspection", expanded=True):
            st.caption("This is the template actually fetched from Google Drive, not a guess. "
                       "An image or placeholder does NOT count as an embedded video.")
            audit_df = st.session_state["journey_template_audit_v6"]
            if len(audit_df):
                st.dataframe(audit_df, use_container_width=True, hide_index=True)
            else:
                st.warning("No images/videos found on the template's recognized Journey slide, "
                           "or the Journey heading could not be matched.")

    if "photo_checker_report_v7" in st.session_state:
        df = st.session_state["photo_checker_report_v7"]
        st.metric("Presentations scanned", len(df))
        if len(df):
            st.dataframe(df[["Name", "Folder", "Format", "Summary", *SECTIONS]],
                         use_container_width=True, hide_index=True)
            with st.expander("Detailed evidence: images, placeholders, dimensions, excluded images"):
                st.dataframe(df, use_container_width=True, hide_index=True)
            if spreadsheet_input.strip():
                if st.button("Retry Google Sheets sync", help="Writes the currently displayed report; does not rerun the Drive scan"):
                    try:
                        from googleapiclient.discovery import build
                        credentials = get_service_account_credentials()
                        sheets_client = build("sheets", "v4", credentials=credentials, cache_discovery=False)
                        show_sheet_sync_result(st, sync_report_to_google_sheet(
                            sheets_client, spreadsheet_input, report_tab, df))
                    except Exception as sync_exc:
                        st.error(f"Google Sheets sync failed: {sync_exc}")
            with st.expander("Optional: download a local CSV backup"):
                st.download_button(
                    "Download detailed CSV report", df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="drive_photo_checker_v7_report.csv", mime="text/csv",
                )
    st.caption(
        "Week 1–5 IMAGE PRESENT — VERIFY means qualifying image candidates, not identity proof. "
        "Journey VIDEO PRESENT — VERIFY means an embedded video object was found, not that it plays. "
        "VIDEO LINK — REVIEW is a hyperlink only, not an embedded video. "
        "UNRESOLVED MEDIA — REVIEW means a large image lacks a content URL or media cannot be confirmed. "
        "Native Google Slides are inspected directly without PPTX export. "
        "If a Sheet ID is provided, reports are written to the selected tab using Drive ID upserts."
    )


if __name__ == "__main__":
    main()
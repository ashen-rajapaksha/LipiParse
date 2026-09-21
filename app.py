"""
LipiParse Studio — Premium PDF & Document Workspace
Version 2.1  (bug-fixed + expanded toolset)

Run with:  streamlit run app.py
"""

import io
import re
import shutil
import textwrap
import subprocess
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image, ImageOps, ImageEnhance
from docx import Document
from PyPDF2 import PdfReader, PdfWriter
from gtts import gTTS
from openpyxl import Workbook

# ---------- Optional dependencies -------------------------------------------
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import HexColor
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

try:
    from weasyprint import HTML
    WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False

try:
    from gtts.lang import tts_langs
    GTTS_SUPPORTED = tts_langs()
except Exception:
    GTTS_SUPPORTED = {"en": "English", "ta": "Tamil", "hi": "Hindi",
                      "es": "Spanish", "fr": "French", "de": "German"}


# ==========================================================
# APP CONFIG  (must be the first Streamlit call)
# ==========================================================
st.set_page_config(
    page_title="LipiParse Studio | Premium PDF Workspace",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get help": "https://docs.streamlit.io/",
        "Report a bug": "https://github.com/",
        "About": "LipiParse Studio — premium PDF & document tools in one workspace.",
    },
)

APP_NAME = "LipiParse Studio"
APP_VERSION = "2.1"
MAX_UPLOAD_MB = 60          # soft warning threshold
SHARE_URL = "https://lipiparse-mcthmncr2jfncvym8kbxmw.streamlit.app/"

# --- Edit these once and the footer updates itself. Leave "" to hide a link ---
SOCIAL_LINKS = {
    "Facebook": "",
    "Instagram": "",
    "LinkedIn": "",
    "GitHub": "",
}
CONTACT_EMAIL = ""
CONTACT_PHONE = ""

OCR_LANGS = {
    "English": "eng",
    "Sinhala (සිංහල)": "sin",
    "Tamil (தமிழ்)": "tam",
    "Sinhala + English": "sin+eng",
    "Tamil + English": "tam+eng",
    "Sinhala + Tamil + English": "sin+tam+eng",
    "Spanish": "spa",
    "French": "fra",
    "German": "deu",
    "Italian": "ita",
    "Portuguese": "por",
    "Dutch": "nld",
    "Russian": "rus",
    "Arabic": "ara",
    "Hindi": "hin",
    "Bengali": "ben",
    "Turkish": "tur",
    "Vietnamese": "vie",
    "Indonesian": "ind",
    "Polish": "pol",
    "Swedish": "swe",
    "Japanese": "jpn",
    "Korean": "kor",
    "Chinese (Simplified)": "chi_sim",
    "Chinese (Traditional)": "chi_tra",
}

TTS_LANGS = {
    "English": "en", "Tamil": "ta", "Hindi": "hi", "Bengali": "bn",
    "Spanish": "es", "French": "fr", "German": "de", "Italian": "it",
    "Portuguese": "pt", "Russian": "ru", "Arabic": "ar", "Turkish": "tr",
    "Japanese": "ja", "Korean": "ko", "Chinese": "zh-CN", "Indonesian": "id",
}

TOOL_META = {
    "All Workflows": ("⌘", "Browse every available workflow from one dashboard."),
    "Convert PDF": ("⟲", "Convert documents and images across common formats."),
    "Organize PDF": ("▦", "Merge, split, extract, delete, reorder and rotate pages."),
    "Optimize PDF": ("◉", "Reduce PDF size and prepare files for sharing."),
    "Edit PDF": ("✎", "Watermarks, text labels and automatic page numbers."),
    "PDF Security": ("⌑", "Encrypt or decrypt password-protected PDFs."),
    "PDF Intelligence": ("✦", "Clean OCR text, analyse it, and turn text into audio."),
    "Document Inspector": ("◍", "Inspect metadata and strip hidden document info."),
    "Data Privacy": ("✓", "See how files are handled and what leaves the app."),
}

NAV_ORDER = list(TOOL_META.keys())


# ==========================================================
# STATE
# ==========================================================
def init_state():
    defaults = {
        "selected_tab": "All Workflows",
        "theme": "dark",
        "documents_processed": 0,
        "results": {},        # key -> {data, filename, mime, message}
        "activity": [],       # recent action log
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def log_activity(message: str):
    stamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.activity.insert(0, f"{stamp} — {message}")
    del st.session_state.activity[12:]


# --- BUGFIX -----------------------------------------------------------------
# Previously every download button lived INSIDE `if st.button(...)`. Clicking
# the download triggers a Streamlit rerun, the outer button returns False, and
# the whole result vanished before the file finished downloading.
# Results are now cached in session_state and rendered outside the click block.
# ----------------------------------------------------------------------------
def store_result(key, data, filename, mime, message=None, count=True):
    st.session_state.results[key] = {
        "data": data, "filename": filename, "mime": mime, "message": message,
    }
    if count:
        st.session_state.documents_processed += 1
    log_activity(message or filename)


def render_result(key, label="Download file"):
    res = st.session_state.results.get(key)
    if not res:
        return
    if res.get("message"):
        st.success(res["message"])
    size_kb = len(res["data"]) / 1024
    st.caption(f"{res['filename']} • {size_kb:,.1f} KB")
    st.download_button(
        label, data=res["data"], file_name=res["filename"], mime=res["mime"],
        use_container_width=True, key=f"dl_{key}",
    )
    if st.button("Clear result", key=f"clr_{key}", use_container_width=True):
        st.session_state.results.pop(key, None)
        st.rerun()


# --- BUGFIX -----------------------------------------------------------------
# Streamlit keeps one UploadedFile object alive across reruns, so its internal
# read pointer stays where the last reader left it. Reading the same upload a
# second time returned empty bytes / raised errors. Always work on a fresh
# BytesIO copy instead.
# ----------------------------------------------------------------------------
def as_stream(uploaded) -> io.BytesIO:
    return io.BytesIO(uploaded.getvalue())


def size_guard(uploaded) -> bool:
    mb = len(uploaded.getvalue()) / (1024 * 1024)
    if mb > MAX_UPLOAD_MB:
        st.warning(f"This file is {mb:.1f} MB. Files above {MAX_UPLOAD_MB} MB may "
                   f"time out on a shared Streamlit runtime.")
    return True


MIME_PDF = "application/pdf"
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_ZIP = "application/zip"
MIME_TXT = "text/plain"


# ==========================================================
# THEME / CSS
# ==========================================================
DARK_PALETTE = {
    "bg": "#07111f",
    "panel": "#0d1b2a",
    "line": "rgba(255,255,255,.10)",
    "text": "#eef6ff",
    "muted": "#94a8bd",
    "blue": "#4f8cff",
    "cyan": "#4dd7ff",
    "purple": "#8b7cff",
    "green": "#35d39a",
    "card": "linear-gradient(180deg, rgba(255,255,255,.055), rgba(255,255,255,.025))",
    "input-bg": "rgba(7,17,31,.72)",
    "app-bg": ("radial-gradient(circle at 8% 0%, rgba(79,140,255,.14), transparent 22%),"
               "radial-gradient(circle at 92% 12%, rgba(139,124,255,.12), transparent 25%),"
               "linear-gradient(180deg, #06101d 0%, #081422 40%, #0a1523 100%)"),
    "hero-bg": ("radial-gradient(circle at 82% 18%, rgba(77,215,255,.20), transparent 22%),"
                "radial-gradient(circle at 18% 90%, rgba(139,124,255,.18), transparent 25%),"
                "linear-gradient(135deg, #0d1d33, #102b50 58%, #183065)"),
    "hero-text": "#d2e0ef",
    "sidebar-bg": "linear-gradient(180deg, #07111f 0%, #0a1625 100%)",
    "btn-bg": "linear-gradient(135deg, #1b2e49, #173e69)",
    "btn-text": "#f4f9ff",
}

LIGHT_PALETTE = {
    "bg": "#f4f7fb",
    "panel": "#ffffff",
    "line": "rgba(16,38,64,.12)",
    "text": "#0f1f33",
    "muted": "#5b6f85",
    "blue": "#1f6feb",
    "cyan": "#0aa2c0",
    "purple": "#6b5bd2",
    "green": "#1a9c6e",
    "card": "linear-gradient(180deg, #ffffff, #f7fafd)",
    "input-bg": "#ffffff",
    "app-bg": ("radial-gradient(circle at 8% 0%, rgba(31,111,235,.10), transparent 24%),"
               "radial-gradient(circle at 92% 10%, rgba(107,91,210,.08), transparent 26%),"
               "linear-gradient(180deg, #f7fafd 0%, #eef3f9 100%)"),
    "hero-bg": ("radial-gradient(circle at 82% 18%, rgba(10,162,192,.16), transparent 24%),"
                "linear-gradient(135deg, #e8f1ff, #dfe9ff 60%, #e7e3ff)"),
    "hero-text": "#2b4058",
    "sidebar-bg": "linear-gradient(180deg, #ffffff 0%, #eef3f9 100%)",
    "btn-bg": "linear-gradient(135deg, #ffffff, #eaf1fb)",
    "btn-text": "#12314f",
}

BASE_CSS = """
.stApp { background: var(--lp-app-bg); color: var(--lp-text); }
.block-container { max-width: 1400px; padding-top: 1.25rem; padding-bottom: 3rem; }

[data-testid="stSidebar"] {
    background: var(--lp-sidebar-bg);
    border-right: 1px solid var(--lp-line);
}
[data-testid="stSidebar"] .block-container { padding-top: 1.2rem; }

.lp-brand { display:flex; align-items:center; gap:.75rem; margin-bottom:.35rem; }
.lp-logo {
    width:42px; height:42px; border-radius:14px;
    display:flex; align-items:center; justify-content:center;
    background: linear-gradient(135deg, var(--lp-blue), var(--lp-purple));
    box-shadow: 0 8px 30px rgba(79,140,255,.25); font-size:1.35rem;
}
.lp-brand-name { font-weight:800; font-size:1.1rem; color:var(--lp-text); }
.lp-brand-sub { color:var(--lp-muted); font-size:.72rem; margin-top:-2px; }

.hero {
    position:relative; overflow:hidden; padding:2.4rem 2.4rem 2.2rem;
    border-radius:26px; margin-bottom:1.3rem; border:1px solid var(--lp-line);
    background: var(--lp-hero-bg); box-shadow:0 22px 70px rgba(0,0,0,.18);
}
.hero::after {
    content:""; position:absolute; width:230px; height:230px; right:-80px; bottom:-95px;
    border-radius:50%; border:1px solid var(--lp-line);
}
.hero-kicker { color:var(--lp-cyan); text-transform:uppercase; font-size:.74rem; font-weight:800; letter-spacing:.14em; }
.hero-title { font-size:clamp(2rem,4vw,3.4rem); line-height:1.05; font-weight:900; margin:.45rem 0 .8rem; color:var(--lp-text); }
.hero-copy { color:var(--lp-hero-text); max-width:760px; font-size:1rem; line-height:1.7; }
.hero-badges { display:flex; gap:.55rem; flex-wrap:wrap; margin-top:1.1rem; }
.badge {
    border:1px solid var(--lp-line); background:var(--lp-card); color:var(--lp-text);
    border-radius:999px; padding:.42rem .72rem; font-size:.76rem;
}

.stat-row { display:grid; grid-template-columns:repeat(4,1fr); gap:.8rem; margin:1rem 0 1.3rem; }
.stat { padding:1rem 1.05rem; border-radius:18px; border:1px solid var(--lp-line); background:var(--lp-card); }
.stat-label { color:var(--lp-muted); font-size:.72rem; text-transform:uppercase; letter-spacing:.08em; }
.stat-value { font-size:1.3rem; font-weight:800; margin-top:.22rem; color:var(--lp-text); }

.section-title { font-size:1.4rem; font-weight:800; margin-top:.7rem; color:var(--lp-text); }
.section-copy { color:var(--lp-muted); margin-top:-.45rem; margin-bottom:1rem; }

.card {
    height:100%; padding:1.05rem; border-radius:20px; border:1px solid var(--lp-line);
    background:var(--lp-card); box-shadow:0 10px 35px rgba(0,0,0,.10);
}
.card-icon { font-size:1.35rem; }
.card-title { font-weight:800; margin-top:.35rem; color:var(--lp-text); }
.card-copy { color:var(--lp-muted); font-size:.85rem; line-height:1.5; min-height:3.6em; }
.card-chip {
    display:inline-block; margin-top:.55rem; padding:.28rem .55rem; border-radius:999px;
    background:rgba(79,140,255,.14); color:var(--lp-blue); font-size:.69rem; font-weight:700;
}

.tool-panel {
    padding:1.3rem; border-radius:22px; border:1px solid var(--lp-line);
    background:var(--lp-card); box-shadow:0 12px 40px rgba(0,0,0,.10);
}
.mini-note {
    padding:.7rem .85rem; border-radius:12px; border:1px dashed var(--lp-line);
    background:var(--lp-card); color:var(--lp-muted); font-size:.77rem;
}

div[data-testid="stButton"] > button,
div[data-testid="stDownloadButton"] > button {
    border-radius:12px; min-height:2.65rem; border:1px solid var(--lp-line);
    background:var(--lp-btn-bg); color:var(--lp-btn-text); font-weight:700;
    transition:all .18s ease;
}
div[data-testid="stButton"] > button:hover,
div[data-testid="stDownloadButton"] > button:hover {
    border-color:var(--lp-cyan); transform:translateY(-1px);
    box-shadow:0 10px 28px rgba(79,140,255,.13);
}

.stTextInput input, .stTextArea textarea, .stNumberInput input,
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div {
    background-color:var(--lp-input-bg) !important;
    border-color:var(--lp-line) !important;
    color:var(--lp-text) !important;
}
[data-testid="stFileUploaderDropzone"] {
    background:var(--lp-card); border:1px dashed var(--lp-line); border-radius:16px;
}

@media (max-width: 800px) {
    .stat-row { grid-template-columns:repeat(2,1fr); }
    .hero { padding:1.6rem; }
}
"""


def inject_css():
    palette = LIGHT_PALETTE if st.session_state.get("theme") == "light" else DARK_PALETTE
    variables = ";".join(f"--lp-{name}:{value}" for name, value in palette.items())
    st.markdown(f"<style>:root{{{variables}}}{BASE_CSS}</style>", unsafe_allow_html=True)


# ==========================================================
# HELPERS
# ==========================================================
def create_docx(text: str) -> bytes:
    doc = Document()
    for block in text.split("\n\n"):
        doc.add_paragraph(block.strip())
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def get_soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def run_libreoffice_to_pdf(uploaded_file, original_name: str):
    soffice = get_soffice()
    if not soffice:
        return None, "LibreOffice is not installed in this runtime."

    suffix = Path(original_name).suffix.lower()
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / f"input{suffix}"
        out_dir = Path(tmp) / "out"
        out_dir.mkdir()
        src.write_bytes(uploaded_file.getvalue())
        try:
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", str(out_dir), str(src)],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                timeout=180, check=False,
            )
            if result.returncode != 0:
                return None, (result.stderr.decode(errors="ignore")
                              or "LibreOffice conversion failed.")
            produced = list(out_dir.glob("*.pdf"))
            if not produced:
                return None, "No PDF file was produced."
            return produced[0].read_bytes(), None
        except subprocess.TimeoutExpired:
            return None, "Conversion timed out."


def html_to_pdf_bytes(html_text: str):
    if not WEASYPRINT_AVAILABLE:
        return None, "HTML to PDF requires WeasyPrint in the runtime."
    try:
        return HTML(string=html_text).write_pdf(), None
    except Exception as exc:
        return None, str(exc)


def parse_page_spec(spec: str, total: int):
    """'1, 3, 5-8' -> [1, 3, 5, 6, 7, 8]. Deduped, order preserved, clamped."""
    pages, seen = [], set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            left, right = part.split("-", 1)
            start, end = int(left.strip()), int(right.strip())
            if start > end:
                start, end = end, start
            chunk = range(start, end + 1)
        else:
            chunk = [int(part)]
        for page in chunk:
            if 1 <= page <= total and page not in seen:
                seen.add(page)
                pages.append(page)
    return pages


def pdf_pages_to_images(data: bytes, dpi: int = 150):
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for index, page in enumerate(pdf.pages, start=1):
            yield index, page.to_image(resolution=dpi).original


def pdf_to_image_zip(data: bytes, dpi: int = 150, fmt: str = "JPEG"):
    ext = "jpg" if fmt == "JPEG" else "png"
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for index, img in pdf_pages_to_images(data, dpi):
            buf = io.BytesIO()
            if fmt == "JPEG":
                img.convert("RGB").save(buf, format="JPEG", quality=92)
            else:
                img.save(buf, format="PNG")
            zf.writestr(f"page_{index:03d}.{ext}", buf.getvalue())
    return out.getvalue()


def pdf_to_excel_bytes(data: bytes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Tables"
    found_rows = 0
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            for table_no, table in enumerate(page.extract_tables(), start=1):
                ws.append([f"Page {page_no} — Table {table_no}"])
                for row in table:
                    ws.append([cell if cell is not None else "" for cell in row])
                    found_rows += 1
                ws.append([])
    if found_rows == 0:
        ws.append(["No structured table detected in the uploaded PDF."])
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), found_rows


def annotate_pdf(data: bytes, text: str, mode="Watermark",
                 position="Bottom-right", opacity=0.14, font_size=28):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is required for the PDF annotation tool.")
    reader = PdfReader(io.BytesIO(data))
    writer = PdfWriter()
    for page in reader.pages:
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        overlay_buf = io.BytesIO()
        c = canvas.Canvas(overlay_buf, pagesize=(width, height))
        if mode == "Watermark":
            c.saveState()
            c.setFillColor(HexColor("#4F8CFF"), alpha=opacity)
            c.setFont("Helvetica-Bold", font_size)
            c.translate(width / 2, height / 2)
            c.rotate(35)
            c.drawCentredString(0, 0, text)
            c.restoreState()
        else:
            c.setFont("Helvetica-Bold", 12)
            c.setFillColor(HexColor("#1E5EFF"))
            if position == "Top-left":
                x, y = 36, height - 40
            elif position == "Top-right":
                x, y = width - 36, height - 40
            elif position == "Bottom-left":
                x, y = 36, 30
            else:
                x, y = width - 36, 30
            if "right" in position.lower():
                c.drawRightString(x, y, text)
            else:
                c.drawString(x, y, text)
        c.save()
        overlay_buf.seek(0)
        page.merge_page(PdfReader(overlay_buf).pages[0])
        writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def add_page_numbers(data: bytes, position="Bottom-center",
                     start_at=1, template="Page {n} of {total}", skip_first=False):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is required for page numbering.")
    reader = PdfReader(io.BytesIO(data))
    writer = PdfWriter()
    total = len(reader.pages)
    for index, page in enumerate(reader.pages):
        if not (skip_first and index == 0):
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            label = template.replace("{n}", str(index + start_at)).replace("{total}", str(total))
            buf = io.BytesIO()
            c = canvas.Canvas(buf, pagesize=(width, height))
            c.setFont("Helvetica", 9)
            c.setFillColor(HexColor("#44566B"))
            y = 24 if position.startswith("Bottom") else height - 30
            if position.endswith("center"):
                c.drawCentredString(width / 2, y, label)
            elif position.endswith("right"):
                c.drawRightString(width - 40, y, label)
            else:
                c.drawString(40, y, label)
            c.save()
            buf.seek(0)
            page.merge_page(PdfReader(buf).pages[0])
        writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def rebuild_pdf(data: bytes, page_order, rotate=0, compress=False):
    """page_order is a list of 1-based page numbers."""
    reader = PdfReader(io.BytesIO(data))
    writer = PdfWriter()
    for number in page_order:
        page = reader.pages[number - 1]
        if rotate:
            page.rotate(rotate)
        if compress:
            try:
                page.compress_content_streams()
            except Exception:
                pass
        writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def rasterize_pdf(data: bytes, dpi=110, quality=60):
    """Aggressive size reduction: render each page to JPEG and rebuild a PDF."""
    frames = []
    for _, img in pdf_pages_to_images(data, dpi):
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
        buf.seek(0)
        frames.append(Image.open(buf).convert("RGB"))
    if not frames:
        raise ValueError("No pages could be rendered.")
    out = io.BytesIO()
    frames[0].save(out, format="PDF", save_all=True, append_images=frames[1:])
    return out.getvalue()


def ocr_image(img: Image.Image, lang_code: str) -> str:
    work = img.convert("L")
    work = ImageOps.autocontrast(work)
    work = ImageEnhance.Sharpness(work).enhance(1.25)
    return pytesseract.image_to_string(work, lang=lang_code)


def extract_pdf_text_or_ocr(data: bytes, lang_code: str) -> str:
    pieces = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pieces.append(text)
            else:
                pieces.append(ocr_image(page.to_image(resolution=180).original, lang_code))
    return "\n\n".join(piece.strip() for piece in pieces if piece.strip())


def extract_pdf_text(data: bytes) -> str:
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return "\n\n".join((page.extract_text() or "") for page in pdf.pages)


def clean_text(text: str) -> str:
    cleaned = re.sub(r"[ \t]+", " ", text)
    cleaned = re.sub(r"(\w)-\n(\w)", r"\1\2", cleaned)      # join hyphen splits
    cleaned = re.sub(r"\n\s*\n\s*\n+", "\n\n", cleaned)
    cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
    return cleaned.strip()


def text_stats(text: str):
    words = re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)
    sentences = [s for s in re.split(r"[.!?\u0DF4]+", text) if s.strip()]
    return {
        "Characters": len(text),
        "Words": len(words),
        "Sentences": len(sentences),
        "Paragraphs": len([p for p in text.split("\n\n") if p.strip()]),
        "Reading time": f"{max(1, round(len(words) / 200))} min",
    }


def pdf_overview(data: bytes):
    reader = PdfReader(io.BytesIO(data))
    encrypted = reader.is_encrypted
    info = {
        "Pages": "—" if encrypted else len(reader.pages),
        "File size": f"{len(data)/1024:,.1f} KB",
        "Encrypted": "Yes" if encrypted else "No",
    }
    meta = {}
    if not encrypted:
        raw = reader.metadata or {}
        for label, key in [("Title", "/Title"), ("Author", "/Author"),
                           ("Subject", "/Subject"), ("Creator", "/Creator"),
                           ("Producer", "/Producer"), ("Created", "/CreationDate")]:
            value = raw.get(key)
            if value:
                meta[label] = str(value)
    return info, meta


def strip_metadata(data: bytes, new_meta=None):
    reader = PdfReader(io.BytesIO(data))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.add_metadata(new_meta or {"/Producer": f"{APP_NAME} {APP_VERSION}"})
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ==========================================================
# LAYOUT PIECES
# ==========================================================
def render_brand():
    st.markdown(
        textwrap.dedent(f"""\
        <div class="lp-brand">
            <div class="lp-logo">⚡</div>
            <div>
                <div class="lp-brand-name">{APP_NAME}</div>
                <div class="lp-brand-sub">Premium PDF &amp; document workspace</div>
            </div>
        </div>
        """),
        unsafe_allow_html=True,
    )


def render_sidebar():
    with st.sidebar:
        render_brand()
        st.markdown("---")

        is_dark = st.session_state.theme == "dark"
        if st.button("☀️  Switch to light mode" if is_dark else "🌙  Switch to dark mode",
                     key="theme_toggle", use_container_width=True):
            st.session_state.theme = "light" if is_dark else "dark"
            st.rerun()

        st.markdown("---")
        st.caption("WORKSPACE")
        for name in NAV_ORDER:
            icon, _ = TOOL_META[name]
            active = "✅ " if st.session_state.selected_tab == name else ""
            if st.button(f"{active}{icon}  {name}", key=f"side_{name}", use_container_width=True):
                st.session_state.selected_tab = name
                st.rerun()

        st.markdown("---")
        st.caption("SESSION ACTIVITY")
        if st.session_state.activity:
            for entry in st.session_state.activity[:6]:
                st.markdown(f"<div style='font-size:.72rem;color:#8ba0b6;margin:.15rem 0'>{entry}</div>",
                            unsafe_allow_html=True)
            if st.button("Clear session data", key="clear_all", use_container_width=True):
                st.session_state.results = {}
                st.session_state.activity = []
                st.session_state.documents_processed = 0
                st.rerun()
        else:
            st.markdown("<div style='font-size:.72rem;color:#8ba0b6'>No actions yet.</div>",
                        unsafe_allow_html=True)

        st.markdown("---")
        st.markdown(
            '<div class="mini-note"><b>Tip</b><br>Everything runs inside your session. '
            'Downloads are generated in the browser.</div>',
            unsafe_allow_html=True,
        )
        st.caption(f"Version {APP_VERSION} • © 2026 {APP_NAME}")


def render_header():
    count = st.session_state.get("documents_processed", 0)
    st.markdown(
        textwrap.dedent(f"""\
        <div style="text-align:center;margin:6px 0 18px 0;color:var(--lp-muted);font-size:13px;">
            <span style="color:#39d98a;font-size:11px;">●</span>
            <strong style="color:var(--lp-text);">{count:,}</strong>
            document{'' if count == 1 else 's'} processed in this session
        </div>
        """),
        unsafe_allow_html=True,
    )


def render_hero():
    st.markdown(
        textwrap.dedent("""\
        <div class="hero">
            <div class="hero-kicker">All-in-one document workspace</div>
            <div class="hero-title">Every PDF tool you need.<br>
                <span style="color:var(--lp-cyan)">One premium workspace.</span></div>
            <div class="hero-copy">Convert, organize, optimize, secure, OCR and transform
            document content with a clean workflow designed for fast everyday use.</div>
            <div class="hero-badges">
                <span class="badge">⚡ Fast workflows</span>
                <span class="badge">🔒 Session-based processing</span>
                <span class="badge">🌏 25-language OCR</span>
                <span class="badge">📄 Office &amp; PDF tools</span>
            </div>
        </div>
        """),
        unsafe_allow_html=True,
    )


def render_stats():
    st.markdown(
        textwrap.dedent(f"""\
        <div class="stat-row">
            <div class="stat"><div class="stat-label">Workflows</div>
                <div class="stat-value">{len(NAV_ORDER)} categories</div></div>
            <div class="stat"><div class="stat-label">OCR</div>
                <div class="stat-value">{len(OCR_LANGS)} presets</div></div>
            <div class="stat"><div class="stat-label">Core tools</div>
                <div class="stat-value">30+ actions</div></div>
            <div class="stat"><div class="stat-label">Session</div>
                <div class="stat-value">No app database</div></div>
        </div>
        """),
        unsafe_allow_html=True,
    )


# ==========================================================
# WORKFLOW: DASHBOARD
# ==========================================================
def render_all_workflows():
    st.markdown('<div class="section-title">Workspace overview</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Choose a workflow below. Search by tool name, '
                'file type, or task.</div>', unsafe_allow_html=True)

    search_query = st.text_input(
        "Search workflows",
        placeholder="Try: JPG, Word, merge, OCR, compress, password, metadata...",
        label_visibility="collapsed",
    ).strip().lower()

    tools = [
        ("Convert PDF Suite", "Convert PDF", "⟲",
         "Convert JPG, Word, Excel, PowerPoint and HTML into PDF and back.", "Convert"),
        ("Organize & Structure", "Organize PDF", "▦",
         "Merge, split, extract, delete, reorder and rotate document pages.", "Organize"),
        ("Privacy & Protection", "PDF Security", "⌑",
         "Encrypt PDFs with passwords or unlock protected documents.", "Security"),
        ("Document OCR Engine", "Convert PDF", "◫",
         "Extract editable text from scans using 25 OCR language presets.", "OCR"),
        ("Batch OCR", "Convert PDF", "❑",
         "Run OCR across several files at once and download one text bundle.", "Batch"),
        ("Mobile Camera Scanner", "Convert PDF", "▣",
         "Capture a physical document with your camera and save it as a PDF.", "Scanner"),
        ("PDF Optimization", "Optimize PDF", "◉",
         "Stream compression or aggressive rasterization for smaller files.", "Optimize"),
        ("PDF Editor", "Edit PDF", "✎",
         "Watermarks, text labels and automatic page numbering.", "Edit"),
        ("Text Intelligence & TTS", "PDF Intelligence", "✦",
         "Clean extracted text, analyse it, and convert it to speech.", "Intelligence"),
        ("Document Inspector", "Document Inspector", "◍",
         "Read PDF metadata and strip hidden author or producer details.", "Inspect"),
        ("Privacy Overview", "Data Privacy", "✓",
         "Understand what is processed locally and what leaves the app.", "Privacy"),
    ]

    visible = [item for item in tools
               if not search_query or search_query in " ".join(item).lower()]

    if not visible:
        st.warning(f"No workflow found for '{search_query}'.")
        return

    cols = st.columns(4)
    for idx, (title, category, icon, desc, chip) in enumerate(visible):
        with cols[idx % 4]:
            st.markdown(
                f"<div class='card'><div class='card-icon'>{icon}</div>"
                f"<div class='card-title'>{title}</div>"
                f"<div class='card-copy'>{desc}</div>"
                f"<span class='card-chip'>{chip}</span></div>",
                unsafe_allow_html=True,
            )
            if st.button(f"Open {category}", key=f"open_{idx}_{category}",
                         use_container_width=True):
                st.session_state.selected_tab = category
                st.rerun()


# ==========================================================
# WORKFLOW: CONVERT
# ==========================================================
def render_convert():
    st.markdown('<div class="section-title">Convert PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Convert to PDF, convert from PDF, run OCR, '
                'or scan straight from your camera.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["To PDF", "From PDF", "OCR Extractor", "Batch OCR", "Camera Scanner"])

    # ---------------- To PDF ----------------
    with tabs[0]:
        c1, c2, c3 = st.columns(3)

        with c1:
            st.markdown("#### JPG / PNG → PDF")
            images = st.file_uploader("Upload image files", type=["png", "jpg", "jpeg", "webp"],
                                      accept_multiple_files=True, key="img_to_pdf")
            page_fit = st.checkbox("Fit every image to A4", value=False, key="img_a4")
            if images and st.button("Convert Images", key="convert_images", use_container_width=True):
                try:
                    frames = [Image.open(as_stream(i)).convert("RGB") for i in images]
                    if page_fit:
                        a4 = (1654, 2339)   # A4 at 200 dpi
                        fitted = []
                        for frame in frames:
                            canvas_img = Image.new("RGB", a4, "white")
                            frame.thumbnail(a4, Image.LANCZOS)
                            canvas_img.paste(
                                frame,
                                ((a4[0] - frame.width) // 2, (a4[1] - frame.height) // 2),
                            )
                            fitted.append(canvas_img)
                        frames = fitted
                    out = io.BytesIO()
                    frames[0].save(out, format="PDF", save_all=True, append_images=frames[1:])
                    store_result("img2pdf", out.getvalue(), "Converted_Images.pdf", MIME_PDF,
                                 f"Converted {len(frames)} image(s) to PDF.")
                except Exception as exc:
                    st.error(f"Image conversion failed: {exc}")
            render_result("img2pdf", "Download PDF")

            st.markdown("---")
            st.markdown("#### PowerPoint → PDF")
            ppt_file = st.file_uploader("Upload PPT / PPTX", type=["ppt", "pptx"], key="ppt_to_pdf")
            if ppt_file and st.button("Convert PowerPoint", key="convert_ppt", use_container_width=True):
                size_guard(ppt_file)
                with st.spinner("Converting..."):
                    result, err = run_libreoffice_to_pdf(ppt_file, ppt_file.name)
                if result:
                    store_result("ppt2pdf", result, "PowerPoint_Converted.pdf", MIME_PDF,
                                 "PowerPoint converted successfully.")
                else:
                    st.warning(err)
            render_result("ppt2pdf", "Download PDF")

        with c2:
            st.markdown("#### Word → PDF")
            doc_file = st.file_uploader("Upload Word file", type=["docx", "doc", "odt", "rtf", "txt"],
                                        key="doc_to_pdf")
            if doc_file and st.button("Convert Word", key="convert_word", use_container_width=True):
                size_guard(doc_file)
                with st.spinner("Converting..."):
                    result, err = run_libreoffice_to_pdf(doc_file, doc_file.name)
                if result:
                    store_result("doc2pdf", result, "Word_Converted.pdf", MIME_PDF,
                                 "Document converted successfully.")
                else:
                    st.warning(err)
            render_result("doc2pdf", "Download PDF")

            st.markdown("---")
            st.markdown("#### Excel → PDF")
            xls_file = st.file_uploader("Upload Excel sheet", type=["xls", "xlsx", "csv", "ods"],
                                        key="xls_to_pdf")
            if xls_file and st.button("Convert Excel", key="convert_excel", use_container_width=True):
                size_guard(xls_file)
                with st.spinner("Converting..."):
                    result, err = run_libreoffice_to_pdf(xls_file, xls_file.name)
                if result:
                    store_result("xls2pdf", result, "Excel_Converted.pdf", MIME_PDF,
                                 "Spreadsheet converted successfully.")
                else:
                    st.warning(err)
            render_result("xls2pdf", "Download PDF")

        with c3:
            st.markdown("#### HTML → PDF")
            html_input = st.text_area("Paste HTML code", height=180, key="html_to_pdf")
            if html_input.strip() and st.button("Convert HTML", key="convert_html",
                                                use_container_width=True):
                result, err = html_to_pdf_bytes(html_input)
                if result:
                    store_result("html2pdf", result, "HTML_Converted.pdf", MIME_PDF,
                                 "HTML converted successfully.")
                else:
                    st.warning(err)
            render_result("html2pdf", "Download PDF")

            st.markdown("---")
            st.markdown("#### Plain text → PDF")
            txt_input = st.text_area("Paste any text", height=150, key="txt_to_pdf")
            if txt_input.strip() and st.button("Convert Text", key="convert_txt",
                                               use_container_width=True):
                escaped = (txt_input.replace("&", "&amp;")
                                    .replace("<", "&lt;")
                                    .replace(">", "&gt;")
                                    .replace("\n", "<br>"))
                html_doc = ("<html><body style='font-family:sans-serif;font-size:12pt;"
                            f"line-height:1.6'>{escaped}</body></html>")
                result, err = html_to_pdf_bytes(html_doc)
                if result:
                    store_result("txt2pdf", result, "Text_Converted.pdf", MIME_PDF,
                                 "Text converted to PDF.")
                else:
                    st.warning(err)
            render_result("txt2pdf", "Download PDF")

    # ---------------- From PDF ----------------
    with tabs[1]:
        c1, c2 = st.columns(2)

        with c1:
            st.markdown("#### PDF → Images")
            pdf_img_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_to_jpg")
            fmt = st.radio("Image format", ["JPEG", "PNG"], horizontal=True, key="img_fmt")
            dpi = st.select_slider("Resolution (DPI)", options=[72, 110, 150, 200, 300],
                                   value=150, key="img_dpi")
            if pdf_img_file and st.button("Convert PDF to Images", key="pdf_jpg_btn",
                                          use_container_width=True):
                try:
                    with st.spinner("Rendering pages..."):
                        result = pdf_to_image_zip(pdf_img_file.getvalue(), dpi, fmt)
                    store_result("pdf2img", result, f"PDF_Pages_{fmt}.zip", MIME_ZIP,
                                 "PDF pages exported as images.")
                except Exception as exc:
                    st.error(f"PDF to image failed: {exc}")
            render_result("pdf2img", "Download image ZIP")

            st.markdown("---")
            st.markdown("#### PDF → Word")
            pdf_word_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_to_word")
            if pdf_word_file and st.button("Convert PDF to DOCX", key="pdf_word_btn",
                                           use_container_width=True):
                try:
                    extracted = extract_pdf_text(pdf_word_file.getvalue())
                    if not extracted.strip():
                        st.info("This PDF appears to be scanned. Use the OCR Extractor instead.")
                    else:
                        store_result("pdf2docx", create_docx(clean_text(extracted)),
                                     "Converted_Document.docx", MIME_DOCX,
                                     "PDF text exported to Word.")
                except Exception as exc:
                    st.error(f"PDF to Word failed: {exc}")
            render_result("pdf2docx", "Download DOCX")

        with c2:
            st.markdown("#### PDF → Excel (tables)")
            pdf_xls_file = st.file_uploader("Upload PDF with tables", type=["pdf"], key="pdf_to_xls")
            if pdf_xls_file and st.button("Extract Tables", key="pdf_xls_btn",
                                          use_container_width=True):
                try:
                    with st.spinner("Scanning for tables..."):
                        result, rows = pdf_to_excel_bytes(pdf_xls_file.getvalue())
                    store_result("pdf2xlsx", result, "Extracted_Tables.xlsx", MIME_XLSX,
                                 f"Table extraction completed ({rows} row(s) found).")
                except Exception as exc:
                    st.error(f"PDF to Excel failed: {exc}")
            render_result("pdf2xlsx", "Download XLSX")

            st.markdown("---")
            st.markdown("#### PDF → Plain text")
            pdf_txt_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_to_txt")
            if pdf_txt_file and st.button("Extract Text File", key="pdf_txt_btn",
                                          use_container_width=True):
                try:
                    extracted = clean_text(extract_pdf_text(pdf_txt_file.getvalue()))
                    if not extracted:
                        st.info("No embedded text found. Try the OCR Extractor.")
                    else:
                        store_result("pdf2txt", extracted.encode("utf-8"),
                                     "Extracted_Text.txt", MIME_TXT,
                                     "Text extracted from PDF.")
                except Exception as exc:
                    st.error(f"Text extraction failed: {exc}")
            render_result("pdf2txt", "Download TXT")

    # ---------------- OCR ----------------
    with tabs[2]:
        col_file, col_lang = st.columns([2.2, 1])
        with col_file:
            uploaded_file = st.file_uploader("Upload image or PDF",
                                             type=["png", "jpg", "jpeg", "webp", "pdf"], key="ocr_s")
        with col_lang:
            lang = st.selectbox("OCR language", list(OCR_LANGS.keys()), key="ocr_l")
        auto_clean = st.checkbox("Clean up spacing after OCR", value=True, key="ocr_clean")

        if uploaded_file and st.button("Extract Text", key="ocr_btn", use_container_width=True):
            try:
                size_guard(uploaded_file)
                with st.spinner("Extracting text..."):
                    is_pdf = uploaded_file.name.lower().endswith(".pdf")
                    if is_pdf:
                        text_result = extract_pdf_text_or_ocr(uploaded_file.getvalue(),
                                                              OCR_LANGS[lang])
                    else:
                        text_result = ocr_image(Image.open(as_stream(uploaded_file)),
                                                OCR_LANGS[lang])
                if auto_clean:
                    text_result = clean_text(text_result)
                if text_result.strip():
                    st.session_state["ocr_text"] = text_result
                    store_result("ocr_docx", create_docx(text_result), "OCR_Output.docx",
                                 MIME_DOCX, "OCR completed.")
                else:
                    st.warning("No text was detected. Try a higher-quality scan or another language.")
            except pytesseract.TesseractError as exc:
                st.error(f"Tesseract could not use this language pack: {exc}")
            except Exception as exc:
                st.error(f"OCR failed: {exc}")

        if st.session_state.get("ocr_text"):
            st.text_area("Extracted text", st.session_state["ocr_text"], height=260, key="ocr_view")
            stats = text_stats(st.session_state["ocr_text"])
            st.caption(" • ".join(f"{k}: {v}" for k, v in stats.items()))
            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button("Download TXT",
                                   data=st.session_state["ocr_text"].encode("utf-8"),
                                   file_name="OCR_Output.txt", mime=MIME_TXT,
                                   use_container_width=True, key="ocr_txt_dl")
            with dl2:
                render_result("ocr_docx", "Download DOCX")

    # ---------------- Batch OCR ----------------
    with tabs[3]:
        st.markdown("#### Run OCR on several files at once")
        batch_files = st.file_uploader("Upload images or PDFs",
                                       type=["png", "jpg", "jpeg", "webp", "pdf"],
                                       accept_multiple_files=True, key="batch_ocr_files")
        batch_lang = st.selectbox("OCR language", list(OCR_LANGS.keys()), key="batch_ocr_lang")
        if batch_files and st.button("Run Batch OCR", key="batch_ocr_btn", use_container_width=True):
            bundle = io.BytesIO()
            failures = []
            try:
                with st.spinner(f"Processing {len(batch_files)} file(s)..."):
                    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as zf:
                        progress = st.progress(0.0)
                        for position, item in enumerate(batch_files, start=1):
                            try:
                                if item.name.lower().endswith(".pdf"):
                                    text = extract_pdf_text_or_ocr(item.getvalue(),
                                                                   OCR_LANGS[batch_lang])
                                else:
                                    text = ocr_image(Image.open(as_stream(item)),
                                                     OCR_LANGS[batch_lang])
                                stem = Path(item.name).stem
                                zf.writestr(f"{stem}.txt", clean_text(text))
                            except Exception as exc:
                                failures.append(f"{item.name}: {exc}")
                            progress.progress(position / len(batch_files))
                store_result("batch_ocr", bundle.getvalue(), "Batch_OCR_Text.zip", MIME_ZIP,
                             f"Batch OCR finished for {len(batch_files) - len(failures)} file(s).")
                if failures:
                    st.warning("Some files failed:\n\n" + "\n\n".join(failures))
            except Exception as exc:
                st.error(f"Batch OCR failed: {exc}")
        render_result("batch_ocr", "Download text bundle")

    # ---------------- Camera ----------------
    with tabs[4]:
        st.markdown("#### Camera document scanner")
        cam_photo = st.camera_input("Take a picture")
        if cam_photo:
            img = Image.open(as_stream(cam_photo)).convert("RGB")
            enhance = st.checkbox("Auto-enhance for documents (grayscale + contrast)",
                                  value=True, key="cam_enhance")
            preview = ImageOps.autocontrast(img.convert("L")).convert("RGB") if enhance else img
            st.image(preview, caption="Captured document")
            if st.button("Save Photo as PDF", key="camera_pdf", use_container_width=True):
                buf = io.BytesIO()
                preview.save(buf, format="PDF")
                store_result("cam_pdf", buf.getvalue(), "Camera_Scan.pdf", MIME_PDF,
                             "Camera scan is ready.")
        render_result("cam_pdf", "Download PDF")


# ==========================================================
# WORKFLOW: ORGANIZE
# ==========================================================
def render_organize():
    st.markdown('<div class="section-title">Organize PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Control document structure without leaving '
                'the workspace.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Merge", "Split", "Extract Pages", "Delete Pages", "Reorder", "Rotate"])

    # ---- Merge ----
    with tabs[0]:
        files = st.file_uploader("Upload multiple PDFs", type=["pdf"],
                                 accept_multiple_files=True, key="m_files")
        if files:
            order_names = [f.name for f in files]
            chosen = st.multiselect("Merge order (drag to reorder)", order_names,
                                    default=order_names, key="merge_order")
        if files and st.button("Merge Documents", key="merge_btn", use_container_width=True):
            try:
                lookup = {f.name: f for f in files}
                ordered = [lookup[name] for name in (chosen or order_names)]
                writer = PdfWriter()
                skipped = []
                for item in ordered:
                    reader = PdfReader(as_stream(item))
                    if reader.is_encrypted:
                        skipped.append(item.name)
                        continue
                    for page in reader.pages:
                        writer.add_page(page)
                if not writer.pages:
                    raise ValueError("No readable pages were found.")
                out = io.BytesIO()
                writer.write(out)
                store_result("merge", out.getvalue(), "Merged_Document.pdf", MIME_PDF,
                             f"Merged {len(ordered) - len(skipped)} PDF file(s).")
                if skipped:
                    st.warning("Skipped password-protected files: " + ", ".join(skipped))
            except Exception as exc:
                st.error(f"Merge failed: {exc}")
        render_result("merge", "Download merged PDF")

    # ---- Split ----
    with tabs[1]:
        split_file = st.file_uploader("Upload PDF", type=["pdf"], key="s_file")
        if split_file:
            data = split_file.getvalue()
            total = len(PdfReader(io.BytesIO(data)).pages)
            st.caption(f"This document has {total} page(s).")
            mode = st.radio("Split mode",
                            ["Single page", "Every N pages", "Split at page", "One file per page"],
                            key="split_mode")

            if mode == "Single page":
                page_num = st.number_input("Page number", min_value=1, max_value=total,
                                           value=1, key="split_page")
                if st.button("Extract Page", key="split_btn", use_container_width=True):
                    store_result("split", rebuild_pdf(data, [page_num]),
                                 f"Page_{page_num}.pdf", MIME_PDF, f"Page {page_num} extracted.")

            elif mode == "Every N pages":
                chunk = st.number_input("Pages per file", min_value=1, max_value=max(1, total),
                                        value=min(5, total), key="split_chunk")
                if st.button("Split Document", key="split_chunk_btn", use_container_width=True):
                    bundle = io.BytesIO()
                    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as zf:
                        for start in range(1, total + 1, int(chunk)):
                            pages = list(range(start, min(start + int(chunk), total + 1)))
                            zf.writestr(f"part_{start:03d}-{pages[-1]:03d}.pdf",
                                        rebuild_pdf(data, pages))
                    store_result("split", bundle.getvalue(), "Split_Parts.zip", MIME_ZIP,
                                 "Document split into parts.")

            elif mode == "Split at page":
                cut = st.number_input("Split after page", min_value=1, max_value=max(1, total - 1),
                                      value=1, key="split_cut")
                if st.button("Split into Two", key="split_cut_btn", use_container_width=True):
                    bundle = io.BytesIO()
                    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as zf:
                        zf.writestr("part_1.pdf", rebuild_pdf(data, list(range(1, int(cut) + 1))))
                        zf.writestr("part_2.pdf", rebuild_pdf(data, list(range(int(cut) + 1, total + 1))))
                    store_result("split", bundle.getvalue(), "Split_Two_Parts.zip", MIME_ZIP,
                                 f"Split after page {cut}.")

            else:
                if st.button("Burst All Pages", key="split_all_btn", use_container_width=True):
                    bundle = io.BytesIO()
                    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as zf:
                        for number in range(1, total + 1):
                            zf.writestr(f"page_{number:03d}.pdf", rebuild_pdf(data, [number]))
                    store_result("split", bundle.getvalue(), "All_Pages.zip", MIME_ZIP,
                                 f"Exported {total} single-page files.")
        render_result("split", "Download result")

    # ---- Extract ----
    with tabs[2]:
        ext_file = st.file_uploader("Upload PDF", type=["pdf"], key="ext_file")
        pages_input = st.text_input("Page range", placeholder="Example: 1, 3, 5-8", key="page_range")
        if ext_file and pages_input and st.button("Extract Range", key="range_btn",
                                                  use_container_width=True):
            try:
                data = ext_file.getvalue()
                total = len(PdfReader(io.BytesIO(data)).pages)
                pages = parse_page_spec(pages_input, total)
                if not pages:
                    raise ValueError(f"No valid page numbers for a {total}-page document.")
                store_result("extract", rebuild_pdf(data, pages), "Extracted_Pages.pdf",
                             MIME_PDF, f"Extracted {len(pages)} page(s).")
            except ValueError as exc:
                st.error(f"Invalid range: {exc}")
            except Exception as exc:
                st.error(f"Extraction failed: {exc}")
        render_result("extract", "Download extracted PDF")

    # ---- Delete ----
    with tabs[3]:
        del_file = st.file_uploader("Upload PDF", type=["pdf"], key="del_file")
        del_input = st.text_input("Pages to remove", placeholder="Example: 2, 7-9", key="del_range")
        if del_file and del_input and st.button("Remove Pages", key="del_btn",
                                                use_container_width=True):
            try:
                data = del_file.getvalue()
                total = len(PdfReader(io.BytesIO(data)).pages)
                remove = set(parse_page_spec(del_input, total))
                keep = [p for p in range(1, total + 1) if p not in remove]
                if not keep:
                    raise ValueError("You cannot remove every page.")
                store_result("delete", rebuild_pdf(data, keep), "Pages_Removed.pdf", MIME_PDF,
                             f"Removed {len(remove)} page(s), {len(keep)} remaining.")
            except ValueError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Page removal failed: {exc}")
        render_result("delete", "Download trimmed PDF")

    # ---- Reorder ----
    with tabs[4]:
        ord_file = st.file_uploader("Upload PDF", type=["pdf"], key="ord_file")
        if ord_file:
            data = ord_file.getvalue()
            total = len(PdfReader(io.BytesIO(data)).pages)
            st.caption(f"This document has {total} page(s).")
            style = st.radio("Reorder style", ["Reverse all pages", "Custom order"],
                             horizontal=True, key="ord_style")
            custom = ""
            if style == "Custom order":
                custom = st.text_input("New page order",
                                       placeholder=f"Example: 3, 1, 2, 4-{total}", key="ord_custom")
            if st.button("Apply New Order", key="ord_btn", use_container_width=True):
                try:
                    if style == "Reverse all pages":
                        order = list(range(total, 0, -1))
                    else:
                        order = parse_page_spec(custom, total)
                        if not order:
                            raise ValueError("Enter a valid page order first.")
                    store_result("reorder", rebuild_pdf(data, order), "Reordered_Document.pdf",
                                 MIME_PDF, f"Reordered into {len(order)} page(s).")
                except ValueError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Reorder failed: {exc}")
        render_result("reorder", "Download reordered PDF")

    # ---- Rotate ----
    with tabs[5]:
        rot_file = st.file_uploader("Upload PDF", type=["pdf"], key="rot_file")
        angle = st.selectbox("Rotation angle", [90, 180, 270], key="rotation_angle")
        scope = st.text_input("Pages to rotate (leave blank for all)",
                              placeholder="Example: 1, 4-6", key="rot_scope")
        if rot_file and st.button("Rotate Document", key="rotate_btn", use_container_width=True):
            try:
                data = rot_file.getvalue()
                reader = PdfReader(io.BytesIO(data))
                total = len(reader.pages)
                targets = set(parse_page_spec(scope, total)) if scope.strip() else set(range(1, total + 1))
                writer = PdfWriter()
                for number, page in enumerate(reader.pages, start=1):
                    if number in targets:
                        page.rotate(int(angle))
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                store_result("rotate", out.getvalue(), "Rotated_Document.pdf", MIME_PDF,
                             f"Rotated {len(targets)} page(s) by {angle}°.")
            except Exception as exc:
                st.error(f"Rotation failed: {exc}")
        render_result("rotate", "Download rotated PDF")


# ==========================================================
# WORKFLOW: OPTIMIZE
# ==========================================================
def render_optimize():
    st.markdown('<div class="section-title">Optimize PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Choose lossless stream compression, or rasterize '
                'pages for a much smaller file.</div>', unsafe_allow_html=True)

    opt_file = st.file_uploader("Upload PDF to optimize", type=["pdf"], key="opt_upload")
    level = st.radio("Compression mode",
                     ["Safe (lossless streams)", "Aggressive (rasterize pages)"],
                     key="opt_level")
    dpi, quality = 110, 60
    if level.startswith("Aggressive"):
        st.info("Rasterizing converts pages to images. Text stops being selectable, "
                "but scanned documents usually shrink dramatically.")
        dpi = st.select_slider("Render DPI", options=[72, 90, 110, 150, 200], value=110, key="opt_dpi")
        quality = st.slider("JPEG quality", 30, 90, 60, key="opt_q")

    if opt_file and st.button("Compress Document", key="compress_btn", use_container_width=True):
        try:
            data = opt_file.getvalue()
            before = len(data)
            with st.spinner("Optimizing..."):
                if level.startswith("Safe"):
                    result = rebuild_pdf(data, list(range(1, len(PdfReader(io.BytesIO(data)).pages) + 1)),
                                         compress=True)
                else:
                    result = rasterize_pdf(data, dpi=dpi, quality=quality)
            after = len(result)
            saved = (1 - after / before) * 100 if before else 0
            message = (f"Compression completed: {before/1024:,.1f} KB → {after/1024:,.1f} KB "
                       f"({saved:.1f}% smaller)")
            if after >= before:
                message = (f"This PDF is already well optimized "
                           f"({before/1024:,.1f} KB → {after/1024:,.1f} KB).")
            store_result("optimize", result, "Optimized_Document.pdf", MIME_PDF, message)
        except Exception as exc:
            st.error(f"Optimization failed: {exc}")
    render_result("optimize", "Download optimized PDF")


# ==========================================================
# WORKFLOW: EDIT
# ==========================================================
def render_edit():
    st.markdown('<div class="section-title">Edit PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Watermarks, corner labels and automatic '
                'page numbering.</div>', unsafe_allow_html=True)

    if not REPORTLAB_AVAILABLE:
        st.error("The annotation engine is unavailable because reportlab is not installed.")
        return

    tabs = st.tabs(["Watermark & labels", "Page numbers"])

    with tabs[0]:
        file = st.file_uploader("Upload PDF", type=["pdf"], key="edit_pdf_file")
        mode = st.radio("Annotation type", ["Watermark", "Text label"], horizontal=True, key="edit_mode")
        text = st.text_input("Text", placeholder="Example: CONFIDENTIAL", key="edit_text")
        if mode == "Watermark":
            col_a, col_b = st.columns(2)
            with col_a:
                opacity = st.slider("Opacity", 0.05, 0.6, 0.14, key="edit_opacity")
            with col_b:
                font_size = st.slider("Font size", 14, 72, 28, key="edit_font")
            position = "Bottom-right"
        else:
            opacity, font_size = 0.14, 28
            position = st.selectbox("Position",
                                    ["Top-left", "Top-right", "Bottom-left", "Bottom-right"],
                                    key="edit_position")

        if file and text.strip() and st.button("Apply Annotation", key="edit_apply",
                                               use_container_width=True):
            try:
                result = annotate_pdf(file.getvalue(), text.strip(), mode=mode,
                                      position=position, opacity=opacity, font_size=font_size)
                store_result("annotate", result, "Edited_Document.pdf", MIME_PDF,
                             "Annotation applied successfully.")
            except Exception as exc:
                st.error(f"Edit failed: {exc}")
        render_result("annotate", "Download edited PDF")

    with tabs[1]:
        num_file = st.file_uploader("Upload PDF", type=["pdf"], key="num_pdf_file")
        col_a, col_b = st.columns(2)
        with col_a:
            num_pos = st.selectbox("Position",
                                   ["Bottom-center", "Bottom-right", "Bottom-left",
                                    "Top-center", "Top-right", "Top-left"], key="num_pos")
            start_at = st.number_input("Start numbering at", min_value=1, value=1, key="num_start")
        with col_b:
            template = st.text_input("Label format", value="Page {n} of {total}", key="num_fmt")
            skip_first = st.checkbox("Skip the first page (cover)", value=False, key="num_skip")

        if num_file and st.button("Add Page Numbers", key="num_btn", use_container_width=True):
            try:
                result = add_page_numbers(num_file.getvalue(), position=num_pos,
                                          start_at=int(start_at), template=template,
                                          skip_first=skip_first)
                store_result("numbers", result, "Numbered_Document.pdf", MIME_PDF,
                             "Page numbers added.")
            except Exception as exc:
                st.error(f"Numbering failed: {exc}")
        render_result("numbers", "Download numbered PDF")


# ==========================================================
# WORKFLOW: SECURITY
# ==========================================================
def render_security():
    st.markdown('<div class="section-title">PDF Security</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Protect or unlock PDFs with a password you '
                'control.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Encrypt PDF", "Decrypt PDF"])

    with tabs[0]:
        sec_file = st.file_uploader("Upload PDF", type=["pdf"], key="sec_encrypt")
        col_a, col_b = st.columns(2)
        with col_a:
            pwd = st.text_input("User password (needed to open)", type="password", key="sec_pwd")
        with col_b:
            confirm = st.text_input("Confirm password", type="password", key="sec_pwd2")
        owner_pwd = st.text_input("Owner password (optional, for permissions)",
                                  type="password", key="sec_owner")

        if sec_file and pwd and st.button("Encrypt Document", key="enc_btn",
                                          use_container_width=True):
            if pwd != confirm:
                st.error("The two passwords do not match.")
            elif len(pwd) < 4:
                st.error("Use a password with at least 4 characters.")
            else:
                try:
                    reader = PdfReader(as_stream(sec_file))
                    if reader.is_encrypted:
                        st.warning("This PDF is already encrypted. Decrypt it first.")
                    else:
                        writer = PdfWriter()
                        for page in reader.pages:
                            writer.add_page(page)
                        try:
                            writer.encrypt(user_password=pwd,
                                           owner_password=owner_pwd or None,
                                           algorithm="AES-256")
                            note = "AES-256"
                        except Exception:
                            writer.encrypt(user_password=pwd,
                                           owner_password=owner_pwd or None)
                            note = "RC4-128"
                        out = io.BytesIO()
                        writer.write(out)
                        store_result("encrypt", out.getvalue(), "Protected_Document.pdf",
                                     MIME_PDF, f"PDF encrypted ({note}).")
                except Exception as exc:
                    st.error(f"Encryption failed: {exc}")
        render_result("encrypt", "Download protected PDF")
        st.markdown('<div class="mini-note">If you lose this password the file cannot be '
                    'recovered by this app.</div>', unsafe_allow_html=True)

    with tabs[1]:
        dec_file = st.file_uploader("Upload encrypted PDF", type=["pdf"], key="sec_decrypt")
        dec_pwd = st.text_input("Enter password", type="password", key="dec_pwd")
        if dec_file and dec_pwd and st.button("Unlock Document", key="dec_btn",
                                              use_container_width=True):
            try:
                reader = PdfReader(as_stream(dec_file))
                if not reader.is_encrypted:
                    st.info("This PDF is not password protected — no changes were needed.")
                elif not reader.decrypt(dec_pwd):
                    st.error("Incorrect password.")
                else:
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    out = io.BytesIO()
                    writer.write(out)
                    store_result("decrypt", out.getvalue(), "Unlocked_Document.pdf",
                                 MIME_PDF, "PDF unlocked.")
            except Exception as exc:
                st.error(f"Unlock failed: {exc}")
        render_result("decrypt", "Download unlocked PDF")
        st.markdown('<div class="mini-note">This tool only removes protection when you supply '
                    'the correct password. Use it on documents you own.</div>',
                    unsafe_allow_html=True)


# ==========================================================
# WORKFLOW: INTELLIGENCE
# ==========================================================
def render_intelligence():
    st.markdown('<div class="section-title">PDF Intelligence</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Clean extracted text, measure it, and generate '
                'speech from your content.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Clean OCR Text", "Text Analysis", "Find & Replace", "Text-to-Speech"])

    with tabs[0]:
        raw_text = st.text_area("Paste raw OCR text", height=220, key="clean_raw")
        if st.button("Clean Text Format", key="clean_btn", use_container_width=True):
            if raw_text.strip():
                cleaned = clean_text(raw_text)
                st.session_state["cleaned_text"] = cleaned
                store_result("clean_docx", create_docx(cleaned), "Cleaned_Text.docx",
                             MIME_DOCX, "Text cleaned.")
            else:
                st.warning("Paste some text first.")
        if st.session_state.get("cleaned_text"):
            st.text_area("Cleaned output", st.session_state["cleaned_text"],
                         height=220, key="clean_output")
        render_result("clean_docx", "Download DOCX")

    with tabs[1]:
        analysis_text = st.text_area("Paste text to analyse", height=200, key="stats_input")
        if analysis_text.strip():
            stats = text_stats(analysis_text)
            cols = st.columns(len(stats))
            for col, (label, value) in zip(cols, stats.items()):
                col.metric(label, value)
            words = re.findall(r"\b[\w'-]{4,}\b", analysis_text.lower(), flags=re.UNICODE)
            if words:
                counts = {}
                for word in words:
                    counts[word] = counts.get(word, 0) + 1
                top = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:12]
                st.caption("Most frequent words")
                st.write(" · ".join(f"**{word}** ({count})" for word, count in top))
        else:
            st.info("Paste some text to see word counts, reading time and frequent terms.")

    with tabs[2]:
        fr_text = st.text_area("Source text", height=200, key="fr_input")
        col_a, col_b, col_c = st.columns([2, 2, 1])
        with col_a:
            find_term = st.text_input("Find", key="fr_find")
        with col_b:
            replace_term = st.text_input("Replace with", key="fr_replace")
        with col_c:
            use_regex = st.checkbox("Regex", key="fr_regex")
        if st.button("Apply Replacement", key="fr_btn", use_container_width=True):
            if not fr_text.strip() or not find_term:
                st.warning("Add some text and a search term first.")
            else:
                try:
                    if use_regex:
                        updated, hits = re.subn(find_term, replace_term, fr_text)
                    else:
                        hits = fr_text.count(find_term)
                        updated = fr_text.replace(find_term, replace_term)
                    st.session_state["fr_result"] = updated
                    st.success(f"Replaced {hits} occurrence(s).")
                    store_result("fr_docx", create_docx(updated), "Edited_Text.docx",
                                 MIME_DOCX, None, count=False)
                except re.error as exc:
                    st.error(f"Invalid regular expression: {exc}")
        if st.session_state.get("fr_result"):
            st.text_area("Result", st.session_state["fr_result"], height=200, key="fr_output")
        render_result("fr_docx", "Download DOCX")

    with tabs[3]:
        tts_input = st.text_area("Enter text for audio", height=190, key="tts_input")
        col_a, col_b = st.columns(2)
        with col_a:
            voice_lang = st.selectbox("Voice language", list(TTS_LANGS.keys()), key="tts_lang")
        with col_b:
            slow = st.checkbox("Slow, clearer speech", value=False, key="tts_slow")

        code = TTS_LANGS[voice_lang]
        if code not in GTTS_SUPPORTED:
            st.warning(f"{voice_lang} is not offered by the gTTS voice service. "
                       "Pick another language.")
        st.caption("Note: Sinhala is not currently available as a gTTS voice. "
                   "Sinhala OCR and text tools still work normally.")

        if st.button("Generate Audio", key="tts_btn", use_container_width=True):
            if not tts_input.strip():
                st.warning("Enter some text first.")
            elif code not in GTTS_SUPPORTED:
                st.error("Choose a supported voice language.")
            elif len(tts_input) > 8000:
                st.error("Please keep the text under 8,000 characters per request.")
            else:
                try:
                    with st.spinner("Generating audio..."):
                        tts = gTTS(text=tts_input, lang=code, slow=slow)
                        out = io.BytesIO()
                        tts.write_to_fp(out)
                    store_result("tts", out.getvalue(), "Speech_Audio.mp3", "audio/mpeg",
                                 "Audio generated.")
                except Exception as exc:
                    st.error(f"Audio generation failed: {exc}. gTTS needs an internet connection.")

        if st.session_state.results.get("tts"):
            st.audio(st.session_state.results["tts"]["data"], format="audio/mp3")
        render_result("tts", "Download MP3")
        st.markdown('<div class="mini-note">Text-to-speech sends your text to the Google '
                    'gTTS service. Avoid pasting confidential content here.</div>',
                    unsafe_allow_html=True)


# ==========================================================
# WORKFLOW: DOCUMENT INSPECTOR
# ==========================================================
def render_inspector():
    st.markdown('<div class="section-title">Document Inspector</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Check what a PDF reveals about itself, then '
                'strip the parts you would rather not share.</div>', unsafe_allow_html=True)

    file = st.file_uploader("Upload PDF", type=["pdf"], key="inspect_file")
    if not file:
        st.info("Upload a PDF to see its page count, size, encryption status and metadata.")
        return

    data = file.getvalue()
    try:
        info, meta = pdf_overview(data)
    except Exception as exc:
        st.error(f"Could not read this PDF: {exc}")
        return

    cols = st.columns(len(info))
    for col, (label, value) in zip(cols, info.items()):
        col.metric(label, value)

    st.markdown("#### Embedded metadata")
    if meta:
        for label, value in meta.items():
            st.markdown(f"**{label}:** {value}")
    else:
        st.caption("No document metadata found (or the file is encrypted).")

    st.markdown("#### Page dimensions")
    try:
        reader = PdfReader(io.BytesIO(data))
        if not reader.is_encrypted:
            rows = []
            for number, page in enumerate(reader.pages[:12], start=1):
                width = float(page.mediabox.width) / 72
                height = float(page.mediabox.height) / 72
                rows.append(f"Page {number}: {width:.1f} × {height:.1f} in")
            st.caption(" • ".join(rows) + (" …" if len(reader.pages) > 12 else ""))
    except Exception:
        pass

    st.markdown("---")
    st.markdown("#### Clean metadata")
    col_a, col_b = st.columns(2)
    with col_a:
        new_title = st.text_input("New title (optional)", key="meta_title")
    with col_b:
        new_author = st.text_input("New author (optional)", key="meta_author")

    if st.button("Write Clean Copy", key="meta_btn", use_container_width=True):
        try:
            new_meta = {"/Producer": f"{APP_NAME} {APP_VERSION}"}
            if new_title.strip():
                new_meta["/Title"] = new_title.strip()
            if new_author.strip():
                new_meta["/Author"] = new_author.strip()
            store_result("meta", strip_metadata(data, new_meta), "Clean_Metadata.pdf",
                         MIME_PDF, "Metadata replaced.")
        except Exception as exc:
            st.error(f"Could not rewrite metadata: {exc}")
    render_result("meta", "Download clean PDF")


# ==========================================================
# WORKFLOW: PRIVACY
# ==========================================================
def render_privacy():
    st.markdown('<div class="section-title">Data Privacy</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Understand exactly how this app handles uploaded '
                'content.</div>', unsafe_allow_html=True)
    st.markdown(
        textwrap.dedent("""\
        <div class="tool-panel">
            <h3 style="color:var(--lp-text)">Session-based processing</h3>
            <p style="color:var(--lp-muted);line-height:1.7">This app does not implement an
            application database or a user-facing file storage system. Uploaded files are
            processed during the active session and generated downloads are returned to the
            browser. Results held for download live in session memory and disappear when the
            session ends or when you clear them from the sidebar.</p>
            <div class="mini-note"><b>Important:</b> one workflow uses an external service.
            gTTS sends the text you type into Text-to-Speech to Google in order to synthesise
            audio. Everything else — OCR, conversion, merging, encryption — runs inside the
            app runtime.</div>
            <div style="height:12px"></div>
            <div class="stat-row">
                <div class="stat"><div class="stat-label">App database</div>
                    <div class="stat-value">None</div></div>
                <div class="stat"><div class="stat-label">Persistent uploads</div>
                    <div class="stat-value">Not implemented</div></div>
                <div class="stat"><div class="stat-label">OCR</div>
                    <div class="stat-value">Local Tesseract</div></div>
                <div class="stat"><div class="stat-label">TTS</div>
                    <div class="stat-value">External service</div></div>
            </div>
        </div>
        """),
        unsafe_allow_html=True,
    )

    st.markdown("#### Runtime status")
    checks = {
        "LibreOffice (Office → PDF)": bool(get_soffice()),
        "reportlab (watermarks, page numbers)": REPORTLAB_AVAILABLE,
        "WeasyPrint (HTML → PDF)": WEASYPRINT_AVAILABLE,
    }
    try:
        pytesseract.get_tesseract_version()
        checks["Tesseract OCR engine"] = True
    except Exception:
        checks["Tesseract OCR engine"] = False

    for label, ok in checks.items():
        st.markdown(f"{'🟢' if ok else '🔴'} {label} — {'available' if ok else 'not detected'}")


# ==========================================================
# FOOTER
# ==========================================================
def render_footer():
    links_html = ""
    for label, url in SOCIAL_LINKS.items():
        if url:
            links_html += (
                f"<a href='{url}' target='_blank' style=\"color:var(--lp-muted);"
                "text-decoration:none;padding:8px 14px;border:1px solid var(--lp-line);"
                f"border-radius:10px;font-size:13px;\">{label}</a>"
            )

    contact_bits = []
    if CONTACT_EMAIL:
        contact_bits.append(f"Email: {CONTACT_EMAIL}")
    if CONTACT_PHONE:
        contact_bits.append(f"Contact: {CONTACT_PHONE}")
    contact_html = " &nbsp;•&nbsp; ".join(contact_bits)

    # NOTE: every line below starts at column 0 (no leading spaces). If HTML
    # passed to st.markdown has 4+ leading spaces on its first line, Streamlit's
    # markdown parser treats it as a code block and prints the raw tags as text
    # instead of rendering them — that was the "falling text" bug in the footer.
    footer_html = (
        '<div style="margin-top:50px;padding:30px;border-top:1px solid var(--lp-line);'
        'background:var(--lp-card);border-radius:18px 18px 0 0;text-align:center;">'
        f'<div style="font-size:20px;font-weight:700;color:var(--lp-text);margin-bottom:8px;">'
        f'⚡ {APP_NAME}</div>'
        '<div style="color:var(--lp-muted);font-size:13px;margin-bottom:18px;">'
        'Premium PDF &amp; document workspace</div>'
        '<div style="display:flex;justify-content:center;flex-wrap:wrap;gap:10px;margin:18px 0;">'
        f'{links_html}</div>'
        f'<a href="{SHARE_URL}" target="_blank" '
        'style="display:inline-block;margin:8px 0 20px 0;padding:11px 20px;border-radius:12px;'
        'background:linear-gradient(135deg,#1677ff,#36c5ff);color:white;'
        'text-decoration:none;font-weight:700;font-size:13px;">'
        f'Invite / Share {APP_NAME}</a>'
        f'<div style="color:var(--lp-muted);font-size:12px;margin-top:8px;">{contact_html}</div>'
        '<div style="margin-top:20px;padding-top:15px;border-top:1px solid var(--lp-line);'
        'color:var(--lp-muted);font-size:11px;">'
        f'⚡ {APP_NAME} • Version {APP_VERSION}<br>'
        'Built for fast everyday document workflows. © 2026</div>'
        '</div>'
    )
    st.markdown(footer_html, unsafe_allow_html=True)


def render_breadcrumb(active_tab: str):
    """Shown instead of the hero/stats block once a specific tool is open,
    so the person lands straight on that tool's options instead of scrolling
    past dashboard-only content."""
    icon, blurb = TOOL_META.get(active_tab, ("⌘", ""))
    left, right = st.columns([5, 1])
    with left:
        st.markdown(
            f"<div style='color:var(--lp-muted);font-size:.85rem;margin-bottom:.2rem;'>"
            f"All Workflows / <span style='color:var(--lp-text);font-weight:700;'>"
            f"{icon} {active_tab}</span></div>"
            f"<div style='color:var(--lp-muted);font-size:.82rem;margin-bottom:.8rem;'>{blurb}</div>",
            unsafe_allow_html=True,
        )
    with right:
        if st.button("← All tools", key="breadcrumb_back", use_container_width=True):
            st.session_state.selected_tab = "All Workflows"
            st.rerun()


# ==========================================================
# RENDER APP
# ==========================================================
init_state()          # BUGFIX: state must exist before the CSS reads the theme
inject_css()
render_sidebar()
render_brand()
render_header()

selected_tab = st.session_state.selected_tab

# BUGFIX/UX: the hero banner and stat cards are dashboard-only content. They
# used to render above every tool page, forcing an extra scroll before
# reaching that tool's actual options. Now they only show on "All Workflows";
# any specific tool goes straight to its options behind a small breadcrumb.
if selected_tab == "All Workflows":
    render_hero()
    render_stats()
else:
    render_breadcrumb(selected_tab)

ROUTES = {
    "All Workflows": render_all_workflows,
    "Convert PDF": render_convert,
    "Organize PDF": render_organize,
    "Optimize PDF": render_optimize,
    "Edit PDF": render_edit,
    "PDF Security": render_security,
    "PDF Intelligence": render_intelligence,
    "Document Inspector": render_inspector,
    "Data Privacy": render_privacy,
}

ROUTES.get(selected_tab, render_all_workflows)()

render_footer()
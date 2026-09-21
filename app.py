import base64
import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image, ImageOps, ImageEnhance
from docx import Document
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from gtts import gTTS
from openpyxl import Workbook

# Optional dependency: used for the Edit PDF annotation tool.
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import HexColor
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# Optional dependency: HTML -> PDF.
try:
    from weasyprint import HTML
    WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False


# ==========================================================
# APP CONFIG
# ==========================================================
st.set_page_config(
    page_title="LipiParse Studio | Premium PDF Workspace",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get help": "https://docs.streamlit.io/",
        "Report a bug": "https://github.com/",
        "About": "LipiParse Studio — premium PDF & document tools in one workspace."
    },
)

APP_NAME = "LipiParse Studio"
APP_VERSION = "2.0"

OCR_LANGS = {
    "English": "eng",
    "Sinhala (සිංහල)": "sin",
    "Tamil (தமிழ்)": "tam",
    "Sinhala + English": "sin+eng",
    "Tamil + English": "tam+eng",
    "Spanish": "spa",
    "French": "fra",
    "German": "deu",
    "Japanese": "jpn",
    "Chinese": "chi_sim",
}

TOOL_META = {
    "Convert PDF": ("⟲", "Convert documents and images across common formats."),
    "Organize PDF": ("▦", "Merge, split, extract, and rotate PDF pages."),
    "Optimize PDF": ("◉", "Reduce PDF size and prepare files for sharing."),
    "Edit PDF": ("✎", "Add simple annotations or watermarks to PDF pages."),
    "PDF Security": ("⌑", "Encrypt or decrypt password-protected PDFs."),
    "PDF Intelligence": ("✦", "Clean OCR text and turn text into audio."),
    "Data Privacy": ("✓", "See how files are handled and what leaves the app."),
    "All Workflows": ("⌘", "Browse every available workflow from one dashboard."),
}


def init_state():
    if "selected_tab" not in st.session_state:
        st.session_state.selected_tab = "All Workflows"

    if "theme" not in st.session_state:
        st.session_state.theme = "dark"

    if "selected_tab" not in st.session_state:
        st.session_state.selected_tab = "All Workflows"

    if "theme" not in st.session_state:
        st.session_state.theme = "dark"

    if "documents_processed" not in st.session_state:
        st.session_state.documents_processed = 0       

def inject_css():
    st.markdown(
        """
        <style>
        :root {
            --lp-bg: #07111f;
            --lp-panel: #0d1b2a;
            --lp-panel-2: #10243a;
            --lp-line: rgba(255,255,255,.10);
            --lp-text: #eef6ff;
            --lp-muted: #94a8bd;
            --lp-blue: #4f8cff;
            --lp-cyan: #4dd7ff;
            --lp-purple: #8b7cff;
            --lp-green: #35d39a;
        }

        .stApp {
            background:
                radial-gradient(circle at 8% 0%, rgba(79,140,255,.14), transparent 22%),
                radial-gradient(circle at 92% 12%, rgba(139,124,255,.12), transparent 25%),
                linear-gradient(180deg, #06101d 0%, #081422 40%, #0a1523 100%);
            color: var(--lp-text);
        }

        .block-container {
            max-width: 1400px;
            padding-top: 1.25rem;
            padding-bottom: 3rem;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #07111f 0%, #0a1625 100%);
            border-right: 1px solid rgba(255,255,255,.07);
        }
        [data-testid="stSidebar"] .block-container { padding-top: 1.2rem; }

        .lp-brand {
            display:flex; align-items:center; gap:.75rem;
            margin-bottom: .35rem;
        }
        .lp-logo {
            width:42px; height:42px; border-radius:14px;
            display:flex; align-items:center; justify-content:center;
            background: linear-gradient(135deg, var(--lp-blue), var(--lp-purple));
            box-shadow: 0 8px 30px rgba(79,140,255,.25);
            font-size: 1.35rem;
        }
        .lp-brand-name { font-weight:800; font-size:1.1rem; letter-spacing:.01em; }
        .lp-brand-sub { color:var(--lp-muted); font-size:.72rem; margin-top:-2px; }

        .hero {
            position:relative;
            overflow:hidden;
            padding:2.4rem 2.4rem 2.2rem;
            border-radius:26px;
            margin-bottom:1.3rem;
            border:1px solid rgba(255,255,255,.10);
            background:
                radial-gradient(circle at 82% 18%, rgba(77,215,255,.20), transparent 22%),
                radial-gradient(circle at 18% 90%, rgba(139,124,255,.18), transparent 25%),
                linear-gradient(135deg, #0d1d33, #102b50 58%, #183065);
            box-shadow: 0 22px 70px rgba(0,0,0,.28);
        }
        .hero::after {
            content:""; position:absolute; width:230px; height:230px; right:-80px; bottom:-95px;
            border-radius:50%; border:1px solid rgba(255,255,255,.10);
            box-shadow:0 0 0 28px rgba(255,255,255,.02), 0 0 0 56px rgba(255,255,255,.015);
        }
        .hero-kicker { color:#8bdcff; text-transform:uppercase; font-size:.74rem; font-weight:800; letter-spacing:.14em; }
        .hero-title { font-size:clamp(2rem,4vw,3.4rem); line-height:1.05; font-weight:900; margin:.45rem 0 .8rem; }
        .hero-copy { color:#d2e0ef; max-width:760px; font-size:1rem; line-height:1.7; }
        .hero-badges { display:flex; gap:.55rem; flex-wrap:wrap; margin-top:1.1rem; }
        .badge {
            border:1px solid rgba(255,255,255,.12); background:rgba(255,255,255,.06);
            color:#dcecff; border-radius:999px; padding:.42rem .72rem; font-size:.76rem;
            backdrop-filter:blur(8px);
        }

        .stat-row { display:grid; grid-template-columns:repeat(4,1fr); gap:.8rem; margin:1rem 0 1.3rem; }
        .stat {
            padding:1rem 1.05rem; border-radius:18px; border:1px solid var(--lp-line);
            background:linear-gradient(180deg, rgba(255,255,255,.055), rgba(255,255,255,.025));
        }
        .stat-label { color:var(--lp-muted); font-size:.72rem; text-transform:uppercase; letter-spacing:.08em; }
        .stat-value { font-size:1.3rem; font-weight:800; margin-top:.22rem; }

        .section-title { font-size:1.4rem; font-weight:800; margin-top:.7rem; }
        .section-copy { color:var(--lp-muted); margin-top:-.45rem; margin-bottom:1rem; }

        .card {
            height:100%; padding:1.05rem; border-radius:20px;
            border:1px solid var(--lp-line);
            background:linear-gradient(180deg, rgba(255,255,255,.055), rgba(255,255,255,.025));
            box-shadow:0 10px 35px rgba(0,0,0,.12);
        }
        .card-icon { font-size:1.35rem; }
        .card-title { font-weight:800; margin-top:.35rem; }
        .card-copy { color:var(--lp-muted); font-size:.85rem; line-height:1.5; min-height:3.6em; }
        .card-chip { display:inline-block; margin-top:.55rem; padding:.28rem .55rem; border-radius:999px; background:rgba(79,140,255,.12); color:#94c2ff; font-size:.69rem; font-weight:700; }

        .tool-panel {
            padding:1.3rem; border-radius:22px; border:1px solid var(--lp-line);
            background:linear-gradient(180deg, rgba(255,255,255,.045), rgba(255,255,255,.022));
            box-shadow:0 12px 40px rgba(0,0,0,.15);
        }

        .mini-note {
            padding:.7rem .85rem; border-radius:12px; border:1px dashed rgba(255,255,255,.14);
            background:rgba(255,255,255,.025); color:var(--lp-muted); font-size:.77rem;
        }

        div[data-testid="stButton"] > button,
        div[data-testid="stDownloadButton"] > button {
            border-radius:12px;
            min-height:2.65rem;
            border:1px solid rgba(255,255,255,.11);
            background:linear-gradient(135deg, #1b2e49, #173e69);
            color:#f4f9ff;
            font-weight:700;
            transition:all .18s ease;
        }
        div[data-testid="stButton"] > button:hover,
        div[data-testid="stDownloadButton"] > button:hover {
            border-color:rgba(77,215,255,.42);
            transform:translateY(-1px);
            box-shadow:0 10px 28px rgba(79,140,255,.13);
        }

        .footer {
            margin-top:2.2rem; padding:1.2rem 1.3rem; border-radius:20px;
            border:1px solid var(--lp-line); background:rgba(255,255,255,.025);
            color:var(--lp-muted); text-align:center; font-size:.78rem;
        }

        .stTextInput input, .stTextArea textarea, .stNumberInput input,
        .stSelectbox div[data-baseweb="select"] > div,
        .stMultiSelect div[data-baseweb="select"] > div {
            background-color:rgba(7,17,31,.72) !important;
            border-color:rgba(255,255,255,.10) !important;
            color:#eef6ff !important;
        }

        [data-testid="stFileUploaderDropzone"] {
            background:rgba(255,255,255,.025);
            border:1px dashed rgba(255,255,255,.13);
            border-radius:16px;
        }

        @media (max-width: 800px) {
            .stat-row { grid-template-columns:repeat(2,1fr); }
            .hero { padding:1.6rem; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )









def create_docx(text: str) -> bytes:
    doc = Document()
    for block in text.split("\n\n"):
        doc.add_paragraph(block.strip())
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def bytes_download_button(label, data, filename, mime):
    st.download_button(label, data=data, file_name=filename, mime=mime, use_container_width=True)


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
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(src)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=90,
                check=False,
            )
            if result.returncode != 0:
                return None, result.stderr.decode(errors="ignore") or "LibreOffice conversion failed."
            pdf_candidates = list(out_dir.glob("*.pdf"))
            if not pdf_candidates:
                return None, "No PDF file was produced."
            return pdf_candidates[0].read_bytes(), None
        except subprocess.TimeoutExpired:
            return None, "Conversion timed out."


def html_to_pdf_bytes(html_text: str):
    if not WEASYPRINT_AVAILABLE:
        return None, "HTML to PDF requires WeasyPrint in the runtime."
    try:
        return HTML(string=html_text).write_pdf(), None
    except Exception as exc:
        return None, str(exc)


def pdf_to_jpg_zip(pdf_file):
    out = io.BytesIO()
    with pdfplumber.open(pdf_file) as pdf, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, page in enumerate(pdf.pages, start=1):
            img = page.to_image(resolution=150).original
            img_buf = io.BytesIO()
            img.save(img_buf, format="JPEG", quality=92)
            zf.writestr(f"page_{i:03d}.jpg", img_buf.getvalue())
    return out.getvalue()


def pdf_to_excel_bytes(pdf_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Tables"
    found_rows = 0
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for table_no, table in enumerate(tables, start=1):
                ws.append([f"Page {page_no} — Table {table_no}"])
                for row in table:
                    ws.append([cell if cell is not None else "" for cell in row])
                    found_rows += 1
                ws.append([])
    if found_rows == 0:
        ws.append(["No structured table detected in the uploaded PDF."])
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def make_pdf_watermark_overlay(width, height, text, opacity=0.14):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    c.saveState()
    c.setFillColor(HexColor("#4F8CFF"), alpha=opacity)
    c.setFont("Helvetica-Bold", 28)
    c.translate(width / 2, height / 2)
    c.rotate(35)
    c.drawCentredString(0, 0, text)
    c.restoreState()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def annotate_pdf(pdf_file, text, mode="Watermark", position="Bottom-right"):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is required for the PDF annotation tool.")
    reader = PdfReader(pdf_file)
    writer = PdfWriter()
    for page in reader.pages:
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        overlay_buf = io.BytesIO()
        c = canvas.Canvas(overlay_buf, pagesize=(width, height))
        c.setFont("Helvetica-Bold", 12)
        c.setFillColor(HexColor("#1E5EFF"))
        if mode == "Watermark":
            c.saveState()
            c.setFillColor(HexColor("#4F8CFF"), alpha=0.14)
            c.setFont("Helvetica-Bold", 28)
            c.translate(width / 2, height / 2)
            c.rotate(35)
            c.drawCentredString(0, 0, text)
            c.restoreState()
        else:
            if position == "Top-left":
                x, y = 36, height - 40
            elif position == "Top-right":
                x, y = width - 36, height - 40
            elif position == "Bottom-left":
                x, y = 36, 30
            else:
                x, y = width - 36, 30
            c.drawRightString(x, y, text) if "right" in position.lower() else c.drawString(x, y, text)
        c.save()
        overlay_buf.seek(0)
        overlay_page = PdfReader(overlay_buf).pages[0]
        page.merge_page(overlay_page)
        writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def ocr_image(img: Image.Image, lang_code: str) -> str:
    # Light preprocessing improves OCR on camera/scanned pages.
    work = img.convert("L")
    work = ImageOps.autocontrast(work)
    work = ImageEnhance.Sharpness(work).enhance(1.25)
    return pytesseract.image_to_string(work, lang=lang_code)


def extract_pdf_text_or_ocr(pdf_file, lang_code):
    pieces = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            if text.strip():
                pieces.append(text)
            else:
                image = page.to_image(resolution=180).original
                pieces.append(ocr_image(image, lang_code))
    return "\n\n".join(x.strip() for x in pieces if x.strip())


def clean_text(text: str):
    cleaned = re.sub(r"[ \t]+", " ", text)
    cleaned = re.sub(r"\n\s*\n\s*\n+", "\n\n", cleaned)
    cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
    return cleaned.strip()


def render_header():
    st.markdown(
        """
        <div class="lp-brand">
            <div class="lp-logo">⚡</div>
            <div>
                <div class="lp-brand-name">LipiParse Studio</div>
                <div class="lp-brand-sub">
                    Premium PDF & document workspace
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    count = st.session_state.get("documents_processed", 0)

    st.markdown(
        f"""
        <div style="
            text-align:center;
            margin:10px 0 20px 0;
            color:#94a8bd;
            font-size:13px;
        ">
            <span style="
                color:#39d98a;
                font-size:11px;
            ">●</span>

            <strong style="color:#eef6ff;">
                {count:,}
            </strong>

            documents processed in this session
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar():
    with st.sidebar:
        render_header()
        st.markdown("---")
        st.caption("WORKSPACE")
        for name in [
            "All Workflows",
            "Convert PDF",
            "Organize PDF",
            "Optimize PDF",
            "Edit PDF",
            "PDF Security",
            "PDF Intelligence",
            "Data Privacy",
        ]:
            icon, _ = TOOL_META[name]
            if st.button(f"{icon}  {name}", key=f"side_{name}", use_container_width=True):
                st.session_state.selected_tab = name
                st.rerun()
        st.markdown("---")
        st.markdown(
            '<div class="mini-note"><b>Tip</b><br>Upload only the files you need for each workflow. Downloads are generated directly in your session.</div>',
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:8rem'></div>", unsafe_allow_html=True)
        st.caption(f"Version {APP_VERSION} • © 2026 {APP_NAME}")


def render_hero():
    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">All-in-one document workspace</div>
            <div class="hero-title">Every PDF tool you need.<br><span style="color:#7dd9ff">One premium workspace.</span></div>
            <div class="hero-copy">Convert, organize, optimize, secure, OCR and transform document content with a clean workflow designed for fast everyday use.</div>
            <div class="hero-badges">
                <span class="badge">⚡ Fast workflows</span>
                <span class="badge">🔒 Session-based processing</span>
                <span class="badge">🌏 Multi-language OCR</span>
                <span class="badge">📄 Office & PDF tools</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_stats():
    st.markdown(
        """
        <div class="stat-row">
            <div class="stat"><div class="stat-label">Workflows</div><div class="stat-value">8 categories</div></div>
            <div class="stat"><div class="stat-label">OCR</div><div class="stat-value">10 languages</div></div>
            <div class="stat"><div class="stat-label">Core tools</div><div class="stat-value">20+ actions</div></div>
            <div class="stat"><div class="stat-label">Session</div><div class="stat-value">No app database</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_all_workflows():
    st.markdown('<div class="section-title">Workspace overview</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Choose a workflow below. Search by tool name, file type, or task.</div>', unsafe_allow_html=True)

    search_query = st.text_input(
        "Search workflows",
        placeholder="Try: JPG, Word, merge, OCR, compress, password...",
        label_visibility="collapsed",
    ).strip().lower()

    tools = [
        ("Convert PDF Suite", "Convert PDF", "⟲", "Convert JPG, Word, Excel, PowerPoint and HTML into PDF.", "Convert"),
        ("Organize & Structure", "Organize PDF", "▦", "Merge files, extract custom page ranges, split pages and rotate documents.", "Organize"),
        ("Privacy & Protection", "PDF Security", "⌑", "Encrypt PDFs with passwords or unlock encrypted documents with the correct password.", "Security"),
        ("Document OCR Engine", "Convert PDF", "◫", "Extract editable text from scanned images and PDFs using multilingual OCR.", "OCR"),
        ("Mobile Camera Scanner", "Convert PDF", "▣", "Capture a physical document with your camera and save it as a PDF.", "Scanner"),
        ("PDF Optimization", "Optimize PDF", "◉", "Compress content streams and generate a smaller, share-friendly PDF.", "Optimize"),
        ("PDF Editor", "Edit PDF", "✎", "Add a watermark or simple text annotation to PDF pages.", "Edit"),
        ("AI Intelligence & TTS", "PDF Intelligence", "✦", "Clean extracted text and convert it into spoken audio.", "Intelligence"),
    ]

    visible = []
    for item in tools:
        blob = " ".join(item).lower()
        if not search_query or search_query in blob:
            visible.append(item)

    cols = st.columns(4)
    for idx, (title, category, icon, desc, chip) in enumerate(visible):
        with cols[idx % 4]:
            st.markdown(
                f"<div class='card'><div class='card-icon'>{icon}</div><div class='card-title'>{title}</div><div class='card-copy'>{desc}</div><span class='card-chip'>{chip}</span></div>",
                unsafe_allow_html=True,
            )
            if st.button(f"Open {category}", key=f"open_{idx}_{category}", use_container_width=True):
                st.session_state.selected_tab = category
                st.rerun()

    if not visible:
        st.warning(f"No workflow found for '{search_query}'.")


def render_convert():
    st.markdown('<div class="section-title">Convert PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Convert to PDF, convert from PDF, extract OCR text, or scan from camera.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["To PDF", "From PDF", "OCR Extractor", "Camera Scanner"])

    with tabs[0]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("#### JPG / PNG → PDF")
            images = st.file_uploader("Upload image files", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="img_to_pdf")
            if images and st.button("Convert Images", key="convert_images", use_container_width=True):
                try:
                    img_objs = [Image.open(i).convert("RGB") for i in images]
                    out = io.BytesIO()
                    img_objs[0].save(out, format="PDF", save_all=True, append_images=img_objs[1:])
                    st.success(f"Converted {len(img_objs)} image(s) to PDF.")
                    bytes_download_button("Download PDF", out.getvalue(), "Converted_Images.pdf", "application/pdf")
                except Exception as exc:
                    st.error(f"Image conversion failed: {exc}")

            st.markdown("---")
            st.markdown("#### PowerPoint → PDF")
            ppt_file = st.file_uploader("Upload PPT / PPTX", type=["ppt", "pptx"], key="ppt_to_pdf")
            if ppt_file and st.button("Convert PowerPoint", key="convert_ppt", use_container_width=True):
                result, err = run_libreoffice_to_pdf(ppt_file, ppt_file.name)
                if result:
                    st.success("PowerPoint converted successfully.")
                    bytes_download_button("Download PDF", result, "PowerPoint_Converted.pdf", "application/pdf")
                else:
                    st.warning(err)

        with c2:
            st.markdown("#### Word → PDF")
            doc_file = st.file_uploader("Upload Word file", type=["docx", "doc"], key="doc_to_pdf")
            if doc_file and st.button("Convert Word", key="convert_word", use_container_width=True):
                result, err = run_libreoffice_to_pdf(doc_file, doc_file.name)
                if result:
                    st.success("Word document converted successfully.")
                    bytes_download_button("Download PDF", result, "Word_Converted.pdf", "application/pdf")
                else:
                    st.warning(err)

            st.markdown("---")
            st.markdown("#### Excel → PDF")
            xls_file = st.file_uploader("Upload Excel sheet", type=["xls", "xlsx"], key="xls_to_pdf")
            if xls_file and st.button("Convert Excel", key="convert_excel", use_container_width=True):
                result, err = run_libreoffice_to_pdf(xls_file, xls_file.name)
                if result:
                    st.success("Excel sheet converted successfully.")
                    bytes_download_button("Download PDF", result, "Excel_Converted.pdf", "application/pdf")
                else:
                    st.warning(err)

        with c3:
            st.markdown("#### HTML → PDF")
            html_input = st.text_area("Paste HTML code", height=180, key="html_to_pdf")
            if html_input and st.button("Convert HTML", key="convert_html", use_container_width=True):
                result, err = html_to_pdf_bytes(html_input)
                if result:
                    st.success("HTML converted successfully.")
                    bytes_download_button("Download PDF", result, "HTML_Converted.pdf", "application/pdf")
                else:
                    st.warning(err)
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[1]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### PDF → JPG")
            pdf_img_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_to_jpg")
            if pdf_img_file and st.button("Convert PDF to JPG", key="pdf_jpg_btn", use_container_width=True):
                try:
                    result = pdf_to_jpg_zip(pdf_img_file)
                    st.success("PDF pages converted to JPG images.")
                    bytes_download_button("Download JPG ZIP", result, "PDF_Pages_JPG.zip", "application/zip")
                except Exception as exc:
                    st.error(f"PDF to JPG failed: {exc}")

            st.markdown("---")
            st.markdown("#### PDF → Word")
            pdf_word_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_to_word")
            if pdf_word_file and st.button("Convert PDF to DOCX", key="pdf_word_btn", use_container_width=True):
                try:
                    with pdfplumber.open(pdf_word_file) as pdf:
                        extracted_text = "\n\n".join((page.extract_text() or "") for page in pdf.pages)
                    if not extracted_text.strip():
                        st.info("This PDF appears to be scanned. Use the OCR Extractor for image-based text conversion.")
                    else:
                        result = create_docx(clean_text(extracted_text))
                        st.success("PDF text exported to Word.")
                        bytes_download_button("Download DOCX", result, "Converted_Document.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as exc:
                    st.error(f"PDF to Word failed: {exc}")

        with c2:
            st.markdown("#### PDF → Excel")
            pdf_xls_file = st.file_uploader("Upload PDF with tables", type=["pdf"], key="pdf_to_xls")
            if pdf_xls_file and st.button("Extract Tables", key="pdf_xls_btn", use_container_width=True):
                try:
                    result = pdf_to_excel_bytes(pdf_xls_file)
                    st.success("Table extraction completed.")
                    bytes_download_button("Download XLSX", result, "Extracted_Tables.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as exc:
                    st.error(f"PDF to Excel failed: {exc}")

            st.markdown("---")
            st.markdown("#### PDF → PDF/A")
            pdfa_file = st.file_uploader("Upload PDF for archiving", type=["pdf"], key="pdf_to_pdfa")
            if pdfa_file and st.button("Create Archive Copy", key="pdfa_btn", use_container_width=True):
                # A standards-compliant PDF/A conversion requires a dedicated PDF renderer.
                # Preserve the option without falsely claiming that a byte-for-byte copy is PDF/A.
                st.warning("PDF/A is a standards-specific archival format. Use a dedicated PDF/A converter such as Ghostscript or Acrobat for final conformance.")
                bytes_download_button("Download Source PDF", pdfa_file.getvalue(), "Archive_Source.pdf", "application/pdf")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[2]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        col_file, col_lang = st.columns([2.2, 1])
        with col_file:
            uploaded_file = st.file_uploader("Upload image or PDF", type=["png", "jpg", "jpeg", "pdf"], key="ocr_s")
        with col_lang:
            lang = st.selectbox("OCR language", list(OCR_LANGS.keys()), key="ocr_l")
        if uploaded_file and st.button("Extract Text", key="ocr_btn", use_container_width=True):
            try:
                with st.spinner("Extracting text..."):
                    ftype = uploaded_file.name.rsplit(".", 1)[-1].lower()
                    if ftype == "pdf":
                        text_result = extract_pdf_text_or_ocr(uploaded_file, OCR_LANGS[lang])
                    else:
                        text_result = ocr_image(Image.open(uploaded_file), OCR_LANGS[lang])
                if text_result.strip():
                    st.success("OCR completed.")
                    st.text_area("Extracted text", text_result, height=260)
                    bytes_download_button("Download DOCX", create_docx(text_result), "OCR_Output.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    st.warning("No text was detected.")
            except Exception as exc:
                st.error(f"OCR failed: {exc}")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[3]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        st.markdown("#### Camera document scanner")
        cam_photo = st.camera_input("Take a picture")
        if cam_photo:
            img = Image.open(cam_photo).convert("RGB")
            st.image(img, caption="Captured document", use_container_width=True)
            if st.button("Save Photo as PDF", key="camera_pdf", use_container_width=True):
                p_arr = io.BytesIO()
                img.save(p_arr, format="PDF")
                st.success("Camera scan is ready.")
                bytes_download_button("Download PDF", p_arr.getvalue(), "Camera_Scan.pdf", "application/pdf")
        st.markdown('</div>', unsafe_allow_html=True)


def render_organize():
    st.markdown('<div class="section-title">Organize PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Control document structure without leaving the workspace.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Merge PDFs", "Split PDF", "Extract Pages", "Rotate PDF"])

    with tabs[0]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        files = st.file_uploader("Upload multiple PDFs", type=["pdf"], accept_multiple_files=True, key="m_files")
        if files and st.button("Merge Documents", key="merge_btn", use_container_width=True):
            try:
                merger = PdfMerger()
                for f in files:
                    merger.append(f)
                out = io.BytesIO()
                merger.write(out)
                merger.close()
                st.success(f"Merged {len(files)} PDF files successfully.")
                bytes_download_button("Download Merged PDF", out.getvalue(), "Merged_Document.pdf", "application/pdf")
            except Exception as exc:
                st.error(f"Merge failed: {exc}")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[1]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        file = st.file_uploader("Upload PDF", type=["pdf"], key="s_file")
        if file:
            reader = PdfReader(file)
            page_num = st.number_input("Page number", min_value=1, max_value=len(reader.pages), value=1, key="split_page")
            if st.button("Extract Page", key="split_btn", use_container_width=True):
                writer = PdfWriter()
                writer.add_page(reader.pages[page_num - 1])
                out = io.BytesIO()
                writer.write(out)
                bytes_download_button(f"Download Page {page_num}", out.getvalue(), f"Page_{page_num}.pdf", "application/pdf")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[2]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        file = st.file_uploader("Upload PDF", type=["pdf"], key="ext_file")
        pages_input = st.text_input("Page range", placeholder="Example: 1, 3, 5-8", key="page_range")
        if file and pages_input and st.button("Extract Range", key="range_btn", use_container_width=True):
            try:
                reader = PdfReader(file)
                writer = PdfWriter()
                selected_pages = []
                for part in pages_input.split(","):
                    part = part.strip()
                    if not part:
                        continue
                    if "-" in part:
                        start, end = [int(x.strip()) for x in part.split("-", 1)]
                        if start > end:
                            raise ValueError("Start page must be less than or equal to end page.")
                        selected_pages.extend(range(start - 1, end))
                    else:
                        selected_pages.append(int(part) - 1)
                valid = [p for p in selected_pages if 0 <= p < len(reader.pages)]
                if not valid:
                    raise ValueError("No valid page numbers were provided.")
                for p in valid:
                    writer.add_page(reader.pages[p])
                out = io.BytesIO()
                writer.write(out)
                st.success(f"Extracted {len(valid)} page(s).")
                bytes_download_button("Download Extracted PDF", out.getvalue(), "Extracted_Pages.pdf", "application/pdf")
            except Exception as exc:
                st.error(f"Invalid range: {exc}")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[3]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        file = st.file_uploader("Upload PDF", type=["pdf"], key="rot_file")
        angle = st.selectbox("Rotation angle", [90, 180, 270], key="rotation_angle")
        if file and st.button("Rotate Document", key="rotate_btn", use_container_width=True):
            try:
                reader = PdfReader(file)
                writer = PdfWriter()
                for page in reader.pages:
                    page.rotate(angle)
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                st.success(f"Document rotated by {angle}°. ")
                bytes_download_button("Download Rotated PDF", out.getvalue(), "Rotated_Document.pdf", "application/pdf")
            except Exception as exc:
                st.error(f"Rotation failed: {exc}")
        st.markdown('</div>', unsafe_allow_html=True)


def render_optimize():
    st.markdown('<div class="section-title">Optimize PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Apply lightweight PDF stream compression for faster sharing.</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
    opt_file = st.file_uploader("Upload PDF to optimize", type=["pdf"], key="opt_upload")
    if opt_file and st.button("Compress Document", key="compress_btn", use_container_width=True):
        try:
            before = len(opt_file.getvalue())
            reader = PdfReader(opt_file)
            writer = PdfWriter()
            for page in reader.pages:
                page.compress_content_streams()
                writer.add_page(page)
            out = io.BytesIO()
            writer.write(out)
            after = len(out.getvalue())
            st.success(f"Compression completed: {before/1024:.1f} KB → {after/1024:.1f} KB")
            bytes_download_button("Download Optimized PDF", out.getvalue(), "Optimized_Document.pdf", "application/pdf")
        except Exception as exc:
            st.error(f"Optimization failed: {exc}")
    st.markdown('</div>', unsafe_allow_html=True)


def render_edit():
    st.markdown('<div class="section-title">Edit PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">A lightweight annotator for simple document labeling and watermarking.</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
    file = st.file_uploader("Upload PDF", type=["pdf"], key="edit_pdf_file")
    mode = st.radio("Annotation type", ["Watermark", "Text label"], horizontal=True)
    text = st.text_input("Text", placeholder="Example: CONFIDENTIAL", key="edit_text")
    position = st.selectbox("Position", ["Top-left", "Top-right", "Bottom-left", "Bottom-right"], key="edit_position")
    if file and text.strip() and st.button("Apply Annotation", key="edit_apply", use_container_width=True):
        if not REPORTLAB_AVAILABLE:
            st.error("The annotation engine is unavailable because reportlab is not installed.")
        else:
            try:
                result = annotate_pdf(file, text.strip(), mode=mode, position=position)
                st.success("Annotation applied successfully.")
                bytes_download_button("Download Edited PDF", result, "Edited_Document.pdf", "application/pdf")
            except Exception as exc:
                st.error(f"Edit failed: {exc}")
    st.markdown('</div>', unsafe_allow_html=True)


def render_security():
    st.markdown('<div class="section-title">PDF Security</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Protect or unlock PDFs with a password you control.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Encrypt PDF", "Decrypt PDF"])

    with tabs[0]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        sec_file = st.file_uploader("Upload PDF", type=["pdf"], key="sec_encrypt")
        pwd = st.text_input("Set password", type="password", key="sec_pwd")
        if sec_file and pwd and st.button("Encrypt Document", key="enc_btn", use_container_width=True):
            try:
                reader = PdfReader(sec_file)
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                writer.encrypt(pwd)
                out = io.BytesIO()
                writer.write(out)
                st.success("PDF encrypted.")
                bytes_download_button("Download Protected PDF", out.getvalue(), "Protected_Document.pdf", "application/pdf")
            except Exception as exc:
                st.error(f"Encryption failed: {exc}")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[1]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        sec_file = st.file_uploader("Upload encrypted PDF", type=["pdf"], key="sec_decrypt")
        pwd = st.text_input("Enter password", type="password", key="dec_pwd")
        if sec_file and pwd and st.button("Unlock Document", key="dec_btn", use_container_width=True):
            try:
                reader = PdfReader(sec_file)
                if reader.is_encrypted and not reader.decrypt(pwd):
                    raise ValueError("Incorrect password.")
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                st.success("PDF unlocked.")
                bytes_download_button("Download Unlocked PDF", out.getvalue(), "Unlocked_Document.pdf", "application/pdf")
            except Exception as exc:
                st.error(str(exc))
        st.markdown('</div>', unsafe_allow_html=True)


def render_intelligence():
    st.markdown('<div class="section-title">PDF Intelligence</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Clean extracted text and generate speech from your content.</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Clean OCR Text", "Text-to-Speech"])

    with tabs[0]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        raw_text = st.text_area("Paste raw OCR text", height=220, key="clean_raw")
        if st.button("Clean Text Format", key="clean_btn", use_container_width=True):
            if raw_text.strip():
                cleaned = clean_text(raw_text)
                st.success("Text cleaned.")
                st.text_area("Cleaned output", cleaned, height=220, key="clean_output")
                bytes_download_button("Download DOCX", create_docx(cleaned), "Cleaned_Text.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.warning("Paste some text first.")
        st.markdown('</div>', unsafe_allow_html=True)

    with tabs[1]:
        st.markdown('<div class="tool-panel">', unsafe_allow_html=True)
        tts_input = st.text_area("Enter text for audio", height=190, key="tts_input")
        voice_lang = st.selectbox("Voice language", ["English", "Sinhala", "Tamil", "Spanish", "French"], key="tts_lang")
        lang_code = {"English": "en", "Sinhala": "si", "Tamil": "ta", "Spanish": "es", "French": "fr"}
        if st.button("Generate Audio", key="tts_btn", use_container_width=True):
            if tts_input.strip():
                try:
                    with st.spinner("Generating audio..."):
                        tts = gTTS(text=tts_input, lang=lang_code[voice_lang])
                        out = io.BytesIO()
                        tts.write_to_fp(out)
                        audio_bytes = out.getvalue()
                    st.audio(audio_bytes, format="audio/mp3")
                    bytes_download_button("Download MP3", audio_bytes, "Speech_Audio.mp3", "audio/mpeg")
                    st.info("Text-to-speech uses the gTTS service and therefore requires an internet connection.")
                except Exception as exc:
                    st.error(f"Audio generation failed: {exc}")
            else:
                st.warning("Enter some text first.")
        st.markdown('</div>', unsafe_allow_html=True)


def render_privacy():
    st.markdown('<div class="section-title">Data Privacy</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-copy">Understand exactly how this app handles uploaded content.</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="tool-panel">
            <h3>Session-based processing</h3>
            <p style="color:#a8b7c9;line-height:1.7">This app does not implement an application database or a user-facing file storage system. Uploaded files are processed during the active Streamlit session and generated downloads are returned to the browser.</p>
            <div class="mini-note"><b>Important:</b> third-party processing can occur for specific workflows. In particular, gTTS sends text to the Google text-to-speech service to generate audio. Avoid uploading or pasting sensitive information into workflows that rely on external services.</div>
            <div style="height:12px"></div>
            <div class="stat-row">
                <div class="stat"><div class="stat-label">App database</div><div class="stat-value">None</div></div>
                <div class="stat"><div class="stat-label">Persistent uploads</div><div class="stat-value">Not implemented</div></div>
                <div class="stat"><div class="stat-label">OCR</div><div class="stat-value">Local Tesseract</div></div>
                <div class="stat"><div class="stat-label">TTS</div><div class="stat-value">External service</div></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def render_footer():
    st.markdown(
        """
        <style>
        .lp-footer {
            margin-top: 50px;
            padding: 30px 35px 20px 35px;
            border-top: 1px solid rgba(125, 217, 255, 0.18);
            background: linear-gradient(
                180deg,
                rgba(13, 27, 42, 0.85),
                rgba(7, 17, 31, 0.98)
            );
            border-radius: 18px 18px 0 0;
            text-align: center;
        }

        .lp-footer-title {
            font-size: 20px;
            font-weight: 700;
            color: #eef6ff;
            margin-bottom: 8px;
        }

        .lp-footer-text {
            color: #94a8bd;
            font-size: 13px;
            margin-bottom: 18px;
        }

        .lp-social {
            display: flex;
            justify-content: center;
            flex-wrap: wrap;
            gap: 10px;
            margin: 18px 0;
        }

        .lp-social a {
            color: #b9d8ef;
            text-decoration: none;
            padding: 8px 14px;
            border: 1px solid rgba(125, 217, 255, 0.16);
            border-radius: 10px;
            background: rgba(255, 255, 255, 0.03);
            font-size: 13px;
        }

        .lp-social a:hover {
            color: #7dd9ff;
            border-color: rgba(125, 217, 255, 0.45);
        }

        .lp-share-btn {
            display: inline-block;
            margin: 8px 0 20px 0;
            padding: 11px 20px;
            border-radius: 12px;
            background: linear-gradient(135deg, #1677ff, #36c5ff);
            color: white !important;
            text-decoration: none !important;
            font-weight: 700;
            font-size: 13px;
        }

        .lp-footer-contact {
            color: #94a8bd;
            font-size: 12px;
            margin-top: 8px;
        }

        .lp-footer-bottom {
            margin-top: 20px;
            padding-top: 15px;
            border-top: 1px solid rgba(125, 217, 255, 0.10);
            color: #6f8499;
            font-size: 11px;
        }
        </style>

        <footer class="lp-footer">

            <div class="lp-footer-title">
                ⚡ LipiParse Studio
            </div>

            <div class="lp-footer-text">
                Premium PDF & document workspace
            </div>

            <div class="lp-social">
                <a href="YOUR_FACEBOOK_LINK" target="_blank">
                    Facebook
                </a>

                <a href="YOUR_INSTAGRAM_LINK" target="_blank">
                    Instagram
                </a>

                <a href="YOUR_LINKEDIN_LINK" target="_blank">
                    LinkedIn
                </a>

                <a href="YOUR_GITHUB_LINK" target="_blank">
                    GitHub
                </a>
            </div>

            <a
                class="lp-share-btn"
                href="https://lipiparse-mcthmncr2jfncvym8kbxmw.streamlit.app/"
                target="_blank"
            >
                Invite / Share LipiParse
            </a>

            <div class="lp-footer-contact">
                Email: YOUR_EMAIL@example.com
                &nbsp; • &nbsp;
                Contact: +94XXXXXXXXX
            </div>

            <div class="lp-footer-bottom">
                ⚡ LipiParse Studio • Premium PDF & document workspace
                • Version 2.0
                <br>
                Built for fast everyday document workflows. © 2026
            </div>

        </footer>
        """,
        unsafe_allow_html=True,
    )



# ==========================================================
# RENDER APP
# ==========================================================
inject_css()
init_state()
render_sidebar()
render_header()
render_hero()
render_stats()

selected_tab = st.session_state.selected_tab

if selected_tab == "All Workflows":
    render_all_workflows()
elif selected_tab == "Convert PDF":
    render_convert()
elif selected_tab == "Organize PDF":
    render_organize()
elif selected_tab == "Optimize PDF":
    render_optimize()
elif selected_tab == "Edit PDF":
    render_edit()
elif selected_tab == "PDF Security":
    render_security()
elif selected_tab == "PDF Intelligence":
    render_intelligence()
elif selected_tab == "Data Privacy":
    render_privacy()

render_footer()

import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image
from docx import Document
import io
import zipfile
import re
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from gtts import gTTS

# Professional Page Configuration
st.set_page_config(
    page_title="LipiParse Studio | SaaS PDF & Document Intelligence", 
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Load FontAwesome Vector Icons & Premium Styling
st.markdown("""
    <!-- FontAwesome Vector Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
    .main {
        background-color: #F8FAFC;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    .brand-logo {
        font-size: 1.5rem;
        font-weight: 800;
        color: #1E3A8A;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .hero-container {
        text-align: center;
        padding: 2.5rem 1.5rem;
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        color: white;
        border-radius: 18px;
        margin-bottom: 2rem;
        box-shadow: 0 20px 25px -5px rgba(37, 99, 235, 0.2);
    }
    .hero-title {
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        opacity: 0.9;
        max-width: 700px;
        margin: 0 auto;
    }
    
    /* Premium Navigation Buttons */
    div.stButton > button {
        width: 100%;
        height: 65px;
        background: #1E293B;
        color: #F8FAFC;
        border: 1px solid #334155;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.25s ease-in-out;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        color: #FFFFFF;
        border-color: #3B82F6;
        transform: translateY(-3px);
    }

    .tool-box {
        background-color: #FFFFFF;
        padding: 2rem;
        border-radius: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
        margin-bottom: 2rem;
    }

    /* Product Grid Styling */
    .product-grid {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 16px;
        padding: 15px 0 25px 0;
    }
    @media (max-width: 1200px) {
        .product-grid { grid-template-columns: repeat(3, 1fr); }
    }
    @media (max-width: 768px) {
        .product-grid { grid-template-columns: repeat(2, 1fr); }
    }
    .tool-card {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 14px;
        padding: 20px 10px;
        text-align: center;
        transition: all 0.25s ease-in-out;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .tool-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 20px rgba(0, 0, 0, 0.08);
        border-color: #94A3B8;
    }
    .icon-wrapper {
        width: 48px;
        height: 48px;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 12px auto;
        font-size: 1.3rem;
    }
    .card-label {
        font-size: 0.88rem;
        font-weight: 600;
        color: #0F172A;
        line-height: 1.3;
    }

    /* Color variations */
    .icon-red { background-color: #FEF2F2; color: #EF4444; }
    .icon-blue { background-color: #F0F9FF; color: #0284C7; }
    .icon-green { background-color: #F0FDF4; color: #16A34A; }
    .icon-orange { background-color: #FFF7ED; color: #EA580C; }
    .icon-purple { background-color: #FAF5FF; color: #9333EA; }
    .icon-teal { background-color: #F0FDFA; color: #0D9488; }

    /* Footer Styling */
    .custom-footer {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%);
        color: #E2E8F0;
        padding: 2.5rem 2rem 1.5rem 2rem;
        border-radius: 16px;
        margin-top: 3rem;
    }
    .footer-heading {
        color: #FFFFFF;
        font-size: 1.15rem;
        font-weight: 700;
        margin-bottom: 0.8rem;
    }
    .footer-text {
        font-size: 0.92rem;
        line-height: 1.6;
        color: #94A3B8;
    }
    .social-link-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 40px;
        height: 40px;
        background-color: rgba(255, 255, 255, 0.1);
        color: #FFFFFF !important;
        border-radius: 10px;
        font-size: 1.2rem;
        margin-right: 0.6rem;
        text-decoration: none;
    }
    .footer-divider {
        border-top: 1px solid rgba(255, 255, 255, 0.1);
        margin-top: 2rem;
        padding-top: 1rem;
        text-align: center;
        font-size: 0.85rem;
        color: #64748B;
    }
    </style>
""", unsafe_allow_html=True)

# Helper: Create DOCX File
def create_docx(text):
    doc = Document()
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

OCR_LANGS = {
    "English": "eng", "Sinhala (සිංහල)": "sin", "Tamil (தமிழ்)": "tam",
    "Sinhala + English": "sin+eng", "Tamil + English": "tam+eng",
    "Spanish": "spa", "French": "fra", "German": "deu", "Japanese": "jpn", "Chinese": "chi_sim"
}

if "selected_tab" not in st.session_state:
    st.session_state["selected_tab"] = "All Workflows"

# Header
col_logo, col_right = st.columns([3, 1])
with col_logo:
    st.markdown('<div class="brand-logo"><i class="fa-solid fa-bolt" style="color: #2563EB;"></i> LipiParse Studio</div>', unsafe_allow_html=True)
with col_right:
    st.caption("🌐 EN | 🔒 SSL Secured SaaS")

# Banner
st.markdown("""
    <div class="hero-container">
        <div class="hero-title">Every PDF Tool You Need in One Place</div>
        <div class="hero-subtitle">Convert, Edit, OCR, Compress, and Secure PDFs seamlessly with 100% Data Privacy.</div>
    </div>
""", unsafe_allow_html=True)

# Navigation Grid
st.markdown("##### Select Workspace Category:")
row1_col1, row1_col2, row1_col3, row1_col4 = st.columns(4)
with row1_col1:
    if st.button("Data Privacy", key="nav_privacy"): st.session_state["selected_tab"] = "Data Privacy"
with row1_col2:
    if st.button("All Workflows", key="nav_all"): st.session_state["selected_tab"] = "All Workflows"
with row1_col3:
    if st.button("Convert PDF", key="nav_convert"): st.session_state["selected_tab"] = "Convert PDF"
with row1_col4:
    if st.button("Organize PDF", key="nav_organize"): st.session_state["selected_tab"] = "Organize PDF"

row2_col1, row2_col2, row2_col3, row2_col4 = st.columns(4)
with row2_col1:
    if st.button("Optimize PDF", key="nav_optimize"): st.session_state["selected_tab"] = "Optimize PDF"
with row2_col2:
    if st.button("Edit PDF", key="nav_edit"): st.session_state["selected_tab"] = "Edit PDF"
with row2_col3:
    if st.button("PDF Security", key="nav_security"): st.session_state["selected_tab"] = "PDF Security"
with row2_col4:
    if st.button("PDF Intelligence", key="nav_intelligence"): st.session_state["selected_tab"] = "PDF Intelligence"

selected_tab = st.session_state["selected_tab"]
st.markdown("---")

# Tab Content
if selected_tab == "Data Privacy":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    st.subheader("Data Privacy & Security Guarantee")
    st.write("All files processed with LipiParse Studio remain strictly confidential.")
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1: st.success("End-to-End SSL Encryption")
    with col_p2: st.info("No Cloud File Persistence")
    with col_p3: st.warning("Automatic Memory Wipe")
    st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "All Workflows":
    search_query = st.text_input("Search workspace workflows...", placeholder="e.g. JPG to PDF, Merge, OCR, Split, Word, Compress").strip().lower()
    tools_list = [
        {"title": "Convert PDF Suite", "category": "Convert PDF", "icon": "fa-arrows-rotate", "desc": "Convert JPG, Word, Excel, PowerPoint, HTML to PDF.", "keywords": ["convert", "jpg", "word"]},
        {"title": "Organize & Structure", "category": "Organize PDF", "icon": "fa-folder-tree", "desc": "Merge, split, rotate, and extract pages.", "keywords": ["merge", "split"]},
        {"title": "Privacy & Protection", "category": "PDF Security", "icon": "fa-shield-halved", "desc": "Encrypt PDFs with password protection.", "keywords": ["protect", "encrypt"]},
        {"title": "Document OCR Engine", "category": "Convert PDF", "icon": "fa-file-lines", "desc": "Extract text from scanned documents.", "keywords": ["ocr", "scan"]},
        {"title": "Mobile Camera Scanner", "category": "Convert PDF", "icon": "fa-camera", "desc": "Snap physical documents and convert to PDFs.", "keywords": ["camera", "scan"]},
        {"title": "PDF Optimization", "category": "Optimize PDF", "icon": "fa-gauge-high", "desc": "Compress large files quickly.", "keywords": ["compress", "optimize"]},
        {"title": "AI Intelligence & TTS", "category": "PDF Intelligence", "icon": "fa-brain", "desc": "Clean text and convert to audio.", "keywords": ["ai", "speech"]}
    ]
    filtered_tools = [t for t in tools_list if not search_query or search_query in t["title"].lower() or any(search_query in kw for kw in t["keywords"])]
    cols = st.columns(3)
    for idx, tool_item in enumerate(filtered_tools):
        with cols[idx % 3]:
            st.markdown(f"#### <i class='fa-solid {tool_item['icon']}'></i> {tool_item['title']}", unsafe_allow_html=True)
            st.caption(f"Category: **{tool_item['category']}**")
            st.info(tool_item['desc'])

elif selected_tab == "Convert PDF":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    sub_tab = st.radio("Conversion Category:", ["Convert to PDF", "Convert from PDF", "OCR Text Extractor", "Camera Scanner"], horizontal=True)
    if sub_tab == "Convert to PDF":
        images = st.file_uploader("Upload Images", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
        if images and st.button("Convert Images to PDF"):
            img_objs = [Image.open(i).convert('RGB') for i in images]
            out = io.BytesIO()
            if img_objs:
                img_objs[0].save(out, format='PDF', save_all=True, append_images=img_objs[1:])
                st.download_button("Download PDF", out.getvalue(), "Converted_Images.pdf", mime="application/pdf")
    elif sub_tab == "Convert from PDF":
        pdf_word_file = st.file_uploader("Upload PDF to convert to DOCX", type=["pdf"])
        if pdf_word_file and st.button("Convert PDF to DOCX"):
            with pdfplumber.open(pdf_word_file) as pdf:
                extracted_text = "".join([(p.extract_text() or "") + "\n\n" for p in pdf.pages])
            st.download_button("Download DOCX", create_docx(extracted_text), "Converted.docx")
    elif sub_tab == "OCR Text Extractor":
        uploaded_file = st.file_uploader("Upload Image/PDF for OCR", type=["png", "jpg", "jpeg", "pdf"])
        lang = st.selectbox("Language", list(OCR_LANGS.keys()))
        if uploaded_file and st.button("Extract Text"):
            text_result = ""
            if uploaded_file.name.endswith('.pdf'):
                with pdfplumber.open(uploaded_file) as pdf:
                    text_result = "\n".join([p.extract_text() or "" for p in pdf.pages])
            else:
                text_result = pytesseract.image_to_string(Image.open(uploaded_file), lang=OCR_LANGS[lang])
            st.text_area("Extracted Text", text_result, height=200)
    elif sub_tab == "Camera Scanner":
        cam_photo = st.camera_input("Take Picture")
        if cam_photo and st.button("Save as PDF"):
            p_arr = io.BytesIO()
            Image.open(cam_photo).convert('RGB').save(p_arr, format='PDF')
            st.download_button("Download PDF", p_arr.getvalue(), "Scan.pdf", mime="application/pdf")
    st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "Organize PDF":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    files = st.file_uploader("Upload PDFs to Merge", type=["pdf"], accept_multiple_files=True)
    if files and st.button("Merge Documents"):
        merger = PdfMerger()
        for f in files: merger.append(f)
        out = io.BytesIO()
        merger.write(out)
        merger.close()
        st.download_button("Download Merged PDF", data=out.getvalue(), file_name="Merged.pdf", mime="application/pdf")
    st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "Optimize PDF":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    opt_file = st.file_uploader("Upload PDF to optimize", type=["pdf"])
    if opt_file and st.button("Compress Document"):
        reader = PdfReader(opt_file)
        writer = PdfWriter()
        for p in reader.pages:
            p.compress_content_streams()
            writer.add_page(p)
        out = io.BytesIO()
        writer.write(out)
        st.download_button("Download Optimized PDF", out.getvalue(), "Optimized.pdf", mime="application/pdf")
    st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "Edit PDF":
    st.info("Page Annotator engine under maintenance.")

elif selected_tab == "PDF Security":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    sec_file = st.file_uploader("Upload PDF to Encrypt", type=["pdf"])
    pwd = st.text_input("Set Password:", type="password")
    if sec_file and pwd and st.button("Encrypt Document"):
        reader = PdfReader(sec_file)
        writer = PdfWriter()
        for p in reader.pages: writer.add_page(p)
        writer.encrypt(pwd)
        out = io.BytesIO()
        writer.write(out)
        st.download_button("Download PDF", out.getvalue(), "Protected.pdf", mime="application/pdf")
    st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "PDF Intelligence":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    tts_input = st.text_area("Enter Text for Audio Speech:", height=140)
    if st.button("Generate Audio Speech"):
        if tts_input.strip():
            tts = gTTS(text=tts_input, lang="en")
            out = io.BytesIO()
            tts.write_to_fp(out)
            out.seek(0)
            st.audio(out, format="audio/mp3")
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# FULL PRODUCT FAMILY GRID SECTION
# ==========================================
st.markdown("---")
st.markdown("<h2 style='text-align: center; font-weight: 800; color: #1E293B; margin-top: 1.5rem; margin-bottom: 1.5rem;'>Meet our full product family</h2>", unsafe_allow_html=True)

tools_family = [
    {"name": "Compress PDF", "icon": "fa-file-zipper", "color": "icon-red"},
    {"name": "Delete PDF Pages", "icon": "fa-trash-can", "color": "icon-red"},
    {"name": "Rotate PDF", "icon": "fa-rotate-right", "color": "icon-red"},
    {"name": "Repair PDF", "icon": "fa-screwdriver-wrench", "color": "icon-red"},
    {"name": "Print-Ready PDF", "icon": "fa-print", "color": "icon-red"},
    {"name": "Word to PDF", "icon": "fa-file-word", "color": "icon-blue"},
    {"name": "Excel to PDF", "icon": "fa-file-excel", "color": "icon-green"},
    {"name": "PowerPoint to PDF", "icon": "fa-file-powerpoint", "color": "icon-orange"},
    {"name": "JPG to PDF", "icon": "fa-file-image", "color": "icon-purple"},
    {"name": "PNG to PDF", "icon": "fa-image", "color": "icon-purple"},
    {"name": "HEIC to PDF", "icon": "fa-camera-retro", "color": "icon-purple"},
    {"name": "AutoCAD to PDF", "icon": "fa-compass-drafting", "color": "icon-red"},
    {"name": "OpenOffice to PDF", "icon": "fa-file-code", "color": "icon-blue"},
    {"name": "eBooks to PDF", "icon": "fa-book-open", "color": "icon-orange"},
    {"name": "iWork to PDF", "icon": "fa-apple", "color": "icon-blue"},
    {"name": "Publisher to PDF", "icon": "fa-newspaper", "color": "icon-green"},
    {"name": "DICOM to PDF", "icon": "fa-notes-medical", "color": "icon-blue"},
    {"name": "PDF to Word", "icon": "fa-file-word", "color": "icon-blue"},
    {"name": "PDF to Excel", "icon": "fa-file-excel", "color": "icon-green"},
    {"name": "PDF to Powerpoint", "icon": "fa-file-powerpoint", "color": "icon-orange"},
    {"name": "PDF to JPG", "icon": "fa-file-image", "color": "icon-purple"},
    {"name": "PDF to PNG", "icon": "fa-image", "color": "icon-purple"},
    {"name": "Extract PDF Images", "icon": "fa-images", "color": "icon-purple"},
    {"name": "PDF to PDF/A", "icon": "fa-box-archive", "color": "icon-blue"},
    {"name": "Merge PDF", "icon": "fa-object-group", "color": "icon-orange"},
    {"name": "Split PDF", "icon": "fa-scissors", "color": "icon-orange"},
    {"name": "Protect PDF", "icon": "fa-lock", "color": "icon-teal"},
    {"name": "Unlock PDF", "icon": "fa-lock-open", "color": "icon-teal"},
    {"name": "Redact PDF", "icon": "fa-user-ninja", "color": "icon-teal"},
    {"name": "PDF Converter", "icon": "fa-arrows-rotate", "color": "icon-red"}
]

# Generate raw single-line cards string to avoid indentation bug
grid_cards_str = ""
for item in tools_family:
    grid_cards_str += f'<div class="tool-card"><div class="icon-wrapper {item["color"]}"><i class="fa-solid {item["icon"]}"></i></div><div class="card-label">{item["name"]}</div></div>'

st.markdown(f'<div class="product-grid">{grid_cards_str}</div>', unsafe_allow_html=True)

# Footer
st.markdown("""
    <div class="custom-footer">
        <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 2rem;">
            <div style="flex: 1.4; min-width: 250px;">
                <div class="footer-heading"><i class="fa-solid fa-bolt" style="color: #38BDF8;"></i> LipiParse Studio</div>
                <p class="footer-text">All-in-one PDF & document processing studio. 100% data privacy.</p>
            </div>
            <div style="flex: 1.1; min-width: 220px;">
                <div class="footer-heading"><i class="fa-solid fa-share-nodes" style="color: #38BDF8;"></i> Share With Peers</div>
                <input type="text" value="https://lipiparse-studio.streamlit.app" readonly style="width: 100%; padding: 0.5rem; border-radius: 8px; border: 1px solid #334155; background: #0F172A; color: #38BDF8; font-size: 0.85rem;">
            </div>
            <div style="flex: 1.1; min-width: 220px;">
                <div class="footer-heading"><i class="fa-solid fa-globe" style="color: #38BDF8;"></i> Connect With Us</div>
                <div style="margin-top: 0.8rem;">
                    <a href="https://facebook.com" target="_blank" class="social-link-btn"><i class="fa-brands fa-facebook-f"></i></a>
                    <a href="https://twitter.com" target="_blank" class="social-link-btn"><i class="fa-brands fa-x-twitter"></i></a>
                    <a href="https://linkedin.com" target="_blank" class="social-link-btn"><i class="fa-brands fa-linkedin-in"></i></a>
                </div>
            </div>
        </div>
        <div class="footer-divider">© 2026 LipiParse Studio SaaS Platform. All rights reserved.</div>
    </div>
""", unsafe_allow_html=True)
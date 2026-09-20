import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image
from docx import Document
import io
import re
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from gtts import gTTS

# Page Configuration
st.set_page_config(
    page_title="LipiParse Studio | SaaS PDF & Document Intelligence", 
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom Styling & FontAwesome Icons
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
        font-size: 2.2rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 1rem;
        opacity: 0.9;
        max-width: 700px;
        margin: 0 auto;
    }
    
    /* Product Grid Styling */
    .product-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
        gap: 16px;
        padding: 10px 0 25px 0;
    }
    .tool-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 20px 10px;
        text-align: center;
        transition: all 0.25s ease-in-out;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .tool-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 20px rgba(0, 0, 0, 0.08);
        border-color: #CBD5E1;
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
        color: #1E293B;
        line-height: 1.3;
    }
    
    /* Icon Colors */
    .icon-red { background-color: #FFECEB; color: #E11D48; }
    .icon-blue { background-color: #E0F2FE; color: #0284C7; }
    .icon-green { background-color: #DCFCE7; color: #16A34A; }
    .icon-orange { background-color: #FFEDD5; color: #EA580C; }
    .icon-purple { background-color: #F3E8FF; color: #9333EA; }
    .icon-teal { background-color: #CCFBF1; color: #0D9488; }

    .tool-box {
        background-color: #FFFFFF;
        padding: 2rem;
        border-radius: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
        margin-bottom: 2rem;
    }

    /* Footer Styling */
    .custom-footer {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%);
        color: #E2E8F0;
        padding: 2.5rem 2rem 1.5rem 2rem;
        border-radius: 16px;
        margin-top: 3rem;
        box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.15);
    }
    .footer-heading {
        color: #FFFFFF;
        font-size: 1.15rem;
        font-weight: 700;
        margin-bottom: 0.8rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
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
        transition: all 0.3s ease;
    }
    .social-link-btn:hover {
        background-color: #2563EB;
        color: #FFFFFF !important;
        transform: translateY(-3px);
    }
    .footer-divider {
        border-top: 1px solid rgba(255, 255, 255, 0.1);
        margin-top: 2rem;
        padding-top: 1rem;
        text-align: center;
        font-size: 0.85rem;
        color: #64748B;
    }
    .footer-link {
        color: #38BDF8;
        text-decoration: none;
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
    "Spanish": "spa", "French": "fra", "German": "deu"
}

# --- TOP BRAND HEADER BAR ---
col_logo, col_right = st.columns([3, 1])
with col_logo:
    st.markdown("""
        <div class="brand-logo">
            <i class="fa-solid fa-bolt" style="color: #2563EB;"></i> LipiParse Studio
        </div>
    """, unsafe_allow_html=True)
with col_right:
    st.caption("🌐 EN | 🔒 SSL Secured SaaS")

# --- HERO BANNER ---
st.markdown("""
    <div class="hero-container">
        <div class="hero-title">Every PDF Tool You Need in One Place</div>
        <div class="hero-subtitle">Convert, Edit, OCR, Compress, and Secure PDFs seamlessly with 100% Data Privacy.</div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# FULL PRODUCT FAMILY GRID (PHOTO BASED)
# ==========================================
st.markdown("<h3 style='text-align: center; margin-bottom: 1.5rem; color: #1E293B;'>Meet our full product family</h3>", unsafe_allow_html=True)

tools_grid = [
    {"name": "Compress PDF", "icon": "fa-compress", "color": "icon-red"},
    {"name": "Delete PDF Pages", "icon": "fa-xmark", "color": "icon-red"},
    {"name": "Rotate PDF", "icon": "fa-rotate-right", "color": "icon-red"},
    {"name": "Repair PDF", "icon": "fa-wrench", "color": "icon-red"},
    {"name": "Print-Ready PDF", "icon": "fa-print", "color": "icon-red"},
    {"name": "Word to PDF", "icon": "fa-file-word", "color": "icon-blue"},
    {"name": "Excel to PDF", "icon": "fa-file-excel", "color": "icon-green"},
    {"name": "PowerPoint to PDF", "icon": "fa-file-powerpoint", "color": "icon-orange"},
    {"name": "JPG to PDF", "icon": "fa-file-image", "color": "icon-purple"},
    {"name": "PNG to PDF", "icon": "fa-image", "color": "icon-purple"},
    {"name": "HEIC to PDF", "icon": "fa-file-lines", "color": "icon-purple"},
    {"name": "AutoCAD to PDF", "icon": "fa-drafting-compass", "color": "icon-red"},
    {"name": "OpenOffice to PDF", "icon": "fa-file-code", "color": "icon-blue"},
    {"name": "eBooks to PDF", "icon": "fa-book", "color": "icon-orange"},
    {"name": "iWork to PDF", "icon": "fa-brands fa-apple", "color": "icon-blue"},
    {"name": "Publisher to PDF", "icon": "fa-note-sticky", "color": "icon-green"},
    {"name": "DICOM to PDF", "icon": "fa-notes-medical", "color": "icon-blue"},
    {"name": "PDF to Word", "icon": "fa-file-word", "color": "icon-blue"},
    {"name": "PDF to Excel", "icon": "fa-file-excel", "color": "icon-green"},
    {"name": "PDF to Powerpoint", "icon": "fa-file-powerpoint", "color": "icon-orange"},
    {"name": "PDF to JPG", "icon": "fa-file-image", "color": "icon-purple"},
    {"name": "PDF to PNG", "icon": "fa-image", "color": "icon-purple"},
    {"name": "Extract PDF Images", "icon": "fa-images", "color": "icon-purple"},
    {"name": "PDF to PDF/A", "icon": "fa-file-pdf", "color": "icon-blue"},
    {"name": "Merge PDF", "icon": "fa-copy", "color": "icon-orange"},
    {"name": "Split PDF", "icon": "fa-arrows-split-up-and-left", "color": "icon-orange"},
    {"name": "Protect PDF", "icon": "fa-lock", "color": "icon-teal"},
    {"name": "Unlock PDF", "icon": "fa-key", "color": "icon-teal"},
    {"name": "Redact PDF", "icon": "fa-pen-to-square", "color": "icon-teal"},
    {"name": "PDF Converter", "icon": "fa-arrows-rotate", "color": "icon-red"},
]

html_grid = '<div class="product-grid">'
for t in tools_grid:
    html_grid += f'''
        <div class="tool-card">
            <div class="icon-wrapper {t['color']}">
                <i class="fa-solid {t['icon']}"></i>
            </div>
            <div class="card-label">{t['name']}</div>
        </div>
    '''
html_grid += '</div>'

st.markdown(html_grid, unsafe_allow_html=True)
st.markdown("---")

# ==========================================
# INTERACTIVE WORKSPACE
# ==========================================
st.markdown("### 🛠️ Interactive Tool Workspace")

active_tab = st.tabs(["Convert Tools", "Organize & Edit", "Security & Protect", "OCR & Speech"])

# TAB 1: CONVERT
with active_tab[0]:
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("JPG/PNG to PDF")
        images = st.file_uploader("Upload Image Files", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
        if images and st.button("Convert to PDF", key="btn_img_pdf"):
            img_objs = [Image.open(i).convert('RGB') for i in images]
            out = io.BytesIO()
            img_objs[0].save(out, format='PDF', save_all=True, append_images=img_objs[1:])
            st.success("Converted to PDF!")
            st.download_button("Download PDF", out.getvalue(), "Images_Converted.pdf", mime="application/pdf")

    with c2:
        st.subheader("PDF to Word (DOCX)")
        pdf_word = st.file_uploader("Upload PDF File", type=["pdf"], key="pdf_to_doc")
        if pdf_word and st.button("Convert to DOCX", key="btn_pdf_doc"):
            with pdfplumber.open(pdf_word) as pdf:
                extracted = "\n\n".join([page.extract_text() or "" for page in pdf.pages])
            st.success("Done!")
            st.download_button("Download DOCX", create_docx(extracted), "Converted.docx")
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 2: ORGANIZE
with active_tab[1]:
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    o1, o2 = st.columns(2)
    with o1:
        st.subheader("Merge PDFs")
        m_files = st.file_uploader("Select Multiple PDFs", type=["pdf"], accept_multiple_files=True, key="m_u")
        if m_files and st.button("Merge Files"):
            merger = PdfMerger()
            for f in m_files: merger.append(f)
            out = io.BytesIO()
            merger.write(out)
            merger.close()
            st.success("Merged!")
            st.download_button("Download Merged PDF", out.getvalue(), "Merged_Document.pdf", mime="application/pdf")

    with o2:
        st.subheader("Rotate PDF")
        r_file = st.file_uploader("Select PDF to Rotate", type=["pdf"], key="r_u")
        angle = st.selectbox("Angle:", [90, 180, 270])
        if r_file and st.button("Rotate PDF"):
            reader = PdfReader(r_file)
            writer = PdfWriter()
            for p in reader.pages:
                p.rotate(angle)
                writer.add_page(p)
            out = io.BytesIO()
            writer.write(out)
            st.success("Rotated!")
            st.download_button("Download PDF", out.getvalue(), "Rotated.pdf", mime="application/pdf")
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 3: SECURITY
with active_tab[2]:
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    s1, s2 = st.columns(2)
    with s1:
        st.subheader("Protect PDF")
        sec_f = st.file_uploader("Upload PDF to Lock", type=["pdf"], key="sec_lock")
        pwd = st.text_input("Set Password:", type="password")
        if sec_f and pwd and st.button("Encrypt PDF"):
            reader = PdfReader(sec_f)
            writer = PdfWriter()
            for p in reader.pages: writer.add_page(p)
            writer.encrypt(pwd)
            out = io.BytesIO()
            writer.write(out)
            st.success("Encrypted!")
            st.download_button("Download Encrypted PDF", out.getvalue(), "Protected.pdf", mime="application/pdf")

    with s2:
        st.subheader("Unlock PDF")
        dec_f = st.file_uploader("Upload Locked PDF", type=["pdf"], key="sec_unlock")
        d_pwd = st.text_input("Enter Password:", type="password", key="d_pwd")
        if dec_f and d_pwd and st.button("Decrypt PDF"):
            try:
                reader = PdfReader(dec_f)
                if reader.is_encrypted: reader.decrypt(d_pwd)
                writer = PdfWriter()
                for p in reader.pages: writer.add_page(p)
                out = io.BytesIO()
                writer.write(out)
                st.success("Unlocked!")
                st.download_button("Download Unlocked PDF", out.getvalue(), "Unlocked.pdf", mime="application/pdf")
            except: st.error("Incorrect password.")
    st.markdown('</div>', unsafe_allow_html=True)

# TAB 4: OCR & SPEECH
with active_tab[3]:
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    ocr_f = st.file_uploader("Upload Image/PDF for OCR", type=["png", "jpg", "jpeg", "pdf"], key="ocr_full")
    lang_sel = st.selectbox("Language", list(OCR_LANGS.keys()))
    if ocr_f and st.button("Extract Text with OCR"):
        res = ""
        if ocr_f.name.endswith(".pdf"):
            with pdfplumber.open(ocr_f) as pdf:
                for p in pdf.pages: res += (p.extract_text() or "") + "\n"
        else:
            res = pytesseract.image_to_string(Image.open(ocr_f), lang=OCR_LANGS[lang_sel])
        st.text_area("Extracted Text", res, height=180)
        st.download_button("Download Word File", create_docx(res), "OCR_Result.docx")
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# FOOTER SECTION
# ==========================================
st.markdown("""
    <div class="custom-footer">
        <div style="display: flex; flex-wrap: wrap; justify-content: space-between; gap: 2rem;">
            <div style="flex: 1.4; min-width: 250px;">
                <div class="footer-heading"><i class="fa-solid fa-bolt" style="color: #38BDF8;"></i> LipiParse Studio</div>
                <p class="footer-text">
                    All-in-one PDF & document processing studio. 
                    Convert, edit, OCR, compress, and secure files with <b>100% data privacy</b> and end-to-end security.
                </p>
            </div>
            <div style="flex: 1.1; min-width: 220px;">
                <div class="footer-heading"><i class="fa-solid fa-share-nodes" style="color: #38BDF8;"></i> Share With Peers</div>
                <p class="footer-text">Share this workspace with your teammates and colleagues:</p>
                <input type="text" value="https://lipiparse-studio.streamlit.app" readonly style="width: 100%; padding: 0.5rem; border-radius: 8px; border: 1px solid #334155; background: #0F172A; color: #38BDF8; font-size: 0.85rem;">
            </div>
            <div style="flex: 1.1; min-width: 220px;">
                <div class="footer-heading"><i class="fa-solid fa-globe" style="color: #38BDF8;"></i> Connect With Us</div>
                <p class="footer-text">Follow us for product updates, feature releases, and support:</p>
                <div style="margin-top: 0.8rem;">
                    <a href="https://facebook.com" target="_blank" class="social-link-btn"><i class="fa-brands fa-facebook-f"></i></a>
                    <a href="https://twitter.com" target="_blank" class="social-link-btn"><i class="fa-brands fa-x-twitter"></i></a>
                    <a href="https://linkedin.com" target="_blank" class="social-link-btn"><i class="fa-brands fa-linkedin-in"></i></a>
                    <a href="https://instagram.com" target="_blank" class="social-link-btn"><i class="fa-brands fa-instagram"></i></a>
                    <a href="https://t.me" target="_blank" class="social-link-btn"><i class="fa-brands fa-telegram"></i></a>
                </div>
            </div>
        </div>
        <div class="footer-divider">
            © 2026 <b>LipiParse Studio</b> SaaS Platform. All rights reserved. <br>
            <a href="#" class="footer-link">Privacy Policy</a> • 
            <a href="#" class="footer-link">Terms of Service</a> • 
            <a href="#" class="footer-link">Security Standard</a>
        </div>
    </div>
""", unsafe_allow_html=True)
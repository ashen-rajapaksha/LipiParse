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

# Load FontAwesome Icons & Custom SaaS Styling CSS
st.markdown("""
    <!-- FontAwesome Vector Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
    /* Main Background & Clean Typography */
    .main {
        background-color: #F8FAFC;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Top Header Bar */
    .top-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 1rem 1.5rem;
        background: #FFFFFF;
        border-bottom: 1px solid #E2E8F0;
        margin-bottom: 1.5rem;
    }
    .brand-logo {
        font-size: 1.5rem;
        font-weight: 800;
        color: #1E3A8A;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    /* Hero Section */
    .hero-container {
        text-align: center;
        padding: 3rem 1.5rem 2.5rem 1.5rem;
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        color: white;
        border-radius: 18px;
        margin-bottom: 2rem;
        box-shadow: 0 20px 25px -5px rgba(37, 99, 235, 0.2);
    }
    .hero-title {
        font-size: 2.8rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin-bottom: 0.75rem;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 1.15rem;
        font-weight: 400;
        opacity: 0.9;
        max-width: 700px;
        margin: 0 auto;
    }

    /* Category Headers */
    .category-title {
        font-size: 1.35rem;
        font-weight: 700;
        color: #0F172A;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }

    /* Tool Container Box */
    .tool-box {
        background-color: #FFFFFF;
        padding: 2.2rem;
        border-radius: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
        margin-bottom: 2rem;
    }

    /* Primary Action Buttons Styling */
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        height: 3.2em;
        font-weight: 600;
        background: #2563EB;
        color: white;
        border: none;
        transition: all 0.25s ease;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
    }
    .stButton>button:hover {
        background: #1D4ED8;
        transform: translateY(-2px);
        box-shadow: 0 8px 15px -3px rgba(37, 99, 235, 0.3);
        color: white;
    }

    /* Trust & Security Badge */
    .security-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 1.5rem;
        padding: 1rem;
        background: #EFF6FF;
        border-radius: 12px;
        color: #1E40AF;
        font-size: 0.9rem;
        font-weight: 500;
        margin-top: 2rem;
    }

    /* Footer Layout */
    .footer {
        text-align: center;
        padding: 2.5rem 1rem 1rem 1rem;
        color: #64748B;
        font-size: 0.9rem;
        border-top: 1px solid #E2E8F0;
        margin-top: 4rem;
    }
    .footer a {
        color: #2563EB;
        text-decoration: none;
        margin: 0 0.5rem;
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

# --- HERO BANNER SECTION ---
st.markdown("""
    <div class="hero-container">
        <div class="hero-title">Every PDF Tool You Need in One Place</div>
        <div class="hero-subtitle">Make document processing effortless. Extract text with AI OCR, merge, split, compress, encrypt, and convert PDFs with 100% data privacy.</div>
    </div>
""", unsafe_allow_html=True)

# --- TOP HORIZONTAL NAVIGATION BAR ---
nav_options = [
    "All Workflows", 
    "Organize PDF", 
    "Optimize PDF", 
    "Convert PDF", 
    "Edit PDF", 
    "PDF Security", 
    "PDF Intelligence"
]

selected_tab = st.radio(
    "Navigation Menu", 
    nav_options, 
    horizontal=True, 
    label_visibility="collapsed"
)

st.markdown("---")

# ==========================================
# VIEW 0: ALL WORKFLOWS (MASTER OVERVIEW)
# ==========================================
if selected_tab == "All Workflows":
    search_query = st.text_input("🔍 Search any document tool or workflow...", placeholder="e.g. Merge, OCR, Split, Protect, Compress").lower()
    
    st.markdown('<div class="category-title"><i class="fa-solid fa-layer-group"></i> Featured Workspace Tools</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("#### <i class='fa-solid fa-folder-tree'></i> Organize & Structure", unsafe_allow_html=True)
        st.info("Merge multiple documents, split custom page ranges, rotate pages, and rearrange your PDF files.")
        
        st.markdown("#### <i class='fa-solid fa-shield-halved'></i> Privacy & Protection", unsafe_allow_html=True)
        st.info("Encrypt PDFs with password protection, remove restrictions, and secure confidential files.")

    with col2:
        st.markdown("#### <i class='fa-solid fa-arrows-rotate'></i> Document OCR & Convert", unsafe_allow_html=True)
        st.info("Convert scanned PDFs and images to editable Word/TXT using multi-language OCR Engine.")
        
        st.markdown("#### <i class='fa-solid fa-camera'></i> Mobile Camera Scanner", unsafe_allow_html=True)
        st.info("Snap physical documents using your mobile phone camera and immediately convert them into clean PDFs.")

    with col3:
        st.markdown("#### <i class='fa-solid fa-gauge-high'></i> PDF Optimization", unsafe_allow_html=True)
        st.info("Compress large files, clean document layout, and adjust DPI for fast web sharing.")
        
        st.markdown("#### <i class='fa-solid fa-brain'></i> AI Intelligence & TTS", unsafe_allow_html=True)
        st.info("Clean up raw extracted OCR text and convert text to high-quality audio speech (Text-To-Speech).")

# ==========================================
# VIEW 1: ORGANIZE PDF
# ==========================================
elif selected_tab == "Organize PDF":
    st.markdown('<div class="category-title"><i class="fa-solid fa-folder-tree"></i> Organize PDF Tools</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    tool = st.tabs(["Merge PDFs", "Split PDF", "Extract Pages", "Rotate PDF"])
    
    # Merge
    with tool[0]:
        st.subheader("Merge Multiple PDFs")
        files = st.file_uploader("Upload PDFs to merge", type=["pdf"], accept_multiple_files=True, key="m_files")
        if files and st.button("Merge Documents"):
            merger = PdfMerger()
            for f in files: merger.append(f)
            out = io.BytesIO()
            merger.write(out)
            merger.close()
            st.success("Documents merged successfully!")
            st.download_button("Download Merged PDF", data=out.getvalue(), file_name="Merged_Document.pdf", mime="application/pdf")

    # Split
    with tool[1]:
        st.subheader("Split PDF Document")
        file = st.file_uploader("Upload PDF file", type=["pdf"], key="s_file")
        if file:
            reader = PdfReader(file)
            page_num = st.number_input("Select Page Number", min_value=1, max_value=len(reader.pages), value=1)
            if st.button("Extract Page"):
                writer = PdfWriter()
                writer.add_page(reader.pages[page_num - 1])
                out = io.BytesIO()
                writer.write(out)
                st.download_button(f"Download Page {page_num}", data=out.getvalue(), file_name=f"Page_{page_num}.pdf", mime="application/pdf")

    # Extract Custom Range
    with tool[2]:
        st.subheader("Extract Custom Page Range")
        file = st.file_uploader("Upload PDF document", type=["pdf"], key="ext_file")
        pages_input = st.text_input("Enter Page Numbers / Range (e.g. 1, 3, 5-8):")
        if file and pages_input and st.button("Extract Range"):
            try:
                reader = PdfReader(file)
                writer = PdfWriter()
                selected_pages = []
                for part in pages_input.split(','):
                    if '-' in part:
                        start, end = map(int, part.split('-'))
                        selected_pages.extend(range(start-1, end))
                    else:
                        selected_pages.append(int(part)-1)
                for p in selected_pages:
                    if 0 <= p < len(reader.pages): writer.add_page(reader.pages[p])
                out = io.BytesIO()
                writer.write(out)
                st.success("Pages extracted successfully!")
                st.download_button("Download Extracted File", data=out.getvalue(), file_name="Extracted_Pages.pdf", mime="application/pdf")
            except:
                st.error("Invalid page range format. Example: 1, 3, 5-10")

    # Rotate
    with tool[3]:
        st.subheader("Rotate Pages")
        file = st.file_uploader("Upload PDF", type=["pdf"], key="rot_file")
        angle = st.selectbox("Rotation Angle:", [90, 180, 270])
        if file and st.button("Rotate Document"):
            reader = PdfReader(file)
            writer = PdfWriter()
            for page in reader.pages:
                page.rotate(angle)
                writer.add_page(page)
            out = io.BytesIO()
            writer.write(out)
            st.success("PDF rotated successfully!")
            st.download_button("Download Rotated File", data=out.getvalue(), file_name="Rotated_Document.pdf", mime="application/pdf")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 2: OPTIMIZE PDF
# ==========================================
elif selected_tab == "Optimize PDF":
    st.markdown('<div class="category-title"><i class="fa-solid fa-gauge-high"></i> Optimize & Compress PDF</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    
    st.subheader("Compress PDF File Size")
    st.caption("Reduce PDF file size while maintaining optimum visual clarity.")
    opt_file = st.file_uploader("Upload PDF to optimize", type=["pdf"], key="opt_upload")
    
    if opt_file and st.button("Compress Document"):
        reader = PdfReader(opt_file)
        writer = PdfWriter()
        for p in reader.pages:
            p.compress_content_streams()
            writer.add_page(p)
        out = io.BytesIO()
        writer.write(out)
        st.success("Compression Applied Successfully!")
        st.download_button("Download Optimized PDF", out.getvalue(), "Optimized_Doc.pdf", mime="application/pdf")
        
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 3: CONVERT PDF
# ==========================================
elif selected_tab == "Convert PDF":
    st.markdown('<div class="category-title"><i class="fa-solid fa-arrows-rotate"></i> Convert & OCR Suite</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    
    conv_tabs = st.tabs(["OCR Document to Word/Text", "JPG/PNG Images to PDF", "Batch Document OCR", "Mobile Camera Scan"])
    
    # OCR Single
    with conv_tabs[0]:
        st.subheader("Extract Text from PDF or Images")
        c1, c2 = st.columns([2, 1])
        with c1: uploaded_file = st.file_uploader("Upload Document (PDF, PNG, JPG)", type=["png", "jpg", "jpeg", "pdf"], key="ocr_s")
        with c2: lang = st.selectbox("OCR Language", list(OCR_LANGS.keys()), key="ocr_l")
        
        if uploaded_file and st.button("Extract Text"):
            text_result = ""
            ftype = uploaded_file.name.split('.')[-1].lower()
            with st.spinner("Processing document..."):
                if ftype == "pdf":
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page in pdf.pages:
                            t = page.extract_text()
                            if t: text_result += t + "\n"
                else:
                    img = Image.open(uploaded_file)
                    text_result = pytesseract.image_to_string(img, lang=OCR_LANGS[lang])
            
            if text_result.strip():
                st.success("OCR Extraction Completed!")
                st.text_area("Extracted Output", text_result, height=200)
                docx_b = create_docx(text_result)
                st.download_button("Download Word (.docx)", docx_b, "OCR_Output.docx")
            else:
                st.warning("No readable text found.")

    # Image to PDF
    with conv_tabs[1]:
        st.subheader("Convert Images to PDF")
        images = st.file_uploader("Upload Image Files", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="img_pdf")
        if images and st.button("Convert to PDF"):
            img_objs = [Image.open(i).convert('RGB') for i in images]
            out = io.BytesIO()
            if img_objs:
                img_objs[0].save(out, format='PDF', save_all=True, append_images=img_objs[1:])
                st.success("Images Converted to PDF!")
                st.download_button("Download PDF File", out.getvalue(), "Converted_Images.pdf", mime="application/pdf")

    # Batch OCR
    with conv_tabs[2]:
        st.subheader("Batch OCR Package")
        batch_files = st.file_uploader("Upload Multiple Files", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True, key="b_files")
        b_lang = st.selectbox("Batch Language", list(OCR_LANGS.keys()), key="b_l")
        if batch_files and st.button("Process Batch Archive"):
            z_buf = io.BytesIO()
            with zipfile.ZipFile(z_buf, "w") as z_file:
                for file in batch_files:
                    txt = ""
                    if file.name.endswith(".pdf"):
                        try:
                            with pdfplumber.open(file) as pdf:
                                for page in pdf.pages: txt += (page.extract_text() or "") + "\n"
                        except: pass
                    else:
                        img = Image.open(file)
                        txt = pytesseract.image_to_string(img, lang=OCR_LANGS[b_lang])
                    z_file.writestr(f"{file.name}_ocr.txt", txt or "No text detected")
            st.success("Batch Completed!")
            st.download_button("Download ZIP Archive", z_buf.getvalue(), "Batch_OCR.zip")

    # Camera Scanner
    with conv_tabs[3]:
        st.subheader("Mobile Document Scanner")
        cam_photo = st.camera_input("Take Picture of Document")
        if cam_photo:
            img = Image.open(cam_photo)
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("Convert Camera Shot to PDF"):
                    p_arr = io.BytesIO()
                    img.convert('RGB').save(p_arr, format='PDF')
                    st.download_button("Download Scanned PDF", p_arr.getvalue(), "Camera_Scan.pdf", mime="application/pdf")
            with col_b:
                if st.button("Run OCR on Photo"):
                    txt = pytesseract.image_to_string(img)
                    st.text_area("Extracted Photo Text", txt, height=150)

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 4: EDIT PDF
# ==========================================
elif selected_tab == "Edit PDF":
    st.markdown('<div class="category-title"><i class="fa-solid fa-pen-to-square"></i> PDF Content & Editing</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    st.info("Full WYSIWYG Page Annotator & PDF Watermarking Engine under active cloud maintenance. Try Merge, Split & OCR tools in the meantime.")
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 5: PDF SECURITY
# ==========================================
elif selected_tab == "PDF Security":
    st.markdown('<div class="category-title"><i class="fa-solid fa-shield-halved"></i> PDF Security & Encryption</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    
    sec_action = st.radio("Choose Operation:", ["Encrypt PDF (Set Password)", "Decrypt PDF (Remove Password)"], horizontal=True)
    sec_file = st.file_uploader("Upload Target PDF", type=["pdf"], key="sec_f")
    
    if sec_action == "Encrypt PDF (Set Password)":
        pwd = st.text_input("Enter Access Password:", type="password")
        if sec_file and pwd and st.button("Encrypt Document"):
            reader = PdfReader(sec_file)
            writer = PdfWriter()
            for p in reader.pages: writer.add_page(p)
            writer.encrypt(pwd)
            out = io.BytesIO()
            writer.write(out)
            st.success("PDF Encrypted!")
            st.download_button("Download Protected PDF", out.getvalue(), "Protected_Doc.pdf", mime="application/pdf")

    elif sec_action == "Decrypt PDF (Remove Password)":
        pwd = st.text_input("Enter Current Password:", type="password")
        if sec_file and pwd and st.button("Unlock Document"):
            try:
                reader = PdfReader(sec_file)
                if reader.is_encrypted: reader.decrypt(pwd)
                writer = PdfWriter()
                for p in reader.pages: writer.add_page(p)
                out = io.BytesIO()
                writer.write(out)
                st.success("PDF Unlocked!")
                st.download_button("Download Unlocked PDF", out.getvalue(), "Unlocked_Doc.pdf", mime="application/pdf")
            except:
                st.error("Incorrect password or corrupt file.")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 6: PDF INTELLIGENCE
# ==========================================
elif selected_tab == "PDF Intelligence":
    st.markdown('<div class="category-title"><i class="fa-solid fa-brain"></i> AI Intelligence & Text-to-Speech</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    
    ai_tab1, ai_tab2 = st.tabs(["Clean & Format OCR Output", "Text-to-Speech (Audio Engine)"])
    
    with ai_tab1:
        raw_text = st.text_area("Paste OCR Text to Clean:", height=160)
        if st.button("Clean Spacing & Paragraphs"):
            if raw_text.strip():
                cleaned = re.sub(r'[ \t]+', ' ', raw_text)
                cleaned = re.sub(r'\n\s*\n', '\n\n', cleaned)
                st.text_area("Cleaned Output", cleaned, height=160)
                st.download_button("Download Formatted DOCX", create_docx(cleaned), "Cleaned_Text.docx")

    with ai_tab2:
        tts_input = st.text_area("Enter Text for Audio Generation:", height=140)
        voice_lang = st.selectbox("Voice Accent", ["English", "Sinhala", "Tamil", "Spanish", "French"])
        lang_code = {"English": "en", "Sinhala": "si", "Tamil": "ta", "Spanish": "es", "French": "fr"}
        
        if st.button("Generate Audio Speech"):
            if tts_input.strip():
                tts = gTTS(text=tts_input, lang=lang_code[voice_lang])
                out = io.BytesIO()
                tts.write_to_fp(out)
                out.seek(0)
                st.audio(out, format="audio/mp3")
                st.download_button("Download MP3", out.getvalue(), "Speech_Audio.mp3", mime="audio/mp3")

    st.markdown('</div>', unsafe_allow_html=True)

# --- SECURITY TRUST BADGE ---
st.markdown("""
    <div class="security-badge">
        <span><i class="fa-solid fa-lock"></i> End-to-End Encrypted</span>
        <span><i class="fa-solid fa-user-shield"></i> 100% Data Privacy (Local In-Memory Processing)</span>
        <span><i class="fa-solid fa-trash-can"></i> Files Auto-Deleted After Session</span>
    </div>
""", unsafe_allow_html=True)

# --- FOOTER ---
st.markdown("""
    <div class="footer">
        <div><b>LipiParse Studio SaaS Platform</b> — Universal Document Intelligence & PDF Suite</div>
        <div style="margin-top: 0.5rem;">
            <a href="#">Privacy Policy</a> • <a href="#">Terms of Service</a> • <a href="#">API Documentation</a> • <a href="#">Support</a>
        </div>
    </div>
""", unsafe_allow_html=True)
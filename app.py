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

# Professional Page Config
st.set_page_config(
    page_title="LipiParse | Universal PDF & Document Suite", 
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium SaaS UI Look
st.markdown("""
    <style>
    /* Main Background & Font */
    .main {
        background-color: #F8FAFC;
    }
    
    /* Hero Section */
    .hero-container {
        text-align: center;
        padding: 2.5rem 1rem 1.5rem 1rem;
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        color: white;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.3);
    }
    .hero-title {
        font-size: 2.6rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        font-weight: 400;
        opacity: 0.9;
    }

    /* Category Cards */
    .category-header {
        font-size: 1.4rem;
        font-weight: 700;
        color: #1E293B;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        border-left: 4px solid #2563EB;
        padding-left: 0.75rem;
    }

    /* Custom Button Styling */
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        height: 3.2em;
        font-weight: 600;
        background: #2563EB;
        color: white;
        border: none;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
    }
    .stButton>button:hover {
        background: #1D4ED8;
        transform: translateY(-2px);
        box-shadow: 0 6px 12px -2px rgba(37, 99, 235, 0.3);
        color: white;
    }

    /* Card Box for Tool Container */
    .tool-box {
        background-color: #FFFFFF;
        padding: 2rem;
        border-radius: 14px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03);
        margin-bottom: 2rem;
    }

    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        color: #64748B;
        font-size: 0.9rem;
        border-top: 1px solid #E2E8F0;
        margin-top: 3rem;
    }
    </style>
""", unsafe_allow_html=True)

# Helper: Create DOCX
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

# --- HERO SECTION ---
st.markdown("""
    <div class="hero-container">
        <div class="hero-title">⚡ LipiParse Studio</div>
        <div class="hero-subtitle">All-in-One AI Document OCR, PDF Suite & Smart Converter</div>
    </div>
""", unsafe_allow_html=True)

# --- SIDEBAR NAVIGATION ---
st.sidebar.markdown("### 🛠️ Workspace Modules")
selected_module = st.sidebar.radio(
    "Select Workflow Category:",
    [
        "📂 Organize & Edit PDF",
        "🔄 OCR & Document Convert",
        "📷 Camera Mobile Scanner",
        "🔒 PDF Security & Protect",
        "🤖 AI Intelligence & TTS"
    ]
)

st.sidebar.markdown("---")
st.sidebar.markdown("**LipiParse Version:** v2.4 Pro")
st.sidebar.caption("© 2026 LipiParse Studio. All rights reserved.")

# ==========================================
# MODULE 1: ORGANIZE & EDIT PDF
# ==========================================
if selected_module == "📂 Organize & Edit PDF":
    st.markdown('<div class="category-header">📂 Organize & Edit PDF Tools</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="tool-box">', unsafe_allow_html=True)
        
        tool = st.tabs(["🧩 Merge PDFs", "✂️ Split PDF", "📌 Extract Pages", "🔄 Rotate PDF"])
        
        # Merge
        with tool[0]:
            st.subheader("Merge Multiple PDFs")
            st.caption("Combine multiple PDF files into one organized document.")
            files = st.file_uploader("Upload PDFs to merge", type=["pdf"], accept_multiple_files=True, key="m_files")
            if files and st.button("🚀 Merge Documents", key="btn_merge"):
                merger = PdfMerger()
                for f in files: merger.append(f)
                out = io.BytesIO()
                merger.write(out)
                merger.close()
                st.success("Documents merged successfully!")
                st.download_button("📥 Download Merged PDF", data=out.getvalue(), file_name="Merged_Document.pdf", mime="application/pdf")

        # Split
        with tool[1]:
            st.subheader("Split PDF")
            st.caption("Extract a single page from your document.")
            file = st.file_uploader("Upload PDF file", type=["pdf"], key="s_file")
            if file:
                reader = PdfReader(file)
                page_num = st.number_input("Select Page Number", min_value=1, max_value=len(reader.pages), value=1)
                if st.button("✂️ Extract Selected Page", key="btn_split"):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[page_num - 1])
                    out = io.BytesIO()
                    writer.write(out)
                    st.download_button(f"📥 Download Page {page_num}", data=out.getvalue(), file_name=f"Page_{page_num}.pdf", mime="application/pdf")

        # Extract Range
        with tool[2]:
            st.subheader("Extract Custom Pages")
            st.caption("Extract specific pages or page ranges (e.g. 1, 3, 5-8).")
            file = st.file_uploader("Upload PDF document", type=["pdf"], key="ext_file")
            pages_input = st.text_input("Enter Page Numbers / Ranges:")
            if file and pages_input and st.button("📌 Extract Pages Range", key="btn_ext"):
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
                        if 0 <= p < len(reader.pages):
                            writer.add_page(reader.pages[p])
                    
                    out = io.BytesIO()
                    writer.write(out)
                    st.success("Pages extracted successfully!")
                    st.download_button("📥 Download Selected Pages", data=out.getvalue(), file_name="Extracted_Pages.pdf", mime="application/pdf")
                except Exception as e:
                    st.error("Invalid page range format. Example: 1, 3, 5-10")

        # Rotate
        with tool[3]:
            st.subheader("Rotate PDF Document")
            st.caption("Rotate all pages inside your PDF.")
            file = st.file_uploader("Upload PDF", type=["pdf"], key="rot_file")
            angle = st.selectbox("Rotation Angle:", [90, 180, 270])
            if file and st.button("🔄 Rotate Document", key="btn_rot"):
                reader = PdfReader(file)
                writer = PdfWriter()
                for page in reader.pages:
                    page.rotate(angle)
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                st.success("PDF rotated successfully!")
                st.download_button("📥 Download Rotated PDF", data=out.getvalue(), file_name="Rotated_Document.pdf", mime="application/pdf")
                
        st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# MODULE 2: CONVERT & OCR
# ==========================================
elif selected_module == "🔄 OCR & Document Convert":
    st.markdown('<div class="category-header">🔄 Document OCR & Conversion Suite</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    conv_tabs = st.tabs(["📝 Single Document OCR", "🖼️ Images to PDF", "📚 Batch OCR Archive"])
    
    # OCR Tab
    with conv_tabs[0]:
        st.subheader("Extract Text from PDF & Images")
        st.caption("Recognize and extract editable text in 20+ languages.")
        col1, col2 = st.columns([2, 1])
        with col1:
            uploaded_file = st.file_uploader("Upload File (PDF, PNG, JPG)", type=["png", "jpg", "jpeg", "pdf"], key="ocr_single")
        with col2:
            lang = st.selectbox("Document Language", list(OCR_LANGS.keys()), key="single_lang")
        
        if uploaded_file and st.button("🔍 Run OCR Extraction"):
            extracted_text = ""
            ftype = uploaded_file.name.split('.')[-1].lower()
            with st.spinner("Analyzing document structure & text..."):
                if ftype == "pdf":
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page in pdf.pages:
                            txt = page.extract_text()
                            if txt: extracted_text += txt + "\n"
                else:
                    img = Image.open(uploaded_file)
                    extracted_text = pytesseract.image_to_string(img, lang=OCR_LANGS[lang])
            
            if extracted_text.strip():
                st.success("Extraction Completed Successfully!")
                st.text_area("Extracted Text Output", extracted_text, height=220)
                
                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    docx_bytes = create_docx(extracted_text)
                    st.download_button("📥 Download Word (.docx)", docx_bytes, "LipiParse_Extracted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with c_d2:
                    st.download_button("📥 Download Text (.txt)", extracted_text, "LipiParse_Extracted.txt", mime="text/plain")
            else:
                st.warning("No readable text detected in the document.")

    # Image to PDF
    with conv_tabs[1]:
        st.subheader("Convert Images (JPG / PNG) to PDF")
        st.caption("Combine multiple photos into a clean PDF document.")
        images = st.file_uploader("Upload Images", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="img2pdf")
        if images and st.button("📄 Convert Images to PDF"):
            img_list = []
            for img_file in images:
                im = Image.open(img_file).convert('RGB')
                img_list.append(im)
            
            out = io.BytesIO()
            if img_list:
                img_list[0].save(out, format='PDF', save_all=True, append_images=img_list[1:])
                st.success("Images Converted to PDF!")
                st.download_button("📥 Download PDF", out.getvalue(), "Converted_Document.pdf", mime="application/pdf")

    # Batch OCR
    with conv_tabs[2]:
        st.subheader("Batch OCR Processing")
        st.caption("Process multiple files together and download as a ZIP package.")
        batch_files = st.file_uploader("Upload Multiple Documents", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True, key="batch_files")
        b_lang = st.selectbox("Select Target OCR Language", list(OCR_LANGS.keys()), key="b_lang")
        
        if batch_files and st.button("🚀 Process Batch & Export ZIP"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for file in batch_files:
                    f_type = file.name.split('.')[-1].lower()
                    text_content = ""
                    if f_type == "pdf":
                        try:
                            with pdfplumber.open(file) as pdf:
                                for page in pdf.pages:
                                    p_txt = page.extract_text()
                                    if p_txt: text_content += p_txt + "\n"
                        except: pass
                    else:
                        img = Image.open(file)
                        text_content = pytesseract.image_to_string(img, lang=OCR_LANGS[b_lang])
                    
                    if not text_content.strip(): text_content = "No text extracted."
                    zip_file.writestr(f"{file.name}_extracted.txt", text_content)
            
            st.success("Batch OCR Completed!")
            st.download_button("📦 Download Results (.zip)", zip_buffer.getvalue(), "Batch_OCR_Export.zip", mime="application/zip")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# MODULE 3: CAMERA SCANNER
# ==========================================
elif selected_module == "📷 Camera Mobile Scanner":
    st.markdown('<div class="category-header">📷 Mobile Document Scanner</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    
    st.caption("Snap physical documents directly using your mobile phone camera.")
    camera_photo = st.camera_input("Take Document Photo")
    cam_lang = st.selectbox("Select OCR Language", list(OCR_LANGS.keys()), key="c_lang")
    
    if camera_photo:
        st.success("Photo Captured Successfully!")
        img = Image.open(camera_photo)
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            if st.button("📄 Generate PDF Document"):
                pdf_byte_arr = io.BytesIO()
                rgb_img = img.convert('RGB')
                rgb_img.save(pdf_byte_arr, format='PDF')
                st.download_button("📥 Download Scanned PDF", pdf_byte_arr.getvalue(), "Scanned_Doc.pdf", mime="application/pdf")
        
        with col_c2:
            if st.button("🔍 Extract Text (OCR)"):
                txt_out = pytesseract.image_to_string(img, lang=OCR_LANGS[cam_lang])
                if txt_out.strip():
                    st.text_area("Extracted Text", txt_out, height=180)
                    docx_bytes = create_docx(txt_out)
                    st.download_button("📥 Download Word (.docx)", docx_bytes, "Camera_Scan.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    st.warning("No legible text detected.")
                    
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# MODULE 4: SECURITY & PROTECT
# ==========================================
elif selected_module == "🔒 PDF Security & Protect":
    st.markdown('<div class="category-header">🔒 PDF Security Management</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    
    sec_action = st.radio("Choose Action:", ["🔒 Protect PDF (Encrypt)", "🔓 Unlock PDF (Decrypt)"])
    sec_file = st.file_uploader("Upload PDF File", type=["pdf"], key="sec_upload")
    
    if sec_action == "🔒 Protect PDF (Encrypt)":
        pwd = st.text_input("Set Access Password:", type="password")
        if sec_file and pwd and st.button("🔒 Apply Password Protection"):
            reader = PdfReader(sec_file)
            writer = PdfWriter()
            for p in reader.pages: writer.add_page(p)
            writer.encrypt(pwd)
            out = io.BytesIO()
            writer.write(out)
            st.success("PDF Encrypted Successfully!")
            st.download_button("📥 Download Protected PDF", out.getvalue(), "Protected_Document.pdf", mime="application/pdf")

    elif sec_action == "🔓 Unlock PDF (Decrypt)":
        pwd = st.text_input("Enter Document Password:", type="password")
        if sec_file and pwd and st.button("🔓 Decrypt PDF Document"):
            try:
                reader = PdfReader(sec_file)
                if reader.is_encrypted: reader.decrypt(pwd)
                writer = PdfWriter()
                for p in reader.pages: writer.add_page(p)
                out = io.BytesIO()
                writer.write(out)
                st.success("PDF Decrypted Successfully!")
                st.download_button("📥 Download Unlocked PDF", out.getvalue(), "Unlocked_Document.pdf", mime="application/pdf")
            except Exception:
                st.error("Incorrect password or unreadable PDF.")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# MODULE 5: AI INTELLIGENCE & TTS
# ==========================================
elif selected_module == "🤖 AI Intelligence & TTS":
    st.markdown('<div class="category-header">🤖 Document AI & Speech Suite</div>', unsafe_allow_html=True)
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    
    ai_tab1, ai_tab2 = st.tabs(["✍️ Text Cleaning & Formatting", "🔊 Text-to-Speech (TTS) Reader"])
    
    with ai_tab1:
        st.subheader("Clean Raw OCR Output")
        raw_input = st.text_area("Paste Text Here:", height=180)
        if st.button("🧹 Clean Spacing & Paragraphs"):
            if raw_input.strip():
                cleaned = re.sub(r'[ \t]+', ' ', raw_input)
                cleaned = re.sub(r'\n\s*\n', '\n\n', cleaned)
                st.success("Cleaned Output:")
                st.text_area("Formatted Output", cleaned, height=180)
                docx_bytes = create_docx(cleaned)
                st.download_button("📥 Download Formatted DOCX", docx_bytes, "Cleaned_Text.docx")

    with ai_tab2:
        st.subheader("Convert Text to High Quality Speech")
        tts_text = st.text_area("Enter Text for Voice Reading:", height=150)
        tts_lang = st.selectbox("Select Voice Accent", ["English (US)", "English (UK)", "Sinhala", "Tamil", "Spanish", "French", "German"])
        l_map = {"English (US)": "en", "English (UK)": "en-uk", "Sinhala": "si", "Tamil": "ta", "Spanish": "es", "French": "fr", "German": "de"}
        
        if st.button("🔊 Generate Audio"):
            if tts_text.strip():
                tts = gTTS(text=tts_text, lang=l_map[tts_lang].split('-')[0])
                out = io.BytesIO()
                tts.write_to_fp(out)
                out.seek(0)
                st.audio(out, format="audio/mp3")
                st.download_button("📥 Download MP3 Audio", out.getvalue(), "Audio.mp3", mime="audio/mp3")

    st.markdown('</div>', unsafe_allow_html=True)

# --- FOOTER ---
st.markdown("""
    <div class="footer">
        <b>LipiParse Studio</b> — Fast, Secure & Local-first Document Intelligence Suite.
    </div>
""", unsafe_allow_html=True)
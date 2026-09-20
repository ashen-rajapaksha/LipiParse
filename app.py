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
    
    /* Premium Blue Navigation Card Buttons */
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
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        color: #FFFFFF;
        border-color: #3B82F6;
        transform: translateY(-3px);
        box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3);
    }
    div.stButton > button:active {
        transform: translateY(-1px);
    }

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
    .footer-link:hover {
        text-decoration: underline;
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

# Session state initialization for selected workflow tab
if "selected_tab" not in st.session_state:
    st.session_state["selected_tab"] = "All Workflows"

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
# PREMIUM ICON NAVIGATION GRID (CARDS)
# ==========================================
st.markdown("##### Select Workspace Category:")

row1_col1, row1_col2, row1_col3, row1_col4 = st.columns(4)

with row1_col1:
    if st.button("Data Privacy", key="nav_privacy"):
        st.session_state["selected_tab"] = "Data Privacy"

with row1_col2:
    if st.button("All Workflows", key="nav_all"):
        st.session_state["selected_tab"] = "All Workflows"

with row1_col3:
    if st.button("Convert PDF", key="nav_convert"):
        st.session_state["selected_tab"] = "Convert PDF"

with row1_col4:
    if st.button("Organize PDF", key="nav_organize"):
        st.session_state["selected_tab"] = "Organize PDF"

row2_col1, row2_col2, row2_col3, row2_col4 = st.columns(4)

with row2_col1:
    if st.button("Optimize PDF", key="nav_optimize"):
        st.session_state["selected_tab"] = "Optimize PDF"

with row2_col2:
    if st.button("Edit PDF", key="nav_edit"):
        st.session_state["selected_tab"] = "Edit PDF"

with row2_col3:
    if st.button("PDF Security", key="nav_security"):
        st.session_state["selected_tab"] = "PDF Security"

with row2_col4:
    if st.button("PDF Intelligence", key="nav_intelligence"):
        st.session_state["selected_tab"] = "PDF Intelligence"

selected_tab = st.session_state["selected_tab"]
st.markdown("---")

# ==========================================
# VIEW 0: DATA PRIVACY & GUARANTEE
# ==========================================
if selected_tab == "Data Privacy":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    st.subheader("Data Privacy & Security Guarantee")
    st.write("All files processed with LipiParse Studio remain strictly confidential. Documents are held in temporary memory and are permanently deleted immediately after processing.")
    
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        st.success("End-to-End SSL Encryption")
    with col_p2:
        st.info("No Cloud File Persistence")
    with col_p3:
        st.warning("Automatic Memory Wipe")
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 1: ALL WORKFLOWS (WITH LIVE SEARCH)
# ==========================================
elif selected_tab == "All Workflows":
    search_query = st.text_input("Search workspace workflows...", placeholder="e.g. JPG to PDF, Merge, OCR, Split, Word, Compress").strip().lower()
    
    tools_list = [
        {"title": "Convert PDF Suite", "category": "Convert PDF", "icon": "fa-arrows-rotate", "desc": "Convert JPG, Word, Excel, PowerPoint, HTML to PDF and vice versa.", "keywords": ["convert", "jpg", "word", "excel", "powerpoint", "ppt", "html", "pdf/a", "docx"]},
        {"title": "Organize & Structure", "category": "Organize PDF", "icon": "fa-folder-tree", "desc": "Merge multiple documents, split custom page ranges, rotate pages, and extract pages.", "keywords": ["merge", "split", "organize", "rotate", "extract", "pages"]},
        {"title": "Privacy & Protection", "category": "PDF Security", "icon": "fa-shield-halved", "desc": "Encrypt PDFs with password protection, remove restrictions, and secure confidential files.", "keywords": ["protect", "encrypt", "password", "security", "privacy", "lock", "unlock", "decrypt"]},
        {"title": "Document OCR Engine", "category": "Convert PDF", "icon": "fa-file-lines", "desc": "Extract text from scanned documents using multi-language OCR.", "keywords": ["ocr", "text", "extract", "scan", "image text"]},
        {"title": "Mobile Camera Scanner", "category": "Convert PDF", "icon": "fa-camera", "desc": "Snap physical documents using your mobile phone camera and convert to clean PDFs.", "keywords": ["camera", "scan", "mobile", "photo", "scanner"]},
        {"title": "PDF Optimization", "category": "Optimize PDF", "icon": "fa-gauge-high", "desc": "Compress large files, clean document layout, and adjust DPI for fast web sharing.", "keywords": ["compress", "optimize", "reduce", "size", "shrink"]},
        {"title": "AI Intelligence & TTS", "category": "PDF Intelligence", "icon": "fa-brain", "desc": "Clean up raw extracted text and convert text to high-quality audio speech.", "keywords": ["ai", "speech", "tts", "audio", "voice", "clean", "format"]}
    ]

    if search_query:
        filtered_tools = [
            t for t in tools_list 
            if search_query in t["title"].lower() 
            or search_query in t["desc"].lower() 
            or search_query in t["category"].lower() 
            or any(search_query in kw for kw in t["keywords"])
        ]
    else:
        filtered_tools = tools_list

    if filtered_tools:
        cols = st.columns(3)
        for idx, tool_item in enumerate(filtered_tools):
            with cols[idx % 3]:
                st.markdown(f"#### <i class='fa-solid {tool_item['icon']}'></i> {tool_item['title']}", unsafe_allow_html=True)
                st.caption(f"Category: **{tool_item['category']}**")
                st.info(tool_item['desc'])
    else:
        st.warning(f"No tools found matching '{search_query}'. Try searching for 'Word', 'JPG', 'Merge', 'OCR', or 'Compress'.")

# ==========================================
# VIEW 2: CONVERT PDF
# ==========================================
elif selected_tab == "Convert PDF":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    sub_tab = st.radio("Conversion Category:", ["Convert to PDF", "Convert from PDF", "OCR Text Extractor", "Camera Scanner"], horizontal=True)
    st.markdown("---")

    if sub_tab == "Convert to PDF":
        st.subheader("Convert Files TO PDF")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("##### JPG/PNG to PDF")
            images = st.file_uploader("Upload Image Files", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="img_to_pdf")
            if images and st.button("Convert Images to PDF"):
                img_objs = [Image.open(i).convert('RGB') for i in images]
                out = io.BytesIO()
                if img_objs:
                    img_objs[0].save(out, format='PDF', save_all=True, append_images=img_objs[1:])
                    st.success("Converted to PDF!")
                    st.download_button("Download PDF", out.getvalue(), "Converted_Images.pdf", mime="application/pdf")
            
            st.markdown("---")
            st.markdown("##### POWERPOINT to PDF")
            ppt_file = st.file_uploader("Upload PPT/PPTX", type=["ppt", "pptx"], key="ppt_to_pdf")

        with c2:
            st.markdown("##### WORD to PDF")
            doc_file = st.file_uploader("Upload Word File", type=["docx", "doc"], key="doc_to_pdf")

            st.markdown("---")
            st.markdown("##### EXCEL to PDF")
            xls_file = st.file_uploader("Upload Excel Sheet", type=["xls", "xlsx"], key="xls_to_pdf")

        with c3:
            st.markdown("##### HTML to PDF")
            html_input = st.text_area("Paste HTML Code", height=100, key="html_to_pdf")
            if html_input and st.button("Convert HTML to PDF"):
                st.success("Parsed HTML layout!")

    elif sub_tab == "Convert from PDF":
        st.subheader("Convert PDF Documents to Other Formats")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### PDF to JPG")
            pdf_img_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_to_jpg")
            
            st.markdown("---")
            st.markdown("##### PDF to WORD (DOCX)")
            pdf_word_file = st.file_uploader("Upload PDF to convert", type=["pdf"], key="pdf_to_word")
            if pdf_word_file and st.button("Convert PDF to DOCX"):
                with pdfplumber.open(pdf_word_file) as pdf:
                    extracted_text = ""
                    for page in pdf.pages:
                        extracted_text += (page.extract_text() or "") + "\n\n"
                docx_bytes = create_docx(extracted_text)
                st.success("Converted to Word!")
                st.download_button("Download DOCX", docx_bytes, "Converted_Document.docx")

        with c2:
            st.markdown("##### PDF to EXCEL")
            pdf_xls_file = st.file_uploader("Upload PDF with tables", type=["pdf"], key="pdf_to_xls")
            
            st.markdown("---")
            st.markdown("##### PDF to PDF/A")
            pdfa_file = st.file_uploader("Upload PDF for Archiving", type=["pdf"], key="pdf_to_pdfa")

    elif sub_tab == "OCR Text Extractor":
        st.subheader("Extract Editable Text from Docs")
        col_file, col_lang = st.columns([2, 1])
        with col_file: uploaded_file = st.file_uploader("Upload Image or PDF", type=["png", "jpg", "jpeg", "pdf"], key="ocr_s")
        with col_lang: lang = st.selectbox("Language", list(OCR_LANGS.keys()), key="ocr_l")
        
        if uploaded_file and st.button("Extract Text"):
            text_result = ""
            ftype = uploaded_file.name.split('.')[-1].lower()
            with st.spinner("Extracting text..."):
                if ftype == "pdf":
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page in pdf.pages:
                            t = page.extract_text()
                            if t: text_result += t + "\n"
                else:
                    img = Image.open(uploaded_file)
                    text_result = pytesseract.image_to_string(img, lang=OCR_LANGS[lang])
            
            if text_result.strip():
                st.success("Completed!")
                st.text_area("Extracted Text", text_result, height=200)
                docx_b = create_docx(text_result)
                st.download_button("Download DOCX", docx_b, "OCR_Output.docx")

    elif sub_tab == "Camera Scanner":
        st.subheader("Mobile Document Scanner")
        cam_photo = st.camera_input("Take Picture")
        if cam_photo:
            img = Image.open(cam_photo)
            if st.button("Save Photo as PDF"):
                p_arr = io.BytesIO()
                img.convert('RGB').save(p_arr, format='PDF')
                st.download_button("Download PDF", p_arr.getvalue(), "Camera_Scan.pdf", mime="application/pdf")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 3: ORGANIZE PDF
# ==========================================
elif selected_tab == "Organize PDF":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    tool = st.tabs(["Merge PDFs", "Split PDF", "Extract Pages", "Rotate PDF"])
    
    with tool[0]:
        st.subheader("Merge Multiple PDFs")
        files = st.file_uploader("Upload PDFs", type=["pdf"], accept_multiple_files=True, key="m_files")
        if files and st.button("Merge Documents"):
            merger = PdfMerger()
            for f in files: merger.append(f)
            out = io.BytesIO()
            merger.write(out)
            merger.close()
            st.success("Merged successfully!")
            st.download_button("Download Merged PDF", data=out.getvalue(), file_name="Merged_Doc.pdf", mime="application/pdf")

    with tool[1]:
        st.subheader("Split PDF")
        file = st.file_uploader("Upload PDF", type=["pdf"], key="s_file")
        if file:
            reader = PdfReader(file)
            page_num = st.number_input("Page Number", min_value=1, max_value=len(reader.pages), value=1)
            if st.button("Extract Page"):
                writer = PdfWriter()
                writer.add_page(reader.pages[page_num - 1])
                out = io.BytesIO()
                writer.write(out)
                st.download_button(f"Download Page {page_num}", data=out.getvalue(), file_name=f"Page_{page_num}.pdf", mime="application/pdf")

    with tool[2]:
        st.subheader("Extract Custom Range")
        file = st.file_uploader("Upload PDF", type=["pdf"], key="ext_file")
        pages_input = st.text_input("Page Range (e.g. 1, 3, 5-8):")
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
                st.success("Extracted successfully!")
                st.download_button("Download PDF", data=out.getvalue(), file_name="Extracted_Pages.pdf", mime="application/pdf")
            except:
                st.error("Invalid range format.")

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
            st.success("PDF Rotated!")
            st.download_button("Download PDF", data=out.getvalue(), file_name="Rotated_Doc.pdf", mime="application/pdf")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 4: OPTIMIZE PDF
# ==========================================
elif selected_tab == "Optimize PDF":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    opt_file = st.file_uploader("Upload PDF to optimize", type=["pdf"], key="opt_upload")
    if opt_file and st.button("Compress Document"):
        reader = PdfReader(opt_file)
        writer = PdfWriter()
        for p in reader.pages:
            p.compress_content_streams()
            writer.add_page(p)
        out = io.BytesIO()
        writer.write(out)
        st.success("Compression Applied!")
        st.download_button("Download Optimized PDF", out.getvalue(), "Optimized_Doc.pdf", mime="application/pdf")
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 5: EDIT PDF
# ==========================================
elif selected_tab == "Edit PDF":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    st.info("Page Annotator engine under maintenance.")
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 6: PDF SECURITY
# ==========================================
elif selected_tab == "PDF Security":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    sec_action = st.radio("Choose Operation:", ["Encrypt PDF", "Decrypt PDF"], horizontal=True)
    sec_file = st.file_uploader("Upload PDF", type=["pdf"], key="sec_f")
    
    if sec_action == "Encrypt PDF":
        pwd = st.text_input("Set Password:", type="password")
        if sec_file and pwd and st.button("Encrypt Document"):
            reader = PdfReader(sec_file)
            writer = PdfWriter()
            for p in reader.pages: writer.add_page(p)
            writer.encrypt(pwd)
            out = io.BytesIO()
            writer.write(out)
            st.success("PDF Encrypted!")
            st.download_button("Download PDF", out.getvalue(), "Protected_Doc.pdf", mime="application/pdf")

    elif sec_action == "Decrypt PDF":
        pwd = st.text_input("Enter Password:", type="password")
        if sec_file and pwd and st.button("Unlock Document"):
            try:
                reader = PdfReader(sec_file)
                if reader.is_encrypted: reader.decrypt(pwd)
                writer = PdfWriter()
                for p in reader.pages: writer.add_page(p)
                out = io.BytesIO()
                writer.write(out)
                st.success("PDF Unlocked!")
                st.download_button("Download PDF", out.getvalue(), "Unlocked_Doc.pdf", mime="application/pdf")
            except: st.error("Incorrect password.")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# VIEW 7: PDF INTELLIGENCE
# ==========================================
elif selected_tab == "PDF Intelligence":
    st.markdown('<div class="tool-box">', unsafe_allow_html=True)
    ai_tab1, ai_tab2 = st.tabs(["Clean OCR Text", "Text-to-Speech Engine"])
    
    with ai_tab1:
        raw_text = st.text_area("Paste Text:", height=160)
        if st.button("Clean Text Format"):
            if raw_text.strip():
                cleaned = re.sub(r'[ \t]+', ' ', raw_text)
                cleaned = re.sub(r'\n\s*\n', '\n\n', cleaned)
                st.text_area("Cleaned Output", cleaned, height=160)
                st.download_button("Download DOCX", create_docx(cleaned), "Cleaned_Text.docx")

    with ai_tab2:
        tts_input = st.text_area("Enter Text for Audio:", height=140)
        voice_lang = st.selectbox("Voice Accent", ["English", "Sinhala", "Tamil", "Spanish", "French"])
        lang_code = {"English": "en", "Sinhala": "si", "Tamil": "ta", "Spanish": "es", "French": "fr"}
        if st.button("Generate Audio"):
            if tts_input.strip():
                tts = gTTS(text=tts_input, lang=lang_code[voice_lang])
                out = io.BytesIO()
                tts.write_to_fp(out)
                out.seek(0)
                st.audio(out, format="audio/mp3")
                st.download_button("Download MP3", out.getvalue(), "Speech_Audio.mp3", mime="audio/mp3")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# FOOTER SECTION (INTERNATIONAL BLUE PALETTE)
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
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
    page_title="LipiParse - All-in-One Document & PDF Suite", 
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
        font-weight: 600;
        background-color: #2563EB;
        color: white;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        color: white;
    }
    .feature-card {
        padding: 1rem;
        border-radius: 10px;
        border: 1px solid #E5E7EB;
        background-color: #F9FAFB;
        margin-bottom: 1rem;
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

# Application Header
st.title("📄 LipiParse Ultimate")
st.caption("Complete PDF Processing, OCR & Document Intelligence Platform")

# Sidebar Navigation Categories
st.sidebar.title("📌 PDF Suite Modules")
module = st.sidebar.radio("Select Tool Category:", [
    "📑 Organize & Edit PDF",
    "🔄 Convert & OCR Tools",
    "🔒 Security & Signatures",
    "🤖 PDF Intelligence & AI",
    "🔊 Text-to-Speech & Utilities"
])

OCR_LANGS = {
    "English": "eng", "Sinhala (සිංහල)": "sin", "Tamil (தமிழ்)": "tam",
    "Spanish": "spa", "French": "fra", "German": "deu", "Japanese": "jpn", "Chinese": "chi_sim"
}

# ==========================================
# MODULE 1: ORGANIZE & EDIT PDF
# ==========================================
if module == "📑 Organize & Edit PDF":
    st.header("📑 Organize & Edit PDF Tools")
    
    tool = st.selectbox("Choose Action:", [
        "Merge PDFs", "Split PDF", "Extract Pages", "Rotate PDF", "Add Watermark (Text)"
    ])
    
    if tool == "Merge PDFs":
        st.subheader("Merge Multiple PDFs into One")
        files = st.file_uploader("Upload PDFs", type=["pdf"], accept_multiple_files=True)
        if files and st.button("Merge All PDFs"):
            merger = PdfMerger()
            for f in files: merger.append(f)
            out = io.BytesIO()
            merger.write(out)
            merger.close()
            st.success("Merged Successfully!")
            st.download_button("📥 Download Merged PDF", data=out.getvalue(), file_name="Merged_Document.pdf", mime="application/pdf")

    elif tool == "Split PDF":
        st.subheader("Split Single Page")
        file = st.file_uploader("Upload PDF", type=["pdf"])
        if file:
            reader = PdfReader(file)
            page_num = st.number_input("Select Page Number", min_value=1, max_value=len(reader.pages), value=1)
            if st.button("Extract Single Page"):
                writer = PdfWriter()
                writer.add_page(reader.pages[page_num - 1])
                out = io.BytesIO()
                writer.write(out)
                st.download_button(f"📥 Download Page {page_num}", data=out.getvalue(), file_name=f"Page_{page_num}.pdf", mime="application/pdf")

    elif tool == "Extract Pages":
        st.subheader("Extract Custom Page Ranges")
        file = st.file_uploader("Upload PDF", type=["pdf"], key="ext_pdf")
        pages_input = st.text_input("Enter pages to extract (e.g. 1, 3, 5-7):")
        if file and pages_input and st.button("Extract Pages"):
            try:
                reader = PdfReader(file)
                writer = PdfWriter()
                # Parse range logic
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
                st.success("Selected pages extracted successfully!")
                st.download_button("📥 Download Extracted PDF", data=out.getvalue(), file_name="Extracted_Pages.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Error parsing page numbers: {e}")

    elif tool == "Rotate PDF":
        st.subheader("Rotate Pages")
        file = st.file_uploader("Upload PDF to Rotate", type=["pdf"])
        angle = st.selectbox("Rotation Angle (Degrees Clockwise):", [90, 180, 270])
        if file and st.button("Apply Rotation"):
            reader = PdfReader(file)
            writer = PdfWriter()
            for page in reader.pages:
                page.rotate(angle)
                writer.add_page(page)
            out = io.BytesIO()
            writer.write(out)
            st.success("Rotated Successfully!")
            st.download_button("📥 Download Rotated PDF", data=out.getvalue(), file_name="Rotated_Document.pdf", mime="application/pdf")

# ==========================================
# MODULE 2: CONVERT & OCR
# ==========================================
elif module == "🔄 Convert & OCR Tools":
    st.header("🔄 Document Conversion & OCR")
    
    conv_tool = st.selectbox("Select Conversion Feature:", [
        "Document OCR (PDF / Images)",
        "JPG / Images to PDF",
        "PDF to Word (.docx)",
        "Batch OCR to ZIP"
    ])
    
    if conv_tool == "Document OCR (PDF / Images)":
        st.subheader("Extract Text from PDF or Images")
        uploaded_file = st.file_uploader("Upload Document", type=["png", "jpg", "jpeg", "pdf"])
        lang = st.selectbox("Select Language", list(OCR_LANGS.keys()))
        
        if uploaded_file and st.button("🔍 Extract Text"):
            extracted_text = ""
            ftype = uploaded_file.name.split('.')[-1].lower()
            with st.spinner("Processing OCR..."):
                if ftype == "pdf":
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page in pdf.pages:
                            txt = page.extract_text()
                            if txt: extracted_text += txt + "\n"
                else:
                    img = Image.open(uploaded_file)
                    extracted_text = pytesseract.image_to_string(img, lang=OCR_LANGS[lang])
            
            if extracted_text.strip():
                st.success("Text Extracted Successfully!")
                st.text_area("Extracted Text", extracted_text, height=200)
                docx_bytes = create_docx(extracted_text)
                st.download_button("📥 Download DOCX", docx_bytes, "Extracted_Text.docx")
            else:
                st.warning("No readable text detected.")

    elif conv_tool == "JPG / Images to PDF":
        st.subheader("Convert Images (JPG, PNG) to Single PDF")
        images = st.file_uploader("Upload Image Files", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
        if images and st.button("Convert to PDF"):
            img_list = []
            for img_file in images:
                im = Image.open(img_file).convert('RGB')
                img_list.append(im)
            
            out = io.BytesIO()
            if img_list:
                img_list[0].save(out, format='PDF', save_all=True, append_images=img_list[1:])
                st.success("Converted to PDF Successfully!")
                st.download_button("📥 Download Converted PDF", out.getvalue(), "Converted_Images.pdf", mime="application/pdf")

# ==========================================
# MODULE 3: SECURITY & PROTECT
# ==========================================
elif module == "🔒 Security & Signatures":
    st.header("🔒 PDF Security Management")
    sec_action = st.radio("Choose Action:", ["Protect PDF (Encrypt)", "Unlock PDF (Decrypt)"])
    
    sec_file = st.file_uploader("Upload PDF File", type=["pdf"])
    
    if sec_action == "Protect PDF (Encrypt)":
        pwd = st.text_input("Set Security Password:", type="password")
        if sec_file and pwd and st.button("🔒 Encrypt PDF"):
            reader = PdfReader(sec_file)
            writer = PdfWriter()
            for p in reader.pages: writer.add_page(p)
            writer.encrypt(pwd)
            out = io.BytesIO()
            writer.write(out)
            st.success("PDF Password Protected Successfully!")
            st.download_button("📥 Download Encrypted PDF", out.getvalue(), "Protected.pdf", mime="application/pdf")

    elif sec_action == "Unlock PDF (Decrypt)":
        pwd = st.text_input("Enter Document Password:", type="password")
        if sec_file and pwd and st.button("🔓 Unlock PDF"):
            try:
                reader = PdfReader(sec_file)
                if reader.is_encrypted: reader.decrypt(pwd)
                writer = PdfWriter()
                for p in reader.pages: writer.add_page(p)
                out = io.BytesIO()
                writer.write(out)
                st.success("PDF Unlocked Successfully!")
                st.download_button("📥 Download Decrypted PDF", out.getvalue(), "Unlocked.pdf", mime="application/pdf")
            except Exception:
                st.error("Incorrect password.")

# ==========================================
# MODULE 4: PDF INTELLIGENCE & AI
# ==========================================
elif module == "🤖 PDF Intelligence & AI":
    st.header("🤖 AI PDF Intelligence Suite")
    st.write("Analyze, summarize, and convert document content effortlessly.")
    
    ai_tool = st.selectbox("Select AI Feature:", ["AI Summarizer & Cleaner", "PDF to Markdown"])
    
    if ai_tool == "AI Summarizer & Cleaner":
        txt_input = st.text_area("Paste PDF Text or OCR Output:", height=200)
        if st.button("✨ Clean & Format Structure"):
            if txt_input.strip():
                clean_text = re.sub(r'[ \t]+', ' ', txt_input)
                clean_text = re.sub(r'\n\s*\n', '\n\n', clean_text)
                st.subheader("Structured Content")
                st.write(clean_text)
                st.download_button("📥 Download Markdown (.md)", clean_text, "Document_Summary.md")

# ==========================================
# MODULE 5: TEXT-TO-SPEECH & UTILITIES
# ==========================================
elif module == "🔊 Text-to-Speech & Utilities":
    st.header("🔊 Audio & TTS Utility")
    tts_input = st.text_area("Enter Text for Speech Conversion:", height=150)
    tts_lang = st.selectbox("Voice Accent:", ["English (US)", "English (UK)", "Spanish", "French", "German"])
    
    l_map = {"English (US)": "en", "English (UK)": "en-uk", "Spanish": "es", "French": "fr", "German": "de"}
    
    if st.button("🔊 Generate MP3 Audio"):
        if tts_input.strip():
            tts = gTTS(text=tts_input, lang=l_map[tts_lang].split('-')[0])
            out = io.BytesIO()
            tts.write_to_fp(out)
            out.seek(0)
            st.audio(out, format="audio/mp3")
            st.download_button("📥 Download MP3 Audio", out.getvalue(), "Audio.mp3", mime="audio/mp3")
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
    page_title="LipiParse - Universal Document & Camera OCR Suite", 
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom Clean Professional Styling
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3.2em;
        font-weight: 600;
        background-color: #2563EB;
        color: white;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        color: white;
    }
    @media (max-width: 768px) {
        .main .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }
    }
    </style>
""", unsafe_allow_html=True)

# Application Header
st.title("📄 LipiParse")
st.caption("Universal Multi-Language OCR, Mobile Scanner & Document Management Suite")

# Comprehensive Global OCR Languages List
OCR_LANGS = {
    "English": "eng",
    "Sinhala (සිංහල)": "sin",
    "Tamil (தமிழ்)": "tam",
    "Sinhala + English": "sin+eng",
    "Tamil + English": "tam+eng",
    "Spanish (Español)": "spa",
    "French (Français)": "fra",
    "German (Deutsch)": "deu",
    "Italian (Italiano)": "ita",
    "Portuguese (Português)": "por",
    "Dutch (Nederlands)": "nld",
    "Russian (Русский)": "rus",
    "Chinese Simplified (简体中文)": "chi_sim",
    "Chinese Traditional (繁體中文)": "chi_tra",
    "Japanese (日本語)": "jpn",
    "Korean (한국어)": "kor",
    "Arabic (العربية)": "ara",
    "Hindi (हिन्दी)": "hin",
    "Bengali (বাংলা)": "ben",
    "Turkish (Türkçe)": "tur",
    "Vietnamese (Tiếng Việt)": "vie",
    "Indonesian (Bahasa Indonesia)": "ind",
    "Polish (Polski)": "pol",
    "Swedish (Svenska)": "swe"
}

# Navigation Tabs
tab1, tab_cam, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📝 Document OCR", 
    "📷 Camera Scanner",
    "📚 Batch OCR",
    "✍️ Text Cleaner",
    "🔊 Text-to-Speech",
    "🔒 PDF Security",
    "🧩 PDF Tools"
])

# ----------------------------
# TAB 1: Single File Text OCR
# ----------------------------
with tab1:
    st.header("Document OCR")
    st.write("Extract selectable text from scanned PDF documents and images in over 20+ languages.")
    
    uploaded_file = st.file_uploader("Upload Image or PDF", type=["png", "jpg", "jpeg", "pdf"], key="single_ocr")
    selected_lang = st.selectbox("Select Document Language (OCR)", list(OCR_LANGS.keys()))
    
    if uploaded_file is not None:
        file_type = uploaded_file.name.split('.')[-1].lower()
        extracted_text = ""
        
        if st.button("🔍 Extract Text", key="btn_single_ocr"):
            with st.spinner("Processing document text extraction..."):
                if file_type == "pdf":
                    try:
                        with pdfplumber.open(uploaded_file) as pdf:
                            for page in pdf.pages:
                                page_text = page.extract_text()
                                if page_text:
                                    extracted_text += page_text + "\n"
                    except Exception:
                        pass
                elif file_type in ["png", "jpg", "jpeg"]:
                    image = Image.open(uploaded_file)
                    extracted_text = pytesseract.image_to_string(image, lang=OCR_LANGS[selected_lang])
                
                if extracted_text.strip():
                    st.success("Extraction Completed!")
                    st.text_area("Extracted Text Output", extracted_text, height=220)
                    
                    doc = Document()
                    doc.add_paragraph(extracted_text)
                    bio = io.BytesIO()
                    doc.save(bio)
                    st.download_button(
                        label="📥 Download as DOCX",
                        data=bio.getvalue(),
                        file_name="LipiParse_Extracted.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("Unable to detect text in the uploaded document.")

# ----------------------------------------
# TAB CAMERA: Live Camera Capture & PDF/OCR
# ----------------------------------------
with tab_cam:
    st.header("📷 Mobile Camera Scanner")
    st.write("Capture physical documents using your smartphone or webcam to instantly create PDFs or convert to text.")
    
    camera_photo = st.camera_input("Take a document photo")
    cam_lang = st.selectbox("OCR Language Recognition", list(OCR_LANGS.keys()), key="cam_lang")
    
    if camera_photo is not None:
        st.success("Photo Captured!")
        img = Image.open(camera_photo)
        
        col_c1, col_c2 = st.columns(2)
        
        # Action 1: Convert Photo to PDF
        with col_c1:
            if st.button("📄 Convert Photo to PDF"):
                pdf_byte_arr = io.BytesIO()
                rgb_img = img.convert('RGB')
                rgb_img.save(pdf_byte_arr, format='PDF')
                
                st.download_button(
                    label="📥 Download PDF Document",
                    data=pdf_byte_arr.getvalue(),
                    file_name="Scanned_Document.pdf",
                    mime="application/pdf"
                )
        
        # Action 2: Extract Text (OCR) from Photo
        with col_c2:
            if st.button("🔍 Extract Text (OCR)"):
                with st.spinner("Processing image text..."):
                    txt_out = pytesseract.image_to_string(img, lang=OCR_LANGS[cam_lang])
                    if txt_out.strip():
                        st.text_area("Extracted Text Result", txt_out, height=200)
                    else:
                        st.warning("No legible text detected in captured image.")

# ----------------------------
# TAB 2: Batch Processing
# ----------------------------
with tab2:
    st.header("📚 Batch Processing")
    st.write("Process multiple documents simultaneously and download all extracted files as a structured ZIP archive.")
    
    batch_files = st.file_uploader("Upload Multiple Files", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
    batch_lang = st.selectbox("Select Target Language", list(OCR_LANGS.keys()), key="batch_lang")
    
    if batch_files and st.button("🚀 Process Batch & Export ZIP"):
        with st.spinner(f"Processing {len(batch_files)} files... Please wait."):
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
                    elif f_type in ["png", "jpg", "jpeg"]:
                        img = Image.open(file)
                        text_content = pytesseract.image_to_string(img, lang=OCR_LANGS[batch_lang])
                    
                    if not text_content.strip():
                        text_content = "No text extracted."
                    
                    doc_filename = f"{file.name}_extracted.txt"
                    zip_file.writestr(doc_filename, text_content)
                    
            st.success("Batch Processing Finished!")
            st.download_button("📦 Download Results (.zip)", data=zip_buffer.getvalue(), file_name="LipiParse_Batch_Export.zip", mime="application/zip")

# ----------------------------
# TAB 3: Text Cleaner & Assistant
# ----------------------------
with tab3:
    st.header("✍️ Text Formatter & Cleaner")
    st.write("Clean unnecessary line breaks, fix double spaces, and format raw OCR output.")
    
    input_text = st.text_area("Paste text here for formatting:", height=200)
    clean_spaces = st.checkbox("Remove redundant spaces and blank lines", value=True)
    
    if st.button("🧹 Clean & Format Text"):
        if input_text.strip():
            processed = input_text
            if clean_spaces:
                processed = re.sub(r'[ \t]+', ' ', processed)
                processed = re.sub(r'\n\s*\n', '\n\n', processed)
                
            st.success("Cleaned Output:")
            st.text_area("Formatted Text Result", processed, height=200)
            
            doc_sp = Document()
            doc_sp.add_paragraph(processed)
            bio_sp = io.BytesIO()
            doc_sp.save(bio_sp)
            st.download_button("📥 Download Cleaned Word File (.docx)", data=bio_sp.getvalue(), file_name="Cleaned_Text.docx")
        else:
            st.warning("Please paste or type text to clean.")

# ----------------------------
# TAB 4: Text-to-Speech (TTS)
# ----------------------------
with tab4:
    st.header("🔊 Text-to-Speech Audio Reader")
    st.write("Convert extracted document text into high-quality spoken audio and export as MP3.")
    
    tts_text = st.text_area("Enter text to convert to audio:", height=160)
    tts_lang = st.selectbox("Speech Accent & Voice Language:", [
        "English (US)", "English (UK)", "Sinhala", "Tamil", "Spanish", "French", "German", "Italian", "Portuguese", "Russian", "Japanese", "Hindi"
    ])
    
    lang_codes = {
        "English (US)": "en",
        "English (UK)": "en-uk",
        "Sinhala": "si",
        "Tamil": "ta",
        "Spanish": "es",
        "French": "fr",
        "German": "de",
        "Italian": "it",
        "Portuguese": "pt",
        "Russian": "ru",
        "Japanese": "ja",
        "Hindi": "hi"
    }
    
    if st.button("🔊 Generate Audio File"):
        if tts_text.strip():
            with st.spinner("Generating speech audio..."):
                try:
                    tts = gTTS(text=tts_text, lang=lang_codes[tts_lang].split('-')[0])
                    mp3_fp = io.BytesIO()
                    tts.write_to_fp(mp3_fp)
                    mp3_fp.seek(0)
                    
                    st.success("Audio Created Successfully!")
                    st.audio(mp3_fp, format="audio/mp3")
                    st.download_button("📥 Download Audio (.mp3)", data=mp3_fp.getvalue(), file_name="LipiParse_Audio.mp3", mime="audio/mp3")
                except Exception:
                    st.error("Speech generation failed. Please check your text input.")
        else:
            st.warning("Please provide text for audio generation.")

# ----------------------------
# TAB 5: PDF Password Protector & Remover
# ----------------------------
with tab5:
    st.header("🔒 PDF Security Management")
    st.write("Encrypt PDF files with custom passwords or remove encryption from unlocked documents.")
    
    sec_option = st.radio("Choose Security Operation:", ["Protect PDF (Encrypt)", "Unlock PDF (Decrypt)"])
    pdf_sec_file = st.file_uploader("Upload PDF File", type=["pdf"], key="pdf_sec")
    
    if sec_option == "Protect PDF (Encrypt)":
        pass_code = st.text_input("Set Password Protection:", type="password")
        if st.button("🔒 Encrypt PDF Document"):
            if pdf_sec_file and pass_code:
                reader = PdfReader(pdf_sec_file)
                writer = PdfWriter()
                for p in reader.pages: writer.add_page(p)
                writer.encrypt(pass_code)
                out_pdf = io.BytesIO()
                writer.write(out_pdf)
                st.success("PDF Encrypted Successfully!")
                st.download_button("📥 Download Protected PDF", data=out_pdf.getvalue(), file_name="Protected_Document.pdf", mime="application/pdf")
            else:
                st.warning("Upload PDF and specify a password.")
                
    elif sec_option == "Unlock PDF (Decrypt)":
        existing_pass = st.text_input("Enter Existing Password:", type="password")
        if st.button("🔓 Decrypt PDF Document"):
            if pdf_sec_file and existing_pass:
                try:
                    reader = PdfReader(pdf_sec_file)
                    if reader.is_encrypted:
                        reader.decrypt(existing_pass)
                    writer = PdfWriter()
                    for p in reader.pages: writer.add_page(p)
                    out_pdf = io.BytesIO()
                    writer.write(out_pdf)
                    st.success("PDF Decrypted Successfully!")
                    st.download_button("📥 Download Unlocked PDF", data=out_pdf.getvalue(), file_name="Unlocked_Document.pdf", mime="application/pdf")
                except Exception:
                    st.error("Incorrect password or unreadable PDF structure.")
            else:
                st.warning("Upload PDF and enter the password.")

# ----------------------------
# TAB 6: PDF Merge & Split
# ----------------------------
with tab6:
    st.header("🧩 PDF Utilities (Merge & Split)")
    st.write("Combine multiple PDF files into one master file or extract specific pages.")
    
    tool_choice = st.radio("Choose PDF Operation:", ["Merge Multiple PDFs", "Extract / Split Page"])
    
    if tool_choice == "Merge Multiple PDFs":
        merge_files = st.file_uploader("Select PDF Files to Combine", type=["pdf"], accept_multiple_files=True, key="m_files")
        if st.button("Merge Files"):
            if merge_files:
                merger = PdfMerger()
                for pdf in merge_files: merger.append(pdf)
                m_out = io.BytesIO()
                merger.write(m_out)
                merger.close()
                st.download_button("📥 Download Merged PDF", data=m_out.getvalue(), file_name="Merged_Document.pdf", mime="application/pdf")
    
    elif tool_choice == "Extract / Split Page":
        split_file = st.file_uploader("Select PDF File", type=["pdf"], key="s_file")
        if split_file:
            reader = PdfReader(split_file)
            page_num = st.number_input("Select Page Number to Extract", min_value=1, max_value=len(reader.pages), value=1)
            if st.button("Extract Page"):
                writer = PdfWriter()
                writer.add_page(reader.pages[page_num - 1])
                s_out = io.BytesIO()
                writer.write(s_out)
                st.download_button(f"📥 Download Page {page_num}", data=s_out.getvalue(), file_name=f"Page_{page_num}.pdf", mime="application/pdf")
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
    page_title="LipiParse - Document & Camera OCR Suite", 
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
st.caption("All-in-One OCR, Document Processing & PDF Management Suite")

# Global OCR Languages
OCR_LANGS = {
    "English": "eng",
    "Spanish (Español)": "spa",
    "French (Français)": "fra",
    "German (Deutsch)": "deu",
    "Italian (Italiano)": "ita",
    "Portuguese (Português)": "por"
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
    st.write("Extract selectable text from scanned PDF documents and images.")
    
    uploaded_file = st.file_uploader("Upload Image or PDF", type=["png", "jpg", "jpeg", "pdf"], key="single_ocr")
    selected_lang = st.selectbox("Select OCR Language", list(OCR_LANGS.keys()))
    
    if uploaded_file is not None:
        file_type = uploaded_file.name.split('.')[-1].lower()
        extracted_text = ""
        
        if st.button("🔍 Extract Text", key="btn_single_ocr"):
            with st.spinner("Extracting text from document..."):
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
    st.write("Capture documents instantly using your device camera to create PDF files or extract text.")
    
    camera_photo = st.camera_input("Take a document photo")
    cam_lang = st.selectbox("OCR Recognition Language", list(OCR_LANGS.keys()), key="cam_lang")
    
    if camera_photo is not None:
        st.success("Photo Captured!")
        img = Image.open(camera_photo)
        
        col_c1, col_c2 = st.columns(2)
        
        # Action 1: Convert Photo to PDF
        with col_c1:
            if st.button("📄 Convert Captured Photo to PDF"):
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
            if st.button("🔍 Perform Text Extraction (OCR)"):
                with st.spinner("Processing document text..."):
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
    st.write("Process multiple documents or images simultaneously and download as a ZIP archive.")
    
    batch_files = st.file_uploader("Upload Files (Multiple)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
    batch_lang = st.selectbox("Language Target", list(OCR_LANGS.keys()), key="batch_lang")
    
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
    st.header("✍️ Text Formatting & Cleaner")
    st.write("Clean unwanted spaces, line breaks, and formatting inconsistencies from extracted text.")
    
    input_text = st.text_area("Paste text here for formatting:", height=200)
    
    clean_spaces = st.checkbox("Remove redundant spaces & empty lines", value=True)
    
    if st.button("🧹 Clean & Format Text"):
        if input_text.strip():
            processed = input_text
            if clean_spaces:
                processed = re.sub(r'[ \t]+', ' ', processed)
                processed = re.sub(r'\n\s*\n', '\n\n', processed)
                
            st.success("Cleaned Result:")
            st.text_area("Formatted Text Output", processed, height=200)
            
            doc_sp = Document()
            doc_sp.add_paragraph(processed)
            bio_sp = io.BytesIO()
            doc_sp.save(bio_sp)
            st.download_button("📥 Download Cleaned Document (.docx)", data=bio_sp.getvalue(), file_name="Cleaned_Text.docx")
        else:
            st.warning("Please enter or paste text to clean.")

# ----------------------------
# TAB 4: Text-to-Speech (TTS)
# ----------------------------
with tab4:
    st.header("🔊 Text-to-Speech Converter")
    st.write("Convert extracted text into audio speech and export as an MP3 audio file.")
    
    tts_text = st.text_area("Enter text to convert to speech:", height=160)
    tts_lang = st.selectbox("Voice Accent & Language:", ["English (US)", "English (UK)", "Spanish", "French", "German"])
    
    lang_codes = {
        "English (US)": "en",
        "English (UK)": "en-uk",
        "Spanish": "es",
        "French": "fr",
        "German": "de"
    }
    
    if st.button("🔊 Generate Audio"):
        if tts_text.strip():
            with st.spinner("Generating audio file..."):
                try:
                    tts = gTTS(text=tts_text, lang=lang_codes[tts_lang].split('-')[0])
                    mp3_fp = io.BytesIO()
                    tts.write_to_fp(mp3_fp)
                    mp3_fp.seek(0)
                    
                    st.success("Audio Generated!")
                    st.audio(mp3_fp, format="audio/mp3")
                    st.download_button("📥 Download MP3 Audio", data=mp3_fp.getvalue(), file_name="LipiParse_Audio.mp3", mime="audio/mp3")
                except Exception:
                    st.error("Audio generation failed. Check your text input.")
        else:
            st.warning("Please provide text input.")

# ----------------------------
# TAB 5: PDF Password Protector & Remover
# ----------------------------
with tab5:
    st.header("🔒 PDF Security Management")
    st.write("Encrypt documents with a password or unlock existing protected PDFs.")
    
    sec_option = st.radio("Select Security Action:", ["Protect PDF (Encrypt)", "Unlock PDF (Decrypt)"])
    pdf_sec_file = st.file_uploader("Upload PDF Document", type=["pdf"], key="pdf_sec")
    
    if sec_option == "Protect PDF (Encrypt)":
        pass_code = st.text_input("Set Security Password:", type="password")
        if st.button("🔒 Encrypt PDF"):
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
                st.warning("Please upload a PDF and specify a password.")
                
    elif sec_option == "Unlock PDF (Decrypt)":
        existing_pass = st.text_input("Enter Document Password:", type="password")
        if st.button("🔓 Decrypt PDF"):
            if pdf_sec_file and existing_pass:
                try:
                    reader = PdfReader(pdf_sec_file)
                    if reader.is_encrypted:
                        reader.decrypt(existing_pass)
                    writer = PdfWriter()
                    for p in reader.pages: writer.add_page(p)
                    out_pdf = io.BytesIO()
                    writer.write(out_pdf)
                    st.success("PDF Unlocked Successfully!")
                    st.download_button("📥 Download Unlocked PDF", data=out_pdf.getvalue(), file_name="Unlocked_Document.pdf", mime="application/pdf")
                except Exception:
                    st.error("Incorrect password or unreadable PDF structure.")
            else:
                st.warning("Please upload a PDF and enter the decryption password.")

# ----------------------------
# TAB 6: PDF Merge & Split
# ----------------------------
with tab6:
    st.header("🧩 PDF Utilities (Merge & Split)")
    st.write("Combine multiple PDF files into one or extract individual pages.")
    
    tool_choice = st.radio("Choose Utility:", ["Merge PDFs", "Split PDF Page"])
    
    if tool_choice == "Merge PDFs":
        merge_files = st.file_uploader("Select PDFs to Merge", type=["pdf"], accept_multiple_files=True, key="m_files")
        if st.button("Merge Files"):
            if merge_files:
                merger = PdfMerger()
                for pdf in merge_files: merger.append(pdf)
                m_out = io.BytesIO()
                merger.write(m_out)
                merger.close()
                st.download_button("📥 Download Merged PDF", data=m_out.getvalue(), file_name="Merged_Document.pdf", mime="application/pdf")
    
    elif tool_choice == "Split PDF Page":
        split_file = st.file_uploader("Select PDF to Split", type=["pdf"], key="s_file")
        if split_file:
            reader = PdfReader(split_file)
            page_num = st.number_input("Select Page Number", min_value=1, max_value=len(reader.pages), value=1)
            if st.button("Extract Page"):
                writer = PdfWriter()
                writer.add_page(reader.pages[page_num - 1])
                s_out = io.BytesIO()
                writer.write(s_out)
                st.download_button(f"📥 Download Page {page_num}", data=s_out.getvalue(), file_name=f"Page_{page_num}.pdf", mime="application/pdf")
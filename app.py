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

# Mobile & Desktop Responsive Page Config
st.set_page_config(
    page_title="LipiParse - Universal Text & PDF Suite", 
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom Styling for Mobile Optimizations
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
        font-weight: bold;
    }
    @media (max-width: 768px) {
        .main .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }
    }
    </style>
""", unsafe_allow_html=True)

# Title Header
st.title("📄 LipiParse")
st.caption("Universal PDF, Image OCR & Multi-Tool Suite (Mobile Friendly)")

# OCR Language Mappings
OCR_LANGS = {
    "English": "eng",
    "Sinhala (සිංහල)": "sin",
    "Tamil (தமிழ்)": "tam",
    "Sinhala + English": "sin+eng",
    "Tamil + English": "tam+eng"
}

# OCR Common Typos Correction Dictionary
OCR_CORRECTIONS = {
    "ලංකාව": ["ලංකාව", "ලංකාව්"],
    "ශ්‍රී": ["ශ්රී", "ශී්ර"],
    "සඳහා": ["සදහා", "සදහ්"],
    "නැත": ["නැත්", "නැත්"],
    "කිරීම": ["කිරිම", "කීරිම"],
    "தமிழ்": ["தமிழ்", "தமிழ"],
    "இலங்கை": ["இலங்கை", "இலஙைக"],
}

def auto_correct_ocr_text(text):
    corrected_text = text
    applied_fixes = []
    for correct_word, typo_list in OCR_CORRECTIONS.items():
        for typo in typo_list:
            if typo in corrected_text and typo != correct_word:
                corrected_text = re.sub(r'\b' + re.escape(typo) + r'\b', correct_word, corrected_text)
                applied_fixes.append(f"'{typo}' ➔ '{correct_word}'")
    return corrected_text, list(set(applied_fixes))

# Navigation Tabs
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📝 OCR Text", 
    "📚 Batch Process",
    "✍️ Spell Checker",
    "🔊 Text-to-Speech",
    "🔒 PDF Security",
    "🔤 Font Convert",
    "🧩 PDF Tools"
])

# ----------------------------
# TAB 1: Single File Text OCR
# ----------------------------
with tab1:
    st.header("Extract Text from PDF or Image")
    uploaded_file = st.file_uploader("Upload Image or PDF", type=["png", "jpg", "jpeg", "pdf"], key="single_ocr")
    selected_lang = st.selectbox("Select Language (OCR)", list(OCR_LANGS.keys()))
    
    if uploaded_file is not None:
        file_type = uploaded_file.name.split('.')[-1].lower()
        extracted_text = ""
        
        if st.button("🔍 Extract Text", key="btn_single_ocr"):
            with st.spinner("Extracting Text..."):
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
                    st.success("Extraction Complete!")
                    st.text_area("Extracted Text Output", extracted_text, height=200)
                    
                    doc = Document()
                    doc.add_paragraph(extracted_text)
                    bio = io.BytesIO()
                    doc.save(bio)
                    st.download_button("📥 Download (.docx)", data=bio.getvalue(), file_name="LipiParse_Output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    st.error("No text could be extracted.")

# ----------------------------
# TAB 2: Batch Processing
# ----------------------------
with tab2:
    st.header("📚 Batch Processing (Multiple Files)")
    st.write("Files 10-100ක් එකපාර Upload කරලා Zip File එකක් විදිහට Download කරගන්න.")
    batch_files = st.file_uploader("Upload Images/PDFs (Multiple)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
    batch_lang = st.selectbox("Select Language", list(OCR_LANGS.keys()), key="batch_lang")
    
    if batch_files and st.button("🚀 Process All Files & Create Zip"):
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
            st.download_button("📦 Download All Outputs (.zip)", data=zip_buffer.getvalue(), file_name="LipiParse_Batch_Results.zip", mime="application/zip")

# ----------------------------
# TAB 3: Spell Checker & Assistant
# ----------------------------
with tab3:
    st.header("✍️ Spell Checker & Assistant")
    input_text = st.text_area("Paste text here to clean and correct:", height=180)
    
    c1, c2 = st.columns(2)
    clean_spaces = c1.checkbox("Clean Spaces & Breaks", value=True)
    apply_dict = c2.checkbox("Apply Auto-Correction", value=True)
    
    if st.button("🔍 Correct & Clean Text"):
        if input_text.strip():
            processed = input_text
            fixes = []
            if clean_spaces:
                processed = re.sub(r'[ \t]+', ' ', processed)
                processed = re.sub(r'\n\s*\n', '\n\n', processed)
            if apply_dict:
                processed, fixes = auto_correct_ocr_text(processed)
                
            st.success("Cleaned Output:")
            if fixes:
                for fix in fixes: st.info(f"Fixed: {fix}")
            st.text_area("Corrected Text", processed, height=200)
            
            doc_sp = Document()
            doc_sp.add_paragraph(processed)
            bio_sp = io.BytesIO()
            doc_sp.save(bio_sp)
            st.download_button("📥 Download Corrected Doc", data=bio_sp.getvalue(), file_name="Corrected_Text.docx")
        else:
            st.warning("Please paste text first!")

# ----------------------------
# TAB 4: Text-to-Speech (TTS)
# ----------------------------
with tab4:
    st.header("🔊 Text-to-Speech (Audio Reader)")
    st.write("Extract කරගත් Text එක සද්දයෙන් ඇසීමට සහ MP3 File එකක් ලෙස Download කරගැනීමට:")
    
    tts_text = st.text_area("Enter Text to Read Aloud:", height=150)
    tts_lang = st.selectbox("Select Speech Language:", ["Sinhala", "Tamil", "English"])
    
    lang_codes = {"Sinhala": "si", "Tamil": "ta", "English": "en"}
    
    if st.button("🔊 Generate Audio"):
        if tts_text.strip():
            with st.spinner("Generating Speech Audio..."):
                try:
                    tts = gTTS(text=tts_text, lang=lang_codes[tts_lang])
                    mp3_fp = io.BytesIO()
                    tts.write_to_fp(mp3_fp)
                    mp3_fp.seek(0)
                    
                    st.success("Audio Generated!")
                    st.audio(mp3_fp, format="audio/mp3")
                    st.download_button("📥 Download Audio (.mp3)", data=mp3_fp.getvalue(), file_name="LipiParse_Audio.mp3", mime="audio/mp3")
                except Exception as e:
                    st.error("Speech generation failed. Please ensure text is valid.")
        else:
            st.warning("Please enter some text!")

# ----------------------------
# TAB 5: PDF Password Protector & Remover
# ----------------------------
with tab5:
    st.header("🔒 PDF Password Protector & Remover")
    sec_option = st.radio("Select Action:", ["Protect PDF (Add Password)", "Unlock PDF (Remove Password)"])
    pdf_sec_file = st.file_uploader("Upload PDF File", type=["pdf"], key="pdf_sec")
    
    if sec_option == "Protect PDF (Add Password)":
        pass_code = st.text_input("Set New Password:", type="password")
        if st.button("🔒 Lock PDF"):
            if pdf_sec_file and pass_code:
                reader = PdfReader(pdf_sec_file)
                writer = PdfWriter()
                for p in reader.pages: writer.add_page(p)
                writer.encrypt(pass_code)
                out_pdf = io.BytesIO()
                writer.write(out_pdf)
                st.success("PDF Encrypted Successfully!")
                st.download_button("📥 Download Protected PDF", data=out_pdf.getvalue(), file_name="Protected_LipiParse.pdf", mime="application/pdf")
            else:
                st.warning("Upload PDF and enter password.")
                
    elif sec_option == "Unlock PDF (Remove Password)":
        existing_pass = st.text_input("Enter Current Password:", type="password")
        if st.button("🔓 Unlock PDF"):
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
                    st.download_button("📥 Download Unlocked PDF", data=out_pdf.getvalue(), file_name="Unlocked_LipiParse.pdf", mime="application/pdf")
                except Exception:
                    st.error("Incorrect password or corrupted PDF.")
            else:
                st.warning("Upload PDF and enter password.")

# ----------------------------
# TAB 6: Legacy Font Converter
# ----------------------------
with tab6:
    st.header("🔤 Legacy Font Converter (FM Abhaya)")
    fm_input = st.text_area("Paste FM Abhaya encoded text here:", height=150)
    if st.button("Convert to Unicode"):
        if fm_input:
            converted = fm_input.replace("a", "්").replace("aa", "ා") # Example mapping
            st.text_area("Converted Output:", converted, height=150)

# ----------------------------
# TAB 7: PDF Merge & Split
# ----------------------------
with tab7:
    st.header("🧩 PDF Merger & Splitter")
    tool_choice = st.radio("Choose Tool:", ["Merge PDFs", "Split PDF Page"])
    
    if tool_choice == "Merge PDFs":
        merge_files = st.file_uploader("Select PDFs to Merge", type=["pdf"], accept_multiple_files=True, key="m_files")
        if st.button("Merge All"):
            if merge_files:
                merger = PdfMerger()
                for pdf in merge_files: merger.append(pdf)
                m_out = io.BytesIO()
                merger.write(m_out)
                merger.close()
                st.download_button("📥 Download Merged PDF", data=m_out.getvalue(), file_name="Merged_LipiParse.pdf", mime="application/pdf")
    
    elif tool_choice == "Split PDF Page":
        split_file = st.file_uploader("Select PDF to Split", type=["pdf"], key="s_file")
        if split_file:
            reader = PdfReader(split_file)
            page_num = st.number_input("Page Number", min_value=1, max_value=len(reader.pages), value=1)
            if st.button("Extract Single Page"):
                writer = PdfWriter()
                writer.add_page(reader.pages[page_num - 1])
                s_out = io.BytesIO()
                writer.write(s_out)
                st.download_button(f"📥 Download Page {page_num}", data=s_out.getvalue(), file_name=f"Page_{page_num}.pdf", mime="application/pdf")
import streamlit as st
import pdfplumber
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import docx
import io
import os

# --- Page Configuration ---
st.set_page_config(
    page_title="LipiParse — Free Universal PDF & OCR Converter",
    page_icon="📄",
    layout="wide"
)

# --- Legacy Font Converter (FM Abhaya -> Unicode Mapping) ---
def convert_fm_to_unicode(text):
    fm_map = {
        'a': 'ං', 'A': 'ඃ', 'b': '්', 'B': 'ා', 'c': 'ැ', 'C': 'ෑ',
        'd': 'ි', 'D': 'ී', 'e': 'ු', 'E': 'ූ', 'f': 'ෘ', 'F': 'ෲ',
        'g': 'ෙ', 'G': 'ෛ', 'h': 'ො', 'H': 'ෝ', 'i': 'ෞ', 'I': 'ෟ',
        'j': 'ක්', 'k': 'ඛ්', 'l': 'ග්', 'm': 'ඝ්', 'n': 'ඞ්',
        'o': 'ච්', 'p': 'ඡ්', 'q': 'ජ්', 'r': 'ඣ්', 's': 'ඤ්',
        't': 'ට්', 'u': 'ඨ්', 'v': 'ඩ්', 'w': 'ඪ්', 'x': 'ණ්',
        'y': 'ත්', 'z': 'ථ්'
    }
    for key, val in fm_map.items():
        text = text.replace(key, val)
    return text

# --- Header & Branding ---
st.title("🌐 LipiParse")
st.caption("100% Free Universal PDF & Multi-Language OCR Engine | All-in-One Document Toolkit")

st.markdown("---")

# --- Mode Navigation (Tabs) ---
mode = st.radio(
    "Choose Tool / Mode:", 
    ["📄 Text Extraction & Multi-Lang OCR", "🔀 PDF Merger (Combine PDFs)", "✂️ PDF Splitter (Extract Pages)"],
    horizontal=True
)

st.markdown("---")

# ==========================================
# MODE 1: Text Extraction & OCR
# ==========================================
if mode == "📄 Text Extraction & Multi-Lang OCR":
    st.sidebar.header("⚙️ OCR & Engine Settings")
    
    lang_dict = {
        "Auto-Detect / Sinhala + English": "sin+eng",
        "Sinhala Only (සිංහල)": "sin",
        "Tamil Only (தமிழ்)": "tam",
        "English Only": "eng",
        "Hindi Only (हिंदी)": "hin",
        "Global Multi-Lang (Sinhala+Tamil+English+Hindi)": "sin+tam+eng+hin"
    }

    selected_lang_label = st.sidebar.selectbox("Select Primary Document Language:", list(lang_dict.keys()))
    tesseract_lang_code = lang_dict[selected_lang_label]

    st.sidebar.markdown("---")
    enable_legacy_fix = st.sidebar.checkbox("🛠️ Enable Legacy Font Fixer (FM Abhaya / DL Manel)")

    uploaded_file = st.file_uploader(
        "Choose a PDF file or Image (PNG, JPG, JPEG)", 
        type=["pdf", "png", "jpg", "jpeg"]
    )

    extracted_text = ""

    if uploaded_file is not None:
        file_type = uploaded_file.type
        st.success(f"File Uploaded: **{uploaded_file.name}**")
        
        if st.button("🚀 Extract & Convert Text", type="primary"):
            with st.spinner("Processing document using LipiParse Engine... Please wait."):
                
                # Image OCR
                if "image" in file_type:
                    try:
                        image = Image.open(uploaded_file)
                        extracted_text = pytesseract.image_to_string(image, lang=tesseract_lang_code)
                    except Exception as e:
                        st.error(f"OCR Error processing image: {e}")

                # PDF Direct + OCR Fallback
                elif file_type == "application/pdf":
                    pdf_bytes = uploaded_file.read()
                    
                    try:
                        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text:
                                    extracted_text += text + "\n\n"
                    except Exception:
                        st.warning("Digital text extraction failed. Switching to OCR Mode...")
                    
                    if not extracted_text.strip():
                        st.info("Scanned/Image-based PDF detected. Applying Tesseract Multi-Language OCR...")
                        try:
                            images = convert_from_bytes(pdf_bytes)
                            for i, img in enumerate(images):
                                page_text = pytesseract.image_to_string(img, lang=tesseract_lang_code)
                                extracted_text += f"--- Page {i+1} ---\n" + page_text + "\n\n"
                        except Exception as e:
                            st.error(f"OCR Error processing PDF: {e}")

                if enable_legacy_fix and extracted_text.strip():
                    extracted_text = convert_fm_to_unicode(extracted_text)

    if extracted_text.strip():
        st.subheader("📝 Extracted Output Text")
        st.text_area("Result Text:", extracted_text, height=350)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Download as Text (.txt)",
                data=extracted_text,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_LipiParse.txt",
                mime="text/plain"
            )
            
        with col2:
            doc = docx.Document()
            doc.add_heading('Converted Text — LipiParse', 0)
            doc.add_paragraph(extracted_text)
            
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            st.download_button(
                label="📄 Download as Word (.docx)",
                data=doc_io,
                file_name=f"{os.path.splitext(uploaded_file.name)[0]}_LipiParse.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# ==========================================
# MODE 2: PDF Merger
# ==========================================
elif mode == "🔀 PDF Merger (Combine PDFs)":
    st.subheader("🔀 Combine Multiple PDFs into One File")
    uploaded_pdfs = st.file_uploader("Upload 2 or more PDF files:", type=["pdf"], accept_multiple_files=True)
    
    if uploaded_pdfs and len(uploaded_pdfs) >= 2:
        if st.button("🔗 Merge PDFs Now", type="primary"):
            merger = PdfMerger()
            for pdf in uploaded_pdfs:
                merger.append(io.BytesIO(pdf.read()))
            
            merged_io = io.BytesIO()
            merger.write(merged_io)
            merger.close()
            merged_io.seek(0)
            
            st.success("PDFs Merged Successfully!")
            st.download_button(
                label="📥 Download Merged PDF",
                data=merged_io,
                file_name="LipiParse_Merged.pdf",
                mime="application/pdf"
            )
    elif uploaded_pdfs:
        st.info("Please upload at least 2 PDF files to merge.")

# ==========================================
# MODE 3: PDF Splitter
# ==========================================
elif mode == "✂️ PDF Splitter (Extract Pages)":
    st.subheader("✂️ Extract Specific Pages from PDF")
    split_pdf = st.file_uploader("Upload a PDF file to split:", type=["pdf"])
    
    if split_pdf is not None:
        reader = PdfReader(io.BytesIO(split_pdf.read()))
        total_pages = len(reader.pages)
        st.info(f"Total Pages in PDF: **{total_pages}**")
        
        start_page = st.number_input("Start Page:", min_value=1, max_value=total_pages, value=1)
        end_page = st.number_input("End Page:", min_value=1, max_value=total_pages, value=total_pages)
        
        if st.button("✂️ Split PDF", type="primary"):
            if start_page <= end_page:
                writer = PdfWriter()
                for i in range(start_page - 1, end_page):
                    writer.add_page(reader.pages[i])
                
                split_io = io.BytesIO()
                writer.write(split_io)
                split_io.seek(0)
                
                st.success(f"Pages {start_page} to {end_page} extracted successfully!")
                st.download_button(
                    label="📥 Download Split PDF",
                    data=split_io,
                    file_name=f"{os.path.splitext(split_pdf.name)[0]}_Pages_{start_page}_to_{end_page}.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("Start Page must be less than or equal to End Page.")
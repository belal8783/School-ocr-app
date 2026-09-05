import os
import io
import re
import time
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image
from pdf2image import convert_from_bytes

# ---------------------------------------------------------
# Unicode Bengali Cleaning Routine
# ---------------------------------------------------------
def clean_bengali_symbols(text):
    if not text:
        return ""
    # অতিরিক্ত ডটেড সার্কেল বা ইনভ্যালিড ক্যারেক্টার রিমুভ
    cleaned_text = re.sub(r'\u25cc', '', text)
    return cleaned_text.replace('◌', '')

def is_arabic(text):
    return bool(re.search(r'[\u0600-\u06FF]', text))

# ---------------------------------------------------------
# Page Config & Branding
# ---------------------------------------------------------
st.set_page_config(
    page_title="Handwritten to Word & PDF Agent",
    page_icon="📝",
    layout="wide"
)

st.title("📝 School Sheet Handwritten to Word & PDF Agent")
st.caption("Developed by: Belal Hossain | নিখুঁত প্রফেশনাল কনভার্সন সিস্টেম")

# ---------------------------------------------------------
# Sidebar Options
# ---------------------------------------------------------
st.sidebar.header("⚙️ কাস্টমাইজেশন সেটিং")

font_size = st.sidebar.slider("মূল টেক্সট ফন্ট সাইজ (Font Size)", min_value=10, max_value=24, value=13, step=1)

bangla_font_name = st.sidebar.selectbox(
    "বাংলা ফন্ট নির্বাচন করুন (ওয়ার্ডের জন্য)",
    ["Kalpurush", "SolaimanLipi", "Siyam Rupali", "Arial"]
)

english_font_name = st.sidebar.selectbox(
    "ইংরেজি ফন্ট নির্বাচন করুন",
    ["Times New Roman", "Calibri", "Arial"]
)

arabic_font_name = st.sidebar.selectbox(
    "আরবি ফন্ট নির্বাচন করুন",
    ["Traditional Arabic", "Amiri", "Scheherazade"]
)

# ---------------------------------------------------------
# API Key Configuration
# ---------------------------------------------------------
api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")

if not api_key:
    api_key = st.sidebar.text_input("Gemini API Key দিন:", type="password")

if not api_key:
    st.info("⚠️ অ্যাপটি চালাতে Gemini API Key প্রয়োজন।")
    st.stop()

genai.configure(api_key=api_key)

# ---------------------------------------------------------
# Model Selection (অফিশিয়াল এবং পাওয়ারফুল Gemini 2.0 Flash)
# ---------------------------------------------------------
MODEL_NAME = 'gemini-2.0-flash'
model = genai.GenerativeModel(MODEL_NAME)

# ---------------------------------------------------------
# Safety Settings (যাতে প্রসেসিং ফিল্টার না হয়)
# ---------------------------------------------------------
safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
]

# ---------------------------------------------------------
# File Upload & Session State Setup
# ---------------------------------------------------------
uploaded_files = st.file_uploader(
    "আপনার শিটের ছবি (JPG, PNG) অথবা PDF ফাইল আপলোড করুন",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

custom_filename = st.text_input("ডাউনলোড ফাইলের নাম লিখুন:", value="Converted_School_Sheet")

if 'final_converted_text' not in st.session_state:
    st.session_state.final_converted_text = ""
if 'conversion_done' not in st.session_state:
    st.session_state.conversion_done = False

# ---------------------------------------------------------
# Master AI Prompt (বাংলা ও গণিত নিখুঁত করার নির্দেশ)
# ---------------------------------------------------------
SYSTEM_PROMPT = """
আপনি একজন বিশেষজ্ঞ OCR ও বাংলা ডকুমেন্ট কনভার্সন সিস্টেম। 
ছবিতে থাকা হ্যান্ডরিটেন বা টাইপ করা টেক্সটগুলো হুবহু সঠিক প্রমিত ইউনিকোড বাংলায় রূপান্তর করুন।
- বাংলা যুক্তবর্ণ, মাত্রা এবং সংখ্যা (যেমন: ১, ২, ৩, ৪) যেন কোনোভাবেই না ভাঙে।
- হেডিং থাকলে শুরুতে `# ` দিন।
- আন্ডারলাইন থাকলে `<u>লেখা</u>` দিন।
- ছক, টেবিল বা লিস্ট থাকলে তা সুন্দরভাবে সাজিয়ে উপস্থাপন করুন।
- কোনো অনাকাঙ্ক্ষিত অক্ষর বা হাবিজাবি চিহ্ন আউটপুটে দেবেন না।
"""

# ---------------------------------------------------------
# Word Document Generator
# ---------------------------------------------------------
def create_word_docx(text, bangla_font, english_font, arabic_font, font_size):
    doc = Document()
    lines = text.split('\n')
    for line in lines:
        if not line.strip():
            continue
            
        p = doc.add_paragraph()
        is_heading = line.startswith('#')
        current_line = re.sub(r'^#+\s*', '', line) if is_heading else line
            
        parts = re.split(r'(<u>.*?</u>|\[ইমেজ নোট:.*?\])', current_line)
        
        for part in parts:
            if not part:
                continue
                
            run = p.add_run()
            
            if part.startswith("[ইমেজ নোট:"):
                run.text = part
                run.font.bold = True
                run.font.color.rgb = RGBColor(180, 50, 50)
                run.font.name = bangla_font
                run.font.size = Pt(font_size)
            elif part.startswith("<u>") and part.endswith("</u>"):
                clean_text = part[3:-4]
                run.text = clean_text
                run.font.underline = True
                run.font.bold = is_heading
                run.font.size = Pt(font_size + 3) if is_heading else Pt(font_size)
                run.font.name = arabic_font if is_arabic(clean_text) else (english_font if clean_text.isascii() else bangla_font)
            else:
                run.text = part
                run.font.bold = is_heading
                run.font.size = Pt(font_size + 3) if is_heading else Pt(font_size)
                
                if is_arabic(part):
                    run.font.name = arabic_font
                elif part.isascii():
                    run.font.name = english_font
                else:
                    run.font.name = bangla_font
                    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# PDF Document Generator
# ---------------------------------------------------------
def create_pdf(text, font_size):
    buffer = io.BytesIO()
    pdf = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    
    custom_style = ParagraphStyle(
        'CustomStyle',
        parent=styles['Normal'],
        fontSize=font_size,
        leading=font_size + 5
    )
    
    story = []
    for line in text.split('\n'):
        if line.strip():
            safe_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(safe_line, custom_style))
            story.append(Spacer(1, 4))
            
    pdf.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Execution Logic with Automatic Rate Limit Retry
# ---------------------------------------------------------
if uploaded_files and st.button("🚀 কনভার্ট শুরু করুন"):
    combined_result = ""
    images_to_process = []

    with st.spinner("ফাইল মেমোরিতে লোড করা হচ্ছে..."):
        for uploaded_file in uploaded_files:
            if uploaded_file.name.lower().endswith(".pdf"):
                pdf_bytes = uploaded_file.read()
                converted_images = convert_from_bytes(pdf_bytes)
                images_to_process.extend(converted_images)
            else:
                images_to_process.append(Image.open(uploaded_file))

    total_pages = len(images_to_process)
    progress_bar = st.progress(0)
    status_text = st.empty()

    for index, img in enumerate(images_to_process):
        max_retries = 5
        success = False
        
        for attempt in range(max_retries):
            status_text.text(f"প্রসেসিং চলছে ({MODEL_NAME}): পৃষ্ঠা {index + 1} / {total_pages} (চেষ্টা: {attempt + 1})")
            try:
                response = model.generate_content(
                    [SYSTEM_PROMPT, img],
                    safety_settings=safety_settings
                )
                
                page_text = ""
                if response and hasattr(response, 'text'):
                    page_text = clean_bengali_symbols(response.text)
                
                if page_text.strip():
                    combined_result += f"\n\n--- পৃষ্ঠা {index + 1} ---\n\n" + page_text
                    success = True
                    break
                else:
                    time.sleep(2)
                
            except Exception as e:
                err_msg = str(e)
                if "429" in err_msg or "Quota" in err_msg or "ResourceExhausted" in err_msg:
                    status_text.warning(f"⚠️ এপিআই লিমিটের জন্য ১০ সেকেন্ড অপেক্ষা করে অটো-রিট্রি করা হচ্ছে (পৃষ্ঠা {index + 1})...")
                    time.sleep(10)
                else:
                    time.sleep(3)
        
        if not success:
            combined_result += f"\n\n--- পৃষ্ঠা {index + 1} ---\n\n[এরর: পৃষ্ঠাটি সঠিকভাবে প্রসেস করা সম্ভব হয়নি।]"
            
        progress_bar.progress((index + 1) / total_pages)

    status_text.empty()
    
    st.session_state.final_converted_text = combined_result
    st.session_state.conversion_done = True

# ---------------------------------------------------------
# Display Output & Downloads
# ---------------------------------------------------------
if st.session_state.conversion_done:
    st.success("✅ কনভার্ট সম্পন্ন হয়েছে!")

    st.subheader("📝 কনভার্ট হওয়া টেক্সট প্রিভিউ:")
    st.text_area("আউটপুট টেক্সট (ইউনিকোড):", st.session_state.final_converted_text, height=350)

    col1, col2 = st.columns(2)
    
    word_file = create_word_docx(st.session_state.final_converted_text, bangla_font_name, english_font_name, arabic_font_name, font_size)
    pdf_file = create_pdf(st.session_state.final_converted_text, font_size)

    with col1:
        st.download_button(
            label="📄 Word (.docx) ডাউনলোড করুন",
            data=word_file,
            file_name=f"{custom_filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    with col2:
        st.download_button(
            label="📕 PDF (.pdf) ডাউনলোড করুন",
            data=pdf_file,
            file_name=f"{custom_filename}.pdf",
            mime="application/pdf"
        )

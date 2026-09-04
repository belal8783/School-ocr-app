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
    cleaned_text = re.sub(r'\u25cc', '', text)
    cleaned_text = cleaned_text.replace('◌', '')
    return cleaned_text

# ---------------------------------------------------------
# Helper to detect Arabic scripts
# ---------------------------------------------------------
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
st.caption("Developed by: Belal Hossain | আপনার স্কুলের হাতের লেখা শিট ও পিডিএফ কনভার্ট করুন")

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
model = genai.GenerativeModel('gemini-3.6-flash')

# ---------------------------------------------------------
# File Upload & Session State setup
# ---------------------------------------------------------
uploaded_files = st.file_uploader(
    "আপনার শিটের ছবি (JPG, PNG) অথবা PDF ফাইল আপলোড করুন",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

custom_filename = st.text_input("ডাউনলোড ফাইলের নাম লিখুন:", value="Converted_School_Sheet")

# সেশন স্টেট (ফাঁকা ডাউনলোড রোধ করার জন্য)
if 'final_converted_text' not in st.session_state:
    st.session_state.final_converted_text = ""
    st.session_state.conversion_successful = False

# ---------------------------------------------------------
# Master AI Prompt
# ---------------------------------------------------------
SYSTEM_PROMPT = """
আপনি একজন বিশেষজ্ঞ OCR ও ডকুমেন্ট কনভার্সন সিস্টেম। 
হ্যান্ডরিটেন শিট থেকে তথ্যগুলো নিখুঁতভাবে বাংলা ইউনিকোড টেক্সটে রূপান্তর করুন।

নির্দেশাবলী:
১. হেডিং থাকলে শুরুতে `# ` ব্যবহার করুন।
২. আন্ডারলাইন থাকলে `<u>লেখা</u>` ট্যাগ ব্যবহার করুন।
৩. কাটাকাটি বা বাতিলকৃত লেখা বাদ দিয়ে সঠিক সংশোধিত রূপটি লিখুন।
৪. ছক বা টেবিল থাকলে তা Markdown Table ফরম্যাটে রাখুন।
৫. কোনো অবোধ্য বা অস্পষ্ট বাংলা যুক্তবর্ণ থাকলে তা ভাঙবেন না, প্রমিত ইউনিকোড বাংলায় লিখুন।
"""

# Safety Settings: AI যেন কোনো লেখা সেন্সর বা ব্লক না করে
safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
]

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
        
        is_heading = False
        current_line = line
        if current_line.startswith('#'):
            is_heading = True
            current_line = re.sub(r'^#+\s*', '', current_line)
            
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
# Main Execution Logic
# ---------------------------------------------------------
if uploaded_files and st.button("🚀 কনভার্ট শুরু করুন"):
    combined_result = ""
    images_to_process = []

    with st.spinner("ফাইল লোড করা হচ্ছে..."):
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
        status_text.text(f"প্রসেসিং চলছে: পৃষ্ঠা {index + 1} / {total_pages}")
        
        if index > 0:
            time.sleep(3) # API Limit রক্ষা করতে বিরতি
            
        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = model.generate_content(
                    [SYSTEM_PROMPT, img],
                    safety_settings=safety_settings
                )
                
                # Check if API blocked it despite settings
                try:
                    page_text = response.text
                except ValueError:
                    page_text = f"[সতর্কতা: পৃষ্ঠা {index + 1} গুগল সেফটি ফিল্টারের কারণে স্কিপ করা হয়েছে।]"

                cleaned_page_text = clean_bengali_symbols(page_text)
                combined_result += f"\n\n--- পৃষ্ঠা {index + 1} ---\n\n"
                combined_result += cleaned_page_text
                break
            except Exception as e:
                if "429" in str(e) or "ResourceExhausted" in str(e):
                    time.sleep(10)
                    continue
                st.error(f"পৃষ্ঠা {index + 1} এ এরর: {e}")
                break
                
        progress_bar.progress((index + 1) / total_pages)

    status_text.empty()
    
    # ডেটা সেশন স্টেটে সেভ করা
    st.session_state.final_converted_text = combined_result
    st.session_state.conversion_successful = True
    
# ---------------------------------------------------------
# Display Results & Download (বাইরে রাখা হয়েছে যাতে রিফ্রেশ হলেও ডেটা না হারায়)
# ---------------------------------------------------------
if st.session_state.conversion_successful:
    st.success("✅ সফলভাবে রূপান্তর সম্পন্ন হয়েছে!")

    st.subheader("📝 কনভার্ট হওয়া টেক্সট প্রিভিউ:")
    st.text_area("আউটপুট টেক্সট (ইউনিকোড):", st.session_state.final_converted_text, height=300)

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

import os
import io
import re
import time
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image
from pdf2image import convert_from_bytes

# ---------------------------------------------------------
# Unicode Bengali Cleaning Routine (Fixing Dotted Circles ◌)
# ---------------------------------------------------------
def clean_bengali_symbols(text):
    if not text:
        return ""
    cleaned_text = re.sub(r'\u25cc', '', text)
    cleaned_text = cleaned_text.replace('◌', '')
    return cleaned_text

# ---------------------------------------------------------
# Unicode to Bijoy (SutonnyMJ) Converter Utility
# ---------------------------------------------------------
def convert_unicode_to_bijoy(text):
    if not text:
        return ""
    conversions = {
        'অ': 'A', 'আ': 'Av', 'ই': 'Bi', 'ঈ': 'C', 'উ': 'D', 'ঊ': 'E', 'ঋ': 'F', 'এ': 'G', 'ঐ': 'H', 'ও': 'I', 'ঔ': 'J',
        'ক': 'k', 'খ': 'L', 'গ': 'M', 'ঘ': 'N', 'ঙ': 'O',
        'চ': 'P', 'ছ': 'Q', 'জ': 'R', 'ঝ': 'S', 'ঞ': 'T',
        'ট': 'U', 'ঠ': 'V', 'ড': 'W', 'ঢ': 'X', 'ণ': 'Y',
        'ত': 'Z', 'থ': '_', 'দ': 'b', 'ধ': 'c', 'ন': 'd',
        'প': 'e', 'ফ': 'f', 'ব': 'g', 'ভ': 'h', 'ম': 'm',
        'য': 'n', 'র': 'r', 'ল': 'l', 'শ': 'k', 'ষ': 'l', 'স': 'm', 'হ': 'n', 'ড়': 'o', 'ঢ়': 'p', 'য়': 'q',
        'া': 'v', 'ি': 'w', 'ী': 'x', 'ু': 'y', 'ূ': 'z', 'ৃ': 'A', 'ে': 'B', 'ৈ': 'C', 'ো': 'Dv', 'ৌ': 'Dv',
        '্': '', 'ং': 's', 'ঃ': 't', 'ঁ': 'u'
    }
    converted_text = text
    for u_char, b_char in conversions.items():
        converted_text = converted_text.replace(u_char, b_char)
    return converted_text

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
st.write("বাংলা, ইংরেজি, আরবি, গণিত, হেডিং-আন্ডারলাইন, কাটাকাটি সংশোধন ও ছক সম্বলিত ছবি/পিডিএফ ফাইল সহজে ওয়ার্ড ও পিডিএফে কনভার্ট করুন।")

# ---------------------------------------------------------
# Sidebar Options
# ---------------------------------------------------------
st.sidebar.header("⚙️ কাস্টমাইজেশন সেটিং")

# ১. ফন্ট সাইজ
font_size = st.sidebar.slider("মূল টেক্সট ফন্ট সাইজ (Font Size)", min_value=10, max_value=24, value=12, step=1)

# ২. বাংলা ফন্ট
bangla_font_type = st.sidebar.selectbox(
    "বাংলা ফন্ট নির্বাচন করুন",
    ["Avro / Unicode (Kalpurush)", "Bijoy 52 (SutonnyMJ)"]
)
bangla_font_name = "SutonnyMJ" if "SutonnyMJ" in bangla_font_type else "Kalpurush"

# ৩. ইংরেজি ফন্ট
english_font_name = st.sidebar.selectbox(
    "ইংরেজি ফন্ট নির্বাচন করুন",
    ["Times New Roman", "Calibri", "Arial"]
)

# ৪. আরবি ফন্ট
arabic_font_name = st.sidebar.selectbox(
    "আরবি ফন্ট নির্বাচন করুন",
    ["Traditional Arabic", "Amiri", "Scheherazade"]
)

# ---------------------------------------------------------
# API Key Configuration
# ---------------------------------------------------------
api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")

if not api_key:
    api_key = st.text_input("Gemini API Key দিন:", type="password")

if not api_key:
    st.info("অ্যাপটি চালাতে Gemini API Key সরবরাহ করতে হবে।")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel('gemini-3.6-flash')

# ---------------------------------------------------------
# File Upload (Images & PDF Allowed)
# ---------------------------------------------------------
uploaded_files = st.file_uploader(
    "আপনার শিটের ছবি (JPG, PNG) অথবা PDF ফাইল আপলোড করুন (সর্বোচ্চ ৫০টি ছবি/১০MB PDF)",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

custom_filename = st.text_input("ডাউনলোড ফাইলের নাম লিখুন:", value="Converted_School_Sheet")

# ---------------------------------------------------------
# Master AI Prompt
# ---------------------------------------------------------
SYSTEM_PROMPT = """
আপনি একজন বিশেষজ্ঞ OCR ও ডকুমেন্ট কনভার্সন সিস্টেম। 
আপনাকে দেয়া হ্যান্ডরিটেন শিট বা পেজের ছবি থেকে তথ্যগুলো নিখুঁতভাবে টেক্সটে রূপান্তর করুন।

প্রধান নির্দেশাবলী:
১. **হেডিং ও আন্ডারলাইন (Headings & Underline):**
   - শিটে যদি কোনো শিরোনাম, প্রশ্নের নম্বর বা প্রধান হেডিং থাকে, সেটির শুরুতে `# ` ব্যবহার করুন (যেমন: `# প্রশ্ন ১: উত্তর লিখ`).
   - শিটে কোনো শব্দ বা লাইনের নিচে দাগ/আন্ডারলাইন টানা থাকলে সেটিকে `<u>লেখা</u>` ট্যাগ দিয়ে চিহ্নিত করুন।

২. **কাটাকাটি বা বাতিলকৃত লেখা (Crossed-out text):**
   - শিটে যদি কোনো শব্দ, সংখ্যা বা লাইন দাগ দিয়ে কেটে দেওয়া থাকে, সেটি আউটপুটে পুরোপুরি বাদ দিন। 
   - কেটে দেওয়ার পর আশেপাশে (উপরে/নিচে/পাশে) যে নতুন সংশোধনটি লেখা হয়েছে, শুধুমাত্র সেটিই আউটপুটে গ্রহণ করুন।

৩. **ছক ও টেবিল (Tables & Grids):**
   - কোনো ছক বা ঘর থাকলে সেটি Markdown Table (যেমন: | কলাম ১ | কলাম ২ |) হিসেবে কলাম ও সারি ঠিক রেখে রূপান্তর করুন। 
   - ঘরের ভেতর বাংলা, ইংরেজি, অংক বা আরবি যাই থাকুক না কেন, নির্দিষ্ট ঘরের ভেতরেই রাখুন।

৪. **ছবি বা ইলাস্ট্রেশনের নির্দেশ (Image Instructions):**
   - শিটে যদি লেখা থাকে "এখানে একটি বাঘের ছবি হবে", "পাখির ছবি আঁকুন" ইত্যাদি, তবে সেটিকে **[ইমেজ নোট: এখানে একটি বাঘের ছবি হবে]** এভাবে ব্র্যাকেটে বোল্ড আকারে তুলে ধরুন।

৫. **বাংলা কার-চিহ্ন ও প্রতীক:**
   - বাংলা আকার, একার, ওকার, ঋ-কার ইত্যাদি কার-চিহ্ন আলাদা লেখা থাকলে কোনো বাড়তি গোল দাগ বা ডটেড চিহ্ন (◌) ব্যবহার করবেন না। সরাসরি শুধু কার-চিহ্নটি (যেমন: া, ি, ো, ৃ) আউটপুটে লিখুন।

৬. **সাধারণ টেক্সট ও ভাষা:**
   - বাংলা, ইংরেজি, আরবি ও গাণিতিক সমীকরণগুলো যেভাবে লেখা আছে হুবহু তুলে আনুন।
"""

# ---------------------------------------------------------
# Advanced Word Document Generator
# ---------------------------------------------------------
def create_word_docx(text, bangla_font, english_font, arabic_font, font_size):
    doc = Document()
    
    lines = text.split('\n')
    for line in lines:
        if not line.strip():
            continue
            
        p = doc.add_paragraph()
        
        # হেডিং চেকিং
        is_heading = False
        current_line = line
        if current_line.startswith('#'):
            is_heading = True
            current_line = re.sub(r'^#+\s*', '', current_line)
            
        # স্বরচিহ্ন ও বিজয়ী কাস্টমাইজেশন
        if bangla_font == "SutonnyMJ":
            current_line = convert_unicode_to_bijoy(current_line)
            
        # আন্ডারলাইন প্রসেসিং (<u>tags</u>)
        parts = re.split(r'(<u>.*?</u>|\[ইমেজ নোট:.*?\])', current_line)
        
        for part in parts:
            if not part:
                continue
                
            run = p.add_run()
            
            # ইমেজ নোট হ্যান্ডলিং
            if part.startswith("[ইমেজ নোট:"):
                run.text = part
                run.font.bold = True
                run.font.color.rgb = RGBColor(180, 50, 50)
                run.font.name = bangla_font
                run.font.size = Pt(font_size)
            # আন্ডারলাইন হ্যান্ডলিং
            elif part.startswith("<u>") and part.endswith("</u>"):
                clean_text = part[3:-4]
                run.text = clean_text
                run.font.underline = True
                if is_heading:
                    run.font.bold = True
                    run.font.size = Pt(font_size + 4)
                else:
                    run.font.size = Pt(font_size)
                run.font.name = arabic_font if is_arabic(clean_text) else (english_font if clean_text.isascii() else bangla_font)
            # সাধারণ টেক্সট ও হেডিং
            else:
                run.text = part
                if is_heading:
                    run.font.bold = True
                    run.font.size = Pt(font_size + 4)
                else:
                    run.font.size = Pt(font_size)
                    
                # ভাষা অনুযায়ী ফন্ট সেট
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
        leading=font_size + 4
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

    with st.spinner("ফাইল প্রস্তুত করা হচ্ছে..."):
        for uploaded_file in uploaded_files:
            if uploaded_file.name.lower().endswith(".pdf"):
                pdf_bytes = uploaded_file.read()
                converted_images = convert_from_bytes(pdf_bytes)
                images_to_process.extend(converted_images)
            else:
                images_to_process.append(Image.open(uploaded_file))

    with st.spinner(f"Gemini AI মোট {len(images_to_process)} টি পেজ প্রসেস করছে..."):
        for index, img in enumerate(images_to_process):
            # API Rate Limit (ResourceExhausted) এড়াতে ২ সেকেন্ড বিরতি
            if index > 0:
                time.sleep(2)
                
            response = model.generate_content([SYSTEM_PROMPT, img])
            cleaned_page_text = clean_bengali_symbols(response.text)
            
            combined_result += f"\n\n--- পৃষ্ঠা {index + 1} ---\n\n"
            combined_result += cleaned_page_text

    st.success("✅ সফলভাবে সব ফাইল কনভার্ট সম্পন্ন হয়েছে!")

    st.subheader("📝 কনভার্ট হওয়া টেক্সট প্রিভিউ:")
    st.text_area("আউটপুট টেক্সট:", combined_result, height=300)

    col1, col2 = st.columns(2)
    
    word_file = create_word_docx(combined_result, bangla_font_name, english_font_name, arabic_font_name, font_size)
    pdf_file = create_pdf(combined_result, font_size)

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

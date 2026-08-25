import streamlit as st
import vertexai
from vertexai.generative_models import GenerativeModel, Part
from docx import Document
import io
from PIL import Image

# Vertex AI ইনিশিয়ালাইজেশন
vertexai.init(project="school-shet-ocr", location="us-central1")
model = GenerativeModel("gemini-2.5-flash")

st.set_page_config(page_title="Advanced School Sheet OCR", layout="wide")
st.title("📚 School Sheet Handwritten to Word & PDF Agent")
st.write("বাংলা, ইংরেজি, আরবি ও গণিতের হাতে লেখা শিটের ছবি কনভার্ট ও এডিট করুন।")

# সেশন স্টেট ইনিশিয়ালাইজ করা (ডাটা ধরে রাখার জন্য)
if "extracted_texts" not in st.session_state:
    st.session_state.extracted_texts = []

uploaded_files = st.file_uploader(
    "আপনার শিটের ছবিগুলো আপলোড করুন (সর্বোচ্চ ৫০টি)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.error("⚠️ একসাথে সর্বোচ্চ ৫০টি ছবি আপলোড করা যাবে।")
    else:
        st.success(f"মোট {len(uploaded_files)}টি ছবি সিলেক্ট করা হয়েছে।")
        
        if st.button("🚀 Process & Extract Text"):
            st.session_state.extracted_texts = [] # নতুন প্রসেসের জন্য খালি করা
            progress_bar = st.progress(0)
            status_text = st.empty()

            prompt = """
            You are an expert OCR agent proficient in Bengali, English, Arabic, and Mathematics. 
            Extract all text, math formulas, equations, symbols, and diagrams from this handwritten school sheet with high accuracy.

            Instructions:
            1. Preserve original layout and sequence.
            2. For Arabic text, maintain right-to-left order, vowels/harakat.
            3. For Mathematics, express equations in clear standard readable format or LaTeX math notation.
            4. For Diagrams, describe them and transcribe labels inside.
            """

            for i, file in enumerate(uploaded_files):
                status_text.text(f"প্রসেস করা হচ্ছে ({i+1}/{len(uploaded_files)}): {file.name}...")
                image_bytes = file.getvalue()
                image_part = Part.from_data(data=image_bytes, mime_type=file.type)
                
                response = model.generate_content([prompt, image_part])
                
                # সেশন স্টেটে টেক্সট সেভ করা
                st.session_state.extracted_texts.append({
                    "file_name": file.name,
                    "text": response.text
                })
                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.success("🎉 সফলভাবে প্রসেস করা হয়েছে! নিচে আপনি এটি রিভিউ ও এডিট করতে পারেন।")

# যদি ডাটা প্রসেসড থাকে, তবে এডিটিং ও ডাউনলোডের অপশন দেখাবে
if st.session_state.extracted_texts:
    st.markdown("---")
    st.subheader("✏️ রিভিউ এবং প্রয়োজনীয় পরিবর্তন করুন (Editable Output)")

    updated_results = []
    
    # প্রতিটি পেজের জন্য আলাদা এডিটর বক্স
    for idx, item in enumerate(st.session_state.extracted_texts):
        st.write(f"**Page {idx+1}: {item['file_name']}**")
        edited_text = st.text_area(
            label=f"এডিট করুন (Page {idx+1})",
            value=item["text"],
            height=200,
            key=f"editor_{idx}"
        )
        updated_results.append({"file_name": item["file_name"], "text": edited_text})

    st.markdown("---")
    st.subheader("📥 ফাইল ডাউনলোড করুন")

    # ১. ওয়ার্ড ফাইল তৈরি
    doc = Document()
    for idx, item in enumerate(updated_results):
        doc.add_heading(f"Page {idx+1} - {item['file_name']}", level=2)
        doc.add_paragraph(item["text"])
        doc.add_page_break()

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)

    st.download_button(
        label="📄 Download Word File (.docx)",
        data=doc_io,
        file_name="Converted_School_Sheets.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
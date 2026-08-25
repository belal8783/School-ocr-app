import os
import google.generativeai as genai
from docx import Document
from io import BytesIO
import PIL.Image
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Streamlit secrets থেকে API Key পড়া
api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))

if api_key:
    genai.configure(api_key=api_key)

st.set_page_config(
    page_title="School Sheet Handwritten to Word & PDF Agent", layout="wide"
)

st.title("Handwritten to Word & PDF Agent")
st.write("বাংলা, ইংরেজি, আরবি ও গণিতের হাতে লেখা শিটের ছবি কনভার্ট ও এডিট করুন।")

# ফাইল আপলোড বক্স
uploaded_files = st.file_uploader(
    "আপনার শিটের ছবিগুলো একসাথে সিলেক্ট করুন (সর্বোচ্চ ৫০টি)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.warning("আপনি ৫০টির বেশি ফাইল সিলেক্ট করেছেন! প্রথম ৫০টি প্রসেস করা হবে।")
        uploaded_files = uploaded_files[:50]

    # নতুন ফিচার: ইউজার নিজের মতো ফাইলের নাম দিতে পারবে
    custom_filename = st.text_input(
        "ডাউনলোড করার ফাইলের নাম লিখুন (এক্সটেনশন ছাড়া):",
        value="Converted_School_Sheets"
    )

    if st.button("কনভার্ট করা শুরু করুন"):
        if not api_key:
            st.error("Gemini API Key পাওয়া যায়নি!")
        else:
            with st.spinner("ছবিগুলো প্রসেস করা হচ্ছে... অনুগ্রহ করে অপেক্ষা করুন"):
                try:
                    model = genai.GenerativeModel("gemini-3.6-flash")

                    full_text = ""
                    for idx, uploaded_file in enumerate(uploaded_files):
                        st.write(f"প্রসেস করা হচ্ছে ({idx+1}/{len(uploaded_files)}): {uploaded_file.name}")
                        image = PIL.Image.open(uploaded_file)

                        prompt = """
                        You are an expert OCR model specializing in converting school sheets, manuscripts, and test papers.
                        Extract all text accurately from the provided image.
                        Preserve languages: Bengali, English, Arabic, and Mathematical formulas.
                        Format tables, headers, and bullet points properly.
                        """

                        response = model.generate_content([prompt, image])
                        full_text += f"\n--- পৃষ্ঠা {idx + 1}: {uploaded_file.name} ---\n\n"
                        full_text += response.text + "\n\n"

                    st.success("সবগুলো ফাইলের টেক্সট সফলভাবে এক্সট্র্যাক্ট করা হয়েছে!")
                    edited_text = st.text_area(
                        "এক্সট্র্যাক্ট করা টেক্সট (প্রয়োজনে এডিট করুন):",
                        full_text,
                        height=350,
                    )

                    # ফাইনাল ফাইল নেম নির্ধারণ (ইউজার নাম না দিলে ডিফল্ট নাম থাকবে)
                    safe_filename = "".join(c for c in custom_filename if c.isalnum() or c in (' ', '_', '-')).strip()
                    if not safe_filename:
                        safe_filename = "Converted_School_Sheets"

                    col1, col2 = st.columns(2)

                    # 1. Word File Generation
                    doc = Document()
                    doc.add_heading("School Sheet Text Output", 0)
                    for para in edited_text.split("\n"):
                        if para.strip():
                            doc.add_paragraph(para)

                    doc_io = BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)

                    with col1:
                        st.download_button(
                            label="📄 Word (.docx) ডাউনলোড করুন",
                            data=doc_io,
                            file_name=f"{safe_filename}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )

                    # 2. PDF Generation
                    pdf_io = BytesIO()
                    c = canvas.Canvas(pdf_io, pagesize=letter)
                    textobject = c.beginText()
                    textobject.setTextOrigin(50, 750)
                    textobject.setFont("Helvetica", 10)

                    for line in edited_text.split("\n"):
                        textobject.textLine(line)
                    c.drawText(textobject)
                    c.showPage()
                    c.save()
                    pdf_io.seek(0)

                    with col2:
                        st.download_button(
                            label="📕 PDF (.pdf) ডাউনলোড করুন",
                            data=pdf_io,
                            file_name=f"{safe_filename}.pdf",
                            mime="application/pdf",
                        )

                except Exception as e:
                    st.error(f"একটি ত্রুটি ঘটেছে: {e}")

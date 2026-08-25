import os
import google.generativeai as genai
from docx import Document
from io import BytesIO
import PIL.Image
import streamlit as st

# Streamlit secrets থেকে API Key পড়া
api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY"))

if api_key:
    genai.configure(api_key=api_key)

st.set_page_config(
    page_title="School Sheet Handwritten to Word & PDF Agent", layout="wide"
)

st.title("Handwritten to Word & PDF Agent")
st.write("বাংলা, ইংরেজি, আরবি ও গণিতের হাতে লেখা শিটের ছবি কনভার্ট ও এডিট করুন।")

# একসাথে সর্বোচ্চ ৫০টি ফাইল আপলোডের সুবিধা
uploaded_files = st.file_uploader(
    "আপনার শিটের ছবিগুলো একসাথে আপলোড করুন (সর্বোচ্চ ৫০টি)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.warning("আপনি ৫০টির বেশি ফাইল সিলেক্ট করেছেন! প্রথম ৫০টি ফাইল প্রসেস করা হবে।")
        uploaded_files = uploaded_files[:50]

    if st.button("কনভার্ট করা শুরু করুন"):
        if not api_key:
            st.error(
                "Gemini API Key পাওয়া যায়নি! অনুগ্রহ করে Streamlit Secrets-এ"
                " GEMINI_API_KEY সেট করুন।"
            )
        else:
            with st.spinner("ছবিগুলো প্রসেস করা হচ্ছে... অনুগ্রহ করে অপেক্ষা করুন"):
                try:
                    # সঠিক ও আপডেট মডেল নাম ব্যবহার করা হলো
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
                        full_text += (
                            f"\n--- পৃষ্ঠা {idx + 1}: {uploaded_file.name} ---\n\n"
                        )
                        full_text += response.text + "\n\n"

                    st.success("সবগুলো ফাইলের টেক্সট সফলভাবে এক্সট্র্যাক্ট করা হয়েছে!")
                    edited_text = st.text_area(
                        "এক্সট্র্যাক্ট করা টেক্সট (প্রয়োজনে এডিট করুন):",
                        full_text,
                        height=400,
                    )

                    # Word File Generation
                    doc = Document()
                    doc.add_heading("School Sheet Text Output", 0)
                    for para in edited_text.split("\n"):
                        if para.strip():
                            doc.add_paragraph(para)

                    bio = BytesIO()
                    doc.save(bio)
                    bio.seek(0)

                    st.download_button(
                        label="Word (.docx) হিসেবে ডাউনলোড করুন",
                        data=bio,
                        file_name="Converted_School_Sheets.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

                except Exception as e:
                    st.error(f"একটি ত্রুটি ঘটেছে: {e}")

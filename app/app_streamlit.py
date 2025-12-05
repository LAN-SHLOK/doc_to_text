# app_streamlit.py
import streamlit as st
import json
from pathlib import Path
import tempfile
from core.doc_to_text import convert_any_to_text, extract_important_details

st.set_page_config(page_title="Doc → TXT Converter", layout="wide")

st.title("Convert Document → TXT (Tesseract + EasyOCR)")
st.markdown("Supports: PDF (text/scanned), DOCX, PPTX, images. Use clean scans for best OCR.")

uploaded = st.file_uploader("Upload a document", type=["pdf", "docx", "pptx", "jpg", "jpeg", "png", "tiff", "bmp", "webp"])
spell_corr = st.checkbox("Enable light spell-correction", value=False)
if uploaded:
    suffix = uploaded.name.split('.')[-1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{suffix}") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    st.info("File saved. Starting conversion...")
    try:
        txt_path, confidence, text = convert_any_to_text(tmp_path, do_spell_correct=spell_corr)
        st.success(f"Saved TXT: {txt_path} — Confidence: {confidence*100:.2f}%")
        st.download_button("Download TXT", data=text, file_name=Path(txt_path).name, mime="text/plain")

        # Important details
        details = extract_important_details(text)
        st.subheader("Important details (heuristic)")
        st.json(details)
        st.download_button("Download details JSON", data=json.dumps(details, indent=2), file_name="important_details.json", mime="application/json")

        st.subheader("Extracted text preview")
        st.text_area("Text preview", value=text[:20000], height=400)

        if confidence < 0.5:
            st.warning("Low OCR confidence. Consider re-uploading a cleaner or higher-DPI scan, or enabling spell-correction.")
    except Exception as e:
        st.error(f"Conversion failed: {e}")

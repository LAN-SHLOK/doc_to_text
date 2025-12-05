# app_streamlit.py

# ====== PATH FIX: ensure "core" package is importable ======
import sys, os
from pathlib import Path

def ensure_core_on_path(max_up=8):
    this_file = Path(__file__).resolve()
    # walk up parents and add the first ancestor that contains "core" dir
    for i in range(0, min(max_up, len(this_file.parents))):
        candidate = this_file.parents[i]
        if (candidate / "core").is_dir():
            cand_str = str(candidate)
            if cand_str not in sys.path:
                sys.path.insert(0, cand_str)
            return True
    # fallback: add repo root guess (two levels up) if nothing found
    fallback = str(this_file.parents[min(2, len(this_file.parents)-1)])
    if fallback not in sys.path:
        sys.path.insert(0, fallback)
    return False

_found = ensure_core_on_path()
# Optional debug line (uncomment while debugging)
# print(f"ensure_core_on_path -> {_found}; sys.path[0:5]={sys.path[:5]}")
# ===========================================================

# now safe to import app deps and core
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

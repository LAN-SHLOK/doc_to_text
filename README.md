# Document Converter

Convert PDFs (scanned or text), DOCX, PPTX and images to plain .txt using Tesseract + EasyOCR.

## Quick start

1. Install system deps:
   - Tesseract OCR (make sure `tesseract` is in PATH)
   - Poppler (for pdf2image) — on Windows set POPPLER_PATH or add to PATH.

2. Create virtualenv and install:
   ```
   pip install -r requirements.txt
   ```

3. Run Streamlit app:
   ```
   streamlit run app/app_streamlit.py
   ```

Files:
- `core/doc_to_text.py` - main converter engine
- `app/app_streamlit.py` - Streamlit UI wrapper

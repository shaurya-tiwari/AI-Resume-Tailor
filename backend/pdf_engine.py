import fitz  # PyMuPDF

def extract_text(pdf_file):
    """Reads PDF file stream and extracts plain text."""
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    return "".join(page.get_text("text") for page in doc)
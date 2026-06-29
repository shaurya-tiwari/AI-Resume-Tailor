from fpdf import FPDF

def generate_pdf(text_content):
    # 1. Initialize PDF
    pdf = FPDF()
    pdf.add_page()
    
    # Auto page break on rakha hai taaki lamba resume next page par automatic chala jaye
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("helvetica", size=11)
    
    # 2. Encoding Fix 
    # AI ke output mein special characters ya bullets hote hain. 
    # Standard PDF fonts crash na hon, isliye text ko safe format ('latin-1') mein clean kar rahe hain.
    clean_text = text_content.encode('latin-1', 'ignore').decode('latin-1')
    
    # 3. Add Text to PDF
    pdf.multi_cell(0, 6, text=clean_text)
    
    # 4. Return as Bytes (Streamlit download_button ko yahi format chahiye hota hai)
    return bytes(pdf.output())
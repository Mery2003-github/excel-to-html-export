import pdfplumber

def extraire_texte(pdf_path):
    texte = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texte += page.extract_text() + "\n"
    return texte

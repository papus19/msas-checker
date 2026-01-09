import streamlit as st
from docx import Document
from docx.shared import Mm

def check_msas_compliance(doc):
    reports = []
    
    # Vérification des Marges [cite: 23]
    section = doc.sections[0]
    margins = {
        "Haut": section.top_margin.mm,
        "Bas": section.bottom_margin.mm,
        "Gauche": section.left_margin.mm,
        "Droite": section.right_margin.mm
    }
    for name, value in margins.items():
        if abs(value - 25) > 1:
            reports.append(f"❌ Marge {name}: {value:.1f}mm détectée. Le guide exige 25mm partout[cite: 23].")
    
    # Vérification de la police du titre [cite: 8, 36, 37]
    if doc.paragraphs:
        first_para = doc.paragraphs[0]
        if first_para.runs:
            font = first_para.runs[0].font
            if font.size and font.size.pt != 20:
                reports.append(f"❌ Titre: Taille de {font.size.pt}pt détectée au lieu de 20pt[cite: 36].")
            if not font.bold:
                reports.append("❌ Titre: Doit être en gras[cite: 8].")

    # Vérification de la casse des sections [cite: 41]
    for para in doc.paragraphs:
        # On cible les titres courts en gras (souvent des sections)
        if para.text.isupper() == False and any(run.bold for run in para.runs) and len(para.text) < 50:
             if para.text in ["INTRODUCTION", "CONCLUSION", "REFERENCES", "REMERCIEMENTS"]:
                 continue # Déjà correct
             else:
                 reports.append(f"⚠️ Section '{para.text[:20]}...': Les titres de sections doivent être en MAJUSCULES[cite: 41].")

    return reports

# --- Interface Streamlit ---
st.set_page_config(page_title="MSAS Compliance Checker", layout="centered")
st.title("Vérificateur de Template MSAS")
st.write("Téléchargez votre article pour vérifier sa conformité au format de la conférence.")

uploaded_file = st.file_uploader("Choisir un fichier .docx", type="docx")

if uploaded_file is not None:
    doc = Document(uploaded_file)
    with st.spinner('Analyse en cours...'):
        errors = check_msas_compliance(doc)
        
        if not errors:
            st.success("Félicitations ! Votre document semble respecter les règles principales du template MSAS.")
        else:
            st.subheader("Points à corriger :")
            for error in errors:
                st.error(error)
    
    st.info("Note : Les articles sont limités à 10 pages maximum[cite: 33].")

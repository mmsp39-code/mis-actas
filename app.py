import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
import os
from PIL import Image, ImageOps, ImageEnhance
import numpy as np
import re

st.set_page_config(page_title="Generador de Actas Profesional", layout="wide")

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'], gpu=False)
reader = load_ocr()

def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

# --- SIDEBAR ---
with st.sidebar:
    st.header("📋 Datos del Acta")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR 1")
    localidad = st.text_input("Ubicación", "Plastitis de Argentina")
    idcoste = st.text_input("IDCOSTE", "IDCOSTE")
    instaladores = st.text_input("Instaladores", "Instaladores")
    fecha = st.text_input("Fecha de Instalación", "17/02/2022")

st.title("📄 Generador de Actas - ADASA & INELCOM")
excel_file = st.file_uploader("1. Sube el Excel", type=['xlsx'])

st.subheader("📸 2. Carga de Fotos")
c1, c2 = st.columns(2)
with c1:
    f_equipos = st.file_uploader("📥 Equipamiento (Con S/N)", accept_multiple_files=True)
with c2:
    f_cartel = st.file_uploader("🪧 Cartel Informativo", accept_multiple_files=True)

if st.button("🚀 GENERAR WORD"):
    if excel_file:
        with st.spinner("Generando documento..."):
            df = pd.read_excel(excel_file)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # --- PORTADA Y CABECERA ---
            if os.path.exists('logo_instituciona.png'):
                p = doc.add_paragraph()
                p.alignment = 1
                p.add_run().add_picture('logo_instituciona.png', width=Inches(4.5))

            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = 1
            doc.add_heading('IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO', 2).alignment = 1

            # Tabla de Datos (Como tu captura)
            tbl = doc.add_table(rows=5, cols=2)
            tbl.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("Ubicación", localidad), ("IDCOSTC", idcoste), ("Instaladores", instaladores), ("Fecha de Instalación", fecha)]
            for i, (k, v) in enumerate(datos):
                tbl.rows[i].cells[0].text, tbl.rows[i].cells[1].text = k, str(v)

            # --- SECCIÓN EQUIPOS (CON S/N) ---
            if f_equipos:
                doc.add_paragraph("")
                # Usamos una tabla invisible para poner las fotos de 2 en 2
                grid = doc.add_table(rows=0, cols=2)
                for i, foto in enumerate(f_equipos):
                    if i % 2 == 0: row_cells = grid.add_row().cells
                    cell = row_cells[i % 2]
                    
                    # IA para detectar S/N (Optimizado para tus fotos)
                    img = ImageOps.grayscale(Image.open(foto))
                    txt = " ".join(reader.readtext(np.array(img), detail=0)).upper()
                    m = re.search(r'(SN-[A-Z0-9-]+|\d{4}[-_]\d{4}[-_]\d{2})', txt)
                    sn = m.group(0).replace("_", "-") if m else "No detectado"
                    
                    cell.paragraphs[0].add_run().add_picture(foto, width=Inches(2.5))
                    cell.add_paragraph(f"S/N: {sn}").alignment = 1

            # --- SECCIÓN CARTEL (SIN S/N) ---
            if f_cartel:
                doc.add_heading('FOTOS DE CARTEL INFORMATIVO', level=1)
                for foto in f_cartel:
                    p = doc.add_paragraph()
                    p.alignment = 1
                    p.add_run().add_picture(foto, width=Inches(4))
                    # TEXTO FIJO PARA CARTEL
                    t_obs = doc.add_paragraph()
                    t_obs.alignment = 1
                    run = t_obs.add_run("Observaciones: Fotografía del cartel de subvenciones de fondos europeos.")
                    run.bold = True
                    run.font.size = Pt(10)

            # --- LOGOS FINALES (ADASA E INELCOM) ---
            doc.add_paragraph("")
            p_final = doc.add_paragraph()
            p_final.alignment = 1
            if os.path.exists('logo_adasa.png'):
                p_final.add_run().add_picture('logo_adasa.png', width=Inches(1.5))
            p_final.add_run("    ")
            if os.path.exists('logo_inelcom.png'):
                p_final.add_run().add_picture('logo_inelcom.png', width=Inches(1.5))

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta generada!")
            st.download_button("📥 DESCARGAR WORD", target.getvalue(), f"Acta_{nombre_edar}.docx")

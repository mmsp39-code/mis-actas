
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
import os
from PIL import Image, ImageOps
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

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("📋 Datos del Acta")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR 1")
    localidad = st.text_input("Ubicación", "Plastitis de Argentina")
    idcoste = st.text_input("IDCOSTE", "IDCOSTE")
    instaladores = st.text_input("Instaladores", "Instaladores")
    fecha = st.text_input("Fecha de Instalación", "17/02/2022")

st.title("📄 Generador de Actas - ADASA & INELCOM")
excel_file = st.file_uploader("1. Sube el Excel de Coordenadas", type=['xlsx'])

st.divider()

# --- REUPERAMOS TODOS LOS APARTADOS ---
st.subheader("📸 2. Carga de Fotos por Sección")
col1, col2 = st.columns(2)

with col1:
    f_portada = st.file_uploader("🖼️ Foto Portada (Puerta/Entorno)", accept_multiple_files=True)
    f_cartel = st.file_uploader("🪧 Cartel Informativo", accept_multiple_files=True)
    f_entrada = st.file_uploader("📥 Entrada / Equipamiento", accept_multiple_files=True)
with col2:
    f_alivio = st.file_uploader("🌊 Alivio (EDAR/Colector)", accept_multiple_files=True)
    f_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    f_graficas = st.file_uploader("📈 Gráficas / Pantallas / Conclusiones", accept_multiple_files=True)

if st.button("🚀 GENERAR ACTA PROFESIONAL"):
    if excel_file:
        with st.spinner("Procesando todas las secciones..."):
            df = pd.read_excel(excel_file)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # --- PORTADA ---
            if os.path.exists('logo_instituciona.png'):
                p = doc.add_paragraph()
                p.alignment = 1
                p.add_run().add_picture('logo_instituciona.png', width=Inches(6))

            doc.add_heading(nombre_edar, 0).alignment = 1
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = 1

            if f_portada:
                p = doc.add_paragraph()
                p.alignment = 1
                p.add_run().add_picture(f_portada[0], width=Inches(4.5))

            tbl = doc.add_table(rows=5, cols=2)
            tbl.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("Ubicación", localidad), ("IDCOSTE", idcoste), ("Instaladores", instaladores), ("Fecha de Instalación", fecha)]
            for i, (k, v) in enumerate(datos):
                tbl.rows[i].cells[0].text, tbl.rows[i].cells[1].text = k, str(v)

            # --- SECCIONES CONFIGURADAS ---
            secciones = [
                ("CARTEL INFORMATIVO", f_cartel, "fijo"),
                ("EQUIPAMIENTO EN ENTRADA", f_entrada, "ia"),
                ("EQUIPAMIENTO EN ALIVIO", f_alivio, "ia"),
                ("EQUIPAMIENTO EN SALIDA", f_salida, "ia"),
                ("GRÁFICAS Y CONCLUSIONES", f_graficas, "fijo")
            ]

            for titulo, lista, modo in secciones:
                if lista:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2) # 2 fotos por fila
                    
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: row_cells = grid.add_row().cells
                        cell = row_cells[i % 2]
                        
                        sn_txt = ""
                        if modo == "ia":
                            # IA para S/N
                            img = ImageOps.grayscale(Image.open(foto))
                            res = " ".join(reader.readtext(np.array(img), detail=0)).upper()
                            m = re.search(r'(SN-[A-Z0-9-]+|\d{4}[-_]\d{4}[-_]\d{2})', res)
                            sn = m.group(0).replace("_", "-") if m else "No detectado"
                            sn_txt = f"S/N: {sn}"
                        elif titulo == "CARTEL INFORMATIVO":
                            sn_txt = "Observaciones: Fotografía del cartel de subvenciones de fondos europeos."
                        
                        cell.paragraphs[0].add_run().add_picture(foto, width=Inches(2.5))
                        if sn_txt:
                            cell.add_paragraph(sn_txt).alignment = 1

            # --- LOGOS FINALES ---
            doc.add_paragraph("")
            p_final = doc.add_paragraph()
            p_final.alignment = 1
            try:
                if os.path.exists('logo_adasa.png'):
                    p_final.add_run().add_picture('logo_adasa.png', width=Inches(1.5))
                p_final.add_run("    ")
                if os.path.exists('logo_inelcom.png'):
                    p_final.add_run().add_picture('logo_inelcom.png', width=Inches(1.5))
            except: pass

            # FIRMA
            doc.add_paragraph("\n\n")
            tbl_f = doc.add_table(rows=2, cols=1)
            tbl_f.style = 'Table Grid'
            tbl_f.rows[0].cells[0].text = "FIRMA Y VALIDACIÓN:"
            tbl_f.rows[1].height = Inches(1.5)

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta completa generada!")
            st.download_button("📥 DESCARGAR WORD", target.getvalue(), f"Acta_{nombre_edar}.docx")
    else:
        st.error("Por favor, sube el Excel primero.")

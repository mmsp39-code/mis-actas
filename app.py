import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
import os
from PIL import Image, ImageOps, ImageEnhance
import numpy as np
import re

# Configuración inicial
st.set_page_config(page_title="Generador de Actas Pro", layout="wide")

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
    st.header("📋 Datos de la EDAR")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR AIN")
    localidad = st.text_input("Localidad", "AIN (Valencia)")
    provincia = st.text_input("Provincia", "Valencia")
    idcoste = st.text_input("IDCOSTE", "0017")
    instaladores = st.text_input("Instaladores", "JOSE / PEPE")
    responsable = st.text_input("Responsable Explotación", "Carlos")
    fecha = st.date_input("Fecha Instalación")

# --- CUERPO PRINCIPAL ---
st.title("📄 Generador de Actas Profesional")
excel_file = st.file_uploader("1. Sube el Excel", type=['xlsx'])

st.subheader("📸 2. Carga de Fotos")
c1, c2 = st.columns(2)
with c1:
    f_puerta = st.file_uploader("🖼️ Puerta", accept_multiple_files=True)
    f_cartel = st.file_uploader("🪧 Cartel", accept_multiple_files=True)
    f_entrada = st.file_uploader("📥 Entrada", accept_multiple_files=True)
with c2:
    f_alivio = st.file_uploader("🌊 Alivio", accept_multiple_files=True)
    f_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    f_pantallas = st.file_uploader("📈 Pantallas", accept_multiple_files=True)

if st.button("🚀 GENERAR ACTA"):
    if excel_file:
        with st.spinner("Procesando fotos y generando Word..."):
            df = pd.read_excel(excel_file)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # --- PORTADA ---
            if os.path.exists('logo_instituciona.png'):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture('logo_instituciona.png', width=Inches(6))

            doc.add_heading(nombre_edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            if f_puerta:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(f_puerta[0], width=Inches(4.5))

            tbl = doc.add_table(rows=6, cols=2)
            tbl.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("LOCALIDAD", localidad), ("IDCOSTE", idcoste), ("INSTALADORES", instaladores), ("RESPONSABLE", responsable), ("FECHA", str(fecha))]
            for i, (k, v) in enumerate(datos):
                tbl.rows[i].cells[0].text = k
                tbl.rows[i].cells[1].text = str(v)

            # Logos inferiores
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                if os.path.exists('logo_adasa.png'): p.add_run().add_picture('logo_adasa.png', width=Inches(1.2))
                p.add_run("    ")
                if os.path.exists('logo_inelcom.png'): p.add_run().add_picture('logo_inelcom.png', width=Inches(1.2))
            except: pass

            # --- TABLA TÉCNICA ---
            doc.add_page_break()
            doc.add_heading('IDENTIFICACIÓN DEL EQUIPAMIENTO', level=1)
            tbl_e = doc.add_table(rows=1, cols=3)
            tbl_e.style = 'Table Grid'
            hdr = tbl_e.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            secciones = [("CARTEL", f_cartel), ("ENTRADA", f_entrada), ("ALIVIO", f_alivio), ("SALIDA", f_salida), ("PANTALLAS", f_pantallas)]
            
            for titulo, lista in secciones:
                if lista:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    # Usamos una tabla para que quepan 2 fotos por fila
                    grid = doc.add_table(rows=0, cols=2)
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: row_cells = grid.add_row().cells
                        cell = row_cells[i % 2]
                        
                        sn, coor = "No detectado", "N/A"
                        if titulo != "CARTEL" and titulo != "PANTALLAS":
                            img = Image.open(foto)
                            # OCR rápido
                            res = " ".join(reader.readtext(np.array(img), detail=0)).upper()
                            m = re.search(r'(\d{4}[-_]\d{4}[-_]\d{2}|SN-[A-Z0-9-]+)', res)
                            if m:
                                sn = m.group(0).replace("_", "-")
                                if c_coord and i < len(df): coor = str(df.iloc[i][c_coord])
                                r_cells = tbl_e.add_row().cells
                                r_cells[0].text, r_cells[1].text, r_cells[2].text = titulo, sn, coor
                        
                        cell.paragraphs[0].add_run().add_picture(foto, width=Inches(2.5))
                        cell.add_paragraph(f"{titulo} S/N: {sn}").alignment = WD_ALIGN_PARAGRAPH.CENTER

            # --- CIERRE ---
            doc.add_page_break()
            doc.add_heading('FIRMA Y VALIDACIÓN', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
            tbl_f = doc.add_table(rows=2, cols=1)
            tbl_f.style = 'Table Grid'
            tbl_f.rows[0].cells[0].text = "FIRMA:"
            tbl_f.rows[1].height = Inches(2)

            # --- GUARDADO ---
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta generada con éxito!")
            st.download_button(label="📥 DESCARGAR EL ARCHIVO WORD", data=target.getvalue(), file_name=f"Acta_{nombre_edar}.docx")
    else:
        st.error("Por favor, sube el archivo Excel primero.")

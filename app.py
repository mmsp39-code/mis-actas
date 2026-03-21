import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
import os
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
import numpy as np
import re

st.set_page_config(page_title="Generador Actas Pro - IA Vision", layout="wide")

@st.cache_resource
def load_ocr():
    # Cargamos el lector optimizado para números
    return easyocr.Reader(['es'], gpu=False)
reader = load_ocr()

def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

def vision_artificial_contraste(pil_image):
    """Transforma la foto para que el grabado del metal sea legible"""
    # 1. Escala de grises
    img = ImageOps.grayscale(pil_image)
    # 2. Aumentar contraste al máximo para ver el grabado
    img = ImageEnhance.Contrast(img).enhance(3.5)
    # 3. Filtro de nitidez para definir los bordes de los números
    img = img.filter(ImageFilter.SHARPEN)
    return img

# --- DATOS EDAR ---
with st.sidebar:
    st.header("📋 Datos EDAR")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR AIN")
    localidad = st.text_input("Localidad", "AIN (Valencia)")
    idcoste = st.text_input("IDCOSTE", "0017")
    instaladores = st.text_input("Instaladores", "JOSE / PEPE")
    responsable = st.text_input("Responsable Explotación", "Nombre")
    fecha = st.date_input("Fecha Instalación")

st.title("📄 Generador de Actas con Visión Artificial")
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
        with st.spinner("La IA está analizando las chapas de los equipos..."):
            df = pd.read_excel(excel_file)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # --- PORTADA ---
            if os.path.exists('logo_instituciona.png'):
                doc.add_paragraph().alignment = 1
                doc.paragraphs[-1].add_run().add_picture('logo_instituciona.png', width=Inches(6))

            doc.add_heading(nombre_edar, 0).alignment = 1
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = 1

            if f_puerta:
                p = doc.add_paragraph()
                p.alignment = 1
                p.add_run().add_picture(f_puerta[0], width=Inches(5))

            tbl = doc.add_table(rows=6, cols=2)
            tbl.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("LOCALIDAD", localidad), ("IDCOSTE", idcoste), ("INSTALADORES", instaladores), ("RESPONSABLE", responsable), ("FECHA", str(fecha))]
            for i, (k, v) in enumerate(datos):
                tbl.rows[i].cells[0].text, tbl.rows[i].cells[1].text = k, str(v)

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
                    grid = doc.add_table(rows=0, cols=2)
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: row_cells = grid.add_row().cells
                        cell = row_cells[i % 2]
                        
                        sn, coor = "No detectado", "N/A"
                        
                        if titulo == "CARTEL":
                            obs = "Observaciones: Fotografía del cartel de subvenciones de fondos europeos."
                        else:
                            # --- PROCESO DE VISIÓN ARTIFICIAL ---
                            img_raw = Image.open(foto)
                            img_mejorada = vision_artificial_contraste(img_raw)
                            
                            # Leer texto
                            txt_ia = " ".join(reader.readtext(np.array(img_mejorada), detail=0)).upper()
                            
                            # Buscar S/N con patrones específicos
                            m = re.search(r'(\d{4}[-_]\d{4}[-_]\d{2}|SN-[A-Z0-9-]+|ITC\d+|[A-Z0-9]{4}-[A-Z0-9]{4})', txt_ia)
                            if m:
                                sn = m.group(0).replace("_", "-")
                                if c_coord and i < len(df): coor = str(df.iloc[i][c_coord])
                                r = tbl_e.add_row().cells
                                r[0].text, r[1].text, r[2].text = titulo, sn, coor
                            
                            obs = f"{titulo} S/N: {sn}"

                        cell.paragraphs[0].add_run().add_picture(foto, width=Inches(3.0))
                        cell.add_paragraph(obs).alignment = 1

            # --- FIRMA ---
            doc.add_page_break()
            doc.add_heading('FIRMA Y VALIDACIÓN', level=1).alignment = 1
            tbl_f = doc.add_table(rows=2, cols=1)
            tbl_f.style = 'Table Grid'
            tbl_f.rows[0].cells[0].text = "FIRMA:"
            tbl_f.rows[1].height = Inches(2)

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta generada con Visión Artificial")
            st.download_button("📥 DESCARGAR WORD", target.getvalue(), f"Acta_{nombre_edar}.docx")

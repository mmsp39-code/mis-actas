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

st.set_page_config(page_title="Generador Actas Pro", layout="wide")

# Lector OCR más ligero
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'], gpu=False)
reader = load_ocr()

def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

def optimizar_imagen_rapido(pil_image):
    """Mejora contraste de forma ligera para no colgar la web"""
    # Convertir a escala de grises y mejorar contraste simple
    bw = ImageOps.grayscale(pil_image)
    enhancer = ImageEnhance.Contrast(bw)
    return enhancer.enhance(2.0)

# --- SIDEBAR ---
with st.sidebar:
    st.header("📋 Datos EDAR")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR ALZIRA")
    localidad = st.text_input("Localidad", "Alzira")
    provincia = st.text_input("Provincia", "Valencia")
    idcoste = st.text_input("IDCOSTE", "0017")
    instaladores = st.text_input("Instaladores", "Técnico 1")
    responsable = st.text_input("Responsable Explotación", "Nombre")
    fecha = st.date_input("Fecha Instalación")

st.title("📄 Generador de Actas Profesional")
excel_file = st.file_uploader("1. Sube el Excel", type=['xlsx'])

st.divider()
st.subheader("📸 2. Carga de Fotos")
col1, col2 = st.columns(2)
with col1:
    foto_puerta = st.file_uploader("🖼️ Puerta", accept_multiple_files=True)
    foto_cartel = st.file_uploader("🪧 Cartel", accept_multiple_files=True)
    fotos_entrada = st.file_uploader("📥 Entrada", accept_multiple_files=True)
with col2:
    fotos_alivio = st.file_uploader("🌊 Alivio", accept_multiple_files=True)
    fotos_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    fotos_graficas = st.file_uploader("📈 Pantallas", accept_multiple_files=True)

if st.button("🚀 GENERAR ACTA"):
    if excel_file:
        progress_bar = st.progress(0)
        with st.spinner("Procesando..."):
            df = pd.read_excel(excel_file)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # --- PORTADA ---
            p_logo = doc.add_paragraph()
            p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if os.path.exists('logo_instituciona.png'):
                p_logo.add_run().add_picture('logo_instituciona.png', width=Inches(5.5))
            
            doc.add_heading(nombre_edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            if foto_puerta:
                p_p = doc.add_paragraph()
                p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_p.add_run().add_picture(foto_puerta[0], width=Inches(4.5))

            tbl = doc.add_table(rows=6, cols=2)
            tbl.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("LOCALIDAD", f"{localidad} ({provincia})"), ("IDCOSTE", idcoste), ("INSTALADORES", instaladores), ("RESPONSABLE", responsable), ("FECHA", str(fecha))]
            for i, (k, v) in enumerate(datos):
                tbl.rows[i].cells[0].text = k
                tbl.rows[i].cells[1].text = str(v)

            p_l = doc.add_paragraph()
            p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                if os.path.exists('logo_adasa.png'): p_l.add_run().add_picture('logo_adasa.png', width=Inches(1))
                p_l.add_run("    ")
                if os.path.exists('logo_inelcom.png'): p_l.add_run().add_picture('logo_inelcom.png', width=Inches(1))
            except: pass

            # --- TABLA EQUIPOS ---
            doc.add_page_break()
            doc.add_heading('IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO', level=1)
            tbl_e = doc.add_table(rows=1, cols=3)
            tbl_e.style = 'Table Grid'
            hdr = tbl_e.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            secciones = [("FOTO CARTEL", foto_cartel), ("ENTRADA", fotos_entrada), ("ALIVIO", fotos_alivio), ("SALIDA", fotos_salida), ("PANTALLAS", fotos_graficas)]
            
            for titulo, lista in secciones:
                if lista:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2)
                    
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: cells = grid.add_row().cells
                        cell = cells[i % 2]
                        
                        sn, coor = "No detectado", "N/A"
                        
                        if titulo == "FOTO CARTEL":
                            obs = "Observaciones: Fotografía del cartel de subvenciones de fondos europeos."
                        else:
                            # Procesar y Leer
                            img_raw = Image.open(foto)
                            img_proc = optimizar_imagen_rapido(img_raw)
                            txt = " ".join(reader.readtext(np.array(img_proc), detail=0)).upper()
                            
                            # Patrón S/N
                            match = re.search(r'(\d{4}-\d{4}-\d{2}|SN-[A-Z0-9-]+)', txt)
                            if match:
                                sn = match.group(0)
                                if c_coord and i < len(df):
                                    coor = str(df.iloc[i][c_coord])
                                r = tbl_e.add_row().cells
                                r[0].text, r[1].text, r[2].text = titulo, sn, coor
                            
                            obs = f"{titulo} S/N: {sn}"

                        cell.paragraphs[0].

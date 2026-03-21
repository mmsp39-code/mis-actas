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

st.set_page_config(page_title="Generador Actas ADASA-INELCOM Pro", layout="wide")

@st.cache_resource
def load_ocr():
    # Cargamos el lector con soporte para números y español
    return easyocr.Reader(['es'], gpu=False)
reader = load_ocr()

def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

def preprocesar_para_ocr(pil_image):
    """
    Prepara la chapa metálica para que la IA lea mejor:
    Convierte a gris, aumenta contraste y aplica un filtro de nitidez.
    """
    # 1. Convertir a escala de grises
    gray_img = ImageOps.grayscale(pil_image)
    # 2. Aumentar el contraste drásticamente (para separar grabado de metal)
    enhancer = ImageEnhance.Contrast(gray_img)
    contrast_img = enhancer.enhance(2.5)
    # 3. Aumentar nitidez
    sharpness = ImageEnhance.Sharpness(contrast_img)
    final_img = sharpness.enhance(2.0)
    return final_img

# --- DATOS BARRA LATERAL ---
with st.sidebar:
    st.header("📋 Datos de la EDAR")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR ALZIRA")
    localidad = st.text_input("Localidad", "Alzira")
    provincia = st.text_input("Provincia", "Valencia")
    idcoste = st.text_input("IDCOSTE", "0017")
    instaladores = st.text_input("Instaladores", "Técnico 1")
    responsable = st.text_input("Responsable Explotación", "Nombre del Responsable")
    fecha = st.date_input("Fecha Instalación")

# --- INTERFAZ CENTRAL ---
st.title("📄 Generador de Actas - IA Mejorada")
excel_file = st.file_uploader("1. Sube el Excel de Coordenadas", type=['xlsx'])

st.divider()
st.subheader("📸 2. Carga de Fotos")
col1, col2 = st.columns(2)
with col1:
    foto_puerta = st.file_uploader("🖼️ Foto Puerta (Portada)", accept_multiple_files=True)
    foto_cartel = st.file_uploader("🪧 Foto Cartel Informativo", accept_multiple_files=True)
    fotos_entrada = st.file_uploader("📥 Entrada", accept_multiple_files=True)
with col2:
    fotos_alivio = st.file_uploader("🌊 Alivio", accept_multiple_files=True)
    fotos_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    fotos_graficas = st.file_uploader("📈 Pantallas", accept_multiple_files=True)

if st.button("📝 GENERAR ACTA FINAL"):
    if excel_file:
        with st.spinner("La IA está 'limpiando' las fotos para leer los S/N..."):
            df = pd.read_excel(excel_file)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # --- PORTADA ---
            p_eu = doc.add_paragraph()
            p_eu.alignment = WD_ALIGN_PARAGRAPH.CENTER
            logo_eu = 'logo_instituciona.png'
            if os.path.exists(logo_eu):
                p_eu.add_run().add_picture(logo_eu, width=Inches(6.2))
            
            doc.add_heading(nombre_edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            if foto_puerta:
                p_p = doc.add_paragraph()
                p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_p.add_run().add_picture(foto_puerta[0], width=Inches(5.0))

            table_info = doc.add_table(rows=6, cols=2)
            table_info.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("LOCALIDAD", f"{localidad} ({provincia})"), 
                     ("IDCOSTE", idcoste), ("INSTALADORES", instaladores), 
                     ("RESPONSABLE", responsable), ("FECHA", str(fecha))]
            for i, (k, v) in enumerate(datos):
                table_info.rows[i].cells[0].text = k
                table_info.rows[i].cells[1].text = str(v)

            p_logos = doc.add_paragraph()
            p_logos.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                if os.path.exists('logo_adasa.png'): p_logos.add_run().add_picture('logo_adasa.png', width=Inches(1.1))
                p_logos.add_run("    ")
                if os.path.exists('logo_inelcom.png'): p_logos.add_run().add_picture('logo_inelcom.png', width=Inches(1.1))
            except: pass

            # --- TABLA TÉCNICA ---
            doc.add_page_break()
            doc.add_heading('IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO', level=1)
            tbl_equip = doc.add_table(rows=1, cols=3)
            tbl_equip.style = 'Table Grid'
            hdr = tbl_equip.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            secciones = [
                ("FOTO CARTEL", foto_cartel), 
                ("ENTRADA", fotos_entrada), 
                ("ALIVIO", fotos_alivio), 
                ("SALIDA", fotos_salida), 
                ("PANTALLAS", fotos_graficas)
            ]
            
            for titulo, lista in secciones:
                if lista:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2)
                    
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: cells = grid.add_row().cells
                        cell = cells[i % 2]
                        
                        sn_detectado = "No detectado"
                        coord_val = "N/A"
                        
                        if titulo == "FOTO CARTEL":
                            obs_text = "Observaciones: Fotografía del cartel de subvenciones de fondos europeos."
                        else:
                            # --- PROCESO IA MEJORADO ---
                            raw_img = Image.open(foto)
                            # Pre-procesamos la imagen para mejorar lectura
                            processed_img = preprocesar_para_ocr(raw_img)
                            
                            # La IA lee la imagen procesada
                            txt_ia = " ".join(reader.readtext(np.array(processed_img), detail=0)).upper()
                            
                            # Buscamos patrones: 0000-0000-00 o SN-XXXX
                            match = re.search(r'(\d{4}-\d{4}-\d{2}|SN-[A-Z0-9-]+)', txt_ia)
                            
                            if match:
                                sn_detectado = match.group(0)
                                if c_coord and i < len(df):
                                    coord_val = str(df.iloc[i][c_coord])
                                
                                # Añadimos a la tabla de resumen
                                r = tbl_equip.add_row().cells
                                r[0].text, r[1].text, r[2].text = titulo, sn_detectado

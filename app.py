import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
import os
from PIL import Image
import numpy as np
import re

st.set_page_config(page_title="Generador Actas ADASA-INELCOM", layout="wide")

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'])
reader = load_ocr()

def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

# --- DATOS BARRA LATERAL ---
with st.sidebar:
    st.header("📋 Datos de la EDAR")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR ALZIRA")
    localidad = st.text_input("Localidad", "Alzira")
    provincia = st.text_input("Provincia", "Valencia")
    idcoste = st.text_input("IDCOSTE", "0017")
    instaladores = st.text_input("Instaladores", "Técnico 1")
    responsable = st.text_input("Responsable Explotación", "Carlos")
    fecha = st.date_input("Fecha Instalación")

# --- INTERFAZ CENTRAL ---
st.title("📄 Generador de Actas Profesionales")
excel_file = st.file_uploader("1. Sube el Excel de Coordenadas", type=['xlsx'])

st.divider()
st.subheader("📸 2. Carga de Fotos por Orden")
col1, col2 = st.columns(2)
with col1:
    foto_puerta = st.file_uploader("🖼️ Foto Puerta (Portada)", accept_multiple_files=True)
    foto_cartel = st.file_uploader("🪧 Foto Cartel Informativo", accept_multiple_files=True)
    fotos_entrada = st.file_uploader("📥 Entrada", accept_multiple_files=True)
with col2:
    fotos_alivio = st.file_uploader("🌊 Alivio", accept_multiple_files=True)
    fotos_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    fotos_graficas = st.file_uploader("📈 Gráficas y Pantallas", accept_multiple_files=True)

if st.button("📝 GENERAR ACTA FINAL"):
    if excel_file:
        with st.spinner("Procesando fotos y detectando números de serie..."):
            df = pd.read_excel(excel_file)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # --- PÁGINA 1: PORTADA ---
            # LOGO EU GRANDE
            p_eu = doc.add_paragraph()
            p_eu.alignment = WD_ALIGN_PARAGRAPH.CENTER
            logo_eu = 'logo_instituciona.png' # Usando el nombre exacto que me dijiste
            if os.path.exists(logo_eu):
                p_eu.add_run().add_picture(logo_eu, width=Inches(6.5))
            
            doc.add_heading(nombre_edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            if foto_puerta:
                p_p = doc.add_paragraph()
                p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_p.add_run().add_picture(foto_puerta[0], width=Inches(5.5))

            # Cuadrícula Datos EDAR (Incluye Responsable)
            table_info = doc.add_table(rows=6, cols=2)
            table_info.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("LOCALIDAD", f"{localidad} ({provincia})"), 
                     ("IDCOSTE", idcoste), ("INSTALADORES", instaladores), 
                     ("RESPONSABLE", responsable), ("FECHA", str(fecha))]
            for i, (k, v) in enumerate(datos):
                table_info.rows[i].cells[0].text = k
                table_info.rows[i].cells[1].text = str(v)

            # Logos Adasa e Inelcom (Tamaño normalizado)
            p_logos = doc.add_paragraph()
            p_logos.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if os.path.exists('logo_adasa.png'): p_logos.add_run().add_picture('logo_adasa.png', width=Inches(1.2))
            p_logos.add_run("    ")
            if os.path.exists('logo_inelcom.png'): p_logos.add_run().add_picture('logo_inelcom.png', width=Inches(1.2))

            # --- PÁGINA 2: IDENTIFICACIÓN EQUIPOS ---
            doc.add_page_break()
            doc.add_heading('IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO', level=1)
            tbl_equip = doc.add_table(rows=1, cols=3)
            tbl_equip.style = 'Table Grid'
            hdr = tbl_equip.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            secciones = [("FOTO CARTEL", foto_cartel), ("ENTRADA", fotos_entrada), ("ALIVIO", fotos_alivio), ("SALIDA", fotos_salida), ("GRÁFICAS", fotos_graficas)]
            
            for titulo, lista in secciones:
                if lista:
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2) # 2 columnas para que quepan más por hoja
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: cells = grid.add_row().cells
                        cell = cells[i % 2]
                        
                        sn_detectado = "No detectado"
                        coord_val = "N/A"
                        
                        if titulo == "FOTO CARTEL":
                            obs_text = "Observaciones: Fotografía del cartel de subvenciones de fondos europeos."
                        else:
                            img = Image.open(foto)
                            txt_ia = " ".join(reader.readtext(np.array(img), detail=0))
                            # Patrón de búsqueda para tus S/N (ej: 2532-0375-20)
                            match = re.search(r'(\d{4}-\d{4}-\d{2}|SN-\w+-\d+)', txt_ia)
                            if match:
                                sn_detectado = match.group(0)
                                if c_coord and i < len(df):
                                    coord_val = str(df.iloc[i][c_coord])
                                r = tbl_equip.add_row().cells
                                r[0].text, r[1].text, r[2].text = titulo, sn_detectado, coord_val
                            obs_text = f"{titulo} S/N: {sn_detectado}"

                        cell.paragraphs[0].add_run().add_picture(foto, width=Inches(2.8))
                        p_desc = cell.add_paragraph(obs_text)
                        p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # --- CONCLUSIONES Y FIRMA ---
            doc.add_page_break()
            c_h = doc.add_heading('CONCLUSIONES', level=1)
            c_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            c_h.runs[0].font.color.rgb = RGBColor(112, 48, 160)
            
            tbl_c = doc.add_table(rows=1, cols=1)
            tbl_c.style = 'Table Grid'
            tbl_c.rows[0].cells[0].text = f"LA INSTALACIÓN EN {nombre_edar}, QUEDA COMPLETADA CORRECTAMENTE Y EN SERVICIO."

            doc.add_paragraph("\n")
            f_h = doc.add_heading('FIRMA Y VALIDACIÓN.', level=1)
            f_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            f_h.runs[0].font.color.rgb = RGBColor(112, 48, 160)
            doc.add_paragraph("Esta Asistencia Técnica de Control, certifica que la instalación ha sido supervisada y verificada.")
            
            tbl_f = doc.add_table(rows=2, cols=1)
            tbl_f.style = 'Table Grid'
            tbl_f.rows[0].cells[0].text = "FIRMA:"
            tbl_f.rows[1].height = Inches(2)

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta generada siguiendo el modelo exacto.")
            st.download_button("📥 DESCARGAR ACTA", target.getvalue(), f"Acta_{nombre_edar}.docx")

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
from PIL import Image
import numpy as np

st.set_page_config(page_title="Generador de Actas Certificadas", layout="wide")

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
    st.header("📋 Datos del Acta")
    nombre_edar = st.text_input("Nombre EDAR", "ALZIRA")
    localidad = st.text_input("Localidad", "Alzira")
    provincia = st.text_input("Provincia", "Valencia")
    idcoste = st.text_input("IDCOSTE", "0017")
    instaladores = st.text_input("Instaladores", "Técnico 1")
    fecha = st.date_input("Fecha Instalación")
    responsable = st.text_input("Responsable Explotación", "Carlos")

# --- INTERFAZ CENTRAL ---
st.title("🚀 Generador de Actas Profesionales")
excel_file = st.file_uploader("1. Sube el Excel de Coordenadas", type=['xlsx'])

st.divider()
st.subheader("📸 2. Carga de Fotos por Orden")
col1, col2 = st.columns(2)
with col1:
    foto_puerta = st.file_uploader("🖼️ Foto de la Puerta (Portada)", accept_multiple_files=True)
    foto_cartel = st.file_uploader("🪧 Foto del Cartel Informativo", accept_multiple_files=True)
    fotos_entrada = st.file_uploader("📥 Entrada (Varios puntos)", accept_multiple_files=True)
with col2:
    fotos_alivio = st.file_uploader("🌊 Alivio (EDAR/Colector)", accept_multiple_files=True)
    fotos_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    fotos_graficas = st.file_uploader("📈 Gráficas y Conclusiones", accept_multiple_files=True)

if st.button("📝 GENERAR ACTA FINAL"):
    if excel_file:
        with st.spinner("Generando documento según el modelo..."):
            df = pd.read_excel(excel_file)
            c_serie = get_column(['serie', 'sn', 's/n'], df)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            c_desc = get_column(['desc', 'nombre', 'punto'], df)

            doc = Document()

            # --- PÁGINA 1: PORTADA ---
            try:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture('logo_institucional.png', width=Inches(5))
            except: pass

            t1 = doc.add_heading(nombre_edar, 0)
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t2 = doc.add_heading('ACTA DE CERTIFICACIÓN', 1)
            t2.alignment = WD_ALIGN_PARAGRAPH.CENTER

            if foto_puerta:
                p_puerta = doc.add_paragraph()
                p_puerta.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_puerta.add_run().add_picture(foto_puerta[0], width=Inches(3.5))

            # Cuadrícula de Datos EDAR
            table_info = doc.add_table(rows=5, cols=2)
            table_info.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("LOCALIDAD", f"{localidad} ({provincia})"), 
                     ("IDCOSTE", idcoste), ("INSTALADORES", instaladores), ("FECHA", str(fecha))]
            for i, (k, v) in enumerate(datos):
                table_info.rows[i].cells[0].text = k
                table_info.rows[i].cells[1].text = str(v)

            # Logos Adasa e Inelcom bajo la tabla
            p_logos = doc.add_paragraph()
            p_logos.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                p_logos.add_run().add_picture('logo_adasa.png', width=Inches(1.2))
                p_logos.add_run("    ")
                p_logos.add_run().add_picture('logo_inelcom.png', width=Inches(1.2))
            except: pass

            # --- PÁGINA 2: IDENTIFICACIÓN EQUIPOS ---
            doc.add_page_break()
            doc.add_heading('IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO', level=1)
            tbl_equip = doc.add_table(rows=1, cols=3)
            tbl_equip.style = 'Table Grid'
            hdr = tbl_equip.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            # --- SECCIONES DE FOTOS ---
            secciones = [
                ("CARTEL INFORMATIVO", foto_cartel),
                ("ENTRADA", fotos_entrada),
                ("ALIVIO", fotos_alivio),
                ("SALIDA", fotos_salida),
                ("GRÁFICAS Y CONCLUSIONES", fotos_graficas)
            ]

            for titulo, lista in secciones:
                if lista:
                    doc.add_heading(titulo, level=1)
                    # Tabla para organizar 4 fotos por hoja (2x2)
                    grid = doc.add_table(rows=0, cols=2)
                    
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: cells = grid.add_row().cells
                        cell = cells[i % 2]
                        
                        # IA para extraer S/N si no es cartel o gráfica
                        sn, coor = "N/A", "N/A"
                        if titulo not in ["CARTEL INFORMATIVO", "GRÁFICAS Y CONCLUSIONES"]:
                            img = Image.open(foto)
                            txt = " ".join(reader.readtext(np.array(img), detail=0))
                            if c_serie:
                                for s in df[c_serie].dropna().astype(str):
                                    if s in txt and len(s) > 3:
                                        sn = s
                                        coor = str(df.loc[df[c_serie] == s, c_coord].values[0])
                                        desc = str(df.loc[df[c_serie] == s, c_desc].values[0]) if c_desc else titulo
                                        # Añadir a la tabla de la pág 2
                                        r = tbl_equip.add_row().cells
                                        r[0].text, r[1].text, r[2].text = desc, sn, coor
                                        break

                        cell.paragraphs[0].add_run().add_picture(foto, width=Inches(2.5))
                        p_desc = cell.add_paragraph(f"{titulo} - S/N: {sn}")
                        p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # --- FINAL: FIRMA ---
            doc.add_paragraph("\n\n")
            tbl_firma = doc.add_table(rows=1, cols=2)
            tbl_firma.rows[0].cells[0].text = f"Firma Instalador:\n\n________________\n{instaladores}"
            tbl_firma.rows[0].cells[1].text = f"Firma Responsable:\n\n________________\n{responsable}"

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta generada! Descárgala aquí debajo.")
            st.download_button("📥 DESCARGAR WORD", target.getvalue(), f"Acta_{nombre_edar}.docx")

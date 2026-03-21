import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import easyocr
import io
from PIL import Image
import numpy as np

# Configuración
st.set_page_config(page_title="Generador de Actas ADASA/INELCOM", layout="wide")

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'])
reader = load_ocr()

# FUNCIÓN CORREGIDA (Para evitar el error de la captura)
def get_column(keywords, df):
    for col in df.columns:
        # Convertimos el nombre de la columna a texto por si acaso hay números
        col_str = str(col)
        if any(key.lower() in col_str.lower() for key in keywords):
            return col
    return None

# --- BARRA LATERAL ---
st.sidebar.header("📍 Ubicación")
nombre_edar = st.sidebar.text_input("Nombre de la EDAR", "EJEMPLO EDAR")
localidad = st.sidebar.text_input("Localidad", "Valencia")
provincia = st.sidebar.text_input("Provincia", "Valencia")
idcoste = st.sidebar.text_input("IDCOSTE", "0017")

st.sidebar.header("👷 Personal y Fecha")
instaladores = st.sidebar.text_input("Nombre Instaladores", "Técnico 1")
fecha = st.sidebar.date_input("Fecha de Instalación")
responsable = st.sidebar.text_input("Responsable Explotación", "")

# --- CUERPO ---
st.title("📄 Generador de Actas")
excel_file = st.file_uploader("1. Sube el Excel", type=['xlsx'])
fotos = st.file_uploader("2. Sube las fotos", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

if st.button("🚀 GENERAR ACTA PROFESIONAL"):
    if excel_file and fotos:
        with st.spinner("Procesando..."):
            df = pd.read_excel(excel_file)
            col_serie = get_column(['serie', 'sn', 's/n', 'numero'], df)
            col_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            doc = Document()
            
            # Logos Cabecera
            try:
                header = doc.sections[0].header
                p = header.paragraphs[0]
                p.add_run().add_picture('logo_institucional.png', width=Inches(6))
            except: pass

            doc.add_heading(f'ACTA DE CERTIFICACIÓN - {nombre_edar}', 0)
            
            # Tabla de Datos Generales
            doc.add_heading('DATOS DE LA INSTALACIÓN', level=1)
            table_info = doc.add_table(rows=5, cols=2)
            table_info.style = 'Table Grid'
            datos_gen = [
                ("EDAR", nombre_edar),
                ("UBICACIÓN", f"{localidad} ({provincia})"),
                ("IDCOSTE", idcoste),
                ("INSTALADORES", instaladores),
                ("FECHA", str(fecha))
            ]
            for i, (clave, valor) in enumerate(datos_gen):
                table_info.rows[i].cells[0].text = clave
                table_info.rows[i].cells[1].text = str(valor)

            # Tabla Técnica
            doc.add_heading('EQUIPAMIENTO INSTALADO', level=1)
            table_equip = doc.add_table(rows=1, cols=3)
            table_equip.style = 'Table Grid'
            hdr = table_equip.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            # Procesar Fotos
            for foto_file in fotos:
                image = Image.open(foto_file)
                img_array = np.array(image)
                result = reader.readtext(img_array, detail=0)
                texto_ia = " ".join(result)
                
                sn_encontrado = "No detectado"
                coord_encontrada = "N/A"
                
                if col_serie and col_coord:
                    # Buscamos el S/N en el texto de la IA
                    for s in df[col_serie].dropna().astype(str):
                        if s.strip() and s.strip() in texto_ia:
                            sn_encontrado = s
                            coord_encontrada = str(df.loc[df[col_serie] == s, col_coord].values[0])
                            
                            r = table_equip.add_row().cells
                            r[0].text, r[1].text, r[2].text = 'SENSOR', sn_encontrado, coord_encontrada
                            break

                doc.add_heading(f"Foto de campo - S/N: {sn_encontrado}", level=2)
                doc.add_picture(foto_file, width=Inches(4))

            # Pie de Página Logos
            try:
                footer = doc.sections[0].footer.paragraphs[0]
                run = footer.add_run()
                run.add_picture('logo_adasa.png', width=Inches(1))
                run.add_run("    ")
                run.add_picture('logo_inelcom.png', width=Inches(1))
            except: pass

            # Descarga
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta generada!")
            st.download_button(label="📥 DESCARGAR WORD", data=target.getvalue(), file_name=f"Acta_{nombre_edar}.docx")
    else:
        st.error("Sube los archivos primero.")

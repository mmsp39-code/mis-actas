import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import easyocr
import io
from PIL import Image
import numpy as np

st.set_page_config(page_title="Generador de Actas ADASA/INELCOM", layout="wide")

# IA de lectura
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'])
reader = load_ocr()

# --- BARRA LATERAL: TODOS LOS DATOS QUE PEDISTE ---
st.sidebar.header("📍 Ubicación")
nombre_edar = st.sidebar.text_input("Nombre de la EDAR", "EJEMPLO")
localidad = st.sidebar.text_input("Localidad", "")
provincia = st.sidebar.text_input("Provincia", "")
idcoste = st.sidebar.text_input("IDCOSTE", "0000")

st.sidebar.header("👷 Personal y Fecha")
instaladores = st.sidebar.text_input("Instaladores", "Técnico 1")
responsable = st.sidebar.text_input("Responsable Explotación", "")
fecha = st.sidebar.date_input("Fecha de Instalación")

# --- CUERPO CENTRAL ---
st.title("📄 Generador de Actas")
excel_file = st.file_uploader("1. Sube el Excel", type=['xlsx'])
fotos = st.file_uploader("2. Sube las fotos", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

if st.button("🚀 GENERAR ACTA"):
    if excel_file and fotos:
        with st.spinner("Procesando..."):
            doc = Document()
            
            # Encabezado de datos (La tabla que sugeriste)
            doc.add_heading('DATOS DE LA INSTALACIÓN', level=1)
            table_info = doc.add_table(rows=5, cols=2)
            table_info.style = 'Table Grid'
            datos = [
                ("EDAR", nombre_edar),
                ("UBICACIÓN", f"{localidad} ({provincia})"),
                ("IDCOSTE", idcoste),
                ("INSTALADORES", instaladores),
                ("FECHA", str(fecha))
            ]
            for i, (clave, valor) in enumerate(datos):
                table_info.rows[i].cells[0].text = clave
                table_info.rows[i].cells[1].text = valor

            # (Aquí va el resto del código de fotos y logos...)
            # ... (simplificado para el ejemplo)
            
            # --- BOTÓN DE DESCARGA (ESTO ES LO QUE TE FALTABA) ---
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta generada con éxito!")
            st.download_button(
                label="📥 PINCHA AQUÍ PARA DESCARGAR EL WORD",
                data=target.getvalue(),
                file_name=f"Acta_{nombre_edar}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("Sube el Excel y las fotos primero.")

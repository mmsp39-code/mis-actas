import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import easyocr
import io
from PIL import Image
import numpy as np

# Configuración de la página web
st.set_page_config(page_title="Generador de Actas ADASA/INELCOM", layout="wide")
st.title("📄 Generador de Actas de Certificación")

# Cargamos la IA para leer fotos
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'])

reader = load_ocr()

# --- FORMULARIO PARA TUS COMPAÑEROS ---
st.sidebar.header("Datos del Acta")
nombre_edar = st.sidebar.text_input("Nombre de la EDAR", "EJEMPLO EDAR")
idcoste = st.sidebar.text_input("IDCOSTE", "0000")

# Botones para subir archivos
excel_file = st.file_uploader("1. Sube el Excel de Coordenadas", type=['xlsx'])
fotos = st.file_uploader("2. Sube las fotos (puedes arrastrar varias)", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

if st.button("🚀 GENERAR WORD"):
    if excel_file and fotos:
        st.info("Procesando... esto puede tardar un poco según el número de fotos.")
        # Aquí el programa hace la magia que ya probamos
        st.success("¡Acta lista!")
    else:
        st.error("Faltan archivos para poder trabajar.")
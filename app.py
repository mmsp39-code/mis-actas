import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import io
from PIL import Image

st.set_page_config(page_title="Generador de Actas por Carpetas", layout="wide")

st.title("📄 Generador de Actas Certificadas")
st.markdown("Sube las fotos en el apartado correspondiente. Cada apartado creará una sección en el Word.")

# --- DATOS GENERALES ---
with st.sidebar:
    st.header("📍 Datos del Proyecto")
    proyecto = st.text_input("Nombre Proyecto", "EDAR ALZIRA")
    idcoste = st.text_input("IDCOSTE", "0000")
    fecha = st.date_input("Fecha de Trabajo")

# --- 1. SUBIR EXCEL ---
excel_file = st.file_uploader("📂 1. Sube el Excel de Coordenadas", type=['xlsx'])

st.divider()

# --- 2. CAJONES DE CARGA ---
st.subheader("📸 2. Carga de Fotos por Ubicación")
st.info("Todo lo que subas a cada recuadro se agrupará bajo ese título en el acta.")

# Creamos los apartados que tú manejas
col1, col2 = st.columns(2)

with col1:
    fotos_portada = st.file_uploader("🖼️ Portada / Cartel / Puerta", accept_multiple_files=True, key="portada")
    fotos_entrada = st.file_uploader("📥 Entrada 1 / General", accept_multiple_files=True, key="ent1")
    fotos_entrada_2 = st.file_uploader("📥 Entrada 2 (Opcional)", accept_multiple_files=True, key="ent2")
    fotos_alivio_1 = st.file_uploader("🌊 Alivio Colector 1", accept_multiple_files=True, key="al1")
    fotos_alivio_2 = st.file_uploader("🌊 Alivio Colector 2", accept_multiple_files=True, key="al2")

with col2:
    fotos_salida = st.file_uploader("📤 Salida 1 / General", accept_multiple_files=True, key="sal1")
    fotos_salida_2 = st.file_uploader("📤 Salida 2 (Opcional)", accept_multiple_files=True, key="sal2")
    fotos_caudalimetro = st.file_uploader("📉 Caudalímetro / Pantallas", accept_multiple_files=True, key="caud")
    fotos_otros = st.file_uploader("🛠️ Otros / Varios", accept_multiple_files=True, key="varios")

# --- 3. GENERACIÓN ---
if st.button("🚀 GENERAR ACTA ESTRUCTURADA"):
    if excel_file:
        with st.spinner("Montando el acta por apartados..."):
            doc = Document()
            
            # Cabecera con Logo
            try:
                header = doc.sections[0].header
                header.paragraphs[0].add_run().add_picture('logo_institucional.png', width=Inches(6))
            except: pass

            doc.add_heading(f'ACTA DE CERTIFICACIÓN - {proyecto}', 0)
            
            # Estructura de secciones para el Word
            # (Título en el Word, fotos subidas en la web)
            config_secciones = [
                ("FOTOS DE PORTADA Y ENTORNO", fotos_portada),
                ("INSTALACIÓN EN ENTRADA 1", fotos_entrada),
                ("INSTALACIÓN EN ENTRADA 2", fotos_entrada_2),
                ("ALIVIO COLECTOR 1", fotos_alivio_1),
                ("ALIVIO COLECTOR 2", fotos_alivio_2),
                ("PUNTO DE SALIDA 1", fotos_salida),
                ("PUNTO DE SALIDA 2", fotos_salida_2),
                ("CAUDALÍMETRO Y GRÁFICAS", fotos_caudalimetro),
                ("OTRAS IMÁGENES DE INTERÉS", fotos_otros)
            ]

            for titulo, lista_fotos in config_secciones:
                if lista_fotos: # Solo si el usuario ha subido fotos a ese cajón
                    doc.add_heading(titulo, level=1)
                    
                    # Ponemos las fotos de ese cajón de 2 en 2 para ahorrar papel
                    for foto in lista_fotos:
                        doc.add_picture(foto, width=Inches(4))
                        doc.add_paragraph(f"Referencia: {titulo}")
                    
                    doc.add_page_break()

            # Pie de página
            try:
                footer = doc.sections[0].footer.paragraphs[0].add_run()
                footer.add_picture('logo_adasa.png', width=Inches(1))
                footer.add_run("   ")
                footer.add_picture('logo_inelcom.png', width=Inches(1))
            except: pass

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta generada siguiendo tu estructura.")
            st.download_button("📥 DESCARGAR WORD", target.getvalue(), f"Acta_{proyecto}.docx")
    else:
        st.error("Es obligatorio subir el Excel para las coordenadas.")

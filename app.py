import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

# Configuración de la página
st.set_page_config(page_title="Generador de Actas EDAR", layout="wide")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("📋 Datos del Acta")
    edar = st.text_input("EDAR", key="edar_name")
    idcoste = st.text_input("IDCOSTE")
    poblacion = st.text_input("Población")
    direccion = st.text_input("Dirección")
    provincia = st.text_input("Provincia")
    fecha = st.text_input("Fecha instalación")
    tecnicos = st.text_input("Técnicos instaladores")
    responsable = st.text_input("Responsable Explotación")
    
    st.divider()
    if st.button("♻️ REINICIAR"):
        st.rerun()

st.title("📄 Generador de Actas Profesional")

# --- SELECTORES DE FOTOS ---
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.error("#### 🖼️ GENERALES")
    f_portada = st.file_uploader("Portada", accept_multiple_files=True, key="u1")
    f_cartel = st.file_uploader("Cartel", accept_multiple_files=True, key="u2")
with c2:
    st.warning("#### 🌊 ALIVIOS")
    f_alivios = st.file_uploader("Fotos Alivios", accept_multiple_files=True, key="u4")
with c3:
    st.success("#### 📥 CAUDAL")
    f_caudal = st.file_uploader("Fotos Caudal", accept_multiple_files=True, key="u9")
with c4:
    st.info("#### 🧪 CALIDAD")
    f_calidad = st.file_uploader("Fotos Calidad", accept_multiple_files=True, key="u13")

# --- BOTÓN DE GENERACIÓN ---
if st.button("🚀 GENERAR DOCUMENTO FINAL", use_container_width=True, type="primary"):
    if not edar:
        st.warning("⚠️ Introduce el nombre de la EDAR.")
    else:
        try:
            doc = Document()

            # 1. INSERTAR LOGOS (Usando tus nombres de archivo de GitHub)
            header_table = doc.add_table(rows=1, cols=3)
            for i, logo in enumerate(["logo_adasa.png", "logo_inelcom.png", "logo_instituciona.png"]):
                if os.path.exists(logo):
                    run = header_table.cell(0, i).paragraphs[0].add_run()
                    run.add_picture(logo, width=Inches(1.2))

            # 2. DATOS DEL ACTA
            doc.add_heading(f"ACTA DE INSTALACIÓN - {edar}", 0)
            table = doc.add_table(rows=4, cols=2)
            table.style = 'Table Grid'
            datos = [("EDAR", edar), ("IDCOSTE", idcoste), ("Población", poblacion), ("Fecha", fecha)]
            for i, (campo, valor) in enumerate(datos):
                table.cell(i, 0).text = campo
                table.cell(i, 1).text = valor

            # 3. PROCESAR FOTOS (Lógica corregida)
            secciones = [("GENERALES", f_portada), ("ALIVIOS", f_alivios), ("CAUDAL", f_caudal), ("CALIDAD", f_calidad)]
            for titulo, archivos in secciones:
                if archivos:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    for foto in archivos:
                        foto_bytes = io.BytesIO(foto.read()) # ESTO ES LO QUE FALTABA
                        doc.add_picture(foto_bytes, width=Inches(4))

            # 4. GUARDAR
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta lista!")
            st.download_button("💾 DESCARGAR WORD", target.getvalue(), f"Acta_{edar}.docx")

        except Exception as e:
            st.error(f"Error: {e}")

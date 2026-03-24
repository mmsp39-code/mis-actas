import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor
import io
import os

# Configuración de página
st.set_page_config(page_title="Generador de Actas Profesional", layout="wide")

# --- BARRA LATERAL: DATOS FIJOS ---
with st.sidebar:
    st.header("📋 Datos del Acta")
    edar = st.text_input("EDAR", "Nombre de la planta")
    idcoste = st.text_input("IDCOSTE", "Referencia IDCOSTE")
    poblacion = st.text_input("Población", "")
    direccion = st.text_input("Dirección", "")
    provincia = st.text_input("Provincia", "")
    fecha = st.text_input("Fecha instalación", "DD/MM/AAAA")
    tecnicos = st.text_input("Técnicos instaladores", "")
    responsable = st.text_input("Responsable Explotación", "")
    
    st.divider()
    st.info("Rellena los campos y carga las fotos para generar el documento.")

st.title("📄 Generador de Actas ADASA & INELCOM")
st.caption("Versión Optimizada: Logos en Portada y Leyendas al Final")

st.divider()

# --- CARGA DE FOTOS ---
st.subheader("📸 Carga de Fotos por Sección")
col1, col2 = st.columns(2)

with col1:
    f_portada = st.file_uploader("🖼️ Foto Portada (Puerta/Entorno)", accept_multiple_files=True)
    f_cartel = st.file_uploader("🪧 Cartel Informativo", accept_multiple_files=True)
    f_entrada = st.file_uploader("📥 Entrada / Equipamiento", accept_multiple_files=True)
with col2:
    f_alivio = st.file_uploader("🌊 Alivio (EDAR/Colector)", accept_multiple_files=True)
    f_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    f_graficas = st.file_uploader("📈 Gráficas / Pantallas / Conclusiones", accept_multiple_files=True)

# --- BOTÓN DE GENERACIÓN ---
if st.button("🚀 GENERAR ACTA PROFESIONAL"):
    if not edar:
        st.error("Por favor, introduce al menos el nombre de la EDAR.")
    else:
        with st.spinner("Generando documento Word..."):
            doc = Document()
            
            # --- 1ª PÁGINA: PORTADA ---
            # 1. Logo Institucional
            if os.path.exists('logo_instituciona.png'):
                p_inst = doc.add_paragraph()
                p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_inst.add_run().add_picture('logo_instituciona.png', width=Inches(5))

            # 2. Títulos
            doc.add_heading(edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 3. Foto de portada
            if f_portada:
                p_portada = doc.add_paragraph()
                p_portada.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_portada.add_run().add_picture(f_portada[0], width=Inches(4.5))

            # 4. Tabla de Datos de la EDAR
            datos_form = [
                ("EDAR", edar), ("IDCOSTE", idcoste), ("Población", poblacion),
                ("Dirección", direccion), ("Provincia", provincia),
                ("Fecha instalación", fecha), ("Técnicos instaladores", tecnicos),
                ("Responsable Explotación", responsable)
            ]
            tbl = doc.add_table(rows=len(datos_form), cols=2)
            tbl.style = 'Table Grid'
            for i, (campo, valor) in enumerate(datos_form):
                tbl.rows[i].cells[0].text = campo
                tbl.rows[i].cells[1].text = str(valor)
            
            # --- NUEVA CORRECCIÓN: LOGOS ADASA & INELCOM EN PORTADA ---
            # Justo después de la tabla, centrados y pequeños
            p_logos_portada = doc.add_paragraph()
            p_logos_portada.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                if os.path.exists('logo_adasa.png'):
                    run_adasa = p_logos_portada.add_run()
                    run_adasa.add_picture('logo_adasa.png', width=Inches(1.2)) # Pequeño
                p_logos_portada.add_run("

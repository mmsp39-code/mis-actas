import streamlit as st
from docx import Document
from docx.shared import Inches
import easyocr
import io
import os
from PIL import Image, ImageOps
import numpy as np
import re

st.set_page_config(page_title="Generador de Actas Profesional", layout="wide")

@st.cache_resource
def load_ocr():
    # Cargamos el lector de IA una sola vez para ahorrar memoria
    return easyocr.Reader(['es'], gpu=False)

reader = load_ocr()

# --- BARRA LATERAL: CAMPOS FIJOS ---
with st.sidebar:
    st.header("📋 Datos del Acta")
    # Campos solicitados por el usuario
    edar = st.text_input("EDAR", "Nombre de la planta")
    idcoste = st.text_input("IDCOSTE", "Referencia IDCOSTE")
    poblacion = st.text_input("Población", "")
    direccion = st.text_input("Dirección", "")
    provincia = st.text_input("Provincia", "")
    fecha = st.text_input("Fecha instalación", "DD/MM/AAAA")
    tecnicos = st.text_input("Técnicos instaladores", "")
    responsable = st.text_input("Responsable Explotación", "")
    
    st.divider()
    st.info("Rellena estos campos y luego carga las fotos en el panel central.")

st.title("📄 Generador de Actas - ADASA & INELCOM")
st.caption("Versión 100% manual (Sin dependencia de Excel)")

st.divider()

# --- CARGA DE FOTOS POR SECCIÓN ---
st.subheader("📸 Carga de Fotos")
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
    with st.spinner("Procesando imágenes y generando documento..."):
        doc = Document()
        
        # --- LOGO PORTADA ---
        if os.path.exists('logo_instituciona.png'):
            p = doc.add_paragraph()
            p.alignment = 1
            p.add_run().add_picture('logo_instituciona.png', width=Inches(5))

        doc.add_heading(edar, 0).alignment = 1
        doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = 1

        # Foto de portada si existe
        if f_portada:
            p = doc.add_paragraph()
            p.alignment = 1
            p.add_run().add_picture(f_portada[0], width=Inches(4.5))

        # --- TABLA DE DATOS (Sustituye al Excel) ---
        tbl = doc.add_table(rows=8, cols=2)
        tbl.style = 'Table Grid'
        
        datos_tabla = [
            ("EDAR", edar),
            ("IDCOSTE", idcoste),
            ("Población", poblacion),
            ("Dirección", direccion),
            ("Provincia", provincia),
            ("Fecha instalación", fecha),
            ("Técnicos instaladores", tecnicos),
            ("Responsable Explotación", responsable)
        ]
        
        for i, (campo, valor) in enumerate(datos_tabla):
            tbl.rows[i].cells[0].text = campo
            tbl.rows[i].cells[1].text = str(valor)

        # --- SECCIONES CON IA PARA NÚMEROS DE SERIE (S/N) ---
        secciones = [
            ("CARTEL INFORMATIVO", f_cartel, "fijo"),
            ("EQUIPAMIENTO EN ENTRADA", f_entrada, "ia"),
            ("EQUIPAMIENTO EN ALIVIO", f_alivio, "ia"),
            ("EQUIPAMIENTO EN SALIDA", f_salida, "ia"),
            ("GRÁFICAS Y CONCLUSIONES", f_graficas, "fijo")
        ]

        for titulo, lista, modo in secciones:
            if lista:
                doc.add_page_break()
                doc.add_heading(titulo, level=1)
                grid = doc.add_table(rows=0, cols=2)
                
                for i, foto in enumerate(lista):
                    if i % 2 == 0: row_cells = grid.add_row().cells
                    cell = row_cells[i % 2]
                    
                    sn_txt = ""
                    if modo == "ia":
                        # Procesamos con la IA para buscar el S/N
                        img = ImageOps.grayscale(Image.open(foto))
                        res = " ".join(reader.readtext(np.array(img), detail=0)).upper()
                        m = re.search(r'(SN-[A-Z0-9-]+|\d{4}[-_]\d{4}[-_]\d{2})', res)
                        sn = m.group(0).replace("_", "-") if m else "No detectado automáticamente"
                        sn_txt = f"S/N: {sn}"
                    elif titulo == "CARTEL INFORMATIVO":
                        sn_txt = "Fotografía del cartel de subvenciones."
                    
                    cell.paragraphs[0].add_run().add_picture(foto, width=Inches(2.5))
                    if sn_txt:
                        cell.add_paragraph(sn_txt).alignment = 1

        # --- FIRMA Y LOGOS FINALES ---
        doc.add_paragraph("\n\n")
        p_final = doc.add_paragraph()
        p_final.alignment = 1
        
        try:
            if os.path.exists('logo_adasa.png'):
                p_final.add_run().add_picture('logo_adasa.png', width=Inches(1.5))
            p_final.add_run("    ")
            if os.path.exists('logo_inelcom.png'):
                p_final.add_run().add_picture('logo_inelcom.png', width=Inches(1.5))
        except: pass

        doc.add_paragraph("\n")
        tbl_f = doc.add_table(rows=2, cols=1)
        tbl_f.style = 'Table Grid'
        tbl_f.rows[0].cells[0].text = "FIRMA Y VALIDACIÓN:"
        tbl_f.rows[1].height = Inches(1.2)

        # Preparar descarga
        target = io.BytesIO()
        doc.save(target)
        st.success(f"✅ Acta de {edar} generada correctamente.")
        st.download_button("📥 DESCARGAR ACTA EN WORD", target.getvalue(), f"Acta_{edar.replace(' ', '_')}.docx")

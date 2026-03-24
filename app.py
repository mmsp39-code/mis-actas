import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

st.set_page_config(page_title="Generador de Actas EDAR - Diseño Pro", layout="wide")

# --- DATOS LATERALES ---
with st.sidebar:
    st.header("📋 Datos del Acta")
    edar = st.text_input("EDAR", key="edar_name")
    idcoste = st.text_input("IDCOSTE", key="id_coste")
    poblacion = st.text_input("Población", key="pob")
    direccion = st.text_input("Dirección", key="dir")
    provincia = st.text_input("Provincia", key="prov")
    fecha = st.text_input("Fecha instalación", key="fec")
    tecnicos = st.text_input("Técnicos instaladores", key="tec")
    responsable = st.text_input("Responsable Explotación", key="resp")
    if st.button("♻️ REINICIAR"):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

st.title("📄 Generador de Actas - Estructura Final")

# --- BLOQUES DE CARGA (Tus 16 selectores) ---
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.error("#### 🖼️ GENERALES")
    f_portada = st.file_uploader("Portada", accept_multiple_files=True, key="u1")
    f_cartel = st.file_uploader("Cartel", accept_multiple_files=True, key="u2")
    f_graficas = st.file_uploader("Gráficas", accept_multiple_files=True, key="u3")
with c2:
    st.warning("#### 🌊 ALIVIOS")
    f_alivio_e1 = st.file_uploader("A. EDAR 1", accept_multiple_files=True, key="u4")
    f_alivio_e2 = st.file_uploader("A. EDAR 2", accept_multiple_files=True, key="u5")
    f_alivio_c1 = st.file_uploader("A. Col 1", accept_multiple_files=True, key="u6")
    f_alivio_c2 = st.file_uploader("A. Col 2", accept_multiple_files=True, key="u7")
    f_alivio_c3 = st.file_uploader("A. Col 3", accept_multiple_files=True, key="u8")
with c3:
    st.success("#### 📥 CAUDA.")
    f_cauda_e1 = st.file_uploader("Ent. 1", accept_multiple_files=True, key="u9")
    f_cauda_e2 = st.file_uploader("Ent. 2", accept_multiple_files=True, key="u10")
    f_cauda_s1 = st.file_uploader("Sal. 1", accept_multiple_files=True, key="u11")
    f_cauda_s2 = st.file_uploader("Sal. 2", accept_multiple_files=True, key="u12")
with c4:
    st.info("#### 🧪 CALIDAD")
    f_calid_e1 = st.file_uploader("Ent. 1", accept_multiple_files=True, key="u13")
    f_calid_e2 = st.file_uploader("Ent. 2", accept_multiple_files=True, key="u14")
    f_calid_s1 = st.file_uploader("Sal. 1", accept_multiple_files=True, key="u15")
    f_calid_s2 = st.file_uploader("Sal. 2", accept_multiple_files=True, key="u16")

# --- GENERACIÓN DEL DOCUMENTO ---
if st.button("🚀 GENERAR ACTA FORMATEADA", use_container_width=True, type="primary"):
    if not edar:
        st.warning("Falta el nombre de la EDAR.")
    else:
        try:
            doc = Document()
            
            # 1. LOGOS ENCABEZADO (Tabla invisible de 3 columnas)
            header_table = doc.add_table(rows=1, cols=3)
            header_table.width = Inches(7)
            logos = ["logo_adasa.png", "logo_inelcom.png", "logo_instituciona.png"]
            for i, l_name in enumerate(logos):
                if os.path.exists(l_name):
                    cell = header_table.cell(0, i)
                    p = cell.paragraphs[0]
                    run = p.add_run()
                    run.add_picture(l_name, width=Inches(1.2))
                    if i == 1: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if i == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            doc.add_heading(f"ACTA DE CONTROL E INSTALACIÓN: {edar}", 0)

            # 2. TABLA DE DATOS (2 columnas, limpia)
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            datos = [("EDAR", edar), ("IDCOSTE", idcoste), ("POBLACIÓN", poblacion), ("PROVINCIA", provincia), ("DIRECCIÓN", direccion), ("FECHA", fecha), ("TÉCNICOS", tecnicos), ("RESPONSABLE", responsable)]
            for c, v in datos:
                row = table.add_row().cells
                row[0].text = c
                row[1].text = str(v)

            # 3. FOTOS EN CUADRÍCULA (2 por fila)
            bloques = [
                ("FOTOS GENERALES", [f_portada, f_cartel, f_graficas]),
                ("SECCIÓN ALIVIOS", [f_alivio_e1, f_alivio_e2, f_alivio_c1, f_alivio_c2, f_alivio_c3]),
                ("SECCIÓN CAUDALÍMETROS", [f_cauda_e1, f_cauda_e2, f_cauda_s1, f_cauda_s2]),
                ("SECCIÓN CALIDAD", [f_calid_e1, f_calid_e2, f_calid_s1, f_calid_s2])
            ]

            for titulo, uploaders in bloques:
                todas_las_fotos = [f for up in uploaders if up for f in up]
                if todas_las_fotos:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    # Creamos tabla para fotos (2 columnas)
                    grid = doc.add_table(rows=0, cols=2)
                    for i in range(0, len(todas_las_fotos), 2):
                        row_cells = grid.add_row().cells
                        for j in range(2):
                            if i + j < len(todas_las_fotos):
                                f = todas_las_fotos[i + j]
                                p = row_cells[j].paragraphs[0]
                                run = p.add_run()
                                run.add_picture(io.BytesIO(f.read()), width=Inches(3.1))
                                row_cells[j].add_paragraph(f.name).alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 4. SECCIÓN DE FIRMAS
            doc.add_page_break()
            doc.add_heading("FIRMAS Y CONFORMIDAD", level=1)
            f_table = doc.add_table(rows=2, cols=2)
            f_table.width = Inches(7)
            f_table.cell(0,0).text = "\n\n__________________________\nFirma Técnico Instalador"
            f_table.cell(0,1).text = "\n\n__________________________\nFirma Responsable Explotación"

            # 5. DESCARGA
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta estructurada correctamente.")
            st.download_button("💾 DESCARGAR ACTA (.docx)", target.getvalue(), f"Acta_{edar}.docx")
        except Exception as e:
            st.error(f"Error: {e}")

import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

# 1. CONFIGURACIÓN
st.set_page_config(page_title="Generador de Actas EDAR", layout="wide")

def restart_application():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# 2. BARRA LATERAL - ENTRADA DE DATOS
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
    
    st.divider()
    if st.button("♻️ REINICIAR TODO", use_container_width=True):
        restart_application()

st.title("📄 Generador de Actas Profesional")

# 3. CARGA DE FOTOS (Tus 16 campos originales)
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
    st.success("#### 📥 CAUDAL.")
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

# 4. BOTÓN DE GENERACIÓN
if st.button("🚀 GENERAR DOCUMENTO FINAL", use_container_width=True, type="primary"):
    if not edar:
        st.warning("Escribe el nombre de la EDAR.")
    else:
        try:
            doc = Document()

            # Logos del encabezado
            header_table = doc.add_table(rows=1, cols=3)
            logos = ["logo_adasa.png", "logo_inelcom.png", "logo_instituciona.png"]
            for i, l_name in enumerate(logos):
                if os.path.exists(l_name):
                    p = header_table.cell(0, i).paragraphs[0]
                    run = p.add_run()
                    run.add_picture(l_name, width=Inches(1.2))

            # Título y tabla de datos
            doc.add_heading(f"ACTA DE INSTALACIÓN: {edar}", 0)
            table = doc.add_table(rows=8, cols=2)
            table.style = 'Table Grid'
            datos_acta = [
                ("EDAR", edar), ("IDCOSTE", idcoste), ("Población", poblacion),
                ("Dirección", direccion), ("Provincia", provincia), ("Fecha", fecha),
                ("Técnicos", tecnicos), ("Responsable", responsable)
            ]
            for i, (campo, valor) in enumerate(datos_acta):
                table.cell(i, 0).text = str(campo)
                table.cell(i, 1).text = str(valor)

            # Inserción de fotos por bloques
            bloques = [
                ("FOTOS GENERALES", [f_portada, f_cartel, f_graficas]),
                ("ALIVIOS", [f_alivio_e1, f_alivio_e2, f_alivio_c1, f_alivio_c2, f_alivio_c3]),
                ("CAUDALÍMETROS", [f_cauda_e1, f_cauda_e2, f_cauda_s1, f_cauda_s2]),
                ("CALIDAD", [f_calid_e1, f_calid_e2, f_calid_s1, f_calid_s2])
            ]

            for titulo_bloque, lista_fotos in bloques:
                if any(lista_fotos):
                    doc.add_page_break()
                    doc.add_heading(titulo_bloque, level=1)
                    for f_upped in lista_fotos:
                        if f_upped:
                            for f in f_upped:
                                img_stream = io.BytesIO(f.read())
                                doc.add_picture(img_stream, width=Inches(3.8))
                                doc.add_paragraph(f"Imagen: {f.name}")

            # Descarga del archivo
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Documento creado!")
            st.download_button("💾 DESCARGAR WORD", target.getvalue(), f"Acta_{edar}.docx")

        except Exception as e:
            st.error(f"Error al generar el documento: {e}")

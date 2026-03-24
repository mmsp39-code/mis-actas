import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

st.set_page_config(page_title="Generador de Actas Ultra-Compacto", layout="wide")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("📋 DATOS DEL ACTA")
    edar = st.text_input("Nombre de la EDAR", key="edar_name").upper()
    idcoste = st.text_input("IDCOSTE", key="id_coste")
    poblacion = st.text_input("Población", key="pob")
    direccion = st.text_input("Dirección", key="dir")
    provincia = st.text_input("Provincia", key="prov")
    fecha = st.text_input("Fecha instalación", key="fec")
    tecnicos = st.text_input("Técnicos instaladores", key="tec")
    responsable = st.text_input("Responsable Explotación", key="resp")

# --- SELECTORES (16 ACCESOS ORIGINALES) ---
st.title("📄 Generador de Acta: Máximo Aprovechamiento")
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.error("#### 🖼️ GENERALES")
    f_portada = st.file_uploader("Portada", accept_multiple_files=False, key="u1")
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
    st.success("#### 📥 CAUDAL")
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

# --- GENERACIÓN ---
if st.button("🚀 GENERAR ACTA COMPACTA (6 FOTOS/PAG)", use_container_width=True, type="primary"):
    if not edar:
        st.warning("Falta el nombre de la EDAR.")
    else:
        try:
            doc = Document()
            
            # 1. LOGO INSTITUCIONAL (Cabecera) [cite: 9]
            if os.path.exists("logo_instituciona.png"):
                p_inst = doc.add_paragraph()
                p_inst.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_inst.add_run().add_picture("logo_instituciona.png", width=Inches(2.5))

            # 2. PORTADA (Nombre EDAR + Acta + Portada) [cite: 10, 11, 15]
            doc.add_heading(edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading("ACTA DE CERTIFICACIÓN", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
            if f_portada:
                p_port = doc.add_paragraph()
                p_port.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_port.add_run().add_picture(io.BytesIO(f_portada.read()), width=Inches(3.5))

            # 3. CUADRÍCULA DE DATOS [cite: 12]
            data_table = doc.add_table(rows=0, cols=2)
            data_table.style = 'Table Grid'
            campos = [("EDAR", edar), ("IDCOSTE", idcoste), ("POBLACIÓN", poblacion), ("FECHA", fecha), ("TÉCNICOS", tecnicos)]
            for c, v in campos:
                row = data_table.add_row().cells
                row[0].text = c
                row[1].text = str(v)
            
            # 4. LOGOS PEQUEÑOS CENTRADOS [cite: 9]
            doc.add_paragraph("\n")
            log_table = doc.add_table(rows=1, cols=2)
            for i, l_name in enumerate(["logo_adasa.png", "logo_inelcom.png"]):
                if os.path.exists(l_name):
                    p_log = log_table.cell(0, i).paragraphs[0]
                    p_log.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_log.add_run().add_picture(l_name, width=Inches(1.0))

            # 5. TABLA EQUIPAMIENTO (Pág 2) [cite: 13, 14]
            doc.add_page_break()
            doc.add_heading("IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO", level=1)
            eq_table = doc.add_table(rows=5, cols=3)
            eq_table.style = 'Table Grid'
            for i, h in enumerate(['EQUIPAMIENTO', 'NÚMERO DE SERIE', 'COORDENADAS']):
                eq_table.cell(0, i).text = h

            # 6. SECCIONES TÉCNICAS (6 FOTOS POR PÁGINA) 
            secciones = [
                ("CARTEL", [f_cartel]),
                ("ALIVIOS", [f_alivio_e1, f_alivio_e2, f_alivio_c1, f_alivio_c2, f_alivio_c3]),
                ("CAUDALÍMETROS", [f_cauda_e1, f_cauda_e2, f_cauda_s1, f_cauda_s2]),
                ("SENSORES CALIDAD", [f_calid_e1, f_calid_e2, f_calid_s1, f_calid_s2])
            ]

            for titulo, uploaders in secciones:
                imgs = [foto for up in uploaders if up for foto in up]
                if imgs:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    # Cuadrícula 2 columnas x N filas
                    grid = doc.add_table(rows=0, cols=2)
                    for i in range(0, len(imgs), 2):
                        row = grid.add_row().cells
                        for j in range(2):
                            if i+j < len(imgs):
                                p_g = row[j].paragraphs[0]
                                p_g.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p_g.add_run().add_picture(io.BytesIO(imgs[i+j].read()), width=Inches(2.3)) # Tamaño para 6/pág
                    # Observación ÚNICA al final de la sección [cite: 5, 20]
                    doc.add_paragraph(f"\nObservaciones {titulo}: ___________________________________________")

            # 7. GRÁFICAS Y CIERRE [cite: 19, 28, 30]
            if f_graficas:
                doc.add_page_break()
                doc.add_heading("GRÁFICAS", level=1)
                for g in f_graficas:
                    doc.add_picture(io.BytesIO(g.read()), width=Inches(4.5))
                doc.add_paragraph("\nObservaciones Gráficas: ___________________________________________")

            doc.add_page_break()
            doc.add_heading("CONCLUSIONES", level=1)
            doc.add_paragraph(f"LA INSTALACIÓN EN EDAR {edar} QUEDA COMPLETADA CORRECTAMENTE.")
            doc.add_paragraph("\n\nFIRMA:__________________________") [cite: 32]

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta Ultra-Compacta generada.")
            st.download_button("💾 DESCARGAR ACTA", target.getvalue(), f"Acta_{edar}.docx")
        except Exception as e:
            st.error(f"Error: {e}")

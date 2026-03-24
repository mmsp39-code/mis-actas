import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

st.set_page_config(page_title="Generador de Actas - Perfección Visual", layout="wide")

# --- DATOS ---
with st.sidebar:
    st.header("📋 DATOS DEL ACTA")
    edar = st.text_input("Nombre de la EDAR", key="edar_name").upper()
    idcoste = st.text_input("IDCOSTE", key="id_coste")
    pob = st.text_input("Población", key="pob")
    dir_inst = st.text_input("Dirección", key="dir")
    prov = st.text_input("Provincia", key="prov")
    fec = st.text_input("Fecha instalación", key="fec")
    tec = st.text_input("Técnicos instaladores", key="tec")
    resp = st.text_input("Responsable Explotación", key="resp")

# --- CARGA DE FOTOS ---
st.title("📄 Generador de Acta: Modelo Exacto")
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.error("#### 🖼️ PORTADA")
    f_portada = st.file_uploader("Foto de Portada", accept_multiple_files=False)
    f_cartel = st.file_uploader("Foto Cartel", accept_multiple_files=True)
with c2:
    st.warning("#### 🌊 ALIVIOS")
    f_alivios = [st.file_uploader(f"Alivio {i}", accept_multiple_files=True, key=f"al{i}") for i in range(1,3)]
with c3:
    st.success("#### 📥 CAUDAL/CALIDAD")
    f_cauda_calid = [st.file_uploader(f"Equipo {i}", accept_multiple_files=True, key=f"eq{i}") for i in range(1,4)]
with c4:
    st.info("#### 📈 CIERRE")
    f_graficas = st.file_uploader("Gráficas", accept_multiple_files=True)

# --- GENERACIÓN ---
if st.button("🚀 GENERAR ACTA IDENTICA AL MODELO", use_container_width=True, type="primary"):
    if not edar:
        st.warning("Introduce el nombre de la EDAR.")
    else:
        try:
            doc = Document()
            
            # 1. LOGO INSTITUCIONAL (ARRIBA DE TODO)
            if os.path.exists("logo_instituciona.png"):
                p_inst = doc.add_paragraph()
                p_inst.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_inst.add_run().add_picture("logo_instituciona.png", width=Inches(2.5))

            # 2. NOMBRE EDAR Y ACTA
            doc.add_paragraph("\n")
            t_edar = doc.add_heading(edar, 0)
            t_edar.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t_acta = doc.add_heading("ACTA DE CERTIFICACIÓN", level=1)
            t_acta.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("\n")

            # 3. FOTO DE PORTADA (DEBAJO DEL TÍTULO)
            if f_portada:
                p_port = doc.add_paragraph()
                p_port.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_port.add_run().add_picture(io.BytesIO(f_portada.read()), width=Inches(4.5))
            
            doc.add_paragraph("\n")

            # 4. CUADRÍCULA DE DATOS
            data_table = doc.add_table(rows=0, cols=2)
            data_table.style = 'Table Grid'
            campos = [("EDAR", edar), ("IDCOSTE", idcoste), ("DIRECCIÓN", dir_inst), 
                      ("POBLACIÓN", pob), ("PROVINCIA", prov), ("FECHA", fec), 
                      ("TÉCNICOS", tec), ("RESPONSABLE", resp)]
            for c, v in campos:
                row = data_table.add_row().cells
                row[0].text = c
                row[1].text = str(v)
            
            doc.add_paragraph("\n")

            # 5. LOGOS ADASA E INELCOM (BAJO LA CUADRÍCULA, CENTRADOS)
            log_table = doc.add_table(rows=1, cols=2)
            log_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for i, l_name in enumerate(["logo_adasa.png", "logo_inelcom.png"]):
                if os.path.exists(l_name):
                    p_log = log_table.cell(0, i).paragraphs[0]
                    p_log.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_log.add_run().add_picture(l_name, width=Inches(1.2))

            # 6. TABLA EQUIPAMIENTO (Siguiente página)
            doc.add_page_break()
            doc.add_heading("IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO", level=1)
            eq_table = doc.add_table(rows=6, cols=3)
            eq_table.style = 'Table Grid'
            for i, h in enumerate(['EQUIPAMIENTO', 'NÚMERO DE SERIE', 'COORDENADAS']):
                eq_table.cell(0, i).text = h

            # 7. SECCIONES DE FOTOS TÉCNICAS (EN PAREJAS)
            secciones = [("FOTO CARTEL", [f_cartel]), ("ALIVIOS", f_alivios), ("EQUIPOS CAUDAL/CALIDAD", f_cauda_calid)]
            for titulo, lista_up in secciones:
                imgs = [foto for up in lista_up if up for foto in up]
                if imgs:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2)
                    for i in range(0, len(imgs), 2):
                        row = grid.add_row().cells
                        for j in range(2):
                            if i+j < len(imgs):
                                p_g = row[j].paragraphs[0]
                                p_g.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p_g.add_run().add_picture(io.BytesIO(imgs[i+j].read()), width=Inches(3.0))
                                row[j].add_paragraph("Observaciones: _________").alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 8. GRÁFICAS (AL FINAL)
            if f_graficas:
                doc.add_page_break()
                doc.add_heading("GRÁFICAS", level=1)
                for g in f_graficas:
                    p_g = doc.add_paragraph()
                    p_g.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_g.add_run().add_picture(io.BytesIO(g.read()), width=Inches(5.0))

            # 9. CONCLUSIÓN Y FIRMA
            doc.add_page_break()
            doc.add_heading("CONCLUSIONES", level=1)
            doc.add_paragraph(f"LA INSTALACIÓN EN EDAR {edar} QUEDA COMPLETADA Y EN SERVICIO.")
            doc.add_paragraph("\n\n\nFIRMA Y VALIDACIÓN\n\nFIRMA:__________________________")

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Estructura corregida al 100% según tu modelo.")
            st.download_button("💾 DESCARGAR ACTA", target.getvalue(), f"Acta_{edar}.docx")
        except Exception as e:
            st.error(f"Error: {e}")

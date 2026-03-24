import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import os

st.set_page_config(page_title="Generador de Actas Certificación - Versión Final", layout="wide")

# --- BARRA LATERAL ---
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
    if st.button("♻️ REINICIAR"):
        st.rerun()

st.title("📄 Generador de Acta: Modelo Certificación")

# --- SELECTORES DE FOTOS ---
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.error("#### 🖼️ PRINCIPALES")
    f_portada = st.file_uploader("1. Portada", accept_multiple_files=True)
    f_cartel = st.file_uploader("2. Cartel", accept_multiple_files=True)
    f_graficas = st.file_uploader("3. Gráficas (Irán al final)", accept_multiple_files=True)
with c2:
    st.warning("#### 🌊 ALIVIOS")
    f_alivios = [st.file_uploader(f"Alivio {i}", accept_multiple_files=True, key=f"al{i}") for i in range(1,6)]
with c3:
    st.success("#### 📥 CAUDAL")
    f_caudal = [st.file_uploader(f"Caudal {i}", accept_multiple_files=True, key=f"ca{i}") for i in range(1,5)]
with c4:
    st.info("#### 🧪 CALIDAD")
    f_calidad = [st.file_uploader(f"Calidad {i}", accept_multiple_files=True, key=f"cli{i}") for i in range(1,5)]

# --- MOTOR DE GENERACIÓN PERFECCIONISTA ---
if st.button("🚀 GENERAR ACTA DE ALTA CALIDAD", use_container_width=True, type="primary"):
    if not edar:
        st.warning("Falta el nombre de la EDAR.")
    else:
        try:
            doc = Document()
            
            # 1. CABECERA TRIPLE LOGO
            head_table = doc.add_table(rows=1, cols=3)
            head_table.width = Inches(7)
            logos = ["logo_adasa.png", "logo_inelcom.png", "logo_instituciona.png"]
            for i, name in enumerate(logos):
                if os.path.exists(name):
                    p = head_table.cell(0, i).paragraphs[0]
                    r = p.add_run()
                    r.add_picture(name, width=Inches(1.4))
                    if i == 1: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if i == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # 2. TÍTULOS DE PORTADA
            doc.add_paragraph("\n")
            t1 = doc.add_heading(edar, 0)
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t2 = doc.add_heading("ACTA DE CERTIFICACIÓN", level=1)
            t2.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 3. CUADRÍCULA DE DATOS (ESTILO ORIGINAL)
            doc.add_paragraph("\n")
            data_table = doc.add_table(rows=0, cols=2)
            data_table.style = 'Table Grid'
            campos = [("EDAR", edar), ("IDCOSTE", idcoste), ("DIRECCIÓN", dir_inst), 
                      ("POBLACIÓN", pob), ("PROVINCIA", prov), ("FECHA", fec), 
                      ("TÉCNICOS", tec), ("RESPONSABLE", resp)]
            for c, v in campos:
                row = data_table.add_row().cells
                row[0].text = c
                row[1].text = str(v)

            # 4. TABLA DE EQUIPAMIENTO VACÍA
            doc.add_page_break()
            doc.add_heading("IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO", level=1)
            eq_table = doc.add_table(rows=6, cols=3)
            eq_table.style = 'Table Grid'
            for i, h in enumerate(['EQUIPAMIENTO', 'NÚMERO DE SERIE', 'COORDENADAS']):
                eq_table.cell(0, i).text = h

            # 5. FOTOS: PORTADA Y CARTEL (1 por página o grandes)
            for tit, f in [("PORTADA", f_portada), ("FOTO CARTEL", f_cartel)]:
                if f:
                    doc.add_page_break()
                    doc.add_heading(tit, level=1)
                    for img in f:
                        doc.add_picture(io.BytesIO(img.read()), width=Inches(5))
                        doc.add_paragraph("Observaciones: ___________________________________")

            # 6. SECCIONES TÉCNICAS (EN PAREJAS 2x2)
            secciones = [("ALIVIOS", f_alivios), ("CAUDALÍMETROS", f_caudal), ("CALIDAD", f_calidad)]
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
                                pic = row[j].paragraphs[0].add_run()
                                pic.add_picture(io.BytesIO(imgs[i+j].read()), width=Inches(3))
                                row[j].add_paragraph("Observaciones: _________")

            # 7. GRÁFICAS (AL FINAL)
            if f_graficas:
                doc.add_page_break()
                doc.add_heading("GRÁFICAS DE FUNCIONAMIENTO", level=1)
                for g in f_graficas:
                    doc.add_picture(io.BytesIO(g.read()), width=Inches(5))

            # 8. CONCLUSIONES Y FIRMA
            doc.add_page_break()
            doc.add_heading("CONCLUSIONES", level=1)
            doc.add_paragraph(f"LA INSTALACIÓN EN EDAR {edar} QUEDA COMPLETADA Y EN SERVICIO.")
            doc.add_paragraph("\n\n\nFIRMA Y VALIDACIÓN\n\nFIRMA:__________________________")

            # FINALIZAR
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta perfeccionada generada.")
            st.download_button("💾 DESCARGAR ACTA PROFESIONAL", target.getvalue(), f"Certificacion_{edar}.docx")
        except Exception as e:
            st.error(f"Error: {e}")

import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

st.set_page_config(page_title="Generador de Actas Adaptable", layout="wide")

# --- BARRA LATERAL (Tus datos de siempre) ---
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

# --- SELECTORES DE FOTOS (Tus 16 accesos para máxima flexibilidad) ---
st.title("📄 Generador de Acta Adaptable")
st.info("Sube fotos solo en los apartados que necesites. Los vacíos no aparecerán en el Word.")

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.error("#### 🖼️ GENERALES")
    f_portada = st.file_uploader("Portada (Principal)", accept_multiple_files=False, key="u1")
    f_cartel = st.file_uploader("Cartel", accept_multiple_files=True, key="u2")
    f_graficas = st.file_uploader("Gráficas (Final)", accept_multiple_files=True, key="u3")

with c2:
    st.warning("#### 🌊 ALIVIOS")
    f_alivio_e1 = st.file_uploader("Alivio EDAR 1", accept_multiple_files=True, key="u4")
    f_alivio_e2 = st.file_uploader("Alivio EDAR 2", accept_multiple_files=True, key="u5")
    f_alivio_c1 = st.file_uploader("Alivio Col 1", accept_multiple_files=True, key="u6")
    f_alivio_c2 = st.file_uploader("Alivio Col 2", accept_multiple_files=True, key="u7")
    f_alivio_c3 = st.file_uploader("Alivio Col 3", accept_multiple_files=True, key="u8")

with c3:
    st.success("#### 📥 CAUDAL")
    f_cauda_e1 = st.file_uploader("Entrada 1", accept_multiple_files=True, key="u9")
    f_cauda_e2 = st.file_uploader("Entrada 2", accept_multiple_files=True, key="u10")
    f_cauda_s1 = st.file_uploader("Salida 1", accept_multiple_files=True, key="u11")
    f_cauda_s2 = st.file_uploader("Salida 2", accept_multiple_files=True, key="u12")

with c4:
    st.info("#### 🧪 CALIDAD")
    f_calid_e1 = st.file_uploader("Sensor Ent. 1", accept_multiple_files=True, key="u13")
    f_calid_e2 = st.file_uploader("Sensor Ent. 2", accept_multiple_files=True, key="u14")
    f_calid_s1 = st.file_uploader("Sensor Sal. 1", accept_multiple_files=True, key="u15")
    f_calid_s2 = st.file_uploader("Sensor Sal. 2", accept_multiple_files=True, key="u16")

# --- LÓGICA DE GENERACIÓN ---
if st.button("🚀 GENERAR ACTA A MEDIDA", use_container_width=True, type="primary"):
    if not edar:
        st.warning("Falta el nombre de la EDAR.")
    else:
        try:
            doc = Document()
            
            # 1. LOGO INSTITUCIONAL (Cabecera derecha)
            if os.path.exists("logo_instituciona.png"):
                p_inst = doc.add_paragraph()
                p_inst.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_inst.add_run().add_picture("logo_instituciona.png", width=Inches(3.0))

            # 2. TÍTULOS PORTADA
            doc.add_paragraph("\n")
            doc.add_heading(edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading("ACTA DE CERTIFICACIÓN", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 3. FOTO PORTADA (Solo si existe)
            if f_portada:
                p_port = doc.add_paragraph()
                p_port.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_port.add_run().add_picture(io.BytesIO(f_portada.read()), width=Inches(4.5))
            
            doc.add_paragraph("\n")

            # 4. CUADRÍCULA DE DATOS
            data_table = doc.add_table(rows=0, cols=2)
            data_table.style = 'Table Grid'
            campos = [("EDAR", edar), ("IDCOSTE", idcoste), ("DIRECCIÓN", direccion), 
                      ("POBLACIÓN", poblacion), ("PROVINCIA", provincia), ("FECHA", fecha), 
                      ("TÉCNICOS", tecnicos), ("RESPONSABLE", responsable)]
            for c, v in campos:
                row = data_table.add_row().cells
                row[0].text = c
                row[1].text = str(v)
            
            # 5. LOGOS ADASA E INELCOM (Bajo la tabla)
            doc.add_paragraph("\n")
            log_table = doc.add_table(rows=1, cols=2)
            for i, l_name in enumerate(["logo_adasa.png", "logo_inelcom.png"]):
                if os.path.exists(l_name):
                    p_log = log_table.cell(0, i).paragraphs[0]
                    p_log.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_log.add_run().add_picture(l_name, width=Inches(1.2))

            # 6. TABLA EQUIPAMIENTO (Página 2)
            doc.add_page_break()
            doc.add_heading("IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO", level=1)
            eq_table = doc.add_table(rows=6, cols=3)
            eq_table.style = 'Table Grid'
            for i, h in enumerate(['EQUIPAMIENTO', 'NÚMERO DE SERIE', 'COORDENADAS']):
                eq_table.cell(0, i).text = h

            # 7. BLOQUES TÉCNICOS ADAPTABLES
            secciones = [
                ("FOTO CARTEL", [f_cartel]),
                ("ALIVIOS", [f_alivio_e1, f_alivio_e2, f_alivio_c1, f_alivio_c2, f_alivio_c3]),
                ("CAUDALÍMETROS", [f_cauda_e1, f_cauda_e2, f_cauda_s1, f_cauda_s2]),
                ("SENSORES CALIDAD", [f_calid_e1, f_calid_e2, f_calid_s1, f_calid_s2])
            ]

            for titulo, lista_up in secciones:
                imgs = [foto for up in lista_up if up for foto in up]
                if imgs: # Solo crea la sección si hay fotos
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

            # 8. GRÁFICAS (Al final de lo técnico)
            if f_graficas:
                doc.add_page_break()
                doc.add_heading("GRÁFICAS DE FUNCIONAMIENTO", level=1)
                for g in f_graficas:
                    p_g = doc.add_paragraph()
                    p_g.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_g.add_run().add_picture(io.BytesIO(g.read()), width=Inches(5.0))
                    doc.add_paragraph("Observaciones: ___________________________________")

            # 9. CIERRE
            doc.add_page_break()
            doc.add_heading("CONCLUSIONES", level=1)
            doc.add_paragraph(f"LA INSTALACIÓN EN EDAR {edar} QUEDA COMPLETADA CORRECTAMENTE.")
            doc.add_paragraph("\n\n\nFIRMA Y VALIDACIÓN\n\nFIRMA:__________________________")

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta generada. Se ha adaptado a los apartados completados.")
            st.download_button("💾 DESCARGAR ACTA", target.getvalue(), f"Acta_{edar}.docx")
        except Exception as e:
            st.error(f"Error: {e}")

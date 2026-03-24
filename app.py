import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

# Configuración de la página
st.set_page_config(page_title="Generador de Actas - Plantilla Certificación", layout="wide")

# --- BARRA LATERAL: DATOS ---
with st.sidebar:
    st.header("📋 Datos del Informe")
    edar = st.text_input("Nombre de la EDAR", key="edar_name")
    idcoste = st.text_input("IDCOSTE", key="id_coste")
    poblacion = st.text_input("Población", key="pob")
    direccion = st.text_input("Dirección", key="dir")
    provincia = st.text_input("Provincia", key="prov")
    fecha = st.text_input("Fecha instalación", key="fec")
    tecnicos = st.text_input("Técnicos instaladores", key="tec")
    responsable = st.text_input("Responsable Explotación", key="resp")
    
    st.divider()
    if st.button("♻️ REINICIAR TODO", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

st.title("📄 Generador de Acta de Certificación")
st.info("Complete los datos a la izquierda y suba las fotos en las secciones correspondientes.")

# --- CARGA DE FOTOS (Tus 16 selectores originales) ---
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

# --- PROCESO DE GENERACIÓN ---
if st.button("🚀 GENERAR ACTA FINAL", use_container_width=True, type="primary"):
    if not edar:
        st.warning("Por favor, indique el nombre de la EDAR.")
    else:
        try:
            doc = Document()
            
            # 1. LOGOS (Encabezado superior)
            header_table = doc.add_table(rows=1, cols=3)
            header_table.width = Inches(7)
            logos = ["logo_adasa.png", "logo_inelcom.png", "logo_instituciona.png"]
            for i, l_name in enumerate(logos):
                if os.path.exists(l_name):
                    p = header_table.cell(0, i).paragraphs[0]
                    run = p.add_run()
                    run.add_picture(l_name, width=Inches(1.2))
                    if i == 1: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if i == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # 2. TÍTULO Y DATOS GENERALES
            doc.add_heading(f"\n{edar.upper()}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading("ACTA DE CERTIFICACIÓN", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph(f"\nEDAR {edar}\nIDCOSTE: {idcoste}\n{direccion}\nPoblación: {poblacion}\nResponsable: {responsable}\nProvincia: {provincia}\nFecha: {fecha}\nTécnicos: {tecnicos}")

            # 3. CUADRÍCULA DE EQUIPAMIENTO (Vacía con enunciados)
            doc.add_page_break()
            doc.add_heading("IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO", level=1)
            table_eq = doc.add_table(rows=6, cols=3) # Filas vacías para rellenar a mano o después
            table_eq.style = 'Table Grid'
            encabezados = ['EQUIPAMIENTO', 'NÚMERO DE SERIE', 'COORDENADAS']
            for i, texto in enumerate(encabezados):
                table_eq.cell(0, i).text = texto
                table_eq.cell(0, i).paragraphs[0].runs[0].bold = True

            # 4. FOTOS Y GRÁFICAS (Secciones)
            bloques = [
                ("FOTO CARTEL Y PORTADA", [f_portada, f_cartel]),
                ("GRÁFICAS", [f_graficas]),
                ("ALIVIOS", [f_alivio_e1, f_alivio_e2, f_alivio_c1, f_alivio_c2, f_alivio_c3]),
                ("CAUDALÍMETROS", [f_cauda_e1, f_cauda_e2, f_cauda_s1, f_cauda_s2]),
                ("CALIDAD", [f_calid_e1, f_calid_e2, f_calid_s1, f_calid_s2])
            ]

            for titulo, uploaders in bloques:
                fotos_a_poner = [f for up in uploaders if up for f in up]
                if fotos_a_poner:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    for foto in fotos_a_poner:
                        # Insertar foto centrada
                        p_foto = doc.add_paragraph()
                        p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run_f = p_foto.add_run()
                        run_f.add_picture(io.BytesIO(foto.read()), width=Inches(4.5))
                        # Observaciones (según plantilla Albufera)
                        doc.add_paragraph(f"Observaciones: _________________________________________________")

            # 5. CONCLUSIÓN
            doc.add_page_break()
            doc.add_heading("CONCLUSIONES", level=1)
            doc.add_paragraph(f"LA INSTALACIÓN EN EDAR {edar}, QUEDA COMPLETADA CORRECTAMENTE Y EN SERVICIO.")

            # 6. FIRMA
            doc.add_paragraph("\n\n\n")
            doc.add_heading("FIRMA Y VALIDACIÓN", level=1)
            doc.add_paragraph("Esta Asistencia Técnica de Control certifica que la instalación ha sido supervisada y verificada.")
            doc.add_paragraph("\n\nFIRMA:__________________________")

            # Guardar y Descargar
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ Acta generada siguiendo el modelo de Albufera Sur.")
            st.download_button("💾 DESCARGAR ACTA (.docx)", target.getvalue(), f"Acta_{edar}.docx")
            
        except Exception as e:
            st.error(f"Error técnico: {e}")

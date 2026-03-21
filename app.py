import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
from PIL import Image
import numpy as np

st.set_page_config(page_title="Generador Actas ADASA-INELCOM", layout="wide")

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'])
reader = load_ocr()

def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

# --- DATOS BARRA LATERAL ---
with st.sidebar:
    st.header("📋 Datos de la EDAR")
    nombre_edar = st.text_input("Nombre EDAR", "EDAR ALZIRA")
    localidad = st.text_input("Localidad", "Alzira")
    provincia = st.text_input("Provincia", "Valencia")
    idcoste = st.text_input("IDCOSTE", "0017")
    instaladores = st.text_input("Instaladores", "Técnico 1")
    fecha = st.date_input("Fecha Instalación")

# --- INTERFAZ CENTRAL ---
st.title("📄 Generador de Actas Profesionales")
excel_file = st.file_uploader("1. Sube el Excel de Coordenadas", type=['xlsx'])

st.divider()
st.subheader("📸 2. Carga de Fotos por Orden")
col1, col2 = st.columns(2)
with col1:
    foto_puerta = st.file_uploader("🖼️ Foto Puerta (Portada)", accept_multiple_files=True)
    foto_cartel = st.file_uploader("🪧 Foto Cartel Informativo", accept_multiple_files=True)
    fotos_entrada = st.file_uploader("📥 Entrada", accept_multiple_files=True)
with col2:
    fotos_alivio = st.file_uploader("🌊 Alivio", accept_multiple_files=True)
    fotos_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    fotos_graficas = st.file_uploader("📈 Gráficas y Pantallas", accept_multiple_files=True)

if st.button("📝 GENERAR ACTA FINAL"):
    if excel_file:
        with st.spinner("Generando acta profesional..."):
            df = pd.read_excel(excel_file)
            c_serie = get_column(['serie', 'sn', 's/n'], df)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            c_desc = get_column(['desc', 'nombre', 'punto'], df)

            doc = Document()
            
            # --- PÁGINA 1: PORTADA CON LOGO EU ---
            try:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # BUSCA EL LOGO INSTITUCIONAL (EU)
                p.add_run().add_picture('logo_institucional.png', width=Inches(5))
            except: 
                st.warning("No se encontró 'logo_institucional.png' en GitHub")

            doc.add_heading(nombre_edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            if foto_puerta:
                p_p = doc.add_paragraph()
                p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_p.add_run().add_picture(foto_puerta[0], width=Inches(3))

            # Cuadrícula Datos EDAR (Calcada a tu foto)
            table_info = doc.add_table(rows=5, cols=2)
            table_info.style = 'Table Grid'
            datos = [("EDAR", nombre_edar), ("LOCALIDAD", f"{localidad} ({provincia})"), ("IDCOSTE", idcoste), ("INSTALADORES", instaladores), ("FECHA", str(fecha))]
            for i, (k, v) in enumerate(datos):
                table_info.rows[i].cells[0].text = k
                table_info.rows[i].cells[1].text = str(v)

            p_logos = doc.add_paragraph()
            p_logos.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                p_logos.add_run().add_picture('logo_adasa.png', width=Inches(1))
                p_logos.add_run("    ")
                p_logos.add_run().add_picture('logo_inelcom.png', width=Inches(1))
            except: pass

            # --- PÁGINA 2: IDENTIFICACIÓN EQUIPOS ---
            doc.add_page_break()
            doc.add_heading('IDENTIFICACIÓN DEL EQUIPAMIENTO INSTALADO', level=1)
            tbl_equip = doc.add_table(rows=1, cols=3)
            tbl_equip.style = 'Table Grid'
            hdr = tbl_equip.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            # --- SECCIONES ---
            secciones = [("FOTO CARTEL", foto_cartel), ("ENTRADA", fotos_entrada), ("ALIVIO", fotos_alivio), ("SALIDA", fotos_salida), ("GRÁFICAS", fotos_graficas)]
            
            for titulo, lista in secciones:
                if lista:
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2)
                    for i, foto in enumerate(lista):
                        if i % 2 == 0: cells = grid.add_row().cells
                        cell = cells[i % 2]
                        
                        sn, coor = "No detectado", "N/A"
                        # IA más agresiva
                        if titulo not in ["FOTO CARTEL", "GRÁFICAS"]:
                            img = Image.open(foto)
                            txt = " ".join(reader.readtext(np.array(img), detail=0)).replace(" ", "").replace("-", "")
                            if c_serie:
                                for _, row in df.iterrows():
                                    s_ex = str(row[c_serie]).replace(" ", "").replace("-", "")
                                    if s_ex in txt and len(s_ex) > 4:
                                        sn, coor = str(row[c_serie]), str(row[c_coord])
                                        d = str(row[c_desc]) if c_desc else titulo
                                        r = tbl_equip.add_row().cells
                                        r[0].text, r[1].text, r[2].text = d, sn, coor
                                        break
                        
                        cell.paragraphs[0].add_run().add_picture(foto, width=Inches(2.5))
                        p_d = cell.add_paragraph(f"{titulo} S/N: {sn}")
                        p_d.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # --- CONCLUSIONES Y FIRMA ---
            doc.add_page_break()
            title_c = doc.add_heading('CONCLUSIONES', level=1)
            title_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_c.runs[0].font.color.rgb = RGBColor(112, 48, 160)
            
            tbl_c = doc.add_table(rows=1, cols=1)
            tbl_c.style = 'Table Grid'
            tbl_c.rows[0].cells[0].text = f"LA INSTALACIÓN EN {nombre_edar}, QUEDA COMPLETADA CORRECTAMENTE Y EN SERVICIO."

            doc.add_paragraph("\n")
            title_f = doc.add_heading('FIRMA Y VALIDACIÓN.', level=1)
            title_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_f.runs[0].font.color.rgb = RGBColor(112, 48, 160)
            doc.add_paragraph("Esta Asistencia Técnica de Control, certifica que la instalación ha sido supervisada y verificada según normativa.")
            
            tbl_f = doc.add_table(rows=2, cols=1)
            tbl_f.style = 'Table Grid'
            tbl_f.rows[0].cells[0].text = "FIRMA:"
            tbl_f.rows[1].height = Inches(2)

            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta lista!")
            st.download_button("📥 DESCARGAR ACTA", target.getvalue(), f"Acta_{nombre_edar}.docx")

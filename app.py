import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    st.info("1. Rellena los datos.\n2. Carga las fotos.\n3. Genera el Word.")

# --- PANEL CENTRAL ---
st.title("📄 Generador de Actas ADASA & INELCOM")
st.caption("Versión Optimizada: Datos manuales y leyendas fijas")

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
            
            # --- PORTADA Y LOGOS ---
            if os.path.exists('logo_instituciona.png'):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture('logo_instituciona.png', width=Inches(5))

            doc.add_heading(edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            if f_portada:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(f_portada[0], width=Inches(4.5))

            # --- TABLA DE DATOS ---
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

            # --- SECCIONES Y LEYENDAS ---
            # Definimos las secciones y qué texto debe llevar cada una
            secciones = [
                ("CARTEL INFORMATIVO", f_cartel, "Fotografía del cartel de subvenciones."),
                ("EQUIPAMIENTO EN ENTRADA", f_entrada, "Fotos ubicación equipos y sondas instalados"),
                ("EQUIPAMIENTO EN ALIVIO", f_alivio, "Fotos ubicación equipos y sondas instalados"),
                ("EQUIPAMIENTO EN SALIDA", f_salida, "Fotos ubicación equipos y sondas instalados"),
                ("GRÁFICAS Y CONCLUSIONES", f_graficas, "") # Sin leyenda específica
            ]

            for titulo, lista, leyenda in secciones:
                if lista:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2)
                    
                    for i, foto in enumerate(lista):
                        if i % 2 == 0:
                            row_cells = grid.add_row().cells
                        cell = row_cells[i % 2]
                        
                        # Imagen
                        run_img = cell.paragraphs[0].add_run()
                        run_img.add_picture(foto, width=Inches(2.5))
                        
                        # Leyenda (si existe para esa sección)
                        if leyenda:
                            p_ley = cell.add_paragraph(leyenda)
                            p_ley.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run_ley = p_ley.runs[0]
                            run_ley.font.size = Pt(9)
                            run_ley.font.bold = True

            # --- FIRMA Y LOGOS FINALES ---
            doc.add_paragraph("\n\n")
            p_final = doc.add_paragraph()
            p_final.alignment = WD_ALIGN_PARAGRAPH.CENTER
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

            # Descarga
            target = io.BytesIO()
            doc.save(target)
            st.success(f"✅ Acta de {edar} lista.")
            st.download_button("📥 DESCARGAR WORD", target.getvalue(), f"Acta_{edar.replace(' ', '_')}.docx")

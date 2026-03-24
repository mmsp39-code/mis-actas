import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

# Configuración de página
st.set_page_config(page_title="Generador de Actas Profesional", layout="wide")

# --- BARRA LATERAL ---
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

# --- PANEL CENTRAL ---
st.title("📄 Generador de Actas ADASA & INELCOM")
st.divider()

# --- CARGA DE FOTOS ---
col1, col2 = st.columns(2)
with col1:
    f_portada = st.file_uploader("🖼️ Foto Portada", accept_multiple_files=True)
    f_cartel = st.file_uploader("🪧 Cartel Informativo", accept_multiple_files=True)
    f_entrada = st.file_uploader("📥 Entrada / Equipamiento", accept_multiple_files=True)
with col2:
    f_alivio = st.file_uploader("🌊 Alivio (EDAR/Colector)", accept_multiple_files=True)
    f_salida = st.file_uploader("📤 Salida", accept_multiple_files=True)
    f_graficas = st.file_uploader("📈 Gráficas / Conclusiones", accept_multiple_files=True)

if st.button("🚀 GENERAR ACTA PROFESIONAL"):
    with st.spinner("Generando Word..."):
        doc = Document()
        
        # 1. Logo Institucional
        if os.path.exists('logo_instituciona.png'):
            p_inst = doc.add_paragraph()
            p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_inst.add_run().add_picture('logo_instituciona.png', width=Inches(5))

        doc.add_heading(edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

        if f_portada:
            p_portada = doc.add_paragraph()
            p_portada.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_portada.add_run().add_picture(f_portada[0], width=Inches(4.5))

        # 2. Tabla de Datos
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

        # 3. Logos ADASA e INELCOM (Aquí estaba el error, ahora corregido)
        p_logos = doc.add_paragraph()
        p_logos.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if os.path.exists('logo_adasa.png'):
            p_logos.add_run().add_picture('logo_adasa.png', width=Inches(1.2))
        
        p_logos.add_run("      ") # Espacio simple entre logos
        
        if os.path.exists('logo_inelcom.png'):
            p_logos.add_run().add_picture('logo_inelcom.png', width=Inches(1.2))

        # 4. Secciones de fotos con leyendas finales
        secciones = [
            ("CARTEL INFORMATIVO", f_cartel, "Fotografía del cartel de subvenciones."),
            ("EQUIPAMIENTO EN ENTRADA", f_entrada, "Fotos instalación y ubicación de equipos y sondas instalados."),
            ("EQUIPAMIENTO EN ALIVIO", f_alivio, "Fotos instalación y ubicación de equipos y sondas instalados."),
            ("EQUIPAMIENTO EN SALIDA", f_salida, "Fotos instalación y ubicación de equipos y sondas instalados."),
            ("GRÁFICAS Y CONCLUSIONES", f_graficas, "")
        ]

        for titulo, lista, leyenda in secciones:
            if lista:
                doc.add_page_break()
                doc.add_heading(titulo, level=1)
                grid = doc.add_table(rows=0, cols=2)
                for i, foto in enumerate(lista):
                    if i % 2 == 0: row_cells = grid.add_row().cells
                    row_cells[i % 2].paragraphs[0].add_run().add_picture(foto, width=Inches(2.5))
                
                if leyenda:
                    p_ley = doc.add_paragraph()
                    run_ley = p_ley.add_run(f"\n{leyenda}")
                    run_ley.font.size = Pt(10)
                    run_ley.font.bold = True
                    run_ley.italic = True

        # 5. Firma final
        doc.add_paragraph("\n\n")
        tbl_f = doc.add_table(rows=2, cols=1)
        tbl_f.style = 'Table Grid'
        tbl_f.rows[0].cells[0].text = "FIRMA Y VALIDACIÓN:"
        tbl_f.rows[1].height = Inches(1.2)

        target = io.BytesIO()
        doc.save(target)
        st.success("✅ Acta generada correctamente")
        st.download_button("📥 DESCARGAR WORD", target.getvalue(), f"Acta_{edar}.docx")

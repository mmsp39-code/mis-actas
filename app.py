import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

# Configuración para Pantalla Completa (Escritorio)
st.set_page_config(page_title="Generador de Actas Profesional", layout="wide")

# --- BARRA LATERAL (DATOS FIJOS) ---
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
    st.info("💡 Consejo: Puedes arrastrar varias fotos a la vez desde tu carpeta directamente a los cuadros de carga.")

# --- PANEL PRINCIPAL ---
st.title("📄 Generador de Actas ADASA & INELCOM")
st.subheader("Panel de Control de Ingeniería - Modo Escritorio")
st.divider()

# --- CARGA DE FOTOS ORGANIZADA (VISTA WEB ANCHA) ---
st.write("### 📸 Gestión de Archivos por Sección")
col_gen, col_aliv, col_caud = st.columns([1, 1, 1])

with col_gen:
    st.error("#### 🖼️ GENERALES")
    f_portada = st.file_uploader("Foto Portada (Principal)", accept_multiple_files=True, key="portada")
    f_cartel = st.file_uploader("Cartel Informativo", accept_multiple_files=True, key="cartel")
    f_graficas = st.file_uploader("Gráficas y Conclusiones", accept_multiple_files=True, key="graficas")

with col_aliv:
    st.warning("#### 🌊 SECCIÓN ALIVIOS")
    f_alivio_e1 = st.file_uploader("Alivio EDAR 1", accept_multiple_files=True)
    f_alivio_e2 = st.file_uploader("Alivio EDAR 2", accept_multiple_files=True)
    f_alivio_c1 = st.file_uploader("Alivio Colector 1", accept_multiple_files=True)
    f_alivio_c2 = st.file_uploader("Alivio Colector 2", accept_multiple_files=True)
    f_alivio_c3 = st.file_uploader("Alivio Colector 3", accept_multiple_files=True)

with col_caud:
    st.success("#### 📥 CAUDALÍMETROS")
    f_cauda_e1 = st.file_uploader("Entrada 1", accept_multiple_files=True)
    f_cauda_e2 = st.file_uploader("Entrada 2", accept_multiple_files=True)
    f_cauda_s1 = st.file_uploader("Salida 1", accept_multiple_files=True)
    f_cauda_s2 = st.file_uploader("Salida 2", accept_multiple_files=True)

st.divider()

# --- LÓGICA DE GENERACIÓN ---
if st.button("🚀 GENERAR DOCUMENTO FINAL (WORD)", use_container_width=True):
    if not edar:
        st.error("Debes indicar el nombre de la EDAR en la barra lateral.")
    else:
        with st.spinner("Procesando imágenes y maquetando el Word..."):
            doc = Document()
            
            # --- PORTADA ---
            if os.path.exists('logo_instituciona.png'):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture('logo_instituciona.png', width=Inches(5))

            doc.add_heading(edar, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('ACTA DE CERTIFICACIÓN', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER

            if f_portada:
                p_portada = doc.add_paragraph()
                p_portada.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_portada.add_run().add_picture(f_portada[0], width=Inches(4.5))

            # TABLA DE DATOS
            datos = [
                ("EDAR", edar), ("IDCOSTE", idcoste), ("Población", poblacion),
                ("Dirección", direccion), ("Provincia", provincia),
                ("Fecha instalación", fecha), ("Técnicos instaladores", tecnicos),
                ("Responsable Explotación", responsable)
            ]
            tbl = doc.add_table(rows=len(datos), cols=2)
            tbl.style = 'Table Grid'
            for i, (k, v) in enumerate(datos):
                tbl.rows[i].cells[0].text, tbl.rows[i].cells[1].text = k, str(v)

            # LOGOS PORTADA (Corrección de espacio y centrado)
            p_logos = doc.add_paragraph()
            p_logos.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if os.path.exists('logo_adasa.png'):
                p_logos.add_run().add_picture('logo_adasa.png', width=Inches(1.2))
            p_logos.add_run("      ") 
            if os.path.exists('logo_inelcom.png'):
                p_logos.add_run().add_picture('logo_inelcom.png', width=Inches(1.2))

            # --- SECCIONES DINÁMICAS ---
            leyenda_eq = "Fotos instalación y ubicación de equipos y sondas instalados."
            
            secciones = [
                ("CARTEL INFORMATIVO", f_cartel, "Fotografía del cartel de subvenciones."),
                ("ALIVIO EDAR 1", f_alivio_e1, leyenda_eq),
                ("ALIVIO EDAR 2", f_alivio_e2, leyenda_eq),
                ("ALIVIO COLECTOR 1", f_alivio_c1, leyenda_eq),
                ("ALIVIO COLECTOR 2", f_alivio_c2, leyenda_eq),
                ("ALIVIO COLECTOR 3", f_alivio_c3, leyenda_eq),
                ("CAUDALÍMETRO ENTRADA 1", f_cauda_e1, leyenda_eq),
                ("CAUDALÍMETRO ENTRADA 2", f_cauda_e2, leyenda_eq),
                ("CAUDALÍMETRO SALIDA 1", f_cauda_s1, leyenda_eq),
                ("CAUDALÍMETRO SALIDA 2", f_cauda_s2, leyenda_eq),
                ("GRÁFICAS Y CONCLUSIONES", f_graficas, "")
            ]

            for titulo, fotos, leyenda in secciones:
                if fotos:
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    grid = doc.add_table(rows=0, cols=2)
                    for i, f in enumerate(fotos):
                        if i % 2 == 0: row = grid.add_row().cells
                        row[i % 2].paragraphs[0].add_run().add_picture(f, width=Inches(2.5))
                    
                    if leyenda:
                        p_ley = doc.add_paragraph()
                        r_ley = p_ley.add_run(f"\n{leyenda}")
                        r_ley.font.size = Pt(10)
                        r_ley.font.bold = True
                        r_ley.italic = True

            # FIRMA
            doc.add_paragraph("\n\n")
            tbl_f = doc.add_table(rows=2, cols=1)
            tbl_f.style = 'Table Grid'
            tbl_f.rows[0].cells[0].text = "FIRMA Y VALIDACIÓN:"
            tbl_f.rows[1].height = Inches(1.2)

            # DESCARGA
            target = io.BytesIO()
            doc.save(target)
            st.success("✅ ¡Acta generada! Haz clic en el botón de abajo para guardarla.")
            st.download_button(label="💾 GUARDAR ARCHIVO WORD", 
                               data=target.getvalue(), 
                               file_name=f"Acta_{edar}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)

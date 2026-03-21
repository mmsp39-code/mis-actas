import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import easyocr
import io
from PIL import Image
import numpy as np

# Configuración profesional de la página
st.set_page_config(page_title="Generador de Actas ADASA/INELCOM - Profesional", layout="wide")

# Cargamos la IA para leer fotos (solo una vez)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'])

reader = load_ocr()

# Función para encontrar columnas en el Excel de forma flexible
def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in str(col).lower() for key in keywords):
            return col
    return None

# --- BARRA LATERAL (DATOS DEL ACTA SIGUIENDO TU EJEMPLO) ---
with st.sidebar:
    st.image('logo_adasa.png', width=100) # (Opcional, si tienes el logo en GitHub)
    st.header("📍 Ubicación y Datos Generales")
    nombre_edar = st.text_input("Nombre de la EDAR", "ALZIRA")
    localidad = st.text_input("Localidad", "Alzira")
    provincia = st.text_input("Provincia", "Valencia")
    idcoste = st.text_input("IDCOSTE", "1234")

    st.header("👷 Personal y Fecha")
    instaladores = st.text_input("Nombre Instaladores", "Técnico 1, Técnico 2")
    fecha = st.date_input("Fecha de Instalación")
    st.divider()
    
# --- CUERPO CENTRAL ---
st.title("📄 Generador de Actas de Certificación (Multi-Punto)")
st.markdown("Sube el Excel con las coordenadas y todas las fotos organizadas en los apartados correspondientes.")

# --- 1. SUBIR EXCEL ---
excel_file = st.file_uploader("📂 1. Sube el Excel de Coordenadas (Debe tener S/N, Coordenadas y Descripción)", type=['xlsx'])

st.divider()

# --- 2. CAJONES DE CARGA POR APARTADO (TÚ CLASIFICAS) ---
st.subheader("📸 2. Carga de Fotos por Ubicación")
st.info("Todo lo que subas a cada cajón se agrupará bajo ese título en el Word.")

col1, col2 = st.columns(2)
with col1:
    fotos_portada = st.file_uploader("🖼️ Portada / Cartel / Puerta / Entorno (Sin S/N)", accept_multiple_files=True, key="p1")
    fotos_entrada_1 = st.file_uploader("📥 Entrada 1 / General", accept_multiple_files=True, key="e1")
    fotos_entrada_2 = st.file_uploader("📥 Entrada 2 (Opcional)", accept_multiple_files=True, key="e2")
    fotos_alivio_1 = st.file_uploader("🌊 Alivio Colector 1", accept_multiple_files=True, key="al1")
    fotos_alivio_2 = st.file_uploader("🌊 Alivio Colector 2", accept_multiple_files=True, key="al2")
with col2:
    fotos_salida_1 = st.file_uploader("📤 Salida 1 / General", accept_multiple_files=True, key="s1")
    fotos_salida_2 = st.file_uploader("📤 Salida 2 (Opcional)", accept_multiple_files=True, key="s2")
    fotos_caudalimetro = st.file_uploader("📉 Caudalímetro y Gráficas", accept_multiple_files=True, key="cd")
    fotos_graficas = st.file_uploader("📈 Pantallas / Gráficas", accept_multiple_files=True, key="gf")

# --- 3. GENERACIÓN ---
if st.button("🚀 GENERAR ACTA PROFESIONAL"):
    if excel_file:
        with st.spinner("Generando portada, recopilando S/N de todas las fotos y cruzando datos..."):
            
            # A. Leer Excel
            df = pd.read_excel(excel_file)
            c_serie = get_column(['serie', 'sn', 's/n'], df)
            c_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            c_desc = get_column(['desc', 'nombre', 'punto'], df) # Buscamos la columna "Entrada 1", etc.
            
            # B. Crear Word
            doc = Document()
            
            # 1. PORTADA CON TABLA Y LOGO (SIGUIENDO TU EJEMPLO)
            try:
                # Logo Institucional centrado
                p_logo = doc.add_paragraph()
                p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_logo.add_run().add_picture('logo_institucional.png', width=Inches(6.5))
            except: pass
            
            # Título principal
            title = doc.add_heading(f'ACTA DE CERTIFICACIÓN DE LA INSTALACIÓN', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"Proyecto: {nombre_edar}", style='Subtitle').alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("") # Espacio
            
            # Tabla de Datos Generales (Ídem a tu captura)
            doc.add_heading('DATOS GENERALES', level=1)
            tbl_info = doc.add_table(rows=5, cols=2)
            tbl_info.style = 'Table Grid'
            
            # Rellenar con los datos laterales
            tbl_datos = [
                ("EDAR / UBICACIÓN", nombre_edar),
                ("LOCALIDAD (PROVINCIA)", f"{localidad} ({provincia})"),
                ("IDCOSTE", idcoste),
                ("TECNICOS DE INSTALACIÓN", instaladores),
                ("FECHA INSTALACIÓN", str(fecha))
            ]
            for i, (k, v) in enumerate(tbl_datos):
                tbl_info.rows[i].cells[0].text = k
                tbl_info.rows[i].cells[1].text = v
                # Formato negrita para las claves
                tbl_info.rows[i].cells[0].paragraphs[0].runs[0].bold = True
                
            doc.add_paragraph("") # Espacio
            
            # 2. TABLA TÉCNICA RECOPILATORIA DE S/N Y COORDENADAS (LO QUE PEDÍAS)
            doc.add_heading('RESUMEN DE EQUIPAMIENTO INSTALADO', level=1)
            doc.add_paragraph("Recopilación automática de los números de serie detectados en las fotos de campo cruzados con el Excel de coordenadas.")
            
            # Tabla vacía (Punto | Nº Serie | Coordenadas)
            tbl_equip = doc.add_table(rows=1, cols=3)
            tbl_equip.style = 'Table Grid'
            hdr = tbl_equip.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'PUNTO MEDIDA', 'Nº SERIE', 'COORDENADAS (WGS84)'
            for cell in hdr: cell.paragraphs[0].runs[0].bold = True
            
            # Diccionario para guardar los S/N únicos y sus datos
            equipos_detectados = {}

            # C. ESTRUCTURA DE SECCIONES (TÍTULO WORD, FOTOS WEB)
            secciones_fotos = [
                ("FOTOS DE PORTADA / CARTEL Y ENTORNO", fotos_portada, False), # False = No buscar S/N
                ("INSTALACIÓN EN ENTRADA 1 / GENERAL", fotos_entrada_1, True),
                ("INSTALACIÓN EN ENTRADA 2", fotos_entrada_2, True),
                ("ALIVIO COLECTOR 1", fotos_alivio_1, True),
                ("ALIVIO COLECTOR 2", fotos_alivio_2, True),
                ("PUNTO DE SALIDA 1 / GENERAL", fotos_salida_1, True),
                ("PUNTO DE SALIDA 2", fotos_salida_2, True),
                ("CAUDALÍMETRO Y PANTALLAS", fotos_caudalimetro, True),
                ("GRÁFICAS Y DATOS", fotos_graficas, False)
            ]

            # D. PROCESAR FOTOS Y RECOPILAR DATOS
            for titulo, fotos_seccion, buscar_sn in secciones_fotos:
                if fotos_seccion:
                    # Página nueva por sección importante
                    doc.add_page_break()
                    doc.add_heading(titulo, level=1)
                    
                    # Para optimizar espacio, usamos tablas ocultas de 2x2 para las fotos de campo
                    if buscar_sn:
                        # Creamos un párrafo para ir metiendo las fotos de 2 en 2
                        current_paragraph = doc.add_paragraph()
                        current_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        count_in_row = 0
                        
                        for i, foto in enumerate(fotos_seccion):
                            
                            # IA Lee S/N
                            sn_foto = "No detectado"
                            coord_foto = "Ver Excel"
                            desc_foto = titulo # Descripción por defecto (la del cajón)
                            
                            img = Image.open(foto)
                            # Redimensionamos para que la IA vaya rápido y el Word no pese 100MB
                            img.thumbnail((1200, 1200))
                            texto_ia = " ".join(reader.readtext(np.array(img), detail=0))
                            
                            # Buscar en el Excel
                            if c_serie and c_coord and buscar_sn:
                                for index_ex, row_ex in df.iterrows():
                                    sn_ex = str(row_ex[c_serie]).strip()
                                    # Asegurarnos de que el S/N es largo y está en la foto
                                    if len(sn_ex) > 4 and sn_ex in texto_ia:
                                        sn_foto = sn_ex
                                        # Guardamos el S/N único para la tabla resumen
                                        equipos_detectados[sn_foto] = (titulo, str(row_ex[c_coord]) if c_coord else "N/A")
                                        # Datos para poner debajo de la foto
                                        desc_foto = f"{titulo} ({str(row_ex[c_desc]) if c_desc else titulo})"
                                        coord_foto = str(row_ex[c_coord])
                                        break
                                        
                            # Insertar foto y explicación (Más pequeñas, 2 por fila si cabe)
                            
                            # Si es la primera de la fila, crea un nuevo párrafo centrado
                            if count_in_row == 0:
                                p_images = doc.add_paragraph()
                                p_images.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            else:
                                # Añadimos espacio entre fotos
                                p_images.add_run("   ")

                            # Añadimos la foto (Tamaño reducido para que quepan varias)
                            run = p_images.add_run()
                            run.add_picture(foto, width=Inches(3.1))
                            
                            # Párrafo explicativo debajo de la foto (Texto pequeño para no ocupar espacio)
                            # Para ponerlo debajo, necesitamos un párrafo nuevo centrado para cada descripción
                            p_desc = doc.add_paragraph()
                            p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            r_desc = p_desc.add_run(f"{desc_foto}")
                            r_desc.font.size = Pt(9)
                            r_desc.bold = True
                            
                            r_ex = p_desc.add_run(f" S/N: {sn_foto}, Coord: {coord_foto}")
                            r_ex.font.size = Pt(8)
                            r_ex.font.color.rgb = RGBColor(100, 100, 100) # Gris oscuro
                            
                            # Lógica para controlar las fotos por fila (intentaremos 2, pero Word es complejo)
                            # Si es la segunda foto, forzamos un espacio antes de la siguiente fila
                            if count_in_row == 1:
                                doc.add_paragraph("")
                                count_in_row = 0
                            else:
                                count_in_row += 1

                    else:
                        # Sección de Portada / Gráficas: Sin S/N, fotos a 2 por fila
                        count_in_row_p = 0
                        
                        for foto_p in fotos_seccion:
                            
                            if count_in_row_p == 0:
                                p_imgs_p = doc.add_paragraph()
                                p_imgs_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            else:
                                p_imgs_p.add_run("   ")
                                
                            run_p = p_imgs_p.add_run()
                            run_p.add_picture(foto_p, width=Inches(3.1))
                            
                            # Título sencillo debajo
                            p_desc_p = doc.add_paragraph()
                            p_desc_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r_p = p_desc_p.add_run(titulo)
                            r_p.font.size = Pt(9)
                            r_p.font.color.rgb = RGBColor(100, 100, 100)
                            
                            if count_in_row_p == 1:
                                doc.add_paragraph("")
                                count_in_row_p = 0
                            else:
                                count_in_row_p += 1

            # E. RELLENAR LA TABLA RESUMEN DE EQUIPOS (ORDENADA)
            # Rellenamos la tabla con los S/N únicos que hemos detectado
            # (Lo hacemos al final para tener la lista completa de todas las secciones)
            
            # Ordenamos los S/N por Punto de Medida (Alivio, Entrada...)
            for sn, (punto, coord) in sorted(equipos_detectados.items(), key=lambda x: x[1][0]):
                r = tbl_equip.add_row().cells
                r[0].text = punto
                r[1].text = sn
                r[2].text = coord
                r[0].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r[1].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r[2].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 6. Logos Finales (Pie de Página Fijo)
            try:
                footer = doc.sections[0].footer.paragraphs[0]
                footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_f = footer.add_run()
                run_f.add_picture('logo_adasa.png', width=Inches(1.2))
                run_f.add_run("   ")
                run_f.add_picture('logo_inelcom.png', width=Inches(1.2))
            except: pass

            # G. Preparar descarga
            target = io.BytesIO()
            doc.save(target)
            
            st.success("✅ Acta profesional generada con éxito con portada, tabla resumen y fotos de campo detalladas.")
            st.download_button(
                label="📥 PINCHA AQUÍ PARA DESCARGAR EL WORD",
                data=target.getvalue(),
                file_name=f"Acta_{nombre_edar}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("Es obligatorio subir el Excel para las coordenadas.")

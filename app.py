import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import easyocr
import io
from PIL import Image
import numpy as np

# Configuración de la página web
st.set_page_config(page_title="Generador de Actas ADASA/INELCOM", layout="wide")

# Cargamos la IA para leer fotos (solo una vez)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['es'])

reader = load_ocr()

# Lógica para encontrar columnas en el Excel de forma flexible
def get_column(keywords, df):
    for col in df.columns:
        if any(key.lower() in col.lower() for key in keywords):
            return col
    return None

# --- BARRA LATERAL (DATOS DEL ACTA) ---
st.sidebar.header("📍 Ubicación")
nombre_edar = st.sidebar.text_input("Nombre de la EDAR", "EJEMPLO EDAR")
localidad = st.sidebar.text_input("Localidad", "Valencia")
provincia = st.sidebar.text_input("Provincia", "Valencia")
idcoste = st.sidebar.text_input("IDCOSTE", "0017")

st.sidebar.header("👷 Personal y Fecha")
instaladores = st.sidebar.text_input("Nombre Instaladores", "Técnico 1, Técnico 2")
fecha = st.sidebar.date_input("Fecha de Instalación")
responsable = st.sidebar.text_input("Responsable Explotación", "Nombre Responsable")

# --- CUERPO CENTRAL ---
st.title("📄 Generador de Actas de Certificación")
st.markdown("Sube el Excel con las coordenadas y todas las fotos de los sensores. La IA las leerá y montará el acta automáticamente.")

excel_file = st.file_uploader("1. Sube el Excel de Coordenadas", type=['xlsx'])
fotos = st.file_uploader("2. Sube las fotos (puedes arrastrar varias)", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

if st.button("🚀 GENERAR ACTA PROFESIONAL"):
    if excel_file and fotos:
        with st.spinner("Procesando fotos y cruzando datos... esto puede tardar unos minutos."):
            # 1. Leer Excel
            df = pd.read_excel(excel_file)
            col_serie = get_column(['serie', 'sn', 's/n', 'numero'], df)
            col_coord = get_column(['coord', 'gps', 'ubicacion'], df)
            
            # 2. Crear Word
            doc = Document()
            
            # Encabezado con Logos (Fijos)
            try:
                header = doc.sections[0].header
                header.paragraphs[0].add_run().add_picture('logo_institucional.png', width=Inches(6))
            except: pass

            # Título y Datos Generales
            doc.add_heading(f'ACTA DE CERTIFICACIÓN - {nombre_edar}', 0)
            
            doc.add_heading('DATOS DE LA INSTALACIÓN', level=1)
            table_info = doc.add_table(rows=5, cols=2)
            table_info.style = 'Table Grid'
            datos_gen = [
                ("EDAR", nombre_edar),
                ("UBICACIÓN", f"{localidad} ({provincia})"),
                ("IDCOSTE", idcoste),
                ("INSTALADORES", instaladores),
                ("FECHA", str(fecha))
            ]
            for i, (clave, valor) in enumerate(datos_gen):
                table_info.rows[i].cells[0].text = clave
                table_info.rows[i].cells[1].text = valor

            # 3. Tabla Técnica (Vacía de momento)
            doc.add_heading('EQUIPAMIENTO INSTALADO', level=1)
            table_equip = doc.add_table(rows=1, cols=3)
            table_equip.style = 'Table Grid'
            hdr = table_equip.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'EQUIPAMIENTO', 'Nº SERIE', 'COORDENADAS'

            # 4. Procesar Fotos subidas
            for foto_file in fotos:
                # Convertir foto para la IA
                image = Image.open(foto_file)
                img_array = np.array(image)
                
                # IA lee el número de serie
                result = reader.readtext(img_array, detail=0)
                texto_ia = " ".join(result)
                
                # Buscar número de serie del Excel en el texto de la foto
                sn_encontrado = "No detectado"
                coord_encontrada = "N/A"
                
                if col_serie and col_coord:
                    for s in df[col_serie].astype(str):
                        if s in texto_ia:
                            sn_encontrado = s
                            coord_encontrada = str(df.loc[df[col_serie] == s, col_coord].values[0])
                            
                            # Añadir fila a la tabla principal
                            r = table_equip.add_row().cells
                            r[0].text, r[1].text, r[2].text = 'SENSOR', sn_encontrado, coord_encontrada
                            break

                # Insertar foto en el Word
                doc.add_heading(f"Foto de campo - S/N: {sn_encontrado}", level=2)
                doc.add_picture(foto_file, width=Inches(4))
                doc.add_paragraph(f"Texto detectado por IA: {texto_ia}")

            # 5. Logos de ADASA / INELCOM al final
            try:
                footer = doc.sections[0].footer.paragraphs[0].add_run()
                footer.add_picture('logo_adasa.png', width=Inches(1.2))
                footer.add_run("    ")
                footer.add_picture('logo_inelcom.png', width=Inches(1.2))
            except: pass

            # 6. Preparar descarga
            target = io.BytesIO()
            doc.save(target)
            
            st.success("✅ ¡Acta generada con éxito con logos, fotos y coordenadas!")
            st.download_button(
                label="📥 PINCHA AQUÍ PARA DESCARGAR EL WORD",
                data=target.getvalue(),
                file_name=f"Acta_{nombre_edar}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("Por favor, sube el Excel y las fotos primero.")

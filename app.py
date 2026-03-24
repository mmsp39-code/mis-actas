import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

st.set_page_config(page_title="Generador de Actas Ultra-Compacto", layout="wide")

# --- BARRA LATERAL ---
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

# --- SELECTORES (16 ACCESOS ORIGINALES) ---
st.title("📄 Generador de Acta: Máximo Aprovechamiento")
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.error("#### 🖼️ GENERALES")
    f_portada = st.file_uploader("Portada", accept_multiple_files=False, key="u1")
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
    st.success("#### 📥 CAUDAL")
    f_cauda_e1 = st.file_uploader("Ent. 1", accept_multiple_files=True, key="u9")
    f_cauda_e2 = st.file_uploader("Ent. 2", accept_multiple_files=True, key="u10")
    f_cauda_s1 = st.file_uploader("Sal. 1", accept_multiple_files=True, key="u11")
    f_cauda_s2 = st.file_uploader("

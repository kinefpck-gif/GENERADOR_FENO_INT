import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- 1. EXTRACCIÓN DE ALTA PRECISIÓN ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina = doc_pdf[0]
    texto_completo = pagina.get_text()

    # RECORTE UNIFICADO: Captura el área exacta de ambas curvas
    # Ajustado para Sunvou: [x0, y0, x1, y1]
    # Este rectángulo captura ambos cuadros de gráficos sin bordes negros innecesarios
    rect_graficos = fitz.Rect(42, 435, 555, 575) 
    pix = pagina.get_pixmap(clip=rect_graficos, matrix=fitz.Matrix(4, 4)) # 4x zoom para nitidez

    def extraer_valor(patron, texto):
        match = re.search(patron, texto, re.IGNORECASE)
        return match.group(1).strip() if match else "---"

    return {
        "feno50": extraer_valor(r"FeN[O0]50[:\s]*(\d+)", texto_completo),
        "temp": extraer_valor(r"Temperatura[:\s]*([\d\.,]+)", texto_completo),
        "presion": extraer_valor(r"Presión[:\s]*([\d\.,]+)", texto_completo),
        "flujo": extraer_valor(r"Tasa de flujo[:\s]*([\d\.,]+)", texto_completo),
        "img_final": pix.tobytes("png")
    }

# --- 2. MOTOR DE REEMPLAZO (MAYÚSCULAS Y PROCEDENCIA) ---
def generar_word_pro(datos_m, datos_e, plantilla_path):
    if not os.path.exists(plantilla_path): return None
    doc = Document(plantilla_path)
    
    # Mapeo con transformación a MAYÚSCULAS para datos personales
    reemplazos = {
        "{{NOMBRE}}": str(datos_m['nombre']).upper(),
        "{{APELLIDOS}}": str(datos_m['apellidos']).upper(),
        "{{RUT}}": str(datos_m['rut']).upper(),
        "{{GENERO}}": str(datos_m['genero']).upper(),
        "{{F. NACIMIENTO}}": str(datos_m['f_nac']).upper(),
        "{{EDAD}}": str(datos_m['edad']).upper(),
        "{{ALTURA}}": str(datos_m['altura']).upper(),
        "{{PESO}}": str(datos_m['peso']).upper(),
        "{{PROCEDENCIA}}": str(datos_m['procedencia']).upper(),
        "{{MEDICO}}": str(datos_m['medico']).upper(),
        "{{OPERADOR}}": str(datos_m['operador']).upper(),
        "{{FECHA_EXAMEN}}": str(datos_m['fecha_ex']).upper(),
        "{{Temperatura}}": datos_e['temp'],
        "{{Presion}}": datos_e['presion'],
        "{{Tasa de flujo}}": datos_e['flujo'],
        "{{FeNO50}}": datos_e['feno50']
    }

    def procesar_parrafo(p):
        # Reemplazo de etiquetas de texto
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, v)
        
        # Inserción de la imagen unificada en el marcador
        if "CURVA_EXHALA" in p.text:
            p.text = p.text.replace("CURVA_EXHALA", "")
            run = p.add_run()
            # 5.2 pulgadas es el ancho ideal para que luzca como en tu Informe2
            run.add_picture(io.BytesIO(datos_e['img_final']), width=Inches(5.2))

    # Aplicar a párrafos y tablas
    for p in doc.paragraphs: procesar_parrafo(p)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs: procesar_parrafo(p)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 3. INTERFAZ ---
st.set_page_config(page_title="INT - FeNO Pro", layout="wide")
st.title("🏥 Generador de Informes FeNO (Nivel Experto)")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Información del Paciente")
    d_m = {
        'nombre': st.text_input("Nombre"),
        'apellidos': st.text_input("Apellidos"),
        'rut': st.text_input("RUT"),
        'genero': st.selectbox("Género", ["HOMBRE", "MUJER"]),
        'f_nac': st.text_input("F. Nacimiento"),
        'edad': st.text_input("Edad"),
        'altura': st.text_input("Altura"),
        'peso': st.text_input("Peso"),
        'procedencia': st.text_input("Procedencia", value="POLI"),
        'medico': st.text_input("Médico"),
        'operador': st.text_input("Operador", value="TM JORGE ESPINOZA"),
        'fecha_ex': st.date_input("Fecha Examen").strftime("%d/%m/%Y")
    }

with col2:
    st.subheader("Carga de Datos")
    pdf_file = st.file_uploader("Subir PDF Sunvou", type="pdf")
    tipo = st.radio("Seleccionar Plantilla:", ["FeNO 50", "FeNO 50-200"])

if st.button("✨ Generar Informe"):
    if pdf_file and d_m['nombre']:
        with st.spinner("Procesando gráficas y datos..."):
            datos_extraidos = procesar_pdf_sunvou(pdf_file)
            
            # Control visual antes de descargar
            st.image(datos_extraidos['img_final'], caption="Área de Gráficos Extraída", width=700)
            
            ruta_plantilla = os.path.join(os.path.dirname(__file__), "plantillas", f"{tipo} Informe.docx")
            archivo_final = generar_word_pro(d_m, datos_extraidos, ruta_plantilla)
            
            if archivo_final:
                st.success("Informe generado correctamente.")
                st.download_button("⬇️ Descargar Informe Word", archivo_final, f"Informe_FeNO_{d_m['rut']}.docx")

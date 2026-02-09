import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- 1. EXTRACCIÓN AVANZADA DE DATOS Y GRÁFICOS ---
def procesar_pdf_experto(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina = doc_pdf[0]
    texto_completo = pagina.get_text()

    # RECORTE QUIRÚRGICO DE CURVAS
    # Coordenadas exactas para Sunvou (ajustadas para eliminar bordes negros)
    # [x0, y0, x1, y1]
    rect_exhala = fitz.Rect(48, 438, 290, 570)   # Curva flujo/tiempo
    rect_analisis = fitz.Rect(308, 438, 550, 570) # Curva de análisis técnica

    # Generar imágenes con alta densidad de píxeles (300 DPI aprox con matrix 3)
    pix1 = pagina.get_pixmap(clip=rect_exhala, matrix=fitz.Matrix(4, 4))
    pix2 = pagina.get_pixmap(clip=rect_analisis, matrix=fitz.Matrix(4, 4))

    # LÓGICA DE EXTRACCIÓN DE DATOS TÉCNICOS
    def extraer(patron, texto):
        match = re.search(patron, texto, re.IGNORECASE)
        return match.group(1).strip() if match else "---"

    datos = {
        "feno50": extraer(r"FeN[O0]50[:\s]*(\d+)", texto_completo),
        "temp": extraer(r"Temperatura[:\s]*([\d\.,]+)", texto_completo),
        "presion": extraer(r"Presión[:\s]*([\d\.,]+)", texto_completo),
        "flujo": extraer(r"Tasa de flujo[:\s]*([\d\.,]+)", texto_completo),
        "img_exhala": pix1.tobytes("png"),
        "img_analisis": pix2.tobytes("png")
    }
    return datos

# --- 2. MOTOR DE REEMPLAZO EN WORD ---
def generar_word_preciso(datos_m, datos_e, plantilla_path):
    if not os.path.exists(plantilla_path): return None
    
    doc = Document(plantilla_path)
    # Unificamos etiquetas para que coincidan con tu Word "Informe2"
    reemplazos = {
        "{{NOMBRE}}": datos_m['nombre'],
        "{{APELLIDOS}}": datos_m['apellidos'],
        "{{RUT}}": datos_m['rut'],
        "{{GENERO}}": datos_m['genero'],
        "{{F. nacimiento}}": datos_m['f_nac'],
        "{{Edad}}": datos_m['edad'],
        "{{Altura}}": datos_m['altura'],
        "{{Peso}}": datos_m['peso'],
        "{{Médico}}": datos_m['medico'],
        "{{Operador}}": datos_m['operador'],
        "{{Fecha de Examen}}": datos_m['fecha_ex'],
        "{{Temperatura}}": datos_e['temp'],
        "{{Presion}}": datos_e['presion'],
        "{{Tasa de flujo}}": datos_e['flujo'],
        "{{FeNO50}}": datos_e['feno50']
    }

    def procesar_texto(p):
        # Reemplazar texto manteniendo formato si es posible
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))
        
        # Inserción de Imágenes por Marcadores Únicos
        if "CURVA_EXHALA" in p.text:
            p.text = p.text.replace("CURVA_EXHALA", "")
            run = p.add_run()
            run.add_picture(io.BytesIO(datos_e['img_exhala']), width=Inches(2.3))
        
        if "CURVA_ANALISIS" in p.text:
            p.text = p.text.replace("CURVA_ANALISIS", "")
            run = p.add_run()
            run.add_picture(io.BytesIO(datos_e['img_analisis']), width=Inches(2.3))

    for p in doc.paragraphs: procesar_texto(p)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs: procesar_texto(p)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 3. INTERFAZ ---
st.set_page_config(page_title="INT FeNO Pro", layout="wide")
st.title("🏥 Generador de Informes FeNO (Nivel Experto)")

c1, c2 = st.columns(2)
with c1:
    st.subheader("Datos Manuales")
    d_m = {
        'nombre': st.text_input("Nombre"),
        'apellidos': st.text_input("Apellidos"),
        'rut': st.text_input("RUT"),
        'genero': st.selectbox("Género", ["Hombre", "Mujer"]),
        'f_nac': st.text_input("F. nacimiento (DD/MM/AAAA)"),
        'edad': st.text_input("Edad"),
        'altura': st.text_input("Altura"),
        'peso': st.text_input("Peso"),
        'medico': st.text_input("Médico"),
        'operador': st.text_input("Operador", "TM Jorge Espinoza"),
        'fecha_ex': st.date_input("Fecha Examen").strftime("%d/%m/%Y")
    }

with c2:
    st.subheader("Archivo PDF")
    pdf_file = st.file_uploader("Cargar reporte Sunvou", type="pdf")
    tipo = st.radio("Plantilla", ["FeNO 50", "FeNO 50-200"])

if st.button("✨ Generar Informe Final"):
    if pdf_file and d_m['nombre']:
        res_e = procesar_pdf_experto(pdf_file)
        
        # Preview de recortes para control de calidad
        st.write("### Control de Calidad de Imágenes")
        v1, v2 = st.columns(2)
        v1.image(res_e['img_exhala'], caption="Gráfico 1: Exhalación")
        v2.image(res_e['img_analisis'], caption="Gráfico 2: Análisis")

        path = os.path.join(os.path.dirname(__file__), "plantillas", f"{tipo} Informe.docx")
        final_doc = generar_word_preciso(d_m, res_e, path)
        
        if final_doc:
            st.success("¡Informe procesado con éxito!")
            st.download_button("⬇️ Descargar Word", final_doc, f"FeNO_{d_m['rut']}.docx")

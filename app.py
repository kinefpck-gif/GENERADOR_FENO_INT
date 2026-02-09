import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- 1. EXTRACCIÓN QUIRÚRGICA DE DATOS Y CURVAS ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina = doc_pdf[0]
    texto_completo = pagina.get_text()

    # RECORTE ÚNICO: Captura ambas curvas juntas con precisión
    # Coordenadas ajustadas para Sunvou [x0, y0, x1, y1]
    rect_unificado = fitz.Rect(40, 435, 560, 580) 
    pix = pagina.get_pixmap(clip=rect_unificado, matrix=fitz.Matrix(4, 4)) # Alta resolución

    def buscar(patron, texto):
        match = re.search(patron, texto, re.IGNORECASE)
        return match.group(1).strip() if match else "---"

    return {
        "feno50": buscar(r"FeN[O0]50[:\s]*(\d+)", texto_completo),
        "temp": buscar(r"Temperatura[:\s]*([\d\.,]+)", texto_completo),
        "presion": buscar(r"Presión[:\s]*([\d\.,]+)", texto_completo),
        "flujo": buscar(r"Tasa de flujo[:\s]*([\d\.,]+)", texto_completo),
        "img_curvas": pix.tobytes("png")
    }

# --- 2. MOTOR DE GENERACIÓN WORD (TODO A MAYÚSCULAS) ---
def generar_word(datos_m, datos_e, plantilla_path):
    if not os.path.exists(plantilla_path): return None
    doc = Document(plantilla_path)
    
    # Mapeo y transformación a MAYÚSCULAS
    reemplazos = {
        "{{NOMBRE}}": str(datos_m['nombre']).upper(),
        "{{APELLIDOS}}": str(datos_m['apellidos']).upper(),
        "{{RUT}}": str(datos_m['rut']).upper(),
        "{{GENERO}}": str(datos_m['genero']).upper(),
        "{{OPERADOR}}": str(datos_m['operador']).upper(),
        "{{MEDICO}}": str(datos_m['medico']).upper(),
        "{{F. NACIMIENTO}}": str(datos_m['f_nac']).upper(),
        "{{EDAD}}": str(datos_m['edad']).upper(),
        "{{ALTURA}}": str(datos_m['altura']).upper(),
        "{{PESO}}": str(datos_m['peso']).upper(),
        "{{PROCEDENCIA}}": str(datos_m['procedencia']).upper(),
        "{{FECHA_EXAMEN}}": str(datos_m['fecha_ex']).upper(),
        "{{TEMPERATURA}}": datos_e['temp'],
        "{{PRESION}}": datos_e['presion'],
        "{{TASA DE FLUJO}}": datos_e['flujo'],
        "{{FeNO50}}": datos_e['feno50']
    }

    def procesar_texto(p):
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, v)
        
        # Inserción de la imagen unificada
        if "CURVA_EXHALA" in p.text:
            p.text = p.text.replace("CURVA_EXHALA", "")
            run = p.add_run()
            run.add_picture(io.BytesIO(datos_e['img_curvas']), width=Inches(5.2))

    for p in doc.paragraphs: procesar_texto(p)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs: procesar_texto(p)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 3. INTERFAZ STREAMLIT ---
st.set_page_config(page_title="INT FeNO - Sistema Experto", layout="wide")
st.title("🏥 Generador de Informes Laboratorio INT")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📝 Datos del Paciente")
    d_m = {
        'nombre': st.text_input("Nombre"),
        'apellidos': st.text_input("Apellidos"),
        'rut': st.text_input("RUT"),
        'genero': st.selectbox("Género", ["Hombre", "Mujer"]),
        'f_nac': st.text_input("F. nacimiento (DD/MM/AAAA)"),
        'edad': st.text_input("Edad"),
        'altura': st.text_input("Altura"),
        'peso': st.text_input("Peso"),
        'procedencia': st.text_input("Procedencia", "Poli"),
        'medico': st.text_input("Médico"),
        'operador': st.text_input("Operador", "TM JORGE ESPINOZA"),
        'fecha_ex': st.date_input("Fecha Examen").strftime("%d/%m/%Y")
    }

with col2:
    st.subheader("📂 Carga de Archivos")
    pdf_file = st.file_uploader("Subir PDF Sunvou", type="pdf")
    tipo = st.radio("Plantilla:", ["FeNO 50", "FeNO 50-200"])

if st.button("🚀 Generar Informe"):
    if pdf_file and d_m['nombre']:
        res = procesar_pdf_sunvou(pdf_file)
        
        st.write("### Vista previa de extracción")
        st.image(res['img_curvas'], caption="Imagen Unificada (CURVA_EXHALA)")

        path = os.path.join(os.path.dirname(__file__), "plantillas", f"{tipo} Informe.docx")
        doc_out = generar_word(d_m, res, path)
        
        if doc_out:
            st.success("✅ Informe generado exitosamente")
            st.download_button("⬇️ Descargar Word", doc_out, f"FeNO_{d_m['rut']}.docx")

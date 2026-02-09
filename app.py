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

    # RECORTE DE CURVAS (Ajustado según Informe2)
    # Curva 1: Curva de Exhalación (Cuadro izquierdo)
    rect_exhala = fitz.Rect(45, 435, 295, 575)
    pix_exhala = pagina.get_pixmap(clip=rect_exhala, matrix=fitz.Matrix(3, 3)) # Alta calidad
    
    # Curva 2: Análisis de Curva (Cuadro derecho)
    rect_analisis = fitz.Rect(305, 435, 555, 575)
    pix_analisis = pagina.get_pixmap(clip=rect_analisis, matrix=fitz.Matrix(3, 3))

    # EXTRACCIÓN TÉCNICA (Búsqueda por palabras clave en el PDF)
    def buscar_valor(patron, texto):
        match = re.search(patron, texto, re.IGNORECASE)
        return match.group(1).strip() if match else "---"

    # Buscamos los valores técnicos específicos
    feno50 = buscar_valor(r"FeN[O0]50[:\s]*(\d+)", texto_completo)
    temp = buscar_valor(r"Temperatura[:\s]*([\d\.]+)", texto_completo)
    presion = buscar_valor(r"Presión[:\s]*([\d\.]+)", texto_completo)
    flujo = buscar_valor(r"Tasa de flujo[:\s]*([\d\.]+)", texto_completo)
    
    return {
        "feno50": feno50,
        "temp": temp,
        "presion": presion,
        "flujo": flujo,
        "img_exhala": pix_exhala.tobytes("png"),
        "img_analisis": pix_analisis.tobytes("png")
    }

# --- 2. GENERACIÓN DEL WORD IDENTICO A LA MUESTRA ---
def generar_word(datos_completos, plantilla_path):
    if not os.path.exists(plantilla_path):
        return None

    doc = Document(plantilla_path)
    
    def procesar_bloque(p):
        original_text = p.text
        # Reemplazo de texto (Sensible a mayúsculas/minúsculas según tus {{etiquetas}})
        for k, v in datos_completos.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))
        
        # Inserción de Imágenes en sus marcadores
        if "CURVA_EXHALA" in original_text:
            p.text = p.text.replace("CURVA_EXHALA", "")
            run = p.add_run()
            run.add_picture(io.BytesIO(datos_completos['img_exhala']), width=Inches(2.4))
        
        if "CURVA_ANALISIS" in original_text:
            p.text = p.text.replace("CURVA_ANALISIS", "")
            run = p.add_run()
            run.add_picture(io.BytesIO(datos_completos['img_analisis']), width=Inches(2.4))

    # Aplicar a todo el documento
    for p in doc.paragraphs: procesar_bloque(p)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs: procesar_bloque(p)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 3. INTERFAZ ---
st.set_page_config(page_title="INT - FeNO", layout="wide")
st.title("🫁 Generador de Informes FeNO - Versión 2.0")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📋 Datos del Paciente")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT")
    genero = st.selectbox("Género", ["Hombre", "Mujer"])
    f_nac = st.text_input("F. Nacimiento (ej: 01/01/1990)")
    edad = st.text_input("Edad")
    altura = st.text_input("Altura")
    peso = st.text_input("Peso")
    medico = st.text_input("Médico")
    operador = st.text_input("Operador", "TM Jorge Espinoza")
    fecha_ex = st.date_input("Fecha Examen")

with col2:
    st.subheader("📄 Reporte Sunvou")
    pdf_file = st.file_uploader("Subir PDF", type="pdf")
    tipo = st.radio("Plantilla", ["FeNO 50", "FeNO 50-200"])

if st.button("🚀 Crear Informe"):
    if pdf_file and nombre:
        res = procesar_pdf_sunvou(pdf_file)
        
        # Mostramos qué encontró para estar seguros
        st.write(f"📈 **Valores:** FeNO50: {res['feno50']} | Temp: {res['temp']} | Flujo: {res['flujo']}")
        
        # Diccionario unificado para el Word
        final_data = {
            "{{NOMBRE}}": nombre, "{{APELLIDOS}}": apellidos, "{{RUT}}": rut,
            "{{GENERO}}": genero, "{{F_NACIMIENTO}}": f_nac, "{{EDAD}}": edad,
            "{{ALTURA}}": altura, "{{PESO}}": peso, "{{MEDICO}}": medico,
            "{{OPERADOR}}": operador, "{{FECHA_EXAMEN}}": fecha_ex.strftime("%d/%m/%Y"),
            "{{FeNO50}}": res['feno50'], "{{Temperatura}}": res['temp'],
            "{{Presion}}": res['presion'], "{{Tasa de flujo}}": res['flujo'],
            "img_exhala": res['img_exhala'], "img_analisis": res['img_analisis']
        }

        path = os.path.join(os.path.dirname(__file__), "plantillas", f"{tipo} Informe.docx")
        doc_final = generar_word(final_data, path)
        
        if doc_final:
            st.success("Informe generado correctamente")
            st.download_button("⬇️ Descargar Word", doc_final, f"Informe_{rut}.docx")

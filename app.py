import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- 1. FUNCIÓN DE EXTRACCIÓN DE DATOS Y GRÁFICOS ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    imagenes_curvas = []

    for pagina in doc_pdf:
        texto_completo += pagina.get_text()
        # Captura el área de las gráficas (mitad inferior del PDF)
        # Rect(x0, y0, x1, y1) - Ajustado para capturar las curvas de Sunvou
        rect = fitz.Rect(40, 380, 560, 820) 
        pix = pagina.get_pixmap(clip=rect, matrix=fitz.Matrix(2, 2))
        imagenes_curvas.append(pix.tobytes("png"))

    # Búsqueda de valores (Soporta O y 0 por el error común del equipo)
    f50 = re.search(r"FeN[O0]50[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    f200 = re.search(r"FeN[O0]200[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    cano = re.search(r"CaN[O0][:\s]*(\d+)", texto_completo, re.IGNORECASE)
    
    return {
        "f50": f50.group(1) if f50 else "---",
        "f200": f200.group(1) if f200 else "---",
        "cano": cano.group(1) if cano else "---",
        "curvas": imagenes_curvas
    }

# --- 2. FUNCIÓN PARA GENERAR EL WORD ---
def generar_word(datos_m, datos_e, plantilla_path):
    if not os.path.exists(plantilla_path):
        return None

    doc = Document(plantilla_path)
    reemplazos = {**datos_m, **datos_e}
    
    def procesar_texto_e_imagen(p):
        # Reemplazar etiquetas de texto
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))
        
        # Insertar imagen en el marcador
        if "CURVA_GRAFICA" in p.text:
            p.text = p.text.replace("CURVA_GRAFICA", "")
            if datos_e['curvas']:
                run = p.add_run()
                img_stream = io.BytesIO(datos_e['curvas'][0])
                run.add_picture(img_stream, width=Inches(5.0))

    # Revisar párrafos y tablas
    for p in doc.paragraphs:
        procesar_texto_e_imagen(p)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    procesar_texto_e_imagen(p)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 3. INTERFAZ DE USUARIO ---
st.set_page_config(page_title="INT - Informe FeNO", layout="wide")
st.title("🏥 Sistema de Informes Laboratorio Función Pulmonar")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📝 Datos del Paciente (Manual)")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT")
    genero = st.selectbox("Género", ["Hombre", "Mujer"])
    f_nac = st.text_input("F. Nacimiento (DD/MM/AAAA)")
    edad = st.text_input("Edad")
    altura = st.text_input("Altura (cm)")
    peso = st.text_input("Peso (kg)")
    medico = st.text_input("Médico Solicitante")
    operador = st.text_input("Operador", value="TM Jorge Espinoza")
    fecha_examen = st.date_input("Fecha de Examen")

with col2:
    st.subheader("📂 Extracción desde PDF")
    pdf_file = st.file_uploader("Subir PDF Sunvou", type="pdf")
    tipo_inf = st.radio("Plantilla de salida:", ["FeNO 50", "FeNO 50-200"])

if st.button("🚀 Generar Informe"):
    if pdf_file and nombre and rut:
        with st.spinner("Procesando..."):
            res_pdf = procesar_pdf_sunvou(pdf_file)
            
            # Mostrar preview de lo extraído
            st.write(f"📊 **Datos extraídos:** FeNO50: {res_pdf['f50']} | FeNO200: {res_pdf['f200']} | CaNO: {res_pdf['cano']}")
            if res_pdf['curvas']:
                st.image(res_pdf['curvas'][0], caption="Previsualización de Gráfica", width=300)

            # Diccionario para el Word
            datos_m = {
                "{{NOMBRE}}": nombre, "{{APELLIDOS}}": apellidos, "{{RUT}}": rut,
                "{{GENERO}}": genero, "{{F_NAC}}": f_nac, "{{EDAD}}": edad,
                "{{ALTURA}}": altura, "{{PESO}}": peso, "{{MEDICO}}": medico,
                "{{OPERADOR}}": operador, "{{FECHA_EXAMEN}}": fecha_examen.strftime("%d/%m/%Y"),
                "{{RAZA}}": "Caucásica", "{{PROCEDENCIA}}": "Poli"
            }
            datos_e = {
                "{{FENO50}}": res_pdf['f50'],
                "{{FENO200}}": res_pdf['f200'],
                "{{CANO}}": res_pdf['cano'],
                "curvas": res_pdf['curvas']
            }

            base_dir = os.path.dirname(os.path.abspath(__file__))
            path = os.path.join(base_dir, "plantillas", f"{tipo_inf} Informe.docx")
            
            doc_final = generar_word(datos_m, datos_e, path)
            
            if doc_final:
                st.success("Informe creado exitosamente")
                st.download_button("⬇️ Descargar Informe", doc_final, f"Informe_{rut}.docx")
            else:
                st.error("No se encontró la plantilla en /plantillas")
    else:
        st.warning("Faltan datos obligatorios (Nombre, RUT o PDF)")

import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- FUNCIÓN DE EXTRACCIÓN MEJORADA ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    imagenes_curvas = []

    for pagina in doc_pdf:
        texto_completo += pagina.get_text()
        
        # Extraer el área de la curva (Sunvou suele tenerlas al final de la página)
        # Ajustamos el área de captura (x0, y0, x1, y1)
        pix = pagina.get_pixmap(clip=fitz.Rect(0, 380, 600, 850))
        imagenes_curvas.append(pix.tobytes("png"))

    # Búsqueda de valores (Flexible para FeNO o FeN0)
    f50_match = re.search(r"FeN[O0]50:\s*(\d+)", texto_completo, re.IGNORECASE)
    f200_match = re.search(r"FeN[O0]200:\s*(\d+)", texto_completo, re.IGNORECASE)
    
    return {
        "f50": f50_match.group(1) if f50_match else "---",
        "f200": f200_match.group(1) if f200_match else "---",
        "curvas": imagenes_curvas
    }

# --- FUNCIÓN PARA GENERAR EL WORD ---
def generar_word(datos_m, datos_e, plantilla_path):
    if not os.path.exists(plantilla_path):
        return None

    doc = Document(plantilla_path)
    reemplazos = {**datos_m, **datos_e}
    
    # Reemplazar en párrafos
    for p in doc.paragraphs:
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))
                
    # Reemplazar en tablas
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    for k, v in reemplazos.items():
                        if k in p.text:
                            p.text = p.text.replace(k, str(v))

    # Insertar la imagen donde diga CURVA_GRAFICA
    for p in doc.paragraphs:
        if "CURVA_GRAFICA" in p.text:
            p.text = p.text.replace("CURVA_GRAFICA", "")
            run = p.add_run()
            if datos_e['curvas']:
                img_stream = io.BytesIO(datos_e['curvas'][0])
                run.add_picture(img_stream, width=Inches(5.0))

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="INT - Laboratorio Función Pulmonar", layout="wide")

st.title("🏥 Extractor FeNO Sunvou a Informe INT")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Datos del Paciente (Manual)")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT")
    edad = st.text_input("Edad")
    genero = st.selectbox("Género", ["Hombre", "Mujer"])
    fecha_examen = st.date_input("Fecha del Examen")

with col2:
    st.subheader("2. Cargar PDF del Equipo")
    pdf_file = st.file_uploader("Arrastra aquí el PDF original de Sunvou", type="pdf")
    tipo_inf = st.radio("Plantilla de salida:", ["FeNO 50", "FeNO 50-200"])

if st.button("🚀 Generar Informe Final"):
    if pdf_file and nombre and rut:
        with st.spinner("Procesando..."):
            # Extraer del PDF
            res_pdf = procesar_pdf_sunvou(pdf_file)
            
            # Preparar datos
            datos_m = {
                "{{NOMBRE}}": nombre, "{{APELLIDOS}}": apellidos,
                "{{RUT}}": rut, "{{EDAD}}": edad, "{{GENERO}}": genero,
                "{{FECHA_EXAMEN}}": str(fecha_examen)
            }
            datos_e = {
                "{{FENO50}}": res_pdf['f50'],
                "{{FENO200}}": res_pdf['f200'],
                "curvas": res_pdf['curvas']
            }
            
            # Ruta de plantilla
            base_dir = os.path.dirname(os.path.abspath(__file__))
            plantilla_path = os.path.join(base_dir, "plantillas", f"{tipo_inf} Informe.docx")
            
            archivo_word = generar_word(datos_m, datos_e, plantilla_path)
            
            if archivo_word:
                st.success("¡Informe procesado!")
                st.download_button(
                    label="⬇️ Descargar Informe .docx",
                    data=archivo_word,
                    file_name=f"FeNO_{rut}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error(f"Error: No se encontró la plantilla en {plantilla_path}")
    else:
        st.warning("Asegúrate de rellenar el nombre, RUT y subir el PDF.")

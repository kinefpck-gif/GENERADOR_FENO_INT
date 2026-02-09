import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- 1. EXTRACCIÓN DE DATOS Y RECORTE DE IMAGEN ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    imagenes_curvas = []

    for pagina in doc_pdf:
        texto_completo += pagina.get_text()
        
        # Coordenadas ajustadas para Sunvou (Captura la mitad inferior del PDF)
        # Rect(x0, y0, x1, y1). Subimos un poco el inicio (380) para atrapar bien la curva.
        rect = fitz.Rect(40, 380, 560, 820) 
        pix = pagina.get_pixmap(clip=rect, matrix=fitz.Matrix(2, 2))
        imagenes_curvas.append(pix.tobytes("png"))

    # Extracción de valores FeNO (Acepta O y 0)
    f50_match = re.search(r"FeN[O0]50[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    f200_match = re.search(r"FeN[O0]200[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    
    return {
        "f50": f50_match.group(1) if f50_match else "---",
        "f200": f200_match.group(1) if f200_match else "---",
        "curvas": imagenes_curvas
    }

# --- 2. GENERACIÓN DEL INFORME WORD ---
def generar_word(datos_m, datos_e, plantilla_path):
    if not os.path.exists(plantilla_path):
        return None

    doc = Document(plantilla_path)
    reemplazos = {**datos_m, **datos_e}
    
    # Función interna para procesar párrafos (ayuda a no repetir código)
    def procesar_parrafo(p):
        # Reemplazar texto
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))
        
        # Insertar Imagen si encuentra el marcador
        if "CURVA_GRAFICA" in p.text:
            p.text = p.text.replace("CURVA_GRAFICA", "")
            if datos_e['curvas']:
                run = p.add_run()
                img_stream = io.BytesIO(datos_e['curvas'][0])
                run.add_picture(img_stream, width=Inches(5.2))

    # Procesar párrafos normales
    for p in doc.paragraphs:
        procesar_parrafo(p)
                
    # Procesar párrafos dentro de tablas (Muy común en tus plantillas)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    procesar_parrafo(p)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 3. INTERFAZ DE USUARIO ---
st.set_page_config(page_title="INT - Laboratorio FeNO", layout="wide")

st.title("🏥 Generador de Informes FeNO - INT")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📝 Datos del Paciente")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT")
    edad = st.text_input("Edad")
    genero = st.selectbox("Género", ["Hombre", "Mujer"])
    fecha_hoy = st.date_input("Fecha Examen")

with col2:
    st.subheader("📂 PDF del Equipo")
    pdf_file = st.file_uploader("Subir PDF de Sunvou", type="pdf")
    tipo_inf = st.radio("Plantilla:", ["FeNO 50", "FeNO 50-200"])

if st.button("🚀 Crear Informe Final"):
    if pdf_file and nombre and rut:
        with st.spinner("Procesando..."):
            res_pdf = procesar_pdf_sunvou(pdf_file)
            
            # Previsualización para el usuario
            st.write(f"✅ **Detectado:** FeNO50: {res_pdf['f50']} | FeNO200: {res_pdf['f200']}")
            if res_pdf['curvas']:
                st.image(res_pdf['curvas'][0], caption="Curva detectada", width=350)

            datos_manuales = {
                "{{NOMBRE}}": nombre, "{{APELLIDOS}}": apellidos,
                "{{RUT}}": rut, "{{EDAD}}": edad, "{{GENERO}}": genero,
                "{{FECHA_EXAMEN}}": fecha_hoy.strftime("%d/%m/%Y")
            }
            datos_extraidos = {
                "{{FENO50}}": res_pdf['f50'],
                "{{FENO200}}": res_pdf['f200'],
                "curvas": res_pdf['curvas']
            }

            base_dir = os.path.dirname(os.path.abspath(__file__))
            plantilla_path = os.path.join(base_dir, "plantillas", f"{tipo_inf} Informe.docx")
            
            archivo = generar_word(datos_manuales, datos_extraidos, plantilla_path)
            
            if archivo:
                st.success("¡Informe listo!")
                st.download_button("⬇️ Descargar Word", archivo, f"FeNO_{rut}.docx")
            else:
                st.error("No se encontró la plantilla en la carpeta 'plantillas'.")
    else:
        st.warning("Completa los datos y sube el PDF.")

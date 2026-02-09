import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
from PIL import Image

# --- FUNCIÓN PARA EXTRAER DATOS Y CURVAS ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    imagenes_curvas = []

    for pagina in doc_pdf:
        texto_completo += pagina.get_text()
        
        # Extraer imágenes o dibujos (las curvas suelen ser gráficos)
        # En los equipos Sunvou, a veces la curva es un dibujo vectorial.
        # Este comando captura el área de la curva como una imagen:
        pix = pagina.get_pixmap(clip=fitz.Rect(50, 400, 550, 800)) # Ajustar coordenadas según reporte
        img_data = pix.tobytes("png")
        imagenes_curvas.append(img_data)

    # Buscar valores numéricos
    f50 = re.search(r"Valor de FeN050:\s*(\d+)", texto_completo)
    f200 = re.search(r"Valor de FeN0200:\s*(\d+)", texto_completo)
    
    return {
        "f50": f50.group(1) if f50 else "---",
        "f200": f200.group(1) if f200 else "---",
        "curvas": imagenes_curvas
    }

# --- FUNCIÓN PARA GENERAR EL WORD FINAL ---
def generar_word(datos_manuales, datos_extraidos, plantilla_path):
    doc = Document(plantilla_path)
    
    # Combinar diccionarios de reemplazo
    reemplazos = {**datos_manuales, **datos_extraidos}
    
    # Reemplazar texto en párrafos y tablas
    for p in doc.paragraphs:
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))
                
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    for k, v in reemplazos.items():
                        if k in p.text:
                            p.text = p.text.replace(k, str(v))

    # INSERTAR IMÁGENES DE LAS CURVAS
    # El código busca un párrafo que diga "IMAGEN_CURVA" y pone la foto ahí
    for p in doc.paragraphs:
        if "CURVA_GRAFICA" in p.text:
            p.text = p.text.replace("CURVA_GRAFICA", "")
            run = p.add_run()
            img_stream = io.BytesIO(datos_extraidos['curvas'][0])
            run.add_picture(img_stream, width=Inches(4.5))

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- INTERFAZ WEB (STREAMLIT) ---
st.set_page_config(page_title="INT - Generador FeNO", layout="wide")
st.title("🫁 Generador de Informes FeNO")

col_izq, col_der = st.columns(2)

with col_izq:
    st.subheader("1. Datos del Paciente (Manual)")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT")
    edad = st.text_input("Edad")
    genero = st.selectbox("Género", ["Hombre", "Mujer"])
    fecha = st.date_input("Fecha del Examen")

with col_der:
    st.subheader("2. Datos del Equipo (PDF)")
    pdf_file = st.file_uploader("Sube el PDF de Sunvou", type="pdf")
    tipo_inf = st.radio("Tipo de informe", ["FeNO 50", "FeNO 50-200"])

if st.button("🚀 Procesar y Generar Informe"):
    if pdf_file and nombre:
        with st.spinner("Extrayendo curvas y datos..."):
            res_pdf = procesar_pdf_sunvou(pdf_file)
            
            datos_m = {
                "{{NOMBRE}}": nombre, "{{APELLIDOS}}": apellidos,
                "{{RUT}}": rut, "{{EDAD}}": edad, "{{GENERO}}": genero,
                "{{FECHA_EXAMEN}}": str(fecha)
            }
            
            datos_e = {
                "{{FENO50}}": res_pdf['f50'],
                "{{FENO200}}": res_pdf['f200'],
                "curvas": res_pdf['curvas']
            }
            
            plantilla = f"plantillas/{tipo_inf} Informe.docx"
            archivo_final = generar_word(datos_m, datos_e, plantilla)
            
            st.success("¡Informe generado con éxito!")
            st.download_button("⬇️ Descargar Informe Final", archivo_final, f"Informe_{rut}.docx")
    else:
        st.error("Por favor, completa los datos y sube el PDF.")

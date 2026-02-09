import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- 1. FUNCIÓN DE EXTRACCIÓN CON RECORTE DE DOBLE CURVA ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina = doc_pdf[0] # El reporte siempre es de 1 página
    texto_completo = pagina.get_text()

    # RECORTE 1: Curva de Exhalación (Cuadro Izquierdo)
    # Coordenadas: [x0, y0, x1, y1]
    rect_exhala = fitz.Rect(40, 440, 300, 580) 
    pix_exhala = pagina.get_pixmap(clip=rect_exhala, matrix=fitz.Matrix(2, 2))
    
    # RECORTE 2: Análisis de Curva (Cuadro Derecho)
    rect_analisis = fitz.Rect(310, 440, 560, 580)
    pix_analisis = pagina.get_pixmap(clip=rect_analisis, matrix=fitz.Matrix(2, 2))

    # Valores numéricos (Robustez para O/0)
    f50 = re.search(r"FeN[O0]50[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    f200 = re.search(r"FeN[O0]200[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    cano = re.search(r"CaN[O0][:\s]*(\d+)", texto_completo, re.IGNORECASE)
    
    return {
        "f50": f50.group(1) if f50 else "---",
        "f200": f200.group(1) if f200 else "---",
        "cano": cano.group(1) if cano else "---",
        "img_exhala": pix_exhala.tobytes("png"),
        "img_analisis": pix_analisis.tobytes("png")
    }

# --- 2. FUNCIÓN DE GENERACIÓN DE WORD ---
def generar_word(datos_m, datos_e, plantilla_path):
    if not os.path.exists(plantilla_path):
        return None

    doc = Document(plantilla_path)
    
    # Unimos todos los datos para el reemplazo de texto
    reemplazos = {**datos_m, **datos_e}
    
    def procesar_párrafos(párrafos):
        for p in párrafos:
            # Reemplazo de etiquetas de texto
            for k, v in reemplazos.items():
                if k in p.text:
                    p.text = p.text.replace(k, str(v))
            
            # Inserción de Curva de Exhalación
            if "CURVA_EXHALA" in p.text:
                p.text = p.text.replace("CURVA_EXHALA", "")
                run = p.add_run()
                run.add_picture(io.BytesIO(datos_e['img_exhala']), width=Inches(2.5))
            
            # Inserción de Análisis de Curva
            if "CURVA_ANALISIS" in p.text:
                p.text = p.text.replace("CURVA_ANALISIS", "")
                run = p.add_run()
                run.add_picture(io.BytesIO(datos_e['img_analisis']), width=Inches(2.5))

    # Procesar párrafos principales y tablas
    procesar_párrafos(doc.paragraphs)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                procesar_párrafos(celda.paragraphs)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 3. INTERFAZ DE USUARIO ---
st.set_page_config(page_title="INT - FeNO", layout="wide")
st.title("🫁 Generador de Informes FeNO - INT")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📝 Datos del Paciente")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT")
    genero = st.selectbox("Género", ["Hombre", "Mujer"])
    f_nac = st.text_input("F. nacimiento")
    edad = st.text_input("Edad")
    altura = st.text_input("Altura")
    peso = st.text_input("Peso")
    medico = st.text_input("Médico")
    operador = st.text_input("Operador", value="TM Jorge Espinoza")
    fecha_examen = st.date_input("Fecha de Examen")

with col2:
    st.subheader("📂 Archivo del Equipo")
    pdf_file = st.file_uploader("Subir PDF de Sunvou", type="pdf")
    tipo_inf = st.radio("Plantilla:", ["FeNO 50", "FeNO 50-200"])

if st.button("🚀 Generar Informe Word"):
    if pdf_file and nombre and rut:
        with st.spinner("Extrayendo curvas y datos..."):
            res = procesar_pdf_sunvou(pdf_file)
            
            # Previsualización para que veas si el recorte es correcto
            st.write("🔎 **Vista previa de recortes:**")
            v1, v2 = st.columns(2)
            v1.image(res['img_exhala'], caption="Curva de Exhalación")
            v2.image(res['img_analisis'], caption="Análisis de Curva")

            # Mapeo de etiquetas
            datos_m = {
                "{{NOMBRE}}": nombre, "{{APELLIDOS}}": apellidos, "{{RUT}}": rut,
                "{{GENERO}}": genero, "{{F_NAC}}": f_nac, "{{EDAD}}": edad,
                "{{ALTURA}}": altura, "{{PESO}}": peso, "{{MEDICO}}": medico,
                "{{OPERADOR}}": operador, "{{FECHA_EXAMEN}}": fecha_examen.strftime("%d/%m/%Y"),
                "{{RAZA}}": "Caucásica", "{{PROCEDENCIA}}": "Poli"
            }
            datos_e = {
                "{{FENO50}}": res['f50'], "{{FENO200}}": res['f200'], "{{CANO}}": res['cano'],
                "img_exhala": res['img_exhala'], "img_analisis": res['img_analisis']
            }

            base_dir = os.path.dirname(os.path.abspath(__file__))
            plantilla_path = os.path.join(base_dir, "plantillas", f"{tipo_inf} Informe.docx")
            
            archivo_word = generar_word(datos_m, datos_e, plantilla_path)
            
            if archivo_word:
                st.success("✅ Informe generado exitosamente")
                st.download_button("⬇️ Descargar Informe", archivo_word, f"FeNO_{rut}.docx")
            else:
                st.error("Error: No se encontró la plantilla en la carpeta /plantillas")
    else:
        st.warning("Faltan datos (Nombre, RUT o PDF)")

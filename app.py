import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# --- FUNCIÓN DE EXTRACCIÓN ROBUSTA ---
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    imagenes_curvas = []

    for pagina in doc_pdf:
        texto_completo += pagina.get_text()
        
        # Captura el área donde Sunvou pone las curvas (coordenadas ajustadas)
        pix = pagina.get_pixmap(clip=fitz.Rect(0, 350, 600, 850))
        imagenes_curvas.append(pix.tobytes("png"))

    # Búsqueda de valores: Soporta "FeNO50", "FeN050", "FeNO 50", etc.
    # El equipo Sunvou suele usar el número '0' en lugar de 'O'
    f50_match = re.search(r"FeN[O0]50[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    f200_match = re.search(r"FeN[O0]200[:\s]*(\d+)", texto_completo, re.IGNORECASE)
    
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
    
    # Unimos datos manuales y extraídos para el reemplazo
    reemplazos = {**datos_m, **datos_e}
    
    # 1. Reemplazo en Párrafos
    for p in doc.paragraphs:
        for k, v in reemplazos.items():
            if k in p.text:
                p.text = p.text.replace(k, str(v))
                
    # 2. Reemplazo en Tablas (donde suelen estar los datos del paciente)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    for k, v in reemplazos.items():
                        if k in p.text:
                            p.text = p.text.replace(k, str(v))

    # 3. Inserción de Imágenes (Busca el marcador CURVA_GRAFICA)
    for p in doc.paragraphs:
        if "CURVA_GRAFICA" in p.text:
            p.text = p.text.replace("CURVA_GRAFICA", "") # Limpia el texto
            run = p.add_run()
            if datos_e['curvas']:
                # Insertamos la primera curva detectada
                img_stream = io.BytesIO(datos_e['curvas'][0])
                run.add_picture(img_stream, width=Inches(5.0))

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- INTERFAZ DE USUARIO (STREAMLIT) ---
st.set_page_config(page_title="INT - Laboratorio FeNO", layout="wide")

st.title("🫁 Generador de Informes FeNO - INT")
st.info("Ingresa los datos del paciente a mano y sube el PDF del equipo para extraer resultados y curvas.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📝 Datos del Paciente")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT (ej: 12.345.678-9)")
    edad = st.text_input("Edad")
    genero = st.selectbox("Género", ["Hombre", "Mujer", "Otro"])
    fecha_hoy = st.date_input("Fecha del Examen")

with col2:
    st.subheader("📂 Datos del Equipo (Sunvou)")
    pdf_file = st.file_uploader("Subir PDF generado por el equipo", type="pdf")
    tipo_inf = st.radio("Plantilla a utilizar:", ["FeNO 50", "FeNO 50-200"])

# --- PROCESAMIENTO ---
if st.button("🚀 Procesar y Descargar Informe"):
    if pdf_file and nombre and rut:
        with st.spinner("Leyendo PDF y generando documento..."):
            # 1. Extraer datos del PDF
            res_pdf = procesar_pdf_sunvou(pdf_file)
            
            # 2. Preparar diccionario de etiquetas para el Word
            datos_manuales = {
                "{{NOMBRE}}": nombre,
                "{{APELLIDOS}}": apellidos,
                "{{RUT}}": rut,
                "{{EDAD}}": edad,
                "{{GENERO}}": genero,
                "{{FECHA_EXAMEN}}": fecha_hoy.strftime("%d/%m/%Y")
            }
            
            datos_extraidos = {
                "{{FENO50}}": res_pdf['f50'],
                "{{FENO200}}": res_pdf['f200'],
                "curvas": res_pdf['curvas']
            }
            
            # Mostrar preview de lo encontrado para seguridad del usuario
            st.write(f"✅ **Datos detectados:** FeNO50: {res_pdf['f50']} ppb | FeNO200: {res_pdf['f200']} ppb")
            
            # 3. Definir ruta de plantilla y generar
            base_dir = os.path.dirname(os.path.abspath(__file__))
            nombre_plantilla = f"{tipo_inf} Informe.docx"
            plantilla_path = os.path.join(base_dir, "plantillas", nombre_plantilla)
            
            archivo_final = generar_word(datos_manuales, datos_extraidos, plantilla_path)
            
            if archivo_final:
                st.success("¡Informe generado con éxito!")
                st.download_button(
                    label="⬇️ Descargar Informe Word",
                    data=archivo_final,
                    file_name=f"Informe_FeNO_{rut}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error(f"Error: No se encontró el archivo '{nombre_plantilla}' en la carpeta 'plantillas'.")
    else:
        st.warning("Por favor completa el Nombre, RUT y sube un archivo PDF.")

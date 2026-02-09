import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os

# ==========================================================
# 1. EXTRACCIÓN EXCLUSIVA DE LA CURVA DE EXHALACIÓN (GRÁFICO)
# ==========================================================
def extraer_curva_exhalacion(pagina):
    """
    Extrae exclusivamente el gráfico de la curva de exhalación
    (línea + ejes), sin título ni márgenes blancos.
    """
    bloques = pagina.get_text("blocks")

    for b in bloques:
        x0, y0, x1, y1, texto, *_ = b

        if "CURVA DE EXHALACIÓN" in texto.upper():
            rect_curva = fitz.Rect(
                x0 + 20,     # elimina margen izquierdo
                y1 + 18,     # baja desde el texto al inicio del gráfico
                x0 + 460,    # ancho real del gráfico
                y1 + 200     # alto real del gráfico
            )

            pix = pagina.get_pixmap(
                clip=rect_curva,
                matrix=fitz.Matrix(4, 4),  # alta resolución clínica
                alpha=False
            )
            return pix.tobytes("png")

    return None


# ==========================================================
# 2. PROCESAMIENTO DEL PDF SUNVOU
# ==========================================================
def procesar_pdf_sunvou(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pagina = doc_pdf[0]
    texto = pagina.get_text()

    img_exhala = extraer_curva_exhalacion(pagina)

    def buscar(patron):
        m = re.search(patron, texto, re.IGNORECASE)
        return m.group(1).strip() if m else "---"

    return {
        "FeNO50": buscar(r"FeN[O0]50[:\s]*(\d+)"),
        "Temperatura": buscar(r"Temperatura[:\s]*([\d\.]+)"),
        "Presion": buscar(r"Presión[:\s]*([\d\.]+)"),
        "Flujo": buscar(r"Tasa de flujo[:\s]*([\d\.]+)"),
        "img_exhala": img_exhala
    }


# ==========================================================
# 3. GENERACIÓN DEL WORD (FORMATO INTACTO)
# ==========================================================
def generar_word(datos, plantilla_path):
    doc = Document(plantilla_path)

    def procesar_parrafo(p):
        texto_original = p.text

        for k, v in datos.items():
            if isinstance(v, str) and k in p.text:
                p.text = p.text.replace(k, v)

        if "CURVA_EXHALA" in texto_original and datos["img_exhala"]:
            p.text = p.text.replace("CURVA_EXHALA", "")
            run = p.add_run()
            run.add_picture(
                io.BytesIO(datos["img_exhala"]),
                width=Inches(2.8)
            )

    for p in doc.paragraphs:
        procesar_parrafo(p)

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    procesar_parrafo(p)

    salida = io.BytesIO()
    doc.save(salida)
    salida.seek(0)
    return salida


# ==========================================================
# 4. INTERFAZ STREAMLIT
# ==========================================================
st.set_page_config(page_title="INT – FeNO", layout="wide")
st.title("🫁 Generador Clínico de Informes FeNO")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📋 Datos del Paciente")
    nombre = st.text_input("Nombre")
    apellidos = st.text_input("Apellidos")
    rut = st.text_input("RUT")
    genero = st.selectbox("Género", ["Hombre", "Mujer"])
    procedencia = st.text_input("Procedencia")
    f_nac = st.text_input("Fecha de nacimiento")
    edad = st.text_input("Edad")
    altura = st.text_input("Altura (cm)")
    peso = st.text_input("Peso (kg)")
    medico = st.text_input("Médico solicitante")
    operador = st.text_input("Operador", "TM Jorge Espinoza")
    fecha = st.date_input("Fecha del examen")

with col2:
    st.subheader("📄 Informe Sunvou")
    pdf = st.file_uploader("Subir PDF", type="pdf")
    tipo = st.radio("Plantilla", ["FeNO 50", "FeNO 50-200"])


# ==========================================================
# 5. EJECUCIÓN
# ==========================================================
if st.button("🚀 Crear informe clínico"):
    if not pdf or not nombre:
        st.error("Faltan datos obligatorios")
    else:
        res = procesar_pdf_sunvou(pdf)

        st.info(
            f"Valores detectados → FeNO50: {res['FeNO50']} | "
            f"T°: {res['Temperatura']} | Flujo: {res['Flujo']}"
        )

        datos_finales = {
            "{{NOMBRE}}": nombre,
            "{{APELLIDOS}}": apellidos,
            "{{RUT}}": rut,
            "{{GENERO}}": genero,
            "{{PROCEDENCIA}}": procedencia,
            "{{F_NACIMIENTO}}": f_nac,
            "{{EDAD}}": edad,
            "{{ALTURA}}": altura,
            "{{PESO}}": peso,
            "{{MEDICO}}": medico,
            "{{OPERADOR}}": operador,
            "{{FECHA_EXAMEN}}": fecha.strftime("%d/%m/%Y"),
            "{{FeNO50}}": res["FeNO50"],
            "{{Temperatura}}": res["Temperatura"],
            "{{Presion}}": res["Presion"],
            "{{Tasa de flujo}}": res["Flujo"],
            "img_exhala": res["img_exhala"]
        }

        plantilla = os.path.join(
            "plantillas", f"{tipo} Informe.docx"
        )

        word = generar_word(datos_finales, plantilla)

        st.success("✅ Informe clínico generado correctamente")
        st.download_button(
            "⬇️ Descargar Word",
            word,
            f"Informe_FeNO_{rut}.docx"
        )


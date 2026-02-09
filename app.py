import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import io
import os
import tempfile
from datetime import datetime
from PIL import Image
import base64
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import pandas as pd

# ==========================================================
# 1. CONFIGURACIÓN DE LA PÁGINA STREAMLIT
# ==========================================================
st.set_page_config(
    page_title="INT – Laboratorio de Función Pulmonar",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado
st.markdown("""
    <style>
    .main-header {
        text-align: center;
        color: #1E3A8A;
        margin-bottom: 1rem;
        font-family: 'Arial', sans-serif;
    }
    .sub-header {
        text-align: center;
        color: #4B5563;
        margin-bottom: 2rem;
        font-family: 'Arial', sans-serif;
    }
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
        font-weight: bold;
        border: none;
        padding: 0.75rem 1.5rem;
        border-radius: 8px;
        font-family: 'Arial', sans-serif;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #2A4BA8;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    .success-box {
        background-color: #D1FAE5;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #10B981;
        margin: 1rem 0;
        font-family: 'Arial', sans-serif;
    }
    .warning-box {
        background-color: #FEF3C7;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #F59E0B;
        margin: 1rem 0;
        font-family: 'Arial', sans-serif;
    }
    .preview-box {
        background-color: #EFF6FF;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px solid #3B82F6;
        margin: 1rem 0;
        font-family: 'Arial', sans-serif;
    }
    .data-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
        font-family: 'Arial', sans-serif;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1E3A8A;
        font-family: 'Arial', sans-serif;
    }
    .tab-content {
        font-family: 'Arial', sans-serif;
    }
    </style>
""", unsafe_allow_html=True)

# Títulos
st.markdown('<h1 class="main-header">🏥 INT – Laboratorio de Función Pulmonar</h1>', unsafe_allow_html=True)
st.markdown('<h3 class="sub-header">Generador de Informes de Óxido Nítrico Exhalado (FeNO)</h3>', unsafe_allow_html=True)

# ==========================================================
# 2. EXTRACCIÓN MEJORADA DE CURVA DE EXHALACIÓN
# ==========================================================
def extraer_curva_exhalacion(pagina):
    """
    Extrae la curva de exhalación del PDF de manera precisa
    """
    try:
        # Obtener todas las imágenes del PDF
        imagenes = pagina.get_images(full=True)
        
        # Buscar la imagen más grande (probablemente la curva)
        max_area = 0
        mejor_imagen = None
        
        for img_index, img_info in enumerate(imagenes):
            try:
                xref = img_info[0]
                img = pagina.parent.extract_image(xref)
                img_bytes = img["image"]
                
                # Calcular área de la imagen
                area = img["width"] * img["height"]
                
                # Buscar la imagen más grande que no sea un logo
                if area > max_area and img["width"] > 100 and img["height"] > 100:
                    # Filtrar imágenes muy pequeñas (posibles íconos)
                    if 20000 < area < 200000:  # Rango típico para gráficos
                        max_area = area
                        mejor_imagen = img_bytes
            except Exception as e:
                continue
        
        if mejor_imagen:
            return mejor_imagen
        
        # Si no encontró por imágenes, buscar por área de pantalla
        texto = pagina.get_text()
        bloques = pagina.get_text("blocks")
        
        # Buscar texto relacionado con curva
        for bloque in bloques:
            x0, y0, x1, y1, texto_bloque, *_ = bloque
            texto_lower = texto_bloque.lower()
            
            if any(palabra in texto_lower for palabra in ['curva', 'exhalación', 'exhalacion', 'graph', 'gráfico']):
                # Expandir área para capturar el gráfico
                rect_curva = fitz.Rect(
                    max(0, x0 - 20),
                    min(pagina.rect.height, y1 + 10),
                    min(pagina.rect.width, x0 + 300),
                    min(pagina.rect.height, y1 + 180)
                )
                
                pix = pagina.get_pixmap(
                    clip=rect_curva,
                    matrix=fitz.Matrix(2.5, 2.5),
                    alpha=False
                )
                return pix.tobytes("png")
        
        # Área por defecto si no encuentra
        rect_curva = fitz.Rect(50, 350, 400, 550)
        pix = pagina.get_pixmap(
            clip=rect_curva,
            matrix=fitz.Matrix(2.5, 2.5),
            alpha=False
        )
        return pix.tobytes("png")
        
    except Exception as e:
        st.warning(f"⚠️ No se pudo extraer la curva: {str(e)}")
        return None

def extraer_datos_pdf_sunvou(pdf_file):
    """
    Extrae datos y gráficos del informe Sunvou con mayor precisión
    """
    try:
        pdf_bytes = pdf_file.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        pagina = doc_pdf[0]
        texto_completo = pagina.get_text()
        
        # Mostrar texto para debugging (opcional)
        if st.session_state.get('debug_mode', False):
            with st.expander("📄 Texto extraído del PDF"):
                st.text(texto_completo[:1000])
        
        # Función mejorada para buscar valores
        def buscar_valor(patron, texto=texto_completo, default="---"):
            try:
                matches = re.findall(patron, texto, re.IGNORECASE | re.MULTILINE | re.DOTALL)
                if matches:
                    valor = str(matches[-1]).strip()
                    # Limpiar el valor
                    valor = re.sub(r'[^\d,.]', '', valor)
                    if valor:
                        return valor.replace(',', '.')
                return default
            except:
                return default
        
        # Patrones mejorados
        datos = {
            "FeNO50": "---",
            "Temperatura": "---",
            "Presion": "---",
            "Flujo": "---",
            "img_curva": None
        }
        
        # Múltiples patrones para cada valor
        patrones_feno = [
            r'FeN[O0]50[:\s]*(\d+[\.,]?\d*)',
            r'FeNO50[:\s]*(\d+[\.,]?\d*)',
            r'Valor de FeNO50[:\s]*(\d+[\.,]?\d*)'
        ]
        
        patrones_temp = [
            r'Temperatura[:\s]*(\d+[\.,]?\d*)',
            r'Temp\.?[:\s]*(\d+[\.,]?\d*)'
        ]
        
        patrones_pres = [
            r'Presión[:\s]*(\d+[\.,]?\d*)',
            r'Pres\.?[:\s]*(\d+[\.,]?\d*)'
        ]
        
        patrones_flujo = [
            r'Tasa de Flujo[:\s]*(\d+[\.,]?\d*)',
            r'Tasa de flujo[:\s]*(\d+[\.,]?\d*)'
        ]
        
        # Buscar con múltiples patrones
        def buscar_con_patrones(patrones):
            for patron in patrones:
                valor = buscar_valor(patron)
                if valor != "---":
                    return valor
            return "---"
        
        datos["FeNO50"] = buscar_con_patrones(patrones_feno)
        datos["Temperatura"] = buscar_con_patrones(patrones_temp)
        datos["Presion"] = buscar_con_patrones(patrones_pres)
        datos["Flujo"] = buscar_con_patrones(patrones_flujo)
        
        # Extraer la curva
        datos["img_curva"] = extraer_curva_exhalacion(pagina)
        
        doc_pdf.close()
        return datos
        
    except Exception as e:
        st.error(f"Error procesando PDF: {str(e)}")
        return None

# ==========================================================
# 3. PROCESAMIENTO DEL DOCUMENTO WORD (CON CURVA)
# ==========================================================
def procesar_documento_word_con_curva(doc_path, datos):
    """
    Reemplaza placeholders en el Word y agrega la curva donde dice CURVA_EXHALA
    """
    try:
        # Cargar el documento Word
        doc = Document(doc_path)
        
        # Reemplazar en párrafos
        for paragraph in doc.paragraphs:
            texto_original = paragraph.text
            
            # Caso especial para CURVA_EXHALA
            if "CURVA_EXHALA" in texto_original and datos.get("img_curva"):
                # Limpiar el párrafo
                paragraph.clear()
                
                # Agregar la imagen de la curva centrada
                run = paragraph.add_run()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run.add_picture(
                    io.BytesIO(datos["img_curva"]),
                    width=Inches(3.7)  # Tamaño similar al original
                )
                continue
            
            # Reemplazar otros placeholders
            for key, value in datos.items():
                if key.startswith("{{") and key.endswith("}}"):
                    placeholder = key
                    if placeholder in texto_original:
                        # Caso especial para FeNO50 - agregar "ppb" en negrita
                        if placeholder == "{{FeNO50}}":
                            paragraph.clear()
                            run = paragraph.add_run(f"{value} ppb")
                            run.font.bold = True
                            run.font.size = Pt(14)
                        else:
                            paragraph.text = texto_original.replace(placeholder, str(value))
        
        # Reemplazar en tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        texto_original = paragraph.text
                        
                        for key, value in datos.items():
                            if key.startswith("{{") and key.endswith("}}"):
                                placeholder = key
                                if placeholder in texto_original:
                                    paragraph.text = texto_original.replace(placeholder, str(value))
        
        # Guardar en archivo temporal
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        doc.save(temp_file.name)
        
        # Leer como bytes
        with open(temp_file.name, 'rb') as f:
            doc_bytes = f.read()
        
        # Limpiar
        os.unlink(temp_file.name)
        
        return doc_bytes
        
    except Exception as e:
        st.error(f"Error procesando Word: {str(e)}")
        return None

# ==========================================================
# 4. GENERACIÓN DE PDF (OPCIONAL)
# ==========================================================
def generar_pdf_simple(datos):
    """
    Genera un PDF simple con los datos (opcional)
    """
    try:
        # Crear archivo temporal
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        pdf_path = temp_pdf.name
        temp_pdf.close()
        
        # Crear PDF
        c = canvas.Canvas(pdf_path, pagesize=letter)
        
        # Título
        c.setFont("Helvetica-Bold", 16)
        c.drawString(100, 750, "INFORME DE PRUEBA FENO")
        c.setFont("Helvetica", 12)
        
        # Datos del paciente
        y_pos = 700
        c.drawString(100, y_pos, f"Paciente: {datos.get('nombre', '')} {datos.get('apellidos', '')}")
        y_pos -= 20
        c.drawString(100, y_pos, f"RUT: {datos.get('rut', '')}")
        y_pos -= 20
        c.drawString(100, y_pos, f"FeNO50: {datos.get('feno50', '')} ppb")
        y_pos -= 20
        c.drawString(100, y_pos, f"Fecha: {datos.get('fecha_examen', '')}")
        
        # Guardar PDF
        c.save()
        
        # Leer bytes
        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()
        
        # Limpiar
        os.unlink(pdf_path)
        
        return pdf_bytes
        
    except Exception as e:
        st.error(f"Error generando PDF: {str(e)}")
        return None

# ==========================================================
# 5. INTERFAZ STREAMLIT MEJORADA
# ==========================================================
def main():
    # Inicializar estado de sesión
    if 'datos_pdf' not in st.session_state:
        st.session_state.datos_pdf = None
    if 'datos_paciente' not in st.session_state:
        st.session_state.datos_paciente = {}
    if 'vista_previa' not in st.session_state:
        st.session_state.vista_previa = False
    
    # Sidebar para controles
    with st.sidebar:
        st.markdown("### ⚙️ Configuración")
        formato_descarga = st.radio(
            "Formato de descarga:",
            ["Word (.docx)", "PDF (.pdf)"],
            index=0
        )
        
        st.markdown("---")
        st.markdown("### 📖 Instrucciones")
        st.markdown("""
        1. Complete datos del paciente
        2. Suba el PDF Sunvou
        3. Extraiga datos automáticamente
        4. Revise la vista previa
        5. Genere y descargue
        """)
        
        # Debug mode
        if st.checkbox("Modo Debug"):
            st.session_state.debug_mode = True
    
    # Pestañas principales
    tab1, tab2, tab3, tab4 = st.tabs(["👤 Datos Paciente", "📄 Cargar PDF", "👁️ Vista Previa", "🚀 Generar"])
    
    with tab1:
        st.markdown("### 👤 Información del Paciente")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            nombre = st.text_input("Nombre *", key="nombre")
            apellidos = st.text_input("Apellidos *", key="apellidos")
            rut = st.text_input("RUT *", key="rut")
            genero = st.selectbox("Género *", ["Seleccione", "Hombre", "Mujer"], key="genero")
            procedencia = st.text_input("Procedencia *", value="Poli", key="procedencia")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            f_nacimiento = st.text_input("Fecha Nacimiento (DD/MM/AAAA) *", key="f_nac")
            edad = st.text_input("Edad (años)", key="edad")
            altura = st.text_input("Altura (cm) *", key="altura")
            peso = st.text_input("Peso (kg) *", key="peso")
            medico = st.text_input("Médico *", key="medico")
            operador = st.text_input("Operador *", value="TM Jorge Espinoza", key="operador")
            fecha_examen = st.date_input("Fecha Examen *", key="fecha_examen")
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Guardar datos del paciente
        if st.button("💾 Guardar Datos Paciente", key="guardar_paciente"):
            st.session_state.datos_paciente = {
                "nombre": nombre,
                "apellidos": apellidos,
                "rut": rut,
                "genero": genero,
                "procedencia": procedencia,
                "f_nacimiento": f_nacimiento,
                "edad": edad,
                "altura": altura,
                "peso": peso,
                "medico": medico,
                "operador": operador,
                "fecha_examen": fecha_examen.strftime("%d/%m/%Y") if fecha_examen else ""
            }
            st.success("✅ Datos del paciente guardados")
    
    with tab2:
        st.markdown("### 📄 Cargar Informe Sunvou")
        
        uploaded_pdf = st.file_uploader(
            "Seleccione el archivo PDF del equipo Sunvou",
            type=["pdf"],
            help="Suba el informe en formato PDF generado por el equipo"
        )
        
        if uploaded_pdf:
            st.success(f"✅ Archivo cargado: {uploaded_pdf.name}")
            st.info(f"📏 Tamaño: {uploaded_pdf.size / 1024:.1f} KB")
            
            # Botón para extraer datos
            if st.button("🔍 Extraer Datos del PDF", type="primary", key="extraer_btn"):
                with st.spinner("Procesando PDF..."):
                    datos_pdf = extraer_datos_pdf_sunvou(uploaded_pdf)
                    
                    if datos_pdf:
                        st.session_state.datos_pdf = datos_pdf
                        
                        # Mostrar datos extraídos
                        st.markdown('<div class="success-box">', unsafe_allow_html=True)
                        st.success("✅ Datos extraídos correctamente")
                        
                        col_val1, col_val2, col_val3, col_val4 = st.columns(4)
                        with col_val1:
                            st.metric("FeNO50", f"{datos_pdf['FeNO50']} ppb")
                        with col_val2:
                            st.metric("Temperatura", f"{datos_pdf['Temperatura']} °C")
                        with col_val3:
                            st.metric("Presión", f"{datos_pdf['Presion']} cmH₂O")
                        with col_val4:
                            st.metric("Flujo", f"{datos_pdf['Flujo']} ml/s")
                        
                        # Mostrar curva si se extrajo
                        if datos_pdf.get("img_curva"):
                            st.markdown("**📈 Curva de Exhalación Extraída:**")
                            st.image(datos_pdf["img_curva"], caption="Gráfico que se insertará en el informe", width=400)
                            st.success("✅ La curva será insertada en 'CURVA_EXHALA'")
                        else:
                            st.warning("⚠️ No se pudo extraer la curva")
                        
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.error("❌ No se pudieron extraer datos del PDF")
    
    with tab3:
        st.markdown("### 👁️ Vista Previa del Informe")
        
        # Verificar que haya datos
        if not st.session_state.datos_paciente:
            st.warning("⚠️ Primero complete los datos del paciente")
        elif not st.session_state.datos_pdf:
            st.warning("⚠️ Primero extraiga los datos del PDF")
        else:
            # Mostrar vista previa
            st.markdown('<div class="preview-box">', unsafe_allow_html=True)
            st.markdown("#### 📋 Resumen del Informe")
            
            # Datos del paciente
            st.markdown("**👤 Datos del Paciente:**")
            col_pre1, col_pre2 = st.columns(2)
            
            with col_pre1:
                dp = st.session_state.datos_paciente
                st.write(f"**Nombre:** {dp.get('nombre', '')} {dp.get('apellidos', '')}")
                st.write(f"**RUT:** {dp.get('rut', '')}")
                st.write(f"**Género:** {dp.get('genero', '')}")
                st.write(f"**Edad:** {dp.get('edad', '')} años")
                st.write(f"**Fecha Nacimiento:** {dp.get('f_nacimiento', '')}")
            
            with col_pre2:
                st.write(f"**Altura:** {dp.get('altura', '')} cm")
                st.write(f"**Peso:** {dp.get('peso', '')} kg")
                st.write(f"**Procedencia:** {dp.get('procedencia', '')}")
                st.write(f"**Médico:** {dp.get('medico', '')}")
                st.write(f"**Operador:** {dp.get('operador', '')}")
                st.write(f"**Fecha Examen:** {dp.get('fecha_examen', '')}")
            
            # Datos técnicos
            st.markdown("**🔬 Datos Técnicos:**")
            col_tec1, col_tec2, col_tec3, col_tec4 = st.columns(4)
            
            with col_tec1:
                st.info(f"**FeNO50:** {st.session_state.datos_pdf.get('FeNO50', '---')} **ppb**")
            with col_tec2:
                st.info(f"**Temperatura:** {st.session_state.datos_pdf.get('Temperatura', '---')} °C")
            with col_tec3:
                st.info(f"**Presión:** {st.session_state.datos_pdf.get('Presion', '---')} cmH₂O")
            with col_tec4:
                st.info(f"**Flujo:** {st.session_state.datos_pdf.get('Flujo', '---')} ml/s")
            
            # Curva de exhalación
            if st.session_state.datos_pdf.get("img_curva"):
                st.markdown("**📈 Curva de Exhalación:**")
                st.image(st.session_state.datos_pdf["img_curva"], 
                        caption="Esta curva se insertará en el informe", 
                        width=350)
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Botón para activar vista previa
            if st.button("✅ Confirmar Vista Previa", type="primary"):
                st.session_state.vista_previa = True
                st.success("✅ Vista previa confirmada. Puede proceder a generar el informe.")
    
    with tab4:
        st.markdown("### 🚀 Generar y Descargar Informe")
        
        # Verificar condiciones
        condiciones = [
            ("Datos del paciente completos", bool(st.session_state.datos_paciente)),
            ("Datos del PDF extraídos", bool(st.session_state.datos_pdf)),
            ("Vista previa confirmada", st.session_state.vista_previa)
        ]
        
        # Mostrar estado
        st.markdown("**Estado del sistema:**")
        for condicion, estado in condiciones:
            if estado:
                st.success(f"✅ {condicion}")
            else:
                st.warning(f"⚠️ {condicion}")
        
        if all(estado for _, estado in condiciones):
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.success("🎉 ¡Todo listo para generar el informe!")
            
            # Botones de generación
            col_gen1, col_gen2 = st.columns(2)
            
            with col_gen1:
                if st.button("📄 Generar Documento Word", type="primary", use_container_width=True):
                    with st.spinner("Generando Word..."):
                        try:
                            # Preparar datos completos
                            datos_completos = {
                                # Datos del paciente
                                "{{NOMBRE}}": st.session_state.datos_paciente.get("nombre", ""),
                                "{{APELLIDOS}}": st.session_state.datos_paciente.get("apellidos", ""),
                                "{{RUT}}": st.session_state.datos_paciente.get("rut", ""),
                                "{{GENERO}}": st.session_state.datos_paciente.get("genero", ""),
                                "{{PROCEDENCIA}}": st.session_state.datos_paciente.get("procedencia", ""),
                                "{{F_NACIMIENTO}}": st.session_state.datos_paciente.get("f_nacimiento", ""),
                                "{{EDAD}}": st.session_state.datos_paciente.get("edad", ""),
                                "{{ALTURA}}": st.session_state.datos_paciente.get("altura", ""),
                                "{{PESO}}": st.session_state.datos_paciente.get("peso", ""),
                                "{{MEDICO}}": st.session_state.datos_paciente.get("medico", ""),
                                "{{OPERADOR}}": st.session_state.datos_paciente.get("operador", ""),
                                "{{FECHA_EXAMEN}}": st.session_state.datos_paciente.get("fecha_examen", ""),
                                
                                # Datos técnicos
                                "{{FeNO50}}": st.session_state.datos_pdf.get("FeNO50", "---"),
                                "{{Temperatura}}": st.session_state.datos_pdf.get("Temperatura", "---"),
                                "{{Presion}}": st.session_state.datos_pdf.get("Presion", "---"),
                                "{{Tasa de flujo}}": st.session_state.datos_pdf.get("Flujo", "---"),
                                
                                # Imagen de la curva
                                "img_curva": st.session_state.datos_pdf.get("img_curva")
                            }
                            
                            # Buscar plantilla
                            plantilla_path = None
                            for ruta in ["FeNO 50 Informe.docx", "plantillas/FeNO 50 Informe.docx"]:
                                if os.path.exists(ruta):
                                    plantilla_path = ruta
                                    break
                            
                            if not plantilla_path:
                                st.error("❌ No se encuentra la plantilla Word")
                                return
                            
                            # Procesar documento
                            doc_bytes = procesar_documento_word_con_curva(plantilla_path, datos_completos)
                            
                            if doc_bytes:
                                # Nombre del archivo
                                nombre_archivo = f"INFORME_FENO_{st.session_state.datos_paciente.get('rut', '')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                                
                                # Botón de descarga
                                st.download_button(
                                    label="⬇️ Descargar Word",
                                    data=doc_bytes,
                                    file_name=nombre_archivo,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                            
                        except Exception as e:
                            st.error(f"❌ Error generando Word: {str(e)}")
            
            with col_gen2:
                if st.button("📊 Generar PDF", type="secondary", use_container_width=True):
                    with st.spinner("Generando PDF..."):
                        try:
                            # Combinar datos para PDF
                            datos_pdf_combinados = {
                                **st.session_state.datos_paciente,
                                "feno50": st.session_state.datos_pdf.get("FeNO50", "---")
                            }
                            
                            # Generar PDF
                            pdf_bytes = generar_pdf_simple(datos_pdf_combinados)
                            
                            if pdf_bytes:
                                nombre_archivo = f"INFORME_FENO_{st.session_state.datos_paciente.get('rut', '')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                                
                                st.download_button(
                                    label="⬇️ Descargar PDF",
                                    data=pdf_bytes,
                                    file_name=nombre_archivo,
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                            
                        except Exception as e:
                            st.error(f"❌ Error generando PDF: {str(e)}")
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Información adicional
            st.markdown("""
            **📋 Características del informe:**
            - ✅ **Formato profesional** del INT
            - ✅ **FeNO50 en negrita** con "ppb"
            - ✅ **Curva de exhalación** insertada
            - ✅ **Fuente Arial** en todo el documento
            - ✅ **Datos completos** del paciente
            """)
        else:
            st.warning("ℹ️ Complete todos los pasos anteriores para generar el informe")

# ==========================================================
# 6. EJECUCIÓN PRINCIPAL
# ==========================================================
if __name__ == "__main__":
    main()
    
    # Pie de página
    st.markdown("---")
    st.caption("© Instituto Nacional del Tórax - Laboratorio de Función Pulmonar | Sistema de Informes FeNO v2.0")

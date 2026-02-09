import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

# ==========================================================
# 1. CONFIGURACIÓN STREAMLIT
# ==========================================================
st.set_page_config(
    page_title="INT - Generador de Informes FeNO",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado para mejor apariencia
st.markdown("""
<style>
    .main-title {
        text-align: center;
        color: #1E3A8A;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    .sub-title {
        text-align: center;
        color: #4B5563;
        font-size: 1.2rem;
        margin-bottom: 2rem;
    }
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
        font-weight: bold;
        border: none;
        padding: 0.75rem 1.5rem;
        border-radius: 8px;
        transition: all 0.3s;
        font-family: 'Arial', sans-serif;
    }
    .stButton>button:hover {
        background-color: #2A4BA8;
        transform: translateY(-2px);
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
    .success-box {
        background-color: #D1FAE5;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #10B981;
        margin: 1rem 0;
        font-family: 'Arial', sans-serif;
    }
    .warning-box {
        background-color: #FEE2E2;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #EF4444;
        margin: 1rem 0;
        font-family: 'Arial', sans-serif;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1E3A8A;
        font-family: 'Arial', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.markdown('<h1 class="main-title">🏥 INT - Laboratorio de Función Pulmonar</h1>', unsafe_allow_html=True)
st.markdown('<h3 class="sub-title">Generador de Informes de Óxido Nítrico Exhalado (FeNO)</h3>', unsafe_allow_html=True)

# ==========================================================
# 2. EXTRACCIÓN DE DATOS DEL PDF SUNVOU
# ==========================================================
def extraer_datos_del_pdf(pdf_file):
    """
    Extrae datos específicos del PDF del equipo Sunvou
    """
    try:
        # Leer el PDF
        pdf_bytes = pdf_file.read()
        doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
        pagina = doc_pdf[0]
        texto_completo = pagina.get_text()
        
        # Función mejorada para buscar valores
        def buscar_valor_seguro(patrones, texto=texto_completo):
            for patron in patrones:
                try:
                    match = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
                    if match:
                        valor = match.group(1).strip()
                        # Limpiar el valor
                        valor = re.sub(r'[^\d,\.]', '', valor)
                        if valor:
                            return valor.replace(',', '.')
                except:
                    continue
            return "---"
        
        # Definir patrones para cada valor
        patrones_feno = [
            r'FeN[O0]50[:\s]*(\d+[\.,]?\d*)',
            r'FeNO50[:\s]*(\d+[\.,]?\d*)',
            r'Valor de FeNO50[:\s]*(\d+[\.,]?\d*)',
            r'FeNO\s*50[:\s]*(\d+[\.,]?\d*)'
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
        
        # Extraer los valores
        datos = {
            "FeNO50": buscar_valor_seguro(patrones_feno),
            "Temperatura": buscar_valor_seguro(patrones_temp),
            "Presion": buscar_valor_seguro(patrones_pres),
            "Tasa de flujo": buscar_valor_seguro(patrones_flujo),
            "img_curva": extraer_curva_exhalacion(pagina)
        }
        
        doc_pdf.close()
        return datos
        
    except Exception as e:
        st.error(f"❌ Error procesando PDF: {str(e)}")
        return None

def extraer_curva_exhalacion(pagina):
    """
    Extrae la imagen de la curva de exhalación del PDF
    """
    try:
        # Buscar texto relacionado con la curva
        texto = pagina.get_text()
        bloques = pagina.get_text("blocks")
        
        # Coordenadas predeterminadas
        x0, y0, x1, y1 = 50, 350, 350, 500
        
        # Buscar bloques con palabras clave
        for bloque in bloques:
            bx0, by0, bx1, by1, btexto, *_ = bloque
            texto_lower = btexto.lower()
            if any(palabra in texto_lower for palabra in ['curva', 'exhalación', 'exhalacion', 'graph']):
                x0, y0, x1, y1 = bx0 - 10, by1 + 5, bx0 + 280, by1 + 130
                break
        
        # Capturar la imagen
        rect_curva = fitz.Rect(x0, y0, x1, y1)
        pix = pagina.get_pixmap(
            clip=rect_curva,
            matrix=fitz.Matrix(2.5, 2.5),
            alpha=False
        )
        
        return pix.tobytes("png")
        
    except Exception as e:
        st.warning(f"⚠️ No se pudo extraer la curva: {str(e)}")
        return None

# ==========================================================
# 3. FUNCIÓN PARA APLICAR FUENTE ARIAL A TODO EL DOCUMENTO
# ==========================================================
def aplicar_fuente_arial_a_todo(doc):
    """
    Aplica fuente Arial a todo el documento Word
    """
    try:
        # Aplicar Arial a todos los estilos
        for style in doc.styles:
            try:
                style.font.name = 'Arial'
            except:
                pass
        
        # Aplicar Arial a todos los párrafos
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Arial'
        
        # Aplicar Arial a todas las tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
        
        return doc
    except Exception as e:
        st.warning(f"⚠️ No se pudo aplicar fuente Arial: {str(e)}")
        return doc

# ==========================================================
# 4. PROCESAMIENTO DEL DOCUMENTO WORD
# ==========================================================
def reemplazar_placeholders_en_documento(doc_path, datos):
    """
    Reemplaza los placeholders {{...}} en el documento Word manteniendo el formato original
    """
    try:
        # Cargar el documento Word
        doc = Document(doc_path)
        
        # Aplicar fuente Arial a todo el documento
        doc = aplicar_fuente_arial_a_todo(doc)
        
        # PRIMERO: Reemplazar en tablas (datos del paciente)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        texto_original = paragraph.text
                        
                        # Reemplazar placeholders
                        for key, value in datos.items():
                            if key.startswith("{{") and key.endswith("}}"):
                                placeholder = key
                                if placeholder in texto_original:
                                    # Limpiar el párrafo primero
                                    paragraph.clear()
                                    
                                    # Agregar el nuevo texto con formato
                                    run = paragraph.add_run(str(value))
                                    run.font.name = 'Arial'
                                    run.font.size = Pt(11)
        
        # SEGUNDO: Reemplazar en párrafos (datos técnicos y curva)
        for paragraph in doc.paragraphs:
            texto_original = paragraph.text
            
            # Caso especial para CURVA_EXHALA
            if "CURVA_EXHALA" in texto_original and datos.get("img_curva"):
                # Limpiar el párrafo
                paragraph.clear()
                
                # Agregar título de la curva
                run_titulo = paragraph.add_run("Curva de Exhalación FeNO50:")
                run_titulo.font.name = 'Arial'
                run_titulo.font.size = Pt(11)
                run_titulo.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Agregar un salto de línea
                paragraph.add_run().add_break()
                
                # Agregar la imagen de la curva
                run_img = paragraph.add_run()
                run_img.add_picture(
                    io.BytesIO(datos["img_curva"]),
                    width=Inches(3.7)
                )
                continue
            
            # Reemplazar datos técnicos en párrafos
            for key, value in datos.items():
                if key.startswith("{{") and key.endswith("}}"):
                    placeholder = key
                    if placeholder in texto_original:
                        # Caso especial para FeNO50 - agregar "ppb" en negrita
                        if placeholder == "{{FeNO50}}":
                            # Limpiar el párrafo
                            paragraph.clear()
                            
                            # Separar el texto antes y después del placeholder
                            parts = texto_original.split(placeholder)
                            if len(parts) >= 2:
                                # Agregar texto antes del placeholder
                                if parts[0].strip():
                                    run_before = paragraph.add_run(parts[0])
                                    run_before.font.name = 'Arial'
                                    run_before.font.size = Pt(11)
                                
                                # Agregar valor FeNO50 en negrita + " ppb"
                                run_value = paragraph.add_run(f"{value} ppb")
                                run_value.font.name = 'Arial'
                                run_value.font.size = Pt(11)
                                run_value.bold = True
                                
                                # Agregar texto después del placeholder
                                if parts[1].strip():
                                    run_after = paragraph.add_run(parts[1])
                                    run_after.font.name = 'Arial'
                                    run_after.font.size = Pt(11)
                        else:
                            # Para otros placeholders, reemplazo normal
                            paragraph.text = paragraph.text.replace(placeholder, str(value))
        
        # TERCERO: Buscar y reemplazar FeNO50 en cualquier parte del documento
        for paragraph in doc.paragraphs:
            texto = paragraph.text
            if "{{FeNO50}}" in texto:
                # Reemplazar con formato especial
                paragraph.clear()
                # Buscar el contexto (por si está en una tabla o lista)
                if "FeNO" in texto:
                    # Mantener el formato "FeNO50" con superíndice si existe
                    if "FeNO~50~" in texto:
                        run_feno = paragraph.add_run("FeNO")
                        run_feno.font.name = 'Arial'
                        run_feno.bold = True
                        
                        run_super = paragraph.add_run("50")
                        run_super.font.name = 'Arial'
                        run_super.font.superscript = True
                        run_super.bold = True
                        
                        run_value = paragraph.add_run(f" {datos['{{FeNO50}}']} ppb")
                        run_value.font.name = 'Arial'
                        run_value.bold = True
                        run_value.font.size = Pt(12)
                    else:
                        run = paragraph.add_run(f"FeNO50: {datos['{{FeNO50}}']} ppb")
                        run.font.name = 'Arial'
                        run.bold = True
        
        # Guardar en un archivo temporal
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        doc.save(temp_file.name)
        
        # Leer el archivo como bytes
        with open(temp_file.name, 'rb') as f:
            doc_bytes = f.read()
        
        # Limpiar archivo temporal
        os.unlink(temp_file.name)
        
        return doc_bytes
        
    except Exception as e:
        st.error(f"❌ Error procesando documento Word: {str(e)}")
        if 'datos' in locals():
            st.write("Datos disponibles:", datos)
        return None

# ==========================================================
# 5. INTERFAZ PRINCIPAL DE LA APLICACIÓN
# ==========================================================
def main():
    # Sidebar para instrucciones
    with st.sidebar:
        st.markdown("### 📋 Instrucciones")
        st.markdown("""
        1. **Complete datos del paciente**
        2. **Suba el PDF del equipo Sunvou**
        3. **Extraiga datos automáticamente**
        4. **Genere el informe final**
        
        ---
        **Requisitos:**
        - PDF del equipo Sunvou
        - Datos completos del paciente
        - Plantilla Word original
        
        **Formato final:**
        - Fuente: Arial en todo el documento
        - FeNO50: Valor en negrita + "ppb"
        - Curva: Insertada automáticamente
        """)
        
        st.markdown("---")
        st.caption("Versión 1.1 - INT Laboratorio")
    
    # Crear pestañas principales
    tab1, tab2, tab3 = st.tabs(["👤 Datos del Paciente", "📄 Carga de PDF", "🎯 Generar Informe"])
    
    with tab1:
        st.markdown("### 👤 Información del Paciente")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            nombre = st.text_input("Nombre *", placeholder="Ej: Juan", key="nombre")
            apellidos = st.text_input("Apellidos *", placeholder="Ej: Pérez González", key="apellidos")
            rut = st.text_input("RUT *", placeholder="Ej: 12.345.678-9", key="rut")
            genero = st.selectbox("Género *", ["Seleccione", "Hombre", "Mujer"], key="genero")
            procedencia = st.text_input("Procedencia *", value="Poli", key="procedencia")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            f_nacimiento = st.text_input("Fecha de Nacimiento *", placeholder="DD/MM/AAAA", key="f_nac")
            edad = st.text_input("Edad", placeholder="Ej: 35", key="edad")
            altura = st.text_input("Altura (cm) *", placeholder="Ej: 175", key="altura")
            peso = st.text_input("Peso (kg) *", placeholder="Ej: 70", key="peso")
            medico = st.text_input("Médico Solicitante *", placeholder="Ej: Dr. Carlos López", key="medico")
            operador = st.text_input("Operador *", value="TM Jorge Espinoza", key="operador")
            fecha_examen = st.date_input("Fecha del Examen *", key="fecha_examen")
            st.markdown('</div>', unsafe_allow_html=True)
    
    with tab2:
        st.markdown("### 📄 Cargar Informe del Equipo Sunvou")
        
        # Subida del PDF
        uploaded_pdf = st.file_uploader(
            "Seleccione el archivo PDF generado por el equipo Sunvou",
            type=["pdf"],
            help="Suba el informe en formato PDF",
            key="pdf_uploader"
        )
        
        if uploaded_pdf:
            st.success(f"✅ Archivo cargado: {uploaded_pdf.name}")
            st.info(f"📏 Tamaño: {uploaded_pdf.size / 1024:.1f} KB")
            
            # Botón para extraer datos
            if st.button("🔍 Extraer Datos del PDF", type="primary", key="extraer_btn"):
                with st.spinner("Procesando PDF..."):
                    datos_pdf = extraer_datos_del_pdf(uploaded_pdf)
                    
                    if datos_pdf:
                        # Guardar en session state
                        st.session_state.datos_pdf = datos_pdf
                        
                        # Mostrar datos extraídos
                        st.markdown('<div class="success-box">', unsafe_allow_html=True)
                        st.success("✅ Datos extraídos correctamente")
                        
                        col_val1, col_val2, col_val3, col_val4 = st.columns(4)
                        with col_val1:
                            st.markdown('<div class="metric-value">', unsafe_allow_html=True)
                            st.metric("FeNO50", f"{datos_pdf['FeNO50']} ppb")
                            st.markdown('</div>', unsafe_allow_html=True)
                        with col_val2:
                            st.metric("Temperatura", f"{datos_pdf['Temperatura']} °C")
                        with col_val3:
                            st.metric("Presión", f"{datos_pdf['Presion']} cmH₂O")
                        with col_val4:
                            st.metric("Flujo", f"{datos_pdf['Tasa de flujo']} ml/s")
                        
                        # Mostrar curva si se extrajo
                        if datos_pdf.get("img_curva"):
                            st.markdown("**📈 Curva de Exhalación Extraída:**")
                            st.image(datos_pdf["img_curva"], use_column_width=True)
                        else:
                            st.warning("⚠️ No se pudo extraer la curva de exhalación")
                        
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                        st.error("❌ No se pudieron extraer datos del PDF")
                        st.markdown("""
                        **Posibles causas:**
                        - El PDF no tiene el formato esperado
                        - El PDF está protegido o es una imagen
                        - Los datos no están en texto legible
                        """)
                        st.markdown('</div>', unsafe_allow_html=True)
    
    with tab3:
        st.markdown("### 🎯 Generar Informe Final")
        
        # Verificar que todos los datos estén listos
        datos_requeridos = [
            ("Nombre", nombre),
            ("Apellidos", apellidos),
            ("RUT", rut),
            ("Género", genero),
            ("Médico", medico),
            ("Operador", operador),
            ("Fecha Nacimiento", f_nacimiento),
            ("Altura", altura),
            ("Peso", peso),
            ("Procedencia", procedencia)
        ]
        
        faltantes = [dato[0] for dato in datos_requeridos if not dato[1] or dato[1] == "Seleccione"]
        
        if not hasattr(st.session_state, 'datos_pdf') or not st.session_state.datos_pdf:
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.error("❌ Primero debe extraer los datos del PDF (pestaña 'Carga de PDF')")
            st.markdown('</div>', unsafe_allow_html=True)
        elif faltantes:
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.error(f"❌ Faltan datos obligatorios del paciente: {', '.join(faltantes)}")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            # Mostrar resumen de datos
            st.markdown("### 📋 Resumen de Datos para el Informe")
            
            col_res1, col_res2 = st.columns(2)
            
            with col_res1:
                st.markdown('<div class="data-card">', unsafe_allow_html=True)
                st.write(f"**👤 Paciente:** {nombre} {apellidos}")
                st.write(f"**📋 RUT:** {rut}")
                st.write(f"**⚧️ Género:** {genero}")
                st.write(f"**🎂 Edad:** {edad} años")
                st.write(f"**📅 Fecha Nacimiento:** {f_nacimiento}")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_res2:
                st.markdown('<div class="data-card">', unsafe_allow_html=True)
                st.write(f"**📏 Altura:** {altura} cm")
                st.write(f"**⚖️ Peso:** {peso} kg")
                st.write(f"**🏥 Procedencia:** {procedencia}")
                st.write(f"**👨‍⚕️ Médico:** {medico}")
                st.write(f"**👨‍💼 Operador:** {operador}")
                st.write(f"**📅 Fecha Examen:** {fecha_examen.strftime('%d/%m/%Y')}")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # Mostrar datos técnicos
            st.markdown("### 🔬 Datos Técnicos Extraídos")
            
            col_tec1, col_tec2, col_tec3, col_tec4 = st.columns(4)
            with col_tec1:
                st.info(f"**FeNO50:** {st.session_state.datos_pdf.get('FeNO50', '---')} ppb")
            with col_tec2:
                st.info(f"**Temperatura:** {st.session_state.datos_pdf.get('Temperatura', '---')} °C")
            with col_tec3:
                st.info(f"**Presión:** {st.session_state.datos_pdf.get('Presion', '---')} cmH₂O")
            with col_tec4:
                st.info(f"**Flujo:** {st.session_state.datos_pdf.get('Tasa de flujo', '---')} ml/s")
            
            # Botón para generar informe
            st.markdown("---")
            if st.button("🚀 GENERAR INFORME COMPLETO", type="primary", use_container_width=True):
                with st.spinner("🔄 Generando informe con formato Arial..."):
                    try:
                        # Preparar todos los datos para reemplazar
                        datos_completos = {
                            # Datos del paciente
                            "{{NOMBRE}}": nombre,
                            "{{APELLIDOS}}": apellidos,
                            "{{RUT}}": rut,
                            "{{GENERO}}": genero,
                            "{{PROCEDENCIA}}": procedencia,
                            "{{F_NACIMIENTO}}": f_nacimiento,
                            "{{EDAD}}": edad,
                            "{{ALTURA}}": altura,
                            "{{PESO}}": peso,
                            "{{MEDICO}}": medico,
                            "{{OPERADOR}}": operador,
                            "{{FECHA_EXAMEN}}": fecha_examen.strftime('%d/%m/%Y'),
                            
                            # Datos técnicos del PDF (CON "ppb" PARA FeNO50)
                            "{{FeNO50}}": f"{st.session_state.datos_pdf.get('FeNO50', '---')}",
                            "{{Temperatura}}": st.session_state.datos_pdf.get("Temperatura", "---"),
                            "{{Presion}}": st.session_state.datos_pdf.get("Presion", "---"),
                            "{{Tasa de flujo}}": st.session_state.datos_pdf.get("Tasa de flujo", "---"),
                            
                            # Imagen de la curva
                            "img_curva": st.session_state.datos_pdf.get("img_curva")
                        }
                        
                        # Buscar la plantilla Word
                        posibles_rutas = [
                            "plantillas/FeNO 50 Informe.docx",
                            "FeNO 50 Informe.docx",
                            "plantillas/FeNO50 Informe.docx",
                            "FeNO50 Informe.docx"
                        ]
                        
                        plantilla_path = None
                        for ruta in posibles_rutas:
                            if os.path.exists(ruta):
                                plantilla_path = ruta
                                break
                        
                        if not plantilla_path:
                            st.error("❌ No se encuentra la plantilla Word")
                            st.info("""
                            **Coloque la plantilla en una de estas ubicaciones:**
                            1. En la carpeta actual: `FeNO 50 Informe.docx`
                            2. En carpeta plantillas: `plantillas/FeNO 50 Informe.docx`
                            """)
                            return
                        
                        st.info(f"📄 Usando plantilla: {plantilla_path}")
                        
                        # Procesar el documento
                        doc_bytes = reemplazar_placeholders_en_documento(plantilla_path, datos_completos)
                        
                        if doc_bytes:
                            # Crear nombre de archivo
                            nombre_simple = f"{nombre}_{apellidos}".replace(" ", "_")
                            nombre_archivo = f"INFORME_FENO_{nombre_simple}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                            
                            # Mostrar éxito
                            st.markdown('<div class="success-box">', unsafe_allow_html=True)
                            st.success("🎉 ¡Informe generado exitosamente!")
                            
                            # Información sobre el formato
                            st.info("""
                            **✅ Características del informe generado:**
                            - ✅ **Fuente Arial** en todo el documento
                            - ✅ **FeNO50 en negrita** con "ppb"
                            - ✅ **Curva de exhalación** insertada
                            - ✅ **Formato original** mantenido
                            - ✅ **Datos del paciente** completos
                            """)
                            
                            # Botón de descarga
                            st.download_button(
                                label="📥 DESCARGAR INFORME EN WORD",
                                data=doc_bytes,
                                file_name=nombre_archivo,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary",
                                use_container_width=True
                            )
                            
                            # Instrucciones finales
                            st.markdown("""
                            **📋 Pasos siguientes:**
                            1. **Abra** el documento descargado en Microsoft Word
                            2. **Verifique** que todos los datos estén correctos
                            3. **Confirme** que la fuente es Arial en todo el documento
                            4. **Guarde como PDF** (Archivo → Guardar como → PDF)
                            5. **Envíe** el PDF al médico y archive
                            """)
                            
                            st.markdown('</div>', unsafe_allow_html=True)
                        else:
                            st.error("❌ Error al generar el documento Word")
                        
                    except Exception as e:
                        st.error(f"❌ Error generando el informe: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())

# ==========================================================
# 6. EJECUCIÓN PRINCIPAL
# ==========================================================
if __name__ == "__main__":
    # Inicializar session state si no existe
    if 'datos_pdf' not in st.session_state:
        st.session_state.datos_pdf = None
    
    # Ejecutar aplicación
    main()
    
    # Pie de página
    st.markdown("---")
    st.caption("© Instituto Nacional del Tórax - Laboratorio de Función Pulmonar | Sistema de Informes FeNO v1.1 | Fuente: Arial")

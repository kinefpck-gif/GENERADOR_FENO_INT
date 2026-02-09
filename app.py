import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import re
import io
import os
import tempfile
from datetime import datetime
from pathlib import Path
import base64

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
        padding: 1rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 2rem;
    }
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: bold;
        border: none;
        padding: 0.75rem 1.5rem;
        border-radius: 8px;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }
    .data-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
    }
    .success-box {
        background: linear-gradient(135deg, #d4fc79 0%, #96e6a1 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #10B981;
        margin: 1rem 0;
    }
    .warning-box {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #EF4444;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.markdown('<h1 class="main-title">🏥 INT - Laboratorio de Función Pulmonar</h1>', unsafe_allow_html=True)
st.markdown('<h3 style="text-align: center; color: #4B5563;">Generador de Informes de Óxido Nítrico Exhalado (FeNO)</h3>', unsafe_allow_html=True)

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
            """
            Busca un valor usando múltiples patrones y retorna el primero encontrado
            """
            for patron in patrones:
                try:
                    match = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
                    if match:
                        valor = match.group(1).strip()
                        # Limpiar el valor (mantener solo números y punto decimal)
                        valor = re.sub(r'[^\d,\.]', '', valor)
                        if valor:
                            # Reemplazar coma por punto si es necesario
                            return valor.replace(',', '.')
                except:
                    continue
            return "---"
        
        # Definir patrones para cada valor (múltiples variaciones posibles)
        patrones_feno = [
            r'FeN[O0]50[:\s]*(\d+[\.,]?\d*)',
            r'FeNO50[:\s]*(\d+[\.,]?\d*)',
            r'Valor de FeNO50[:\s]*(\d+[\.,]?\d*)',
            r'FeNO\s*50[:\s]*(\d+[\.,]?\d*)',
            r'FeNO\s*50[:\s]*(\d+)'
        ]
        
        patrones_temp = [
            r'Temperatura[:\s]*(\d+[\.,]?\d*)',
            r'Temp\.?[:\s]*(\d+[\.,]?\d*)',
            r'°C[:\s]*(\d+[\.,]?\d*)'
        ]
        
        patrones_pres = [
            r'Presión[:\s]*(\d+[\.,]?\d*)',
            r'Pres\.?[:\s]*(\d+[\.,]?\d*)',
            r'cmH2O[:\s]*(\d+[\.,]?\d*)'
        ]
        
        patrones_flujo = [
            r'Tasa de Flujo[:\s]*(\d+[\.,]?\d*)',
            r'Tasa de flujo[:\s]*(\d+[\.,]?\d*)',
            r'Flujo[:\s]*(\d+[\.,]?\d*)',
            r'ml/s[:\s]*(\d+[\.,]?\d*)'
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
        
        # Coordenadas predeterminadas (ajustar según sea necesario)
        x0, y0, x1, y1 = 50, 350, 350, 500
        
        # Buscar bloques con palabras clave de curva
        for bloque in bloques:
            bx0, by0, bx1, by1, btexto, *_ = bloque
            if any(palabra in btexto.lower() for palabra in ['curva', 'exhalación', 'exhalacion', 'graph', 'plot']):
                # Ajustar coordenadas basadas en el texto encontrado
                x0, y0, x1, y1 = bx0 - 10, by1 + 5, bx0 + 280, by1 + 130
                break
        
        # Capturar la imagen
        rect_curva = fitz.Rect(x0, y0, x1, y1)
        pix = pagina.get_pixmap(
            clip=rect_curva,
            matrix=fitz.Matrix(2.5, 2.5),  # Buena resolución
            alpha=False
        )
        
        return pix.tobytes("png")
        
    except Exception as e:
        st.warning(f"⚠️ No se pudo extraer la curva: {str(e)}")
        return None

# ==========================================================
# 3. PROCESAMIENTO DEL DOCUMENTO WORD (MANTENIENDO FORMATO)
# ==========================================================
def reemplazar_placeholders_en_documento(doc_path, datos):
    """
    Reemplaza los placeholders {{...}} en el documento Word manteniendo el formato original
    """
    try:
        # Cargar el documento Word
        doc = Document(doc_path)
        
        # Reemplazar en párrafos
        for paragraph in doc.paragraphs:
            texto_original = paragraph.text
            
            # Reemplazar todos los placeholders encontrados
            for key, value in datos.items():
                if key.startswith("{{") and key.endswith("}}"):
                    placeholder = key
                    if placeholder in texto_original:
                        # Reemplazar manteniendo el formato
                        paragraph.text = paragraph.text.replace(placeholder, str(value))
            
            # Caso especial para CURVA_EXHALA
            if "CURVA_EXHALA" in texto_original and datos.get("img_curva"):
                # Limpiar el texto del placeholder
                paragraph.text = paragraph.text.replace("CURVA_EXHALA", "")
                
                # Agregar la imagen de la curva
                run = paragraph.add_run()
                run.add_picture(
                    io.BytesIO(datos["img_curva"]),
                    width=Inches(3.7)  # Tamaño según el documento original
                )
        
        # Reemplazar en tablas
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
                                    paragraph.text = paragraph.text.replace(placeholder, str(value))
        
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
        return None

# ==========================================================
# 4. INTERFAZ PRINCIPAL DE LA APLICACIÓN
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
        """)
        
        # Mostrar versión
        st.markdown("---")
        st.caption("Versión 1.0 - INT Laboratorio")
    
    # Crear pestañas principales
    tab1, tab2, tab3 = st.tabs(["👤 Datos del Paciente", "📄 Carga de PDF", "🎯 Generar Informe"])
    
    with tab1:
        st.markdown("### 👤 Información del Paciente")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            nombre = st.text_input("Nombre *", placeholder="Ej: Juan")
            apellidos = st.text_input("Apellidos *", placeholder="Ej: Pérez González")
            rut = st.text_input("RUT *", placeholder="Ej: 12.345.678-9")
            genero = st.selectbox("Género *", ["Seleccione", "Hombre", "Mujer"])
            procedencia = st.text_input("Procedencia *", value="Poli")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            f_nacimiento = st.text_input("Fecha de Nacimiento *", placeholder="DD/MM/AAAA")
            edad = st.text_input("Edad", placeholder="Ej: 35")
            altura = st.text_input("Altura (cm) *", placeholder="Ej: 175")
            peso = st.text_input("Peso (kg) *", placeholder="Ej: 70")
            medico = st.text_input("Médico Solicitante *", placeholder="Ej: Dr. Carlos López")
            operador = st.text_input("Operador *", value="TM Jorge Espinoza")
            fecha_examen = st.date_input("Fecha del Examen *")
            st.markdown('</div>', unsafe_allow_html=True)
    
    with tab2:
        st.markdown("### 📄 Cargar Informe del Equipo Sunvou")
        
        # Subida del PDF
        uploaded_pdf = st.file_uploader(
            "Seleccione el archivo PDF generado por el equipo Sunvou",
            type=["pdf"],
            help="Suba el informe en formato PDF"
        )
        
        if uploaded_pdf:
            st.success(f"✅ Archivo cargado: {uploaded_pdf.name}")
            
            # Botón para extraer datos
            if st.button("🔍 Extraer Datos del PDF", type="primary"):
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
                            st.metric("FeNO50", f"{datos_pdf['FeNO50']} ppb")
                        with col_val2:
                            st.metric("Temperatura", f"{datos_pdf['Temperatura']} °C")
                        with col_val3:
                            st.metric("Presión", f"{datos_pdf['Presion']} cmH2O")
                        with col_val4:
                            st.metric("Flujo", f"{datos_pdf['Tasa de flujo']} ml/s")
                        
                        # Mostrar curva si se extrajo
                        if datos_pdf.get("img_curva"):
                            st.image(datos_pdf["img_curva"], caption="Curva de Exhalación Extraída", width=400)
                        
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                        st.error("❌ No se pudieron extraer datos del PDF")
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
        
        if not hasattr(st.session_state, 'datos_pdf'):
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.error("❌ Primero debe extraer los datos del PDF")
            st.markdown('</div>', unsafe_allow_html=True)
        elif faltantes:
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.error(f"❌ Faltan datos obligatorios: {', '.join(faltantes)}")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            # Mostrar resumen de datos
            st.markdown("### 📋 Resumen de Datos")
            
            col_res1, col_res2 = st.columns(2)
            
            with col_res1:
                st.info(f"**Paciente:** {nombre} {apellidos}")
                st.info(f"**RUT:** {rut}")
                st.info(f"**Género:** {genero}")
                st.info(f"**Edad:** {edad} años")
                st.info(f"**Fecha Nacimiento:** {f_nacimiento}")
            
            with col_res2:
                st.info(f"**Altura:** {altura} cm")
                st.info(f"**Peso:** {peso} kg")
                st.info(f"**Procedencia:** {procedencia}")
                st.info(f"**Médico:** {medico}")
                st.info(f"**Operador:** {operador}")
                st.info(f"**Fecha Examen:** {fecha_examen.strftime('%d/%m/%Y')}")
            
            # Botón para generar informe
            if st.button("🚀 GENERAR INFORME COMPLETO", type="primary", use_container_width=True):
                with st.spinner("🔄 Generando informe..."):
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
                            
                            # Datos técnicos del PDF
                            "{{FeNO50}}": st.session_state.datos_pdf.get("FeNO50", "---"),
                            "{{Temperatura}}": st.session_state.datos_pdf.get("Temperatura", "---"),
                            "{{Presion}}": st.session_state.datos_pdf.get("Presion", "---"),
                            "{{Tasa de flujo}}": st.session_state.datos_pdf.get("Tasa de flujo", "---"),
                            
                            # Imagen de la curva
                            "img_curva": st.session_state.datos_pdf.get("img_curva")
                        }
                        
                        # Ruta a la plantilla Word (ajustar según tu estructura)
                        plantilla_path = "plantillas/FeNO 50 Informe.docx"
                        
                        # Verificar que la plantilla existe
                        if not os.path.exists(plantilla_path):
                            # Intentar ubicación alternativa
                            plantilla_path = "FeNO 50 Informe.docx"
                            
                        if not os.path.exists(plantilla_path):
                            st.error(f"❌ No se encuentra la plantilla Word: {plantilla_path}")
                            st.info("""
                            **Solución:**
                            1. Asegúrese de que el archivo 'FeNO 50 Informe.docx' esté en la misma carpeta
                            2. O cree una carpeta 'plantillas/' y colóquelo allí
                            """)
                            return
                        
                        # Procesar el documento
                        doc_bytes = reemplazar_placeholders_en_documento(plantilla_path, datos_completos)
                        
                        if doc_bytes:
                            # Crear nombre de archivo
                            nombre_archivo = f"INFORME_FENO_{nombre}_{apellidos}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                            
                            # Mostrar éxito y botón de descarga
                            st.markdown('<div class="success-box">', unsafe_allow_html=True)
                            st.success("🎉 ¡Informe generado exitosamente!")
                            
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
                            3. **Guarde como PDF** (Archivo → Guardar como → PDF)
                            4. **Envíe** el PDF al médico y archive
                            """)
                            
                            st.markdown('</div>', unsafe_allow_html=True)
                        
                    except Exception as e:
                        st.error(f"❌ Error generando el informe: {str(e)}")
                        st.info("""
                        **Posibles soluciones:**
                        1. Verifique que la plantilla Word esté en la ubicación correcta
                        2. Asegúrese de que el Word no esté abierto en otro programa
                        3. Revise que tenga permisos de escritura
                        """)

# ==========================================================
# 5. EJECUCIÓN PRINCIPAL
# ==========================================================
if __name__ == "__main__":
    # Inicializar session state si no existe
    if 'datos_pdf' not in st.session_state:
        st.session_state.datos_pdf = None
    
    # Ejecutar aplicación
    main()
    
    # Pie de página
    st.markdown("---")
    st.caption("© Instituto Nacional del Tórax - Laboratorio de Función Pulmonar | Sistema de Informes FeNO v1.0")

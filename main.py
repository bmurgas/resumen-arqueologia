import streamlit as st
from docx import Document
from docx.shared import Inches
import io

# Configuración de la interfaz
st.set_page_config(page_title="Generador de Resumen Arqueológico", layout="wide")

def extraer_datos_con_imagenes(archivo_binario, nombre_archivo):
    """Extrae texto e imágenes de los anexos diarios."""
    try:
        doc = Document(io.BytesIO(archivo_binario))
    except Exception as e:
        st.error(f"No se pudo abrir {nombre_archivo}. Asegúrate que sea un .docx válido.")
        return []

    fichas_extraidas = []
    
    # Cada anexo puede tener una o más tablas (fichas)
    for tabla in doc.tables:
        ficha = {"fecha": None, "actividad": "", "fotos_binarias": []}
        seccion_fotos = False
        
        for fila in tabla.rows:
            if len(fila.cells) < 2: continue
            
            # Limpieza de texto de la columna izquierda
            encabezado = fila.cells[0].text.strip()
            
            # Captura de datos básicos
            if "Fecha" in encabezado:
                ficha["fecha"] = fila.cells[1].text.strip()
            
            if "Descripción de la actividad" in encabezado:
                ficha["actividad"] = fila.cells[1].text.strip()
            
            # Identificación de la sección de fotos
            if "Registro fotográfico" in encabezado:
                seccion_fotos = True
                continue
            
            # Extracción de imágenes reales
            if seccion_fotos and ficha["fecha"]:
                for celda in fila.cells:
                    for parrafo in celda.paragraphs:
                        for run in parrafo.runs:
                            # Buscamos el rId de la imagen en el XML del documento
                            blips = run._element.xpath('.//a:blip')
                            for blip in blips:
                                try:
                                    rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                    image_part = doc.part.related_parts[rId]
                                    if 'image' in image_part.content_type:
                                        ficha["fotos_binarias"].append(image_part.blob)
                                except:
                                    continue

        # Si la tabla tenía una fecha, es una ficha válida
        if ficha["fecha"]:
            fichas_extraidas.append(ficha)
            
    return fichas_extraidas

def generar_documento_resumen(lista_total):
    """Construye el Word final con la tabla de 3 columnas."""
    nuevo_doc = Document()
    nuevo_doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    
    # Creamos la tabla con el formato del usuario (3 columnas)
    tabla = nuevo_doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    
    # Encabezados
    hdr = tabla.rows[0].cells
    hdr[0].text = 'Fecha'
    hdr[1].text = 'Actividades realizadas durante el MAP'
    hdr[2].text = 'Imagen de la actividad'

    # Llenamos la tabla fila por fila (sin agrupar, respetando la regla 2.3)
    for ficha in lista_total:
        fila_celdas = tabla.add_row().cells
        fila_celdas[0].text = ficha["fecha"]
        fila_celdas[1].text = ficha["actividad"]
        
        # Insertar fotos en la tercera columna
        parrafo_foto = fila_celdas[2].paragraphs[0]
        if not ficha["fotos_binarias"]:
            parrafo_foto.add_run("[Sin registro fotográfico]")
            
        for img_blob in ficha["fotos_binarias"]:
            try:
                run = parrafo_foto.add_run()
                run.add_picture(io.BytesIO(img_blob), width=Inches(2.0))
                parrafo_foto.add_run("\n") # Salto de línea entre fotos
            except:
                parrafo_foto.add_run("\n[Error en formato de imagen]\n")

    # Guardar en memoria
    buffer = io.BytesIO()
    nuevo_doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ STREAMLIT ---
st.title("📂 Generador de Tabla Resumen MAP")
st.write("Sube los anexos diarios (.docx) y genera el acumulado mensual automáticamente.")

archivos_subidos = st.file_uploader("Subir archivos", type="docx", accept_multiple_files=True)

if archivos_subidos:
    if st.button("🛠️ Generar Word Final"):
        datos_completos = []
        
        # Procesamos cada archivo subido
        for archivo in archivos_subidos:
            fichas = extraer_datos_con_imagenes(archivo.read(), archivo.name)
            datos_completos.extend(fichas)
        
        if datos_completos:
            # Creamos el archivo final
            archivo_word = generar_documento_resumen(datos_completos)
            
            st.success(f"Se procesaron {len(datos_completos)} fichas.")
            st.download_button(
                label="⬇️ Descargar Tabla Resumen.docx",
                data=archivo_word,
                file_name="Resumen_Monitoreo_Acumulado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("No se encontraron datos válidos en los archivos subidos.")

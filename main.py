import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="Generador de Resumen Arqueológico", layout="wide")

def obtener_imagenes_de_celda(celda, doc_source):
    """
    Busca imágenes (inline o ancladas/flotantes) dentro del XML de una celda.
    """
    imagenes_encontradas = []
    
    # 1. Buscar imágenes INLINE (las normales)
    blips = celda._element.xpath('.//a:blip')
    
    # 2. Buscar imágenes ANCHORED (flotantes)
    # A veces Word guarda las fotos dentro de estructuras 'graphicData' anidadas
    if not blips:
        blips = celda._element.xpath('.//pic:blipFill/a:blip')
    
    for blip in blips:
        try:
            embed_attr = blip.get(qn('r:embed'))
            if embed_attr:
                image_part = doc_source.part.related_parts[embed_attr]
                # Verificar que sea realmente una imagen
                if 'image' in image_part.content_type:
                    imagenes_encontradas.append(image_part.blob)
        except Exception as e:
            continue
            
    return imagenes_encontradas

def procesar_documento(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except:
        st.error(f"Error leyendo {nombre_archivo}")
        return []

    fichas = []
    
    # Recorrer todas las tablas del anexo
    for i, tabla in enumerate(doc.tables):
        datos_ficha = {"fecha": None, "actividad": "", "fotos": []}
        seccion_fotos_activa = False
        
        for fila in tabla.rows:
            # Protección contra filas vacías
            if not fila.cells: continue
            
            # Texto de la primera columna para identificar secciones
            try:
                texto_col1 = fila.cells[0].text.strip()
            except:
                texto_col1 = ""

            # 1. FECHA
            if "Fecha" in texto_col1 and len(fila.cells) > 1:
                datos_ficha["fecha"] = fila.cells[1].text.strip()

            # 2. ACTIVIDAD
            if "Descripción de la actividad" in texto_col1 and len(fila.cells) > 1:
                datos_ficha["actividad"] = fila.cells[1].text.strip()
            
            # 3. DETECTAR SECCIÓN FOTOS
            # A veces dice "Registro fotográfico" o "Registro fotográfico actividad realizada"
            if "Registro fotográfico" in texto_col1:
                seccion_fotos_activa = True
                continue # Saltamos la fila del título
            
            # 4. EXTRAER FOTOS
            if seccion_fotos_activa and datos_ficha["fecha"]:
                # Buscamos fotos en TODAS las celdas de esta fila
                for celda in fila.cells:
                    imgs = obtener_imagenes_de_celda(celda, doc)
                    if imgs:
                        datos_ficha["fotos"].extend(imgs)

        # Si encontramos fecha, guardamos la ficha
        if datos_ficha["fecha"]:
            fichas.append(datos_ficha)
            
    return fichas

def generar_word_salida(datos_totales):
    doc_final = Document()
    doc_final.add_heading('Tabla Resumen Monitoreo', 0)
    
    # Crear tabla idéntica a tu ejemplo: 3 columnas
    tabla = doc_final.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False 
    
    # Encabezados
    headers = tabla.rows[0].cells
    headers[0].text = "Fecha"
    headers[1].text = "Actividades realizadas durante el MAP"
    headers[2].text = "Imagen de la actividad"
    
    # Ajustar anchos (Aproximación para que se parezca al tuyo)
    for cell in tabla.columns[0].cells: cell.width = Inches(0.8)
    for cell in tabla.columns[1].cells: cell.width = Inches(3.5)
    for cell in tabla.columns[2].cells: cell.width = Inches(2.5)

    for item in datos_totales:
        row = tabla.add_row().cells
        
        # Columna 1: Fecha
        row[0].text = item.get("fecha", "")
        
        # Columna 2: Actividad
        # Limpiamos saltos de línea extra
        actividad_limpia = item.get("actividad", "").replace("\n\n", "\n")
        row[1].text = actividad_limpia
        
        # Columna 3: Imágenes
        celda_img = row[2]
        parrafo = celda_img.paragraphs[0]
        
        if not item["fotos"]:
            parrafo.add_run("[Sin imágenes]")
        else:
            for img_blob in item["fotos"]:
                try:
                    run = parrafo.add_run()
                    run.add_picture(io.BytesIO(img_blob), width=Inches(2.2))
                    parrafo.add_run("\n") # Salto de línea entre foto y foto
                except:
                    pass
                    
    buffer = io.BytesIO()
    doc_final.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ STREAMLIT ---
st.title("Generador de Tabla Resumen MAP (V3 - Extracción Profunda)")

st.warning("⚠️ Asegúrate de que tu requirements.txt tenga: python-docx, streamlit, lxml")

debug_mode = st.checkbox("Ver detalles del proceso (Diagnóstico)")
archivos = st.file_uploader("Sube los anexos (.docx)", accept_multiple_files=True)

if archivos and st.button("Generar Tabla"):
    todos_los_datos = []
    
    barra = st.progress(0)
    for i, archivo in enumerate(archivos):
        # Procesar
        datos = procesar_documento(archivo.read(), archivo.name)
        todos_los_datos.extend(datos)
        
        # Diagnóstico en pantalla
        if debug_mode:
            st.write(f"📄 **{archivo.name}**: Se encontraron {len(datos)} fichas.")
            for d in datos:
                st.write(f"- Fecha: {d['fecha']} | Fotos encontradas: {len(d['fotos'])}")
        
        barra.progress((i + 1) / len(archivos))
        
    if todos_los_datos:
        # Ordenar cronológicamente
        todos_los_datos.sort(key=lambda x: x['fecha'] if x['fecha'] else "")
        
        doc_binario = generar_word_salida(todos_los_datos)
        
        st.success(f"¡Listo! Se generó una tabla con {len(todos_los_datos)} filas.")
        st.download_button(
            "⬇️ Descargar Tabla Resumen.docx",
            data=doc_binario,
            file_name="Tabla_Resumen_Final.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("No se pudo extraer información. Revisa el formato de los anexos.")

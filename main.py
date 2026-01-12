import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
import io

# Configuración de la página
st.set_page_config(page_title="Generador de Resumen Arqueológico", layout="wide")

def obtener_imagenes_seguro(celda, doc_relacionado):
    """
    Busca imágenes dentro de una celda sin causar errores de namespace.
    """
    imagenes = []
    
    # Intentamos acceder al elemento XML crudo
    try:
        xml_element = celda._element
    exceptAttributeError:
        return []

    # 1. Búsqueda estándar (imágenes inline)
    # python-docx ya sabe qué es 'a:blip', no necesitamos pasar namespaces
    blips = xml_element.xpath('.//a:blip')
    
    # 2. Búsqueda de respaldo (imágenes ancladas/flotantes)
    if not blips:
        blips = xml_element.xpath('.//pic:blipFill/a:blip')

    for blip in blips:
        try:
            # Obtenemos el ID de la relación (r:embed)
            # Usamos qn de python-docx para obtener el nombre cualificado correcto
            embed_code = blip.get(qn('r:embed'))
            
            if embed_code:
                part = doc_relacionado.part.related_parts[embed_code]
                # Verificamos que sea imagen
                if 'image' in part.content_type:
                    imagenes.append(part.blob)
        except Exception as e:
            # Si una imagen falla, la saltamos pero seguimos con las demás
            continue
            
    return imagenes

def escanear_documento_v5(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"No se pudo leer el archivo {nombre_archivo}: {e}")
        return [], []

    fichas = []
    logs = []
    
    # Recorremos todas las tablas
    for i, tabla in enumerate(doc.tables):
        ficha = {"fecha": None, "actividad": "", "fotos": []}
        buscando_fotos = False
        
        for fila in tabla.rows:
            # Convertimos toda la fila a una sola cadena de texto para buscar palabras clave
            # Esto ayuda si el formato cambia ligeramente entre fichas
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # --- 1. DETECTAR FECHA ---
            if "Fecha" in texto_fila and not ficha["fecha"]:
                # Buscamos la celda que tiene la palabra "Fecha"
                for j, celda in enumerate(fila.cells):
                    if "Fecha" in celda.text:
                        # Asumimos que la fecha está en la celda siguiente (j+1)
                        if j + 1 < len(fila.cells):
                            valor = fila.cells[j+1].text.strip()
                            if valor: # Solo si no está vacío
                                ficha["fecha"] = valor
                                logs.append(f"✅ Fecha detectada: {valor}")
                        break
            
            # --- 2. DETECTAR ACTIVIDAD ---
            # Buscamos variantes comunes del título
            if "Descripción de la actividad" in texto_fila or "Actividad" in texto_fila:
                for j, celda in enumerate(fila.cells):
                    # Si encontramos el título en esta celda...
                    if "Descripción" in celda.text or "Actividad" in celda.text:
                        # ... el contenido debería estar a la derecha
                        if j + 1 < len(fila.cells):
                            texto = fila.cells[j+1].text.strip()
                            # Guardamos si es un texto relevante (más largo que el título)
                            if len(texto) > 1:
                                ficha["actividad"] = texto
                        break

            # --- 3. DETECTAR SECCIÓN FOTOS ---
            if "Registro fotográfico" in texto_fila or "Imagen de la actividad" in texto_fila:
                buscando_fotos = True
                continue # Saltamos la línea del título

            # --- 4. EXTRAER FOTOS ---
            if buscando_fotos:
                # Si estamos en la zona de fotos, escaneamos TODAS las celdas de la fila
                for celda in fila.cells:
                    imgs = obtener_imagenes_seguro(celda, doc)
                    if imgs:
                        ficha["fotos"].extend(imgs)

        # Si al terminar de leer la tabla (ficha) tenemos fecha, la guardamos
        if ficha["fecha"]:
            fichas.append(ficha)

    return fichas, logs

def generar_word_final(datos):
    doc = Document()
    doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    
    # Configurar encabezados
    hdr = tabla.rows[0].cells
    hdr[0].text = "Fecha"
    hdr[1].text = "Actividades realizadas durante el MAP"
    hdr[2].text = "Imagen de la actividad"
    
    # Anchos
    for c in tabla.columns[0].cells: c.width = Inches(0.9) # Fecha
    for c in tabla.columns[1].cells: c.width = Inches(3.5) # Actividad
    for c in tabla.columns[2].cells: c.width = Inches(2.5) # Fotos
    
    for item in datos:
        row = tabla.add_row().cells
        row[0].text = str(item["fecha"])
        row[1].text = str(item["actividad"])
        
        # Columna de Fotos
        parrafo_img = row[2].paragraphs[0]
        if not item["fotos"]:
            parrafo_img.add_run("[Sin fotos detectadas]")
        else:
            for img_bytes in item["fotos"]:
                try:
                    run = parrafo_img.add_run()
                    run.add_picture(io.BytesIO(img_bytes), width=Inches(2.1))
                    parrafo_img.add_run("\n")
                except:
                    pass
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ DE USUARIO ---
st.title("Generador MAP V5 (Corrección de Imágenes)")
st.info("Sube los anexos. Esta versión corrige el error 'TypeError' y busca mejor el texto.")

# Checkbox para ver qué está pasando internamente
ver_logs = st.checkbox("Mostrar detalles de extracción (Útil si algo sale vacío)")

archivos = st.file_uploader("Sube Anexos Word (.docx)", accept_multiple_files=True)

if archivos and st.button("Generar Tabla Resumen"):
    fichas_totales = []
    
    barra = st.progress(0)
    
    for i, archivo in enumerate(archivos):
        bytes_archivo = archivo.read()
        fichas, logs = escanear_documento_v5(bytes_archivo, archivo.name)
        fichas_totales.extend(fichas)
        
        # Mostrar logs si el usuario quiere
        if ver_logs:
            st.text(f"--- Procesando: {archivo.name} ---")
            for l in logs:
                st.text(l)
            if not fichas:
                st.warning(f"⚠️ No se detectaron fichas en {archivo.name}")
        
        barra.progress((i + 1) / len(archivos))
        
    if fichas_totales:
        # Ordenar por fecha (simple string sort)
        fichas_totales.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        
        doc_final = generar_word_final(fichas_totales)
        
        st.success(f"¡Proceso completado! Se generaron {len(fichas_totales)} filas.")
        st.download_button(
            label="⬇️ Descargar Tabla Resumen.docx",
            data=doc_final,
            file_name="Resumen_Acumulado_MAP.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("No se pudo extraer ninguna ficha válida. Revisa los mensajes de detalle arriba.")

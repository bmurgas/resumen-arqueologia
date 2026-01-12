import streamlit as st
from docx import Document
from docx.shared import Inches
import io
from lxml import etree

# Configuración de página
st.set_page_config(page_title="Generador de Resumen Arqueológico", layout="wide")

# Mapas de nombres (Namespaces) para encontrar imágenes escondidas en el XML de Word
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

def obtener_imagenes_profundo(celda, doc_relacionado):
    """Busca imágenes usando XPATH y namespaces explícitos."""
    imagenes = []
    # Buscamos la etiqueta 'blip' (imagen) en cualquier profundidad dentro de la celda
    # Esto encuentra imágenes inline, ancladas, en tablas anidadas, etc.
    blips = celda._element.xpath('.//a:blip', namespaces=NAMESPACES)
    
    for blip in blips:
        try:
            # Obtenemos el ID de la relación (r:embed)
            embed_code = blip.get(f"{{{NAMESPACES['r']}}}embed")
            if embed_code:
                part = doc_relacionado.part.related_parts[embed_code]
                if 'image' in part.content_type:
                    imagenes.append(part.blob)
        except Exception as e:
            continue
    return imagenes

def escanear_documento(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except:
        st.error(f"Error crítico leyendo {nombre_archivo}")
        return [], []

    fichas_encontradas = []
    logs = [] # Para mostrar en pantalla qué está pasando
    
    # Recorremos todas las tablas
    for i, tabla in enumerate(doc.tables):
        ficha_actual = {"fecha": None, "actividad": "", "fotos": []}
        buscando_fotos = False
        
        for fila in tabla.rows:
            # Convertimos toda la fila a texto para buscar palabras clave sin importar la columna
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # 1. DETECTAR FECHA
            if "Fecha" in texto_fila and ficha_actual["fecha"] is None:
                # Intentamos buscar el valor de la fecha.
                # Estrategia: Buscar la celda que tenga "Fecha" y tomar la siguiente.
                for j, celda in enumerate(fila.cells):
                    if "Fecha" in celda.text:
                        # Intentamos tomar la celda de la derecha
                        if j + 1 < len(fila.cells):
                            ficha_actual["fecha"] = fila.cells[j+1].text.strip()
                            logs.append(f"✅ Fecha encontrada en tabla {i}: {ficha_actual['fecha']}")
                        break
            
            # 2. DETECTAR ACTIVIDAD (Buscamos coincidencias parciales)
            if "Descripción de la actividad" in texto_fila or "Actividad" in texto_fila:
                # Evitamos confundir el encabezado con el contenido si están en la misma fila
                for j, celda in enumerate(fila.cells):
                    if "Descripción" in celda.text or "Actividad" in celda.text:
                         if j + 1 < len(fila.cells):
                             texto_act = fila.cells[j+1].text.strip()
                             if texto_act and texto_act != ficha_actual["actividad"]:
                                 ficha_actual["actividad"] = texto_act
                                 logs.append(f"📝 Actividad detectada: {texto_act[:30]}...")
                         break

            # 3. DETECTAR SECCIÓN FOTOS
            if "Registro fotográfico" in texto_fila or "Imagen de la actividad" in texto_fila:
                buscando_fotos = True
                continue # Saltamos la línea del título

            # 4. EXTRAER FOTOS
            if buscando_fotos:
                # Buscamos en CADA celda de la fila actual
                fotos_fila = []
                for celda in fila.cells:
                    imgs = obtener_imagenes_profundo(celda, doc)
                    fotos_fila.extend(imgs)
                
                if fotos_fila:
                    ficha_actual["fotos"].extend(fotos_fila)
                    logs.append(f"📷 {len(fotos_fila)} foto(s) extraída(s).")
        
        # Guardamos la ficha si tiene al menos fecha
        if ficha_actual["fecha"]:
            fichas_encontradas.append(ficha_actual)

    return fichas_encontradas, logs

def generar_word_final(datos):
    doc = Document()
    doc.add_heading('Tabla Resumen Monitoreo', 0)
    
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    
    hdr = tabla.rows[0].cells
    hdr[0].text = "Fecha"
    hdr[1].text = "Actividades realizadas durante el MAP"
    hdr[2].text = "Imagen de la actividad"
    
    # Anchos fijos para que se vea bien
    for c in tabla.columns[0].cells: c.width = Inches(0.9)
    for c in tabla.columns[1].cells: c.width = Inches(3.5)
    for c in tabla.columns[2].cells: c.width = Inches(2.5)
    
    for item in datos:
        row = tabla.add_row().cells
        row[0].text = str(item["fecha"])
        row[1].text = str(item["actividad"])
        
        parrafo_img = row[2].paragraphs[0]
        if not item["fotos"]:
            parrafo_img.add_run("[Sin fotos]")
        
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

# --- INTERFAZ ---
st.title("Generador MAP V4 (Escáner Profundo)")
st.info("Esta versión busca en toda la fila y usa namespaces XML para las fotos.")

archivos = st.file_uploader("Sube Anexos", accept_multiple_files=True)
debug = st.checkbox("Mostrar registro de proceso (Logs)")

if archivos and st.button("Generar"):
    total_fichas = []
    
    for arch in archivos:
        bytes_archivo = arch.read()
        fichas, logs = escanear_documento(bytes_archivo, arch.name)
        total_fichas.extend(fichas)
        
        if debug:
            st.write(f"**Procesando {arch.name}:**")
            for l in logs:
                st.text(l)
                
    if total_fichas:
        # Ordenar cronológicamente
        total_fichas.sort(key=lambda x: x['fecha'] if x['fecha'] else "")
        
        word_bytes = generar_word_final(total_fichas)
        st.success(f"¡Éxito! Se generaron {len(total_fichas)} filas.")
        st.download_button("Descargar Tabla Resumen.docx", word_bytes, "Resumen_MAP.docx")
    else:
        st.error("No se encontró ninguna ficha válida (con Fecha). Revisa los logs.")

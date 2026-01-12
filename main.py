import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="Generador MAP V6 (Memoria Persistente)", layout="wide")

def obtener_imagenes_xml(elemento_padre, doc_relacionado):
    """
    Busca imágenes recursivamente en el XML del elemento (fila o celda).
    Encuentra imágenes inline, ancladas, en grupos, etc.
    """
    imagenes = []
    # Buscamos cualquier etiqueta 'blip' en cualquier profundidad
    # 'a:blip' es la etiqueta estándar para imágenes en Word OpenXML
    blips = elemento_padre.xpath('.//a:blip')
    
    for blip in blips:
        try:
            # Extraer el ID de la relación (r:embed)
            embed_code = blip.get(qn('r:embed'))
            if embed_code:
                part = doc_relacionado.part.related_parts[embed_code]
                if 'image' in part.content_type:
                    imagenes.append(part.blob)
        except:
            continue
    return imagenes

def procesar_archivo_inteligente(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"Error leyendo {nombre_archivo}: {e}")
        return []

    fichas_extraidas = []
    
    # MEMORIA PERSISTENTE
    # Estas variables recuerdan el estado aunque cambiemos de tabla
    estado_actual = {
        "fecha": None,
        "actividad": "",
        "fotos": [],
        "origen": nombre_archivo
    }
    
    # Bandera para saber si estamos en zona de fotos
    capturando_fotos = False

    # Recorremos TODAS las tablas del documento en orden
    for tabla in doc.tables:
        for fila in tabla.rows:
            # Texto completo de la fila para búsqueda flexible
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # --- 1. DETECCIÓN DE FECHA ---
            # Si encontramos una NUEVA fecha, guardamos la ficha anterior y empezamos una nueva
            if "Fecha" in texto_fila:
                # Buscar el valor de la fecha en la fila (buscamos algo que parezca fecha o esté en celda derecha)
                fecha_encontrada = None
                for celda in fila.cells:
                    txt = celda.text.strip()
                    # Si la celda no es la palabra clave "Fecha" y tiene longitud razonable, es el valor
                    if "Fecha" not in txt and len(txt) > 5: 
                        fecha_encontrada = txt
                        break
                
                if fecha_encontrada:
                    # Si ya teníamos una ficha abierta con fecha, la cerramos y guardamos
                    if estado_actual["fecha"]:
                        fichas_extraidas.append(estado_actual)
                        # Reiniciamos estado
                        estado_actual = {
                            "fecha": None, "actividad": "", "fotos": [], "origen": nombre_archivo
                        }
                    
                    # Iniciamos la nueva fecha
                    estado_actual["fecha"] = fecha_encontrada
                    capturando_fotos = False # Reseteamos fotos al cambiar fecha

            # --- 2. DETECCIÓN DE ACTIVIDAD ---
            # Solo buscamos actividad si ya tenemos fecha activa
            if estado_actual["fecha"]:
                if "Descripción de la actividad" in texto_fila or "Actividad" in texto_fila:
                    # Buscamos la celda con el texto más largo en esta fila
                    texto_mas_largo = ""
                    for celda in fila.cells:
                        txt = celda.text.strip()
                        # Ignoramos los títulos
                        if "Descripción" in txt or "Actividad" in txt: continue
                        if len(txt) > len(texto_mas_largo):
                            texto_mas_largo = txt
                    
                    if len(texto_mas_largo) > 2:
                        estado_actual["actividad"] = texto_mas_largo

            # --- 3. DETECCIÓN DE FOTOS ---
            if "Registro fotográfico" in texto_fila:
                capturando_fotos = True
                continue # No buscamos fotos en la fila del título

            if capturando_fotos and estado_actual["fecha"]:
                # Buscamos fotos en el XML de toda la fila
                try:
                    imgs = obtener_imagenes_xml(fila._element, doc)
                    if imgs:
                        estado_actual["fotos"].extend(imgs)
                except:
                    pass

    # Al terminar el documento, guardamos la última ficha si quedó pendiente
    if estado_actual["fecha"]:
        fichas_extraidas.append(estado_actual)

    return fichas_extraidas

def generar_word_v6(datos):
    doc = Document()
    doc.add_heading('Tabla Resumen Monitoreo', 0)
    
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    
    headers = tabla.rows[0].cells
    headers[0].text = "Fecha"
    headers[1].text = "Actividades realizadas durante el MAP"
    headers[2].text = "Imagen de la actividad"
    
    # Ajuste de anchos
    for c in tabla.columns[0].cells: c.width = Inches(0.9)
    for c in tabla.columns[1].cells: c.width = Inches(3.5)
    for c in tabla.columns[2].cells: c.width = Inches(2.5)

    for item in datos:
        row = tabla.add_row().cells
        row[0].text = str(item["fecha"])
        row[1].text = str(item["actividad"])
        
        celda_img = row[2]
        p = celda_img.paragraphs[0]
        
        if not item["fotos"]:
            p.add_run("[Sin fotos]")
        else:
            for img_blob in item["fotos"]:
                try:
                    run = p.add_run()
                    run.add_picture(io.BytesIO(img_blob), width=Inches(2.0))
                    p.add_run("\n")
                except:
                    pass
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ ---
st.title("Generador MAP V6 (Memoria Persistente)")
st.info("Esta versión soluciona el problema cuando las fotos o textos están en tablas divididas.")

archivos = st.file_uploader("Sube Anexos", accept_multiple_files=True)
debug = st.checkbox("Ver Logs de Extracción")

if archivos and st.button("Generar Tabla"):
    todas = []
    for arch in archivos:
        fichas = procesar_archivo_inteligente(arch.read(), arch.name)
        todas.extend(fichas)
        
        if debug:
            st.write(f"📂 {arch.name}: {len(fichas)} fichas detectadas.")
            for f in fichas:
                st.write(f"   - Fecha: {f['fecha']} | Actividad: {len(f['actividad'])} caracteres | Fotos: {len(f['fotos'])}")

    if todas:
        # Ordenar cronológicamente
        todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        
        word_out = generar_word_v6(todas)
        st.success(f"¡Éxito! Tabla generada con {len(todas)} filas.")
        st.download_button("Descargar Tabla Resumen.docx", word_out, "Tabla_Resumen.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.error("No se encontraron datos. Verifica si los archivos tienen la palabra 'Fecha'.")

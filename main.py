import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="Generador MAP V9 (Final)", layout="wide")

def obtener_imagenes_de_elemento(elemento_xml, doc_relacionado):
    """
    Busca imágenes recursivamente dentro de cualquier elemento XML (celda o párrafo).
    """
    imagenes = []
    # Buscamos etiquetas 'a:blip' que son las imágenes en Word
    blips = elemento_xml.xpath('.//a:blip')
    for blip in blips:
        try:
            embed_code = blip.get(qn('r:embed'))
            if embed_code:
                part = doc_relacionado.part.related_parts[embed_code]
                if 'image' in part.content_type:
                    imagenes.append(part.blob)
        except:
            continue
    return imagenes

def obtener_texto_celda_abajo(tabla, fila_idx, col_idx):
    """Intenta obtener el texto de la celda justo debajo (fila+1, misma columna)."""
    try:
        # Verificamos si existe la fila siguiente
        if fila_idx + 1 < len(tabla.rows):
            fila_siguiente = tabla.rows[fila_idx + 1]
            # Verificamos si existe la columna (cuidado con celdas combinadas)
            if col_idx < len(fila_siguiente.cells):
                return fila_siguiente.cells[col_idx].text.strip()
    except:
        pass
    return ""

def procesar_archivo_v9(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"Error leyendo {nombre_archivo}: {e}")
        return []

    fichas_extraidas = []
    
    # MEMORIA PERSISTENTE (Para heredar fecha si la ficha no la tiene)
    fecha_persistente = "Sin Fecha"

    # Recorremos todas las tablas (fichas)
    for tabla in doc.tables:
        
        datos_ficha = {
            "fecha_propia": None,
            "actividad": "",
            "hallazgos": "", 
            "items_foto": [] # Lista de diccionarios: {'blob': img, 'leyenda': txt}
        }
        
        en_seccion_fotos = False
        
        # Analizamos fila por fila
        for r_idx, fila in enumerate(tabla.rows):
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # --- 1. FECHA (Heredable) ---
            if "Fecha" in texto_fila:
                for celda in fila.cells:
                    t = celda.text.strip()
                    # Si parece una fecha (no es el título 'Fecha' y es larga)
                    if "Fecha" not in t and len(t) > 5:
                        datos_ficha["fecha_propia"] = t
                        fecha_persistente = t # Actualizamos la memoria
                        break
            
            # --- 2. ACTIVIDAD ---
            if "Descripción de la actividad" in texto_fila:
                mejor_texto = ""
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Descripción" in t or "Actividad" in t: continue
                    if len(t) > len(mejor_texto):
                        mejor_texto = t
                if mejor_texto:
                    datos_ficha["actividad"] = mejor_texto

            # --- 3. HALLAZGOS (Lógica estricta de la "X") ---
            if "Ausencia" in texto_fila:
                # Revisamos si alguna celda de esta fila tiene una "X"
                tiene_x = any(c.text.strip().upper() == "X" for c in fila.cells)
                if tiene_x:
                    datos_ficha["hallazgos"] = "Ausencia de hallazgos arqueológicos no previstos."
            
            if "Presencia" in texto_fila:
                tiene_x = any(c.text.strip().upper() == "X" for c in fila.cells)
                if tiene_x:
                    datos_ficha["hallazgos"] = "PRESENCIA de hallazgos arqueológicos."

            # --- 4. FOTOS Y LEYENDAS ---
            # Detectamos inicio de sección
            if "Registro fotográfico" in texto_fila:
                en_seccion_fotos = True
                continue # Saltamos la fila del título

            if en_seccion_fotos:
                # Recorremos CADA CELDA de la fila buscando fotos
                for c_idx, celda in enumerate(fila.cells):
                    # Usamos la función de extracción robusta (V7) pero por celda
                    imgs = obtener_imagenes_de_elemento(celda._element, doc)
                    
                    if imgs:
                        # ¡Encontramos fotos! Ahora buscamos su leyenda.
                        
                        # Opción A: Texto en la misma celda
                        texto_leyenda = celda.text.strip()
                        
                        # Opción B: Si la celda está vacía, miramos la celda DE ABAJO
                        if not texto_leyenda:
                            texto_leyenda = obtener_texto_celda_abajo(tabla, r_idx, c_idx)
                        
                        # Guardamos cada foto encontrada con esa leyenda
                        for img in imgs:
                            datos_ficha["items_foto"].append({
                                "blob": img,
                                "leyenda": texto_leyenda
                            })

        # --- FIN DE LA TABLA ---
        # Determinamos fecha final (Propia o Heredada)
        fecha_final = datos_ficha["fecha_propia"] if datos_ficha["fecha_propia"] else fecha_persistente
        
        # Validamos Hallazgos (Si no se marcó nada, lo dejamos vacío o ponemos default si prefieres)
        # if not datos_ficha["hallazgos"]: datos_ficha["hallazgos"] = "Ausencia (No verificado)"

        # Solo guardamos si hay contenido relevante
        if datos_ficha["actividad"] or datos_ficha["items_foto"]:
            
            # Construimos el texto central combinado
            texto_central = datos_ficha["actividad"]
            if datos_ficha["hallazgos"]:
                texto_central += f"\n\n[Hallazgos: {datos_ficha['hallazgos']}]"
            
            fichas_extraidas.append({
                "fecha": fecha_final,
                "texto_central": texto_central,
                "fotos": datos_ficha["items_foto"]
            })

    return fichas_extraidas

def generar_word_v9(datos):
    doc = Document()
    doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    
    headers = tabla.rows[0].cells
    headers[0].text = "Fecha"
    headers[1].text = "Actividades realizadas durante el MAP"
    headers[2].text = "Imagen de la actividad"
    
    # Anchos de columna
    for c in tabla.columns[0].cells: c.width = Inches(0.9)
    for c in tabla.columns[1].cells: c.width = Inches(3.5)
    for c in tabla.columns[2].cells: c.width = Inches(2.5)

    for item in datos:
        row = tabla.add_row().cells
        row[0].text = str(item["fecha"])
        row[1].text = str(item["texto_central"])
        
        # Columna de Fotos
        celda_img = row[2]
        parrafo = celda_img.paragraphs[0]
        
        if not item["fotos"]:
            parrafo.add_run("[Sin fotos]")
        else:
            for i, foto_obj in enumerate(item["fotos"]):
                try:
                    # Insertamos imagen
                    run = parrafo.add_run()
                    run.add_picture(io.BytesIO(foto_obj["blob"]), width=Inches(2.0))
                    
                    # Insertamos leyenda
                    if foto_obj["leyenda"]:
                        run_text = parrafo.add_run(f"\n{foto_obj['leyenda']}")
                        run_text.font.size = Pt(9)
                        run_text.italic = True
                    
                    # Espacio entre fotos
                    if i < len(item["fotos"]) - 1:
                        parrafo.add_run("\n\n")
                except:
                    continue
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ ---
st.title("Generador MAP V9 (Final)")
st.info("Incluye: Hallazgos por X, Fotos robustas y Leyendas automáticas (misma celda o abajo).")

archivos = st.file_uploader("Sube Anexos (.docx)", accept_multiple_files=True)
debug = st.checkbox("Ver Logs")

if archivos and st.button("Generar Tabla"):
    todas = []
    bar = st.progress(0)
    
    for i, a in enumerate(archivos):
        fichas = procesar_archivo_v9(a.read(), a.name)
        todas.extend(fichas)
        bar.progress((i+1)/len(archivos))
        
        if debug:
            st.write(f"📄 {a.name}: {len(fichas)} fichas.")
            for f in fichas:
                st.write(f"   📅 {f['fecha']} | Fotos: {len(f['fotos'])}")

    if todas:
        todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        doc_out = generar_word_v9(todas)
        st.success(f"¡Listo! Se generó la tabla con {len(todas)} registros.")
        st.download_button("Descargar Tabla Resumen.docx", doc_out, "Resumen_MAP_Final.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.error("No se encontraron datos.")

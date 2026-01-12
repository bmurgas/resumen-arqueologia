import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="Generador MAP V10 (Anti-Duplicados)", layout="wide")

def obtener_imagenes_con_id(elemento_xml, doc_relacionado):
    """
    Busca imágenes y retorna su ID único (rId) y el contenido binario.
    Esto permite filtrar duplicados después.
    """
    resultados = [] # Lista de tuplas (rId, blob)
    blips = elemento_xml.xpath('.//a:blip')
    for blip in blips:
        try:
            embed_code = blip.get(qn('r:embed'))
            if embed_code:
                part = doc_relacionado.part.related_parts[embed_code]
                if 'image' in part.content_type:
                    resultados.append((embed_code, part.blob))
        except:
            continue
    return resultados

def obtener_texto_celda_abajo(tabla, fila_idx, col_idx):
    """Intenta obtener el texto de la celda justo debajo (fila+1, misma columna)."""
    try:
        if fila_idx + 1 < len(tabla.rows):
            fila_siguiente = tabla.rows[fila_idx + 1]
            if col_idx < len(fila_siguiente.cells):
                return fila_siguiente.cells[col_idx].text.strip()
    except:
        pass
    return ""

def procesar_archivo_v10(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"Error leyendo {nombre_archivo}: {e}")
        return []

    fichas_extraidas = []
    fecha_persistente = "Sin Fecha"

    for tabla in doc.tables:
        datos_ficha = {
            "fecha_propia": None,
            "actividad": "",
            "hallazgos": "", 
            "items_foto": [] 
        }
        
        # CONJUNTOS PARA EVITAR DUPLICADOS EN ESTA FICHA
        rids_procesados = set()   # Guarda los ID de las fotos ya añadidas
        celdas_procesadas = set() # Guarda los ID de memoria de las celdas leídas
        
        en_seccion_fotos = False
        
        for r_idx, fila in enumerate(tabla.rows):
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # --- 1. FECHA ---
            if "Fecha" in texto_fila:
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Fecha" not in t and len(t) > 5:
                        datos_ficha["fecha_propia"] = t
                        fecha_persistente = t
                        break
            
            # --- 2. ACTIVIDAD ---
            if "Descripción de la actividad" in texto_fila:
                mejor_texto = ""
                # Evitar leer la misma celda merged multiples veces
                celdas_fila_vistas = set()
                for celda in fila.cells:
                    if celda in celdas_fila_vistas: continue
                    celdas_fila_vistas.add(celda)
                    
                    t = celda.text.strip()
                    if "Descripción" in t or "Actividad" in t: continue
                    if len(t) > len(mejor_texto):
                        mejor_texto = t
                if mejor_texto:
                    datos_ficha["actividad"] = mejor_texto

            # --- 3. HALLAZGOS ---
            if "Ausencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "Ausencia de hallazgos arqueológicos no previstos."
            
            if "Presencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "PRESENCIA de hallazgos arqueológicos."

            # --- 4. FOTOS Y LEYENDAS (Con Anti-Duplicación) ---
            if "Registro fotográfico" in texto_fila:
                en_seccion_fotos = True
                continue 

            if en_seccion_fotos:
                for c_idx, celda in enumerate(fila.cells):
                    # FILTRO 1: Si esta celda ya la leímos en esta fila (caso celdas combinadas), saltar
                    if celda in celdas_procesadas:
                        continue
                    celdas_procesadas.add(celda)

                    # Obtenemos imágenes con su ID único
                    lista_imgs_ids = obtener_imagenes_con_id(celda._element, doc)
                    
                    if lista_imgs_ids:
                        # Buscamos leyenda (misma celda o abajo)
                        texto_leyenda = celda.text.strip()
                        if not texto_leyenda:
                            texto_leyenda = obtener_texto_celda_abajo(tabla, r_idx, c_idx)
                        
                        for rId, blob in lista_imgs_ids:
                            # FILTRO 2: Si este ID de imagen ya está en la ficha, saltar
                            if rId in rids_procesados:
                                continue
                            
                            # Si es nueva, la guardamos y marcamos como procesada
                            rids_procesados.add(rId)
                            datos_ficha["items_foto"].append({
                                "blob": blob,
                                "leyenda": texto_leyenda
                            })
                
                # Limpiamos el set de celdas al cambiar de fila (solo nos importa duplicados horizontales en merged cells)
                celdas_procesadas.clear() 

        # --- FIN DE FICHA ---
        fecha_final = datos_ficha["fecha_propia"] if datos_ficha["fecha_propia"] else fecha_persistente
        
        if datos_ficha["actividad"] or datos_ficha["items_foto"]:
            texto_central = datos_ficha["actividad"]
            if datos_ficha["hallazgos"]:
                texto_central += f"\n\n[Hallazgos: {datos_ficha['hallazgos']}]"
            
            fichas_extraidas.append({
                "fecha": fecha_final,
                "texto_central": texto_central,
                "fotos": datos_ficha["items_foto"]
            })

    return fichas_extraidas

def generar_word_v10(datos):
    doc = Document()
    doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    
    headers = tabla.rows[0].cells
    headers[0].text = "Fecha"
    headers[1].text = "Actividades realizadas durante el MAP"
    headers[2].text = "Imagen de la actividad"
    
    for c in tabla.columns[0].cells: c.width = Inches(0.9)
    for c in tabla.columns[1].cells: c.width = Inches(3.5)
    for c in tabla.columns[2].cells: c.width = Inches(2.5)

    for item in datos:
        row = tabla.add_row().cells
        row[0].text = str(item["fecha"])
        row[1].text = str(item["texto_central"])
        
        celda_img = row[2]
        parrafo = celda_img.paragraphs[0]
        
        if not item["fotos"]:
            parrafo.add_run("[Sin fotos]")
        else:
            for i, foto_obj in enumerate(item["fotos"]):
                try:
                    run = parrafo.add_run()
                    run.add_picture(io.BytesIO(foto_obj["blob"]), width=Inches(2.0))
                    
                    if foto_obj["leyenda"]:
                        run_text = parrafo.add_run(f"\n{foto_obj['leyenda']}")
                        run_text.font.size = Pt(9)
                        run_text.italic = True
                    
                    if i < len(item["fotos"]) - 1:
                        parrafo.add_run("\n\n")
                except:
                    continue
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ ---
st.title("Generador MAP V10 (Sin Duplicados)")
st.info("Esta versión filtra fotos repetidas causadas por celdas combinadas.")

archivos = st.file_uploader("Sube Anexos (.docx)", accept_multiple_files=True)
debug = st.checkbox("Ver Logs")

if archivos and st.button("Generar Tabla"):
    todas = []
    bar = st.progress(0)
    
    for i, a in enumerate(archivos):
        fichas = procesar_archivo_v10(a.read(), a.name)
        todas.extend(fichas)
        bar.progress((i+1)/len(archivos))
        
        if debug:
            st.write(f"📄 {a.name}: {len(fichas)} fichas.")
            for f in fichas:
                st.write(f"   📅 {f['fecha']} | Fotos únicas: {len(f['fotos'])}")

    if todas:
        todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        doc_out = generar_word_v10(todas)
        st.success("Tabla generada correctamente.")
        st.download_button("Descargar Tabla Resumen.docx", doc_out, "Resumen_MAP_V10.docx")
    else:
        st.error("No se encontraron datos.")

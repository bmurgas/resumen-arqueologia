import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="Generador MAP V8 (Leyendas + Hallazgos X)", layout="wide")

def obtener_imagenes_xml(celda, doc_relacionado):
    """
    Extrae las imágenes binarias de una celda.
    """
    imagenes = []
    try:
        # Buscamos etiquetas de imagen (a:blip)
        blips = celda._element.xpath('.//a:blip')
        for blip in blips:
            embed_code = blip.get(qn('r:embed'))
            if embed_code:
                part = doc_relacionado.part.related_parts[embed_code]
                if 'image' in part.content_type:
                    imagenes.append(part.blob)
    except:
        pass
    return imagenes

def procesar_archivo_v8(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"Error leyendo {nombre_archivo}: {e}")
        return []

    fichas_extraidas = []
    
    # MEMORIA PERSISTENTE PARA FECHA
    fecha_actual_persistente = "Sin Fecha Inicial"

    # Recorremos TABLA POR TABLA
    for tabla_idx, tabla in enumerate(doc.tables):
        
        datos_ficha = {
            "fecha_propia": None,
            "actividad": "",
            "hallazgos": "", # Se llenará según la X
            "fotos_con_leyenda": [] # Lista de diccionarios {img, leyenda}
        }
        
        # Estructuras para mapear fotos y textos por coordenadas (Fila, Columna)
        mapa_fotos = {} # Key: (r, c), Value: [blob, blob...]
        mapa_textos = {} # Key: (r, c), Value: "Texto"
        seccion_fotos_inicio = -1 # Fila donde empiezan las fotos
        
        # --- PASO 1: ESCANEO DE LA TABLA ---
        for r_idx, fila in enumerate(tabla.rows):
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # A) FECHA
            if "Fecha" in texto_fila:
                for celda in fila.cells:
                    txt = celda.text.strip()
                    if "Fecha" not in txt and len(txt) > 6:
                        datos_ficha["fecha_propia"] = txt
                        fecha_actual_persistente = txt
                        break
            
            # B) ACTIVIDAD
            if "Descripción de la actividad" in texto_fila:
                mejor_texto = ""
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Descripción" in t or "Actividad" in t: continue
                    if len(t) > len(mejor_texto):
                        mejor_texto = t
                if mejor_texto:
                    datos_ficha["actividad"] = mejor_texto

            # C) HALLAZGOS (Lógica de la "X")
            # Buscamos si la fila corresponde a Ausencia o Presencia
            if "Ausencia" in texto_fila:
                # Revisamos TODAS las celdas de esta fila buscando una "X"
                for celda in fila.cells:
                    if celda.text.strip().upper() == "X":
                        datos_ficha["hallazgos"] = "Ausencia de hallazgos arqueológicos no previstos."
            
            if "Presencia" in texto_fila:
                for celda in fila.cells:
                    if celda.text.strip().upper() == "X":
                        datos_ficha["hallazgos"] = "PRESENCIA de hallazgos arqueológicos."

            # D) REGISTRO FOTOGRÁFICO (Mapeo)
            if "Registro fotográfico" in texto_fila:
                seccion_fotos_inicio = r_idx
                continue
            
            if seccion_fotos_inicio != -1 and r_idx > seccion_fotos_inicio:
                # Estamos en la zona de fotos. Guardamos todo lo que vemos.
                for c_idx, celda in enumerate(fila.cells):
                    # 1. Guardar Texto
                    texto_celda = celda.text.strip()
                    if texto_celda:
                        mapa_textos[(r_idx, c_idx)] = texto_celda
                    
                    # 2. Guardar Fotos
                    imgs = obtener_imagenes_xml(celda._element, doc)
                    if imgs:
                        mapa_fotos[(r_idx, c_idx)] = imgs

        # --- PASO 2: ASOCIAR FOTOS CON LEYENDAS ---
        # Recorremos las fotos encontradas y buscamos su leyenda
        # Prioridad: 1. Texto en la misma celda. 2. Texto en la celda de ABRAJO.
        for (r, c), lista_imgs in mapa_fotos.items():
            leyenda_encontrada = ""
            
            # Intento 1: Misma celda
            if (r, c) in mapa_textos:
                leyenda_encontrada = mapa_textos[(r, c)]
            
            # Intento 2: Celda de abajo (r+1), misma columna
            # Muchos formatos ponen la foto arriba y el texto abajo
            if not leyenda_encontrada:
                if (r + 1, c) in mapa_textos:
                    leyenda_encontrada = mapa_textos[(r + 1, c)]
            
            # Guardamos cada foto con su leyenda
            for img_blob in lista_imgs:
                datos_ficha["fotos_con_leyenda"].append({
                    "blob": img_blob,
                    "leyenda": leyenda_encontrada
                })

        # --- PASO 3: CONSOLIDAR ---
        fecha_final = datos_ficha["fecha_propia"] if datos_ficha["fecha_propia"] else fecha_actual_persistente
        
        # Si no se detectó ninguna X, ponemos un default seguro (opcional) o lo dejamos vacío
        # Si quieres forzar Ausencia si no hay X, descomenta la siguiente línea:
        # if not datos_ficha["hallazgos"]: datos_ficha["hallazgos"] = "Ausencia (No marcado)"

        if datos_ficha["actividad"] or datos_ficha["fotos_con_leyenda"]:
            
            texto_col_central = datos_ficha["actividad"]
            if datos_ficha["hallazgos"]:
                texto_col_central += f"\n\n[Hallazgos: {datos_ficha['hallazgos']}]"
            
            fichas_extraidas.append({
                "fecha": fecha_final,
                "texto_central": texto_col_central,
                "fotos": datos_ficha["fotos_con_leyenda"]
            })

    return fichas_extraidas

def generar_word_v8(datos):
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
                    
                    # Agregar la leyenda debajo de la foto
                    if foto_obj["leyenda"]:
                        p_leyenda = parrafo.add_run(f"\n{foto_obj['leyenda']}\n")
                        p_leyenda.font.size = Pt(9) # Letra un poco más pequeña para la leyenda
                        p_leyenda.italic = True
                    else:
                        parrafo.add_run("\n")
                        
                    # Espacio extra entre fotos si hay más de una
                    if i < len(item["fotos"]) - 1:
                        parrafo.add_run("\n")
                except:
                    pass
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ ---
st.title("Generador MAP V8 (Leyendas + Hallazgos por 'X')")
st.info("Ahora busca la 'X' para definir Hallazgos y busca el texto debajo de la foto.")

archivos = st.file_uploader("Sube Anexos", accept_multiple_files=True)
debug = st.checkbox("Debug")

if archivos and st.button("Generar Tabla"):
    todas = []
    bar = st.progress(0)
    for i, a in enumerate(archivos):
        fichas = procesar_archivo_v8(a.read(), a.name)
        todas.extend(fichas)
        bar.progress((i+1)/len(archivos))
        
        if debug:
            st.write(f"**{a.name}**: {len(fichas)} fichas.")
            for f in fichas:
                st.write(f"- {f['fecha']} | Hallazgo: {f['texto_central'][-20:]} | Fotos: {len(f['fotos'])}")

    if todas:
        todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        doc_out = generar_word_v8(todas)
        st.success("Tabla generada con éxito.")
        st.download_button("Descargar Tabla Resumen.docx", doc_out, "Resumen_MAP_V8.docx")
    else:
        st.error("No se encontraron datos.")

import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm # Importamos Cm para medidas exactas
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH # Para centrar texto/fotos
from docx.enum.table import WD_TABLE_ALIGNMENT # Para centrar la tabla
import io

st.set_page_config(page_title="Generador MAP V12 (Formato Franklin)", layout="wide")

# --- BLOQUE DE EXTRACCIÓN (LÓGICA V10 INTACTA) ---
def obtener_imagenes_con_id(elemento_xml, doc_relacionado):
    resultados = [] 
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
    try:
        if fila_idx + 1 < len(tabla.rows):
            fila_siguiente = tabla.rows[fila_idx + 1]
            if col_idx < len(fila_siguiente.cells):
                return fila_siguiente.cells[col_idx].text.strip()
    except:
        pass
    return ""

def procesar_archivo_v12(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"Error leyendo {nombre_archivo}: {e}")
        return []

    fichas_extraidas = []
    fecha_persistente = "Sin Fecha"

    for tabla in doc.tables:
        datos_ficha = {
            "fecha_propia": None, "actividad": "", "hallazgos": "", "items_foto": [] 
        }
        rids_procesados = set()
        celdas_procesadas = set()
        en_seccion_fotos = False
        
        for r_idx, fila in enumerate(tabla.rows):
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # 1. FECHA
            if "Fecha" in texto_fila:
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Fecha" not in t and len(t) > 5:
                        datos_ficha["fecha_propia"] = t
                        fecha_persistente = t
                        break
            
            # 2. ACTIVIDAD
            if "Descripción de la actividad" in texto_fila:
                mejor_texto = ""
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

            # 3. HALLAZGOS
            if "Ausencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "-Durante la jornada no se registran hallazgos arqueológicos protegidos por la Ley 17.288 Sobre Monumentos Nacionales."
            if "Presencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "Se identifican hallazgos arqueológicos."

            # 4. FOTOS
            if "Registro fotográfico" in texto_fila:
                en_seccion_fotos = True
                continue 

            if en_seccion_fotos:
                for c_idx, celda in enumerate(fila.cells):
                    if celda in celdas_procesadas: continue
                    celdas_procesadas.add(celda)

                    lista_imgs_ids = obtener_imagenes_con_id(celda._element, doc)
                    if lista_imgs_ids:
                        texto_leyenda = celda.text.strip()
                        if not texto_leyenda:
                            texto_leyenda = obtener_texto_celda_abajo(tabla, r_idx, c_idx)
                        
                        for rId, blob in lista_imgs_ids:
                            if rId in rids_procesados: continue
                            rids_procesados.add(rId)
                            datos_ficha["items_foto"].append({
                                "blob": blob, "leyenda": texto_leyenda
                            })
                celdas_procesadas.clear() 

        fecha_final = datos_ficha["fecha_propia"] if datos_ficha["fecha_propia"] else fecha_persistente
        
        if datos_ficha["actividad"] or datos_ficha["items_foto"]:
            texto_central = datos_ficha["actividad"]
            if datos_ficha["hallazgos"]:
                texto_central += f"\n\n[Hallazgos: {datos_ficha['hallazgos']}]"
            
            fichas_extraidas.append({
                "fecha": fecha_final, "texto_central": texto_central, "fotos": datos_ficha["items_foto"]
            })

    return fichas_extraidas

# --- BLOQUE DE GENERACIÓN WORD (FORMATO APLICADO AQUÍ) ---
def generar_word_con_formato(datos):
    doc = Document()
    
    # Título (Opcional, mantenemos estándar)
    titulo = doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 1. ALINEAR TABLA AL CENTRO
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER 

    # --- Configuración de Encabezados ---
    headers = tabla.rows[0].cells
    titulos = ["Fecha", "Actividades realizadas durante el MAP", "Imagen de la actividad"]
    
    for i, texto in enumerate(titulos):
        parrafo = headers[i].paragraphs[0]
        run = parrafo.add_run(texto)
        # 2. FUENTE FRANKLIN GOTHIC BOOK 9
        run.font.name = 'Franklin Gothic Book'
        run.font.size = Pt(9)
        run.bold = True
        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Ajuste de anchos (Necesitamos al menos 8cm en la col 3 para la foto)
    # 8 cm es aprox 3.15 pulgadas.
    for c in tabla.columns[0].cells: c.width = Cm(2.5) # Fecha
    for c in tabla.columns[1].cells: c.width = Cm(7.5) # Actividad
    for c in tabla.columns[2].cells: c.width = Cm(8.5) # Fotos (Un poco más que 8cm)

    # --- Llenado de Filas ---
    for item in datos:
        row = tabla.add_row().cells
        
        # COLUMNA 1: FECHA
        p_fecha = row[0].paragraphs[0]
        p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fecha = p_fecha.add_run(str(item["fecha"]))
        r_fecha.font.name = 'Franklin Gothic Book'
        r_fecha.font.size = Pt(9)

        # COLUMNA 2: TEXTO
        p_act = row[1].paragraphs[0]
        p_act.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # Justificado se ve mejor
        r_act = p_act.add_run(str(item["texto_central"]))
        r_act.font.name = 'Franklin Gothic Book'
        r_act.font.size = Pt(9)
        
        # COLUMNA 3: FOTOS
        celda_img = row[2]
        p_img = celda_img.paragraphs[0]
        # 4. FOTO ALINEADA AL CENTRO DE SU COLUMNA
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER 
        
        if not item["fotos"]:
            r_sin = p_img.add_run("[Sin fotos]")
            r_sin.font.name = 'Franklin Gothic Book'
            r_sin.font.size = Pt(9)
        else:
            for i, foto_obj in enumerate(item["fotos"]):
                try:
                    run = p_img.add_run()
                    # 3. TAMAÑO DE FOTO 6 CM ALTO x 8 CM ANCHO
                    # Nota: Forzar ambas medidas puede deformar la imagen si no es 4:3.
                    # Si prefieres mantener proporción, usa solo width. Aquí obedezco tu orden.
                    run.add_picture(io.BytesIO(foto_obj["blob"]), width=Cm(8), height=Cm(6))
                    
                    if foto_obj["leyenda"]:
                        r_leyenda = p_img.add_run(f"\n{foto_obj['leyenda']}")
                        r_leyenda.font.name = 'Franklin Gothic Book'
                        r_leyenda.font.size = Pt(9)
                        r_leyenda.italic = True
                    
                    if i < len(item["fotos"]) - 1:
                        p_img.add_run("\n\n")
                except:
                    continue
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ ---
st.title("Generador MAP V12 (Estilo Franklin)")
st.info("Formato: Tabla Centrada, Fuente Franklin Gothic Book 9, Fotos 8x6 cm centradas.")

archivos = st.file_uploader("Sube Anexos (.docx)", accept_multiple_files=True)
debug = st.checkbox("Ver Logs")

if archivos and st.button("Generar Tabla"):
    todas = []
    bar = st.progress(0)
    
    for i, a in enumerate(archivos):
        fichas = procesar_archivo_v12(a.read(), a.name)
        todas.extend(fichas)
        bar.progress((i+1)/len(archivos))

    if todas:
        todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        doc_out = generar_word_con_formato(todas)
        st.success("Tabla generada con el formato solicitado.")
        st.download_button("Descargar Tabla Resumen.docx", doc_out, "Resumen_MAP_Estilo.docx")
    else:
        st.error("No se encontraron datos.")

import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

st.set_page_config(page_title="Generador MAP V13 (Extracción Exacta)", layout="wide")

# --- FUNCIONES AUXILIARES ---

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

def procesar_archivo_v13(archivo_bytes, nombre_archivo):
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
            "actividad": "",     # La actividad general
            "categoria": "",     # SA o HA
            "descripcion_item": "", # La descripción del lítico/hallazgo
            "hallazgos_check": "",  # Si hubo X en Ausencia/Presencia
            "items_foto": [] 
        }
        
        rids_procesados = set()
        celdas_procesadas = set()
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
            
            # --- 2. EXTRACCIÓN POR PARES (Titulo -> Valor al lado) ---
            # Recorremos las celdas para encontrar los títulos y sacar el valor de la derecha
            for c_idx, celda in enumerate(fila.cells):
                txt_celda = celda.text.strip()
                
                # A) CATEGORÍA
                if "Categoría" in txt_celda: # Busca "Categoría (SA/HA):"
                    # Intentamos tomar la celda siguiente (c_idx + 1)
                    if c_idx + 1 < len(fila.cells):
                        valor = fila.cells[c_idx + 1].text.strip()
                        if valor: datos_ficha["categoria"] = valor
                
                # B) DESCRIPCIÓN (Específica del item)
                # Ojo: A veces dice "Descripción de la actividad" (esa es otra).
                # Buscamos "Descripción" a secas o que NO tenga "actividad" para no confundir
                elif "Descripción" in txt_celda and "actividad" not in txt_celda:
                    if c_idx + 1 < len(fila.cells):
                        valor = fila.cells[c_idx + 1].text.strip()
                        if valor: datos_ficha["descripcion_item"] = valor

            # --- 3. ACTIVIDAD GENERAL ---
            if "Descripción de la actividad" in texto_fila:
                mejor_texto = ""
                celdas_vistas = set()
                for celda in fila.cells:
                    if celda in celdas_vistas: continue
                    celdas_vistas.add(celda)
                    t = celda.text.strip()
                    # Ignoramos el título mismo
                    if "Descripción" in t or "Actividad" in t: continue
                    if len(t) > len(mejor_texto):
                        mejor_texto = t
                if mejor_texto:
                    datos_ficha["actividad"] = mejor_texto

            # --- 4. HALLAZGOS (Check X) ---
            if "Ausencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos_check"] = "Ausencia de hallazgos arqueológicos no previstos."
            if "Presencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos_check"] = "PRESENCIA de hallazgos arqueológicos."

            # --- 5. FOTOS ---
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

        # --- CONSOLIDACIÓN DE DATOS ---
        fecha_final = datos_ficha["fecha_propia"] if datos_ficha["fecha_propia"] else fecha_persistente
        
        # Validar si guardamos la ficha
        hay_info = (datos_ficha["actividad"] or datos_ficha["items_foto"] or 
                    datos_ficha["categoria"] or datos_ficha["descripcion_item"])
        
        if hay_info:
            # Construimos el Texto Central combinando todo lo encontrado
            partes_texto = []
            
            if datos_ficha["actividad"]:
                partes_texto.append(datos_ficha["actividad"])
            
            # Agregamos Categoría y Descripción si existen
            if datos_ficha["categoria"]:
                partes_texto.append(f"Categoría: {datos_ficha['categoria']}")
                
            if datos_ficha["descripcion_item"]:
                partes_texto.append(f"Descripción: {datos_ficha['descripcion_item']}")
            
            # Agregamos el check de hallazgos al final
            if datos_ficha["hallazgos_check"]:
                partes_texto.append(f"\n[Hallazgos: {datos_ficha['hallazgos_check']}]")
            
            texto_consolidado = "\n\n".join(partes_texto)
            
            fichas_extraidas.append({
                "fecha": fecha_final, 
                "texto_central": texto_consolidado, 
                "fotos": datos_ficha["items_foto"]
            })

    return fichas_extraidas

# --- GENERACIÓN WORD (ESTILO FRANKLIN MANTENIDO) ---
def generar_word_v13(datos):
    doc = Document()
    
    # Título
    titulo = doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER 

    # Encabezados
    headers = tabla.rows[0].cells
    titulos = ["Fecha", "Actividades realizadas durante el MAP", "Imagen de la actividad"]
    
    for i, texto in enumerate(titulos):
        parrafo = headers[i].paragraphs[0]
        run = parrafo.add_run(texto)
        run.font.name = 'Franklin Gothic Book'
        run.font.size = Pt(9)
        run.bold = True
        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Anchos
    for c in tabla.columns[0].cells: c.width = Cm(2.5) 
    for c in tabla.columns[1].cells: c.width = Cm(7.5) 
    for c in tabla.columns[2].cells: c.width = Cm(8.5) 

    for item in datos:
        row = tabla.add_row().cells
        
        # 1. FECHA
        p_fecha = row[0].paragraphs[0]
        p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fecha = p_fecha.add_run(str(item["fecha"]))
        r_fecha.font.name = 'Franklin Gothic Book'
        r_fecha.font.size = Pt(9)

        # 2. TEXTO (Ahora incluye Categoría y Descripción)
        p_act = row[1].paragraphs[0]
        p_act.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_act = p_act.add_run(str(item["texto_central"]))
        r_act.font.name = 'Franklin Gothic Book'
        r_act.font.size = Pt(9)
        
        # 3. FOTOS
        celda_img = row[2]
        p_img = celda_img.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER 
        
        if not item["fotos"]:
            r_sin = p_img.add_run("[Sin fotos]")
            r_sin.font.name = 'Franklin Gothic Book'
            r_sin.font.size = Pt(9)
        else:
            for i, foto_obj in enumerate(item["fotos"]):
                try:
                    run = p_img.add_run()
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
st.title("Generador MAP V13 (Extracción de Categoría y Descripción)")
st.info("Ahora extrae correctamente los campos 'Categoría' y 'Descripción' de sus celdas adyacentes.")

archivos = st.file_uploader("Sube Anexos (.docx)", accept_multiple_files=True)
debug = st.checkbox("Ver Logs")

if archivos and st.button("Generar Tabla"):
    todas = []
    bar = st.progress(0)
    
    for i, a in enumerate(archivos):
        fichas = procesar_archivo_v13(a.read(), a.name)
        todas.extend(fichas)
        bar.progress((i+1)/len(archivos))

    if todas:
        todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        doc_out = generar_word_v13(todas)
        st.success("Tabla generada con datos completos.")
        st.download_button("Descargar Tabla Resumen.docx", doc_out, "Resumen_MAP_V13.docx")
    else:
        st.error("No se encontraron datos.")

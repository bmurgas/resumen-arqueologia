import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
import io

# Configuración de la página
st.set_page_config(page_title="Generador MAP V7 (Fecha Persistente + Hallazgos)", layout="wide")

def obtener_imagenes_xml(elemento_padre, doc_relacionado):
    """Busca imágenes recursivamente en el XML (celdas, grupos, etc)."""
    imagenes = []
    blips = elemento_padre.xpath('.//a:blip')
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

def procesar_archivo_v7(archivo_bytes, nombre_archivo):
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"Error leyendo {nombre_archivo}: {e}")
        return []

    fichas_extraidas = []
    
    # MEMORIA PERSISTENTE (Para fichas sin fecha)
    fecha_actual_persistente = "Sin Fecha Inicial" 

    # Recorremos el documento TABLA POR TABLA
    # Asumimos que cada tabla (o conjunto de tablas seguidas) es una ficha
    for i, tabla in enumerate(doc.tables):
        
        # Datos de ESTA ficha específica
        datos_ficha = {
            "fecha_propia": None, # La fecha que trae la tabla (si trae)
            "actividad": "",
            "hallazgos": "",
            "fotos": []
        }
        seccion_fotos = False
        
        # Analizamos fila por fila
        for fila in tabla.rows:
            # Texto limpio de toda la fila para búsquedas
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # --- 1. DETECTAR FECHA ---
            if "Fecha" in texto_fila:
                # Buscamos el valor de la fecha
                for celda in fila.cells:
                    txt = celda.text.strip()
                    # Si no es la etiqueta "Fecha" y tiene números/guiones, es el valor
                    if "Fecha" not in txt and len(txt) > 6:
                        datos_ficha["fecha_propia"] = txt
                        # ACTUALIZAMOS LA MEMORIA GLOBAL
                        fecha_actual_persistente = txt 
                        break
            
            # --- 2. DETECTAR ACTIVIDAD ---
            if "Descripción de la actividad" in texto_fila:
                # Buscamos el texto más largo de la fila que no sea el título
                mejor_texto = ""
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Descripción" in t or "Actividad" in t: continue
                    if len(t) > len(mejor_texto):
                        mejor_texto = t
                if mejor_texto:
                    datos_ficha["actividad"] = mejor_texto

            # --- 3. DETECTAR HALLAZGOS (NUEVO) ---
            if "Hallazgos arqueológicos" in texto_fila or "Hallazgos" in texto_fila:
                # Lógica simple: Si dice Ausencia y hay una X, es ausencia
                if "Ausencia" in texto_fila and ("X" in texto_fila or "x" in texto_fila):
                    datos_ficha["hallazgos"] = "Ausencia de hallazgos arqueológicos no previstos."
                elif "Presencia" in texto_fila and ("X" in texto_fila or "x" in texto_fila):
                    datos_ficha["hallazgos"] = "PRESENCIA de hallazgos arqueológicos."
                else:
                    # Si no es checkbox, intentamos capturar el texto descriptivo
                    # Buscamos celda con texto que no sea el título
                    for celda in fila.cells:
                        t = celda.text.strip()
                        if "Hallazgos" in t or "Presencia" in t or "Ausencia" in t: continue
                        if len(t) > 5:
                            datos_ficha["hallazgos"] = t
                            break

            # --- 4. DETECTAR FOTOS ---
            if "Registro fotográfico" in texto_fila:
                seccion_fotos = True
                continue

            if seccion_fotos:
                imgs = obtener_imagenes_xml(fila._element, doc)
                if imgs:
                    datos_ficha["fotos"].extend(imgs)

        # --- FINAL DE LA TABLA: GUARDAR ---
        # Usamos la fecha propia si existe, si no, la persistente
        fecha_final = datos_ficha["fecha_propia"] if datos_ficha["fecha_propia"] else fecha_actual_persistente
        
        # Solo guardamos si la ficha tiene algo útil (Actividad o Fotos)
        # Esto evita guardar tablas vacías o de formato
        if datos_ficha["actividad"] or datos_ficha["fotos"] or datos_ficha["hallazgos"]:
            
            # Formateamos el texto final para la columna central
            texto_completo = datos_ficha["actividad"]
            if datos_ficha["hallazgos"]:
                texto_completo += f"\n\n[Hallazgos: {datos_ficha['hallazgos']}]"

            item = {
                "fecha": fecha_final,
                "texto_columna_central": texto_completo,
                "fotos": datos_ficha["fotos"]
            }
            fichas_extraidas.append(item)

    return fichas_extraidas

def generar_word_v7(datos):
    doc = Document()
    doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    
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
        row[1].text = str(item["texto_columna_central"]) # Incluye actividad + hallazgos
        
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
st.title("Generador MAP V7 (Completo)")
st.info("Incluye: Fecha heredada, Hallazgos y Fotos.")

archivos = st.file_uploader("Sube Anexos (.docx)", accept_multiple_files=True)
debug = st.checkbox("Ver Logs")

if archivos and st.button("Generar Tabla"):
    todas = []
    barra = st.progress(0)
    
    for i, arch in enumerate(archivos):
        fichas = procesar_archivo_v7(arch.read(), arch.name)
        todas.extend(fichas)
        barra.progress((i + 1) / len(archivos))
        
        if debug:
            st.write(f"**{arch.name}**: {len(fichas)} registros.")

    if todas:
        # Ordenar cronológicamente (Intento simple)
        todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
        
        word_out = generar_word_v7(todas)
        st.success(f"¡Éxito! {len(todas)} registros generados.")
        st.download_button("Descargar Tabla Resumen.docx", word_out, "Resumen_MAP.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.error("No se encontraron datos.")

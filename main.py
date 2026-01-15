import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import pandas as pd
import pdfplumber
import re

# --- CONFIGURACIÓN GLOBAL ---
st.set_page_config(page_title="Arqueología - Suite de Herramientas", layout="wide")

# ==========================================
#      LÓGICA GENERADOR WORD (INTACTA)
# ==========================================

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
            
            if "Fecha" in texto_fila:
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Fecha" not in t and len(t) > 5:
                        datos_ficha["fecha_propia"] = t
                        fecha_persistente = t
                        break
            
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

            if "Ausencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "Ausencia de hallazgos arqueológicos no previstos."
            if "Presencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "PRESENCIA de hallazgos arqueológicos."

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

def generar_word_con_formato(datos):
    doc = Document()
    titulo = doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    tabla.autofit = False
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER 

    headers = tabla.rows[0].cells
    titulos = ["Fecha", "Actividades realizadas durante el MAP", "Imagen de la actividad"]
    
    for i, texto in enumerate(titulos):
        parrafo = headers[i].paragraphs[0]
        run = parrafo.add_run(texto)
        run.font.name = 'Franklin Gothic Book'
        run.font.size = Pt(9)
        run.bold = True
        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for c in tabla.columns[0].cells: c.width = Cm(2.5) 
    for c in tabla.columns[1].cells: c.width = Cm(7.5) 
    for c in tabla.columns[2].cells: c.width = Cm(8.5) 

    for item in datos:
        row = tabla.add_row().cells
        p_fecha = row[0].paragraphs[0]
        p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fecha = p_fecha.add_run(str(item["fecha"]))
        r_fecha.font.name = 'Franklin Gothic Book'
        r_fecha.font.size = Pt(9)

        p_act = row[1].paragraphs[0]
        p_act.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_act = p_act.add_run(str(item["texto_central"]))
        r_act.font.name = 'Franklin Gothic Book'
        r_act.font.size = Pt(9)
        
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

# ==========================================
#      LÓGICA EXTRACTOR PDF (CORREGIDA V21 - EL FILTRO INTELIGENTE)
# ==========================================

def limpiar_descripcion(texto_sucio):
    """
    Función 'Portero': Analiza si el texto capturado es realmente una descripción
    o si capturamos el título de la siguiente sección por error.
    """
    if not texto_sucio:
        return ""
    
    texto_sucio = texto_sucio.strip()
    
    # Palabras que indican el inicio de la SIGUIENTE sección.
    # Si el texto empieza con esto, significa que NO había descripción.
    palabras_prohibidas = [
        "CRONOLOGÍA", "OBSERVACIONES", "INTERPRETACIÓN", 
        "REGISTRO", "BIBLIOGRAFÍA", "TIPOLOGÍA"
    ]
    
    # Verificamos si lo que capturamos empieza DIRECTAMENTE con una palabra prohibida
    for palabra in palabras_prohibidas:
        if texto_sucio.upper().startswith(palabra):
            return "" # Era el título, así que la descripción es vacía.
            
    # Limpieza extra: A veces queda un "Otro" colgado de la tabla anterior
    if texto_sucio.startswith("Otro"):
        texto_sucio = texto_sucio.replace("Otro", "", 1).strip()

    return texto_sucio

def extraer_datos_pdf(archivo_bytes):
    datos_extraidos = []
    
    with pdfplumber.open(io.BytesIO(archivo_bytes)) as pdf:
        for pagina in pdf.pages:
            info = {}
            
            # --- 1. INTENTO POR TABLAS (Prioritario para celda de al lado) ---
            tablas = pagina.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    fila_segura = [str(c).strip() if c else "" for c in fila]
                    
                    for i, celda in enumerate(fila_segura):
                        # Categoría
                        if "Categoría" in celda and "SA/HA" in celda:
                            if i + 1 < len(fila_segura) and fila_segura[i+1]:
                                info["Categoría"] = fila_segura[i+1]

                        # Descripción en TABLA
                        if celda == "Descripción" or ("Descripción" in celda and "actividad" not in celda.lower()):
                            if i + 1 < len(fila_segura):
                                desc_tabla = fila_segura[i+1]
                                # Aplicamos el filtro también aquí por si acaso
                                info["Descripción"] = limpiar_descripcion(desc_tabla)

                        # Otros campos
                        if "ID Sitio" in celda:
                            if i + 1 < len(fila_segura) and fila_segura[i+1]: info["ID Sitio"] = fila_segura[i+1]
                        if "Fecha" in celda:
                            if i + 1 < len(fila_segura) and fila_segura[i+1]: info["Fecha"] = fila_segura[i+1]
                        if "Responsable" in celda:
                            if i + 1 < len(fila_segura) and fila_segura[i+1]: info["Responsable"] = fila_segura[i+1]
                        if "Coord. Central Norte" in celda:
                            if i + 1 < len(fila_segura): info["Coord. Norte"] = fila_segura[i+1]
                        if "Coord. Central Este" in celda:
                            if i + 1 < len(fila_segura): info["Coord. Este"] = fila_segura[i+1]

            # --- 2. INTENTO POR TEXTO (Respaldo inteligente) ---
            texto = pagina.extract_text()
            if texto:
                # Descripción: Regex que captura hasta la siguiente sección
                if "Descripción" not in info or not info["Descripción"]:
                    # Busca "Descripción", salta líneas opcionales, captura todo (.*?) 
                    # hasta que se encuentre con CRONOLOGIA, OBSERVACIONES, etc.
                    palabras_cierre = r"(?:CRONOLOGÍA|OBSERVACIONES|INTERPRETACIÓN|REGISTRO|BIBLIOGRAFÍA|OBSERVACIÓN|TIPOLOGÍA)"
                    patron_desc = r"Descripción\s*\n+(.*?)(?=\n\s*" + palabras_cierre + r"|$)"
                    
                    match_desc = re.search(patron_desc, texto, re.DOTALL | re.IGNORECASE)
                    
                    if match_desc:
                        raw_text = match_desc.group(1)
                        # AQUÍ LA MAGIA: Limpiamos y verificamos si es solo el título siguiente
                        info["Descripción"] = limpiar_descripcion(raw_text)
                    else:
                        info["Descripción"] = ""

                # Respaldos Regex estándar
                if not info.get("Categoría"):
                    match = re.search(r"Categoría.*?:\s*(.*?)(?:\n|Responsable|ID|$)", texto)
                    if match: info["Categoría"] = match.group(1).strip()
                if not info.get("ID Sitio"):
                    match = re.search(r"ID Sitio:?\s*([A-Za-z0-9\-]+)", texto)
                    if match: info["ID Sitio"] = match.group(1)
                if not info.get("Fecha"):
                    match = re.search(r"Fecha:?\s*(\d{2}[-/]\d{2}[-/]\d{4})", texto)
                    if match: info["Fecha"] = match.group(1)
                if not info.get("Responsable"):
                    match_resp = re.search(r"Responsable:?\s*(.*?)(?:\n|$)", texto)
                    if match_resp: info["Responsable"] = match_resp.group(1).strip()
                if not info.get("Coord. Norte"):
                    match_n = re.search(r"Norte:?\s*(\d{6,8})", texto)
                    if match_n: info["Coord. Norte"] = match_n.group(1)
                if not info.get("Coord. Este"):
                    match_e = re.search(r"Este:?\s*(\d{5,7})", texto)
                    if match_e: info["Coord. Este"] = match_e.group(1)

            # Guardar si hay datos mínimos
            if info.get("ID Sitio") or info.get("Fecha"):
                datos_extraidos.append(info)
                
    return datos_extraidos

# ==========================================
#          PÁGINAS DE LA APP
# ==========================================

def mostrar_pagina_word():
    st.title("Generador Word MAP")
    st.markdown("Crea la tabla resumen mensual a partir de los anexos diarios.")
    st.info("Configuración: Franklin Gothic Book 9 | Fotos 8x6 cm | Centrado")

    archivos = st.file_uploader("Subir Anexos Word (.docx)", accept_multiple_files=True, key="word_up")
    
    if archivos and st.button("Generar Informe Word"):
        todas = []
        bar = st.progress(0)
        for i, a in enumerate(archivos):
            fichas = procesar_archivo_v12(a.read(), a.name)
            todas.extend(fichas)
            bar.progress((i+1)/len(archivos))

        if todas:
            todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
            doc_out = generar_word_con_formato(todas)
            st.success("✅ Informe Word generado.")
            st.download_button("Descargar Word", doc_out, "Resumen_MAP.docx")
        else:
            st.error("No se encontraron datos.")

def mostrar_pagina_pdf():
    st.title("Extractor de Fichas PDF a Excel")
    st.markdown("Extrae datos de fichas de hallazgo (Tablas o Texto).")
    
    archivo_pdf = st.file_uploader("Subir PDF de Hallazgos (.pdf)", type="pdf", key="pdf_up")
    
    if archivo_pdf and st.button("Procesar PDF y Crear Excel"):
        with st.spinner("Escaneando PDF..."):
            datos = extraer_datos_pdf(archivo_pdf.read())
            
            if datos:
                df = pd.DataFrame(datos)
                cols_deseadas = ["ID Sitio", "Fecha", "Categoría", "Descripción", "Responsable", "Coord. Norte", "Coord. Este"]
                cols_finales = [c for c in cols_deseadas if c in df.columns]
                extras = [c for c in df.columns if c not in cols_deseadas]
                df = df[cols_finales + extras]
                
                st.success(f"✅ Se extrajeron {len(df)} fichas.")
                st.dataframe(df) 
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name="Hallazgos")
                
                st.download_button(
                    label="⬇️ Descargar Planilla Excel",
                    data=buffer.getvalue(),
                    file_name="Base_Datos_Hallazgos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No se encontraron datos. El PDF podría ser una imagen escaneada (sin texto seleccionable).")

# ==========================================
#        MENÚ LATERAL
# ==========================================

st.sidebar.title("Arqueología App")
opcion = st.sidebar.radio("Herramientas:", ["Generador Word (MAP)", "Extractor PDF a Excel"])

if opcion == "Generador Word (MAP)":
    mostrar_pagina_word()
elif opcion == "Extractor PDF a Excel":
    mostrar_pagina_pdf()

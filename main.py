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
            
            # FECHA
            if "Fecha" in texto_fila:
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Fecha" not in t and len(t) > 5:
                        datos_ficha["fecha_propia"] = t
                        fecha_persistente = t
                        break
            
            # ACTIVIDAD
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

            # HALLAZGOS
            if "Ausencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "Ausencia de hallazgos arqueológicos no previstos."
            if "Presencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos"] = "PRESENCIA de hallazgos arqueológicos."

            # FOTOS
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
#      LÓGICA EXTRACTOR PDF (NUEVO)
# ==========================================

def extraer_datos_pdf(archivo_bytes):
    """
    Lee un PDF y busca patrones específicos usando Regex.
    """
    datos_extraidos = []
    
    with pdfplumber.open(io.BytesIO(archivo_bytes)) as pdf:
        # Asumimos que cada página es una ficha o contiene una ficha
        # Si una ficha ocupa más de una página, esto captura por página.
        
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto: continue
            
            info = {}
            
            # --- 1. ID SITIO ---
            # Buscamos "ID Sitio" o patrones tipo "HA-01"
            # Regex: Busca "ID Sitio:" seguido de algo, O busca "HA-" seguido de números
            match_id = re.search(r"ID Sitio:?\s*([A-Za-z0-9\-]+)", texto, re.IGNORECASE)
            if not match_id:
                # Intento alternativo: buscar código tipo HA-XX
                match_id = re.search(r"(HA-\d+|SA-\d+)", texto)
            info["ID Sitio"] = match_id.group(1) if match_id else "No encontrado"

            # --- 2. FECHA ---
            # Busca patrones DD-MM-AAAA o DD/MM/AAAA
            match_fecha = re.search(r"(\d{2}[-/]\d{2}[-/]\d{4})", texto)
            info["Fecha"] = match_fecha.group(1) if match_fecha else "No encontrada"

            # --- 3. COORDENADAS ---
            # Busca "Norte" seguido de números (6-7 dígitos)
            match_norte = re.search(r"Norte:?\s*.*?(\d{6,8})", texto, re.IGNORECASE | re.DOTALL)
            info["Coord. Norte"] = match_norte.group(1) if match_norte else ""

            # Busca "Este" seguido de números
            match_este = re.search(r"Este:?\s*.*?(\d{5,7})", texto, re.IGNORECASE | re.DOTALL)
            info["Coord. Este"] = match_este.group(1) if match_este else ""

            # --- 4. CATEGORÍA (SA/HA) ---
            # Busca "Categoría" y toma las palabras siguientes, o busca directamente HA/SA aislado
            match_cat = re.search(r"Categoría.*?(SA|HA)", texto, re.IGNORECASE | re.DOTALL)
            if match_cat:
                info["Categoría"] = match_cat.group(1)
            else:
                # Si no encuentra "Categoría:", busca si hay un "HA" o "SA" suelto que no sea el ID
                if "HA" in texto and "SA" not in texto: info["Categoría"] = "HA"
                elif "SA" in texto and "HA" not in texto: info["Categoría"] = "SA"
                else: info["Categoría"] = ""

            # --- 5. RESPONSABLE ---
            # Busca "Responsable:" y captura hasta el salto de línea
            match_resp = re.search(r"Responsable:?\s*(.*?)(?:\n|$)", texto, re.IGNORECASE)
            if match_resp:
                # Limpiamos si agarró caracteres raros
                clean_resp = match_resp.group(1).replace("\n", " ").strip()
                info["Responsable"] = clean_resp
            else:
                info["Responsable"] = ""

            # --- 6. DESCRIPCIÓN ---
            # Esta es difícil. Buscamos bloques de texto comunes.
            # Intento 1: Buscar "Descripción" y tomar lo que sigue
            match_desc = re.search(r"Descripción.*?\n(.*?)(?:\n\n|Evidencias|Cronología|$)", texto, re.IGNORECASE | re.DOTALL)
            if match_desc:
                info["Descripción"] = match_desc.group(1).strip()
            else:
                # Intento 2: Buscar "Descripción de las evidencias"
                match_desc2 = re.search(r"descripción de las evidencias\s*\n(.*?)(?:\n\n|Asociación|$)", texto, re.IGNORECASE | re.DOTALL)
                info["Descripción"] = match_desc2.group(1).strip() if match_desc2 else ""

            # Solo agregamos si encontramos al menos un ID o una Fecha para evitar páginas vacías
            if info["ID Sitio"] != "No encontrado" or info["Fecha"] != "No encontrada":
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
    st.markdown("Extrae ID, Coordenadas, Categoría, Descripción y Responsables de las fichas PDF.")
    
    archivo_pdf = st.file_uploader("Subir PDF de Hallazgos (.pdf)", type="pdf", key="pdf_up")
    
    if archivo_pdf and st.button("Procesar PDF y Crear Excel"):
        with st.spinner("Leyendo PDF... esto puede tomar unos segundos."):
            datos = extraer_datos_pdf(archivo_pdf.read())
            
            if datos:
                df = pd.DataFrame(datos)
                
                # Reordenar columnas para que salga bonito
                columnas_orden = ["ID Sitio", "Fecha", "Categoría", "Coord. Norte", "Coord. Este", "Responsable", "Descripción"]
                # Asegurarnos de que existan en el DF (por si acaso)
                cols_existentes = [c for c in columnas_orden if c in df.columns]
                df = df[cols_existentes]
                
                st.success(f"✅ Se extrajeron {len(df)} registros.")
                st.dataframe(df) # Muestra una vista previa en la web
                
                # Generar Excel en memoria
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
                st.error("No se pudo extraer información. Verifica que el PDF tenga texto seleccionable (no sea solo imagen escaneada).")

# ==========================================
#        MENÚ LATERAL
# ==========================================

st.sidebar.title("Arqueología App")
opcion = st.sidebar.radio("Herramientas:", ["Generador Word (MAP)", "Extractor PDF a Excel"])

if opcion == "Generador Word (MAP)":
    mostrar_pagina_word()
elif opcion == "Extractor PDF a Excel":
    mostrar_pagina_pdf()

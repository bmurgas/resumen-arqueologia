import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT # Para poner la hoja horizontal
import io
import pandas as pd
import pdfplumber
import re
from PIL import Image as PILImage # Necesario para procesar las fotos del PDF

# --- CONFIGURACIÓN GLOBAL ---
st.set_page_config(page_title="Arqueología - Suite de Herramientas", layout="wide")

# ==========================================
#      LÓGICA GENERADOR WORD MAP (INTACTA)
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
#      LÓGICA EXTRACTOR PDF A EXCEL (INTACTA)
# ==========================================

def extraer_datos_pdf(archivo_bytes):
    datos_extraidos = []
    palabras_fin = ["CRONOLOGÍA", "OBSERVACIONES", "INTERPRETACIÓN", "REGISTRO", "BIBLIOGRAFÍA", "TIPOLOGÍA", "ASOCIACIÓN"]
    
    with pdfplumber.open(io.BytesIO(archivo_bytes)) as pdf:
        for pagina in pdf.pages:
            info = {}
            
            # 1. TABLAS
            tablas = pagina.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    fila = [str(c).strip() if c else "" for c in fila]
                    for i, celda in enumerate(fila):
                        if "Categoría" in celda and "SA/HA" in celda:
                            if i + 1 < len(fila): info["Categoría"] = fila[i+1]
                        if "ID Sitio" in celda:
                            if i + 1 < len(fila): info["ID Sitio"] = fila[i+1]
                        if "Fecha" in celda:
                            if i + 1 < len(fila): info["Fecha"] = fila[i+1]
                        if "Responsable" in celda:
                            if i + 1 < len(fila): info["Responsable"] = fila[i+1]
                        if "Coord. Central Norte" in celda:
                            if i + 1 < len(fila): info["Coord. Norte"] = fila[i+1]
                        if "Coord. Central Este" in celda:
                            if i + 1 < len(fila): info["Coord. Este"] = fila[i+1]

            # 2. TEXTO
            texto = pagina.extract_text()
            if texto:
                if not info.get("Descripción"):
                    patron = r"Descripción\s*\n*(.*?)(?=" + "|".join(palabras_fin) + r"|$)"
                    match = re.search(patron, texto, re.DOTALL | re.IGNORECASE)
                    if match:
                        contenido = match.group(1).strip()
                        es_titulo = False
                        for p in palabras_fin:
                            if contenido.upper().startswith(p):
                                es_titulo = True; break
                        if es_titulo: info["Descripción"] = ""
                        else:
                            if contenido.startswith("Otro"): contenido = contenido.replace("Otro", "", 1).strip()
                            info["Descripción"] = contenido
                    else: info["Descripción"] = ""

                if not info.get("ID Sitio"):
                    match = re.search(r"ID Sitio:?\s*([A-Za-z0-9\-]+)", texto)
                    if match: info["ID Sitio"] = match.group(1)
                if not info.get("Fecha"):
                    match = re.search(r"Fecha:?\s*(\d{2}[-/]\d{2}[-/]\d{4})", texto)
                    if match: info["Fecha"] = match.group(1)
                if not info.get("Categoría"):
                    match = re.search(r"Categoría.*?:\s*(.*?)(?:\n|Responsable|ID|$)", texto)
                    if match: info["Categoría"] = match.group(1).strip()
                if not info.get("Coord. Norte"):
                    match = re.search(r"Norte:?\s*(\d{6,8})", texto)
                    if match: info["Coord. Norte"] = match.group(1)
                if not info.get("Coord. Este"):
                    match = re.search(r"Este:?\s*(\d{5,7})", texto)
                    if match: info["Coord. Este"] = match.group(1)
                if not info.get("Responsable"):
                    match = re.search(r"Responsable:?\s*(.*?)(?:\n|$)", texto)
                    if match: info["Responsable"] = match.group(1).strip()

            if info.get("ID Sitio") or info.get("Fecha"):
                datos_extraidos.append(info)
                
    return datos_extraidos

# ==========================================
#   NUEVA LÓGICA: GENERADOR FICHAS CON FOTO (V23)
# ==========================================

def generar_fichas_con_foto(archivo_bytes):
    """
    1. Extrae datos (Lógica V22).
    2. Extrae la 4ta foto de la página (Lógica 'Fotografía detalle').
    3. Genera Word en Horizontal.
    """
    datos_completos = []
    
    # 1. Extracción de Datos y Fotos
    palabras_fin = ["CRONOLOGÍA", "OBSERVACIONES", "INTERPRETACIÓN", "REGISTRO", "BIBLIOGRAFÍA", "TIPOLOGÍA", "ASOCIACIÓN"]
    
    with pdfplumber.open(io.BytesIO(archivo_bytes)) as pdf:
        for pagina in pdf.pages:
            info = {}
            
            # --- A. DATOS (Misma lógica que extractor Excel) ---
            tablas = pagina.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    fila = [str(c).strip() if c else "" for c in fila]
                    for i, celda in enumerate(fila):
                        if "Categoría" in celda and "SA/HA" in celda:
                            if i + 1 < len(fila): info["Categoría"] = fila[i+1]
                        if "ID Sitio" in celda:
                            if i + 1 < len(fila): info["ID Sitio"] = fila[i+1]
                        if "Fecha" in celda:
                            if i + 1 < len(fila): info["Fecha"] = fila[i+1]
                        if "Responsable" in celda:
                            if i + 1 < len(fila): info["Responsable"] = fila[i+1]
                        if "Coord. Central Norte" in celda:
                            if i + 1 < len(fila): info["Coord. Norte"] = fila[i+1]
                        if "Coord. Central Este" in celda:
                            if i + 1 < len(fila): info["Coord. Este"] = fila[i+1]

            texto = pagina.extract_text()
            if texto:
                if not info.get("Descripción"):
                    patron = r"Descripción\s*\n*(.*?)(?=" + "|".join(palabras_fin) + r"|$)"
                    match = re.search(patron, texto, re.DOTALL | re.IGNORECASE)
                    if match:
                        contenido = match.group(1).strip()
                        es_titulo = False
                        for p in palabras_fin:
                            if contenido.upper().startswith(p):
                                es_titulo = True; break
                        if es_titulo: info["Descripción"] = ""
                        else:
                            if contenido.startswith("Otro"): contenido = contenido.replace("Otro", "", 1).strip()
                            info["Descripción"] = contenido
                    else: info["Descripción"] = ""
                
                # Respaldos
                if not info.get("ID Sitio"):
                    match = re.search(r"ID Sitio:?\s*([A-Za-z0-9\-]+)", texto)
                    if match: info["ID Sitio"] = match.group(1)
                # (Se omiten otros respaldos por brevedad, asumiendo que el PDF es estándar, 
                # pero se pueden agregar igual que en la función anterior)

            # --- B. EXTRACCIÓN DE LA 4TA FOTO ---
            # pdfplumber retorna las imágenes como diccionarios con coordenadas
            imagenes_pagina = pagina.images
            foto_bytes = None
            
            # Necesitamos la 4ta foto (índice 3). 
            if len(imagenes_pagina) >= 4:
                img_data = imagenes_pagina[3] # La cuarta foto
                # Recortamos la imagen de la página original
                # bbox = (x0, top, x1, bottom)
                bbox = (img_data['x0'], img_data['top'], img_data['x1'], img_data['bottom'])
                
                try:
                    # Usamos .crop() y .to_image() de pdfplumber
                    cropped_page = pagina.crop(bbox)
                    img_obj = cropped_page.to_image(resolution=300) # Alta resolución
                    
                    # Convertir a bytes para Word
                    buf = io.BytesIO()
                    img_obj.save(buf, format="JPEG")
                    foto_bytes = buf.getvalue()
                except:
                    foto_bytes = None
            
            # Guardamos todo
            if info.get("ID Sitio") or info.get("Fecha") or foto_bytes:
                info["foto_blob"] = foto_bytes
                datos_completos.append(info)

    return datos_completos

def crear_doc_horizontal(datos):
    doc = Document()
    
    # Cambiar orientación a Horizontal (Landscape)
    section = doc.sections[0]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height
    
    # Márgenes un poco más estrechos para aprovechar espacio
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    for i, item in enumerate(datos):
        # Título de la Ficha
        id_sitio = item.get('ID Sitio', 'Sin ID')
        doc.add_heading(f"Ficha de Registro - {id_sitio}", level=1)
        
        # --- TABLA DE DATOS ---
        tabla = doc.add_table(rows=0, cols=2)
        tabla.style = 'Table Grid'
        
        # Lista de campos a mostrar
        campos = [
            ("ID Sitio", item.get("ID Sitio", "")),
            ("Fecha", item.get("Fecha", "")),
            ("Responsable", item.get("Responsable", "")),
            ("Categoría (SA/HA)", item.get("Categoría", "")),
            ("Coord. Norte", item.get("Coord. Norte", "")),
            ("Coord. Este", item.get("Coord. Este", "")),
            ("Descripción", item.get("Descripción", ""))
        ]
        
        for clave, valor in campos:
            row = tabla.add_row().cells
            row[0].text = clave
            row[0].paragraphs[0].runs[0].bold = True
            row[1].text = str(valor)
            
        doc.add_paragraph("\n") # Espacio
        
        # --- FOTO DETALLE ---
        if item.get("foto_blob"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            # Agregamos la foto con un tamaño razonable para hoja horizontal
            # Altura de unos 10-12 cm debería caber bien
            try:
                run.add_picture(io.BytesIO(item["foto_blob"]), height=Cm(11))
                p.add_run("\nFotografía detalle (Extraída de PDF)").italic = True
            except:
                p.add_run("[Error al insertar imagen]")
        else:
            doc.add_paragraph("[No se detectó la 4ta foto en la página original]")

        # Salto de página entre fichas
        if i < len(datos) - 1:
            doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

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
        else: st.error("No se encontraron datos.")

def mostrar_pagina_pdf_excel():
    st.title("Extractor de Fichas PDF a Excel")
    st.markdown("Extrae datos de fichas de hallazgo.")
    archivo_pdf = st.file_uploader("Subir PDF de Hallazgos (.pdf)", type="pdf", key="pdf_up")
    if archivo_pdf and st.button("Procesar PDF y Crear Excel"):
        with st.spinner("Escaneando PDF..."):
            datos = extraer_datos_pdf(archivo_pdf.read())
            if datos:
                df = pd.DataFrame(datos)
                orden_columnas = ["ID Sitio", "Coord. Norte", "Coord. Este", "Categoría", "Descripción", "Fecha", "Responsable"]
                cols_finales = [c for c in orden_columnas if c in df.columns]
                extras = [c for c in df.columns if c not in orden_columnas]
                df_final = df[cols_finales + extras]
                df_final = df_final.rename(columns={"Categoría": "Categoría (SA/HA)"})
                st.success(f"✅ Se extrajeron {len(df_final)} fichas.")
                st.dataframe(df_final) 
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name="Hallazgos")
                st.download_button("⬇️ Descargar Planilla Excel", buffer.getvalue(), "Base_Datos_Hallazgos.xlsx")
            else: st.error("No se encontraron datos.")

def mostrar_pagina_fichas_con_fotos():
    st.title("Generador de Fichas Individuales (Con Foto)")
    st.markdown("Genera un documento Word con los datos extraídos + la **cuarta foto** (Fotografía detalle) de cada ficha.")
    st.info("Formato: Hoja Horizontal (Landscape) para mejor ajuste.")
    
    archivo_pdf = st.file_uploader("Subir PDF de Hallazgos (.pdf)", type="pdf", key="pdf_foto_up")
    
    if archivo_pdf and st.button("Generar Fichas Word"):
        with st.spinner("Procesando PDF (Texto e Imágenes)..."):
            datos_foto = generar_fichas_con_foto(archivo_pdf.read())
            
            if datos_foto:
                st.success(f"✅ Se generaron {len(datos_foto)} fichas con foto.")
                doc_bytes = crear_doc_horizontal(datos_foto)
                st.download_button(
                    label="⬇️ Descargar Fichas con Fotos (.docx)",
                    data=doc_bytes,
                    file_name="Fichas_Con_Fotos.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error("No se pudieron extraer datos o fotos del PDF.")

# ==========================================
#        MENÚ LATERAL
# ==========================================

st.sidebar.title("Arqueología App")
opcion = st.sidebar.radio("Herramientas:", ["Generador Word (MAP)", "Extractor PDF a Excel", "Generador Fichas con Foto"])

if opcion == "Generador Word (MAP)":
    mostrar_pagina_word()
elif opcion == "Extractor PDF a Excel":
    mostrar_pagina_pdf_excel()
elif opcion == "Generador Fichas con Foto":
    mostrar_pagina_fichas_con_fotos()

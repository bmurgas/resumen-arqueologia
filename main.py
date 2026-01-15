import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
import io
import pandas as pd
import pdfplumber
import re
from PIL import Image as PILImage

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
#   NUEVO PROCESADOR MAESTRO (DESDE WORD)
# ==========================================

def procesar_maestro_desde_word(archivo_bytes, nombre_archivo):
    """
    Lee DOCX de fichas de hallazgo.
    Extrae datos y FOTOS de la celda 'Fotografía detalle'.
    """
    try:
        doc = Document(io.BytesIO(archivo_bytes))
    except Exception as e:
        st.error(f"Error leyendo {nombre_archivo}: {e}")
        return []

    fichas = []
    
    # Iteramos tablas buscando "ID Sitio" para identificar una ficha
    for tabla in doc.tables:
        info = {
            "ID Sitio": "", "Coord. Norte": "", "Coord. Este": "", 
            "Categoría": "", "Descripción": "", "Fecha": "", 
            "Responsable": "", "Cronología": "", "foto_blob": None
        }
        
        es_ficha = False
        crono_partes = []

        # Recorremos filas
        for r_idx, fila in enumerate(tabla.rows):
            for c_idx, celda in enumerate(fila.cells):
                txt = celda.text.strip()
                
                # --- IDENTIFICACIÓN ---
                if "ID Sitio" in txt:
                    if c_idx + 1 < len(fila.cells):
                        info["ID Sitio"] = fila.cells[c_idx+1].text.strip()
                        es_ficha = True
                
                if "Fecha" in txt:
                    if c_idx + 1 < len(fila.cells):
                        info["Fecha"] = fila.cells[c_idx+1].text.strip()
                        
                if "Responsable" in txt:
                    if c_idx + 1 < len(fila.cells):
                        info["Responsable"] = fila.cells[c_idx+1].text.strip()

                if "Categoría" in txt:
                    if c_idx + 1 < len(fila.cells):
                        info["Categoría"] = fila.cells[c_idx+1].text.strip()

                if "Coord. Central Norte" in txt:
                    if c_idx + 1 < len(fila.cells):
                        info["Coord. Norte"] = fila.cells[c_idx+1].text.strip()
                if "Coord. Central Este" in txt:
                    if c_idx + 1 < len(fila.cells):
                        info["Coord. Este"] = fila.cells[c_idx+1].text.strip()

                # --- DESCRIPCIÓN ---
                if "Descripción" == txt or "Descripción:" == txt: # Exacto
                    if c_idx + 1 < len(fila.cells):
                         # Evitar confundir con "Descripción de las evidencias"
                         vecino = fila.cells[c_idx+1].text.strip()
                         if vecino: info["Descripción"] = vecino

                # --- CRONOLOGÍA (Detectar X en celda vecina) ---
                opciones = ["Prehispánico", "Subactual", "Incierto", "Histórico"]
                for op in opciones:
                    if op in txt:
                        if c_idx + 1 < len(fila.cells):
                            val_vecino = fila.cells[c_idx+1].text.strip().upper()
                            if "X" in val_vecino:
                                crono_partes.append(op)
                
                if "Periodo específico" in txt:
                    if c_idx + 1 < len(fila.cells):
                        val = fila.cells[c_idx+1].text.strip()
                        if val: crono_partes.append(val)

                # --- FOTO DETALLE ---
                # Si encontramos la etiqueta, buscamos la imagen en la celda de ARRIBA (misma columna, fila anterior)
                if "Fotografía detalle" in txt:
                    if r_idx > 0: # Asegurar que hay fila arriba
                        celda_arriba = tabla.rows[r_idx - 1].cells[c_idx]
                        
                        # Extraer imagen de esa celda
                        imgs = obtener_imagenes_con_id(celda_arriba._element, doc)
                        if imgs:
                            # Tomamos la primera imagen de esa celda (generalmente es única)
                            info["foto_blob"] = imgs[0][1] # blob

        if crono_partes:
            info["Cronología"] = ", ".join(list(set(crono_partes)))

        # Guardamos ficha si tiene ID
        if es_ficha and info["ID Sitio"]:
            fichas.append(info)

    return fichas

def crear_doc_tabla_horizontal(datos):
    doc = Document()
    
    section = doc.sections[0]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    doc.add_heading("Fichas de Hallazgos (Resumen)", 0)

    tabla = doc.add_table(rows=1, cols=9)
    tabla.style = 'Table Grid'
    
    titulos = ["ID Sitio", "Coord. Norte", "Coord. Este", "Cat. (SA/HA)", "Descripción", "Fecha", "Responsable", "Cronología", "Foto"]
    headers = tabla.rows[0].cells
    for i, t in enumerate(titulos):
        headers[i].text = t
        headers[i].paragraphs[0].runs[0].bold = True

    for item in datos:
        row = tabla.add_row().cells
        row[0].text = str(item.get("ID Sitio", ""))
        row[1].text = str(item.get("Coord. Norte", ""))
        row[2].text = str(item.get("Coord. Este", ""))
        row[3].text = str(item.get("Categoría", ""))
        row[4].text = str(item.get("Descripción", ""))
        row[5].text = str(item.get("Fecha", ""))
        row[6].text = str(item.get("Responsable", ""))
        row[7].text = str(item.get("Cronología", ""))
        
        if item.get("foto_blob"):
            p = row[8].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                run = p.add_run()
                run.add_picture(io.BytesIO(item["foto_blob"]), width=Cm(4.5)) 
            except:
                p.add_run("[Err]")
        else:
            row[8].text = "[Sin Foto]"

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

def mostrar_pagina_maestro_word():
    st.title("Procesador Maestro (Desde Word)")
    st.markdown("Extrae datos y fotos desde las Fichas de Hallazgo originales en Word.")
    st.info("Genera 2 archivos: Excel con datos y Word con fichas horizontales.")
    
    archivos = st.file_uploader("Subir Fichas de Hallazgo (.docx)", accept_multiple_files=True, key="maestro_up")
    
    if archivos and st.button("Procesar Archivos"):
        todos_datos = []
        bar = st.progress(0)
        for i, a in enumerate(archivos):
            datos = procesar_maestro_desde_word(a.read(), a.name)
            todos_datos.extend(datos)
            bar.progress((i+1)/len(archivos))
            
        if todos_datos:
            st.success(f"✅ Se procesaron {len(todos_datos)} fichas correctamente.")
            
            # 1. Generar Excel
            df = pd.DataFrame(todos_datos)
            # Limpiar columna foto para el excel
            df_excel = df.drop(columns=["foto_blob"], errors='ignore')
            orden = ["ID Sitio", "Coord. Norte", "Coord. Este", "Categoría", "Descripción", "Fecha", "Responsable", "Cronología"]
            cols = [c for c in orden if c in df_excel.columns]
            df_excel = df_excel[cols]
            
            buf_excel = io.BytesIO()
            with pd.ExcelWriter(buf_excel, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False, sheet_name="Hallazgos")
                
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("⬇️ Descargar Excel", buf_excel.getvalue(), "Base_Datos_Hallazgos.xlsx")
            
            # 2. Generar Word con Fotos
            buf_word = crear_doc_tabla_horizontal(todos_datos)
            with col2:
                st.download_button("⬇️ Descargar Fichas Word", buf_word.getvalue(), "Fichas_Con_Fotos.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
            st.dataframe(df_excel)
        else:
            st.error("No se encontraron fichas válidas (con ID Sitio) en los documentos.")

# ==========================================
#        MENÚ LATERAL
# ==========================================

st.sidebar.title("Arqueología App")
opcion = st.sidebar.radio("Herramientas:", [
    "Generador Word (MAP)", 
    "Procesador Maestro (Desde Word)"
])

if opcion == "Generador Word (MAP)":
    mostrar_pagina_word()
elif opcion == "Procesador Maestro (Desde Word)":
    mostrar_pagina_maestro_word()

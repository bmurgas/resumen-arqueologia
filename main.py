import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Plataforma Arqueología", layout="wide")

# --- 2. FUNCIONES DE EXTRACCIÓN (Lógica V13) ---
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

def procesar_archivo_completo(archivo_bytes, nombre_archivo):
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
            "categoria": "",
            "descripcion_item": "",
            "hallazgos_check": "",
            "items_foto": [] 
        }
        
        rids_procesados = set()
        celdas_procesadas = set()
        en_seccion_fotos = False
        
        for r_idx, fila in enumerate(tabla.rows):
            texto_fila = " ".join([c.text.strip() for c in fila.cells]).strip()
            
            # A) FECHA
            if "Fecha" in texto_fila:
                for celda in fila.cells:
                    t = celda.text.strip()
                    if "Fecha" not in t and len(t) > 5:
                        datos_ficha["fecha_propia"] = t
                        fecha_persistente = t
                        break
            
            # B) EXTRACCIÓN EXACTA (Celda vecina)
            for c_idx, celda in enumerate(fila.cells):
                txt = celda.text.strip()
                
                # B.1 Categoría (Busca celda derecha)
                if "Categoría" in txt: 
                    if c_idx + 1 < len(fila.cells):
                        datos_ficha["categoria"] = fila.cells[c_idx + 1].text.strip()
                
                # B.2 Descripción del Hallazgo (Busca celda derecha, evita confundir con Actividad)
                elif "Descripción" in txt and "actividad" not in txt:
                    if c_idx + 1 < len(fila.cells):
                        datos_ficha["descripcion_item"] = fila.cells[c_idx + 1].text.strip()

            # C) ACTIVIDAD GENERAL
            if "Descripción de la actividad" in texto_fila:
                mejor_texto = ""
                celdas_vistas = set()
                for celda in fila.cells:
                    if celda in celdas_vistas: continue
                    celdas_vistas.add(celda)
                    t = celda.text.strip()
                    if "Descripción" in t or "Actividad" in t: continue
                    if len(t) > len(mejor_texto):
                        mejor_texto = t
                if mejor_texto:
                    datos_ficha["actividad"] = mejor_texto

            # D) HALLAZGOS (Check X)
            if "Ausencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos_check"] = "Ausencia de hallazgos arqueológicos no previstos."
            if "Presencia" in texto_fila and any(c.text.strip().upper() == "X" for c in fila.cells):
                datos_ficha["hallazgos_check"] = "PRESENCIA de hallazgos arqueológicos."

            # E) FOTOS
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

        # CONSOLIDACIÓN
        fecha_final = datos_ficha["fecha_propia"] if datos_ficha["fecha_propia"] else fecha_persistente
        
        hay_info = (datos_ficha["actividad"] or datos_ficha["items_foto"] or 
                    datos_ficha["categoria"] or datos_ficha["descripcion_item"])
        
        if hay_info:
            partes_texto = []
            if datos_ficha["actividad"]:
                partes_texto.append(datos_ficha["actividad"])
            if datos_ficha["categoria"]:
                partes_texto.append(f"Categoría: {datos_ficha['categoria']}")
            if datos_ficha["descripcion_item"]:
                partes_texto.append(f"Descripción: {datos_ficha['descripcion_item']}")
            if datos_ficha["hallazgos_check"]:
                partes_texto.append(f"\n[Hallazgos: {datos_ficha['hallazgos_check']}]")
            
            texto_consolidado = "\n\n".join(partes_texto)
            
            fichas_extraidas.append({
                "fecha": fecha_final, 
                "texto_central": texto_consolidado, 
                "fotos": datos_ficha["items_foto"]
            })

    return fichas_extraidas

# --- 3. GENERACIÓN WORD (ESTILO FRANKLIN + FOTOS 8x6) ---
def generar_word_estilo_final(datos):
    doc = Document()
    
    # Título
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

    # Anchos (cm)
    for c in tabla.columns[0].cells: c.width = Cm(2.5) 
    for c in tabla.columns[1].cells: c.width = Cm(7.5) 
    for c in tabla.columns[2].cells: c.width = Cm(8.5) 

    for item in datos:
        row = tabla.add_row().cells
        
        # COL 1: Fecha
        p_fecha = row[0].paragraphs[0]
        p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fecha = p_fecha.add_run(str(item["fecha"]))
        r_fecha.font.name = 'Franklin Gothic Book'
        r_fecha.font.size = Pt(9)

        # COL 2: Texto
        p_act = row[1].paragraphs[0]
        p_act.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_act = p_act.add_run(str(item["texto_central"]))
        r_act.font.name = 'Franklin Gothic Book'
        r_act.font.size = Pt(9)
        
        # COL 3: Foto
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

# --- 4. INTERFAZ (MENÚ LATERAL - SIN CLAVE) ---
st.sidebar.title("Menú Principal")
opcion = st.sidebar.radio("Seleccione una herramienta:", ["Generador de Informes (Word)", "Extractor PDF (Próximamente)"])

if opcion == "Generador de Informes (Word)":
    st.title("Generador de Tabla Resumen MAP")
    st.markdown("---")
    st.info("Sube los anexos diarios (.docx). El sistema extraerá Actividades, Categorías, Descripciones y Fotos.")

    archivos = st.file_uploader("Sube los archivos aquí", type=["docx"], accept_multiple_files=True)

    if archivos and st.button("Generar Informe"):
        todas_fichas = []
        bar = st.progress(0)
        
        for i, a in enumerate(archivos):
            fichas = procesar_archivo_completo(a.read(), a.name)
            todas_fichas.extend(fichas)
            bar.progress((i+1)/len(archivos))

        if todas_fichas:
            todas_fichas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
            doc_final = generar_word_estilo_final(todas_fichas)
            st.success(f"¡Listo! Se procesaron {len(todas_fichas)} registros.")
            st.download_button("⬇️ Descargar Word", doc_final, "Resumen_MAP_Final.docx")
        else:
            st.warning("No se encontraron datos en los archivos.")

elif opcion == "Extractor PDF (Próximamente)":
    st.title("Herramienta PDF a Excel")
    st.info("Aquí podrás subir tus PDF antiguos para pasarlos a Excel.")

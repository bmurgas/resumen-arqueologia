import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

# --- 1. CONFIGURACIÓN GLOBAL ---
st.set_page_config(page_title="Arqueología - Suite de Herramientas", layout="wide")

# ==========================================
#      BLOQUE DE LÓGICA (GENERADOR WORD)
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
    # Lógica de extracción V12 (Intacta)
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
    # Lógica de generación V12 (Intacta con estilos Franklin)
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
        
        # FECHA
        p_fecha = row[0].paragraphs[0]
        p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_fecha = p_fecha.add_run(str(item["fecha"]))
        r_fecha.font.name = 'Franklin Gothic Book'
        r_fecha.font.size = Pt(9)

        # TEXTO
        p_act = row[1].paragraphs[0]
        p_act.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_act = p_act.add_run(str(item["texto_central"]))
        r_act.font.name = 'Franklin Gothic Book'
        r_act.font.size = Pt(9)
        
        # FOTOS
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
#          PÁGINAS DE LA APLICACIÓN
# ==========================================

def mostrar_pagina_word():
    st.title("Generador MAP (Informe Word)")
    st.markdown("### Convierte Anexos Diarios en Tabla Resumen")
    st.info("Formato: Franklin Gothic Book 9 | Fotos 8x6 cm | Centrado")

    archivos = st.file_uploader("Cargar Anexos (.docx)", accept_multiple_files=True, key="upload_word")
    debug = st.checkbox("Ver detalles técnicos", key="debug_word")

    if archivos and st.button("Generar Informe", key="btn_word"):
        todas = []
        bar = st.progress(0)
        
        for i, a in enumerate(archivos):
            fichas = procesar_archivo_v12(a.read(), a.name)
            todas.extend(fichas)
            bar.progress((i+1)/len(archivos))
            
            if debug:
                st.write(f"📄 {a.name}: {len(fichas)} fichas detectadas.")

        if todas:
            todas.sort(key=lambda x: x['fecha'] if x['fecha'] else "ZZZ")
            doc_out = generar_word_con_formato(todas)
            st.success("✅ Informe generado exitosamente.")
            st.download_button("⬇️ Descargar Word", doc_out, "Resumen_MAP_V12.docx")
        else:
            st.error("⚠️ No se encontraron datos válidos.")

def mostrar_pagina_pdf():
    st.title("Extractor de Datos (PDF a Excel)")
    st.markdown("---")
    st.warning("🚧 **En construcción**")
    st.write("Aquí implementaremos la función para transformar informes PDF antiguos o escaneados directamente a planillas Excel.")
    
    # Placeholder visual para que veas cómo quedaría
    st.file_uploader("Cargar PDF (Demostración)", type="pdf", disabled=True)
    st.button("Convertir a Excel", disabled=True)

# ==========================================
#        MENÚ DE NAVEGACIÓN PRINCIPAL
# ==========================================

# Menú lateral
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/1087/1087840.png", width=100) # Icono genérico (opcional)
st.sidebar.title("Menú Principal")

opcion_seleccionada = st.sidebar.radio(
    "Selecciona una herramienta:",
    ["📄 Generador Word MAP", "📊 Extractor PDF (Próx.)"]
)

st.sidebar.markdown("---")
st.sidebar.info("Versión 13.0 - Multi-Herramienta")

# Lógica de visualización
if opcion_seleccionada == "📄 Generador Word MAP":
    mostrar_pagina_word()
elif opcion_seleccionada == "📊 Extractor PDF (Próx.)":
    mostrar_pagina_pdf()

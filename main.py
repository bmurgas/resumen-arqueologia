import streamlit as st
from docx import Document
from docx.shared import Inches
import io

def extraer_fichas_con_fotos(archivo_binario):
    # Abrimos el documento desde la memoria
    doc = Document(io.BytesIO(archivo_binario))
    lista_fichas = []
    
    # Buscamos todas las tablas (cada ficha es una tabla o serie de tablas)
    for tabla in doc.tables:
        ficha = {"fecha": "Sin fecha", "actividad": "", "imagenes": []}
        en_seccion_fotos = False
        
        for fila in tabla.rows:
            # Limpiamos el texto de la primera celda para identificar la sección
            if not fila.cells or len(fila.cells) < 2: continue
            encabezado = fila.cells[0].text.strip()
            
            # 1. Capturar Fecha
            if "Fecha" in encabezado:
                ficha["fecha"] = fila.cells[1].text.strip()
            
            # 2. Capturar Actividad
            if "Descripción de la actividad" in encabezado:
                ficha["actividad"] = fila.cells[1].text.strip()
            
            # 3. Identificar sección de fotos
            if "Registro fotográfico" in encabezado:
                en_seccion_fotos = True
                continue # Saltamos la fila del encabezado
            
            # 4. Extraer imágenes binarias si estamos en la sección correcta
            if en_seccion_fotos:
                for celda in fila.cells:
                    for parrafo in celda.paragraphs:
                        for run in parrafo.runs:
                            # Buscamos elementos de imagen (blip) en el XML del run
                            blips = run._element.xpath('.//a:blip')
                            for blip in blips:
                                rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                image_part = doc.part.related_parts[rId]
                                # Guardamos los bytes de la imagen y su extensión
                                ficha["imagenes"].append(image_part.blob)

        # Si encontramos una fecha, asumimos que es una ficha válida y la guardamos
        if ficha["fecha"] != "Sin fecha":
            lista_fichas.append(ficha)
            
    return lista_fichas

def crear_word_final(datos_acumulados):
    dest_doc = Document()
    # Configuración de la tabla: 3 columnas [Fecha, Actividad, Imagen]
    tabla_final = dest_doc.add_table(rows=1, cols=3)
    tabla_final.style = 'Table Grid'
    
    # Encabezados
    etiquetas = ['Fecha', 'Actividades realizadas durante el MAP', 'Imagen de la actividad']
    for i, texto in enumerate(etiquetas):
        tabla_final.rows[0].cells[i].text = texto

    # Llenado de filas (sin agrupar, una por ficha encontrada)
    for ficha in datos_acumulados:
        row = tabla_final.add_row().cells
        row[0].text = ficha["fecha"]
        row[1].text = ficha["actividad"]
        
        # Insertar imágenes en la tercera columna
        parrafo_foto = row[2].paragraphs[0]
        for img_blob in ficha["imagenes"]:
            run = parrafo_foto.add_run()
            # Insertamos la imagen desde el flujo de bytes
            run.add_picture(io.BytesIO(img_blob), width=Inches(2.0))
            parrafo_foto.add_run("\n") # Espacio entre fotos si hay varias

    # Guardar resultado en memoria para descarga
    salida_memoria = io.BytesIO()
    dest_doc.save(salida_memoria)
    salida_memoria.seek(0)
    return salida_memoria

# --- INTERFAZ DE USUARIO (STREAMLIT) ---
st.set_page_config(page_title="ArqueoTab Gen", layout="wide")
st.title("📂 Generador de Tabla Resumen Arqueológica")
st.info("Sube tus archivos de Anexo (Word) y la app creará la tabla acumulada con fotos.")

archivos = st.file_uploader("Selecciona los archivos .docx", type="docx", accept_multiple_files=True)

if archivos:
    if st.button("🚀 Procesar y Generar Tabla"):
        todas_las_fichas = []
        for a in archivos:
            fichas_archivo = extraer_fichas_con_fotos(a.read())
            todas_las_fichas.extend(fichas_archivo)
        
        if todas_las_fichas:
            # Generamos el archivo
            archivo_final = crear_word_final(todas_las_fichas)
            
            st.success(f"Se procesaron {len(todas_las_fichas)} fichas correctamente.")
            st.download_button(
                label="⬇️ Descargar Tabla Resumen (.docx)",
                data=archivo_final,
                file_name="Tabla_Resumen_MAP_Acumulada.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("No se detectaron fichas válidas en los archivos subidos.")
    
    if st.button("Generar Word Final"):
        doc_binario = generar_word_resumen(lista_total)

        st.download_button("Descargar Tabla Resumen.docx", doc_binario, "Resumen_MAP.docx")

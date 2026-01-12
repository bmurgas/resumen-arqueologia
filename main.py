import streamlit as st
from docx import Document
from docx.shared import Inches
import io

def extraer_datos_fichas(archivo):
    doc = Document(archivo)
    fichas_del_archivo = []
    
    # Buscamos la estructura de la ficha
    # Nota: Cada documento puede tener varias fichas (una tras otra)
    # Identificamos el inicio por la celda "Fecha"
    for tabla in doc.tables:
        datos = {"fecha": "", "actividad": "", "fotos": []}
        capturando_fotos = False
        
        for fila in tabla.rows:
            texto_id = fila.cells[0].text.strip()
            
            if "Fecha" in texto_id:
                datos["fecha"] = fila.cells[1].text.strip()
            
            if "Descripción de la actividad" in texto_id:
                datos["actividad"] = fila.cells[1].text.strip()

            # Localizar sección de fotos
            if "Registro fotográfico" in texto_id:
                capturando_fotos = True
                continue
            
            if capturando_fotos:
                # Extraer imágenes de las celdas
                for celda in fila.cells:
                    for paragraph in celda.paragraphs:
                        for run in paragraph.runs:
                            if run._element.xpath('.//a:blip'):
                                # Guardamos el objeto de imagen
                                datos["fotos"].append(run)
        
        if datos["fecha"]:
            fichas_del_archivo.append(datos)
            
    return fichas_del_archivo

def generar_word_resumen(datos_totales):
    nuevo_doc = Document()
    nuevo_doc.add_heading('Tabla Resumen Monitoreo Arqueológico', 0)
    
    # Crear tabla de 3 columnas según tu formato
    tabla = nuevo_doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    hdr_cells = tabla.rows[0].cells
    hdr_cells[0].text = 'Fecha'
    hdr_cells[1].text = 'Actividades realizadas durante el MAP'
    hdr_cells[2].text = 'Imagen de la actividad'

    # Ordenar por fecha antes de escribir
    datos_ordenados = sorted(datos_totales, key=lambda x: x['fecha'])

    for item in datos_ordenados:
        row_cells = tabla.add_row().cells
        row_cells[0].text = item['fecha']
        row_cells[1].text = item['actividad']
        
        # Insertar fotos en la tercera columna
        paragraph = row_cells[2].paragraphs[0]
        for foto_run in item['fotos']:
            # Aquí la lógica para copiar la imagen de un doc a otro
            # Nota: python-docx requiere extraer el stream de la imagen original
            pass # (La lógica de copiado de imagen binaria se añade en la versión final)

    target = io.BytesIO()
    nuevo_doc.save(target)
    return target.getvalue()

# Interfaz Streamlit
st.title("Generador de Resumen Arqueológico")
uploaded_files = st.file_uploader("Sube los anexos diarios", type="docx", accept_multiple_files=True)

if uploaded_files:
    lista_total = []
    for f in uploaded_files:
        lista_total.extend(extraer_datos_fichas(f))
    
    if st.button("Generar Word Final"):
        doc_binario = generar_word_resumen(lista_total)
        st.download_button("Descargar Tabla Resumen.docx", doc_binario, "Resumen_MAP.docx")
import streamlit as st
import pandas as pd
import zipfile
import io
import xml.etree.ElementTree as ET
import re

def extraer_datos_kml(kml_content):
    """Lee el código XML de un archivo KML y extrae los Puntos."""
    xml_string = kml_content.decode('utf-8', errors='ignore')
    xml_string = re.sub(r'\sxmlns="[^"]+"', '', xml_string, count=1)
    
    try:
        root = ET.fromstring(xml_string)
    except Exception as e:
        return []

    datos = []
    for placemark in root.findall('.//Placemark'):
        nombre = placemark.find('name')
        nombre_txt = nombre.text if nombre is not None else "Sin nombre"

        punto = placemark.find('.//Point/coordinates')
        if punto is not None and punto.text:
            coords_texto = punto.text.strip()
            partes = [p.strip() for p in coords_texto.split(',')]
            
            lon = partes[0] if len(partes) > 0 else ""
            lat = partes[1] if len(partes) > 1 else ""
            alt = partes[2] if len(partes) > 2 else "0"

            datos.append({
                "Nombre del Punto": nombre_txt,
                "Latitud (Y)": lat,
                "Longitud (X)": lon,
                "Altura (Z)": alt
            })
    return datos

def mostrar_pagina():
    """Función principal que será llamada desde el main.py"""
    st.title("🗺️ Extractor de KMZ/KML a Excel")
    st.markdown("Sube tus archivos geográficos y obtén una planilla Excel con sus datos de forma instantánea.")
    
    archivos = st.file_uploader("Sube tus archivos (.kml o .kmz)", type=['kml', 'kmz'], accept_multiple_files=True, key="kmz_to_excel_up")

    if archivos and st.button("Extraer Datos a Excel"):
        todos_los_puntos = []
        
        with st.spinner("Procesando archivos geográficos..."):
            for archivo in archivos:
                nombre_archivo = archivo.name
                contenido = archivo.read()

                if nombre_archivo.lower().endswith('.kmz'):
                    try:
                        with zipfile.ZipFile(io.BytesIO(contenido)) as z:
                            kml_filename = next((name for name in z.namelist() if name.lower().endswith('.kml')), None)
                            if kml_filename:
                                kml_content = z.read(kml_filename)
                                puntos = extraer_datos_kml(kml_content)
                                for p in puntos: p["Archivo Origen"] = nombre_archivo
                                todos_los_puntos.extend(puntos)
                    except Exception as e:
                        st.error(f"No se pudo leer el archivo {nombre_archivo}: {e}")
                else:
                    puntos = extraer_datos_kml(contenido)
                    for p in puntos: p["Archivo Origen"] = nombre_archivo
                    todos_los_puntos.extend(puntos)

        if todos_los_puntos:
            df = pd.DataFrame(todos_los_puntos)
            columnas = ["Archivo Origen", "Nombre del Punto", "Latitud (Y)", "Longitud (X)", "Altura (Z)"]
            df = df[columnas]

            st.success(f"✅ ¡Éxito! Se extrajeron {len(df)} puntos geográficos.")
            st.dataframe(df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Coordenadas_Extraidas")

            st.download_button(
                label="⬇️ Descargar Planilla Excel",
                data=buffer.getvalue(),
                file_name="Coordenadas_Extraidas_KMZ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron Puntos en los archivos subidos.")

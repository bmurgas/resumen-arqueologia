import streamlit as st
import pandas as pd
import zipfile
import io
import xml.etree.ElementTree as ET
import re
from pyproj import Transformer # Importamos para la conversión a UTM

def extraer_datos_kml(kml_content):
    """
    Lee el código XML de un archivo KML, extrae los puntos
    y convierte sus coordenadas geográficas a UTM Huso 18S.
    """
    xml_string = kml_content.decode('utf-8', errors='ignore')
    xml_string = re.sub(r'\sxmlns="[^"]+"', '', xml_string, count=1)
    
    try:
        root = ET.fromstring(xml_string)
        # Configurar conversor: WGS84 (Lat/Lon) -> UTM Huso 18S (EPSG:32718)
        transformer = Transformer.from_crs("epsg:4326", "epsg:32718", always_xy=True)
    except Exception:
        return []

    datos = []
    
    for placemark in root.findall('.//Placemark'):
        nombre = placemark.find('name')
        nombre_txt = nombre.text if nombre is not None else "Sin nombre"

        punto = placemark.find('.//Point/coordinates')
        if punto is not None and punto.text:
            coords_texto = punto.text.strip()
            partes = [p.strip() for p in coords_texto.split(',')]
            
            lon_str = partes[0] if len(partes) > 0 else ""
            lat_str = partes[1] if len(partes) > 1 else ""
            alt = partes[2] if len(partes) > 2 else "0"

            # Si tenemos Latitud y Longitud válidas, calculamos el UTM
            utm_este = ""
            utm_norte = ""
            
            if lon_str and lat_str:
                try:
                    lon_float = float(lon_str)
                    lat_float = float(lat_str)
                    # Transformación matemática a metros (UTM)
                    este_float, norte_float = transformer.transform(lon_float, lat_float)
                    
                    # Formatear como números enteros o con 2 decimales para el Excel
                    utm_este = round(este_float, 2)
                    utm_norte = round(norte_float, 2)
                except:
                    pass

            datos.append({
                "Nombre del Punto": nombre_txt,
                "Latitud (Y)": lat_str,
                "Longitud (X)": lon_str,
                "UTM Este (X)": utm_este,
                "UTM Norte (Y)": utm_norte,
                "Altura (Z)": alt
            })

    return datos

def mostrar_pagina():
    """Función principal que es llamada desde el menú de main.py"""
    st.title("🗺️ Extractor de KMZ/KML a Excel (Con UTM)")
    st.markdown("Sube tus archivos geográficos para extraer sus datos en coordenadas Geográficas y **UTM (Huso 18S)**.")
    
    archivos = st.file_uploader("Sube tus archivos (.kml o .kmz)", type=['kml', 'kmz'], accept_multiple_files=True, key="kmz_to_excel_up")

    if archivos and st.button("Extraer Datos a Excel"):
        todos_los_puntos = []
        
        with st.spinner("Procesando archivos y calculando coordenadas UTM..."):
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
            
            # Ordenamos las columnas incluyendo los nuevos datos UTM
            columnas = [
                "Archivo Origen", 
                "Nombre del Punto", 
                "UTM Este (X)", 
                "UTM Norte (Y)", 
                "Latitud (Y)", 
                "Longitud (X)", 
                "Altura (Z)"
            ]
            df = df[columnas]

            st.success(f"✅ ¡Éxito! Se extrajeron {len(df)} puntos con sus respectivas coordenadas UTM.")
            st.dataframe(df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Coordenadas_UTM")

            st.download_button(
                label="⬇️ Descargar Planilla Excel",
                data=buffer.getvalue(),
                file_name="Coordenadas_Extraidas_UTM.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron puntos espaciales válidos en los archivos cargados.")

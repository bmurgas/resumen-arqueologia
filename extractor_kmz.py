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
    y convierte sus coordenadas geográficas a UTM Huso 19S (19K).
    """
    xml_string = kml_content.decode('utf-8', errors='ignore')
    xml_string = re.sub(r'\sxmlns="[^"]+"', '', xml_string, count=1)
    
    try:
        root = ET.fromstring(xml_string)
        # CONFIGURACIÓN: WGS84 (Lat/Lon) -> UTM Huso 19S / EPSG:32719 (Huso 19K)
        transformer = Transformer.from_crs("epsg:4326", "epsg:32719", always_xy=True)
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
                    # Transformación matemática a metros (UTM Huso 19)
                    este_float, norte_float = transformer.transform(lon_float, lat_float)
                    
                    # Redondeamos a 2 decimales para el Excel
                    utm_este = round(este_float, 2)
                    utm_norte = round(norte_float, 2)
                except:
                    pass

            datos.append({
                "Nombre del Punto": nombre_txt,
                "Latitud (Y)": lat_str,
                "Longitud (X)": lon_str,
                "UTM Este (X) - Huso 19": utm_este,
                "UTM Norte (Y) - Huso 19": utm_norte,
                "Altura (Z)": alt
            })

    return datos

def mostrar_pagina():
    """Función principal que es llamada desde el menú de main.py"""
    st.title("🗺️ Extractor de KMZ/KML a Excel (Huso 19K)")
    st.markdown("Sube tus archivos geográficos para extraer sus datos en coordenadas Geográficas y **UTM (Huso 19K)**.")
    
    archivos = st.file_uploader("Sube tus archivos (.kml o .kmz)", type=['kml', 'kmz'], accept_multiple_files=True, key="kmz_to_excel_up")

    if archivos and st.button("Extraer Datos a Excel"):
        todos_los_puntos = []
        
        with st.spinner("Procesando archivos y calculando coordenadas UTM Huso 19..."):
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
            
            # Ordenamos las columnas incluyendo los nuevos datos UTM Huso 19
            columnas = [
                "Archivo Origen", 
                "Nombre del Punto", 
                "UTM Este (X) - Huso 19", 
                "UTM Norte (Y) - Huso 19", 
                "Latitud (Y)", 
                "Longitud (X)", 
                "Altura (Z)"
            ]
            df = df[columnas]

            st.success(f"✅ ¡Éxito! Se extrajeron {len(df)} puntos con coordenadas UTM calculadas para el Huso 19.")
            st.dataframe(df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Coordenadas_UTM_19S")

            st.download_button(
                label="⬇️ Descargar Planilla Excel",
                data=buffer.getvalue(),
                file_name="Coordenadas_Extraidas_UTM_19.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron puntos espaciales válidos en los archivos cargados.")

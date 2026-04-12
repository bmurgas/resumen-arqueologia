import streamlit as st
import io
import pandas as pd
import zipfile
import re
import json 
from pyproj import Transformer
try:
    import fitz  # PyMuPDF
except ImportError:
    pass

# --- Funciones Auxiliares solo para este módulo ---
def limpiar_coordenada(texto):
    texto_limpio = str(texto).replace(".", "").replace(" ", "").strip()
    texto_limpio = texto_limpio.replace(",", ".")
    try:
        return float(texto_limpio)
    except:
        return None

def crear_kml_texto(puntos):
    kml_header = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Hallazgos Arqueológicos</name>"""
    kml_footer = """
  </Document>
</kml>"""
    kml_body = ""
    for p in puntos:
        kml_body += f"""
    <Placemark>
      <name>{p['nombre']}</name>
      <description>{p['desc']}</description>
      <Point>
        <coordinates>{p['lon']},{p['lat']},0</coordinates>
      </Point>
    </Placemark>"""
    return kml_header + kml_body + kml_footer

# --- LA LÓGICA CORRECTA DE EXTRACCIÓN (Línea por línea + Saltos) ---
def procesar_pdf_recoleccion_regex_gis(pdf_bytes, nombre_archivo):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        st.error(f"Error abriendo PDF {nombre_archivo}: {e}")
        return []

    fichas = []

    for pagina in doc:
        texto_completo = pagina.get_text("text")
        # Leemos línea por línea
        lineas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
        
        if len(lineas) < 10:
            continue

        ficha = {
            "Responsable": "", "Sitio": "", "Hallazgo Previsto": "",
            "Cuadrante": "", "Dimensión": "", "Fecha": "",
            "UTM Norte": "", "UTM Este": "", "Material": "", "Superficie": ""
        }

        # Bucle de lectura secuencial que armamos ayer
        for i, linea in enumerate(lineas):
            lin_lower = linea.lower().replace(":", "").strip()

            if lin_lower == "responsable":
                if i + 1 < len(lineas): ficha["Responsable"] = lineas[i+1]
            
            elif lin_lower == "sitio":
                if i + 1 < len(lineas): ficha["Sitio"] = lineas[i+1]
            
            elif lin_lower == "hallazgo previsto":
                pass # Se captura con Regex abajo por seguridad
            
            elif lin_lower == "cuadrante":
                if i + 1 < len(lineas): ficha["Cuadrante"] = lineas[i+1]
            
            elif lin_lower in ["dimensión", "dimension"]:
                if i + 1 < len(lineas): ficha["Dimensión"] = lineas[i+1]
            
            elif lin_lower == "fecha":
                if i + 1 < len(lineas): ficha["Fecha"] = lineas[i+1]
            
            elif lin_lower == "material":
                if i + 1 < len(lineas):
                    # El famoso salto de columna si viene la palabra superficie
                    if lineas[i+1].lower().replace(":", "").strip() == "superficie":
                        if i + 2 < len(lineas): ficha["Material"] = lineas[i+2]
                    else:
                        ficha["Material"] = lineas[i+1]
            
            elif lin_lower == "superficie":
                if i + 1 < len(lineas):
                    # El salto inverso
                    if i - 1 >= 0 and lineas[i-1].lower().replace(":", "").strip() == "material":
                        if i + 2 < len(lineas): ficha["Superficie"] = lineas[i+2]
                    else:
                        ficha["Superficie"] = lineas[i+1]
            
            elif lin_lower.startswith("utm norte"):
                val = linea[len("UTM Norte"):].strip()
                val = re.sub(r'^[:\-\s]+', '', val)
                if val:
                    ficha["UTM Norte"] = val
                elif i + 1 < len(lineas):
                    ficha["UTM Norte"] = lineas[i+1]
                    
            elif lin_lower.startswith("utm este"):
                val = linea[len("UTM Este"):].strip()
                val = re.sub(r'^[:\-\s]+', '', val)
                if val:
                    ficha["UTM Este"] = val
                elif i + 1 < len(lineas):
                    ficha["UTM Este"] = lineas[i+1]

        # FASE DE LIMPIEZA
        etiquetas_conocidas = ["sitio", "responsable", "cuadrante", "dimensión", "dimension", "fecha", "material", "superficie", "coordenadas", "identificación", "procedencia y material cultural"]
        for key in list(ficha.keys()):
            val_limpio = str(ficha[key]).lower().strip()
            if val_limpio in etiquetas_conocidas or val_limpio == key.lower():
                ficha[key] = ""

        # RESPALDOS DE SEGURIDAD (REGEX)
        if not ficha["Fecha"]:
            m = re.search(r"(\d{2}/\d{2}/\d{4})", texto_completo)
            if m: ficha["Fecha"] = m.group(1)

        if not ficha["Hallazgo Previsto"] or len(ficha["Hallazgo Previsto"]) < 4:
            m = re.search(r"(HLU_HP_\d+|HP_\d+)", texto_completo)
            if m: ficha["Hallazgo Previsto"] = m.group(1)

        # Solo guardar si hay datos clave
        if ficha["Sitio"] or ficha["Responsable"]:
            fichas.append(ficha)

    return fichas

# --- La Interfaz Visual de este módulo ---
def ejecutar_interfaz():
    st.title("Generador Base de Datos y GIS (Módulo Actualizado)")
    st.markdown("Extrae datos mediante patrones lógicos secuenciales y convierte coordenadas UTM para QGIS y Google Earth.")
    
    archivos = st.file_uploader("Subir Fichas PDF (.pdf)", accept_multiple_files=True, key="pdf_recoleccion_up_nuevo")
    if archivos and st.button("Procesar Fichas y Crear Mapas"):
        todas_las_fichas = []
        bar = st.progress(0)
        for i, a in enumerate(archivos):
            fichas_extraidas = procesar_pdf_recoleccion_regex_gis(a.read(), a.name)
            todas_las_fichas.extend(fichas_extraidas)
            bar.progress((i+1)/len(archivos))
            
        if todas_las_fichas:
            columnas_ordenadas = [
                "Responsable", "Sitio", "Hallazgo Previsto", "Cuadrante", 
                "Dimensión", "Fecha", "UTM Norte", "UTM Este", "Material", "Superficie"
            ]
            df = pd.DataFrame(todas_las_fichas)[columnas_ordenadas]
            st.success(f"✅ Se extrajeron {len(df)} registros correctamente.")
            st.dataframe(df)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Hallazgos Previstos")
            
            kmz_buffer = None
            geojson_str = None
            try:
                transformer = Transformer.from_crs("epsg:32718", "epsg:4326", always_xy=True)
                puntos_kml = []
                features_geojson = []
                
                for f in todas_las_fichas:
                    n_val = limpiar_coordenada(f.get("UTM Norte", ""))
                    e_val = limpiar_coordenada(f.get("UTM Este", ""))
                    
                    if n_val and e_val:
                        lon, lat = transformer.transform(e_val, n_val)
                        nombre = f.get("Hallazgo Previsto", f.get("Sitio", "Sin ID"))
                        desc = f"Material: {f.get('Material', '')} | Superficie: {f.get('Superficie', '')} | Fecha: {f.get('Fecha', '')}"
                        
                        puntos_kml.append({"nombre": nombre, "desc": desc, "lat": lat, "lon": lon})
                        
                        features_geojson.append({
                            "type": "Feature",
                            "properties": {
                                "ID_Hallazgo": nombre,
                                "Sitio": f.get("Sitio", ""),
                                "Cuadrante": f.get("Cuadrante", ""),
                                "Material": f.get("Material", ""),
                                "Superficie": f.get("Superficie", ""),
                                "Fecha": f.get("Fecha", "")
                            },
                            "geometry": {
                                "type": "Point",
                                "coordinates": [lon, lat]
                            }
                        })
                        
                if puntos_kml:
                    kml_content = crear_kml_texto(puntos_kml)
                    kmz_buffer = io.BytesIO()
                    with zipfile.ZipFile(kmz_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        zf.writestr("doc.kml", kml_content)
                    kmz_buffer.seek(0)
                    
                    geojson_data = {
                        "type": "FeatureCollection",
                        "features": features_geojson
                    }
                    geojson_str = json.dumps(geojson_data)
            except Exception as e:
                st.warning("⚠️ No se pudieron procesar los archivos espaciales. Asegúrate de tener la librería 'pyproj' instalada.")

            st.markdown("### Descargas Disponibles")
            col1, col2, col3 = st.columns(3)
            
            col1.download_button(
                label="📊 Descargar Excel", 
                data=buffer.getvalue(), 
                file_name="Base_Datos_Recoleccion_Superficial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            if kmz_buffer:
                col2.download_button(
                    label="🌍 Descargar KMZ", 
                    data=kmz_buffer.getvalue(), 
                    file_name="Geometrias_Recoleccion.kmz",
                    mime="application/vnd.google-earth.kmz"
                )
            
            if geojson_str:
                col3.download_button(
                    label="🗺️ Descargar GeoJSON", 
                    data=geojson_str, 
                    file_name="Geometrias_Recoleccion.geojson",
                    mime="application/geo+json"
                )
        else:
            st.error("No se encontraron datos de recolección válidos en los PDFs subidos.")

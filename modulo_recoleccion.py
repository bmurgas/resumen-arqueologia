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

def procesar_pdf_recoleccion_regex_gis(pdf_bytes, nombre_archivo):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        st.error(f"Error abriendo PDF {nombre_archivo}: {e}")
        return []

    fichas = []
    for pagina in doc:
        texto = pagina.get_text("text")
        
        if "Sitio" not in texto and "Responsable" not in texto:
            continue

        ficha = {
            "Responsable": "", "Sitio": "", "Hallazgo Previsto": "",
            "Cuadrante": "", "Dimensión": "", "Fecha": "",
            "UTM Norte": "", "UTM Este": "", "Material": "", "Superficie": ""
        }

        m = re.search(r"Sitio\s*\n+([^\n]+)", texto)
        if m: ficha["Sitio"] = m.group(1).strip()
        
        m = re.search(r"Cuadrante\s*\n+([^\n]+)", texto)
        if m: ficha["Cuadrante"] = m.group(1).strip()

        m = re.search(r"Responsable\s*\n+([^\n]+)", texto)
        if m: ficha["Responsable"] = m.group(1).strip()

        m = re.search(r"([^\n]+)\s*\n+Hallazgo Previsto", texto)
        if m: 
            ficha["Hallazgo Previsto"] = m.group(1).strip()
        else:
            m = re.search(r"(HLU_HP_\d+|HP_\d+)", texto)
            if m: ficha["Hallazgo Previsto"] = m.group(1)

        m = re.search(r"Dimensi[oó]n\s*\n+([^\n]+)", texto, re.IGNORECASE)
        if m: ficha["Dimensión"] = m.group(1).strip()

        m = re.search(r"Fecha\s*\n+([^\n]+)", texto)
        if m: 
            ficha["Fecha"] = m.group(1).strip()
        if not ficha["Fecha"] or not re.search(r"\d", ficha["Fecha"]):
            m = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
            if m: ficha["Fecha"] = m.group(1)

        m = re.search(r"UTM Norte\s*(?:\n\s*)?(\d+)", texto)
        if m: ficha["UTM Norte"] = m.group(1).strip()

        m = re.search(r"UTM Este\s*(?:\n\s*)?(\d+)", texto)
        if m: ficha["UTM Este"] = m.group(1).strip()

        m = re.search(r"Material\s*\n+([^\n]+)", texto)
        if m: ficha["Material"] = m.group(1).strip()

        m = re.search(r"Superficie\s*\n+([^\n]+)", texto)
        if m: ficha["Superficie"] = m.group(1).strip()

        etiquetas_conocidas = [
            "Sitio", "Cuadrante", "Responsable", "Hallazgo Previsto", 
            "Dimensión", "Dimension", "Fecha", "UTM Norte", "UTM Este", 
            "Material", "Superficie", "Altura", "Descripción", "Cronología", 
            "COORDENADAS", "IDENTIFICACIÓN", "PROCEDENCIA Y MATERIAL CULTURAL"
        ]
        
        for k, v in ficha.items():
            for et in etiquetas_conocidas:
                if v.lower() == et.lower():
                    ficha[k] = ""
                    break

        if ficha["Sitio"] or ficha["Responsable"]:
            fichas.append(ficha)

    return fichas

def ejecutar_interfaz():
    st.title("Generador Base de Datos y GIS (Recolección Superficial)")
    st.markdown("Extrae datos mediante patrones lógicos (Regex) y convierte coordenadas UTM para QGIS y Google Earth.")
    
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

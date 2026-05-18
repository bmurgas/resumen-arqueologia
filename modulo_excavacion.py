import streamlit as st
import pandas as pd
import io
import re
try:
    import fitz  # PyMuPDF
except ImportError:
    pass

def extraer_datos_excavacion(pdf_bytes, nombre_archivo):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        st.error(f"Error abriendo PDF {nombre_archivo}: {e}")
        return None

    # Diccionario base con la estructura exacta que pide el Excel
    ficha = {
        "Sitio": "", "Unidad": "", "C. Norte": "", "C. Este": "", "Dimensión": "", "Fecha": "", "Responsable": "",
        # Nivel Superficial
        "Capa_Sup": "", "Litico_Sup": "", "Osteofauna_Sup": "", "Malacologico_Sup": "", "Vidrio_Sup": "", "Metal_Sup": "", "Ceramica_Sup": "", "Otros_Sup": "",
        # Nivel I
        "Capa_I": "", "Litico_I": "", "Osteofauna_I": "", "Malacologico_I": "", "Vidrio_I": "", "Metal_I": "", "Ceramica_I": "", "Otros_I": "",
        # Nivel II
        "Capa_II": "", "Litico_II": "", "Osteofauna_II": "", "Malacologico_II": "", "Vidrio_II": "", "Metal_II": "", "Ceramica_II": "", "Otros_II": "",
        # Nivel III
        "Capa_III": "", "Litico_III": "", "Osteofauna_III": "", "Malacologico_III": "", "Vidrio_III": "", "Metal_III": "", "Ceramica_III": "", "Otros_III": "",
        # Nivel IV
        "Capa_IV": "", "Litico_IV": "", "Osteofauna_IV": "", "Malacologico_IV": "", "Vidrio_IV": "", "Metal_IV": "", "Ceramica_IV": "", "Otros_IV": "",
        # Nivel V
        "Capa_V": "", "Litico_V": "", "Osteofauna_V": "", "Malacologico_V": "", "Vidrio_V": "", "Metal_V": "", "Ceramica_V": "", "Otros_V": "",
        # Observaciones
        "Obs_Sup": "", "Obs_I": "", "Obs_II": "", "Obs_III": "", "Obs_IV": "", "Obs_V": ""
    }

    texto_completo = ""
    for pagina in doc:
        texto_completo += pagina.get_text("text") + "\n"

    lineas = [l.strip() for l in texto_completo.split('\n') if l.strip()]

    # 1. Extracción de los datos de Cabecera (Identificación)
    for i, linea in enumerate(lineas):
        # Buscamos la fila que tiene los títulos y tomamos los valores de la fila de abajo
        if "Sitio" in linea and "Unidad" in lineas[i+1] if i+1 < len(lineas) else False:
            # En tu PDF los valores están unas 7 líneas más abajo porque primero lee todos los títulos
            # Buscaremos los valores iterando desde donde estamos
            try:
                # Utilizamos una búsqueda por Regex para la cabecera por mayor seguridad
                m_sitio = re.search(r"(HLU-\d+|Sitio\s*([A-Za-z0-9\-]+))", texto_completo)
                if m_sitio: ficha["Sitio"] = m_sitio.group(1).replace("Sitio", "").strip()
                
                m_unidad = re.search(r"(HLU-HP-\d+|Unidad\s*([A-Za-z0-9\-]+))", texto_completo)
                if m_unidad: ficha["Unidad"] = m_unidad.group(1).replace("Unidad", "").strip()

                m_norte = re.search(r"C\. Norte\s*\n+(\d+)", texto_completo)
                if m_norte: ficha["C. Norte"] = m_norte.group(1)

                m_este = re.search(r"C\. Este\s*\n+(\d+)", texto_completo)
                if m_este: ficha["C. Este"] = m_este.group(1)

                m_dim = re.search(r"(\d+\s*[mM]\s*[xX]\s*\d+\s*[mM])", texto_completo)
                if m_dim: ficha["Dimensión"] = m_dim.group(1)

                m_fecha = re.search(r"(\d{2}[-/]\d{2}[-/]\d{4})", texto_completo)
                if m_fecha: ficha["Fecha"] = m_fecha.group(1)

                # Responsable suele estar justo después de la fecha en la cabecera
                for j, l in enumerate(lineas):
                    if l == ficha["Fecha"] and j + 1 < len(lineas):
                        if not re.match(r"^\d", lineas[j+1]): # Si no es un número, es el nombre
                            ficha["Responsable"] = lineas[j+1]
                            break
            except:
                pass

    # 2. Extracción de la Tabla de Materiales por Nivel
    # Mapeo de sufijos según el nivel detectado en la línea
    niveles_map = {
        "superficial": "_Sup",
        "0-10": "_I",
        "10-20": "_II",
        "20-30": "_III",
        "30-40": "_IV",
        "40-50": "_V"
    }

    for i, linea in enumerate(lineas):
        linea_lower = linea.lower()
        sufijo_actual = None
        
        # Identificar en qué nivel estamos iterando
        for clave, sufijo in niveles_map.items():
            if clave in linea_lower:
                sufijo_actual = sufijo
                break
        
        # Si detectamos un nivel, tomamos los siguientes 8 valores (Capa, Lítico, Osteo... Otros)
        if sufijo_actual and (i + 8) < len(lineas):
            # A veces el lector de PDF divide mal. Verificamos que no estemos leyendo otro nivel
            if "disturbada" in lineas[i+1].lower() or len(lineas[i+1]) < 15:
                ficha[f"Capa{sufijo_actual}"] = lineas[i+1]
                ficha[f"Litico{sufijo_actual}"] = lineas[i+2]
                ficha[f"Osteofauna{sufijo_actual}"] = lineas[i+3]
                ficha[f"Malacologico{sufijo_actual}"] = lineas[i+4]
                ficha[f"Vidrio{sufijo_actual}"] = lineas[i+5]
                ficha[f"Metal{sufijo_actual}"] = lineas[i+6]
                ficha[f"Ceramica{sufijo_actual}"] = lineas[i+7]
                ficha[f"Otros{sufijo_actual}"] = lineas[i+8]

    # 3. Extracción de Observaciones (Asumiendo que vienen debajo como "Observaciones Nivel X")
    for i, linea in enumerate(lineas):
        linea_lower = linea.lower()
        if "observaci" in linea_lower:
            # Aquí podrías capturar la línea siguiente dependiendo de cómo se estructure el PDF.
            # Se ha dejado preparado el contenedor.
            pass 

    return ficha

def ejecutar_interfaz():
    st.title("Generador Excel (Fichas de Excavación)")
    st.markdown("Extrae los datos de la matriz de excavación (materiales por niveles) y genera el Excel en formato extendido horizontal.")
    
    archivos = st.file_uploader("Subir Fichas de Excavación PDF (.pdf)", accept_multiple_files=True, key="pdf_excavacion_up")
    
    if archivos and st.button("Procesar Fichas de Excavación"):
        datos_extraidos = []
        bar = st.progress(0)
        
        for i, a in enumerate(archivos):
            ficha = extraer_datos_excavacion(a.read(), a.name)
            if ficha:
                datos_extraidos.append(ficha)
            bar.progress((i+1)/len(archivos))
            
        if datos_extraidos:
            # Convertimos a DataFrame
            df = pd.DataFrame(datos_extraidos)
            
            # Formatear el Excel con la estructura exacta (dos filas de encabezado)
            # Primero, creamos una lista con los encabezados reales (fila 2 del Excel)
            columnas_finales = [
                "Sitio", "Unidad", "C. Norte", "C. Este", "Dimensión", "Fecha", "Responsable",
                "Capa", "Litico", "Osteofauna", "Malacológico", "Vidrio", "Metal", "Cerámica", "Otros", # Sup
                "Capa", "Litico", "Osteofauna", "Malacológico", "Vidrio", "Metal", "Cerámica", "Otros", # I
                "Capa", "Litico", "Osteofauna", "Malacológico", "Vidrio", "Metal", "Cerámica", "Otros", # II
                "Capa", "Litico", "Osteofauna", "Malacológico", "Vidrio", "Metal", "Cerámica", "Otros", # III
                "Capa", "Litico", "Osteofauna", "Malacológico", "Vidrio", "Metal", "Cerámica", "Otros", # IV
                "Capa", "Litico", "Osteofauna", "Malacológico", "Vidrio", "Metal", "Cerámica", "Otros", # V
                "Observacion nivel Superficial:", "Observacion nivel I (0-10 cm):", "Observacion nivel II (10-20 cm):", 
                "Observacion nivel III (20-30 cm):", "Observacion nivel IV (30-40 cm):", "Observacion nivel V (40-50 cm):"
            ]
            
            # Cabecera superior (fila 1 del Excel) con los niveles agrupados
            header_niveles = (
                [""] * 7 + 
                ["Superficial"] + [""] * 7 + 
                ["I (0-10 cm)"] + [""] * 7 +
                ["II (10-20 cm)"] + [""] * 7 +
                ["III (20-30 cm)"] + [""] * 7 +
                ["IV (30-40 cm)"] + [""] * 7 +
                ["V (40-50 cm)"] + [""] * 7 +
                [""] * 6
            )

            df.columns = columnas_finales
            
            # Crear un nuevo dataframe para inyectar la fila superior
            df_header = pd.DataFrame([columnas_finales], columns=header_niveles)
            df.columns = header_niveles
            df_final = pd.concat([df_header, df], ignore_index=True)

            st.success(f"✅ Se procesaron {len(datos_extraidos)} fichas de excavación.")
            st.dataframe(df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Hoja1")
            
            st.download_button(
                label="📊 Descargar Excel de Excavación", 
                data=buffer.getvalue(), 
                file_name="Base_Datos_Excavacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se pudieron extraer datos de los archivos proporcionados.")

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

    # Diccionario base con sufijos únicos
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
    try:
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

        for j, l in enumerate(lineas):
            if l == ficha["Fecha"] and j + 1 < len(lineas):
                if not re.match(r"^\d", lineas[j+1]): 
                    ficha["Responsable"] = lineas[j+1]
                    break
    except:
        pass

    # 2. Extracción de la Tabla de Materiales por Nivel (CANDADO ANTI-ERRORES)
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
        
        # EL CANDADO: Si la línea menciona observaciones o fotos, lo saltamos rotundamente
        if "observaci" in linea_lower or "registro" in linea_lower or "foto" in linea_lower:
            continue

        sufijo_actual = None
        for clave, sufijo in niveles_map.items():
            if clave in linea_lower:
                sufijo_actual = sufijo
                break
        
        if sufijo_actual and (i + 8) < len(lineas):
            # SEGUNDO CANDADO: Solo guardamos si esa capa está vacía (evita que se sobreescriba)
            if not ficha[f"Capa{sufijo_actual}"]:
                
                # Para estar 100% seguros, revisamos que haya números en los materiales
                materiales = [lineas[i+2], lineas[i+3], lineas[i+4], lineas[i+5], lineas[i+6], lineas[i+7], lineas[i+8]]
                numeros_encontrados = sum(1 for m in materiales if m.isdigit() or m == "0")
                
                if numeros_encontrados >= 3 or len(lineas[i+1]) <= 15:
                    ficha[f"Capa{sufijo_actual}"] = lineas[i+1]
                    ficha[f"Litico{sufijo_actual}"] = lineas[i+2]
                    ficha[f"Osteofauna{sufijo_actual}"] = lineas[i+3]
                    ficha[f"Malacologico{sufijo_actual}"] = lineas[i+4]
                    ficha[f"Vidrio{sufijo_actual}"] = lineas[i+5]
                    ficha[f"Metal{sufijo_actual}"] = lineas[i+6]
                    ficha[f"Ceramica{sufijo_actual}"] = lineas[i+7]
                    ficha[f"Otros{sufijo_actual}"] = lineas[i+8]

    # 3. Extracción de Observaciones Reales
    for i, linea in enumerate(lineas):
        linea_lower = linea.lower()
        if "observaci" in linea_lower:
            sufijo_obs = None
            if "superficial" in linea_lower: sufijo_obs = "_Sup"
            elif "0-10" in linea_lower or " i " in linea_lower or " 1 " in linea_lower: sufijo_obs = "_I"
            elif "10-20" in linea_lower or " ii " in linea_lower or " 2 " in linea_lower: sufijo_obs = "_II"
            elif "20-30" in linea_lower or " iii " in linea_lower or " 3 " in linea_lower: sufijo_obs = "_III"
            elif "30-40" in linea_lower or " iv " in linea_lower or " 4 " in linea_lower: sufijo_obs = "_IV"
            elif "40-50" in linea_lower or " v " in linea_lower or " 5 " in linea_lower: sufijo_obs = "_V"

            if sufijo_obs:
                obs_texto = ""
                # Si hay texto después de los dos puntos
                if ":" in linea:
                    obs_texto = linea.split(":", 1)[1].strip()
                
                # Si no había texto al lado, tomamos la línea de abajo
                if not obs_texto and i + 1 < len(lineas):
                    if "registro" not in lineas[i+1].lower() and "observaci" not in lineas[i+1].lower():
                        obs_texto = lineas[i+1].strip()
                
                ficha[f"Obs{sufijo_obs}"] = obs_texto

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
            df = pd.DataFrame(datos_extraidos)
            
            # 1. MOSTRAR EN PANTALLA
            st.success(f"✅ Se procesaron {len(datos_extraidos)} fichas de excavación.")
            st.dataframe(df)

            # 2. GENERAR EXCEL
            fila1 = (
                ["", "", "", "", "", "", ""] + 
                ["Superficial"] + [""] * 7 + 
                ["I (0-10 cm)"] + [""] * 7 +
                ["II (10-20 cm)"] + [""] * 7 +
                ["III (20-30 cm)"] + [""] * 7 +
                ["IV (30-40 cm)"] + [""] * 7 +
                ["V (40-50 cm)"] + [""] * 7 +
                [""] * 6
            )
            
            fila2 = [
                "Sitio", "Unidad", "C. Norte", "C. Este", "Dimensión", "Fecha", "Responsable"
            ] + ["Capa", "Litico", "Osteofauna", "Malacológico", "Vidrio", "Metal", "Cerámica", "Otros"] * 6 + [
                "Observacion nivel Superficial:", "Observacion nivel I (0-10 cm):", "Observacion nivel II (10-20 cm):", 
                "Observacion nivel III (20-30 cm):", "Observacion nivel IV (30-40 cm):", "Observacion nivel V (40-50 cm):"
            ]

            datos_excel = [fila1, fila2] + df.values.tolist()
            df_export = pd.DataFrame(datos_excel)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, header=False, sheet_name="Hoja1")
            
            st.download_button(
                label="📊 Descargar Excel de Excavación", 
                data=buffer.getvalue(), 
                file_name="Base_Datos_Excavacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se pudieron extraer datos de los archivos proporcionados.")

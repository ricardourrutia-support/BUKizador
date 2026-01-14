import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import difflib

# --- CONFIGURACIÓN MINIMALISTA ---
st.set_page_config(
    page_title="BUKizador",
    page_icon="🤖",
    layout="centered", # Centrado para look más limpio/app móvil
    initial_sidebar_state="collapsed"
)

# CSS para esconder elementos innecesarios y dar look limpio
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stApp {background-color: #FAFAFA;}
    h1 {color: #2C3E50; font-family: 'Helvetica', sans-serif; font-weight: 700;}
    .stButton>button {
        background-color: #2C3E50; color: white; border-radius: 8px; 
        border: none; padding: 10px 24px; font-weight: 600;
    }
    .stButton>button:hover {background-color: #34495E; color: white;}
    </style>
""", unsafe_allow_html=True)

st.title("🤖 BUKizador")
st.caption("Transformación e Inyección de Turnos Inteligente")

# --- LÓGICA CORE (MOTOR) ---

def limpiar_texto(texto):
    if pd.isna(texto): return ""
    texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.upper().strip()

def buscar_rut_inteligente(nombres_unicos, df_colaboradores):
    mapa_ruts = {}
    df_colab = df_colaboradores.copy()
    df_colab['Nombre_Clean'] = df_colab['Nombre del Colaborador'].apply(limpiar_texto)
    rut_lookup = df_colab.set_index('Nombre_Clean')['RUT'].to_dict()
    lista_nombres = df_colab['Nombre_Clean'].unique()

    for nombre in nombres_unicos:
        if not nombre or pd.isna(nombre): continue
        n_clean = limpiar_texto(nombre)
        partes = n_clean.split()
        
        # 1. Coincidencia Exacta de Partes
        matches = [real for real in lista_nombres if all(p in real for p in partes)]
        
        if len(matches) == 1:
            mapa_ruts[nombre] = rut_lookup[matches[0]]
        elif len(matches) > 1:
            mapa_ruts[nombre] = "ERROR: Múltiples coincidencias"
        else:
            # 2. Coincidencia Difusa (Fuzzy)
            posibles = difflib.get_close_matches(n_clean, lista_nombres, n=1, cutoff=0.7)
            if posibles:
                mapa_ruts[nombre] = rut_lookup[posibles[0]]
            else:
                mapa_ruts[nombre] = "ERROR: No encontrado"
    return mapa_ruts

def normalizar_horarios(serie):
    s = serie.astype(str).str.upper().str.strip()
    res = pd.Series(index=s.index, dtype='object')
    
    # Detección de Libres
    mask_libre = s.str.contains('LIBRE', na=False) | s.isna() | (s == 'NAN')
    res[mask_libre] = 'L'
    
    # Regex HH:MM
    patron = r"(\d{1,2}):(\d{2})"
    s_proc = s[~mask_libre]
    
    if s_proc.empty: return res
    
    extracted = s_proc.str.findall(patron)
    
    def formatear(match):
        if not isinstance(match, list) or len(match) < 2: return "ERROR_FORMATO"
        return f"{int(match[0][0]):02d}:{match[0][1]}-{int(match[-1][0]):02d}:{match[-1][1]}"

    res[~mask_libre] = extracted.apply(formatear)
    return res

def cargar_plantilla_robusta(archivo):
    """Intenta cargar la plantilla ya sea como Excel real o CSV disfrazado."""
    try:
        return pd.read_excel(archivo)
    except:
        archivo.seek(0)
        try:
            # Intentar como CSV con punto y coma
            return pd.read_csv(archivo, sep=';', engine='python')
        except:
            archivo.seek(0)
            # Intentar como CSV con coma
            return pd.read_csv(archivo, sep=',', engine='python')

# --- INTERFAZ DE USUARIO ---

with st.container():
    col1, col2 = st.columns(2)
    archivo_input = col1.file_uploader("1. Excel Supervisores", type=["xlsx"])
    archivo_plantilla = col2.file_uploader("2. Plantilla BUK", type=["xls", "xlsx", "csv"])

if archivo_input and archivo_plantilla:
    with st.spinner('BUKizando datos...'):
        try:
            # 1. CARGA Y PROCESAMIENTO
            xls = pd.ExcelFile(archivo_input)
            df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=2)
            df_base = pd.read_excel(xls, sheet_name='Base de Colaboradores')
            df_cods = pd.read_excel(xls, sheet_name='Codificación de Turnos')

            # Melt
            col_nom = df_turnos.columns[0]
            cols_fechas = [c for c in df_turnos.columns if c != col_nom]
            df_long = df_turnos.melt(id_vars=[col_nom], value_vars=cols_fechas, var_name='Fecha', value_name='Turno_Raw')
            df_long = df_long.dropna(subset=[col_nom])

            # Logic
            mapa = buscar_rut_inteligente(df_long[col_nom].unique(), df_base)
            df_long['RUT'] = df_long[col_nom].map(mapa)
            
            # Normalización
            df_long['Turno_Norm'] = normalizar_horarios(df_long['Turno_Raw'])
            df_cods['Horario_Norm'] = normalizar_horarios(df_cods['Horario'])
            
            dic_turnos = dict(zip(df_cods['Horario_Norm'], df_cods['Sigla']))
            dic_turnos['L'] = 'L'
            
            df_long['Sigla'] = df_long['Turno_Norm'].map(dic_turnos)
            
            # Pivot (Datos listos para inyectar)
            df_pivot = df_long.pivot(index='RUT', columns='Fecha', values='Sigla')

            # 2. INYECCIÓN EN PLANTILLA
            df_template = cargar_plantilla_robusta(archivo_plantilla)
            cols_template = df_template.columns.tolist()
            
            filas_nuevas = []
            ruts_validos = [r for r in df_long['RUT'].unique() if "ERROR" not in str(r)]

            # Crear datos para cada colaborador
            for rut in ruts_validos:
                fila = {}
                # Buscar info extra del colaborador
                info_colab = df_base[df_base['RUT'] == rut].iloc[0] if not df_base[df_base['RUT'] == rut].empty else {}
                
                for col in cols_template:
                    col_u = col.upper()
                    
                    # Estrategias de llenado
                    if 'RUT' in col_u or 'EMPLEADO' in col_u:
                        fila[col] = rut
                    elif col in df_pivot.columns: # Es una fecha (ej: 2026-01-01)
                        val = df_pivot.loc[rut, col]
                        fila[col] = val if pd.notna(val) else ""
                    elif 'NOMBRE' in col_u and not info_colab.empty:
                        fila[col] = info_colab.get('Nombre del Colaborador', '')
                    elif 'AREA' in col_u and not info_colab.empty:
                        fila[col] = info_colab.get('Área', '')
                    elif 'SUPERVISOR' in col_u and not info_colab.empty:
                        fila[col] = info_colab.get('Supervisor', '')
                    else:
                        fila[col] = "" # Mantener estructura vacía si no sabemos qué es
                
                filas_nuevas.append(fila)
            
            df_final = pd.DataFrame(filas_nuevas)
            # Reordenar columnas estrictamente como la plantilla
            df_final = df_final[cols_template]

            # 3. RESULTADO
            st.success(f"✅ ¡Listo! {len(df_final)} colaboradores procesados.")
            
            # Descarga CSV (Formato universal para BUK)
            csv_buffer = df_final.to_csv(index=False, sep=';', encoding='utf-8-sig')
            
            st.download_button(
                label="📥 Descargar BUKizador Output (.csv)",
                data=csv_buffer,
                file_name="Carga_Masiva_BUK_Final.csv",
                mime="text/csv",
                use_container_width=True
            )
            
            # Muestra de errores si existen
            errores = df_long[df_long['Sigla'].isna()]['Turno_Norm'].unique()
            if len(errores) > 0:
                with st.expander("⚠️ Ver turnos no reconocidos"):
                    st.write("Estos horarios no estaban en el maestro de codificación:")
                    st.write(errores)

        except Exception as e:
            st.error("Hubo un problema procesando los archivos.")
            st.code(f"Error: {e}")
            st.info("Tip: Revisa que el Input tenga las 3 hojas correctas y la Plantilla sea legible.")

else:
    # Estado vacío (Empty State) bonito
    st.markdown("""
    <div style="text-align: center; color: #95a5a6; padding: 50px;">
        <h3>Esperando archivos...</h3>
        <p>Sube el Excel de Supervisores y tu Plantilla BUK para comenzar la magia.</p>
    </div>
    """, unsafe_allow_html=True)

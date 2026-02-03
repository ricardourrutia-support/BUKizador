import streamlit as st
import pandas as pd
import io
import unicodedata
import difflib

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="BUKizador Interactivo", page_icon="🤖", layout="centered")

# Estilos
st.markdown("""
    <style>
    .stApp {background-color: #FAFAFA;}
    h1 {color: #2C3E50;}
    .stButton>button {width: 100%; border-radius: 8px;}
    .reportview-container .main .block-container{padding-top: 2rem;}
    /* Resaltar caja de corrección */
    .stForm {background-color: #ffffff; padding: 20px; border-radius: 10px; border: 1px solid #ddd; box-shadow: 0 2px 5px rgba(0,0,0,0.05);}
    </style>
""", unsafe_allow_html=True)

st.title("🤖 BUKizador Interactivo")
st.caption("Fase 1: Análisis | Fase 2: Corrección Humana | Fase 3: Descarga")

# --- ESTADO DE LA SESIÓN (MEMORIA) ---
if 'etapa' not in st.session_state:
    st.session_state.etapa = 'carga' # carga -> correccion -> descarga
if 'mapa_manual' not in st.session_state:
    st.session_state.mapa_manual = {}
if 'df_long_cache' not in st.session_state:
    st.session_state.df_long_cache = None
if 'df_base_cache' not in st.session_state:
    st.session_state.df_base_cache = None
if 'nombres_pendientes' not in st.session_state:
    st.session_state.nombres_pendientes = []

# --- FUNCIONES ---

def limpiar_texto(texto):
    if pd.isna(texto): return ""
    texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.upper().strip()

def normalizar_fecha_universal(valor):
    if pd.isna(valor) or str(valor).strip() == "": return None
    try:
        if isinstance(valor, pd.Timestamp): return valor.strftime('%Y-%m-%d')
        dt = pd.to_datetime(valor, dayfirst=True, errors='coerce')
        if pd.notna(dt): return dt.strftime('%Y-%m-%d')
        return str(valor).strip()
    except: return str(valor).strip()

def normalizar_horarios(serie):
    s = serie.astype(str).str.upper().str.strip()
    res = pd.Series(index=s.index, dtype='object')
    mask_libre = s.str.contains('LIBRE', na=False) | s.isna() | (s == 'NAN')
    res[mask_libre] = 'L'
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
    try: return pd.read_excel(archivo)
    except:
        archivo.seek(0)
        try: return pd.read_csv(archivo, sep=';', engine='python')
        except: 
            archivo.seek(0)
            return pd.read_csv(archivo, sep=',', engine='python')

# --- LOGICA DE COINCIDENCIA ---

def analizar_nombres(nombres_unicos, df_colab):
    """
    Retorna:
    1. mapa_seguro: {NombreInput: RUT} (Coincidencias 100% seguras)
    2. pendientes: [NombreInput] (Los que no cruzaron y necesitan ayuda humana)
    """
    mapa_seguro = {}
    pendientes = []
    
    df_colab['Nombre_Clean'] = df_colab['Nombre del Colaborador'].apply(limpiar_texto)
    df_colab = df_colab.dropna(subset=['RUT'])
    
    rut_lookup = df_colab.set_index('Nombre_Clean')['RUT'].to_dict()
    lista_nombres_reales = df_colab['Nombre_Clean'].unique()

    for nombre in nombres_unicos:
        if not nombre or pd.isna(nombre): continue
        n_clean = limpiar_texto(nombre)
        partes = n_clean.split()
        
        # 1. Estrategia Exacta
        matches = [real for real in lista_nombres_reales if all(p in real for p in partes)]
        
        if len(matches) == 1:
            mapa_seguro[nombre] = rut_lookup[matches[0]]
        else:
            # Si hay dudas o no existe, va a pendientes
            pendientes.append(nombre)
            
    return mapa_seguro, pendientes

# --- INTERFAZ ---

st.info("💡 Sube tus archivos. Si hay nombres mal escritos, la app te pedirá ayuda antes de descargar.")

col1, col2 = st.columns(2)
archivo_input = col1.file_uploader("1. Excel Supervisores", type=["xlsx"], key="input")
archivo_plantilla = col2.file_uploader("2. Plantilla BUK", type=["xls", "xlsx", "csv"], key="plantilla")

# Botón de reinicio si se cargan nuevos archivos
if archivo_input and archivo_plantilla and st.session_state.etapa == 'carga':
    if st.button("🔍 Analizar y Buscar Coincidencias"):
        with st.spinner("Leyendo datos..."):
            try:
                # Carga inicial
                xls = pd.ExcelFile(archivo_input)
                df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=2)
                df_base = pd.read_excel(xls, sheet_name='Base de Colaboradores')
                df_cods = pd.read_excel(xls, sheet_name='Codificación de Turnos')
                
                # Pre-procesamiento
                df_base.columns = [str(c).strip() for c in df_base.columns]
                col_nom = df_turnos.columns[0]
                
                # Normalizar fechas columnas
                mapa_cols_fechas = {}
                cols_fechas_originales = []
                for c in df_turnos.columns:
                    if c != col_nom:
                        fn = normalizar_fecha_universal(c)
                        if fn:
                            mapa_cols_fechas[c] = fn
                            cols_fechas_originales.append(c)
                
                # Melt
                df_long = df_turnos.melt(id_vars=[col_nom], value_vars=cols_fechas_originales, var_name='Fecha_Original', value_name='Turno_Raw')
                df_long = df_long.dropna(subset=[col_nom])
                df_long['Fecha_Norm'] = df_long['Fecha_Original'].map(mapa_cols_fechas)
                
                # Procesar Horarios una sola vez
                df_long['Turno_Norm'] = normalizar_horarios(df_long['Turno_Raw'])
                df_cods['Horario_Norm'] = normalizar_horarios(df_cods['Horario'])
                dic_turnos = dict(zip(df_cods['Horario_Norm'], df_cods['Sigla']))
                dic_turnos['L'] = 'L'
                df_long['Sigla'] = df_long['Turno_Norm'].map(dic_turnos)

                # ANÁLISIS DE NOMBRES
                nombres_unicos = df_long[col_nom].unique()
                mapa_seguro, pendientes = analizar_nombres(nombres_unicos, df_base)
                
                # Guardar en sesión
                st.session_state.df_long_cache = df_long
                st.session_state.df_base_cache = df_base
                st.session_state.mapa_manual = mapa_seguro # Iniciamos con los seguros
                st.session_state.nombres_pendientes = pendientes
                st.session_state.col_nom_input = col_nom # Nombre columna input
                
                st.session_state.etapa = 'correccion' # Avanzar etapa
                st.rerun() # Recargar pantalla

            except Exception as e:
                st.error(f"Error al leer archivos: {e}")

# --- FASE 2: CORRECCIÓN MANUAL ---
if st.session_state.etapa == 'correccion':
    
    st.divider()
    pendientes = st.session_state.nombres_pendientes
    df_base = st.session_state.df_base_cache
    
    # Preparar lista de opciones para el Combobox
    # Formato: "NOMBRE REAL (RUT)"
    df_base['Opcion_Display'] = df_base['Nombre del Colaborador'].astype(str) + " (" + df_base['RUT'].astype(str) + ")"
    opciones_base = df_base['Opcion_Display'].tolist()
    opciones_base.sort()
    
    if len(pendientes) > 0:
        st.warning(f"⚠️ Se encontraron {len(pendientes)} nombres que no coinciden exactamente. Por favor corrígelos manualmente.")
        
        with st.form("form_correcciones"):
            st.write("### 🛠️ Panel de Corrección")
            
            # Diccionario temporal para guardar selecciones del form
            nuevas_correcciones = {}
            
            col_form1, col_form2 = st.columns([1, 2])
            
            for i, mal_nombre in enumerate(pendientes):
                # Intentar sugerir el más parecido por defecto
                n_clean = limpiar_texto(mal_nombre)
                lista_nombres_clean = df_base['Nombre del Colaborador'].apply(limpiar_texto).tolist()
                sugerencia_idx = 0
                
                posibles = difflib.get_close_matches(n_clean, lista_nombres_clean, n=1, cutoff=0.4)
                if posibles:
                    # Buscar el índice de esa sugerencia en la lista de opciones
                    nombre_real_sugerido = df_base[df_base['Nombre del Colaborador'].apply(limpiar_texto) == posibles[0]].iloc[0]['Opcion_Display']
                    try:
                        sugerencia_idx = opciones_base.index(nombre_real_sugerido)
                    except:
                        sugerencia_idx = 0
                
                st.write(f"**{i+1}. Input:** `{mal_nombre}`")
                seleccion = st.selectbox(
                    f"Corresponde a:", 
                    options=opciones_base, 
                    index=sugerencia_idx,
                    key=f"sel_{i}"
                )
                st.markdown("---")
                
                # Extraer RUT de la selección "NOMBRE (RUT)"
                rut_elegido = seleccion.split("(")[-1].replace(")", "").strip()
                nuevas_correcciones[mal_nombre] = rut_elegido

            confirmar = st.form_submit_button("✅ Confirmar Correcciones y Generar")
            
            if confirmar:
                # Unir mapas
                st.session_state.mapa_manual.update(nuevas_correcciones)
                st.session_state.etapa = 'descarga'
                st.rerun()
    else:
        st.success("✅ ¡Todos los nombres coincidieron perfectamente!")
        if st.button("Continuar a Descarga"):
            st.session_state.etapa = 'descarga'
            st.rerun()

# --- FASE 3: GENERACIÓN Y DESCARGA ---
if st.session_state.etapa == 'descarga':
    st.divider()
    st.write("### 🚀 Generando Archivo Final...")
    
    try:
        # Recuperar datos de memoria
        df_long = st.session_state.df_long_cache
        df_base = st.session_state.df_base_cache
        mapa_final = st.session_state.mapa_manual
        col_nom = st.session_state.col_nom_input
        
        # APLICAR MAPA CORREGIDO
        df_long['RUT'] = df_long[col_nom].map(mapa_final)
        
        # PIVOTE FINAL
        df_pivot = df_long.pivot(index='RUT', columns='Fecha_Norm', values='Sigla')
        
        # LLENAR PLANTILLA
        df_template = cargar_plantilla_robusta(archivo_plantilla)
        cols_template = df_template.columns.tolist()
        
        filas_nuevas = []
        ruts_procesar = df_long['RUT'].unique()
        
        for rut in ruts_procesar:
            if pd.isna(rut): continue
            
            fila = {}
            # Buscar datos maestros usando el RUT corregido
            datos_maestros = df_base[df_base['RUT'] == rut]
            info_colab = datos_maestros.iloc[0] if not datos_maestros.empty else pd.Series()
            
            for col in cols_template:
                col_u = str(col).upper().strip()
                col_fecha_norm = normalizar_fecha_universal(col)
                
                if 'RUT' in col_u or 'EMPLEADO' in col_u:
                    fila[col] = rut
                elif col_fecha_norm and col_fecha_norm in df_pivot.columns:
                    val = df_pivot.loc[rut, col_fecha_norm]
                    fila[col] = val if pd.notna(val) else ""
                elif ('NOMBRE' in col_u or 'NAME' in col_u) and not info_colab.empty:
                    fila[col] = info_colab.get('Nombre del Colaborador', '')
                elif ('AREA' in col_u or 'ÁREA' in col_u) and not info_colab.empty:
                    fila[col] = info_colab.get('Área', info_colab.get('Area', ''))
                elif ('SUPERVISOR' in col_u) and not info_colab.empty:
                    fila[col] = info_colab.get('Supervisor', '')
                else:
                    fila[col] = ""
            filas_nuevas.append(fila)
            
        df_final = pd.DataFrame(filas_nuevas)
        df_final = df_final[cols_template] # Orden estricto
        
        # EXPORTAR
        output = io.BytesIO()
        df_final.to_excel(output, index=False, engine='xlwt')
        
        st.success("✅ Archivo generado exitosamente.")
        
        col_down, col_reset = st.columns(2)
        
        col_down.download_button(
            label="📥 Descargar Output (.xls)",
            data=output.getvalue(),
            file_name="Carga_BUK_Corregida.xls",
            mime="application/vnd.ms-excel"
        )
        
        if col_reset.button("🔄 Comenzar de nuevo"):
            st.session_state.clear()
            st.rerun()

    except Exception as e:
        st.error(f"Error en generación: {e}")

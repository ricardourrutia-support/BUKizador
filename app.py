import streamlit as st
import pandas as pd
import io
import unicodedata
import difflib

# --- CONFIGURACIÓN MINIMALISTA ---
st.set_page_config(
    page_title="BUKizador",
    page_icon="🤖",
    layout="centered",
    initial_sidebar_state="collapsed"
)

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
    .reportview-container .main .block-container{padding-top: 2rem;}
    </style>
""", unsafe_allow_html=True)

st.title("🤖 BUKizador")
st.caption("Motor V3: Sincronización de Fechas Inteligente")

# --- LÓGICA CORE ---

def limpiar_texto(texto):
    if pd.isna(texto): return ""
    texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.upper().strip()

def normalizar_fecha_universal(valor):
    """
    Convierte CUALQUIER formato de fecha (Excel, string, timestamp) 
    a un string estándar 'YYYY-MM-DD' para asegurar comparaciones perfectas.
    """
    if pd.isna(valor) or str(valor).strip() == "":
        return None
    
    try:
        # Si ya es timestamp de pandas
        if isinstance(valor, pd.Timestamp):
            return valor.strftime('%Y-%m-%d')
        
        # Intentamos convertir lo que sea a fecha
        # dayfirst=True ayuda con formatos latinos (DD/MM/YYYY)
        dt = pd.to_datetime(valor, dayfirst=True, errors='coerce')
        
        if pd.notna(dt):
            return dt.strftime('%Y-%m-%d')
        
        # Si falla (ej: es un texto que no es fecha "Nombre"), devolvemos string limpio
        return str(valor).strip()
    except:
        return str(valor).strip()

def buscar_rut_inteligente(nombres_unicos, df_colaboradores):
    mapa_ruts = {}
    df_colab = df_colaboradores.copy()
    df_colab['Nombre_Clean'] = df_colab['Nombre del Colaborador'].apply(limpiar_texto)
    
    # Limpiar duplicados vacíos
    df_colab = df_colab.dropna(subset=['RUT'])
    rut_lookup = df_colab.set_index('Nombre_Clean')['RUT'].to_dict()
    lista_nombres = df_colab['Nombre_Clean'].unique()

    for nombre in nombres_unicos:
        if not nombre or pd.isna(nombre): continue
        n_clean = limpiar_texto(nombre)
        partes = n_clean.split()
        
        matches = [real for real in lista_nombres if all(p in real for p in partes)]
        
        if len(matches) == 1:
            mapa_ruts[nombre] = rut_lookup[matches[0]]
        elif len(matches) > 1:
            mapa_ruts[nombre] = "ERROR: Múltiples"
        else:
            posibles = difflib.get_close_matches(n_clean, lista_nombres, n=1, cutoff=0.7)
            if posibles:
                mapa_ruts[nombre] = rut_lookup[posibles[0]]
            else:
                mapa_ruts[nombre] = "ERROR: No encontrado"
    return mapa_ruts

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
    try:
        return pd.read_excel(archivo)
    except:
        archivo.seek(0)
        try:
            return pd.read_csv(archivo, sep=';', engine='python')
        except:
            archivo.seek(0)
            return pd.read_csv(archivo, sep=',', engine='python')

# --- INTERFAZ ---

with st.container():
    col1, col2 = st.columns(2)
    archivo_input = col1.file_uploader("1. Excel Supervisores", type=["xlsx"])
    archivo_plantilla = col2.file_uploader("2. Plantilla BUK", type=["xls", "xlsx", "csv"])

if archivo_input and archivo_plantilla:
    with st.spinner('Sincronizando fechas y procesando...'):
        try:
            # 1. CARGA INPUT
            xls = pd.ExcelFile(archivo_input)
            df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=2)
            df_base = pd.read_excel(xls, sheet_name='Base de Colaboradores')
            df_cods = pd.read_excel(xls, sheet_name='Codificación de Turnos')

            # Limpieza headers base (fix problema Area)
            df_base.columns = [str(c).strip() for c in df_base.columns]

            # 2. PROCESAMIENTO FECHAS INPUT
            # Detectar columna de nombre (asumimos la primera)
            col_nom = df_turnos.columns[0]
            
            # Crear mapa de fechas normalizadas para el input
            # { "2026-01-01 00:00:00": "2026-01-01", ... }
            mapa_cols_fechas = {}
            cols_fechas_originales = []
            
            for c in df_turnos.columns:
                if c != col_nom:
                    fecha_norm = normalizar_fecha_universal(c)
                    if fecha_norm: # Si es una fecha válida
                        mapa_cols_fechas[c] = fecha_norm
                        cols_fechas_originales.append(c)

            # 3. MELT & PROCESAMIENTO
            df_long = df_turnos.melt(id_vars=[col_nom], value_vars=cols_fechas_originales, var_name='Fecha_Original', value_name='Turno_Raw')
            df_long = df_long.dropna(subset=[col_nom])
            
            # Asignar fecha normalizada
            df_long['Fecha_Norm'] = df_long['Fecha_Original'].map(mapa_cols_fechas)

            # Rut
            mapa_rut = buscar_rut_inteligente(df_long[col_nom].unique(), df_base)
            df_long['RUT'] = df_long[col_nom].map(mapa_rut)
            
            # Horarios
            df_long['Turno_Norm'] = normalizar_horarios(df_long['Turno_Raw'])
            df_cods['Horario_Norm'] = normalizar_horarios(df_cods['Horario'])
            dic_turnos = dict(zip(df_cods['Horario_Norm'], df_cods['Sigla']))
            dic_turnos['L'] = 'L'
            df_long['Sigla'] = df_long['Turno_Norm'].map(dic_turnos)
            
            # Pivot usando FECHA NORMALIZADA como columna
            df_pivot = df_long.pivot(index='RUT', columns='Fecha_Norm', values='Sigla')

            # 4. CRUCE CON PLANTILLA
            df_template = cargar_plantilla_robusta(archivo_plantilla)
            cols_template_orig = df_template.columns.tolist()
            
            filas_nuevas = []
            ruts_validos = [r for r in df_long['RUT'].unique() if "ERROR" not in str(r)]

            # Contadores para debug
            fechas_encontradas = 0
            fechas_totales_plantilla = 0

            for rut in ruts_validos:
                fila = {}
                # Info colaborador
                datos_maestros = df_base[df_base['RUT'] == rut]
                info_colab = datos_maestros.iloc[0] if not datos_maestros.empty else pd.Series()
                
                for col in cols_template_orig:
                    col_u = str(col).upper().strip()
                    
                    # Intentar normalizar encabezado plantilla
                    col_fecha_norm = normalizar_fecha_universal(col)
                    
                    # Lógica de llenado
                    if 'RUT' in col_u or 'EMPLEADO' in col_u:
                        fila[col] = rut
                    
                    # MAGIA DE FECHAS: Comparamos normalizado vs normalizado
                    elif col_fecha_norm and col_fecha_norm in df_pivot.columns:
                        val = df_pivot.loc[rut, col_fecha_norm]
                        fila[col] = val if pd.notna(val) else ""
                        if rut == ruts_validos[0]: dates_matched = True # Flag para debug
                        
                    # Datos maestros
                    elif ('NOMBRE' in col_u or 'NAME' in col_u) and not info_colab.empty:
                        fila[col] = info_colab.get('Nombre del Colaborador', '')
                    elif ('AREA' in col_u or 'ÁREA' in col_u) and not info_colab.empty:
                        # Busca 'Area' o 'Área'
                        val_area = info_colab.get('Área', info_colab.get('Area', ''))
                        fila[col] = val_area
                    elif ('SUPERVISOR' in col_u) and not info_colab.empty:
                        fila[col] = info_colab.get('Supervisor', '')
                    else:
                        fila[col] = ""
                
                filas_nuevas.append(fila)
            
            df_final = pd.DataFrame(filas_nuevas)
            df_final = df_final[cols_template_orig] # Orden estricto

            # 5. SALIDA
            output = io.BytesIO()
            df_final.to_excel(output, index=False, engine='xlwt')
            
            st.success(f"✅ Proceso completado. {len(df_final)} filas generadas.")
            
            st.download_button(
                label="📥 Descargar .XLS Final",
                data=output.getvalue(),
                file_name="Carga_BUK_V3.xls",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )

            # --- DEBUGGING PANEL (Solo si algo sale raro) ---
            with st.expander("🛠 Panel de Diagnóstico (Abre si faltan datos)"):
                st.write("### Chequeo de Fechas")
                fechas_input = sorted(list(df_pivot.columns))
                st.write(f"**Fechas detectadas en Input (Formato YYYY-MM-DD):** {len(fechas_input)}")
                st.code(fechas_input[:5]) # Mostrar primeras 5

                # Ver qué fechas detectó en la plantilla
                fechas_plantilla_detectadas = []
                for c in cols_template_orig:
                    fn = normalizar_fecha_universal(c)
                    if fn and fn in fechas_input:
                        fechas_plantilla_detectadas.append(f"{c} -> {fn} (MATCH ✅)")
                    elif fn: # Parece fecha pero no está en input
                        fechas_plantilla_detectadas.append(f"{c} -> {fn} (NO MATCH ❌)")
                
                st.write(f"**Fechas coincidentes con Plantilla:** {len([x for x in fechas_plantilla_detectadas if 'MATCH' in x])}")
                st.write(fechas_plantilla_detectadas)

        except Exception as e:
            st.error(f"Error Crítico: {e}")

else:
    st.info("Esperando archivos...")

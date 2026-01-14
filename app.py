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
    layout="centered",
    initial_sidebar_state="collapsed"
)

# CSS Minimalista
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
st.caption("Generador de .XLS (97-2003) para BUK")

# --- LÓGICA CORE ---

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
    """Detecta si es Excel real o CSV disfrazado."""
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
    archivo_plantilla = col2.file_uploader("2. Plantilla BUK (.xls)", type=["xls", "xlsx", "csv"])

if archivo_input and archivo_plantilla:
    with st.spinner('Procesando a formato antiguo .XLS...'):
        try:
            # 1. PROCESAMIENTO
            xls = pd.ExcelFile(archivo_input)
            df_turnos = pd.read_excel(xls, sheet_name='Turnos Formato Supervisor', header=2)
            df_base = pd.read_excel(xls, sheet_name='Base de Colaboradores')
            df_cods = pd.read_excel(xls, sheet_name='Codificación de Turnos')

            col_nom = df_turnos.columns[0]
            cols_fechas = [c for c in df_turnos.columns if c != col_nom]
            df_long = df_turnos.melt(id_vars=[col_nom], value_vars=cols_fechas, var_name='Fecha', value_name='Turno_Raw')
            df_long = df_long.dropna(subset=[col_nom])

            mapa = buscar_rut_inteligente(df_long[col_nom].unique(), df_base)
            df_long['RUT'] = df_long[col_nom].map(mapa)
            
            df_long['Turno_Norm'] = normalizar_horarios(df_long['Turno_Raw'])
            df_cods['Horario_Norm'] = normalizar_horarios(df_cods['Horario'])
            dic_turnos = dict(zip(df_cods['Horario_Norm'], df_cods['Sigla']))
            dic_turnos['L'] = 'L'
            df_long['Sigla'] = df_long['Turno_Norm'].map(dic_turnos)
            
            df_pivot = df_long.pivot(index='RUT', columns='Fecha', values='Sigla')

            # 2. INYECCIÓN
            df_template = cargar_plantilla_robusta(archivo_plantilla)
            cols_template = df_template.columns.tolist()
            
            filas_nuevas = []
            ruts_validos = [r for r in df_long['RUT'].unique() if "ERROR" not in str(r)]

            for rut in ruts_validos:
                fila = {}
                info_colab = df_base[df_base['RUT'] == rut].iloc[0] if not df_base[df_base['RUT'] == rut].empty else {}
                
                for col in cols_template:
                    col_u = col.upper()
                    if 'RUT' in col_u or 'EMPLEADO' in col_u:
                        fila[col] = rut
                    elif col in df_pivot.columns: 
                        val = df_pivot.loc[rut, col]
                        fila[col] = val if pd.notna(val) else ""
                    elif 'NOMBRE' in col_u and not info_colab.empty:
                        fila[col] = info_colab.get('Nombre del Colaborador', '')
                    elif 'AREA' in col_u and not info_colab.empty:
                        fila[col] = info_colab.get('Área', '')
                    elif 'SUPERVISOR' in col_u and not info_colab.empty:
                        fila[col] = info_colab.get('Supervisor', '')
                    else:
                        fila[col] = ""
                
                filas_nuevas.append(fila)
            
            df_final = pd.DataFrame(filas_nuevas)
            df_final = df_final[cols_template]

            # 3. GUARDADO .XLS
            output = io.BytesIO()
            # IMPORTANTE: Aquí es donde fallaba si pandas era muy nuevo
            df_final.to_excel(output, index=False, engine='xlwt')
            
            st.success(f"✅ ¡Listo! Archivo .xls (Legacy) generado.")
            
            st.download_button(
                label="📥 Descargar Output (.xls)",
                data=output.getvalue(),
                file_name="Carga_BUK.xls",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )

        except ImportError as e:
            st.error("Error de Versión de Librerías")
            st.warning("El servidor instaló una versión de Pandas demasiado nueva que no soporta .xls. Asegúrate de poner 'pandas<2.0.0' en requirements.txt")
            st.code(e)
            
        except Exception as e:
            st.error(f"Error Técnico: {e}")

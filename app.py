import streamlit as st
import pandas as pd
import io
import os
import re
import unicodedata
import difflib
import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="BUKizador v3", page_icon="✈️", layout="centered")

st.markdown("""
    <style>
    .stApp {background-color: #FAFAFA;}
    .block-container {padding-top: 1rem !important;}
    h1 {color: #1a1a2e;}
    .stButton>button {width: 100%; border-radius: 8px; font-weight: 600;}
    .match-ok {color: #27ae60; font-weight: bold;}
    .match-warn {color: #e67e22; font-weight: bold;}
    .match-err {color: #e74c3c; font-weight: bold;}
    div[data-testid="stForm"] {
        background-color: #ffffff; padding: 20px; border-radius: 10px;
        border: 1px solid #ddd; box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    div[data-testid="stImage"] img {border-radius: 12px;}
    </style>
""", unsafe_allow_html=True)

# --- HEADER ---
if os.path.exists("header.png"):
    st.image("header.png", use_container_width=True)

st.title("✈️ BUKizador v3")
st.caption("Input 1: Turnos 360 (supervisores) · Input 2: Importador BUK (.xls)")

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIONES AUXILIARES
# ═══════════════════════════════════════════════════════════════════════════════

def limpiar_texto(texto):
    """Normaliza texto: quita acentos, mayúsculas, espacios extra."""
    if pd.isna(texto) or texto is None:
        return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto.upper().strip()


def normalizar_hora(texto):
    """
    Convierte cualquier formato de hora a 'HH:MM' estándar.
    Maneja: '8:00', '08:00', '8:30', datetime.time, etc.
    """
    if pd.isna(texto) or str(texto).strip() in ['-', '', 'nan']:
        return None
    texto = str(texto).strip()
    # Si es un time object
    if isinstance(texto, datetime.time):
        return f"{texto.hour:02d}:{texto.minute:02d}"
    # Extraer HH:MM con regex
    match = re.search(r'(\d{1,2}):(\d{2})', texto)
    if match:
        h, m = int(match.group(1)), match.group(2)
        return f"{h:02d}:{m}"
    # Solo número (ej: "8" → "08:00")
    match = re.match(r'^(\d{1,2})$', texto)
    if match:
        return f"{int(match.group(1)):02d}:00"
    return None


def extraer_rango_horario(texto):
    """
    Extrae (entrada, salida) de un texto como '08:00 - 19:00' o '09:00-20:00'.
    Retorna tupla de strings normalizados o ('LIBRE', 'LIBRE') o None si error.
    """
    if pd.isna(texto):
        return None
    texto = str(texto).strip().upper()
    
    if texto in ['', 'NAN']:
        return None
    
    # Detectar "Libre" / "Descanso"
    if 'LIBRE' in texto or 'DESCANSO' in texto:
        return ('LIBRE', 'LIBRE')
    
    # Normalizar separadores
    texto_sep = re.sub(r'\s*[-–—]\s*', '-', texto)  # guiones
    texto_sep = re.sub(r'\s+A\s+|\s+AL\s+', '-', texto_sep)  # "a" / "al"
    
    # Extraer todos los patrones HH:MM
    patron = r'(\d{1,2}):(\d{2})'
    matches = re.findall(patron, texto_sep)
    
    if len(matches) >= 2:
        h1, m1 = int(matches[0][0]), matches[0][1]
        h2, m2 = int(matches[-1][0]), matches[-1][1]
        entrada = f"{h1:02d}:{m1}"
        salida = f"{h2:02d}:{m2}"
        return (entrada, salida)
    
    # Un solo HH:MM no es un rango válido
    return None


def detectar_fila_fechas(df):
    """Encuentra la fila que contiene fechas (datetime) en el DataFrame."""
    for i in range(min(10, len(df))):
        count_dates = 0
        for j in range(1, min(40, df.shape[1])):
            val = df.iloc[i, j]
            if isinstance(val, (datetime.datetime, pd.Timestamp)):
                count_dates += 1
        if count_dates >= 5:  # al menos 5 fechas
            return i
    return None


def parsear_hoja_turnos(df, nombre_hoja):
    """
    Parsea una hoja de turnos del formato 360.
    Retorna DataFrame con columnas: [Nombre, Fecha, Turno_Raw, Rol]
    """
    fila_fechas = detectar_fila_fechas(df)
    if fila_fechas is None:
        return pd.DataFrame()
    
    # Extraer fechas de esa fila
    fechas = {}
    for j in range(1, df.shape[1]):
        val = df.iloc[fila_fechas, j]
        if isinstance(val, (datetime.datetime, pd.Timestamp)):
            fechas[j] = pd.Timestamp(val).strftime('%Y-%m-%d')
    
    if not fechas:
        return pd.DataFrame()
    
    # Determinar dónde empiezan los datos
    fila_data = fila_fechas + 1
    # Saltar filas de encabezado como "Cargo", "Nombre", "Supervisor"
    while fila_data < len(df):
        val = df.iloc[fila_data, 0]
        if pd.isna(val):
            fila_data += 1
            continue
        val_str = str(val).strip().upper()
        if val_str in ['CARGO', 'NOMBRE', 'SUPERVISOR', '']:
            fila_data += 1
            continue
        break
    
    # Determinar rol y mes a partir del nombre de la hoja
    nombre_upper = nombre_hoja.upper()
    if 'ANFITRION' in nombre_upper:
        rol = 'ANFITRION'
    elif 'AGENTE' in nombre_upper:
        rol = 'AGENTE'
    elif 'COORDINADOR' in nombre_upper:
        rol = 'COORDINADOR'
    elif 'SUPERVISOR' in nombre_upper:
        rol = 'SUPERVISOR'
    else:
        rol = 'OTRO'
    
    # Detectar mes del nombre de la hoja para tiebreaking de solapamientos
    MESES_ES = {
        'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
        'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12
    }
    mes_hoja = None
    for nombre_mes, num_mes in MESES_ES.items():
        if nombre_mes in nombre_upper:
            mes_hoja = num_mes
            break
    
    registros = []
    PALABRAS_HEADER = {'NOMBRE', 'CARGO', 'SUPERVISOR', 'COLABORADOR', 'NOMBRE DEL COLABORADOR', 'TRABAJADOR', 'EMPLEADO', 'RUT'}
    for i in range(fila_data, len(df)):
        nombre = df.iloc[i, 0]
        if pd.isna(nombre):
            continue
        # Saltar filas con valores numéricos (filas de totales/resumen)
        if isinstance(nombre, (int, float)):
            continue
        nombre_str = str(nombre).strip()
        if nombre_str in ['.', '', 'nan', 'NaN']:
            continue
        # Debe contener al menos una letra
        if not any(c.isalpha() for c in nombre_str):
            continue
        # Saltar palabras que son encabezados de columna (no son nombres reales)
        if nombre_str.upper() in PALABRAS_HEADER:
            continue
        
        for col_idx, fecha_str in fechas.items():
            turno_raw = df.iloc[i, col_idx] if col_idx < df.shape[1] else None
            registros.append({
                'Nombre_Input': nombre_str,
                'Fecha': fecha_str,
                'Turno_Raw': turno_raw,
                'Rol': rol,
                'Hoja': nombre_hoja,
                'Mes_Hoja': mes_hoja
            })
    
    return pd.DataFrame(registros)


def construir_mapa_siglas(df_turnos_semanales):
    """
    Construye un diccionario: (entrada, salida, rol) → sigla
    a partir de la hoja turnosSemanales del importador BUK.
    """
    df = df_turnos_semanales.copy()
    df.columns = ['Nombre', 'Sigla', 'Dia', 'Entrada', 'Salida', 'ColIn', 'ColOut']
    df = df.iloc[1:]  # Quitar header
    
    # Agrupar por sigla para obtener entrada/salida únicos
    mapa = {}
    for sigla in df['Sigla'].unique():
        sub = df[df['Sigla'] == sigla]
        entradas = [str(e).strip() for e in sub['Entrada'].unique() if str(e).strip() != '-']
        salidas = [str(s).strip() for s in sub['Salida'].unique() if str(s).strip() != '-']
        nombre = str(sub['Nombre'].iloc[0]).strip().upper()
        
        if not entradas or not salidas:
            # Es un turno sin horario (D, F, L, P, V, C)
            continue
        
        entrada_norm = normalizar_hora(entradas[0])
        salida_norm = normalizar_hora(salidas[0])
        
        if entrada_norm and salida_norm:
            # Determinar a qué rol pertenece esta sigla
            roles = []
            if 'ANF' in sigla.upper() or 'ANFITRION' in nombre:
                roles.append('ANFITRION')
            if 'AGE' in sigla.upper() or 'AGENTE' in nombre:
                roles.append('AGENTE')
            if 'COO' in sigla.upper() or 'COORDINADOR' in nombre:
                roles.append('COORDINADOR')
            if 'SUP' in sigla.upper() or 'SUPERVISOR' in nombre:
                roles.append('SUPERVISOR')
            if 'INDUC' in sigla.upper() or 'INDUCCION' in nombre or 'INDUCCIÓN' in nombre:
                roles.append('INDUCCION')
            if 'BASE' in sigla.upper():
                roles = ['ANFITRION', 'AGENTE', 'COORDINADOR', 'SUPERVISOR', 'OTRO']
            
            if not roles:
                roles = ['OTRO']
            
            for rol in roles:
                key = (entrada_norm, salida_norm, rol)
                mapa[key] = sigla
    
    return mapa


def turno_a_sigla(turno_raw, rol, mapa_siglas):
    """Convierte un turno en texto humano a su sigla BUK."""
    rango = extraer_rango_horario(turno_raw)
    
    if rango is None:
        return None  # Celda vacía o no parseable
    
    if rango == ('LIBRE', 'LIBRE'):
        return 'L'  # Libre
    
    entrada, salida = rango
    
    # Buscar con rol exacto
    key = (entrada, salida, rol)
    if key in mapa_siglas:
        return mapa_siglas[key]
    
    # Manejar medianoche: "00:00" como salida → probar con "23:59"
    if salida == '00:00':
        key_midnight = (entrada, '23:59', rol)
        if key_midnight in mapa_siglas:
            return mapa_siglas[key_midnight]
    
    # Fallback: buscar en cualquier rol
    for (e, s, r), sigla in mapa_siglas.items():
        if e == entrada and s == salida:
            return sigla
    
    # Fallback medianoche en cualquier rol
    if salida == '00:00':
        for (e, s, r), sigla in mapa_siglas.items():
            if e == entrada and s == '23:59':
                return sigla
    
    return None  # No encontrado


def matching_nombres(nombres_input, nombres_buk):
    """
    Hace matching inteligente entre nombres cortos (input) y nombres completos (BUK).
    Retorna: (mapa_seguro, pendientes)
      - mapa_seguro: {nombre_input: nombre_buk}
      - pendientes: [nombre_input, ...] que necesitan corrección manual
    """
    nombres_buk_clean = {limpiar_texto(n): n for n in nombres_buk}
    lista_clean = list(nombres_buk_clean.keys())
    
    mapa_seguro = {}
    pendientes = []
    
    for nombre in nombres_input:
        n_clean = limpiar_texto(nombre)
        partes = n_clean.split()
        
        if not partes:
            continue
        
        # Estrategia 1: Todas las palabras del input aparecen en algún nombre BUK
        matches = [real for real in lista_clean if all(p in real for p in partes)]
        
        if len(matches) == 1:
            mapa_seguro[nombre] = nombres_buk_clean[matches[0]]
        elif len(matches) > 1:
            # Intentar desempatar: el que tenga menos "basura" extra
            best = min(matches, key=lambda x: len(x) - len(n_clean))
            mapa_seguro[nombre] = nombres_buk_clean[best]
        else:
            # Estrategia 2: Coincidencia difusa
            posibles = difflib.get_close_matches(n_clean, lista_clean, n=1, cutoff=0.6)
            if posibles:
                mapa_seguro[nombre] = nombres_buk_clean[posibles[0]]
            else:
                pendientes.append(nombre)
    
    return mapa_seguro, pendientes


# ═══════════════════════════════════════════════════════════════════════════════
# ESTADO DE SESIÓN
# ═══════════════════════════════════════════════════════════════════════════════

if 'etapa' not in st.session_state:
    st.session_state.etapa = 'carga'
if 'df_all_turnos' not in st.session_state:
    st.session_state.df_all_turnos = None
if 'mapa_siglas' not in st.session_state:
    st.session_state.mapa_siglas = None
if 'mapa_nombres' not in st.session_state:
    st.session_state.mapa_nombres = {}
if 'pendientes' not in st.session_state:
    st.session_state.pendientes = []
if 'nombres_buk' not in st.session_state:
    st.session_state.nombres_buk = []
if 'df_buk_header' not in st.session_state:
    st.session_state.df_buk_header = None
if 'df_buk_data' not in st.session_state:
    st.session_state.df_buk_data = None
if 'buk_bytes' not in st.session_state:
    st.session_state.buk_bytes = None
if 'hojas_mes' not in st.session_state:
    st.session_state.hojas_mes = []
if 'turnos_no_encontrados' not in st.session_state:
    st.session_state.turnos_no_encontrados = []

# ═══════════════════════════════════════════════════════════════════════════════
# FASE 1: CARGA DE ARCHIVOS
# ═══════════════════════════════════════════════════════════════════════════════

st.info("💡 Sube los 2 archivos: el Excel de turnos (formato 360) y el importador BUK (.xls)")

col1, col2 = st.columns(2)
archivo_360 = col1.file_uploader("📋 Turnos 360 (supervisores)", type=["xlsx"], key="input360")
archivo_buk = col2.file_uploader("📦 Importador BUK", type=["xls", "xlsx"], key="inputbuk")

if archivo_360 and archivo_buk and st.session_state.etapa == 'carga':
    
    try:
        xls360 = pd.ExcelFile(archivo_360)
        hojas = xls360.sheet_names
        
        if not hojas:
            st.error("El archivo 360 no tiene hojas.")
            st.stop()
        
        # Pre-escanear cada hoja para mostrar su rango de fechas
        st.write("**Hojas detectadas en el archivo 360:**")
        info_hojas = []
        hojas_validas = []
        for h in hojas:
            df_tmp = pd.read_excel(xls360, sheet_name=h, header=None)
            ff = detectar_fila_fechas(df_tmp)
            if ff is None:
                info_hojas.append(f"  ⚠️ `{h}` — sin formato de fechas reconocible (será omitida)")
                continue
            fechas_tmp = []
            for j in range(1, df_tmp.shape[1]):
                v = df_tmp.iloc[ff, j]
                if isinstance(v, (datetime.datetime, pd.Timestamp)):
                    fechas_tmp.append(pd.Timestamp(v))
            if fechas_tmp:
                rango = f"{min(fechas_tmp).strftime('%d-%m-%Y')} → {max(fechas_tmp).strftime('%d-%m-%Y')}"
                info_hojas.append(f"  ✅ `{h}` — {rango} ({len(fechas_tmp)} días)")
                hojas_validas.append(h)
        
        for line in info_hojas:
            st.markdown(line)
        
        if not hojas_validas:
            st.error("Ninguna hoja tiene un formato de fechas válido.")
            st.stop()
        
        st.write("")
        hojas_seleccionadas = st.multiselect(
            "📋 Hojas a procesar (puedes seleccionar varias para cubrir rangos entre meses):",
            options=hojas_validas,
            default=hojas_validas
        )
        
        st.caption("💡 Si una fecha aparece en varias hojas, se prefiere la hoja cuyo mes coincida con la fecha (ej: 1-abr se toma de 'Abril', no de 'Marzo').")
        
        if not hojas_seleccionadas:
            st.warning("Selecciona al menos una hoja.")
            st.stop()
        
        if st.button("🔍 Analizar y Procesar", type="primary"):
            with st.spinner("Leyendo y procesando datos..."):
                
                # ── LEER IMPORTADOR BUK ──
                st.session_state.buk_bytes = archivo_buk.read()
                archivo_buk.seek(0)
                
                buk_name = archivo_buk.name.lower()
                if buk_name.endswith('.xls') and not buk_name.endswith('.xlsx'):
                    xls_buk = pd.ExcelFile(archivo_buk, engine='xlrd')
                    st.session_state.buk_is_xls = True
                else:
                    xls_buk = pd.ExcelFile(archivo_buk, engine='openpyxl')
                    st.session_state.buk_is_xls = False
                
                # Hoja turnosColaboradores
                df_tc_raw = pd.read_excel(xls_buk, sheet_name='turnosColaboradores', header=None)
                header_row = df_tc_raw.iloc[0].tolist()
                df_tc = df_tc_raw.iloc[1:].copy()
                df_tc.columns = header_row
                df_tc = df_tc.reset_index(drop=True)
                
                st.session_state.df_buk_header = header_row
                st.session_state.df_buk_data = df_tc
                
                nombres_buk = df_tc['Nombre del Colaborador'].tolist()
                ruts_buk = df_tc['RUT'].tolist()
                st.session_state.nombres_buk = nombres_buk
                nombre_a_rut = dict(zip(nombres_buk, ruts_buk))
                st.session_state.nombre_a_rut = nombre_a_rut
                
                # Hoja turnosSemanales (codificación)
                df_ts_raw = pd.read_excel(xls_buk, sheet_name='turnosSemanales', header=None)
                mapa_siglas = construir_mapa_siglas(df_ts_raw)
                st.session_state.mapa_siglas = mapa_siglas
                
                # ── LEER TURNOS 360 (TODAS LAS HOJAS SELECCIONADAS) ──
                all_turnos = []
                for hoja in hojas_seleccionadas:
                    df_hoja = pd.read_excel(xls360, sheet_name=hoja, header=None)
                    df_parsed = parsear_hoja_turnos(df_hoja, hoja)
                    if not df_parsed.empty:
                        all_turnos.append(df_parsed)
                
                if not all_turnos:
                    st.error("No se pudieron parsear turnos de las hojas seleccionadas.")
                    st.stop()
                
                df_all = pd.concat(all_turnos, ignore_index=True)
                
                # ── RESOLUCIÓN DE SOLAPAMIENTOS ──
                # Para (Nombre, Fecha, Rol) duplicados, preferir la hoja cuyo mes coincida
                # con el mes de la fecha. Si ninguno coincide, tomar el primero.
                df_all['_fecha_dt'] = pd.to_datetime(df_all['Fecha'])
                df_all['_mes_fecha'] = df_all['_fecha_dt'].dt.month
                df_all['_match_mes'] = (df_all['Mes_Hoja'] == df_all['_mes_fecha']).astype(int)
                
                # Ordenar: primero los que matchean mes (1), luego por hoja (estable)
                df_all = df_all.sort_values(
                    by=['Nombre_Input', 'Fecha', 'Rol', '_match_mes'],
                    ascending=[True, True, True, False]
                )
                df_all = df_all.drop_duplicates(
                    subset=['Nombre_Input', 'Fecha', 'Rol'],
                    keep='first'
                )
                df_all = df_all.drop(columns=['_fecha_dt', '_mes_fecha', '_match_mes'])
                df_all = df_all.reset_index(drop=True)
                
                st.session_state.df_all_turnos = df_all
                st.session_state.hojas_mes = hojas_seleccionadas
                
                # ── MATCHING DE NOMBRES ──
                nombres_input = df_all['Nombre_Input'].unique().tolist()
                mapa, pendientes = matching_nombres(nombres_input, nombres_buk)
                
                st.session_state.mapa_nombres = mapa
                st.session_state.pendientes = pendientes
                
                st.session_state.etapa = 'correccion'
                st.rerun()
    
    except Exception as e:
        st.error(f"Error al leer archivos: {e}")
        import traceback
        st.code(traceback.format_exc())


# ═══════════════════════════════════════════════════════════════════════════════
# FASE 2: CORRECCIÓN DE NOMBRES
# ═══════════════════════════════════════════════════════════════════════════════

if st.session_state.etapa == 'correccion':
    st.divider()
    
    pendientes = st.session_state.pendientes
    nombres_buk = st.session_state.nombres_buk
    mapa = st.session_state.mapa_nombres
    
    # Mostrar matches automáticos
    n_auto = len(mapa)
    n_total = n_auto + len(pendientes)
    
    st.success(f"✅ {n_auto} de {n_total} nombres emparejados automáticamente.")
    
    with st.expander("Ver matches automáticos", expanded=False):
        for inp, buk in sorted(mapa.items()):
            st.write(f"  `{inp}` → **{buk}**")
    
    if pendientes:
        st.warning(f"⚠️ {len(pendientes)} nombres necesitan corrección manual.")
        
        # Opciones para selectbox
        opciones = sorted(nombres_buk)
        
        with st.form("form_correcciones"):
            st.write("### 🛠️ Corrección Manual")
            
            correcciones = {}
            for i, nombre_mal in enumerate(pendientes):
                # Sugerir el más parecido
                n_clean = limpiar_texto(nombre_mal)
                nombres_buk_clean = [limpiar_texto(n) for n in opciones]
                sugerencia_idx = 0
                
                posibles = difflib.get_close_matches(n_clean, nombres_buk_clean, n=1, cutoff=0.3)
                if posibles:
                    idx_real = nombres_buk_clean.index(posibles[0])
                    sugerencia_idx = idx_real
                
                st.write(f"**{i+1}.** Input del supervisor: `{nombre_mal}`")
                seleccion = st.selectbox(
                    f"Corresponde a:",
                    options=["❌ NO EXISTE EN BUK (omitir)"] + opciones,
                    index=sugerencia_idx + 1,  # +1 por la opción de omitir
                    key=f"corr_{i}"
                )
                st.markdown("---")
                
                if seleccion != "❌ NO EXISTE EN BUK (omitir)":
                    correcciones[nombre_mal] = seleccion
            
            confirmar = st.form_submit_button("✅ Confirmar y Generar", type="primary")
            
            if confirmar:
                st.session_state.mapa_nombres.update(correcciones)
                st.session_state.etapa = 'descarga'
                st.rerun()
    else:
        st.success("🎉 ¡Todos los nombres coincidieron perfectamente!")
        if st.button("▶️ Continuar a Generar Archivo", type="primary"):
            st.session_state.etapa = 'descarga'
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# FASE 3: GENERACIÓN Y DESCARGA
# ═══════════════════════════════════════════════════════════════════════════════

if st.session_state.etapa == 'descarga':
    st.divider()
    st.write("### 🚀 Generando Archivo Final...")
    
    try:
        df_all = st.session_state.df_all_turnos
        mapa_nombres = st.session_state.mapa_nombres
        mapa_siglas = st.session_state.mapa_siglas
        df_buk = st.session_state.df_buk_data
        header_buk = st.session_state.df_buk_header
        nombre_a_rut = st.session_state.nombre_a_rut
        
        # ── Filtrar fechas que existen en el importador BUK ──
        # Las fechas del BUK están como DD-MM-YYYY en el header
        fechas_buk = {}
        for col in header_buk:
            if col in ['Nombre del Colaborador', 'RUT', 'Área', 'Supervisor']:
                continue
            if col is None or pd.isna(col):
                continue
            # Intentar parsear como fecha
            try:
                dt = pd.to_datetime(str(col), format='%d-%m-%Y', errors='raise')
                fechas_buk[dt.strftime('%Y-%m-%d')] = col
            except:
                try:
                    dt = pd.to_datetime(str(col), dayfirst=True, errors='raise')
                    fechas_buk[dt.strftime('%Y-%m-%d')] = col
                except:
                    pass
        
        # ── Convertir turnos a siglas ──
        df_all['Nombre_BUK'] = df_all['Nombre_Input'].map(mapa_nombres)
        # Filtrar solo los que tienen match
        df_con_match = df_all[df_all['Nombre_BUK'].notna()].copy()
        
        # Obtener RUT
        df_con_match['RUT'] = df_con_match['Nombre_BUK'].map(nombre_a_rut)
        
        # Mapear turnos a siglas
        turnos_no_encontrados = set()
        
        def resolver_sigla(row):
            sigla = turno_a_sigla(row['Turno_Raw'], row['Rol'], mapa_siglas)
            if sigla is None and pd.notna(row['Turno_Raw']) and str(row['Turno_Raw']).strip() not in ['', 'nan']:
                turnos_no_encontrados.add(str(row['Turno_Raw']).strip())
                return f"REVISAR:{str(row['Turno_Raw']).strip()}"
            return sigla
        
        df_con_match['Sigla'] = df_con_match.apply(resolver_sigla, axis=1)
        st.session_state.turnos_no_encontrados = list(turnos_no_encontrados)
        
        # ── Construir el DataFrame de salida con la estructura BUK ──
        # Para cada RUT en el BUK, llenar las columnas de fecha con la sigla correspondiente
        df_output = df_buk.copy()
        
        # Determinar cuál es la última fecha del importador (para el truco de D final)
        fechas_iso_ordenadas = sorted(fechas_buk.keys())
        ultima_fecha_iso = fechas_iso_ordenadas[-1] if fechas_iso_ordenadas else None
        ultima_col_buk = fechas_buk.get(ultima_fecha_iso) if ultima_fecha_iso else None
        
        for idx, row_buk in df_output.iterrows():
            rut = row_buk['RUT']
            # Buscar turnos de este colaborador
            turnos_colab = df_con_match[df_con_match['RUT'] == rut]
            
            for fecha_iso, col_buk in fechas_buk.items():
                # Último día del importador → siempre D (truco de configuración BUK)
                if col_buk == ultima_col_buk:
                    df_output.at[idx, col_buk] = 'D'
                    continue
                
                if turnos_colab.empty:
                    # Sin datos del 360 para este colaborador → L (Libre)
                    df_output.at[idx, col_buk] = 'L'
                    continue
                
                turno_dia = turnos_colab[turnos_colab['Fecha'] == fecha_iso]
                if not turno_dia.empty:
                    sigla = turno_dia.iloc[0]['Sigla']
                    if sigla is not None:
                        df_output.at[idx, col_buk] = sigla
                    else:
                        # Celda vacía en el input 360 → L
                        df_output.at[idx, col_buk] = 'L'
                else:
                    # Fecha existe en BUK pero no en el 360 → L
                    df_output.at[idx, col_buk] = 'L'
        
        # ── Detección de turnos problemáticos ──
        # Recorrer df_output buscando celdas REVISAR:... para construir lista de problemas
        problemas = []
        for idx, row_buk in df_output.iterrows():
            rut = row_buk['RUT']
            nombre = row_buk['Nombre del Colaborador']
            for fi, cb in fechas_buk.items():
                val = df_output.at[idx, cb]
                if isinstance(val, str) and val.startswith('REVISAR:'):
                    turno_raw = val.replace('REVISAR:', '', 1)
                    # Buscar el rol desde df_con_match
                    rol_rows = df_con_match[(df_con_match['RUT'] == rut) & (df_con_match['Fecha'] == fi)]
                    rol = rol_rows.iloc[0]['Rol'] if not rol_rows.empty else 'N/A'
                    problemas.append({
                        'key': f"{rut}__{fi}",
                        'rut': rut,
                        'nombre': nombre,
                        'fecha_iso': fi,
                        'fecha_display': cb,
                        'rol': rol,
                        'turno_raw': turno_raw,
                        'idx': idx,
                        'col': cb,
                    })
        
        # Inicializar resoluciones y estado en session_state
        if 'resoluciones_problemas' not in st.session_state:
            st.session_state.resoluciones_problemas = {}
        if 'correcciones_estado' not in st.session_state:
            st.session_state.correcciones_estado = 'pendiente'
        
        # Si no hay problemas, marcar como aplicadas automáticamente
        if not problemas:
            st.session_state.correcciones_estado = 'aplicadas'
        else:
            # Si hay problemas nuevos (keys que no están en las resoluciones guardadas),
            # volver a modo pendiente para que el usuario los resuelva
            keys_actuales = {p['key'] for p in problemas}
            keys_resueltas = set(st.session_state.resoluciones_problemas.keys())
            keys_nuevas = keys_actuales - keys_resueltas
            if keys_nuevas:
                st.session_state.correcciones_estado = 'pendiente'
            # También limpiar resoluciones obsoletas (de problemas que ya no existen)
            keys_obsoletas = keys_resueltas - keys_actuales
            for k in keys_obsoletas:
                del st.session_state.resoluciones_problemas[k]
        
        # ── Panel de resolución de turnos problemáticos ──
        if problemas:
            st.subheader("🚨 Turnos no codificados")
            
            if st.session_state.correcciones_estado == 'pendiente':
                # ─── Modo edición: mostrar panel y bloquear descarga ───
                st.error(f"Se encontraron **{len(problemas)}** celdas con turnos que no existen en la base maestra de BUK. Debes resolver cada una antes de descargar.")
                
                siglas_disponibles = sorted(set(mapa_siglas.values()))
                with st.expander(f"📖 Ver siglas disponibles en BUK ({len(siglas_disponibles)})"):
                    st.write(", ".join(f"`{s}`" for s in siglas_disponibles))
                    st.write("Siglas adicionales sin horario: `D` (Descanso), `F` (Festivo), `L` (Licencia/Libre), `P` (Permiso), `V` (Vacación), `C` (Compensado)")
                
                st.markdown("---")
                
                for i, p in enumerate(problemas):
                    key = p['key']
                    st.markdown(f"**{i+1}. {p['nombre']}** · {p['fecha_display']} · `{p['rol']}`")
                    st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;Turno reportado por el supervisor: **`{p['turno_raw']}`**")
                    
                    st.radio(
                        "Acción:",
                        options=[
                            "🔧 (1) Actualizar la base de turnos en BUK [recomendado — no modifica este archivo]",
                            "🚫 (2) Omitir este colaborador del archivo final",
                            "✍️ (3) Asignar una sigla manualmente para esta celda",
                        ],
                        key=f"accion_{key}",
                        index=0,
                    )
                    
                    st.text_input(
                        "Sigla manual (solo si elegiste opción 3):",
                        key=f"sigla_{key}",
                        placeholder="Ej: ANFDIU1, L, D"
                    )
                    
                    st.markdown("---")
                
                # Botón de aplicar (gate)
                if st.button("✅ Aplicar correcciones y continuar", type="primary", use_container_width=True):
                    # Leer todos los estados de los widgets y persistir decisiones
                    pendientes_msg = []
                    nuevas_resoluciones = {}
                    
                    for p in problemas:
                        key = p['key']
                        accion_val = st.session_state.get(f"accion_{key}", "")
                        sigla_val = str(st.session_state.get(f"sigla_{key}", "")).strip().upper()
                        
                        if "(3)" in accion_val:
                            if sigla_val:
                                nuevas_resoluciones[key] = {'tipo': 'manual', 'sigla': sigla_val}
                            else:
                                pendientes_msg.append(f"{p['nombre']} · {p['fecha_display']}")
                        elif "(2)" in accion_val:
                            nuevas_resoluciones[key] = {'tipo': 'omitir'}
                        else:
                            nuevas_resoluciones[key] = {'tipo': 'bdmaestra'}
                    
                    if pendientes_msg:
                        st.warning(f"⏸️ Elegiste opción (3) pero no escribiste sigla para: **{', '.join(pendientes_msg)}**")
                    else:
                        st.session_state.resoluciones_problemas = nuevas_resoluciones
                        st.session_state.correcciones_estado = 'aplicadas'
                        st.rerun()
                
                # Bloquear el resto del flujo
                st.stop()
            
            else:
                # ─── Modo aplicadas: mostrar resumen y permitir re-editar ───
                resumen_manual = []
                resumen_omitir = []
                resumen_bdm = []
                
                for p in problemas:
                    res = st.session_state.resoluciones_problemas.get(p['key'], {'tipo': 'bdmaestra'})
                    tipo = res['tipo']
                    linea = f"**{p['nombre']}** · {p['fecha_display']} · `{p['turno_raw']}`"
                    if tipo == 'manual':
                        resumen_manual.append(f"{linea} → **`{res['sigla']}`**")
                    elif tipo == 'omitir':
                        resumen_omitir.append(linea)
                    else:
                        resumen_bdm.append(linea)
                
                st.success(f"✅ {len(problemas)} correcciones aplicadas")
                
                if resumen_manual:
                    st.markdown("**✍️ Sigla manual asignada:**")
                    for l in resumen_manual:
                        st.markdown(f"- {l}")
                if resumen_omitir:
                    st.markdown("**🚫 Colaborador omitido del archivo:**")
                    for l in resumen_omitir:
                        st.markdown(f"- {l}")
                if resumen_bdm:
                    st.markdown("**🔧 Se mantiene como `REVISAR:...` (actualizar base BUK manualmente):**")
                    for l in resumen_bdm:
                        st.markdown(f"- {l}")
                
                if st.button("✏️ Editar correcciones"):
                    st.session_state.correcciones_estado = 'pendiente'
                    st.rerun()
                
                st.divider()
        
        # ── Aplicar resoluciones a df_output (siempre, si estado='aplicadas') ──
        ruts_omitidos_por_problema = set()
        for p in problemas:
            res = st.session_state.resoluciones_problemas.get(p['key'], {'tipo': 'bdmaestra'})
            if res['tipo'] == 'manual':
                df_output.at[p['idx'], p['col']] = res['sigla']
            elif res['tipo'] == 'omitir':
                ruts_omitidos_por_problema.add(p['rut'])
            # 'bdmaestra' → REVISAR: se queda en la celda
        
        # ── Análisis de estado por colaborador ──
        # Clasifica cada fila como: OK / sin datos / con errores (REVISAR)
        estado_filas = []
        for idx, row_buk in df_output.iterrows():
            rut = row_buk['RUT']
            nombre = row_buk['Nombre del Colaborador']
            area = row_buk.get('Área', '')
            supervisor = row_buk.get('Supervisor', '')
            turnos_colab = df_con_match[df_con_match['RUT'] == rut]
            
            # Contar errores REVISAR en la fila del output
            n_revisar = 0
            for fi, cb in fechas_buk.items():
                val = df_output.at[idx, cb]
                if isinstance(val, str) and val.startswith('REVISAR'):
                    n_revisar += 1
            
            if turnos_colab.empty:
                estado = "⚠️ Sin datos 360"
                motivo = "no tiene registros en el archivo 360"
            elif n_revisar > 0:
                estado = f"🔴 {n_revisar} turnos con error"
                motivo = f"{n_revisar} celdas con formato no reconocido"
            else:
                estado = "✅ OK"
                motivo = f"{len(turnos_colab)} turnos cargados correctamente"
            
            estado_filas.append({
                'idx_original': idx,
                'RUT': rut,
                'Nombre': nombre,
                'Área': area,
                'Supervisor': supervisor,
                'Estado': estado,
                'Detalle': motivo,
            })
        
        df_estado = pd.DataFrame(estado_filas)
        
        # ── Panel de revisión y exclusión ──
        st.subheader("📋 Revisión de Colaboradores")
        st.caption("Marca los colaboradores que quieres **excluir** del archivo final. BUK permite reducir filas pero no columnas.")
        
        # Contadores rápidos
        n_ok = (df_estado['Estado'].str.startswith('✅')).sum()
        n_sin = (df_estado['Estado'].str.startswith('⚠️')).sum()
        n_err = (df_estado['Estado'].str.startswith('🔴')).sum()
        col_st1, col_st2, col_st3 = st.columns(3)
        col_st1.metric("✅ OK", int(n_ok))
        col_st2.metric("⚠️ Sin datos 360", int(n_sin))
        col_st3.metric("🔴 Con errores", int(n_err))
        
        # Inicializar exclusiones en session_state
        if 'ruts_excluidos' not in st.session_state:
            st.session_state.ruts_excluidos = set()
        
        # Botones de acción rápida
        col_b1, col_b2, col_b3 = st.columns(3)
        if col_b1.button("Excluir 'Sin datos 360'"):
            sin_datos = df_estado[df_estado['Estado'].str.startswith('⚠️')]['RUT'].tolist()
            st.session_state.ruts_excluidos.update(sin_datos)
            st.rerun()
        if col_b2.button("Excluir 'Con errores'"):
            con_err = df_estado[df_estado['Estado'].str.startswith('🔴')]['RUT'].tolist()
            st.session_state.ruts_excluidos.update(con_err)
            st.rerun()
        if col_b3.button("Incluir todos"):
            st.session_state.ruts_excluidos = set()
            st.rerun()
        
        # Tabla editable con data_editor
        df_tabla = df_estado[['RUT', 'Nombre', 'Área', 'Supervisor', 'Estado', 'Detalle']].copy()
        df_tabla.insert(0, 'Excluir', df_tabla['RUT'].isin(st.session_state.ruts_excluidos))
        
        df_editada = st.data_editor(
            df_tabla,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Excluir': st.column_config.CheckboxColumn('❌ Excluir', default=False, width="small"),
                'RUT': st.column_config.TextColumn('RUT', disabled=True, width="small"),
                'Nombre': st.column_config.TextColumn('Nombre', disabled=True),
                'Área': st.column_config.TextColumn('Área', disabled=True, width="small"),
                'Supervisor': st.column_config.TextColumn('Supervisor', disabled=True),
                'Estado': st.column_config.TextColumn('Estado', disabled=True, width="small"),
                'Detalle': st.column_config.TextColumn('Detalle', disabled=True),
            },
            key='editor_exclusion'
        )
        
        # Actualizar exclusiones desde el editor
        ruts_excluidos_manual = set(df_editada[df_editada['Excluir']]['RUT'].tolist())
        st.session_state.ruts_excluidos = ruts_excluidos_manual
        
        # El set final = exclusiones del editor ∪ omisiones del panel de problemas
        ruts_excluidos_final = ruts_excluidos_manual | ruts_omitidos_por_problema
        
        if ruts_omitidos_por_problema:
            st.caption(f"ℹ️ Además de los marcados arriba, se excluirán **{len(ruts_omitidos_por_problema)}** colaboradores por decisión del panel de turnos no codificados.")
        
        if ruts_excluidos_final:
            st.info(f"🗑️ Se excluirán **{len(ruts_excluidos_final)}** colaboradores del archivo final. Quedarán **{len(df_output) - len(ruts_excluidos_final)}** filas.")
        
        # Filtrar el df_output real
        df_output_final = df_output[~df_output['RUT'].isin(ruts_excluidos_final)].copy().reset_index(drop=True)
        
        st.divider()
        
        # ── Vista Previa ──
        st.subheader("Vista Previa del Archivo Final")
        
        # Mostrar solo las primeras columnas y algunas fechas
        cols_preview = ['Nombre del Colaborador', 'RUT']
        fecha_cols = [c for c in header_buk if c not in ['Nombre del Colaborador', 'RUT', 'Área', 'Supervisor'] and c is not None and not pd.isna(c)]
        cols_preview.extend(fecha_cols[:10])
        cols_existentes = [c for c in cols_preview if c in df_output_final.columns]
        st.dataframe(df_output_final[cols_existentes].head(15), use_container_width=True)
        
        # ── Alertas ──
        if ultima_col_buk:
            st.info(f"📌 Último día del importador (**{ultima_col_buk}**) → **D** (Descanso) para todos. Los días sin turno asignado → **L** (Libre).")
        
        nombres_sin_match = set(df_all['Nombre_Input'].unique()) - set(mapa_nombres.keys())
        if nombres_sin_match:
            st.warning(f"⚠️ {len(nombres_sin_match)} nombres omitidos (sin match BUK): {', '.join(nombres_sin_match)}")
        
        # ── Estadísticas ──
        total_celdas = len(df_con_match)
        celdas_ok = df_con_match['Sigla'].notna().sum()
        celdas_revisar = df_con_match['Sigla'].astype(str).str.startswith('REVISAR').sum()
        
        col_s1, col_s2, col_s3 = st.columns(3)
        col_s1.metric("Total turnos procesados", total_celdas)
        col_s2.metric("Codificados OK", int(celdas_ok - celdas_revisar))
        col_s3.metric("Por revisar", int(celdas_revisar))
        
        # ── Generar archivo de salida ──
        # Escribir como .xls
        output = io.BytesIO()
        try:
            # Intentar con xlwt (formato .xls nativo)
            import xlwt
            wb = xlwt.Workbook()
            
            # ── Hoja 1: turnosColaboradores (con datos modificados) ──
            ws1 = wb.add_sheet('turnosColaboradores')
            for j, col_name in enumerate(header_buk):
                ws1.write(0, j, col_name if col_name is not None else '')
            
            for i, (_, row) in enumerate(df_output_final.iterrows()):
                for j, col_name in enumerate(header_buk):
                    val = row.get(col_name, '')
                    if pd.isna(val):
                        val = ''
                    ws1.write(i + 1, j, str(val) if val != '' else '')
            
            # ── Hoja 2: turnosSemanales (copiar tal cual del original) ──
            # Re-leer el BUK original
            archivo_buk_bytes = st.session_state.buk_bytes
            buk_io = io.BytesIO(archivo_buk_bytes)
            buk_engine = 'xlrd' if st.session_state.get('buk_is_xls', False) else 'openpyxl'
            xls_buk_re = pd.ExcelFile(buk_io, engine=buk_engine)
            
            hojas_copiar = ['turnosSemanales', 'turnosFlexibles', 'turnosTransitorios']
            for nombre_hoja in hojas_copiar:
                try:
                    df_hoja = pd.read_excel(xls_buk_re, sheet_name=nombre_hoja, header=None)
                    ws = wb.add_sheet(nombre_hoja)
                    for i in range(len(df_hoja)):
                        for j in range(len(df_hoja.columns)):
                            val = df_hoja.iloc[i, j]
                            if pd.isna(val):
                                ws.write(i, j, '')
                            else:
                                ws.write(i, j, str(val))
                except Exception as e_h:
                    st.warning(f"No se pudo copiar la hoja '{nombre_hoja}': {e_h}")
            
            wb.save(output)
            formato_salida = "xls"
            
        except ImportError:
            # Fallback: usar openpyxl para xlsx
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_output_final.to_excel(writer, index=False, sheet_name='turnosColaboradores')
                
                # Copiar otras hojas
                archivo_buk_bytes = st.session_state.buk_bytes
                buk_io = io.BytesIO(archivo_buk_bytes)
                buk_engine = 'xlrd' if st.session_state.get('buk_is_xls', False) else 'openpyxl'
                xls_buk_re = pd.ExcelFile(buk_io, engine=buk_engine)
                
                for nombre_hoja in ['turnosSemanales', 'turnosFlexibles', 'turnosTransitorios']:
                    try:
                        df_hoja = pd.read_excel(xls_buk_re, sheet_name=nombre_hoja, header=None)
                        df_hoja.to_excel(writer, index=False, header=False, sheet_name=nombre_hoja)
                    except:
                        pass
            
            formato_salida = "xlsx"
        
        # ── Botones de descarga ──
        st.divider()
        col_d1, col_d2 = st.columns(2)
        
        col_d1.download_button(
            label=f"📥 Descargar Importador BUK (.{formato_salida})",
            data=output.getvalue(),
            file_name=f"Importador_BUK_Cargado.{formato_salida}",
            mime="application/vnd.ms-excel",
            type="primary"
        )
        
        if col_d2.button("🔄 Comenzar de nuevo"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
    
    except Exception as e:
        st.error(f"Error en generación: {e}")
        import traceback
        st.code(traceback.format_exc())
        
        if st.button("🔄 Reiniciar"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

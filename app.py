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
    
    # Determinar rol a partir del nombre de la hoja
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
    
    registros = []
    for i in range(fila_data, len(df)):
        nombre = df.iloc[i, 0]
        if pd.isna(nombre):
            continue
        nombre_str = str(nombre).strip()
        if nombre_str in ['.', '', 'nan', 'NaN']:
            continue
        
        for col_idx, fecha_str in fechas.items():
            turno_raw = df.iloc[i, col_idx] if col_idx < df.shape[1] else None
            registros.append({
                'Nombre_Input': nombre_str,
                'Fecha': fecha_str,
                'Turno_Raw': turno_raw,
                'Rol': rol
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
    
    # Detectar meses disponibles
    try:
        xls360 = pd.ExcelFile(archivo_360)
        hojas = xls360.sheet_names
        
        # Extraer meses únicos de los nombres de hojas
        meses_detectados = set()
        for h in hojas:
            partes = h.split()
            if len(partes) >= 2:
                meses_detectados.add(partes[-1])  # "Abril", "Marzo", etc.
        
        meses_detectados = sorted(meses_detectados)
        
        if not meses_detectados:
            st.error("No se detectaron meses en los nombres de hojas del archivo 360.")
            st.stop()
        
        mes_seleccionado = st.selectbox("📅 Selecciona el mes a procesar:", meses_detectados)
        
        hojas_del_mes = [h for h in hojas if mes_seleccionado.lower() in h.lower()]
        
        if hojas_del_mes:
            st.write(f"Se procesarán las hojas: **{', '.join(hojas_del_mes)}**")
        
        if st.button("🔍 Analizar y Procesar", type="primary"):
            with st.spinner("Leyendo y procesando datos..."):
                
                # ── LEER IMPORTADOR BUK ──
                st.session_state.buk_bytes = archivo_buk.read()
                archivo_buk.seek(0)
                
                # Intentar leer con el motor adecuado
                # .xls → xlrd, .xlsx → openpyxl
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
                
                # Crear mapa nombre→RUT
                nombre_a_rut = dict(zip(nombres_buk, ruts_buk))
                st.session_state.nombre_a_rut = nombre_a_rut
                
                # Hoja turnosSemanales (codificación)
                df_ts_raw = pd.read_excel(xls_buk, sheet_name='turnosSemanales', header=None)
                mapa_siglas = construir_mapa_siglas(df_ts_raw)
                st.session_state.mapa_siglas = mapa_siglas
                
                # ── LEER TURNOS 360 ──
                all_turnos = []
                for hoja in hojas_del_mes:
                    df_hoja = pd.read_excel(xls360, sheet_name=hoja, header=None)
                    df_parsed = parsear_hoja_turnos(df_hoja, hoja)
                    if not df_parsed.empty:
                        all_turnos.append(df_parsed)
                
                if not all_turnos:
                    st.error("No se pudieron parsear turnos de las hojas seleccionadas.")
                    st.stop()
                
                df_all = pd.concat(all_turnos, ignore_index=True)
                st.session_state.df_all_turnos = df_all
                st.session_state.hojas_mes = hojas_del_mes
                
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
        
        # ── Mostrar Preview ──
        st.subheader("Vista Previa")
        
        # Mostrar solo las primeras columnas y algunas fechas
        cols_preview = ['Nombre del Colaborador', 'RUT']
        fecha_cols = [c for c in header_buk if c not in ['Nombre del Colaborador', 'RUT', 'Área', 'Supervisor'] and c is not None and not pd.isna(c)]
        cols_preview.extend(fecha_cols[:10])
        cols_existentes = [c for c in cols_preview if c in df_output.columns]
        st.dataframe(df_output[cols_existentes].head(15), use_container_width=True)
        
        # ── Alertas ──
        if ultima_col_buk:
            st.info(f"📌 Último día del importador (**{ultima_col_buk}**) → **D** (Descanso) para todos. Los días sin turno asignado → **L** (Libre).")
        
        if turnos_no_encontrados:
            st.error(f"🚨 {len(turnos_no_encontrados)} formatos de turno no se pudieron codificar:")
            for t in sorted(turnos_no_encontrados):
                st.write(f"  • `{t}`")
            st.write("Estos aparecen como 'REVISAR:...' en el archivo.")
        
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
        # Reconstruir con header
        df_final = pd.DataFrame([header_buk], columns=header_buk)
        df_final = pd.concat([df_final, df_output], ignore_index=True)
        
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
            
            for i, (_, row) in enumerate(df_output.iterrows()):
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
                df_output.to_excel(writer, index=False, sheet_name='turnosColaboradores')
                
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

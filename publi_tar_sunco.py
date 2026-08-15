import streamlit as st
import pandas as pd
import base64
from pathlib import Path

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Tarifas – SUNCOLOMBIA GENERACIÓN",
    page_icon="☀️",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* ---- Header ---- */
.header-hero {
    text-align: center;
    padding: 2.5rem 1rem 1.5rem 1rem;
}
.header-hero h1 {
    font-size: 2.4rem;
    font-weight: 700;
    color: #1a1a2e;
    margin-bottom: 0.4rem;
}
.header-hero h1 span {
    color: #7c3aed;
}
.header-hero p {
    font-size: 1.05rem;
    color: #6b7280;
    font-weight: 400;
}

/* ---- Logo bar ---- */
.logo-bar {
    display: flex;
    justify-content: center;
    align-items: center;
    margin-bottom: 0.5rem;
}

/* ---- Filter row ---- */
.filter-label {
    font-size: 0.75rem;
    font-weight: 700;
    color: #7c3aed;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 0.25rem;
}

/* ---- Apply button ---- */
div.stButton > button {
    background: #7c3aed;
    color: white;
    font-weight: 600;
    font-size: 1rem;
    border: none;
    border-radius: 8px;
    padding: 0.7rem 2rem;
    width: 100%;
    transition: background 0.2s;
}
div.stButton > button:hover {
    background: #6d28d9;
    color: white;
}

/* ---- Metric cards ---- */
.metric-section {
    background: #fff;
    border: 1.5px solid #e5e7eb;
    border-radius: 16px;
    padding: 1.4rem 1.6rem 1rem 1.6rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 1px 6px rgba(124,58,237,0.06);
}
.metric-section h3 {
    font-size: 0.8rem;
    font-weight: 700;
    color: #7c3aed;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 1rem;
    border-bottom: 2px solid #ede9fe;
    padding-bottom: 0.5rem;
}
.metric-grid {
    display: grid;
    gap: 0.8rem;
}
.metric-pill {
    display: inline-block;
    background: #7c3aed;
    color: white;
    font-weight: 600;
    font-size: 0.78rem;
    border-radius: 20px;
    padding: 0.35rem 0.9rem;
    text-align: center;
    width: 100%;
    box-sizing: border-box;
}
.metric-value {
    font-size: 1.35rem;
    font-weight: 700;
    color: #1a1a2e;
    text-align: center;
    margin-top: 0.2rem;
}
.metric-unit {
    font-size: 0.72rem;
    color: #9ca3af;
    text-align: center;
    margin-bottom: 0.4rem;
}
.metric-card {
    text-align: center;
    padding: 0.5rem 0.2rem;
}
.card-container {
    border: 1px solid #ede9fe;
    border-radius: 12px;
    padding: 0.6rem;
}

/* ---- Divider ---- */
hr.purple-divider {
    border: none;
    border-top: 2px solid #ede9fe;
    margin: 1.5rem 0;
}

/* ---- Footer ---- */
.footer-box {
    background: #f5f3ff;
    border-left: 4px solid #7c3aed;
    border-radius: 10px;
    padding: 1.2rem 1.6rem;
    margin-top: 2rem;
    font-size: 0.88rem;
    color: #374151;
    line-height: 1.7;
}
.footer-box strong {
    color: #5b21b6;
}
.footer-bottom {
    text-align: center;
    font-size: 0.78rem;
    color: #9ca3af;
    margin-top: 1.2rem;
    padding-bottom: 1rem;
}

/* ---- No-results ---- */
.no-result {
    text-align: center;
    color: #9ca3af;
    font-size: 1rem;
    padding: 2.5rem 0;
}

/* ---- Data table ---- */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
}

/* ---- Selectbox label override ---- */
.stSelectbox label {
    font-size: 0.75rem !important;
    font-weight: 700 !important;
    color: #7c3aed !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
</style>
""", unsafe_allow_html=True)

# ── Logo helper ───────────────────────────────────────────────────────────────
def img_to_b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

logo_path = Path("suncolombia_logo.png")
if logo_path.exists():
    logo_b64 = img_to_b64(logo_path)
    st.markdown(f"""
    <div class="logo-bar">
        <img src="data:image/png;base64,{logo_b64}" height="60" alt="SUNCOLOMBIA GENERACIÓN">
    </div>
    """, unsafe_allow_html=True)

# ── Hero title ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-hero">
    <h1>Conoce <span>nuestras tarifas</span> de servicio de energía sostenible</h1>
    <p>Información detallada y accesible para usuarios en Zonas No Interconectadas (ZNI)</p>
</div>
""", unsafe_allow_html=True)

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    xl = pd.ExcelFile("Base_Tarifas_Publicacion.xlsx")
    dfs = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df.columns = [str(c).strip() for c in df.columns]
        # Normalize text columns
        for col in ["MES", "MUNICIPIO", "TIPO DE SISTEMA"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()
        if "AÑO" in df.columns:
            df["AÑO"] = pd.to_numeric(df["AÑO"], errors="coerce").astype("Int64")
        dfs[sheet] = df
    return dfs

MES_ORDER = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
             "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

data = load_data()
sheet_names = list(data.keys())   # e.g. ["CREG 166", "CREG 101 026"]

# ── Sidebar / filter row ──────────────────────────────────────────────────────
st.markdown("---")
col_res, col_yr, col_mo, col_mun, col_btn = st.columns([2, 1.5, 1.5, 2, 1.5])

with col_res:
    st.markdown('<p class="filter-label">Resolución CREG</p>', unsafe_allow_html=True)
    resolucion = st.selectbox("Resolución CREG", options=["Todas"] + sheet_names,
                               label_visibility="collapsed", key="res")

# Build combined unique values for filters
def get_options(field):
    vals = set()
    for sn, df in data.items():
        if resolucion != "Todas" and sn != resolucion:
            continue
        if field in df.columns:
            vals.update(df[field].dropna().astype(str).unique())
    return sorted(vals)

años = get_options("AÑO")
municipios = get_options("MUNICIPIO")
meses_raw = get_options("MES")
meses = [m for m in MES_ORDER if m in meses_raw] + [m for m in meses_raw if m not in MES_ORDER]

with col_yr:
    st.markdown('<p class="filter-label">Año</p>', unsafe_allow_html=True)
    año_sel = st.selectbox("Año", ["Todos"] + años, label_visibility="collapsed", key="yr")

with col_mo:
    st.markdown('<p class="filter-label">Mes</p>', unsafe_allow_html=True)
    mes_sel = st.selectbox("Mes", ["Todos"] + meses, label_visibility="collapsed", key="mo")

with col_mun:
    st.markdown('<p class="filter-label">Municipio</p>', unsafe_allow_html=True)
    mun_sel = st.selectbox("Municipio", ["Todos"] + municipios, label_visibility="collapsed", key="mun")

with col_btn:
    st.markdown('<p class="filter-label">&nbsp;</p>', unsafe_allow_html=True)
    aplicar = st.button("Aplicar filtros", key="apply")

# ── Filter data ───────────────────────────────────────────────────────────────
def filter_df(df):
    out = df.copy()
    if año_sel != "Todos" and "AÑO" in out.columns:
        out = out[out["AÑO"].astype(str) == str(año_sel)]
    if mes_sel != "Todos" and "MES" in out.columns:
        out = out[out["MES"] == mes_sel]
    if mun_sel != "Todos" and "MUNICIPIO" in out.columns:
        out = out[out["MUNICIPIO"] == mun_sel]
    return out

sheets_to_show = sheet_names if resolucion == "Todas" else [resolucion]

# ── Helpers: render sections ──────────────────────────────────────────────────
def fmt(val, decimals=2):
    try:
        return f"{float(val):,.{decimals}f}"
    except Exception:
        return str(val)

def render_creg166(df):
    cols_map = {
        "CO_DIA":       ("CO_DIA",       "$/día",  "Costo Operación"),
        "GAOM_DIA":     ("GAOM_DIA",     "$/día",  "G+A+O+M"),
        "CM":           ("CM",           "$",      "Costo Medio"),
        "GM":           ("GM",           "$",      "Gasto Medio"),
        "CU_DIA":       ("CU_DIA",       "$/día",  "Costo Unitario"),
        "SUBSIDIO_DIA": ("SUBSIDIO_DIA", "$/día",  "Subsidio"),
        "TARIFA_DIA":   ("TARIFA_DIA",   "$/día",  "Tarifa"),
    }

    st.markdown('<div class="metric-section"><h3>⚡ Componentes CREG 166 de 2020</h3>', unsafe_allow_html=True)
    if df.empty:
        st.markdown('<p class="no-result">Sin datos para los filtros seleccionados.</p>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        return

    row = df.iloc[0]
    grid_cols = st.columns(len(cols_map))
    for i, (col_key, (field, unit, label)) in enumerate(cols_map.items()):
        with grid_cols[i]:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-pill">{label}</div>
                <div class="metric-value">{fmt(row.get(field, '---'))}</div>
                <div class="metric-unit">{unit}</div>
            </div>
            """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if len(df) > 1:
        with st.expander("Ver tabla completa – CREG 166"):
            st.dataframe(df.reset_index(drop=True), use_container_width=True)


def render_creg101(df):
    cols_map = {
        "WHD":      ("WHD",      "Wh/día", "WHD"),
        "AMGCnu_m": ("AMGCnu_m", "$/kWh",  "AMGCnu_m"),
        "AMGCvi_m": ("AMGCvi_m", "$/kWh",  "AMGCvi_m"),
        "AMGCau_m": ("AMGCau_m", "$/kWh",  "AMGCau_m"),
        "AMGCnf_m": ("AMGCnf_m", "$/kWh",  "AMGCnf_m"),
        "AMGCro_m": ("AMGCro_m", "$/kWh",  "AMGCro_m"),
    }
    cols_map2 = {
        "INVERSION": ("INVERSION",   "$",      "Inversión"),
        "AMGCm":     ("AMGCm",       "$/kWh",  "AMGC_m"),
        "Empresa SIN": ("Empresa SIN", "",     "Empresa SIN"),
        "Tarifa SIN":  ("Tarifa SIN",  "$/kWh", "Tarifa SIN"),
    }

    st.markdown('<div class="metric-section"><h3>☀️ Componentes CREG 101-026 de 2022</h3>', unsafe_allow_html=True)
    if df.empty:
        st.markdown('<p class="no-result">Sin datos para los filtros seleccionados.</p>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        return

    row = df.iloc[0]

    # Row 1
    grid1 = st.columns(len(cols_map))
    for i, (col_key, (field, unit, label)) in enumerate(cols_map.items()):
        with grid1[i]:
            val = row.get(field, "---")
            display = fmt(val, 0) if field == "WHD" else fmt(val)
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-pill">{label}</div>
                <div class="metric-value">{display}</div>
                <div class="metric-unit">{unit}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Row 2
    grid2 = st.columns(len(cols_map2))
    for i, (col_key, (field, unit, label)) in enumerate(cols_map2.items()):
        with grid2[i]:
            val = row.get(field, "---")
            display = val if field == "Empresa SIN" else fmt(val)
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-pill">{label}</div>
                <div class="metric-value" style="font-size:1.1rem">{display}</div>
                <div class="metric-unit">{unit}</div>
            </div>
            """, unsafe_allow_html=True)

    # Tipo de sistema badge
    if "TIPO DE SISTEMA" in row.index and str(row["TIPO DE SISTEMA"]) not in ("", "nan", "NAN"):
        st.markdown(f"""
        <p style="text-align:center; margin-top:0.8rem; color:#6b7280; font-size:0.85rem;">
            Tipo de sistema: <strong style="color:#7c3aed">{row['TIPO DE SISTEMA']}</strong>
        </p>
        """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    if len(df) > 1:
        with st.expander("Ver tabla completa – CREG 101-026"):
            st.dataframe(df.reset_index(drop=True), use_container_width=True)


# ── Context info ──────────────────────────────────────────────────────────────
def context_bar(df):
    parts = []
    if año_sel != "Todos":
        parts.append(f"**Año:** {año_sel}")
    if mes_sel != "Todos":
        parts.append(f"**Mes:** {mes_sel.capitalize()}")
    if mun_sel != "Todos":
        parts.append(f"**Municipio:** {mun_sel.capitalize()}")
    if parts:
        label = " &nbsp;|&nbsp; ".join(parts)
        st.markdown(f"""
        <div style="background:#f5f3ff;border-radius:8px;padding:0.5rem 1rem;
                    font-size:0.88rem;color:#374151;margin-bottom:1rem;">
            📍 {label} &nbsp;·&nbsp; <span style="color:#7c3aed">{len(df)} registro(s) encontrado(s)</span>
        </div>""", unsafe_allow_html=True)


# ── Render results ────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)

for sn in sheets_to_show:
    df_raw = data[sn]
    df_f = filter_df(df_raw)

    context_bar(df_f)

    if "CREG 166" in sn.upper():
        render_creg166(df_f)
    elif "101" in sn or "026" in sn:
        render_creg101(df_f)
    else:
        st.subheader(sn)
        st.dataframe(df_f.reset_index(drop=True), use_container_width=True)

# ── Regulatory footer ─────────────────────────────────────────────────────────
st.markdown("""
<div class="footer-box">
    <strong>📋 Marco regulatorio aplicable</strong><br><br>
    Las tarifas de SUNCOLOMBIA GENERACIÓN SAS ESP fueron calculadas inicialmente bajo la
    <strong>Resolución CREG 166 de 2020</strong>, que estableció la metodología tarifaria
    para los Sistemas Individuales de Servicio Público de Energía Eléctrica (SISFV) en Zonas
    No Interconectadas (ZNI). Esta resolución definía componentes como el Costo de Operación
    (CO), los Gastos de Administración, Operación y Mantenimiento (GAOM), el Costo Medio (CM),
    el Gasto Medio (GM), el Costo Unitario Diario (CU_DIA) y la Tarifa Diaria, junto con el
    subsidio aplicable a los estratos 1, 2 y 3.<br><br>
    A partir de <strong>noviembre de 2023</strong>, la <strong>Resolución CREG 101-026 de 2022</strong>
    derogó la CREG 166 de 2020 e introdujo una nueva metodología tarifaria para los SISFV en ZNI.
    La nueva resolución incorpora componentes como el WHD (Watios Hora Diarios), los Años de Marca
    de Garantía del Componente (AMGC) diferenciados por tipo (nuevos, vida útil, actualización,
    no funcionales y reposición), y establece la comparación con la Tarifa SIN del operador de red
    interconectado de referencia, garantizando condiciones tarifarias equitativas para los usuarios
    de estas zonas rurales.
</div>
<div class="footer-bottom">
    SUNCOLOMBIA GENERACIÓN SAS ESP &nbsp;·&nbsp; Información tarifaria ZNI &nbsp;·&nbsp;
    Resoluciones CREG 166/2020 y CREG 101-026/2022
</div>
""", unsafe_allow_html=True)

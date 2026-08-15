import streamlit as st
import pandas as pd
import base64
import traceback
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Tarifas – SUNCOLOMBIA GENERACIÓN",
    page_icon="☀️",
    layout="wide",
)

BASE_DIR = Path(__file__).parent


# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.hero { text-align:center; padding:2rem 1rem 1rem; }
.hero h1 { font-size:2.2rem; font-weight:700; color:#1a1a2e; }
.hero h1 span { color:#7c3aed; }
.hero p { color:#6b7280; font-size:1rem; margin-top:0.3rem; }
.pill {
    display:block; background:#7c3aed; color:#fff;
    font-weight:600; font-size:0.75rem; border-radius:20px;
    padding:0.3rem 0.6rem; text-align:center; margin-bottom:0.25rem;
}
.val { font-size:1.25rem; font-weight:700; color:#1a1a2e; text-align:center; }
.unit { font-size:0.7rem; color:#9ca3af; text-align:center; margin-bottom:0.5rem; }
.card-box {
    border:1.5px solid #ede9fe; border-radius:14px;
    padding:1rem 1.2rem; margin-bottom:1.2rem;
    background:#fff; box-shadow:0 1px 6px rgba(124,58,237,0.07);
}
.card-title {
    font-size:0.75rem; font-weight:700; color:#7c3aed;
    text-transform:uppercase; letter-spacing:0.07em;
    border-bottom:2px solid #ede9fe; padding-bottom:0.4rem; margin-bottom:1rem;
}
.footer-box {
    background:#f5f3ff; border-left:4px solid #7c3aed;
    border-radius:10px; padding:1.2rem 1.6rem;
    margin-top:2rem; font-size:0.86rem; color:#374151; line-height:1.75;
}
.footer-box strong { color:#5b21b6; }
.footer-bottom { text-align:center; font-size:0.75rem; color:#9ca3af; margin-top:1rem; }
div.stButton > button {
    background:#7c3aed; color:#fff; font-weight:600;
    border:none; border-radius:8px; width:100%; padding:0.65rem;
}
div.stButton > button:hover { background:#6d28d9; color:#fff; }
</style>
""", unsafe_allow_html=True)

# ── Logo ──────────────────────────────────────────────────────────────────────
logo_path = BASE_DIR / "suncolombia_logo.png"
if logo_path.exists():
    try:
        logo_b64 = base64.b64encode(logo_path.read_bytes()).decode()
        st.markdown(
            f'<div style="text-align:center;padding:1rem 0 0.5rem">'
            f'<img src="data:image/png;base64,{logo_b64}" height="55" alt="SUNCOLOMBIA"></div>',
            unsafe_allow_html=True,
        )
    except Exception:
        pass

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>Conoce <span>nuestras tarifas</span> de servicio de energía sostenible</h1>
  <p>Información detallada y accesible para usuarios en Zonas No Interconectadas (ZNI)</p>
</div>
""", unsafe_allow_html=True)

# ── Cargar Excel ──────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    candidates = [
        BASE_DIR / "Base_Tarifas_Publicacion.xlsx",
        BASE_DIR / "Base Tarifas Publicacion.xlsx",
        BASE_DIR / "base_tarifas_publicacion.xlsx",
    ]
    excel_path = next((p for p in candidates if p.exists()), None)
    if excel_path is None:
        return None, f"Excel no encontrado. Archivos: {[p.name for p in BASE_DIR.iterdir()]}"
    try:
        xl = pd.ExcelFile(excel_path)
        dfs = {}
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            df.columns = [str(c).strip() for c in df.columns]
            for col in ["MES", "MUNICIPIO", "TIPO DE SISTEMA"]:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.upper()
            if "AÑO" in df.columns:
                df["AÑO"] = pd.to_numeric(df["AÑO"], errors="coerce").astype("Int64")
            dfs[sheet] = df
        return dfs, None
    except Exception as e:
        return None, f"Error leyendo Excel: {e}\n{traceback.format_exc()}"

data, load_error = load_data()

if load_error:
    st.error(load_error)
    st.stop()

sheet_names = list(data.keys())

# ── Filtros ───────────────────────────────────────────────────────────────────
st.markdown("---")
MES_ORDER = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
             "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

c1, c2, c3, c4, c5 = st.columns([2, 1.5, 1.5, 2, 1.5])

with c1:
    resolucion = st.selectbox("**Resolución CREG**", ["Todas"] + sheet_names, key="res")

def opts(field):
    vals = set()
    for sn, df in data.items():
        if resolucion != "Todas" and sn != resolucion:
            continue
        if field in df.columns:
            vals.update(df[field].dropna().astype(str).unique())
    return sorted(vals)

años      = opts("AÑO")
municipios = opts("MUNICIPIO")
meses_raw  = opts("MES")
meses = [m for m in MES_ORDER if m in meses_raw] + [m for m in meses_raw if m not in MES_ORDER]

with c2:
    año_sel = st.selectbox("**Año**", ["Todos"] + años, key="yr")
with c3:
    mes_sel = st.selectbox("**Mes**", ["Todos"] + meses, key="mo")
with c4:
    mun_sel = st.selectbox("**Municipio**", ["Todos"] + municipios, key="mun")
with c5:
    st.markdown("<br>", unsafe_allow_html=True)
    st.button("Aplicar filtros", key="apply")

# ── Filtrar ───────────────────────────────────────────────────────────────────
def filter_df(df):
    out = df.copy()
    if año_sel != "Todos" and "AÑO" in out.columns:
        out = out[out["AÑO"].astype(str) == str(año_sel)]
    if mes_sel != "Todos" and "MES" in out.columns:
        out = out[out["MES"] == mes_sel]
    if mun_sel != "Todos" and "MUNICIPIO" in out.columns:
        out = out[out["MUNICIPIO"] == mun_sel]
    return out

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(val, dec=2):
    try:
        return f"{float(val):,.{dec}f}"
    except Exception:
        return str(val)

def metric_col(col, label, val, unit):
    with col:
        st.markdown(
            f'<div class="pill">{label}</div>'
            f'<div class="val">{val}</div>'
            f'<div class="unit">{unit}</div>',
            unsafe_allow_html=True,
        )

# ── Render CREG 166 ───────────────────────────────────────────────────────────
def render_166(df):
    st.markdown('<div class="card-box"><div class="card-title">⚡ Componentes CREG 166 de 2020</div>', unsafe_allow_html=True)
    if df.empty:
        st.info("Sin datos para los filtros seleccionados.")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    row = df.iloc[0]
    fields = [
        ("CO_DIA",       "Costo Operación", "$/día"),
        ("GAOM_DIA",     "G+A+O+M",         "$/día"),
        ("CM",           "Costo Medio",      "$"),
        ("GM",           "Gasto Medio",      "$"),
        ("CU_DIA",       "Costo Unitario",   "$/día"),
        ("SUBSIDIO_DIA", "Subsidio",         "$/día"),
        ("TARIFA_DIA",   "Tarifa",           "$/día"),
    ]
    cols = st.columns(len(fields))
    for col, (f, lbl, unit) in zip(cols, fields):
        metric_col(col, lbl, fmt(row.get(f, "---")), unit)
    st.markdown("</div>", unsafe_allow_html=True)
    if len(df) > 1:
        with st.expander("Ver tabla completa – CREG 166"):
            st.dataframe(df.reset_index(drop=True), use_container_width=True)

# ── Render CREG 101-026 ───────────────────────────────────────────────────────
def render_101(df):
    st.markdown('<div class="card-box"><div class="card-title">☀️ Componentes CREG 101-026 de 2022</div>', unsafe_allow_html=True)
    if df.empty:
        st.info("Sin datos para los filtros seleccionados.")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    row = df.iloc[0]
    row1 = [
        ("WHD",      "WHD",      "Wh/día", 0),
        ("AMGCnu_m", "AMGCnu_m", "$/kWh",  2),
        ("AMGCvi_m", "AMGCvi_m", "$/kWh",  2),
        ("AMGCau_m", "AMGCau_m", "$/kWh",  2),
        ("AMGCnf_m", "AMGCnf_m", "$/kWh",  2),
        ("AMGCro_m", "AMGCro_m", "$/kWh",  2),
    ]
    cols1 = st.columns(len(row1))
    for col, (f, lbl, unit, dec) in zip(cols1, row1):
        metric_col(col, lbl, fmt(row.get(f, "---"), dec), unit)
    st.markdown("<br>", unsafe_allow_html=True)
    row2 = [
        ("INVERSION",   "Inversión",   "$",     2),
        ("AMGCm",       "AMGC_m",      "$/kWh", 2),
        ("Empresa SIN", "Empresa SIN", "",      -1),
        ("Tarifa SIN",  "Tarifa SIN",  "$/kWh", 2),
    ]
    cols2 = st.columns(len(row2))
    for col, (f, lbl, unit, dec) in zip(cols2, row2):
        v = str(row.get(f, "---")) if dec == -1 else fmt(row.get(f, "---"), dec)
        metric_col(col, lbl, v, unit)
    if "TIPO DE SISTEMA" in row.index:
        tipo = str(row["TIPO DE SISTEMA"])
        if tipo not in ("", "nan", "NAN"):
            st.markdown(
                f'<p style="text-align:center;color:#6b7280;font-size:0.85rem;margin-top:0.8rem">'
                f'Tipo de sistema: <strong style="color:#7c3aed">{tipo}</strong></p>',
                unsafe_allow_html=True,
            )
    st.markdown("</div>", unsafe_allow_html=True)
    if len(df) > 1:
        with st.expander("Ver tabla completa – CREG 101-026"):
            st.dataframe(df.reset_index(drop=True), use_container_width=True)

# ── Mostrar resultados ────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
sheets_to_show = sheet_names if resolucion == "Todas" else [resolucion]

for sn in sheets_to_show:
    df_f = filter_df(data[sn])
    n = len(df_f)
    if año_sel != "Todos" or mes_sel != "Todos" or mun_sel != "Todos":
        st.caption(f"📍 {n} registro(s) — {sn}")
    if "166" in sn:
        render_166(df_f)
    else:
        render_101(df_f)

# ── Footer regulatorio ────────────────────────────────────────────────────────
st.markdown("""
<div class="footer-box">
  <strong>📋 Marco regulatorio aplicable</strong><br><br>
  Las tarifas de SUNCOLOMBIA GENERACIÓN SAS ESP fueron calculadas inicialmente bajo la
  <strong>Resolución CREG 166 de 2020</strong>, que estableció la metodología tarifaria
  para los Sistemas Individuales de Servicio Público de Energía Eléctrica (SISFV) en Zonas
  No Interconectadas (ZNI). Esta resolución definía componentes como el Costo de Operación (CO),
  los Gastos de Administración, Operación y Mantenimiento (GAOM), el Costo Medio (CM), el Gasto
  Medio (GM), el Costo Unitario Diario (CU_DIA) y la Tarifa Diaria, junto con el subsidio aplicable
  a los estratos 1, 2 y 3.<br><br>
  A partir de <strong>noviembre de 2023</strong>, la <strong>Resolución CREG 101-026 de 2022</strong>
  derogó la CREG 166 de 2020 e introdujo una nueva metodología tarifaria para los SISFV en ZNI.
  La nueva resolución incorpora componentes como el WHD (Watios Hora Diarios), los Años de Marca de
  Garantía del Componente (AMGC) diferenciados por tipo (nuevos, vida útil, actualización, no
  funcionales y reposición), y establece la comparación con la Tarifa SIN del operador de red
  interconectado de referencia, garantizando condiciones tarifarias equitativas para los usuarios
  de estas zonas rurales.
</div>
<div class="footer-bottom">
  SUNCOLOMBIA GENERACIÓN SAS ESP &nbsp;·&nbsp; Información tarifaria ZNI &nbsp;·&nbsp;
  Resoluciones CREG 166/2020 y CREG 101-026/2022
</div>
""", unsafe_allow_html=True)

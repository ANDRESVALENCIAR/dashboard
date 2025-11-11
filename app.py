import pandas as pd
import streamlit as st
from datetime import datetime
import io

# ---------- CONFIGURACIÓN DE PÁGINA ----------
st.set_page_config(
    page_title="Dashboard CEO – Proyectos",
    page_icon="📊",
    layout="wide",
)

# ---------- ESTILOS BÁSICOS ----------
# Un poquito de CSS para que se vea más limpio
st.markdown(
    """
    <style>
    /* Ocultar menú y footer de Streamlit si quieres más limpio */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .big-title {
        font-size: 36px;
        font-weight: 700;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 16px;
        color: #6c757d;
        margin-bottom: 20px;
    }
    .kpi-card {
        padding: 16px;
        border-radius: 12px;
        border: 1px solid #e5e5e5;
        background-color: #ffffff;
        text-align: center;
    }
    .kpi-label {
        font-size: 14px;
        color: #6c757d;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 700;
        margin-top: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- CONFIGURACIÓN DE PESOS ----------
WEIGHTS = {
    "Impacto_ventas": 2.0,
    "Tiempo_impl": 1.5,
    "Facilidad": 1.5,
    "Alineacion_vision": 2.0,
    "Diferenciacion": 1.0,
    "Riesgo_bajo": 1.0,
}

# ---------- FUNCIONES AUXILIARES ----------

def to_number(x):
    """Convierte cosas como '30 dias' o '0,5' a número."""
    if isinstance(x, str):
        digits = "".join(ch for ch in x if ch.isdigit() or ch in [".", ","])
        digits = digits.replace(",", ".")
        try:
            return float(digits) if digits != "" else None
        except Exception:
            return None
    return x

def limpiar_numericos(df):
    numeric_cols = [
        "Impacto_ventas",
        "Tiempo_impl",
        "Facilidad",
        "Alineacion_vision",
        "Diferenciacion",
        "Riesgo_bajo",
    ]

    for col in numeric_cols:
        if col not in df.columns:
            continue

        if col == "Tiempo_impl":
            df[col] = df[col].apply(to_number)

        elif col == "Riesgo_bajo":
            def map_riesgo(v):
                if isinstance(v, str):
                    v_clean = v.strip().lower()
                    if v_clean in ["si", "sí"]:
                        return 5.0
                    if v_clean == "no":
                        return 1.0
                return to_number(v)
            df[col] = df[col].apply(map_riesgo)

        else:
            df[col] = pd.to_numeric(df[col].apply(to_number), errors="coerce")

    return df

def calcular_score(row):
    score = 0.0
    for col, w in WEIGHTS.items():
        if col in row and row[col] is not None and not pd.isna(row[col]):
            score += row[col] * w
    return score

def color_por_proyecto(row, hoy):
    estado = str(row.get("Estado_manual", "")).strip().lower()
    etd = row.get("ETD", None)

    if estado == "completado":
        return "✅ VERDE – Completado"

    if isinstance(etd, str):
        etd_parsed = pd.to_datetime(etd, errors="coerce")
        etd = etd_parsed.date() if pd.notna(etd_parsed) else None
    elif isinstance(etd, pd.Timestamp):
        etd = etd.date()
    elif pd.isna(etd):
        etd = None

    score = row["Score"]

    if etd is None:
        if score >= 30:
            return "🟢 VERDE – Alta prioridad (sin fecha ETD)"
        elif score >= 22:
            return "🟢 Verde media – Importante (sin fecha ETD)"
        else:
            return "⚪ Parking – Baja prioridad (sin fecha ETD)"

    if etd < hoy:
        return "🔴 ROJO – Atrasado"

    dias_restantes = (etd - hoy).days

    if dias_restantes <= 30:
        return "🟠 AMARILLO – Urgente (<30 días)"

    if score >= 30:
        return "🟢 VERDE – Prioridad estratégica"
    elif score >= 22:
        return "🟢 Verde media – Importante"
    else:
        return "⚪ Parking – Baja prioridad"

def procesar_excel(uploaded_file):
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    header_row = None
    for i in range(len(raw)):
        if str(raw.iloc[i, 0]).strip().lower() == "proyecto":
            header_row = i
            break

    if header_row is None:
        df = pd.read_excel(uploaded_file, sheet_name=0)
    else:
        header = raw.iloc[header_row]
        df = raw.iloc[header_row + 1 :].copy()
        df.columns = header

    df = limpiar_numericos(df)
    df["Score"] = df.apply(calcular_score, axis=1)
    hoy = datetime.today().date()
    df["Semaforo"] = df.apply(lambda r: color_por_proyecto(r, hoy), axis=1)
    return df

# ---------- UI PRINCIPAL ----------

st.markdown('<p class="big-title">📊 Dashboard CEO – Tablero de Proyectos</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-title">Sube tu archivo de proyectos en Excel y el sistema te calcula <b>Score</b> y <b>Semáforo</b> para priorizar.</p>',
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx"])

# Sidebar: filtros
st.sidebar.header("🔍 Filtros")
st.sidebar.write("Aplica filtros para enfocarte en lo que importa.")

if uploaded_file is not None:
    df = procesar_excel(uploaded_file)
    df_ordenado = df.sort_values(by="Score", ascending=False)

    # Filtros en sidebar
    estados_unicos = df_ordenado["Semaforo"].unique().tolist()
    filtro_estado = st.sidebar.multiselect(
        "Filtrar por semáforo",
        options=estados_unicos,
        default=estados_unicos,
    )

    if "Dueño" in df_ordenado.columns:
        duenos_unicos = sorted(df_ordenado["Dueño"].dropna().unique().tolist())
        filtro_dueno = st.sidebar.multiselect(
            "Filtrar por dueño",
            options=duenos_unicos,
            default=duenos_unicos,
        )
    else:
        filtro_dueno = []

    df_filtrado = df_ordenado.copy()
    if filtro_estado:
        df_filtrado = df_filtrado[df_filtrado["Semaforo"].isin(filtro_estado)]
    if "Dueño" in df_filtrado.columns and filtro_dueno:
        df_filtrado = df_filtrado[df_filtrado["Dueño"].isin(filtro_dueno)]

    # ---------- KPIs ----------
    total_proj = len(df_ordenado)
    en_rojo = (df_ordenado["Semaforo"].str.contains("ROJO")).sum()
    en_amarillo = (df_ordenado["Semaforo"].str.contains("AMARILLO")).sum()
    en_verde = (df_ordenado["Semaforo"].str.contains("VERDE")).sum()
    en_parking = (df_ordenado["Semaforo"].str.contains("Parking")).sum()

    st.markdown("### Resumen general")

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown('<div class="kpi-card"><div class="kpi-label">Total proyectos</div><div class="kpi-value">{}</div></div>'.format(total_proj), unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="kpi-card"><div class="kpi-label">🔴 Atrasados</div><div class="kpi-value">{}</div></div>'.format(en_rojo), unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="kpi-card"><div class="kpi-label">🟠 Urgentes (&lt;30 días)</div><div class="kpi-value">{}</div></div>'.format(en_amarillo), unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="kpi-card"><div class="kpi-label">🟢 En buena vía</div><div class="kpi-value">{}</div></div>'.format(en_verde), unsafe_allow_html=True)
    with col5:
        st.markdown('<div class="kpi-card"><div class="kpi-label">⚪ Parking</div><div class="kpi-value">{}</div></div>'.format(en_parking), unsafe_allow_html=True)

    st.markdown("---")

    # ---------- GRÁFICA TOP PROYECTOS ----------
    st.subheader("Top proyectos por Score")

    top_chart = df_ordenado.copy()
    if "Proyecto" in top_chart.columns:
        top_chart = top_chart.head(10)[["Proyecto", "Score"]].set_index("Proyecto")
        st.bar_chart(top_chart)

    # ---------- TABLA PRINCIPAL ----------
    st.markdown("### Lista de proyectos (filtrados y ordenados por Score)")
    st.dataframe(df_filtrado, use_container_width=True)

    # ---------- TOP 5 ----------
    st.markdown("---")
    st.subheader("Top 5 proyectos recomendados para foco")

    cols_top = [c for c in ["Proyecto", "Dueño", "Score", "Semaforo", "ETD", "Estado_manual"] if c in df_ordenado.columns]
    st.table(df_ordenado[cols_top].head(5).reset_index(drop=True))

    # ---------- DESCARGA EXCEL ----------
    st.markdown("---")
    st.subheader("Descargar Excel enriquecido")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_ordenado.to_excel(writer, index=False, sheet_name="Proyectos_con_score")
    data_xlsx = output.getvalue()

    st.download_button(
        label="📥 Descargar Excel con Score y Semáforo",
        data=data_xlsx,
        file_name="proyectos_con_score.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("👆 Sube un archivo Excel para ver el tablero.")

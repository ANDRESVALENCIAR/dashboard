import pandas as pd
import streamlit as st
from datetime import datetime
import io

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
    """Limpia columnas numéricas según la lógica de tu archivo."""
    numeric_cols = ["Impacto_ventas", "Tiempo_impl", "Facilidad",
                    "Alineacion_vision", "Diferenciacion", "Riesgo_bajo"]

    for col in numeric_cols:
        if col not in df.columns:
            continue

        if col == "Tiempo_impl":
            df[col] = df[col].apply(to_number)

        elif col == "Riesgo_bajo":
            # Mapear SI/NO a números
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

    # Si ya está completado
    if estado == "completado":
        return "✅ VERDE – Completado"

    # Parsear ETD
    if isinstance(etd, str):
        etd_parsed = pd.to_datetime(etd, errors="coerce")
        etd = etd_parsed.date() if pd.notna(etd_parsed) else None
    elif isinstance(etd, pd.Timestamp):
        etd = etd.date()
    elif pd.isna(etd):
        etd = None

    score = row["Score"]

    # Sin fecha ETD: solo Score decide
    if etd is None:
        if score >= 30:
            return "🟢 VERDE – Alta prioridad (sin fecha ETD)"
        elif score >= 22:
            return "🟢 Verde media – Importante (sin fecha ETD)"
        else:
            return "⚪ Parking – Baja prioridad (sin fecha ETD)"

    # Con fecha ETD
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
    """Lee el Excel, limpia, calcula Score y Semáforo."""
    # Leer siempre la primera hoja
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # Detectar fila de encabezados (buscamos donde esté 'Proyecto')
    header_row = None
    for i in range(len(raw)):
        if str(raw.iloc[i, 0]).strip().lower() == "proyecto":
            header_row = i
            break

    if header_row is None:
        # Si no encuentra, asumimos que la primera fila ya es encabezado normal
        df = pd.read_excel(uploaded_file, sheet_name=0)
    else:
        header = raw.iloc[header_row]
        df = raw.iloc[header_row + 1 :].copy()
        df.columns = header

    # Limpiar numericos
    df = limpiar_numericos(df)

    # Calcular Score
    df["Score"] = df.apply(calcular_score, axis=1)

    # Calcular Semáforo
    hoy = datetime.today().date()
    df["Semaforo"] = df.apply(lambda r: color_por_proyecto(r, hoy), axis=1)

    return df

# ---------- APP STREAMLIT ----------

st.set_page_config(page_title="Tablero de Proyectos – Dark Board", layout="wide")

st.title("🎯 Tablero de Proyectos – Dark Board")
st.write("Sube tu archivo de proyectos en Excel y te calculo **Score** y **Semáforo** para priorizar.")

uploaded_file = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = procesar_excel(uploaded_file)

    # Ordenar por Score
    df_ordenado = df.sort_values(by="Score", ascending=False)

    # Resumen
    st.subheader("Resumen general")
    total_proj = len(df_ordenado)
    en_rojo = (df_ordenado["Semaforo"].str.contains("ROJO")).sum()
    en_amarillo = (df_ordenado["Semaforo"].str.contains("AMARILLO")).sum()
    en_verde = (df_ordenado["Semaforo"].str.contains("VERDE")).sum()
    en_parking = (df_ordenado["Semaforo"].str.contains("Parking")).sum()

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total proyectos", total_proj)
    col2.metric("🔴 Atrasados", en_rojo)
    col3.metric("🟠 Urgentes (<30 días)", en_amarillo)
    col4.metric("🟢 En buena vía", en_verde)
    col5.metric("⚪ Parking", en_parking)

    st.markdown("---")

    # Tabla
    st.subheader("Lista de proyectos (ordenados por Score)")
    st.dataframe(df_ordenado, use_container_width=True)

    # Top 5
    st.markdown("---")
    st.subheader("Top 5 proyectos recomendados para foco")
    cols_top = [c for c in ["Proyecto", "Dueño", "Score", "Semaforo", "ETD", "Estado_manual"] if c in df_ordenado.columns]
    st.table(df_ordenado[cols_top].head(5).reset_index(drop=True))

    # Descargar Excel enriquecido
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

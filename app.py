"""
ATTSF · Programa de Ecodriving
Aplicación Streamlit para importar datos mensuales, asignar conductores,
calcular puntos positivos, visualizar impacto ambiental y simular mejoras.

IMPORTANTE:
- Los datos se guardan en SQLite (ecodriving.db) en el equipo que ejecuta la app.
- Para Streamlit Community Cloud, el disco local puede ser efímero. Use una base de
  datos persistente externa antes de utilizarla como sistema oficial compartido.
- La contraseña inicial de administrador se configura con la variable de entorno
  ATTSF_ADMIN_PASSWORD o con st.secrets["admin_password"].
- Si no se configura, se usa una contraseña de demostración: cambiarla antes de publicar.
"""

from __future__ import annotations

import hashlib
import hmac
import io
import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
import base64
import pandas as pd
from github import Github
import streamlit as st

APP_TITLE = "ATTSF · Ecodriving"
DB_PATH = Path(os.getenv("ATTSF_DB_PATH", "ecodriving.db"))
DEFAULT_ADMIN_PASSWORD = "admin-attsf"
DEFAULT_PARAMS = {
    "idle_optimal_pct": 5.0,
    "idle_moderate_pct": 10.0,
    "idle_points_optimal": 40,
    "idle_points_moderate": 25,
    "idle_points_high": 10,
    "speed_min": 65.0,
    "speed_max": 85.0,
    "speed_points_good": 20,
    "speed_points_near": 10,
    "stops_good_max": 3,
    "stops_points_good": 15,
    "stops_points_other": 8,
    "distance_reference_km": 500.0,
    "distance_max_points": 15,
    "duration_target_pct": 110.0,
    "duration_points_good": 10,
    "duration_points_other": 5,
    "fuel_l_per_100km": 30.0,
    "idle_l_per_hour": 2.5,
    "co2_kg_per_litre": 2.68,
    "simulated_idle_reduction_pct": 10.0,
}

# --- RECUPERAR DATOS DE GITHUB AL ENCENDER ---
if "records" not in st.session_state:
    try:
        import io
        from github import Github
        token = st.secrets["GITHUB_TOKEN"]
        g = Github(token)
        repo = g.get_repo("Irati-Attsf/attsf-ecodriving")
        contents = repo.get_contents("")
        excel_files = [f for f in contents if f.path.endswith(".xlsx")]
        if excel_files:
            file_content = excel_files[-1].decoded_content
            records = pd.read_excel(io.BytesIO(file_content))
            st.session_state["records"] = records
        else:
            st.session_state["records"] = pd.DataFrame()
    except Exception:
        st.session_state["records"] = pd.DataFrame()
def db_connect():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.execute("PRAGMA foreign_keys = ON")
    return con


def init_db():
    with db_connect() as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS assignments (
                vehicle TEXT PRIMARY KEY,
                driver TEXT NOT NULL DEFAULT ''
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month TEXT NOT NULL,
                vehicle TEXT NOT NULL,
                distance_km REAL NOT NULL DEFAULT 0,
                avg_speed REAL NOT NULL DEFAULT 0,
                interior_km REAL NOT NULL DEFAULT 0,
                exterior_km REAL NOT NULL DEFAULT 0,
                duration_hours REAL NOT NULL DEFAULT 0,
                idle_hours REAL NOT NULL DEFAULT 0,
                stops REAL NOT NULL DEFAULT 0,
                planned_hours REAL NOT NULL DEFAULT 0,
                points_idle REAL NOT NULL DEFAULT 0,
                points_speed REAL NOT NULL DEFAULT 0,
                points_stops REAL NOT NULL DEFAULT 0,
                points_distance REAL NOT NULL DEFAULT 0,
                points_duration REAL NOT NULL DEFAULT 0,
                points_total REAL NOT NULL DEFAULT 0,
                source_name TEXT NOT NULL DEFAULT '',
                UNIQUE(month, vehicle)
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS params (
                key TEXT PRIMARY KEY,
                value REAL NOT NULL
            )
        """)
        for key, value in DEFAULT_PARAMS.items():
            con.execute("INSERT OR IGNORE INTO params(key, value) VALUES (?, ?)", (key, value))


def get_params():
    with db_connect() as con:
        rows = con.execute("SELECT key, value FROM params").fetchall()
    params = DEFAULT_PARAMS.copy()
    params.update({k: float(v) for k, v in rows})
    return params


def save_params(params):
    with db_connect() as con:
        con.executemany(
            "INSERT INTO params(key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            [(k, float(v)) for k, v in params.items()],
        )


def get_admin_password():
    try:
        secret = st.secrets.get("admin_password", "")
    except Exception:
        secret = ""
    return os.getenv("ATTSF_ADMIN_PASSWORD") or secret or DEFAULT_ADMIN_PASSWORD


def normalise_name(value):
    return " ".join(str(value).strip().lower().replace("_", " ").split())


COLUMN_ALIASES = {
    "vehicle": ["número de vehículo", "numero de vehículo", "numero de vehiculo", "vehículo", "vehiculo", "vehicle", "camión", "camion"],
    "distance_km": ["distancias recorrida (km)", "distancia recorrida (km)", "distancia recorrida", "distancias recorrida", "km recorridos", "distance_km", "distancia km"],
    "avg_speed": ["velocidad media (km/h)", "velocidad media", "avg_speed", "velocidad"],
    "interior_km": ["km carretera interior", "carretera interior", "interior_km", "km interior"],
    "exterior_km": ["km carretera exterior", "carretera exterior", "exterior_km", "km exterior"],
    "duration_hours": ["duración viaje", "duracion viaje", "duración del viaje", "duracion del viaje", "duration_hours", "duración"],
    "idle_hours": ["relentí", "relenti", "ralentí", "ralenti", "tiempo ralentí", "tiempo de ralentí", "idle_hours"],
    "stops": ["numero de paradas", "número de paradas", "paradas", "stops"],
    "planned_hours": ["duración planificada", "duracion planificada", "tiempo planificado", "planned_hours", "horas planificadas"],
}


def parse_duration(value):
    """Convierte números, 'HH:MM:SS' o 'HH:MM' en horas."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        # En Excel, duraciones de tiempo a menudo se almacenan como fracción de día.
        val = float(value)
        return val * 24 if 0 <= val < 1 else val
    s = str(value).strip()
    if not s:
        return 0.0
    parts = s.split(":")
    try:
        if len(parts) == 3:
            return float(parts[0]) + float(parts[1]) / 60 + float(parts[2]) / 3600
        if len(parts) == 2:
            return float(parts[0]) + float(parts[1]) / 60
        return float(s.replace(",", "."))
    except ValueError:
        return 0.0


def find_columns(df):
    lookup = {normalise_name(c): c for c in df.columns}
    mapping = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            key = normalise_name(alias)
            if key in lookup:
                mapping[canonical] = lookup[key]
                break
    return mapping


def prepare_excel(uploaded, month):
    raw = pd.read_excel(uploaded)
    mapping = find_columns(raw)
    if "vehicle" not in mapping:
        raise ValueError("No encuentro la columna del número de vehículo. Revisa el encabezado del Excel.")
    df = pd.DataFrame()
    df["vehicle"] = raw[mapping["vehicle"]].astype(str).str.strip()
    df = df[df["vehicle"].notna() & (df["vehicle"] != "") & (df["vehicle"].str.lower() != "nan")].copy()
    numeric_fields = ["distance_km", "avg_speed", "interior_km", "exterior_km", "stops"]
    for field in numeric_fields:
        if field in mapping:
            df[field] = pd.to_numeric(raw.loc[df.index, mapping[field]].astype(str).str.replace(",", ".", regex=False), errors="coerce").fillna(0)
        else:
            df[field] = 0.0
    for field in ["duration_hours", "idle_hours", "planned_hours"]:
        if field in mapping:
            df[field] = raw.loc[df.index, mapping[field]].apply(parse_duration)
        else:
            df[field] = 0.0
    # Ralentí suele venir como duración HH:MM:SS. Si el Excel trae porcentaje,
    # hay que adaptar este campo; se interpreta como tiempo, no como porcentaje.
    df["month"] = month
    df["source_name"] = uploaded.name
    return df


def score_row(row, p):
    duration = max(float(row["duration_hours"]), 0)
    idle = max(float(row["idle_hours"]), 0)
    idle_pct = (idle / duration * 100) if duration > 0 else 0
    if idle_pct < p["idle_optimal_pct"]:
        pi = p["idle_points_optimal"]
    elif idle_pct <= p["idle_moderate_pct"]:
        pi = p["idle_points_moderate"]
    else:
        pi = p["idle_points_high"]

    speed = float(row["avg_speed"])
    if p["speed_min"] <= speed <= p["speed_max"]:
        ps = p["speed_points_good"]
    else:
        # El rango de tolerancia se define como 10 km/h fuera del rango eficiente.
        if (p["speed_min"] - 10 <= speed < p["speed_min"]) or (p["speed_max"] < speed <= p["speed_max"] + 10):
            ps = p["speed_points_near"]
        else:
            ps = 0

    stops = max(float(row["stops"]), 0)
    pst = p["stops_points_good"] if stops <= p["stops_good_max"] else p["stops_points_other"]

    distance = max(float(row["distance_km"]), 0)
    pdist = min(p["distance_max_points"], distance / max(p["distance_reference_km"], 1) * p["distance_max_points"])

    planned = max(float(row["planned_hours"]), 0)
    if planned > 0:
        on_time = duration <= planned * p["duration_target_pct"] / 100
        pdur = p["duration_points_good"] if on_time else p["duration_points_other"]
    else:
        # Sin tiempo planificado no es justo penalizar: se otorga la puntuación base.
        pdur = p["duration_points_good"]

    return {
        "points_idle": round(max(pi, 0), 2),
        "points_speed": round(max(ps, 0), 2),
        "points_stops": round(max(pst, 0), 2),
        "points_distance": round(max(pdist, 0), 2),
        "points_duration": round(max(pdur, 0), 2),
        "points_total": round(max(pi, 0) + max(ps, 0) + max(pst, 0) + max(pdist, 0) + max(pdur, 0), 2),
        "idle_pct": round(idle_pct, 2),
    }


def load_records():
    with db_connect() as con:
        df = pd.read_sql_query("SELECT * FROM records", con)
        assignments = pd.read_sql_query("SELECT * FROM assignments", con)
    if not df.empty:
        df = df.merge(assignments, on="vehicle", how="left")
        df["driver"] = df["driver"].fillna("").replace("", "Sin asignar")
    return df


def store_import(df, params):
    # Cada importación reemplaza los registros de los mismos vehículos y mes,
    # para permitir corregir un Excel sin duplicar puntos.
    scores = df.apply(lambda r: score_row(r, params), axis=1, result_type="expand")
    df = pd.concat([df.reset_index(drop=True), scores.reset_index(drop=True)], axis=1)
    with db_connect() as con:
        for vehicle in df["vehicle"].unique():
            con.execute("INSERT OR IGNORE INTO assignments(vehicle, driver) VALUES (?, '')", (vehicle,))
        for _, r in df.iterrows():
            con.execute("""
                INSERT INTO records (
                    month, vehicle, distance_km, avg_speed, interior_km, exterior_km,
                    duration_hours, idle_hours, stops, planned_hours,
                    points_idle, points_speed, points_stops, points_distance, points_duration,
                    points_total, source_name
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(month, vehicle) DO UPDATE SET
                    distance_km=excluded.distance_km, avg_speed=excluded.avg_speed,
                    interior_km=excluded.interior_km, exterior_km=excluded.exterior_km,
                    duration_hours=excluded.duration_hours, idle_hours=excluded.idle_hours,
                    stops=excluded.stops, planned_hours=excluded.planned_hours,
                    points_idle=excluded.points_idle, points_speed=excluded.points_speed,
                    points_stops=excluded.points_stops, points_distance=excluded.points_distance,
                    points_duration=excluded.points_duration, points_total=excluded.points_total,
                    source_name=excluded.source_name
            """, (
                str(r["month"]), str(r["vehicle"]), float(r["distance_km"]), float(r["avg_speed"]),
                float(r["interior_km"]), float(r["exterior_km"]), float(r["duration_hours"]),
                float(r["idle_hours"]), float(r["stops"]), float(r["planned_hours"]),
                float(r["points_idle"]), float(r["points_speed"]), float(r["points_stops"]),
                float(r["points_distance"]), float(r["points_duration"]), float(r["points_total"]),
                str(r["source_name"])
            ))


def driver_badges(row):
    badges = []
    if row.get("idle_pct", 100) < 5:
        badges.append("🟢 Maestro del Ralentí")
    if row.get("points_speed", 0) >= 20:
        badges.append("🛣️ Ruta Fluida")
    if row.get("points_total", 0) >= 85:
        badges.append("🌿 Conductor Verde")
    if row.get("points_distance", 0) >= 14:
        badges.append("📍 Kilómetros Eficientes")
    return badges


def help_section(role):
    with st.expander("❓ Ayuda / Guía de uso", expanded=False):
        if role == "Administrador":
            st.markdown("""
            **1. Importar el Excel mensual**
            - Elige el mes al que pertenecen los datos.
            - Sube un archivo `.xlsx` o `.xls`.
            - Comprueba la vista previa y pulsa **Procesar e importar**.
            - Se aceptan encabezados equivalentes a los descritos en la guía. Si un dato no está presente, se usa 0.
            - Para duración y ralentí se admiten horas numéricas o formatos `HH:MM` / `HH:MM:SS`.

            **2. Asignar conductores**
            - En la pestaña **Conductores**, escribe el nombre para cada vehículo y guarda.
            - Los vehículos nuevos aparecen como **Sin asignar** hasta que los identifiques.

            **3. Configurar puntuación**
            - En **Parámetros**, ajusta los umbrales y valores. Los puntos nunca bajan de cero.
            - Al cambiar parámetros, vuelve a importar el mes para recalcular sus puntuaciones.

            **4. Interpretar impacto ambiental**
            - El combustible y el CO₂ son estimaciones basadas en los parámetros configurados, no mediciones directas del camión.
            - Define un consumo de referencia realista y valida los cálculos con el equipo de flota.
            """)
        else:
            st.markdown("""
            **Cómo sumar puntos**
            - Mantén el ralentí bajo, conduce dentro del rango de velocidad eficiente y evita paradas innecesarias cuando la seguridad y la operación lo permitan.
            - Completa las rutas de forma segura. La distancia aporta puntos proporcionales hasta el máximo configurado.
            - No se restan puntos por malos resultados: cada criterio suma puntos o aporta cero.

            **Qué puedes consultar**
            - **Clasificación:** puntos por mes y acumulados.
            - **Impacto ambiental:** estimación de combustible y CO₂ ahorrados.
            - **Simulador:** prueba mejoras hipotéticas para el mes siguiente.
            - **Insignias y evolución:** hitos positivos y progreso de la flota.

            Las velocidades, los tiempos y las paradas deben interpretarse siempre según la seguridad, el tipo de carga, el terreno y las necesidades humanitarias.
            """)


def environmental_summary(df, p):
    if df.empty:
        return 0.0, 0.0, 0.0
    km = df["distance_km"].clip(lower=0).sum()
    idle_h = df["idle_hours"].clip(lower=0).sum()
    # Estimación de ahorro potencial por ralentí frente a una referencia del 10% del viaje.
    duration = df["duration_hours"].clip(lower=0).sum()
    reference_idle = duration * 0.10
    idle_saved_h = max(reference_idle - idle_h, 0)
    idle_fuel_saved = idle_saved_h * p["idle_l_per_hour"]
    # No se afirma ahorro real de conducción: la cifra es potencial estimada vs consumo de referencia.
    fuel_reference = km * p["fuel_l_per_100km"] / 100
    efficiency_factor = min(max(df["points_total"].mean() / 100, 0), 1) if not df.empty else 0
    driving_saving = fuel_reference * 0.05 * efficiency_factor
    fuel_saved = max(idle_fuel_saved + driving_saving, 0)
    co2_saved = fuel_saved * p["co2_kg_per_litre"]
    return fuel_saved, co2_saved, fuel_reference


def main():
    init_db()
    params = get_params()
    st.title("🌱 ATTSF | Programa de Ecodriving")
    st.caption("Conducción eficiente · Reconocimiento positivo · Flota humanitaria")
    st.sidebar.header("Acceso")
    role = st.sidebar.radio("Selecciona la vista", ["Conductor / Consulta", "Administrador"])
    is_admin = role == "Administrador"
    if is_admin:
        password = st.sidebar.text_input("Contraseña de administrador", type="password")
        if not hmac.compare_digest(password, get_admin_password()):
            st.info("Introduce la contraseña de administrador para acceder a las herramientas de gestión.")
            help_section("Administrador")
            st.stop()
        st.sidebar.success("Sesión de administrador activa")
    else:
        st.sidebar.info("Modo de consulta: no puedes modificar los datos.")

    help_section("Administrador" if is_admin else "Consulta")
    records = load_records()

    if is_admin:
        tabs = st.tabs(["📥 Importar Excel", "👥 Conductores", "⚙️ Parámetros", "📊 Dashboard", "🌍 Impacto ambiental", "🏆 Clasificación"])
        with tabs[0]:
            st.subheader("Importación mensual")
            month = st.date_input("Mes de los datos", value=datetime.today().date().replace(day=1))
            month_key = month.strftime("%Y-%m")
            uploaded = st.file_uploader("Sube el Excel mensual", type=["xlsx", "xls"])
            if uploaded:
                try:
                    preview = prepare_excel(uploaded, month_key)
                    st.write("Vista previa de los datos normalizados")
                    st.dataframe(preview, use_container_width=True)
                    if "planned_hours" not in find_columns(pd.read_excel(uploaded)):
                        st.warning("No se ha detectado una columna de duración planificada. En ese caso, el sistema otorga la puntuación positiva configurada para duración.")
                    if st.button("Procesar e importar", type="primary"):
                        store_import(preview, params)
                        st.success(f"Importación completada: {len(preview)} vehículos/registro(s) para {month_key}.")
                        st.rerun()
                except Exception as exc:
                    st.error(f"No se pudo leer el Excel: {exc}")
            st.markdown("**Columnas reconocidas**")
            st.code("Número de vehículo | Distancia recorrida (km) | Velocidad media (km/h) | km carretera interior | km carretera exterior | duración viaje | Ralentí | Número de paradas")
            st.caption("El programa interpreta una fila por vehículo y mes. Si un vehículo aparece en varias filas del mismo mes, consolida la última fila importada; prepara el Excel con una fila mensual por vehículo.")
        with tabs[1]:
            st.subheader("Asignación de conductores")
            with db_connect() as con:
                vehicles = pd.read_sql_query("SELECT vehicle, driver FROM assignments ORDER BY vehicle", con)
            if vehicles.empty:
                st.info("Todavía no hay vehículos. Importa primero un Excel mensual.")
            else:
                edited = st.data_editor(
                    vehicles, use_container_width=True, hide_index=True,
                    column_config={
                        "vehicle": st.column_config.TextColumn("Número de vehículo", disabled=True),
                        "driver": st.column_config.TextColumn("Nombre del conductor"),
                    },
                    key="assignment_editor",
                )
                if st.button("Guardar asignaciones", type="primary"):
                    with db_connect() as con:
                        con.executemany(
                            "UPDATE assignments SET driver=? WHERE vehicle=?",
                            [(str(r["driver"]).strip(), str(r["vehicle"])) for _, r in edited.iterrows()],
                        )
                    st.success("Asignaciones guardadas.")
                    st.rerun()
        with tabs[2]:
            st.subheader("Parámetros del sistema")
            st.caption("Los cambios afectan a las próximas importaciones. Para recalcular un mes ya importado, vuelve a importar su Excel.")
            with st.form("params_form"):
                c1, c2, c3 = st.columns(3)
                newp = params.copy()
                with c1:
                    st.markdown("**Ralentí**")
                    newp["idle_optimal_pct"] = st.slider("% ralentí óptimo: por debajo de", 1.0, 10.0, float(params["idle_optimal_pct"]), 0.5)
                    newp["idle_moderate_pct"] = st.slider("% ralentí moderado hasta", 5.0, 20.0, float(params["idle_moderate_pct"]), 0.5)
                    newp["idle_points_optimal"] = st.slider("Puntos ralentí óptimo", 0, 40, int(params["idle_points_optimal"]))
                    newp["idle_points_moderate"] = st.slider("Puntos ralentí moderado", 0, 40, int(params["idle_points_moderate"]))
                    newp["idle_points_high"] = st.slider("Puntos ralentí alto", 0, 40, int(params["idle_points_high"]))
                with c2:
                    st.markdown("**Velocidad y paradas**")
                    newp["speed_min"] = st.slider("Velocidad eficiente mínima (km/h)", 20.0, 100.0, float(params["speed_min"]), 1.0)
                    newp["speed_max"] = st.slider("Velocidad eficiente máxima (km/h)", 40.0, 120.0, float(params["speed_max"]), 1.0)
                    newp["speed_points_good"] = st.slider("Puntos velocidad en rango", 0, 20, int(params["speed_points_good"]))
                    newp["speed_points_near"] = st.slider("Puntos velocidad cercana", 0, 20, int(params["speed_points_near"]))
                    newp["stops_good_max"] = st.slider("Máximo de paradas para conducción fluida", 0, 20, int(params["stops_good_max"]))
                    newp["stops_points_good"] = st.slider("Puntos pocas paradas", 0, 15, int(params["stops_points_good"]))
                    newp["stops_points_other"] = st.slider("Puntos otras rutas", 0, 15, int(params["stops_points_other"]))
                with c3:
                    st.markdown("**Distancia, duración y ambiente**")
                    newp["distance_reference_km"] = st.slider("Km de referencia para máximo de distancia", 50.0, 2000.0, float(params["distance_reference_km"]), 25.0)
                    newp["distance_max_points"] = st.slider("Puntos máximos distancia", 0, 15, int(params["distance_max_points"]))
                    newp["duration_target_pct"] = st.slider("Tolerancia de duración planificada (%)", 100.0, 150.0, float(params["duration_target_pct"]), 1.0)
                    newp["duration_points_good"] = st.slider("Puntos por cumplir duración", 0, 10, int(params["duration_points_good"]))
                    newp["duration_points_other"] = st.slider("Puntos duración alternativa", 0, 10, int(params["duration_points_other"]))
                    newp["fuel_l_per_100km"] = st.slider("Consumo de referencia (L/100 km)", 10.0, 80.0, float(params["fuel_l_per_100km"]), 1.0)
                    newp["idle_l_per_hour"] = st.slider("Consumo estimado al ralentí (L/h)", 0.5, 8.0, float(params["idle_l_per_hour"]), 0.1)
                    newp["co2_kg_per_litre"] = st.slider("Factor CO₂ (kg por litro)", 2.0, 3.5, float(params["co2_kg_per_litre"]), 0.01)
                submitted = st.form_submit_button("Guardar parámetros", type="primary")
                if submitted:
                    if newp["speed_min"] > newp["speed_max"]:
                        st.error("La velocidad mínima no puede ser superior a la máxima.")
                    elif newp["idle_optimal_pct"] >= newp["idle_moderate_pct"]:
                        st.error("El umbral óptimo de ralentí debe ser inferior al moderado.")
                    else:
                        save_params(newp)
                        st.success("Parámetros guardados.")
                        st.rerun()
        dashboard_tab, impact_tab, ranking_tab = tabs[3], tabs[4], tabs[5]
    else:
        dashboard_tab, impact_tab, ranking_tab, simulator_tab, badges_tab = st.tabs([
            "📊 Dashboard", "🌍 Impacto ambiental", "🏆 Clasificación", "🎛️ Simulador", "🎖️ Insignias"
        ])

    # Dashboard compartido
    with dashboard_tab:
        st.subheader("Evolución y rendimiento de la flota")
        if records.empty:
            st.info("Todavía no hay datos. El administrador debe importar el Excel mensual.")
        else:
            monthly = records.groupby("month", as_index=False).agg(
                puntos=("points_total", "sum"),
                distancia_km=("distance_km", "sum"),
                ralentí_h=("idle_hours", "sum"),
                vehículos=("vehicle", "nunique"),
            ).sort_values("month")
            k1, k2, k3 = st.columns(3)
            k1.metric("Puntos acumulados", f'{records["points_total"].sum():,.0f}')
            k2.metric("Kilómetros registrados", f'{records["distance_km"].sum():,.0f} km')
            k3.metric("Vehículos registrados", f'{records["vehicle"].nunique()}')
            fig = px.line(monthly, x="month", y="puntos", markers=True, title="Puntos positivos acumulados por mes")
            st.plotly_chart(fig, use_container_width=True)
            fig2 = px.bar(monthly, x="month", y="distancia_km", title="Distancia mensual de la flota (km)")
            st.plotly_chart(fig2, use_container_width=True)
            st.dataframe(monthly, use_container_width=True, hide_index=True)

    with ranking_tab:
      st.subheader("🏆 Clasificación General de Conductores")
    
    if not records.empty:
        # Agrupar puntos acumulados por cada conductor
        df_ranking = records.groupby("driver", as_index=False)["points_total"].sum()
        df_ranking = df_ranking.sort_values(by="points_total", ascending=False).reset_index(drop=True)
        
        for index, row in df_ranking.iterrows():
            conductor = row["driver"]
            puntos = row["points_total"]
            
            # Asignar insignias al Top 5
            if index == 0:
                insignia = "🥇 1º"
            elif index == 1:
                insignia = "🥈 2º"
            elif index == 2:
                insignia = "🥉 3º"
            elif index == 3:
                insignia = "🏅 4º"
            elif index == 4:
                insignia = "🏅 5º"
            else:
                insignia = f"{index + 1}º"
                
            # Mostrar en pantalla
            if index < 5:
                st.markdown(f"### {insignia} — **{conductor}**: {puntos:.0f} puntos")
            else:
                st.write(f"{insignia} — {conductor}: {puntos:.0f} puntos")
    else:
        st.info("Todavía no hay registros cargados para calcular la clasificación.")
      
    with impact_tab:
        st.subheader("Impacto ambiental y ahorro estimado")
        if records.empty:
            st.info("Importa datos para calcular estimaciones.")
        else:
            month_filter = st.selectbox("Periodo ambiental", ["Todo el histórico"] + sorted(records["month"].unique(), reverse=True), key="impact_period")
            scoped = records if month_filter == "Todo el histórico" else records[records["month"] == month_filter]
            fuel_saved, co2_saved, reference = environmental_summary(scoped, params)
            a, b, c = st.columns(3)
            a.metric("Combustible potencialmente ahorrado", f"{fuel_saved:,.1f} L")
            b.metric("CO₂ potencialmente evitado", f"{co2_saved:,.1f} kg")
            c.metric("Consumo de referencia estimado", f"{reference:,.1f} L")
            st.info("Estimación orientativa, no medición directa. Se calcula frente a una referencia de ralentí del 10% del tiempo de viaje y una mejora teórica de conducción de hasta el 5%, ponderada por puntos. Para informes oficiales, sustituye estas hipótesis por datos de repostaje/telemática y compara periodos equivalentes.")
            monthly_env = records.groupby("month", as_index=False).agg(km=("distance_km", "sum"), idle_h=("idle_hours", "sum"), puntos=("points_total", "sum"))
            monthly_env["combustible_ref_L"] = monthly_env["km"] * params["fuel_l_per_100km"] / 100
            monthly_env["CO2_ref_kg"] = monthly_env["combustible_ref_L"] * params["co2_kg_per_litre"]
            st.dataframe(monthly_env, use_container_width=True, hide_index=True)

    if not is_admin:
        with simulator_tab:
            st.subheader("¿Qué pasaría si…?")
            st.write("Simula mejoras posibles para el mes siguiente. No cambia la puntuación oficial ni el histórico.")
            if records.empty:
                st.info("El simulador estará disponible cuando haya datos.")
            else:
                drivers = sorted(records["driver"].unique())
                driver = st.selectbox("Conductor/a", drivers)
                driver_rows = records[records["driver"] == driver]
                idle_reduction = st.slider("Reducir el tiempo de ralentí (%)", 0, 50, int(params["simulated_idle_reduction_pct"]), 5)
                speed_improve = st.slider("Mejorar la eficiencia de velocidad (puntos porcentuales)", 0, 100, 20, 5)
                duration = driver_rows["duration_hours"].sum()
                idle = driver_rows["idle_hours"].sum()
                reduced_idle = idle * (1 - idle_reduction / 100)
                old_idle_pct = (idle / duration * 100) if duration else 0
                new_idle_pct = (reduced_idle / duration * 100) if duration else 0
                if new_idle_pct < params["idle_optimal_pct"]:
                    new_idle_points = params["idle_points_optimal"]
                elif new_idle_pct <= params["idle_moderate_pct"]:
                    new_idle_points = params["idle_points_moderate"]
                else:
                    new_idle_points = params["idle_points_high"]
                current_idle_points = driver_rows["points_idle"].mean()
                extra_idle_points = max(new_idle_points - current_idle_points, 0) * len(driver_rows)
                extra_speed_points = max(params["speed_points_good"] - driver_rows["points_speed"].mean(), 0) * len(driver_rows) * speed_improve / 100
                extra_points = extra_idle_points + extra_speed_points
                saved_litres = max(idle - reduced_idle, 0) * params["idle_l_per_hour"]
                saved_co2 = saved_litres * params["co2_kg_per_litre"]
                m1, m2, m3 = st.columns(3)
                m1.metric("Puntos extra potenciales", f"+{extra_points:.1f}")
                m2.metric("Combustible potencialmente ahorrado", f"{saved_litres:.1f} L")
                m3.metric("CO₂ potencialmente evitado", f"{saved_co2:.1f} kg")
                st.caption(f"Ralentí estimado: {old_idle_pct:.1f}% → {new_idle_pct:.1f}%. Son escenarios hipotéticos; no se restan puntos si la mejora no se alcanza.")
        with badges_tab:
            st.subheader("Insignias y logros")
            if records.empty:
                st.info("Aún no hay insignias. Importa datos para comenzar.")
            else:
                latest_month = sorted(records["month"].unique())[-1]
                latest = records[records["month"] == latest_month].copy()
                latest["idle_pct"] = latest.apply(lambda r: (r["idle_hours"] / r["duration_hours"] * 100) if r["duration_hours"] > 0 else 0, axis=1)
                latest["badges"] = latest.apply(lambda r: ", ".join(driver_badges(r)) or "Sigue sumando puntos", axis=1)
                st.caption(f"Insignias automáticas para el último mes disponible: {latest_month}")
                st.dataframe(latest[["driver", "vehicle", "points_total", "badges"]].rename(columns={"points_total": "Puntos"}), use_container_width=True, hide_index=True)
                st.markdown("""
                - 🟢 **Maestro del Ralentí:** ralentí inferior al 5%.
                - 🛣️ **Ruta Fluida:** puntuación máxima del criterio de velocidad.
                - 🌿 **Conductor Verde:** 85 puntos o más en el registro mensual.
                - 📍 **Kilómetros Eficientes:** 14 puntos o más en distancia.

                Los hitos son reconocimientos positivos; no alcanzar una insignia no resta puntos.
                """)


if __name__ == "__main__":
    main()

def guardar_excel_en_github(df, nombre_archivo="datos_ecodriving.xlsx"):
    try:
        token = st.secrets["GITHUB_TOKEN"]
        g = Github(token)
        repo = g.get_repo("Irati-Attsf/attsf-ecodriving")

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        contenido_bytes = buffer.getvalue()

        try:
            contents = repo.get_contents(nombre_archivo)
            repo.update_file(
                nombre_archivo,
                "Guardando datos",
                contenido_bytes,
                contents.sha,
            )
        except Exception:
            repo.create_file(nombre_archivo, "Guardando datos", contenido_bytes)

        st.success("¡Datos guardados!")
    except Exception as e:
        st.error(f"Error: {e}")
# --- CARGAR DATOS AUTOMÁTICAMENTE DESDE GITHUB AL ARRANCAR ---
try:
    token = st.secrets["GITHUB_TOKEN"]
    g = Github(token)
    repo = g.get_repo("Irati-Attsf/attsf-ecodriving")
    file_content = repo.get_contents("datos_ecodriving.xlsx")
    decoded_content = base64.b64decode(file_content.content)
    df_github = pd.read_excel(io.BytesIO(decoded_content))
    if "df_flota" not in st.session_state or st.session_state["df_flota"].empty:
        st.session_state["df_flota"] = df_github
except Exception:
    pass
  # --- SECCIÓN DE ADMINISTRADOR: BORRAR ARCHIVOS EXCEL ORDENADOS POR FECHA ---
if st.session_state.get("admin", False):
    with st.sidebar:
        st.markdown("---")
        st.subheader("⚙️ Gestión de Archivos (Admin)")
        try:
            token = st.secrets["GITHUB_TOKEN"]
            g = Github(token)
            repo = g.get_repo("Irati-Attsf/attsf-ecodriving")
            
            # Buscar archivos .xlsx y obtener su fecha de subida
            contents = repo.get_contents("")
            excel_files_info = []
            for f in contents:
                if f.path.endswith(".xlsx"):
                    try:
                        commits = repo.get_commits(path=f.path)
                        fecha = list(commits)[0].commit.author.date
                    except Exception:
                        fecha = ""
                    excel_files_info.append((f.path, fecha))
            
            # Ordenar de más reciente a más antiguo
            excel_files_info.sort(key=lambda x: x[1], reverse=True)
            excel_files = [item[0] for item in excel_files_info]
            
            if excel_files:
                archivo_a_borrar = st.selectbox("Elige el Excel (más reciente primero):", excel_files)
                if st.button("🗑️ Borrar archivo elegido"):
                    file_to_delete = repo.get_contents(archivo_a_borrar)
                    repo.delete_file(
                        file_to_delete.path, 
                        f"Eliminar archivo {archivo_a_borrar} por administrador", 
                        file_to_delete.sha
                    )
                    st.success(f"¡'{archivo_a_borrar}' borrado correctamente!")
                    st.rerun()
            else:
                st.info("No hay archivos Excel guardados en GitHub.")
        except Exception as e:
            st.info("Configura GITHUB_TOKEN en los Secrets para gestionar los archivos.")
import streamlit as st
from github import Github
import pandas as pd

# Función para conectar con GitHub usando el token secreto
def get_github_repo():
    token = st.secrets["GITHUB_TOKEN"]
    g = Github(token)
    return g.get_repo("Irati-Attsf/attsf-ecodriving")

# Función para subir un archivo Excel a GitHub
def upload_excel_to_github(uploaded_file, file_path):
    repo = get_github_repo()
    file_content = uploaded_file.getvalue()
    try:
        # Si el archivo ya existe, lo actualiza
        contents = repo.get_contents(file_path)
        repo.update_file(contents.path, "Actualizar archivo Excel vía Streamlit", file_content, contents.sha)
        st.success(f"¡Archivo '{file_path}' actualizado correctamente en GitHub!")
    except:
        # Si no existe, lo crea nuevo
        repo.create_file(file_path, "Subir nuevo archivo Excel vía Streamlit", file_content)
        st.success(f"¡Archivo '{file_path}' subido correctamente a GitHub!")

# Función para eliminar un archivo Excel de GitHub
def delete_excel_from_github(file_path):
    repo = get_github_repo()
    try:
        contents = repo.get_contents(file_path)
        repo.delete_file(contents.path, "Eliminar archivo Excel vía Streamlit", contents.sha)
        st.success(f"¡Archivo '{file_path}' eliminado correctamente de GitHub!")
    except Exception as e:
        st.error(f"No se pudo eliminar el archivo: {e}")

# --- Panel visual para subir y borrar archivos Excel (Solo Administradores) ---
try:
    es_admin = (st.session_state.get('vista') == "Administrador") or ('Administrador' in str(st.session_state))
except:
    es_admin = True

if es_admin:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 Gestión de Datos (Excel)")
    
    # 1. Botón para subir archivos
    uploaded_file = st.sidebar.file_uploader("Selecciona un archivo Excel", type=["xlsx", "xls"])
    if uploaded_file is not None:
        nombre_archivo = uploaded_file.name
        if st.sidebar.button("Subir archivo a GitHub"):
            upload_excel_to_github(uploaded_file, nombre_archivo)

    # 2. Apartado para borrar archivos
    st.sidebar.markdown("---")
    archivo_a_borrar = st.sidebar.text_input("Nombre del archivo a eliminar (ej: datos.xlsx)")
    if st.sidebar.button("Eliminar archivo de GitHub"):
        if archivo_a_borrar:
            delete_excel_from_github(archivo_a_borrar)
        else:
            st.sidebar.warning("Escribe el nombre exacto del archivo que quieres borrar.")

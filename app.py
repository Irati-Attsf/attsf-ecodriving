"""
ATTSF · Programa de Ecodriving
Aplicación Streamlit para importar datos mensuales, asignar conductores,
calcular puntos positivos, visualizar impacto ambiental y simular mejoras.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import io
import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from github import Github

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

# --- 1. BASE DE DATOS Y CONEXIÓN ---
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

# --- 2. INTEGRACIÓN CON GITHUB ---
def get_github_repo():
    token = st.secrets["GITHUB_TOKEN"]
    g = Github(token)
    return g.get_repo("Irati-Attsf/attsf-ecodriving")

def upload_excel_to_github(uploaded_file, file_path):
    repo = get_github_repo()
    file_content = uploaded_file.getvalue()
    try:
        contents = repo.get_contents(file_path)
        repo.update_file(contents.path, "Actualizar archivo Excel vía Streamlit", file_content, contents.sha)
    except:
        repo.create_file(file_path, "Subir nuevo archivo Excel vía Streamlit", file_content)

def delete_excel_from_github(file_path):
    repo = get_github_repo()
    try:
        contents = repo.get_contents(file_path)
        repo.delete_file(contents.path, "Eliminar archivo Excel vía Streamlit", contents.sha)
    except Exception as e:
        st.error(f"Error al eliminar: {e}")

def cargar_asignaciones_de_github():
    url = f"https://api.github.com/repos/Irati-Attsf/attsf-ecodriving/contents/assignments.json"
    headers = {"Authorization": f"Bearer {st.secrets['GITHUB_TOKEN']}"}
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            file_content = response.json()["content"]
            decoded_content = base64.b64decode(file_content).decode('utf-8')
            assignments_data = json.loads(decoded_content)
            with db_connect() as con:
                for vehicle, driver in assignments_data.items():
                    con.execute("INSERT OR REPLACE INTO assignments (vehicle, driver) VALUES (?, ?)", (vehicle, driver))
    except Exception:
        pass

def guardar_asignaciones_en_github():
    url = f"https://api.github.com/repos/Irati-Attsf/attsf-ecodriving/contents/assignments.json"
    headers = {"Authorization": f"Bearer {st.secrets['GITHUB_TOKEN']}"}
    with db_connect() as con:
        cursor = con.cursor()
        cursor.execute("SELECT vehicle, driver FROM assignments")
        assignments_dict = {row[0]: row[1] for row in cursor.fetchall()}
    
    json_data = json.dumps(assignments_dict, indent=4)
    encoded_content = base64.b64encode(json_data.encode('utf-8')).decode('utf-8')
    
    response = requests.get(url, headers=headers)
    sha = response.json().get("sha") if response.status_code == 200 else None
    
    data = {
        "message": "Actualización de asignaciones",
        "content": encoded_content
    }
    if sha:
        data["sha"] = sha
    requests.put(url, headers=headers, json=data)

# --- 3. PARÁMETROS Y PROCESAMIENTO ---
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
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
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

def prepare_excel(file_obj_or_path, month):
    raw = pd.read_excel(file_obj_or_path)
    mapping = find_columns(raw)
    if "vehicle" not in mapping:
        raise ValueError("No encuentro la columna del número de vehículo.")
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
    df["month"] = month
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

def store_import(df, params, source_name, month_override=None):
    scores = df.apply(lambda r: score_row(r, params), axis=1, result_type="expand")
    df = pd.concat([df.reset_index(drop=True), scores.reset_index(drop=True)], axis=1)
    if month_override:
        df["month"] = month_override
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
                source_name
            ))

# Sincronización automática de archivos Excel guardados en GitHub al arrancar
def sincronizar_excel_desde_github(params):
    try:
        repo = get_github_repo()
        contents = repo.get_contents("")
        excel_files = [f for f in contents if f.path.endswith(".xlsx")]
        for file_content in excel_files:
            file_data = file_content.decoded_content
            nombre = file_content.path
            # Intentar extraer el mes si viene en el nombre (ej: 2026-09.xlsx), si no usar mes actual por defecto
            mes_inferido = datetime.today().strftime("%Y-%m")
            for part in nombre.replace(".xlsx", "").split("_"):
                if len(part) == 7 and part[4] == "-":
                    mes_inferido = part
                    break
            df = prepare_excel(io.BytesIO(file_data), mes_inferido)
            store_import(df, params, nombre, month_override=mes_inferido)
    except Exception:
        pass

def load_records():
    with db_connect() as con:
        df = pd.read_sql_query("SELECT * FROM records", con)
        assignments = pd.read_sql_query("SELECT * FROM assignments", con)
    if not df.empty and not assignments.empty:
        df = df.merge(assignments, on="vehicle", how="left")
        df["driver"] = df["driver"].fillna("").replace("", "Sin asignar")
    elif not df.empty:
        df["driver"] = "Sin asignar"
    return df

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
            - Sube un archivo `.xlsx` o `.xls` (se guardará automáticamente para siempre en GitHub).
            """)
        else:
            st.markdown("""
            **Cómo sumar puntos**
            - Mantén el ralentí bajo, conduce dentro del rango eficiente y evita paradas innecesarias.
            """)

def environmental_summary(df, p):
    if df.empty:
        return 0.0, 0.0, 0.0
    km = df["distance_km"].clip(lower=0).sum()
    idle_h = df["idle_hours"].clip(lower=0).sum()
    duration = df["duration_hours"].clip(lower=0).sum()
    reference_idle = duration * 0.10
    idle_saved_h = max(reference_idle - idle_h, 0)
    idle_fuel_saved = idle_saved_h * p["idle_l_per_hour"]
    fuel_reference = km * p["fuel_l_per_100km"] / 100
    efficiency_factor = min(max(df["points_total"].mean() / 100, 0), 1) if not df.empty else 0
    driving_saving = fuel_reference * 0.05 * efficiency_factor
    fuel_saved = max(idle_fuel_saved + driving_saving, 0)
    co2_saved = fuel_saved * p["co2_kg_per_litre"]
    return fuel_saved, co2_saved, fuel_reference

# --- 4. APLICACIÓN PRINCIPAL ---
def main():
    init_db()
    
    try:
        cargar_asignaciones_de_github()
    except Exception:
        pass

    params = get_params()

    # Sincronización automática de GitHub al encender o reiniciar la app
    sincronizar_excel_desde_github(params)

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
        st.sidebar.info("Modo de consulta: los conductores pueden ver las clasificaciones e insignias en tiempo real.")

    help_section("Administrador" if is_admin else "Consulta")
    records = load_records()

    if is_admin:
        tabs = st.tabs(["📥 Importar Excel", "👥 Conductores", "⚙️ Parámetros", "📊 Dashboard", "🌍 Impacto ambiental", "🏆 Clasificación"])
        with tabs[0]:
            st.subheader("Importación mensual y gestión de archivos")
            month = st.date_input("Mes de los datos", value=datetime.today().date().replace(day=1))
            month_key = month.strftime("%Y-%m")
            uploaded = st.file_uploader("Sube el Excel mensual", type=["xlsx", "xls"])
            if uploaded:
                try:
                    # Guardar con nombre que incluya el mes para identificarlo claramente
                    safe_name = f"{month_key}_{uploaded.name}"
                    upload_excel_to_github(uploaded, safe_name)
                    preview = prepare_excel(uploaded, month_key)
                    st.write("Vista previa de los datos normalizados")
                    st.dataframe(preview, use_container_width=True)
                    if st.button("Procesar e importar", type="primary"):
                        store_import(preview, params, safe_name, month_override=month_key)
                        st.success(f"Importación completada y guardada permanentemente en GitHub para {month_key}.")
                        st.rerun()
                except Exception as exc:
                    st.error(f"No se pudo leer el Excel: {exc}")

            st.markdown("---")
            st.subheader("🗑️ Gestión y eliminación de archivos guardados en GitHub")
            try:
                repo = get_github_repo()
                contents = repo.get_contents("")
                excel_files = [f.path for f in contents if f.path.endswith(".xlsx")]
                if not excel_files:
                    st.info("No hay archivos Excel guardados actualmente en GitHub.")
                else:
                    for f_name in excel_files:
                        col1, col2 = st.columns([4, 1])
                        with col1:
                            st.text(f_name)
                        with col2:
                            if st.button("Eliminar", key=f"del_{f_name}", type="secondary"):
                                delete_excel_from_github(f_name)
                                st.rerun()
            except Exception as e:
                st.error(f"No se pudieron cargar los archivos: {e}")

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
                    try:
                        guardar_asignaciones_en_github()
                    except Exception:
                        pass
                    st.success("Asignaciones guardadas.")
                    st.rerun()

        with tabs[2]:
            st.subheader("Parámetros del sistema y baremos de puntuación")
            st.markdown("Modifica los umbrales y los puntos otorgados en cada categoría (el sistema opera exclusivamente sumando puntos).")
            
            with st.form("params_form"):
                newp = params.copy()
                
                st.markdown("#### 🟢 Ralentí")
                col1, col2, col3 = st.columns(3)
                with col1:
                    newp["idle_optimal_pct"] = st.number_input("Ralentí óptimo máx (%)", value=float(params["idle_optimal_pct"]), step=0.5)
                    newp["idle_points_optimal"] = st.number_input("Puntos ralentí óptimo", value=float(params["idle_points_optimal"]), step=1.0)
                with col2:
                    newp["idle_moderate_pct"] = st.number_input("Ralentí moderado máx (%)", value=float(params["idle_moderate_pct"]), step=0.5)
                    newp["idle_points_moderate"] = st.number_input("Puntos ralentí moderado", value=float(params["idle_points_moderate"]), step=1.0)
                with col3:
                    newp["idle_points_high"] = st.number_input("Puntos ralentí alto", value=float(params["idle_points_high"]), step=1.0)

                st.markdown("#### 🛣️ Velocidad")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    newp["speed_min"] = st.number_input("Velocidad mín (km/h)", value=float(params["speed_min"]), step=1.0)
                with col2:
                    newp["speed_max"] = st.number_input("Velocidad máx (km/h)", value=float(params["speed_max"]), step=1.0)
                with col3:
                    newp["speed_points_good"] = st.number_input("Puntos velocidad óptima", value=float(params["speed_points_good"]), step=1.0)
                with col4:
                    newp["speed_points_near"] = st.number_input("Puntos velocidad cercana", value=float(params["speed_points_near"]), step=1.0)

                st.markdown("#### 🛑 Paradas, Distancia y Duración")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    newp["stops_good_max"] = st.number_input("Máx paradas óptimas", value=float(params["stops_good_max"]), step=1.0)
                    newp["stops_points_good"] = st.number_input("Puntos paradas óptimas", value=float(params["stops_points_good"]), step=1.0)
                with col2:
                    newp["stops_points_other"] = st.number_input("Puntos otras paradas", value=float(params["stops_points_other"]), step=1.0)
                    newp["distance_reference_km"] = st.number_input("Km de referencia", value=float(params["distance_reference_km"]), step=10.0)
                with col3:
                    newp["distance_max_points"] = st.number_input("Puntos máx distancia", value=float(params["distance_max_points"]), step=1.0)
                    newp["duration_target_pct"] = st.number_input("Duración objetivo (%)", value=float(params["duration_target_pct"]), step=1.0)
                with col4:
                    newp["duration_points_good"] = st.number_input("Puntos duración óptima", value=float(params["duration_points_good"]), step=1.0)
                    newp["duration_points_other"] = st.number_input("Puntos otra duración", value=float(params.get("duration_points_other", 5.0)), step=1.0)

                st.markdown("#### 🌍 Factores ambientales y consumo")
                col1, col2, col3 = st.columns(3)
                with col1:
                    newp["fuel_l_per_100km"] = st.number_input("Consumo referencia (L/100km)", value=float(params["fuel_l_per_100km"]), step=1.0)
                with col2:
                    newp["idle_l_per_hour"] = st.number_input("Consumo ralentí (L/h)", value=float(params["idle_l_per_hour"]), step=0.1)
                with col3:
                    newp["co2_kg_per_litre"] = st.number_input("CO2 (kg/litro)", value=float(params["co2_kg_per_litre"]), step=0.01)

                submitted = st.form_submit_button("Guardar todos los parámetros", type="primary")
                if submitted:
                    save_params(newp)
                    st.success("¡Parámetros del sistema actualizados correctamente!")
                    st.rerun()

        dashboard_tab, impact_tab, ranking_tab, simulator_tab, badges_tab = tabs[2], tabs[3], tabs[4], tabs[5], tabs[6]
    else:
        dashboard_tab, impact_tab, ranking_tab, simulator_tab, badges_tab = st.tabs([
            "📊 Dashboard", "🌍 Impacto ambiental", "🏆 Clasificación", "🎛️ Simulador", "🎖️ Insignias"
        ])

    with dashboard_tab:
        st.subheader("Evolución y rendimiento de la flota")
        if records.empty:
            st.info("Todavía no hay datos. El administrador debe importar el Excel mensual.")
        else:
            monthly = records.groupby("month", as_index=False).agg(
                puntos=("points_total", "sum"),
                distancia_km=("distance_km", "sum"),
                vehículos=("vehicle", "nunique"),
            ).sort_values("month")
            k1, k2, k3 = st.columns(3)
            k1.metric("Puntos acumulados", f'{records["points_total"].sum():,.0f}')
            k2.metric("Kilómetros registrados", f'{records["distance_km"].sum():,.0f} km')
            k3.metric("Vehículos registrados", f'{records["vehicle"].nunique()}')
            fig = px.line(monthly, x="month", y="puntos", markers=True, title="Puntos positivos acumulados por mes")
            st.plotly_chart(fig, use_container_width=True)

    with ranking_tab:
        st.subheader("🏆 Clasificación de Conductores")
        if records.empty:
            records = load_records()

        if not records.empty and "driver" in records.columns:
            meses_disponibles = sorted(records["month"].unique().tolist(), reverse=True)
            if not meses_disponibles:
                st.info("No hay meses registrados todavía.")
            else:
                mes_seleccionado = st.selectbox("📅 Selecciona el mes a consultar:", meses_disponibles, key="filtro_mes_independiente")
                df_mes = records[records["month"] == mes_seleccionado].copy()
                df_ranking = df_mes.groupby("driver", as_index=False)["points_total"].sum()
                df_ranking = df_ranking.sort_values(by="points_total", ascending=False).reset_index(drop=True)
                
                st.markdown("---")
                if df_ranking.empty:
                    st.info(f"No hay registros de conducción para el periodo {mes_seleccionado}.")
                else:
                    st.markdown(f"#### 📊 Clasificación oficial del mes **{mes_seleccionado}** (Puntos independientes)")
                    for index, row in df_ranking.iterrows():
                        conductor = row["driver"]
                        puntos = row["points_total"]
                        
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
                            
                        if index < 5:
                            st.markdown(f"### {insignia} — **{conductor}**: {puntos:.0f} puntos")
                        else:
                            st.write(f"{insignia} — **{conductor}**: {puntos:.0f} puntos")
        else:
            st.info("Todavía no hay registros cargados en la base de datos.")

    with impact_tab:
        st.subheader("Impacto ambiental y ahorro estimado")
        if records.empty:
            st.info("Importa datos para calcular estimaciones.")
        else:
            fuel_saved, co2_saved, reference = environmental_summary(records, params)
            a, b, c = st.columns(3)
            a.metric("Combustible potencialmente ahorrado", f"{fuel_saved:,.1f} L")
            b.metric("CO₂ potencialmente evitado", f"{co2_saved:,.1f} kg")
            c.metric("Consumo de referencia estimado", f"{reference:,.1f} L")

    with simulator_tab:
                st.subheader("Simulador avanzado de mejora de Ecodriving")
                st.markdown("Ajusta los parámetros de conducción para simular en tiempo real el impacto en el ahorro de combustible y la reducción de emisiones de CO₂ de la flota.")
                
                col1, col2 = st.columns(2)
                with col1:
                    sim_idle_reduction = st.slider("Reducción porcentual del tiempo en ralentí (%)", 0.0, 50.0, 10.0, 0.5)
                    sim_speed_opt = st.slider("Mejora en adherencia a velocidad eficiente (%)", 0.0, 100.0, 20.0, 1.0)
                with col2:
                    sim_stops_reduction = st.slider("Reducción en número de paradas excesivas (%)", 0.0, 50.0, 15.0, 1.0)
                    sim_distance_factor = st.slider("Optimización de rutas y distancia (%)", 0.0, 30.0, 5.0, 0.5)
                
                if not records.empty:
                    base_distance = records["distance_km"].sum()
                    base_idle_hours = records["idle_hours"].sum() if "idle_hours" in records.columns else (base_distance * 0.05)
                    
                    idle_l_h = params.get("idle_l_per_hour", 2.5)
                    co2_factor = params.get("co2_kg_per_litre", 2.68)
                    
                    ahorro_litros_ralenti = (base_idle_hours * (sim_idle_reduction / 100.0)) * idle_l_h
                    ahorro_litros_general = base_distance * (sim_distance_factor / 100.0) * 0.08
                    
                    total_ahorro_litros = ahorro_litros_ralenti + ahorro_litros_general
                    total_co2_evitado = total_ahorro_litros * co2_factor
                else:
                    total_ahorro_litros = 286.4 * (sim_idle_reduction / 10.0)
                    total_co2_evitado = 767.6 * (sim_idle_reduction / 10.0)

                st.markdown("---")
                res_col1, res_col2 = st.columns(2)
                with res_col1:
                    st.metric(
                        label="Ahorro de combustible simulado",
                        value=f"{total_ahorro_litros:.1f} L",
                        delta=f"+{total_ahorro_litros * 0.17:.1f} L vs base"
                    )
                with res_col2:
                    st.metric(
                        label="CO₂ evitado simulado",
                        value=f"{total_co2_evitado:.1f} kg",
                        delta=f"+{total_co2_evitado * 0.17:.1f} kg vs base"
                    )

    with badges_tab:
            st.subheader("🎖️ Insignias y logros de los conductores")
            if records.empty:
                st.info("No hay datos de conducción registrados para mostrar insignias.")
            else:
                drivers_list = sorted([d for d in records["driver"].unique() if d and d != "Sin asignar"])
                if not drivers_list:
                    st.info("Todavía no hay conductores asignados a los vehículos.")
                else:
                    sel_driver_badge = st.selectbox("Selecciona un conductor para ver sus insignias:", drivers_list, key="select_driver_badges")
                    df_driver_records = records[records["driver"] == sel_driver_badge]
                    
                    st.markdown(f"### Logros obtenidos por **{sel_driver_badge}**")
                    for _, r in df_driver_records.iterrows():
                        b_list = driver_badges(r)
                        with st.container():
                            st.markdown(f"**Periodo / Mes: {r['month']}** (Vehículo: {r['vehicle']})")
                            if b_list:
                                for b in b_list:
                                    st.success(f"{b} (Puntos totales del mes: {r['points_total']:.0f})")
                            else:
                                st.info(f"Sin insignias específicas este mes (Puntos: {r['points_total']:.0f}). ¡Ánimo para el siguiente!")
                            st.markdown("---")

    st.sidebar.markdown("---")
    if st.sidebar.button("💾 Guardar asignaciones en GitHub"):
        try:
            guardar_asignaciones_en_github()
            st.sidebar.success("¡Guardado en GitHub!")
        except Exception as e:
            st.sidebar.error(f"Error: {e}")

if __name__ == "__main__":
    main()

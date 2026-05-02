"""
app.py — Algoritmo Prescriptivo Táctico · LaLiga 2025-26
=========================================================
Aplicación Streamlit que, dados los parámetros de un partido próximo
(equipos, árbitro, lluvia, alineación propia), genera un informe táctico
prescriptivo en cinco bloques:

  Bloque 1 · Resumen General      → predicción ML + IS + forma reciente + H2H
  Bloque 2 · Estrategia Ofensiva  → cómo atacar al rival
  Bloque 3 · Estrategia Defensiva → cómo neutralizar al rival
  Bloque 4 · Jugadores Clave      → alineación estimada rival + calidad de plantilla
  Bloque 5 · Contexto             → árbitro y condiciones meteorológicas

Predicción vía Random Forest (Modelo B, 111 features diferenciales):
  ŷ > +0.5  → Victoria probable   |  ŷ ∈ [−0.5, +0.5] → Empate probable
  ŷ < −0.5  → Derrota probable    |  Accuracy: 49.6% (baseline azar: 36.3%)

Fuentes de datos:
  Base de Datos.xlsx  (Equipos, Clasificación, Partidos, Lluvias, Árbitros,
    Jugadores, Porteros, Lesionados y Sancionados, Duplas Peligrosas)
  Datos WhoScored.xlsx       → features tácticas de equipos (85 métricas)
  Jugadores Unificados.xlsx  → features de portero y plantilla para el modelo
  modelos/modelo_b_rf.pkl    → Random Forest entrenado (670 partidos, K-Fold 5)
  modelos/feature_names.pkl  → orden exacto de los 111 features

Uso: streamlit run app.py
"""

import base64
import pickle
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ══════════════════════════════════════════════════════════════════
# RUTAS
# ══════════════════════════════════════════════════════════════════
_BASE        = Path(__file__).parent
EXCEL        = str(_BASE / "Base de datos.xlsx")
WS_PATH      = str(_BASE / "Tablas Excel" / "Datos WhoScored.xlsx")
JUG_UNIF_PATH= str(_BASE / "Tablas Excel" / "Jugadores Unificados.xlsx")
MODEL_PATH   = str(_BASE / "modelos" / "modelo_b_rf.pkl")
FEAT_PATH    = str(_BASE / "modelos" / "feature_names.pkl")

FORMACIONES = ["4-3-3", "4-4-2", "4-2-3-1", "4-1-4-1", "3-4-3",
               "3-5-2", "5-3-2", "5-4-1", "4-5-1", "4-3-2-1"]

FORMACIONES_SLOTS = {
    "4-3-3": [
        ("Portero",           "portero"),
        ("Lateral Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Lateral Izquierdo", "defensa"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Extremo Derecho",   "extremo"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
    ],
    "4-4-2": [
        ("Portero",           "portero"),
        ("Lateral Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Lateral Izquierdo", "defensa"),
        ("Extremo Derecho",   "extremo"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
        ("Delantero Centro",  "delantero"),
    ],
    "4-2-3-1": [
        ("Portero",           "portero"),
        ("Lateral Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Lateral Izquierdo", "defensa"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Extremo Derecho",   "extremo"),
        ("Mediapunta",        "centro"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
    ],
    "4-1-4-1": [
        ("Portero",           "portero"),
        ("Lateral Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Lateral Izquierdo", "defensa"),
        ("Pivote",            "centro"),
        ("Extremo Derecho",   "extremo"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
    ],
    "3-4-3": [
        ("Portero",           "portero"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Extremo Derecho",   "extremo"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
    ],
    "3-5-2": [
        ("Portero",              "portero"),
        ("Defensa Central",      "defensa"),
        ("Defensa Central",      "defensa"),
        ("Defensa Central",      "defensa"),
        ("Carrilero Derecho",    "defensa"),
        ("Centrocampista",       "centro"),
        ("Centrocampista",       "centro"),
        ("Centrocampista",       "centro"),
        ("Carrilero Izquierdo",  "defensa"),
        ("Delantero Centro",     "delantero"),
        ("Delantero Centro",     "delantero"),
    ],
    "5-3-2": [
        ("Portero",              "portero"),
        ("Carrilero Derecho",    "defensa"),
        ("Defensa Central",      "defensa"),
        ("Defensa Central",      "defensa"),
        ("Defensa Central",      "defensa"),
        ("Carrilero Izquierdo",  "defensa"),
        ("Centrocampista",       "centro"),
        ("Centrocampista",       "centro"),
        ("Centrocampista",       "centro"),
        ("Delantero Centro",     "delantero"),
        ("Delantero Centro",     "delantero"),
    ],
    "5-4-1": [
        ("Portero",              "portero"),
        ("Carrilero Derecho",    "defensa"),
        ("Defensa Central",      "defensa"),
        ("Defensa Central",      "defensa"),
        ("Defensa Central",      "defensa"),
        ("Carrilero Izquierdo",  "defensa"),
        ("Extremo Derecho",      "extremo"),
        ("Centrocampista",       "centro"),
        ("Centrocampista",       "centro"),
        ("Extremo Izquierdo",    "extremo"),
        ("Delantero Centro",     "delantero"),
    ],
    "4-5-1": [
        ("Portero",           "portero"),
        ("Lateral Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Lateral Izquierdo", "defensa"),
        ("Extremo Derecho",   "extremo"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
    ],
    "4-3-2-1": [
        ("Portero",           "portero"),
        ("Lateral Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Lateral Izquierdo", "defensa"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Mediapunta",        "centro"),
        ("Mediapunta",        "centro"),
        ("Delantero Centro",  "delantero"),
    ],
}

# ── Normalización de nombres de equipo ────────────────────────────
TEAM_ALIASES = {
    "Deportivo Alaves":   "Alavés",
    "Athletic Bilbao":    "Athletic",
    "Atletico Madrid":    "Atlético",
    "Celta Vigo":         "Celta",
    "Rayo Vallecano":     "Rayo",
    "Real Betis":         "Betis",
    "Real Madrid":        "Real Madrid",
    "Real Sociedad":      "Real Sociedad",
    "UD Las Palmas":      "Las Palmas",
    "Valencia CF":        "Valencia",
}

def normalize_team(name: str) -> str:
    """
    Traduce el nombre de un equipo al formato unificado que usa el resto de la base de datos.

    WhoScored usa nombres diferentes a FBref y estadisticaslaliga.es (p. ej. "Athletic Bilbao"
    frente a "Athletic"). Esta función aplica el diccionario TEAM_ALIASES para que todos los
    módulos hablen el mismo idioma al cruzar tablas.
    """
    if pd.isna(name):
        return ""
    s = str(name).strip()
    return TEAM_ALIASES.get(s, s)


# ══════════════════════════════════════════════════════════════════
# FEATURES WHOSCORED — 85 columnas (mismo orden que modelos.py)
# ══════════════════════════════════════════════════════════════════
WS_COLS = [
    "rating_gen", "tiros_pp_gen", "tirosAP_pp_gen", "regates_pp_gen",
    "faltasFavor_pp_gen", "fueraJuego_pp_gen", "aereos_gen", "aciertoPasePct_gen",
    "xG_fav_gen", "xGDif_fav_gen", "tiros_fav_gen", "xGTiros_fav_gen",
    "xG_con_gen", "xGDif_con_gen",
    "tiros_contra_gen", "entradas_pp_gen", "intercep_pp_gen", "faltas_pp_gen",
    "entrada_exito_gen", "entrada_fallo_gen", "intercepciones_gen",
    "despejes_gen", "bloqueados_tiros_gen", "bloqueados_centros_gen", "bloqueados_pases_gen",
    "paradas_total_gen", "paradas_pequeña_gen", "paradas_area_gen", "paradas_fuera_gen",
    "balonesAereos_ganados_gen", "balonesAereos_perdidos_gen",
    "pases_total_gen", "pases_largosPrecisos_gen", "pases_largosImprecisos_gen",
    "pases_cortosPrecisos_gen", "pases_cortosImprecisos_gen",
    "pasesClave_corto_gen", "pasesClave_largo_gen",
    "pase_centros_fav_gen", "pase_alHueco_fav_gen",
    "pase_centros_con_gen", "pase_alHueco_con_gen",
    "asistencias_centro_gen", "asistencias_corner_gen", "asistencias_alHueco_gen",
    "asistencias_tiroLibre_gen", "asistencias_banda_gen",
    "gol_juegoAbierto_fav_gen", "gol_contraataque_fav_gen",
    "gol_balonParado_fav_gen", "gol_penalty_fav_gen",
    "gol_juegoAbierto_con_gen", "gol_contraataque_con_gen",
    "gol_balonParado_con_gen", "gol_penalty_con_gen",
    "goles_areaPequeña_gen", "goles_area_gen", "goles_fueraArea_gen",
    "tiros_areaPequeña_gen", "tiros_area_gen", "tiros_fueraArea_gen",
    "zonasTiro_areaPequeña_fav_gen", "zonasTiro_areaPenalty_fav_gen", "zonasTiro_areaFuera_fav_gen",
    "zonasTiro_areaPequeña_con_gen", "zonasTiro_areaPenalty_con_gen", "zonasTiro_areaFuera_con_gen",
    "direccionesTiro_izquierda_fav_gen", "direccionesTiro_centro_fav_gen", "direccionesTiro_derecha_fav_gen",
    "direccionesTiro_izquierda_con_gen", "direccionesTiro_centro_con_gen", "direccionesTiro_derecha_con_gen",
    "zonas_accion_ataque_gen", "zonas_accion_mediocampo_gen", "zonas_accion_defensa_gen",
    "tarjetas_amarilla_gen", "tarjetas_roja_gen",
    "perdida_desposeido_gen", "perdida_toqueFallido_gen",
    "regates_exitosos_gen", "regates_fallidos_gen",
    "ladosAtaque_izquierda_gen", "ladosAtaque_centro_gen", "ladosAtaque_derecha_gen",
]


# ══════════════════════════════════════════════════════════════════
# UTILIDADES NUMÉRICAS
# ══════════════════════════════════════════════════════════════════
def pct_num(series):
    """Convierte columnas de porcentaje (p. ej. "34%") a número decimal."""
    return (series.astype(str)
            .str.replace("%", "", regex=False)
            .pipe(pd.to_numeric, errors="coerce"))

def minmax_series(s):
    """
    Normaliza una serie entre 0 y 1 usando mínimo-máximo.

    Si todos los valores son iguales devuelve 0.5 para evitar divisiones por cero.
    Se usa para calcular el Índice de Éxito (IS) del equipo.
    """
    mn, mx = s.min(), s.max()
    if mx > mn:
        return (s - mn) / (mx - mn)
    return pd.Series(0.5, index=s.index)

def minmax_group(series, groups):
    """
    Normaliza una serie entre 0 y 1 de forma independiente dentro de cada grupo.

    En lugar de comparar a todos los jugadores entre sí, compara cada jugador
    solo con los de su misma posición (porteros vs porteros, defensas vs defensas, etc.)
    para que el IS individual sea justo independientemente del rol.
    """
    result = pd.Series(0.0, index=series.index)
    for g in groups.unique():
        mask = groups == g
        s = series[mask]
        mn, mx = s.min(), s.max()
        result[mask] = (s - mn) / (mx - mn) if mx > mn else 0.5
    return result

def _to_num(s):
    """Convierte una serie a numérico ignorando errores (celdas vacías o texto quedan como NaN)."""
    return pd.to_numeric(s, errors="coerce")


# ══════════════════════════════════════════════════════════════════
# CARGA DE DATOS (cacheada)
# ══════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Cargando base de datos…")
def load_data():
    """
    Lee y procesa todas las hojas de la Base de Datos.xlsx y los archivos complementarios.

    Las hojas no tienen cabecera estándar, por lo que se accede a cada columna por índice
    numérico (documentado en el CLAUDE.md). El resultado se devuelve como un conjunto de
    DataFrames ya limpios, normalizados y cruzados, listos para usar en el informe.

    Se cachea con @st.cache_data para que Streamlit no relea el Excel en cada interacción
    del usuario — solo al arrancar la app o cuando cambia el fichero.

    Devuelve:
        df            — métricas de equipo + clasificación + IS calculado
        teams         — lista ordenada de nombres de equipo
        arb           — estadísticas por árbitro
        arbitros      — lista de nombres de árbitros
        jug           — estadísticas individuales de jugadores + porteros con IS_indiv
        les           — jugadores lesionados o sancionados
        ausentes      — lista de jugadores con probabilidad 0% de jugar
        dup           — duplas goleador-asistidor por equipo
        ws_df         — 85 métricas WhoScored por equipo (para el modelo ML)
        por_feats_ml  — métricas del portero titular de cada equipo (para el modelo ML)
        jug_feats_ml  — resumen de la plantilla de cada equipo (para el modelo ML)
        partidos_df   — historial de partidos con resultado
        forma_actual  — forma reciente (últimos 5 partidos) por equipo
    """
    # ── Equipos ──────────────────────────────────────────────────
    raw_eq = pd.read_excel(EXCEL, sheet_name="Equipos", header=None)
    data   = raw_eq.iloc[4:24].copy().reset_index(drop=True)

    eq = pd.DataFrame({
        "Equipo":          data.iloc[:, 0].apply(normalize_team),
        "Pos%":            pd.to_numeric(data.iloc[:, 3],   errors="coerce") * 100,
        "Rating":          pd.to_numeric(data.iloc[:, 4],   errors="coerce"),
        "Goles":           pd.to_numeric(data.iloc[:, 7],   errors="coerce"),
        "xG":              pd.to_numeric(data.iloc[:, 57],  errors="coerce"),
        "xGDif":           pd.to_numeric(data.iloc[:, 60],  errors="coerce"),
        "xG_contra":       pd.to_numeric(data.iloc[:, 63],  errors="coerce"),
        "xGDif_contra":    pd.to_numeric(data.iloc[:, 66],  errors="coerce"),
        "xG_tiro":         pd.to_numeric(data.iloc[:, 69],  errors="coerce"),
        "Tiros_pp":        pd.to_numeric(data.iloc[:, 81],  errors="coerce"),
        "Tiros_contra_pp": pd.to_numeric(data.iloc[:, 84],  errors="coerce"),
        "Tiros_puerta_pp": pd.to_numeric(data.iloc[:, 87],  errors="coerce"),
        "Pct_tiro_area":   pct_num(data.iloc[:, 102]),
        "Entradas_exit":   pd.to_numeric(data.iloc[:, 170], errors="coerce"),
        "Entradas_fall":   pd.to_numeric(data.iloc[:, 173], errors="coerce"),
        "Intercepciones":  pd.to_numeric(data.iloc[:, 176], errors="coerce"),
        "Despejes":        pd.to_numeric(data.iloc[:, 179], errors="coerce"),
        "Bloqueos_tiro":   pd.to_numeric(data.iloc[:, 182], errors="coerce"),
        "Aereos_gan":      pd.to_numeric(data.iloc[:, 191], errors="coerce"),
        "Aereos_per":      pd.to_numeric(data.iloc[:, 194], errors="coerce"),
        "Total_pases":     pd.to_numeric(data.iloc[:, 197], errors="coerce"),
        "Precision_pase":  pd.to_numeric(data.iloc[:, 200], errors="coerce"),
        "Pases_L_prec":    pd.to_numeric(data.iloc[:, 209], errors="coerce"),
        "Pases_L_impr":    pd.to_numeric(data.iloc[:, 212], errors="coerce"),
        "PasesClave_cort": pd.to_numeric(data.iloc[:, 224], errors="coerce"),
        "Zona_Def":        pct_num(data.iloc[:, 266]),
        "Zona_Med":        pct_num(data.iloc[:, 269]),
        "Zona_Ata":        pct_num(data.iloc[:, 272]),
        "Perdidas":        pd.to_numeric(data.iloc[:, 275], errors="coerce")
                           + pd.to_numeric(data.iloc[:, 278], errors="coerce"),
        "Faltas_com":      pd.to_numeric(data.iloc[:, 281], errors="coerce"),
        "Amarillas":       pd.to_numeric(data.iloc[:, 284], errors="coerce"),
        "Rojas":           pd.to_numeric(data.iloc[:, 287], errors="coerce"),
        "Paradas_pp":      pd.to_numeric(data.iloc[:, 152], errors="coerce"),
        "SavePct":         pd.to_numeric(data.iloc[:, 153], errors="coerce"),
        "P0":              pd.to_numeric(data.iloc[:, 154], errors="coerce"),
        "Reg_exit":        pd.to_numeric(data.iloc[:, 138], errors="coerce"),
    }).reset_index(drop=True)

    eq["Efic_entrada"] = eq["Entradas_exit"] / (eq["Entradas_exit"] + eq["Entradas_fall"]) * 100
    eq["Pct_aereo"]    = eq["Aereos_gan"] / (eq["Aereos_gan"] + eq["Aereos_per"]) * 100

    # ── Clasificación ─────────────────────────────────────────────
    raw_cla = pd.read_excel(EXCEL, sheet_name="Clasificación", header=None)
    cla = pd.DataFrame({
        "Equipo": raw_cla.iloc[3:23, 1].apply(normalize_team).values,
        "PJ":     _to_num(raw_cla.iloc[3:23, 2]),
        "GF":     _to_num(raw_cla.iloc[3:23, 6]),
        "GC":     _to_num(raw_cla.iloc[3:23, 7]),
        "Pts":    _to_num(raw_cla.iloc[3:23, 9]),
    }).reset_index(drop=True)

    # ── Lluvias ───────────────────────────────────────────────────
    raw_ll  = pd.read_excel(EXCEL, sheet_name="Lluvias", header=None)
    lluvia = pd.DataFrame({
        "Equipo":      raw_ll.iloc[2:22, 0].apply(normalize_team).values,
        "Total_mm":    _to_num(raw_ll.iloc[2:22, 3]),
        "Dias_lluvia": _to_num(raw_ll.iloc[2:22, 4]),
    }).reset_index(drop=True)

    # ── Merge maestro ─────────────────────────────────────────────
    df = eq.merge(cla, on="Equipo", how="inner").merge(lluvia, on="Equipo", how="inner")
    df = df.loc[:, ~df.columns.duplicated()].copy()
    df["Pts_pp"] = df["Pts"] / df["PJ"]
    df["xG_pp"]  = df["xG"]  / df["PJ"]
    df["GC_pp"]  = df["GC"]  / df["PJ"]
    df["GF_pp"]  = df["GF"]  / df["PJ"]
    df["GD_pp"]  = (df["GF"] - df["GC"]) / df["PJ"]
    df["IS"] = (
        0.35 * minmax_series(df["xG_pp"]) +
        0.35 * minmax_series(-df["GC_pp"]) +
        0.30 * minmax_series(df["Rating"])
    )

    teams = sorted(df["Equipo"].dropna().unique().tolist())

    # ── Árbitros ──────────────────────────────────────────────────
    raw_arb = pd.read_excel(EXCEL, sheet_name="Árbitros", header=None)
    arb = pd.DataFrame({
        "Arbitro":     raw_arb.iloc[2:, 0].ffill().values,
        "Partidos":    _to_num(raw_arb.iloc[2:, 2]),
        "FaltasPP":    _to_num(raw_arb.iloc[2:, 4]),
        "AmarillasPP": _to_num(raw_arb.iloc[2:, 6]),
        "RojasPP":     _to_num(raw_arb.iloc[2:, 8]),
        "PenaltisPP":  _to_num(raw_arb.iloc[2:, 12]),
        "PctLocal":    _to_num(raw_arb.iloc[2:, 13]) * 100,
        "PctEmpate":   _to_num(raw_arb.iloc[2:, 14]) * 100,
        "PctVisitante":_to_num(raw_arb.iloc[2:, 15]) * 100,
    })
    arb = arb[arb["Partidos"].notna()].copy()
    arb["Arbitro"] = arb["Arbitro"].apply(lambda x: str(x).strip() if pd.notna(x) else "")
    arb = arb[arb["Arbitro"] != ""].reset_index(drop=True)
    arbitros = sorted(arb["Arbitro"].unique().tolist())

    # ── Porteros ──────────────────────────────────────────────────
    raw_por = pd.read_excel(EXCEL, sheet_name="Porteros", header=None)
    por = pd.DataFrame({
        "Jugador":  raw_por.iloc[3:, 0].values,
        "Posicion": "Portero",
        "Equipo":   raw_por.iloc[3:, 3].apply(normalize_team).values,
        "Minutos":  _to_num(raw_por.iloc[3:, 20]),
        "Rating":   _to_num(raw_por.iloc[3:,  5]),
    }).dropna(subset=["Minutos", "Equipo"])
    por["PosGrupo"]   = "Portero"
    por["OfScore90"]  = 0.0
    por["DefScore90"] = 0.0
    por["norm_of"]    = 0.5
    por["norm_def"]   = 0.5

    # ── Jugadores ─────────────────────────────────────────────────
    raw_jug = pd.read_excel(EXCEL, sheet_name="Jugadores", header=None)
    jug = pd.DataFrame({
        "Jugador":   raw_jug.iloc[3:, 0].values,
        "Posicion":  raw_jug.iloc[3:, 2].values,
        "Equipo":    raw_jug.iloc[3:, 3].apply(normalize_team).values,
        "Minutos":   _to_num(raw_jug.iloc[3:, 20]),
        "PctMin":    _to_num(raw_jug.iloc[3:, 24]),
        "Rating":    _to_num(raw_jug.iloc[3:,  5]),
        "xG_90":     _to_num(raw_jug.iloc[3:, 139]),
        "Asist_90":  _to_num(raw_jug.iloc[3:,  43]),
        "Entradas":  _to_num(raw_jug.iloc[3:, 235]),
        "Interc":    _to_num(raw_jug.iloc[3:, 241]),
        "Despejes":  _to_num(raw_jug.iloc[3:, 244]),
        "Bloqueos":  _to_num(raw_jug.iloc[3:, 247]),
    }).dropna(subset=["Minutos", "Equipo"])
    jug = jug[jug["Minutos"] >= 400].reset_index(drop=True)

    _POR = {"GK","P","PT","POR","GOR"}
    _DEF = {"DC","DFC","DL","DR","LD","LI","CB","LB","RB","DF","DFD","DFI"}
    _DEL = {"FC","FW","SS","ST","CF","EI","ED","EC","DEL","AT","FWL","FWR"}

    def pos_group(pos_str):
        if pd.isna(pos_str) or str(pos_str).strip() == "":
            return "Centrocampista"
        p = str(pos_str).split(",")[0].strip().upper()
        if p in _POR: return "Portero"
        if p in _DEL or p.startswith("F") or p.startswith("SS"): return "Delantero"
        if p.startswith("E"): return "Delantero"
        if p in _DEF or p.startswith("L"): return "Defensa"
        return "Centrocampista"

    jug["PosGrupo"]  = jug["Posicion"].apply(pos_group)
    jug["OfScore90"] = jug["xG_90"].fillna(0) + jug["Asist_90"].fillna(0)
    jug["DefScore90"] = (
        jug[["Entradas","Interc","Despejes","Bloqueos"]]
        .fillna(0).sum(axis=1) / (jug["Minutos"] / 90)
    )
    rat_f = jug["Rating"].fillna(jug.groupby("PosGrupo")["Rating"].transform("median"))
    jug["norm_of"]   = minmax_group(jug["OfScore90"],  jug["PosGrupo"])
    jug["norm_def"]  = minmax_group(jug["DefScore90"], jug["PosGrupo"])
    jug["norm_rat"]  = minmax_group(rat_f,             jug["PosGrupo"])
    jug["IS_indiv"]  = (0.35 * jug["norm_of"]
                        + 0.35 * jug["norm_def"]
                        + 0.30 * jug["norm_rat"])
    min_eq = jug.groupby("Equipo")["Minutos"].sum().rename("Min_equipo")
    jug = jug.merge(min_eq, on="Equipo", how="left")
    jug["IS_contrib"] = jug["IS_indiv"] * (jug["Minutos"] / jug["Min_equipo"])

    rat_por = por["Rating"].fillna(por["Rating"].median())
    por["norm_rat"]  = minmax_group(rat_por, por["PosGrupo"])
    por["IS_indiv"]  = 0.30 * por["norm_rat"]
    min_eq_por = jug.groupby("Equipo")["Minutos"].sum().rename("Min_equipo")
    por = por.merge(min_eq_por, on="Equipo", how="left")
    por["IS_contrib"] = por["IS_indiv"] * (por["Minutos"] / por["Min_equipo"].fillna(1))
    for col in ["PctMin", "xG_90", "Asist_90", "Entradas", "Interc", "Despejes", "Bloqueos"]:
        por[col] = np.nan
    jug = pd.concat([jug, por], ignore_index=True)

    # ── Lesionados y Sancionados ──────────────────────────────────
    raw_les = pd.read_excel(EXCEL, sheet_name="Lesionados y Sancionados", header=None)
    les = pd.DataFrame({
        "Jugador":   raw_les.iloc[2:, 0].values,
        "Equipo":    raw_les.iloc[2:, 1].apply(normalize_team).values,
        "Motivo":    raw_les.iloc[2:, 2].values,
        "Vuelta":    raw_les.iloc[2:, 5].values,
        "ProbJugar": _to_num(raw_les.iloc[2:, 6]),
    }).dropna(subset=["Jugador"]).reset_index(drop=True)
    les["Jugador"] = les["Jugador"].astype(str).str.strip()
    les["Equipo"]  = les["Equipo"].astype(str).str.strip()
    ausentes = les[les["ProbJugar"] == 0]["Jugador"].tolist()

    # ── Duplas Peligrosas ─────────────────────────────────────────
    raw_dup = pd.read_excel(EXCEL, sheet_name="Duplas Peligrosas", header=None)
    dup = pd.DataFrame({
        "Goleador":   raw_dup.iloc[1:, 1].values,
        "Asistidor":  raw_dup.iloc[1:, 2].values,
        "Equipo":     raw_dup.iloc[1:, 3].apply(normalize_team).values,
        "Frecuencia": _to_num(raw_dup.iloc[1:, 4]),
    }).dropna(subset=["Frecuencia"]).reset_index(drop=True)

    # ── WhoScored features (para modelo ML y recomendaciones) ────
    ws_df = pd.DataFrame()
    try:
        raw_ws = pd.read_excel(WS_PATH, header=0)
        raw_ws["equipo"] = raw_ws["equipo"].apply(normalize_team)
        raw_ws = raw_ws.set_index("equipo")
        cols_disp = [c for c in WS_COLS if c in raw_ws.columns]
        ws_df = raw_ws[cols_disp].copy()
        for col in ws_df.columns:
            ws_df[col] = _to_num(ws_df[col].astype(str).str.replace("%", ""))
    except Exception:
        pass

    # ── Portero features para modelo ML ──────────────────────────
    por_feats_ml = pd.DataFrame()
    try:
        raw_pm = pd.read_excel(JUG_UNIF_PATH, sheet_name="Porteros", header=None)
        df_pm = pd.DataFrame({
            "Jugador":       raw_pm.iloc[3:, 0].astype(str).values,
            "Equipo":        raw_pm.iloc[3:, 3].apply(normalize_team).values,
            "Min":           _to_num(raw_pm.iloc[3:, 7]),
            "por_GA90":      _to_num(raw_pm.iloc[3:, 22]),
            "por_pct_par":   _to_num(raw_pm.iloc[3:, 25]),
            "por_pct_p0":    _to_num(raw_pm.iloc[3:, 30]),
            "por_rating":    _to_num(raw_pm.iloc[3:, 15]),
            "por_par_AreaPeq":   _to_num(raw_pm.iloc[3:, 31]),
            "por_par_AreaPen":   _to_num(raw_pm.iloc[3:, 34]),
            "por_par_FueraArea": _to_num(raw_pm.iloc[3:, 37]),
            "por_pct_pk":    _to_num(raw_pm.iloc[3:, 44]),
            "por_pct_pases": _to_num(raw_pm.iloc[3:, 45]),
            "por_pases_pp":  _to_num(raw_pm.iloc[3:, 48]),
        }).dropna(subset=["Min"])
        por_feats_ml = (
            df_pm.sort_values("Min", ascending=False)
            .groupby("Equipo").first()
            .drop(columns=["Min"])
        )
    except Exception:
        pass

    # ── Jugador features para modelo ML ──────────────────────────
    jug_feats_ml = pd.DataFrame()
    try:
        raw_jm = pd.read_excel(JUG_UNIF_PATH, sheet_name="Jugadores", header=None)
        df_jm = pd.DataFrame({
            "Jugador":  raw_jm.iloc[3:, 0].astype(str).values,
            "Equipo":   raw_jm.iloc[3:, 3].apply(normalize_team).values,
            "Posicion": raw_jm.iloc[3:, 2].astype(str).values,
            "Min":      _to_num(raw_jm.iloc[3:, 7]),
            "rating":   _to_num(raw_jm.iloc[3:, 22]),
            "g90":      _to_num(raw_jm.iloc[3:, 32]),
            "ast90":    _to_num(raw_jm.iloc[3:, 33]),
            "xg90":     _to_num(raw_jm.iloc[3:, 61]),
            "ent90":    _to_num(raw_jm.iloc[3:, 247]),
            "interc90": _to_num(raw_jm.iloc[3:, 250]),
            "despe90":  _to_num(raw_jm.iloc[3:, 253]),
            "bloq90":   _to_num(raw_jm.iloc[3:, 256]),
        }).dropna(subset=["Min"])
        df_jm = df_jm[df_jm["Min"] >= 400].copy()

        def _pg(p):
            t = str(p).split(",")[0].strip().upper()
            if t in {"GK","P","PT","POR","GOR"}: return "Portero"
            if t in {"FC","FW","SS","ST","CF","EI","ED","EC","DEL","AT","FWL","FWR"} \
               or t.startswith("F") or t.startswith("E"): return "Delantero"
            if t in {"DC","DFC","DL","DR","LD","LI","CB","LB","RB","DF","DFD","DFI"} \
               or t.startswith("L"): return "Defensa"
            return "Centrocampista"

        df_jm["PG"] = df_jm["Posicion"].apply(_pg)
        df_jm["OfScore90"]  = df_jm["xg90"].fillna(0) + df_jm["ast90"].fillna(0)
        df_jm["DefScore90"] = df_jm[["ent90","interc90","despe90","bloq90"]].fillna(0).sum(axis=1)
        rat_ml = df_jm["rating"].fillna(df_jm.groupby("PG")["rating"].transform("median"))
        df_jm["norm_of"]  = minmax_group(df_jm["OfScore90"],  df_jm["PG"])
        df_jm["norm_def"] = minmax_group(df_jm["DefScore90"], df_jm["PG"])
        df_jm["norm_rat"] = minmax_group(rat_ml,               df_jm["PG"])
        df_jm["IS_ml"]    = 0.35*df_jm["norm_of"] + 0.35*df_jm["norm_def"] + 0.30*df_jm["norm_rat"]

        df_jm["total_goles"] = df_jm["g90"].fillna(0) * df_jm["Min"] / 90
        df_jm["total_ast"]   = df_jm["ast90"].fillna(0) * df_jm["Min"] / 90

        filas_jm = []
        for equipo, grp in df_jm.groupby("Equipo"):
            by_is    = grp.sort_values("IS_ml",      ascending=False)
            by_min   = grp.sort_values("Min",        ascending=False)
            by_goles = grp.sort_values("total_goles", ascending=False)
            by_ast   = grp.sort_values("total_ast",   ascending=False)
            tot        = grp["Min"].sum()
            team_goles = grp["total_goles"].sum()
            team_ast   = grp["total_ast"].sum()
            filas_jm.append({
                "Equipo":               equipo,
                "jug_pilar_is":         by_is["IS_ml"].iloc[0],
                "jug_avg_is_top5":      by_is["IS_ml"].head(5).mean(),
                "jug_avg_rating_top11": by_min["rating"].head(11).mean(),
                "jug_pilar_pct_min":    by_is["Min"].iloc[0] / tot if tot > 0 else np.nan,
                "jug_top_scorer_g90":   grp["g90"].fillna(0).max(),
                "jug_avg_xg90_top11":   by_min["xg90"].fillna(0).head(11).mean(),
                "jug_top_scorer_name":  by_goles["Jugador"].iloc[0] if len(by_goles) > 0 else None,
                "jug_top_scorer_goles": int(round(by_goles["total_goles"].iloc[0])) if len(by_goles) > 0 else None,
                "jug_top_scorer_pct":   by_goles["total_goles"].iloc[0] / team_goles * 100 if team_goles > 0 else None,
                "jug_top_assist_name":  by_ast["Jugador"].iloc[0] if len(by_ast) > 0 else None,
                "jug_top_assist_ast":   int(round(by_ast["total_ast"].iloc[0])) if len(by_ast) > 0 else None,
                "jug_top_assist_pct":   by_ast["total_ast"].iloc[0] / team_ast * 100 if team_ast > 0 else None,
            })
        jug_feats_ml = pd.DataFrame(filas_jm).set_index("Equipo")
    except Exception:
        pass

    # ── Partidos (forma reciente + H2H) ──────────────────────────
    partidos_df = pd.DataFrame()
    forma_actual = {}
    try:
        raw_par = pd.read_excel(EXCEL, sheet_name="Partidos", header=None)
        data_par = raw_par.iloc[2:].reset_index(drop=True)
        score_par = data_par.iloc[:, 5].astype(str)
        mask_par  = score_par.str.contains("–", na=False)
        home_par  = data_par.iloc[:, 4][mask_par].apply(normalize_team)
        away_par  = data_par.iloc[:, 6][mask_par].apply(normalize_team)
        goles_par = score_par[mask_par].str.split("–", expand=True)
        lluvia_par = (data_par.iloc[:, 8][mask_par]
                      .astype(str).str.lower()
                      .str.contains("llovió|lluvia|sí", na=False)
                      .astype(int))
        partidos_df = pd.DataFrame({
            "Home":       home_par.values,
            "Away":       away_par.values,
            "lluvia":     lluvia_par.values,
            "goles_home": pd.to_numeric(goles_par[0].str.strip(), errors="coerce").values,
            "goles_away": pd.to_numeric(goles_par[1].str.strip(), errors="coerce").values,
        }).dropna(subset=["goles_home"]).reset_index(drop=True)

        # Calcular forma actual (últimos 5 partidos)
        historial = {}
        for _, row_p in partidos_df.iterrows():
            h, a = row_p["Home"], row_p["Away"]
            gh, ga = row_p["goles_home"], row_p["goles_away"]
            pts_h = 3 if gh > ga else (1 if gh == ga else 0)
            pts_a = 3 if ga > gh else (1 if ga == gh else 0)
            res_h = "V" if gh > ga else ("E" if gh == ga else "D")
            res_a = "V" if ga > gh else ("E" if ga == gh else "D")
            historial.setdefault(h, []).append({"gf": gh, "gc": ga, "pts": pts_h, "res": res_h, "rival": a})
            historial.setdefault(a, []).append({"gf": ga, "gc": gh, "pts": pts_a, "res": res_a, "rival": h})

        for equipo, hist in historial.items():
            recent = hist[-5:]
            forma_actual[equipo] = {
                "gf":    np.mean([x["gf"]  for x in recent]),
                "gc":    np.mean([x["gc"]  for x in recent]),
                "pts":   np.mean([x["pts"] for x in recent]),
                "racha": [x["res"] for x in recent],
            }
    except Exception:
        pass

    return (df, teams, arb, arbitros, jug, les, ausentes, dup,
            ws_df, por_feats_ml, jug_feats_ml, partidos_df, forma_actual)


@st.cache_resource(show_spinner="Cargando modelo predictivo…")
def load_model():
    """
    Carga el modelo Random Forest (Modelo B) y la lista de variables que espera.

    El modelo está serializado en disco como archivo .pkl para no tener que
    reentrenarlo cada vez. Si el archivo no se encuentra (p. ej. en entornos de demo),
    devuelve None y la app sigue funcionando sin predicción ML.
    """
    try:
        with open(MODEL_PATH, "rb") as f:
            modelo = pickle.load(f)
        with open(FEAT_PATH, "rb") as f:
            feature_names = pickle.load(f)
        return modelo, feature_names
    except Exception:
        return None, None


# ══════════════════════════════════════════════════════════════════
# PREDICCIÓN ML
# ══════════════════════════════════════════════════════════════════

def construir_vector(home, away, lluvia_bin,
                     ws_df, df_main, por_feats_ml, jug_feats_ml,
                     forma_actual, feature_names):
    """
    Construye el vector de 111 features para el partido home vs away.
    Perspectiva siempre del equipo local: Δ = home − away.
    """
    row = {}
    df_idx = df_main.set_index("Equipo") if "Equipo" in df_main.columns else df_main

    # 85 WhoScored features
    for col in WS_COLS:
        feat = f"d_{col}"
        if feat not in feature_names:
            continue
        h = float(ws_df.loc[home, col]) if (home in ws_df.index and col in ws_df.columns) else np.nan
        a = float(ws_df.loc[away, col]) if (away in ws_df.index and col in ws_df.columns) else np.nan
        row[feat] = h - a if (pd.notna(h) and pd.notna(a)) else 0.0

    # 5 clasificación features
    CLAS_MAP = {"Pts_pp":"Pts_pp","GF_pp":"GF_pp","GC_pp":"GC_pp",
                "GD_pp":"GD_pp","Pos_pct":"Pos%"}
    for feat_sfx, src in CLAS_MAP.items():
        feat = f"d_{feat_sfx}"
        if feat not in feature_names:
            continue
        h = float(df_idx.loc[home, src]) if home in df_idx.index else 0.0
        a = float(df_idx.loc[away, src]) if away in df_idx.index else 0.0
        row[feat] = h - a

    # 10 portero features
    POR_COLS = ["por_GA90","por_pct_par","por_pct_p0","por_rating",
                "por_par_AreaPeq","por_par_AreaPen","por_par_FueraArea",
                "por_pct_pk","por_pct_pases","por_pases_pp"]
    for col in POR_COLS:
        feat = f"d_{col}"
        if feat not in feature_names:
            continue
        h = float(por_feats_ml.loc[home, col]) if (not por_feats_ml.empty and home in por_feats_ml.index and col in por_feats_ml.columns) else 0.0
        a = float(por_feats_ml.loc[away, col]) if (not por_feats_ml.empty and away in por_feats_ml.index and col in por_feats_ml.columns) else 0.0
        row[feat] = h - a

    # 6 jugador features
    JUG_COLS = ["jug_pilar_is","jug_avg_is_top5","jug_avg_rating_top11",
                "jug_pilar_pct_min","jug_top_scorer_g90","jug_avg_xg90_top11"]
    for col in JUG_COLS:
        feat = f"d_{col}"
        if feat not in feature_names:
            continue
        h = float(jug_feats_ml.loc[home, col]) if (not jug_feats_ml.empty and home in jug_feats_ml.index and col in jug_feats_ml.columns) else 0.0
        a = float(jug_feats_ml.loc[away, col]) if (not jug_feats_ml.empty and away in jug_feats_ml.index and col in jug_feats_ml.columns) else 0.0
        row[feat] = h - a

    # Contextuales
    row["es_local"] = 1
    row["lluvia"]   = 1 if lluvia_bin else 0

    # Forma reciente
    fh = forma_actual.get(home, {})
    fa = forma_actual.get(away, {})
    row["d_forma_gf"]  = fh.get("gf", 0.0)  - fa.get("gf",  0.0)
    row["d_forma_gc"]  = fh.get("gc", 0.0)  - fa.get("gc",  0.0)
    row["d_forma_pts"] = fh.get("pts", 0.0) - fa.get("pts", 0.0)

    return np.array([row.get(f, 0.0) for f in feature_names], dtype=float)


def predecir_partido(mi_equipo, rival, es_local, lluvia,
                     ws_df, df_main, por_feats_ml, jug_feats_ml,
                     forma_actual, modelo, feature_names):
    """
    Predice el resultado desde la perspectiva de mi_equipo.
    Devuelve (yhat, clasificacion, confianza) o (None, None, None) si el modelo no está disponible.
    yhat > 0 → favorable para mi_equipo.
    """
    if modelo is None or feature_names is None:
        return None, None, None

    if es_local:
        vec = construir_vector(mi_equipo, rival, lluvia,
                               ws_df, df_main, por_feats_ml, jug_feats_ml,
                               forma_actual, feature_names)
        yhat = float(modelo.predict([vec])[0])
    else:
        # El rival es el local real → construimos rival vs mi_equipo, luego negamos
        vec = construir_vector(rival, mi_equipo, lluvia,
                               ws_df, df_main, por_feats_ml, jug_feats_ml,
                               forma_actual, feature_names)
        yhat = -float(modelo.predict([vec])[0])

    if yhat > 0.5:
        clasificacion = "Victoria"
    elif yhat < -0.5:
        clasificacion = "Derrota"
    else:
        clasificacion = "Empate"

    abs_y = abs(yhat)
    if abs_y > 1.5:
        confianza = "Alta"
    elif abs_y > 0.5:
        confianza = "Media"
    else:
        confianza = "Baja (partido muy incierto)"

    return yhat, clasificacion, confianza


def get_h2h(partidos_df, eq1, eq2, n=5):
    """Últimos n partidos entre eq1 y eq2."""
    if partidos_df.empty:
        return pd.DataFrame()
    mask = (
        ((partidos_df["Home"] == eq1) & (partidos_df["Away"] == eq2)) |
        ((partidos_df["Home"] == eq2) & (partidos_df["Away"] == eq1))
    )
    return partidos_df[mask].tail(n).copy().reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════
# LÓGICA PRESCRIPTIVA
# ══════════════════════════════════════════════════════════════════

def percentil(val, serie):
    """Calcula el percentil de un valor dentro de una serie (0–100). Ej: 85 → mejor que el 85% de equipos."""
    return round(float(np.mean(serie <= val) * 100))

def diferencia_etiqueta(delta, umbral_alto=0.15, umbral_bajo=-0.15):
    """
    Convierte la diferencia de IS entre equipos en una etiqueta visual y su color asociado.

    Retorna ("FAVORABLE"/"EQUILIBRADO"/"DESFAVORABLE", tipo_alerta) para usar en
    el cuadro central del encabezado del informe.
    """
    if delta > umbral_alto:
        return "FAVORABLE", "success"
    if delta < umbral_bajo:
        return "DESFAVORABLE", "error"
    return "EQUILIBRADO", "warning"

def jugadores_para_slot(jug_df, grupo, ausentes):
    """
    Devuelve la lista de jugadores disponibles para un puesto concreto de la alineación.

    Filtra por grupo posicional (portero, defensa, centro, extremo, delantero), excluye
    a los jugadores ausentes por lesión o sanción, y los ordena de más a menos minutos
    jugados para que el mejor candidato aparezca primero en el selector.
    """
    GRUPO_MAP = {
        "portero":   ["Portero"],
        "defensa":   ["Defensa"],
        "centro":    ["Centrocampista"],
        "extremo":   ["Delantero"],
        "delantero": ["Delantero"],
    }
    grupos_buscados = set(GRUPO_MAP.get(grupo, []))

    def cumple(pos_str):
        if pd.isna(pos_str):
            return False
        partes = [p.strip().capitalize() for p in str(pos_str).split(",")]
        return bool(grupos_buscados & set(partes))

    candidatos = jug_df[jug_df["Posicion"].apply(cumple)]
    candidatos = candidatos[~candidatos["Jugador"].isin(ausentes)]
    nombres = candidatos.sort_values("Minutos", ascending=False)["Jugador"].tolist()
    return nombres if nombres else ["(sin datos)"]

def ranking_is(df, equipo):
    """Devuelve la posición del equipo en el ranking de IS (1 = mejor de la liga)."""
    orden = df.sort_values("IS", ascending=False)["Equipo"].tolist()
    return orden.index(equipo) + 1 if equipo in orden else "?"


def generar_informe(local, rival, arbitro_sel, lluvia_partido,
                    lineup_local, df, arb, jug, les, ausentes, dup,
                    ws_df, por_feats_ml, jug_feats_ml, partidos_df,
                    forma_actual, modelo, feature_names, es_local=True):
    """
    Reúne todos los datos necesarios para generar el informe táctico del partido.

    Cruza métricas de equipo, jugadores, árbitro, porteros, forma reciente, historial
    de enfrentamientos directos y predicción del modelo ML, y los devuelve en un
    único diccionario que los 5 bloques del informe consumen directamente.

    El diccionario resultante contiene, entre otros:
      - row_local / row_rival : métricas de cada equipo
      - is_delta / estado_label / estado_color : valoración comparativa del IS
      - rival_lineup : alineación estimada del rival (por minutos jugados)
      - arb_* / liga_* : estadísticas del árbitro y medias de la liga
      - yhat / clasificacion / confianza : predicción del modelo Random Forest
      - forma_local / forma_rival : forma reciente (últimos 5 partidos)
      - h2h_df : historial de enfrentamientos directos
      - gol_* / pase_* / lado_* / por_* : métricas tácticas WhoScored
    """
    row_local = df[df["Equipo"] == local].iloc[0]
    row_rival = df[df["Equipo"] == rival].iloc[0]

    ausentes_local = les[(les["Equipo"] == local) & (les["ProbJugar"] == 0)]["Jugador"].tolist()
    ausentes_rival = les[(les["Equipo"] == rival) & (les["ProbJugar"] == 0)]["Jugador"].tolist()

    jug_rival      = jug[jug["Equipo"] == rival].copy()
    jug_rival_disp = jug_rival[~jug_rival["Jugador"].isin(ausentes_rival)]
    rival_lineup   = (jug_rival_disp
                      .sort_values("Minutos", ascending=False)
                      .head(11)[["Jugador","Posicion","Minutos","IS_indiv"]]
                      .reset_index(drop=True))

    jug_local      = jug[jug["Equipo"] == local].copy()
    jug_local_disp = jug_local[~jug_local["Jugador"].isin(ausentes_local)]

    is_delta = float(row_local["IS"]) - float(row_rival["IS"])
    estado_label, estado_color = diferencia_etiqueta(is_delta)

    arb_row       = arb[arb["Arbitro"] == arbitro_sel].head(1)
    arb_faltas    = float(arb_row["FaltasPP"].values[0])    if len(arb_row) else None
    arb_amarillas = float(arb_row["AmarillasPP"].values[0]) if len(arb_row) else None
    arb_rojas     = float(arb_row["RojasPP"].values[0])     if len(arb_row) and pd.notna(arb_row["RojasPP"].values[0]) else None
    arb_penaltis  = float(arb_row["PenaltisPP"].values[0])  if len(arb_row) and pd.notna(arb_row["PenaltisPP"].values[0]) else None
    arb_pct_local = float(arb_row["PctLocal"].values[0])    if len(arb_row) and pd.notna(arb_row["PctLocal"].values[0]) else None
    arb_pct_emp   = float(arb_row["PctEmpate"].values[0])   if len(arb_row) and pd.notna(arb_row["PctEmpate"].values[0]) else None
    arb_pct_vis   = float(arb_row["PctVisitante"].values[0]) if len(arb_row) and pd.notna(arb_row["PctVisitante"].values[0]) else None
    liga_faltas    = float(arb["FaltasPP"].mean())
    liga_amarillas = float(arb["AmarillasPP"].mean())
    liga_rojas     = float(arb["RojasPP"].mean())
    liga_penaltis  = float(arb["PenaltisPP"].mean())
    liga_pct_local = float(arb["PctLocal"].mean())
    liga_pct_vis   = float(arb["PctVisitante"].mean())

    dup_rival = (dup[dup["Equipo"] == rival]
                 .sort_values("Frecuencia", ascending=False)
                 .head(3))

    # ── Predicción ML ─────────────────────────────────────────────
    yhat, clasificacion, confianza = predecir_partido(
        local, rival, es_local, lluvia_partido,
        ws_df, df, por_feats_ml, jug_feats_ml,
        forma_actual, modelo, feature_names
    )

    # ── Forma reciente ────────────────────────────────────────────
    forma_local = forma_actual.get(local, {})
    forma_rival = forma_actual.get(rival, {})

    # ── H2H ───────────────────────────────────────────────────────
    h2h_df = get_h2h(partidos_df, local, rival)

    # ── Métricas WhoScored para recomendaciones ───────────────────
    def ws_val(equipo, col):
        if ws_df.empty or equipo not in ws_df.index or col not in ws_df.columns:
            return None
        v = ws_df.loc[equipo, col]
        return float(v) if pd.notna(v) else None

    # Goles por tipo de situación
    gol_abierto_local  = ws_val(local, "gol_juegoAbierto_fav_gen")
    gol_abierto_rival  = ws_val(rival, "gol_juegoAbierto_fav_gen")
    gol_bparado_local  = ws_val(local, "gol_balonParado_fav_gen")
    gol_bparado_rival  = ws_val(rival, "gol_balonParado_fav_gen")
    gol_ct_local       = ws_val(local, "gol_contraataque_fav_gen")
    gol_ct_rival       = ws_val(rival, "gol_contraataque_fav_gen")

    # Pases al hueco
    alhueco_local = ws_val(local, "pase_alHueco_fav_gen")
    alhueco_rival = ws_val(rival, "pase_alHueco_fav_gen")

    # Goles por zona de tiro
    goles_area_local   = ws_val(local, "goles_area_gen")
    goles_area_rival   = ws_val(rival, "goles_area_gen")
    goles_fuera_local  = ws_val(local, "goles_fueraArea_gen")
    goles_fuera_rival  = ws_val(rival, "goles_fueraArea_gen")
    tiros_area_rival   = ws_val(rival, "tiros_area_gen")
    tiros_peq_rival    = ws_val(rival, "tiros_areaPequeña_gen")

    # Lados de ataque
    lado_izq_rival = ws_val(rival, "ladosAtaque_izquierda_gen")
    lado_cen_rival = ws_val(rival, "ladosAtaque_centro_gen")
    lado_der_rival = ws_val(rival, "ladosAtaque_derecha_gen")

    # Tarjetas rojas
    rojas_local = float(row_local["Rojas"]) if pd.notna(row_local["Rojas"]) else None
    rojas_rival = float(row_rival["Rojas"]) if pd.notna(row_rival["Rojas"]) else None

    # ── Portero rival desde Jugadores Unificados ─────────────────
    por_ga90_rival    = None
    por_pct_p0_rival  = None
    por_pct_par_rival = None
    por_name_rival    = None
    por_ga90_local    = None
    por_pct_p0_local  = None
    por_pct_par_local = None
    por_name_local    = None
    if not por_feats_ml.empty:
        if rival in por_feats_ml.index:
            por_ga90_rival    = float(por_feats_ml.loc[rival, "por_GA90"])    if pd.notna(por_feats_ml.loc[rival, "por_GA90"])    else None
            por_pct_p0_rival  = float(por_feats_ml.loc[rival, "por_pct_p0"])  if pd.notna(por_feats_ml.loc[rival, "por_pct_p0"])  else None
            por_pct_par_rival = float(por_feats_ml.loc[rival, "por_pct_par"]) if pd.notna(por_feats_ml.loc[rival, "por_pct_par"]) else None
            por_name_rival    = str(por_feats_ml.loc[rival, "Jugador"])        if "Jugador" in por_feats_ml.columns else None
        if local in por_feats_ml.index:
            por_ga90_local    = float(por_feats_ml.loc[local, "por_GA90"])    if pd.notna(por_feats_ml.loc[local, "por_GA90"])    else None
            por_pct_p0_local  = float(por_feats_ml.loc[local, "por_pct_p0"])  if pd.notna(por_feats_ml.loc[local, "por_pct_p0"])  else None
            por_pct_par_local = float(por_feats_ml.loc[local, "por_pct_par"]) if pd.notna(por_feats_ml.loc[local, "por_pct_par"]) else None
            por_name_local    = str(por_feats_ml.loc[local, "Jugador"])        if "Jugador" in por_feats_ml.columns else None

    # ── Calidad de plantilla (jug_feats_ml) ──────────────────────
    def jf(equipo, col):
        if jug_feats_ml.empty or equipo not in jug_feats_ml.index or col not in jug_feats_ml.columns:
            return None
        v = jug_feats_ml.loc[equipo, col]
        return float(v) if pd.notna(v) else None

    def jfs(equipo, col):
        if jug_feats_ml.empty or equipo not in jug_feats_ml.index or col not in jug_feats_ml.columns:
            return None
        v = jug_feats_ml.loc[equipo, col]
        return str(v) if pd.notna(v) and str(v) not in ("nan", "None", "") else None

    def jfi(equipo, col):
        if jug_feats_ml.empty or equipo not in jug_feats_ml.index or col not in jug_feats_ml.columns:
            return None
        v = jug_feats_ml.loc[equipo, col]
        return int(v) if pd.notna(v) else None

    return {
        "row_local": row_local, "row_rival": row_rival,
        "is_delta": is_delta, "estado_label": estado_label, "estado_color": estado_color,
        "rival_lineup": rival_lineup, "jug_local_disp": jug_local_disp,
        "arb_faltas": arb_faltas, "arb_amarillas": arb_amarillas,
        "arb_rojas": arb_rojas, "arb_penaltis": arb_penaltis,
        "arb_pct_local": arb_pct_local, "arb_pct_emp": arb_pct_emp, "arb_pct_vis": arb_pct_vis,
        "liga_faltas": liga_faltas, "liga_amarillas": liga_amarillas,
        "liga_rojas": liga_rojas, "liga_penaltis": liga_penaltis,
        "liga_pct_local": liga_pct_local, "liga_pct_vis": liga_pct_vis,
        "ausentes_local": ausentes_local, "ausentes_rival": ausentes_rival,
        "dup_rival": dup_rival, "lluvia": lluvia_partido, "es_local": es_local,
        # ML
        "yhat": yhat, "clasificacion": clasificacion, "confianza": confianza,
        # Forma y H2H
        "forma_local": forma_local, "forma_rival": forma_rival, "h2h_df": h2h_df,
        # WhoScored tácticas
        "gol_abierto_local": gol_abierto_local, "gol_abierto_rival": gol_abierto_rival,
        "gol_bparado_local": gol_bparado_local, "gol_bparado_rival": gol_bparado_rival,
        "gol_ct_local": gol_ct_local, "gol_ct_rival": gol_ct_rival,
        "alhueco_local": alhueco_local, "alhueco_rival": alhueco_rival,
        "goles_area_local": goles_area_local, "goles_area_rival": goles_area_rival,
        "goles_fuera_local": goles_fuera_local, "goles_fuera_rival": goles_fuera_rival,
        "tiros_area_rival": tiros_area_rival, "tiros_peq_rival": tiros_peq_rival,
        "lado_izq_rival": lado_izq_rival, "lado_cen_rival": lado_cen_rival,
        "lado_der_rival": lado_der_rival,
        "rojas_local": rojas_local, "rojas_rival": rojas_rival,
        # Portero
        "por_ga90_rival": por_ga90_rival, "por_pct_p0_rival": por_pct_p0_rival,
        "por_pct_par_rival": por_pct_par_rival, "por_name_rival": por_name_rival,
        "por_ga90_local": por_ga90_local, "por_pct_p0_local": por_pct_p0_local,
        "por_pct_par_local": por_pct_par_local, "por_name_local": por_name_local,
        # Plantilla
        "jf_pilar_is_local":      jf(local, "jug_pilar_is"),
        "jf_avg_is_top5_local":   jf(local, "jug_avg_is_top5"),
        "jf_avg_rat11_local":     jf(local, "jug_avg_rating_top11"),
        "jf_xg11_local":          jf(local, "jug_avg_xg90_top11"),
        "jf_top_scorer_local":    jf(local, "jug_top_scorer_g90"),
        "jf_scorer_name_local":   jfs(local, "jug_top_scorer_name"),
        "jf_scorer_goles_local":  jfi(local, "jug_top_scorer_goles"),
        "jf_scorer_pct_local":    jf(local,  "jug_top_scorer_pct"),
        "jf_assist_name_local":   jfs(local, "jug_top_assist_name"),
        "jf_assist_ast_local":    jfi(local, "jug_top_assist_ast"),
        "jf_assist_pct_local":    jf(local,  "jug_top_assist_pct"),
        "jf_pilar_is_rival":      jf(rival, "jug_pilar_is"),
        "jf_avg_is_top5_rival":   jf(rival, "jug_avg_is_top5"),
        "jf_avg_rat11_rival":     jf(rival, "jug_avg_rating_top11"),
        "jf_xg11_rival":          jf(rival, "jug_avg_xg90_top11"),
        "jf_top_scorer_rival":    jf(rival, "jug_top_scorer_g90"),
        "jf_scorer_name_rival":   jfs(rival, "jug_top_scorer_name"),
        "jf_scorer_goles_rival":  jfi(rival, "jug_top_scorer_goles"),
        "jf_scorer_pct_rival":    jf(rival,  "jug_top_scorer_pct"),
        "jf_assist_name_rival":   jfs(rival, "jug_top_assist_name"),
        "jf_assist_ast_rival":    jfi(rival, "jug_top_assist_ast"),
        "jf_assist_pct_rival":    jf(rival,  "jug_top_assist_pct"),
    }


# ══════════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Algoritmo Prescriptivo LaLiga",
    page_icon="⚽",
    layout="wide",
)

_PITCH_SVG = """
<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 620' preserveAspectRatio='xMidYMid slice'>
  <rect width='400' height='620' fill='#0e2410'/>
  <!-- Boundary: full rect — single clean top line, penalty/goal area tops omitted to avoid stacking -->
  <rect x='22' y='22' width='356' height='576' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- Halfway line -->
  <line x1='22' y1='310' x2='378' y2='310' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- Center circle r=50 (9.15m) + spot -->
  <circle cx='200' cy='310' r='50' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <circle cx='200' cy='310' r='3.5' fill='rgba(255,255,255,0.70)'/>
  <!-- TOP penalty area: 3-sided, no top border (outer rect top already provides it) -->
  <path d='M 95,22 L 95,113 L 305,113 L 305,22' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- TOP goal area: 3-sided, no top border -->
  <path d='M 152,22 L 152,52 L 248,52 L 248,22' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- TOP goal: 7.32m → 38px wide -->
  <rect x='181' y='9' width='38' height='14' fill='none' stroke='rgba(255,255,255,0.38)' stroke-width='2'/>
  <!-- TOP penalty spot: 11m → 60px → y=82 -->
  <circle cx='200' cy='82' r='3' fill='rgba(255,255,255,0.70)'/>
  <!-- TOP penalty arc: center(200,82) r=50, exits PA at y=113, x=±39 -->
  <path d='M 161,113 A 50,50 0 0,0 239,113' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- BOTTOM penalty area: y=507..598 = 91px -->
  <rect x='95' y='507' width='210' height='91' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- BOTTOM goal area: y=568..598 = 30px -->
  <rect x='152' y='568' width='96' height='30' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- BOTTOM goal -->
  <rect x='181' y='598' width='38' height='13' fill='none' stroke='rgba(255,255,255,0.38)' stroke-width='2'/>
  <!-- BOTTOM penalty spot: y=538 -->
  <circle cx='200' cy='538' r='3' fill='rgba(255,255,255,0.70)'/>
  <!-- BOTTOM penalty arc: center(200,538) r=50, exits PA at y=507, x=±39 -->
  <path d='M 161,507 A 50,50 0 0,1 239,507' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <!-- Corner arcs r=12 -->
  <path d='M 22,34 A 12,12 0 0,1 34,22' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <path d='M 366,22 A 12,12 0 0,1 378,34' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <path d='M 378,586 A 12,12 0 0,1 366,598' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
  <path d='M 34,598 A 12,12 0 0,1 22,586' fill='none' stroke='rgba(255,255,255,0.50)' stroke-width='2'/>
</svg>
""".strip()

_PITCH_B64 = base64.b64encode(_PITCH_SVG.encode()).decode()

st.markdown(f"""
<style>
.metric-card {{
    background: #f8f9fa;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 12px;
    border-left: 4px solid #1a73e8;
}}
.section-header {{
    font-size: 1.55rem;
    font-weight: 800;
    color: #F5C842;
    border-bottom: 2px solid #F5C842;
    padding-bottom: 6px;
    margin-top: 24px;
    margin-bottom: 12px;
}}
.tag-success {{ color: #27ae60; font-weight: bold; }}
.tag-warning {{ color: #e67e22; font-weight: bold; }}
.tag-error   {{ color: #e74c3c; font-weight: bold; }}
.forma-V {{ color: #27ae60; font-weight: bold; }}
.forma-E {{ color: #e67e22; font-weight: bold; }}
.forma-D {{ color: #e74c3c; font-weight: bold; }}

/* Campo de fútbol directamente en el área principal (no en sidebar) */
[data-testid="stMain"] {{
    background-image: url('data:image/svg+xml;base64,{_PITCH_B64}');
    background-size: cover;
    background-position: center top;
    background-attachment: local;
}}
[data-testid="stMain"] > div {{
    background-color: transparent;
}}
/* Sidebar con fondo sólido oscuro */
[data-testid="stSidebar"] {{
    background-color: #141820;
}}
</style>
""", unsafe_allow_html=True)

st.title("Algoritmo Prescriptivo Táctico — LaLiga 2025-26")
st.caption("Herramienta de apoyo a la decisión para cuerpos técnicos")

# ── Carga de datos ────────────────────────────────────────────────
(df, teams, arb, arbitros, jug, les, ausentes, dup,
 ws_df, por_feats_ml, jug_feats_ml, partidos_df, forma_actual) = load_data()

modelo_b, feature_names = load_model()

# ── SIDEBAR ───────────────────────────────────────────────────────
with st.sidebar:
    st.header("🗓️ Parámetros del partido")

    mi_equipo = st.selectbox("Tu equipo", teams, index=None, placeholder="Elige tu equipo")
    rival_options = [t for t in teams if t != mi_equipo] if mi_equipo else teams
    rival = st.selectbox("Rival", rival_options, index=None, placeholder="Elige el rival")

    es_local = st.radio(
        "¿Dónde juegas?",
        ["🏠 En casa (local)", "✈️ Fuera (visitante)"],
        index=0,
        horizontal=True,
    ) == "🏠 En casa (local)"

    local = mi_equipo

    arbitro_sel  = st.selectbox("Árbitro designado", ["(Desconocido)"] + arbitros)
    lluvia_partido = st.toggle("¿Se prevé lluvia?", value=False)

    st.divider()
    st.header("📋 Tu alineación")

    formacion = st.selectbox("Formación", FORMACIONES, index=0)

    if mi_equipo:
        ausentes_local_confirmed = les[
            (les["Equipo"] == local) & (les["ProbJugar"] == 0)
        ]["Jugador"].tolist()
        jug_equipo = jug[jug["Equipo"] == local].copy()

        if ausentes_local_confirmed:
            st.caption(f"⚠️ Excluidos (0% prob.): {', '.join(ausentes_local_confirmed)}")

        slots = FORMACIONES_SLOTS.get(formacion, [])

        asignados = set()
        defaults  = []
        for etiqueta, grupo in slots:
            candidatos_ord = jugadores_para_slot(jug_equipo, grupo, ausentes_local_confirmed)
            idx_default = 0
            for j, nombre in enumerate(candidatos_ord):
                if nombre not in asignados:
                    idx_default = j
                    asignados.add(nombre)
                    break
            defaults.append((candidatos_ord, idx_default))

        lineup_local = []
        for i, ((etiqueta, grupo), (candidatos, idx_default)) in enumerate(zip(slots, defaults)):
            col_lbl, col_jug = st.columns([2, 3])
            with col_lbl:
                st.markdown(f"<small>{etiqueta}</small>", unsafe_allow_html=True)
            with col_jug:
                jug_sel = st.selectbox(etiqueta, candidatos, index=idx_default,
                                       key=f"jug_{i}", label_visibility="collapsed")
            lineup_local.append({"Posicion": etiqueta, "Jugador": jug_sel})

        generar = st.button("🔍 Generar informe táctico", type="primary", use_container_width=True)
    else:
        st.caption("Selecciona tu equipo para configurar la alineación.")
        lineup_local = []
        generar = False


# ══════════════════════════════════════════════════════════════════
# PANTALLA PRINCIPAL
# ══════════════════════════════════════════════════════════════════

if not generar:
    st.info("Configura el partido en el panel izquierdo y pulsa **Generar informe táctico**.")
    st.markdown('<div class="section-header">Tabla La Liga 25-26</div>', unsafe_allow_html=True)
    tabla_is = (df[["Equipo","IS","Pts","xG_pp","GC_pp","Rating"]]
                .sort_values("IS", ascending=False)
                .reset_index(drop=True))
    tabla_is.index = tabla_is.index + 1
    _filas_home = ""
    for pos, row in tabla_is.iterrows():
        bg = "#1a2540" if pos % 2 == 1 else "#1e2b4a"
        _filas_home += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:9px 14px;color:#8892b0;text-align:center;font-weight:600">{pos}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;font-weight:700">{row["Equipo"]}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;text-align:center;font-weight:700">{row["IS"]:.3f}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;text-align:center">{int(row["Pts"])}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;text-align:center">{row["xG_pp"]:.2f}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;text-align:center">{row["GC_pp"]:.2f}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;text-align:center">{row["Rating"]:.2f}</td>'
            f'</tr>'
        )
    st.markdown(f"""
<div style="border-radius:12px;overflow:hidden;border:1px solid #2d3d6b;background:#1a2540">
<table style="width:100%;border-collapse:collapse;font-size:0.9rem">
  <thead>
    <tr style="background:#1e3a6e">
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">#</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:left;font-size:0.88rem;letter-spacing:0.05em">EQUIPO</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">IS</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">PTS</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">xG / PARTIDO</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">GC / PARTIDO</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">RATING</th>
    </tr>
  </thead>
  <tbody>{_filas_home}</tbody>
</table>
</div>""", unsafe_allow_html=True)

    st.markdown(
        '<div class="section-header" style="margin-top:28px">Mapa ofensivo-defensivo · LaLiga 25-26</div>',
        unsafe_allow_html=True
    )
    _med_xg = float(df["xG_pp"].median())
    _med_gc = float(df["GC_pp"].median())
    _xg_max = float(df["xG_pp"].max()) + 0.18
    _xg_min = float(df["xG_pp"].min()) - 0.18
    _gc_max = float(df["GC_pp"].max()) + 0.18
    _gc_min = float(df["GC_pp"].min()) - 0.18
    _df_oth = df[~df["Equipo"].isin([local, rival])].dropna(subset=["xG_pp","GC_pp"])
    _df_loc = df[df["Equipo"] == local].dropna(subset=["xG_pp","GC_pp"]) if local else pd.DataFrame()
    _df_riv = df[df["Equipo"] == rival].dropna(subset=["xG_pp","GC_pp"]) if rival else pd.DataFrame()

    _fig_sc = go.Figure()

    # Sombreado de cuadrantes (capa inferior)
    # Y axis invertido: gc_min (valor bajo) → visualmente arriba = buena defensa
    for (_qx0, _qx1, _qy0, _qy1, _qcol) in [
        (_med_xg, _xg_max, _gc_min, _med_gc, "rgba(39,174,96,0.09)"),   # Élite
        (_xg_min, _med_xg, _gc_min, _med_gc, "rgba(74,158,202,0.09)"),  # Defensivos
        (_med_xg, _xg_max, _med_gc, _gc_max, "rgba(230,126,34,0.09)"),  # Ofensivos
        (_xg_min, _med_xg, _med_gc, _gc_max, "rgba(231,76,60,0.09)"),   # En dificultades
    ]:
        _fig_sc.add_shape(
            type="rect", x0=_qx0, x1=_qx1, y0=_qy0, y1=_qy1,
            fillcolor=_qcol, line=dict(width=0), layer="below",
        )

    # Líneas de medianas
    _fig_sc.add_vline(x=_med_xg, line=dict(color="#3d4870", dash="dot", width=1))
    _fig_sc.add_hline(y=_med_gc, line=dict(color="#3d4870", dash="dot", width=1))

    # Resto de equipos
    _fig_sc.add_trace(go.Scatter(
        x=_df_oth["xG_pp"], y=_df_oth["GC_pp"],
        mode="markers+text",
        marker=dict(color="#4a5568", size=10, line=dict(color="#2d3d6b", width=1)),
        text=_df_oth["Equipo"], textposition="top center",
        textfont=dict(color="#8892b0", size=10),
        name="Resto de equipos",
        hovertemplate="%{text}<br>xG/Partido: %{x:.2f}<br>GC/Partido: %{y:.2f}<extra></extra>",
    ))
    if not _df_loc.empty:
        _fig_sc.add_trace(go.Scatter(
            x=_df_loc["xG_pp"], y=_df_loc["GC_pp"],
            mode="markers+text",
            marker=dict(color="#4a9eca", size=16, line=dict(color="#ffffff", width=2)),
            text=[local], textposition="top center",
            textfont=dict(color="#4a9eca", size=12),
            name=local,
            hovertemplate=f"{local}<br>xG/Partido: %{{x:.2f}}<br>GC/Partido: %{{y:.2f}}<extra></extra>",
        ))
    if not _df_riv.empty:
        _fig_sc.add_trace(go.Scatter(
            x=_df_riv["xG_pp"], y=_df_riv["GC_pp"],
            mode="markers+text",
            marker=dict(color="#F5C842", size=16, line=dict(color="#ffffff", width=2)),
            text=[rival], textposition="top center",
            textfont=dict(color="#F5C842", size=12),
            name=rival,
            hovertemplate=f"{rival}<br>xG/Partido: %{{x:.2f}}<br>GC/Partido: %{{y:.2f}}<extra></extra>",
        ))

    # Etiquetas de cuadrante en esquinas (coords paper, fuera de los datos)
    # Con Y invertido: paper y=1 = arriba = bajo GC = buena defensa
    for (_lx, _ly, _ltxt, _lxa, _lya, _lcol) in [
        (0.99, 0.99, "ÉLITE",            "right", "top",    "rgba(39,174,96,0.85)"),
        (0.01, 0.99, "DEFENSIVOS",       "left",  "top",    "rgba(74,158,202,0.85)"),
        (0.99, 0.01, "OFENSIVOS",        "right", "bottom", "rgba(230,126,34,0.85)"),
        (0.01, 0.01, "EN DIFICULTADES",  "left",  "bottom", "rgba(231,76,60,0.85)"),
    ]:
        _fig_sc.add_annotation(
            x=_lx, y=_ly, text=_ltxt,
            xref="paper", yref="paper",
            showarrow=False, xanchor=_lxa, yanchor=_lya,
            font=dict(color=_lcol, size=10, family="Arial Black"),
            bgcolor="rgba(26,32,53,0.55)", borderpad=5,
        )

    _fig_sc.update_layout(
        xaxis=dict(
            title=dict(text="xG / Partido (capacidad ofensiva)", font=dict(color="#ffffff", size=11)),
            color="#ffffff", gridcolor="#1e2b4a", zerolinecolor="#2d3d6b",
            tickfont=dict(color="#ffffff"),
            range=[_xg_min, _xg_max],
        ),
        yaxis=dict(
            title=dict(text="GC / Partido  ↑ mejor defensa", font=dict(color="#ffffff", size=11)),
            color="#ffffff", gridcolor="#1e2b4a", zerolinecolor="#2d3d6b",
            tickfont=dict(color="#ffffff"),
            range=[_gc_max, _gc_min],
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#1a2035",
        legend=dict(font=dict(color="#ccd6f6", size=11), bgcolor="rgba(0,0,0,0)"),
        margin=dict(t=20, b=50, l=60, r=20),
        height=480,
    )
    st.plotly_chart(_fig_sc, use_container_width=True)

    # Análisis del mapa
    if local and rival and not _df_loc.empty and not _df_riv.empty:
        def _quadrant_sc(xg, gc):
            if xg >= _med_xg and gc <= _med_gc: return "élite"
            if xg <  _med_xg and gc <= _med_gc: return "defensivo"
            if xg >= _med_xg and gc >  _med_gc: return "ofensivo"
            return "en dificultades"
        _qdesc = {
            "élite":            "sólido en ataque y defensa",
            "defensivo":        "sólido defensivamente pero con potencial ofensivo limitado",
            "ofensivo":         "peligroso en ataque pero vulnerable atrás",
            "en dificultades":  "con carencias tanto ofensivas como defensivas",
        }
        _ql = _quadrant_sc(float(_df_loc["xG_pp"].iloc[0]), float(_df_loc["GC_pp"].iloc[0]))
        _qr = _quadrant_sc(float(_df_riv["xG_pp"].iloc[0]), float(_df_riv["GC_pp"].iloc[0]))
        _al1 = f"<strong>{local}</strong> es {_qdesc[_ql]} esta temporada."
        if _ql == _qr:
            _al2 = f"<strong>{rival}</strong> comparte cuadrante — el enfrentamiento será muy equilibrado tácticamente."
        else:
            _al2 = f"<strong>{rival}</strong> es {_qdesc[_qr]} — el partido plantea un contraste táctico claro entre ambos estilos."
        st.markdown(
            f'<div style="background:#142032;border:1px solid rgba(74,158,202,0.45);'
            f'border-radius:10px;padding:13px 18px;font-size:0.9rem;line-height:1.7;'
            f'color:#ccd6f6;margin-top:4px">{_al1}<br>{_al2}</div>',
            unsafe_allow_html=True
        )

else:
    inf = generar_informe(
        local, rival, arbitro_sel, lluvia_partido,
        lineup_local, df, arb, jug, les, ausentes, dup,
        ws_df, por_feats_ml, jug_feats_ml, partidos_df,
        forma_actual, modelo_b, feature_names,
        es_local=es_local,
    )

    rl = inf["row_local"]
    rr = inf["row_rival"]

    # ── Cabecera ──────────────────────────────────────────────────
    icono_mi    = "🏠" if es_local else "✈️"
    icono_rival = "✈️" if es_local else "🏠"

    is_media = float(df["IS"].mean())
    delta_local = float(rl["IS"]) - is_media
    delta_rival = float(rr["IS"]) - is_media

    def _is_bubble(delta):
        if delta >= 0:
            col = "#2ecc71"; bg = "#2ecc7120"; arrow = "↑"
        else:
            col = "#e74c3c"; bg = "#e74c3c20"; arrow = "↓"
        return (
            f'<div style="display:inline-block;background:{bg};border:1px solid {col};'
            f'border-radius:20px;padding:4px 14px;color:{col};font-size:0.85rem;'
            f'font-weight:700;margin-top:10px">'
            f'{arrow} {delta:+.3f} vs media Liga</div>'
        )

    def _card(nombre, is_val, rank, icono, delta):
        condicion = "Local" if icono == "🏠" else "Visitante"
        return f"""
<div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
padding:26px 20px;text-align:center;border:1px solid #3d4870;
box-shadow:0 4px 20px rgba(0,0,0,0.4)">
  <div style="font-size:0.82rem;color:#8892b0;letter-spacing:0.08em;
  text-transform:uppercase;margin-bottom:6px">{icono} {condicion}</div>
  <div style="font-size:1.55rem;font-weight:800;color:#ccd6f6;
  margin-bottom:12px;line-height:1.2">{nombre}</div>
  <div style="font-size:2.6rem;font-weight:900;color:#64b5f6;
  line-height:1;margin-bottom:4px">{is_val:.3f}</div>
  <div style="font-size:0.78rem;color:#ffffff;margin-bottom:2px">
  Índice de Éxito (IS) · #{rank} en LaLiga</div>
  {_is_bubble(delta)}
</div>"""

    estado_col = {"FAVORABLE": "#2ecc71", "EQUILIBRADO": "#f39c12",
                  "DESFAVORABLE": "#e74c3c"}[inf["estado_label"]]
    center_html = f"""
<div style="display:flex;flex-direction:column;align-items:center;
justify-content:center;min-height:215px;text-align:center">
  <div style="font-size:0.9rem;color:#ffffff;margin-bottom:10px;
  letter-spacing:0.1em">VS</div>
  <div style="font-size:1rem;font-weight:800;color:{estado_col};
  background:{estado_col}20;border:1.5px solid {estado_col};
  border-radius:10px;padding:8px 14px;letter-spacing:0.05em">
  {inf["estado_label"]}</div>
  <div style="font-size:0.75rem;color:#ffffff;margin-top:10px">
  ΔIS {inf['is_delta']:+.3f}</div>
</div>"""

    c1, c2, c3 = st.columns([5, 2, 5])
    with c1:
        st.markdown(_card(local, rl["IS"], ranking_is(df, local),
                          icono_mi, delta_local), unsafe_allow_html=True)
    with c2:
        st.markdown(center_html, unsafe_allow_html=True)
    with c3:
        st.markdown(_card(rival, rr["IS"], ranking_is(df, rival),
                          icono_rival, delta_rival), unsafe_allow_html=True)

    if inf["ausentes_local"]:
        st.warning(f"Bajas confirmadas ({local}): {', '.join(inf['ausentes_local'])}")
    if inf["ausentes_rival"]:
        st.info(f"Bajas confirmadas ({rival}): {', '.join(inf['ausentes_rival'])}")

    st.divider()

    # ════════════════════════════════════════
    # BLOQUE 1 — RESUMEN GENERAL
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">📊 Bloque 1 · Resumen General</div>',
                unsafe_allow_html=True)

    is_delta = inf["is_delta"]

    # ── Predicción ML ─────────────────────────────────────────────
    if inf["yhat"] is not None:
        yhat  = inf["yhat"]
        clas  = inf["clasificacion"]
        color_ml = {"Victoria": "success", "Empate": "warning", "Derrota": "error"}[clas]
        signo = "+" if yhat >= 0 else ""
        st.markdown("**Predicción del Modelo Predictivo (Random Forest)**")

        col_ml_l, col_ml1, col_ml2, col_ml_r = st.columns([1, 2, 2, 1])
        card_col = {"Victoria": "#2ecc71", "Empate": "#f39c12", "Derrota": "#e74c3c"}[clas]
        with col_ml1:
            st.markdown(f"""
<div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
padding:22px 16px;text-align:center;border:1px solid #3d4870;
box-shadow:0 4px 20px rgba(0,0,0,0.4)">
  <div style="font-size:0.82rem;color:#8892b0;letter-spacing:0.08em;
  text-transform:uppercase;margin-bottom:8px">Diferencia esperada</div>
  <div style="font-size:2.4rem;font-weight:900;color:#64b5f6;line-height:1">
  {signo}{yhat:.2f} goles</div>
</div>""", unsafe_allow_html=True)
        with col_ml2:
            st.markdown(f"""
<div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
padding:22px 16px;text-align:center;border:1px solid #3d4870;
box-shadow:0 4px 20px rgba(0,0,0,0.4)">
  <div style="font-size:0.82rem;color:#8892b0;letter-spacing:0.08em;
  text-transform:uppercase;margin-bottom:8px">Resultado probable</div>
  <div style="font-size:2.4rem;font-weight:900;color:{card_col};line-height:1">
  {clas}</div>
</div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)
        _clas_texto = (
            f"**{clas.lower()}**"
            if clas.lower() == "empate"
            else f"**{clas.lower()}** para {local}"
        )
        getattr(st, color_ml)(
            f"El modelo estima {_clas_texto} con una diferencia esperada "
            f"de {signo}{yhat:.2f} goles."
        )
    else:
        st.caption("_Modelo predictivo no disponible._")

    st.markdown("---")

    # ── Métricas IS ───────────────────────────────────────────────
    def _delta_color(val, positive_good=True):
        if val > 0:
            col = "#2ecc71" if positive_good else "#e74c3c"
            arrow = "↑"
        elif val < 0:
            col = "#e74c3c" if positive_good else "#2ecc71"
            arrow = "↓"
        else:
            col = "#8892b0"; arrow = "→"
        return col, arrow

    d_is  = is_delta
    d_pts = rl['Pts_pp'] - rr['Pts_pp']
    d_xg  = rl['xG_pp']  - rr['xG_pp']
    d_gc  = rl['GC_pp']  - rr['GC_pp']

    col_is,  arr_is  = _delta_color(d_is,  positive_good=True)
    col_pts, arr_pts = _delta_color(d_pts, positive_good=True)
    col_xg,  arr_xg  = _delta_color(d_xg,  positive_good=True)
    col_gc,  arr_gc  = _delta_color(d_gc,  positive_good=False)

    def _stat_cell(label, value, delta_val, d_col, d_arr):
        return f"""
<div style="flex:1;text-align:center;padding:0 12px;border-right:1px solid #3d4870">
  <div style="font-size:0.78rem;color:#8892b0;letter-spacing:0.07em;
  text-transform:uppercase;margin-bottom:6px">{label}</div>
  <div style="font-size:2rem;font-weight:900;color:#ccd6f6;line-height:1;
  margin-bottom:6px">{value}</div>
  <div style="font-size:0.8rem;color:{d_col};font-weight:700">
  {d_arr} {delta_val:+.3f} vs rival</div>
</div>"""

    st.markdown(f"<div style='font-size:0.85rem;color:#8892b0;font-weight:600;"
                f"letter-spacing:0.06em;text-transform:uppercase;margin-bottom:8px'>"
                f"Indicadores de {local}</div>", unsafe_allow_html=True)

    st.markdown(f"""
<div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
padding:22px 10px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
display:flex;flex-direction:row;align-items:center">
  {_stat_cell("IS", f"{rl['IS']:.3f}", d_is, col_is, arr_is)}
  {_stat_cell("Pts / partido", f"{rl['Pts_pp']:.2f}", d_pts, col_pts, arr_pts)}
  {_stat_cell("xG / partido", f"{rl['xG_pp']:.2f}", d_xg, col_xg, arr_xg)}
  <div style="flex:1;text-align:center;padding:0 12px">
    <div style="font-size:0.78rem;color:#8892b0;letter-spacing:0.07em;
    text-transform:uppercase;margin-bottom:6px">GC / partido</div>
    <div style="font-size:2rem;font-weight:900;color:#ccd6f6;line-height:1;
    margin-bottom:6px">{rl['GC_pp']:.2f}</div>
    <div style="font-size:0.8rem;color:{col_gc};font-weight:700">
    {arr_gc} {d_gc:+.3f} vs rival</div>
  </div>
</div>""", unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom:16px'></div>", unsafe_allow_html=True)

    # Valoración narrativa según ΔIS y localía
    condicion = "jugando en casa" if es_local else "jugando como visitante"
    if is_delta > 0.20:
        valoracion = (
            f"**{local}** afronta este partido con una ventaja objetiva clara sobre **{rival}** "
            f"(ΔIS = {is_delta:+.3f}), {condicion}. La estrategia debe orientarse a **consolidar el dominio** "
            f"y explotar las debilidades del rival sin asumir riesgos innecesarios."
        )
    elif is_delta > 0.05:
        valoracion = (
            f"El enfrentamiento es **ligeramente favorable** para {local} (ΔIS = {is_delta:+.3f}), {condicion}. "
            f"El margen es pequeño: la eficiencia en las transiciones y la solidez defensiva "
            f"serán determinantes."
        )
    elif is_delta > -0.05:
        if es_local:
            valoracion = (
                f"Partido **muy equilibrado** (ΔIS = {is_delta:+.3f}). Jugando en casa, "
                f"el factor local puede inclinar la balanza. "
                f"En un duelo tan ajustado, un error defensivo o una ocasión bien resuelta puede ser la diferencia."
            )
        else:
            valoracion = (
                f"Partido **muy equilibrado** (ΔIS = {is_delta:+.3f}). Como visitante, "
                f"un punto puede ser un resultado válido — gestiona bien los momentos de presión rival. "
                f"La eficiencia en las pocas ocasiones que generes será clave."
            )
    elif is_delta > -0.20:
        valoracion = (
            f"**{rival}** parte con ventaja objetiva (ΔIS = {is_delta:+.3f}) y {condicion}. "
            f"Se recomienda una propuesta defensivamente sólida, con apuesta por el contraataque "
            f"rápido y la explotación de los puntos débiles del rival identificados a continuación."
        )
    else:
        if es_local:
            valoracion = (
                f"**{local}** enfrenta en casa a un rival notablemente superior (ΔIS = {is_delta:+.3f}). "
                f"La clave estará en la **organización defensiva compacta**, minimizar los xG encajados "
                f"y aprovechar cualquier situación a balón parado."
            )
        else:
            valoracion = (
                f"**{local}** visita a un rival notablemente superior (ΔIS = {is_delta:+.3f}). "
                f"Como visitante ante un rival más fuerte, la prioridad es **no encajar en los primeros minutos** "
                f"y aprovechar el contraataque y el balón parado como herramientas principales."
            )
    valoracion_html = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', valoracion)
    st.markdown(
        f'<div style="background:#142032;border:1px solid rgba(74,158,202,0.45);'
        f'border-radius:10px;padding:14px 18px;color:#ccd6f6;font-size:0.95rem;line-height:1.6;'
        f'margin-bottom:16px">'
        f'{valoracion_html}</div>',
        unsafe_allow_html=True
    )

    # Tabla comparativa
    _metricas = [
        ("IS",                "Rendimiento global",          f"{rl['IS']:.3f}",              f"{rr['IS']:.3f}"),
        ("Pts / Partido",     "Puntos por partido",          f"{rl['Pts_pp']:.2f}",          f"{rr['Pts_pp']:.2f}"),
        ("xG / Partido",      "Ocasiones de gol generadas",  f"{rl['xG_pp']:.2f}",           f"{rr['xG_pp']:.2f}"),
        ("GC / Partido",      "Goles encajados por partido", f"{rl['GC_pp']:.2f}",           f"{rr['GC_pp']:.2f}"),
        ("Rating",            "Valoración media WhoScored",  f"{rl['Rating']:.2f}",          f"{rr['Rating']:.2f}"),
        ("Posesión %",        "Control del balón",           f"{rl['Pos%']:.1f}%",           f"{rr['Pos%']:.1f}%"),
        ("Precisión pase %",  "Eficacia en el pase",         f"{rl['Precision_pase']:.1f}%", f"{rr['Precision_pase']:.1f}%"),
        ("Tiros / Partido",   "Volumen ofensivo",            f"{rl['Tiros_pp']:.1f}",        f"{rr['Tiros_pp']:.1f}"),
        ("Tiros contra / Partido", "Presión recibida",            f"{rl['Tiros_contra_pp']:.1f}", f"{rr['Tiros_contra_pp']:.1f}"),
        ("Faltas cometidas",  "Agresividad táctica",         f"{rl['Faltas_com']:.1f}",      f"{rr['Faltas_com']:.1f}"),
        ("Tarjetas rojas",    "Disciplina — expulsiones",
         f"{round(rl['Rojas'] * rl['PJ'])}" if pd.notna(rl['Rojas']) else "–",
         f"{round(rr['Rojas'] * rr['PJ'])}" if pd.notna(rr['Rojas']) else "–"),
    ]

    _filas = ""
    for i, (metrica, desc, val_l, val_r) in enumerate(_metricas):
        bg = "#1a2540" if i % 2 == 0 else "#1e2b4a"
        _filas += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:9px 14px;color:#ffffff;font-weight:600;white-space:nowrap">{metrica}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;font-size:0.82rem">{desc}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;font-weight:700;text-align:center">{val_l}</td>'
            f'<td style="padding:9px 14px;color:#ffffff;font-weight:700;text-align:center">{val_r}</td>'
            f'</tr>'
        )

    st.markdown(f"""
<div style="border-radius:12px;overflow:hidden;border:1px solid #2d3d6b;background:#1a2540">
<table style="width:100%;border-collapse:collapse;font-size:0.9rem">
  <thead>
    <tr style="background:#1e3a6e">
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:left;font-size:0.88rem;letter-spacing:0.05em">MÉTRICA</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:left;font-size:0.88rem;letter-spacing:0.05em">DESCRIPCIÓN</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">{local.upper()}</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">{rival.upper()}</th>
    </tr>
  </thead>
  <tbody>{_filas}</tbody>
</table>
</div>""", unsafe_allow_html=True)

    # ── Radar comparativo ─────────────────────────────────────────
    st.markdown(
        '<div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;'
        'text-transform:uppercase;margin:22px 0 10px">Perfil táctico comparativo</div>',
        unsafe_allow_html=True
    )
    def _norm_radar(col, val):
        mn, mx = float(col.min()), float(col.max())
        return 0.0 if mx == mn else (float(val) - mn) / (mx - mn) * 100

    _r_labels = [
        "xG / Partido", "Solidez\ndefensiva", "IS",
        "Posesión", "Precisión\npase", "Tiros a puerta\n/ Partido",
    ]
    _r_local = [
        _norm_radar(df["xG_pp"],          rl["xG_pp"]),
        _norm_radar(-df["GC_pp"],         -rl["GC_pp"]),
        _norm_radar(df["IS"],              rl["IS"]),
        _norm_radar(df["Pos%"],            rl["Pos%"]),
        _norm_radar(df["Precision_pase"],  rl["Precision_pase"]),
        _norm_radar(df["Tiros_puerta_pp"], rl["Tiros_puerta_pp"]),
    ]
    _r_rival = [
        _norm_radar(df["xG_pp"],          rr["xG_pp"]),
        _norm_radar(-df["GC_pp"],         -rr["GC_pp"]),
        _norm_radar(df["IS"],              rr["IS"]),
        _norm_radar(df["Pos%"],            rr["Pos%"]),
        _norm_radar(df["Precision_pase"],  rr["Precision_pase"]),
        _norm_radar(df["Tiros_puerta_pp"], rr["Tiros_puerta_pp"]),
    ]
    _r_labels_c = _r_labels + [_r_labels[0]]
    _r_local_c  = _r_local  + [_r_local[0]]
    _r_rival_c  = _r_rival  + [_r_rival[0]]

    _fig_radar = go.Figure()
    _fig_radar.add_trace(go.Scatterpolar(
        r=_r_local_c, theta=_r_labels_c, fill="toself",
        fillcolor="rgba(74,158,202,0.20)",
        line=dict(color="#4a9eca", width=2),
        name=local,
        hovertemplate="%{theta}: %{r:.0f}/100<extra>" + local + "</extra>",
    ))
    _fig_radar.add_trace(go.Scatterpolar(
        r=_r_rival_c, theta=_r_labels_c, fill="toself",
        fillcolor="rgba(229,115,115,0.20)",
        line=dict(color="#e57373", width=2),
        name=rival,
        hovertemplate="%{theta}: %{r:.0f}/100<extra>" + rival + "</extra>",
    ))
    _fig_radar.update_layout(
        polar=dict(
            bgcolor="rgba(26,32,53,0.6)",
            radialaxis=dict(
                visible=True, range=[0, 100], showticklabels=False,
                gridcolor="#2d3d6b", linecolor="#2d3d6b",
            ),
            angularaxis=dict(
                tickfont=dict(color="#ccd6f6", size=11),
                gridcolor="#2d3d6b", linecolor="#2d3d6b",
            ),
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(font=dict(color="#ccd6f6", size=12), bgcolor="rgba(0,0,0,0)", orientation="h", y=-0.08),
        margin=dict(t=20, b=40, l=60, r=60),
        height=400,
    )
    st.plotly_chart(_fig_radar, use_container_width=True)

    # ── Forma reciente ────────────────────────────────────────────
    fl = inf["forma_local"]
    fr = inf["forma_rival"]

    st.markdown("<div style='margin-top:20px'></div>", unsafe_allow_html=True)

    def _badge(r):
        cfg = {"V": ("#27ae60", "#27ae6025"), "E": ("#e67e22", "#e67e2225"), "D": ("#e74c3c", "#e74c3c25")}
        col, bg = cfg.get(r, ("#8892b0", "#88929025"))
        return (f'<span style="background:{bg};border:1px solid {col};border-radius:6px;'
                f'padding:5px 11px;color:{col};font-weight:800;font-size:1rem">{r}</span>')

    def _forma_card(nombre, forma):
        if not forma:
            return (f'<div style="background:linear-gradient(145deg,#1a2035,#242c4a);'
                    f'border-radius:14px;padding:20px;border:1px solid #3d4870;'
                    f'box-shadow:0 4px 20px rgba(0,0,0,0.4);text-align:center">'
                    f'<div style="color:#8892b0;font-size:0.85rem">Sin datos de forma</div></div>')
        badges = " ".join(_badge(r) for r in forma.get("racha", []))
        stats = (
            '<div style="display:flex;justify-content:center;gap:24px">'
            f'<div style="text-align:center">'
            f'<div style="font-size:0.72rem;color:#8892b0;letter-spacing:0.04em;margin-bottom:4px">Media de puntos</div>'
            f'<div style="font-size:1.4rem;font-weight:900;color:#64b5f6">{forma.get("pts",0):.1f}</div>'
            f'</div>'
            f'<div style="text-align:center">'
            f'<div style="font-size:0.72rem;color:#8892b0;letter-spacing:0.04em;margin-bottom:4px">Goles a favor</div>'
            f'<div style="font-size:1.4rem;font-weight:900;color:#64b5f6">{forma.get("gf",0):.1f}</div>'
            f'</div>'
            f'<div style="text-align:center">'
            f'<div style="font-size:0.72rem;color:#8892b0;letter-spacing:0.04em;margin-bottom:4px">Goles en contra</div>'
            f'<div style="font-size:1.4rem;font-weight:900;color:#64b5f6">{forma.get("gc",0):.1f}</div>'
            f'</div>'
            '</div>'
        )
        return (
            f'<div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;'
            f'padding:22px 24px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);text-align:center">'
            f'<div style="font-size:0.88rem;color:#8892b0;letter-spacing:0.06em;'
            f'text-transform:uppercase;margin-bottom:14px">Forma reciente · <strong style="color:#ccd6f6">{nombre}</strong></div>'
            f'<div style="display:flex;justify-content:center">'
            f'<div style="display:inline-flex;flex-direction:column;align-items:stretch">'
            f'<div style="display:flex;gap:8px">{badges}</div>'
            f'<div style="display:flex;align-items:center;margin:10px 0 18px">'
            f'<div style="flex:1;height:1px;background:#4a9eca"></div>'
            f'<span style="color:#4a9eca;font-size:0.75rem;margin-left:3px">▶</span>'
            f'</div>'
            f'</div>'
            f'</div>'
            f'{stats}'
            f'</div>'
        )

    col_fl, col_fr = st.columns(2)
    with col_fl:
        st.markdown(_forma_card(local, fl), unsafe_allow_html=True)
    with col_fr:
        st.markdown(_forma_card(rival, fr), unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom:16px'></div>", unsafe_allow_html=True)

    # ── H2H ───────────────────────────────────────────────────────
    h2h = inf["h2h_df"]
    if not h2h.empty:
        def _res_h2h_html(r, mi):
            gh, ga = r["goles_home"], r["goles_away"]
            ganaste  = (r["Home"] == mi and gh > ga) or (r["Away"] == mi and ga > gh)
            empataste = gh == ga
            if ganaste:   return '<span style="color:#27ae60;font-weight:700">Victoria</span>'
            if empataste: return '<span style="color:#e67e22;font-weight:700">Empate</span>'
            return '<span style="color:#e74c3c;font-weight:700">Derrota</span>'

        _h2h_filas = ""
        for i, (_, row_h) in enumerate(h2h.iterrows()):
            bg = "#1a2540" if i % 2 == 0 else "#1e2b4a"
            partido = f"{row_h['Home']} {int(row_h['goles_home'])}–{int(row_h['goles_away'])} {row_h['Away']}"
            res     = _res_h2h_html(row_h, local)
            _h2h_filas += (
                f'<tr style="background:{bg}">'
                f'<td style="padding:9px 14px;color:#ffffff">{partido}</td>'
                f'<td style="padding:9px 14px;text-align:center;font-size:1rem">{res}</td>'
                f'</tr>'
            )

        st.markdown(f"""
<div style="border-radius:12px;overflow:hidden;border:1px solid #2d3d6b;background:#1a2540">
<table style="width:100%;border-collapse:collapse;font-size:0.9rem">
  <thead>
    <tr style="background:#1e3a6e">
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:left;font-size:0.88rem;letter-spacing:0.05em">HISTORIAL · {local.upper()} VS {rival.upper()}</th>
      <th style="padding:11px 14px;color:#F5C842;font-weight:800;text-align:center;font-size:0.88rem;letter-spacing:0.05em">RESULTADO</th>
    </tr>
  </thead>
  <tbody>{_h2h_filas}</tbody>
</table>
</div>""", unsafe_allow_html=True)
    else:
        st.caption(f"_No hay enfrentamientos previos registrados entre {local} y {rival} en los datos disponibles._")

    # ════════════════════════════════════════
    # BLOQUE 2 — ATAQUE
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">⚔️ Bloque 2 · Estrategia Ofensiva</div>',
                unsafe_allow_html=True)

    def _info_box(context, action=None):
        ctx = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', context)
        if action:
            act = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', action)
            html = (
                f'<div style="background:#142032;border:1px solid rgba(74,158,202,0.45);'
                f'border-radius:10px;overflow:hidden;font-size:0.95rem;line-height:1.6;margin-bottom:10px">'
                f'<div style="padding:14px 18px;color:#ccd6f6">{ctx}</div>'
                f'<div style="background:#1c3050;padding:11px 18px;'
                f'border-top:1px solid rgba(74,158,202,0.45)">'
                f'<span style="color:#F5C842;font-weight:700;font-size:0.78rem;'
                f'letter-spacing:0.06em;text-transform:uppercase">▸ Acción&nbsp;&nbsp;</span>'
                f'<span style="color:#ffffff;font-size:0.97rem">{act}</span></div>'
                f'</div>'
            )
        else:
            html = (
                f'<div style="background:#142032;border:1px solid rgba(74,158,202,0.45);'
                f'border-radius:10px;padding:14px 18px;font-size:0.95rem;'
                f'line-height:1.6;margin-bottom:10px;color:#ccd6f6">{ctx}</div>'
            )
        st.markdown(html, unsafe_allow_html=True)

    # ── Tarjetas potencial ofensivo / vulnerabilidad defensiva ───────
    _xg_rival_pp = rr['xG_contra'] / rr['PJ'] if rr['PJ'] > 0 else 0
    _por_rival_val = f"{rr['Paradas_pp']:.1f}"
    _por_rival_lbl = "Paradas / Partido"

    col_at1, col_at2 = st.columns(2)
    with col_at1:
        st.markdown(f"""
        <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
        padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
        margin-bottom:16px">
            <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:16px;text-align:center">
                Potencial ofensivo · {local}</div>
            <div style="display:flex;justify-content:space-around">
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">xG / partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rl['xG_pp']:.2f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Disparos a puerta / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rl['Tiros_puerta_pp']:.1f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">% tiros desde el área</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rl['Pct_tiro_area']:.0f}%</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_at2:
        st.markdown(f"""
        <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
        padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
        margin-bottom:16px">
            <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:16px;text-align:center">
                Vulnerabilidad defensiva · {rival}</div>
            <div style="display:flex;justify-content:space-around">
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">xG concedido / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{_xg_rival_pp:.2f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Disparos concedidos / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rr['Tiros_contra_pp']:.1f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">{_por_rival_lbl}</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{_por_rival_val}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Portero rival ──────────────────────────────────────────────
    if inf["por_ga90_rival"] is not None or inf["por_pct_p0_rival"] is not None or inf["por_pct_par_rival"] is not None:
        _items_por_r = []
        if inf["por_ga90_rival"] is not None:
            _items_por_r.append(("Goles encajados / Partido", f"{inf['por_ga90_rival']:.2f}"))
        if inf["por_pct_par_rival"] is not None:
            _items_por_r.append(("% Paradas", f"{inf['por_pct_par_rival']:.0f}%"))
        if inf["por_pct_p0_rival"] is not None:
            _items_por_r.append(("% Portería a cero", f"{inf['por_pct_p0_rival']:.0f}%"))
        _por_flex_r = "".join(
            f'<div style="text-align:center">'
            f'<div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">{lbl}</div>'
            f'<div style="color:#64b5f6;font-size:1.3rem;font-weight:700">{val}</div></div>'
            for lbl, val in _items_por_r
        )
        st.markdown(f"""
        <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
        padding:20px 22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
        margin-bottom:16px">
            <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:4px;text-align:center">
                Portero rival · {rival}</div>
            <div style="color:#ccd6f6;font-size:1.0rem;font-weight:600;text-align:center;margin-bottom:14px">
                {inf['por_name_rival'] if inf['por_name_rival'] else ''}</div>
            <div style="display:flex;justify-content:space-around">{_por_flex_r}</div>
        </div>
        """, unsafe_allow_html=True)

        if inf["por_ga90_rival"] is not None:
            if inf["por_ga90_rival"] > 1.5:
                _info_box(
                    f"El portero de {rival} encaja **{inf['por_ga90_rival']:.2f} goles por 90 minutos jugados**, una cifra alta.",
                    "Genera volumen de tiros: este portero es un punto débil del rival."
                )
            elif inf["por_ga90_rival"] < 0.9:
                _info_box(
                    f"El portero de {rival} es muy sólido: **{inf['por_ga90_rival']:.2f} goles encajados por 90 minutos**.",
                    "Necesitarás generar ocasiones de muy alta calidad para batirle."
                )

    # ── Distribución goleadora de {local} ────────────────────────
    ga_l = inf["gol_abierto_local"]
    gb_l = inf["gol_bparado_local"]
    gc_l = inf["gol_ct_local"]
    # Guardamos datos del rival para Bloque 3
    ga_r = inf["gol_abierto_rival"]
    gb_r = inf["gol_bparado_rival"]
    gc_r = inf["gol_ct_rival"]

    if any(v is not None for v in [ga_l, gb_l, gc_l]):
        rows_tipo = [
            ("Juego abierto", f"{ga_l:.1f}" if ga_l is not None else "–"),
            ("Balón parado",  f"{gb_l:.1f}" if gb_l is not None else "–"),
            ("Contraataque",  f"{gc_l:.1f}" if gc_l is not None else "–"),
        ]
        rows_html_tipo = "".join(
            f'<tr style="background:{"#1a2540" if i % 2 == 0 else "#1e2b4a"}">'
            f'<td style="padding:10px 16px;color:#ffffff;text-align:left">{t}</td>'
            f'<td style="padding:10px 16px;color:#ffffff;text-align:center;font-weight:700">{v}</td></tr>'
            for i, (t, v) in enumerate(rows_tipo)
        )
        st.markdown(f"""
        <div style="margin-bottom:16px">
            <div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;
            text-transform:uppercase;margin-bottom:8px">Distribución goleadora de {local}</div>
            <div style="border-radius:12px;overflow:hidden;border:1px solid #2d3d6b;background:#1a2540">
                <table style="width:100%;border-collapse:collapse">
                    <thead><tr style="background:#1e3a6e">
                        <th style="padding:10px 16px;color:#F5C842;text-align:left;font-weight:600">Situación</th>
                        <th style="padding:10px 16px;color:#F5C842;text-align:center;font-weight:600">Goles marcados</th>
                    </tr></thead>
                    <tbody>{rows_html_tipo}</tbody>
                </table>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Recomendaciones ofensivas ─────────────────────────────────
    recs_ataque = []

    if rl["xG_pp"] > rr["xG_contra"] / rr["PJ"] * 1.15:
        recs_ataque.append((
            f"Superioridad ofensiva: generas **{rl['xG_pp']:.2f} goles esperados por partido** frente a "
            f"{rr['xG_contra']/rr['PJ']:.2f} que concede {rival}.",
            "Mantén el volumen de tiro y prioriza disparos desde dentro del área."
        ))
    elif rl["xG_pp"] < rr["xG_contra"] / rr["PJ"] * 0.85:
        recs_ataque.append((
            f"El rival es defensivamente sólido: solo concede **{rr['xG_contra']/rr['PJ']:.2f} goles esperados por partido** "
            f"(calidad de disparo propia: {rl['xG_tiro']:.3f} xG por tiro).",
            "Prioriza calidad de disparo sobre volumen. Busca situaciones claras antes de rematar."
        ))
    else:
        recs_ataque.append((
            f"El rival concede **{rr['xG_contra']/rr['PJ']:.2f} goles esperados por partido**, "
            f"cercano a lo que generas tú ({rl['xG_pp']:.2f}).",
            "Los cambios de ritmo y las transiciones rápidas serán clave para desequilibrar."
        ))

    if rl["Pct_tiro_area"] < 45:
        recs_ataque.append((
            f"Solo el **{rl['Pct_tiro_area']:.0f}%** de tus disparos parten desde dentro del área "
            f"(media Liga ≈50%).",
            "Trabaja combinaciones interiores para mejorar la calidad de las oportunidades de gol."
        ))

    if inf["alhueco_local"] is not None and not ws_df.empty:
        med = ws_df["pase_alHueco_fav_gen"].median() if "pase_alHueco_fav_gen" in ws_df.columns else None
        if med is not None and inf["alhueco_local"] > med * 1.2:
            recs_ataque.append((
                f"{local} genera **{inf['alhueco_local']:.1f} pases al hueco por partido** "
                f"(media de la Liga: {med:.1f}).",
                "Aprovecha esta ventaja en las transiciones para atacar la espalda defensiva rival."
            ))

    if rl["Reg_exit"] and rr["Entradas_fall"]:
        if rl["Reg_exit"] > df["Reg_exit"].quantile(0.65):
            recs_ataque.append((
                f"Capacidad de regate por encima de la media (**{rl['Reg_exit']:.1f} regates exitosos por partido**).",
                "Utiliza los perfiles rápidos en banda para desequilibrar en el 1vs1."
            ))

    if rr["Intercepciones"] > df["Intercepciones"].quantile(0.70):
        recs_ataque.append((
            f"**{rival}** es uno de los equipos con más intercepciones de la Liga "
            f"({rr['Intercepciones']:.0f}).",
            "Varía el juego con pases en profundidad y evita el exceso de combinación corta "
            "en zonas de presión."
        ))

    if rl["Pos%"] > 55:
        recs_ataque.append((
            f"Posesión elevada ({rl['Pos%']:.1f}%) no garantiza ocasiones de calidad: "
            f"los equipos con más posesión generan un menor % de tiros desde dentro del área "
            f"(r=−0.644, p<0.01).",
            "Cuida la verticalidad: la posesión debe ser productiva, no estéril."
        ))

    if recs_ataque:
        st.markdown(
            '<div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;'
            'text-transform:uppercase;margin-bottom:8px;margin-top:8px">'
            'Recomendaciones ofensivas</div>',
            unsafe_allow_html=True
        )
        for ctx, act in recs_ataque:
            _info_box(ctx, act)

    # ════════════════════════════════════════
    # BLOQUE 3 — DEFENSA
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">🛡️ Bloque 3 · Estrategia Defensiva</div>',
                unsafe_allow_html=True)

    col_def1, col_def2 = st.columns(2)
    with col_def1:
        st.markdown(f"""
        <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
        padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
        margin-bottom:16px">
            <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:16px;text-align:center">
                Amenaza ofensiva · {rival}</div>
            <div style="display:flex;justify-content:space-around">
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">xG / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rr['xG_pp']:.2f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Disparos a puerta / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rr['Tiros_puerta_pp']:.1f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">% tiros desde el área</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rr['Pct_tiro_area']:.0f}%</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_def2:
        st.markdown(f"""
        <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
        padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
        margin-bottom:16px">
            <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:16px;text-align:center">
                Solidez defensiva · {local}</div>
            <div style="display:flex;justify-content:space-around">
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">xG concedido / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rl['xG_contra']/rl['PJ']:.2f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Disparos concedidos / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rl['Tiros_contra_pp']:.1f}</div>
                </div>
                <div style="text-align:center">
                    <div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Paradas / Partido</div>
                    <div style="color:#64b5f6;font-size:1.5rem;font-weight:700">{rl['Paradas_pp']:.1f}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Portero propio ─────────────────────────────────────────────
    if inf["por_ga90_local"] is not None or inf["por_pct_p0_local"] is not None or inf["por_pct_par_local"] is not None:
        _items_por_l = []
        if inf["por_ga90_local"] is not None:
            _items_por_l.append(("Goles encajados / Partido", f"{inf['por_ga90_local']:.2f}"))
        if inf["por_pct_par_local"] is not None:
            _items_por_l.append(("% Paradas", f"{inf['por_pct_par_local']:.0f}%"))
        if inf["por_pct_p0_local"] is not None:
            _items_por_l.append(("% Portería a cero", f"{inf['por_pct_p0_local']:.0f}%"))
        _por_flex_l = "".join(
            f'<div style="text-align:center">'
            f'<div style="color:#8892b0;font-size:0.75rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">{lbl}</div>'
            f'<div style="color:#64b5f6;font-size:1.3rem;font-weight:700">{val}</div></div>'
            for lbl, val in _items_por_l
        )
        st.markdown(f"""
        <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
        padding:20px 22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
        margin-bottom:16px">
            <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:4px;text-align:center">
                Tu portero · {local}</div>
            <div style="color:#ccd6f6;font-size:1.0rem;font-weight:600;text-align:center;margin-bottom:14px">
                {inf['por_name_local'] if inf['por_name_local'] else ''}</div>
            <div style="display:flex;justify-content:space-around">{_por_flex_l}</div>
        </div>
        """, unsafe_allow_html=True)

        if inf["por_ga90_local"] is not None:
            if inf["por_ga90_local"] < 0.9:
                _info_box(
                    f"Tu portero está en un momento excelente: **{inf['por_ga90_local']:.2f} goles encajados por 90 minutos**.",
                    "Aprovecha la solidez del portero como activo: el equipo puede competir con un bloque más bajo si la situación lo requiere."
                )
            elif inf["por_ga90_local"] > 1.5:
                _info_box(
                    f"Tu portero está encajando muchos goles: **{inf['por_ga90_local']:.2f} por 90 minutos**.",
                    "Prioriza la organización defensiva para reducir la carga de trabajo bajo palos."
                )

    # ── Patrones ofensivos del rival (inteligencia defensiva) ─────
    if any(v is not None for v in [ga_r, gb_r, gc_r]):
        st.markdown(
            f'<div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;'
            f'text-transform:uppercase;margin-bottom:8px">Cómo ataca {rival}</div>',
            unsafe_allow_html=True
        )
        if gb_r is not None and ga_r is not None and (ga_r + gb_r) > 0:
            pct_bp_rival = gb_r / (ga_r + gb_r) * 100
            if pct_bp_rival > 35:
                _info_box(
                    f"{rival} es peligroso en balón parado: el **{pct_bp_rival:.0f}%** "
                    f"de sus goles provienen de balón parado.",
                    "Atención especial en córners, faltas directas y segundas jugadas."
                )
            elif pct_bp_rival < 20:
                _info_box(
                    f"{rival} apenas marca en balón parado (**{pct_bp_rival:.0f}%** de sus goles).",
                    "Mantén el bloque compacto en transiciones — el peligro llega del juego dinámico, no del balón parado."
                )
        if gc_r is not None and gc_r > 0:
            _info_box(
                f"{rival} anotó **{gc_r:.1f} goles de contraataque** esta temporada.",
                "Cuidado con las transiciones: gestiona la recuperación del bloque al perder el balón."
                if gc_r >= 3 else
                "Presiona en alto con seguridad — el contraataque rival no representa una amenaza significativa."
            )

    if inf["goles_area_rival"] is not None and inf["goles_fuera_rival"] is not None:
        tot_gol_rival = inf["goles_area_rival"] + inf["goles_fuera_rival"]
        if tot_gol_rival > 0:
            pct_dentro = inf["goles_area_rival"] / tot_gol_rival * 100
            _info_box(
                f"El **{pct_dentro:.0f}%** de los goles de {rival} provienen de dentro del área.",
                "Bloquea centros y cierres al segundo palo."
                if pct_dentro > 70 else
                "También remata desde fuera: mantén el bloque bajo pero compacto."
            )

    if inf["alhueco_rival"] is not None and not ws_df.empty:
        med_hueco = ws_df["pase_alHueco_fav_gen"].median() if "pase_alHueco_fav_gen" in ws_df.columns else None
        if med_hueco is not None and abs(inf["alhueco_rival"] - med_hueco) >= 0.3:
            _info_box(
                f"Pases al hueco de {rival}: **{inf['alhueco_rival']:.1f} por partido** "
                f"({'por encima' if inf['alhueco_rival'] > med_hueco else 'por debajo'} "
                f"de la media de la Liga: {med_hueco:.1f}).",
                "Vigila la espalda defensiva: el rival busca el espacio con pases filtrados."
                if inf["alhueco_rival"] > med_hueco else
                "Mantén la línea defensiva alta — el rival raramente busca el espacio a la espalda."
            )

    if all(v is not None for v in [inf["lado_izq_rival"], inf["lado_cen_rival"], inf["lado_der_rival"]]):
        tot_lado = inf["lado_izq_rival"] + inf["lado_cen_rival"] + inf["lado_der_rival"]
        if tot_lado > 0:
            pi  = inf["lado_izq_rival"] / tot_lado * 100
            pc  = inf["lado_cen_rival"] / tot_lado * 100
            pd_r = inf["lado_der_rival"] / tot_lado * 100
            lado_dom = max(("izquierda", pi), ("centro", pc), ("derecha", pd_r), key=lambda x: x[1])
            if lado_dom[1] > 40:
                _info_box(
                    f"{rival} ataca preferentemente por **{lado_dom[0]}** "
                    f"({lado_dom[1]:.0f}% de sus acciones ofensivas).",
                    "Refuerza ese costado defensivo y limita el desborde por esa banda."
                )

    # ── Recomendaciones defensivas ─────────────────────────────────
    recs_defensa = []

    if rr["xG_pp"] > df["xG_pp"].quantile(0.75):
        recs_defensa.append((
            f"El rival es muy peligroso ofensivamente: **{rr['xG_pp']:.2f} goles esperados por partido** (top 25% de la Liga).",
            "Considera bloque medio bajo y limita los espacios a la espalda de la defensa."
        ))
    elif rr["xG_pp"] < df["xG_pp"].quantile(0.30):
        recs_defensa.append((
            f"El rival tiene un ataque limitado: **{rr['xG_pp']:.2f} goles esperados por partido**.",
            "La defensa puede subir la línea para presionar la salida de balón rival."
        ))
    else:
        recs_defensa.append((
            f"El rival tiene un ataque moderado: **{rr['xG_pp']:.2f} goles esperados por partido**.",
            "Mantén bloque medio con atención especial en las transiciones."
        ))

    if inf["tiros_peq_rival"] is not None and not ws_df.empty:
        med_peq = ws_df["tiros_areaPequeña_gen"].median() if "tiros_areaPequeña_gen" in ws_df.columns else None
        if med_peq is not None and inf["tiros_peq_rival"] > med_peq * 1.3:
            recs_defensa.append((
                f"{rival} finaliza mucho desde el área pequeña: "
                f"**{inf['tiros_peq_rival']:.1f} tiros por partido** desde el área pequeña (media Liga: {med_peq:.1f}).",
                "El cierre de centros al segundo palo es prioritario."
            ))

    if rr["Pct_tiro_area"] > 55:
        recs_defensa.append((
            f"{rival} genera el **{rr['Pct_tiro_area']:.0f}%** de sus tiros desde dentro del área.",
            "Fundamental el cierre de centros: puntos de penalti y segundos palos son zonas críticas."
        ))

    if rr["Pct_aereo"] and rl["Pct_aereo"]:
        if rr["Pct_aereo"] > 55 and rl["Pct_aereo"] < 50:
            recs_defensa.append((
                f"Desventaja aérea: {rival} gana el **{rr['Pct_aereo']:.0f}%** de los duelos "
                f"aéreos vs {rl['Pct_aereo']:.0f}% propio.",
                "Reduce los despejes largos y favorece la salida de balón jugada."
            ))
        elif rl["Pct_aereo"] > rr["Pct_aereo"] + 10:
            recs_defensa.append((
                f"Superioridad aérea: **{rl['Pct_aereo']:.0f}%** vs {rr['Pct_aereo']:.0f}% del rival.",
                "Utiliza el juego de largo a balón parado como herramienta estratégica."
            ))

    if rr["Zona_Ata"] and rr["Zona_Ata"] > 35:
        recs_defensa.append((
            f"{rival} pasa el **{rr['Zona_Ata']:.0f}%** del tiempo en zona de ataque.",
            "Prepara la salida de balón bajo presión alta."
        ))

    if rl["GC_pp"] > df["GC_pp"].quantile(0.70):
        recs_defensa.append((
            f"Solidez defensiva por mejorar: encajas **{rl['GC_pp']:.2f} goles por partido** "
            f"(percentil {percentil(rl['GC_pp'], df['GC_pp'])}% en la Liga).",
            "Prioridad: reducir los xG concedidos desde dentro del área."
        ))

    if inf["rojas_local"] is not None and inf["rojas_local"] > df["Rojas"].quantile(0.70):
        recs_defensa.append((
            f"{local} acumula **{inf['rojas_local']:.0f} tarjetas rojas** esta temporada "
            f"(percentil {percentil(inf['rojas_local'], df['Rojas'].dropna())}% en la Liga).",
            "Alto riesgo de inferioridad numérica — gestiona la disciplina."
        ))

    if recs_defensa:
        st.markdown(
            '<div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;'
            'text-transform:uppercase;margin-bottom:8px;margin-top:8px">'
            'Recomendaciones defensivas</div>',
            unsafe_allow_html=True
        )
        for ctx, act in recs_defensa:
            _info_box(ctx, act)

    # ── Duplas peligrosas a neutralizar ───────────────────────────
    if not inf["dup_rival"].empty:
        dup_rows_html = "".join(
            f'<tr style="background:{"#1a2540" if i % 2 == 0 else "#1e2b4a"}">'
            f'<td style="padding:10px 16px;color:#ffffff;text-align:left">'
            f'<strong>{row_dup["Goleador"]}</strong></td>'
            f'<td style="padding:10px 16px;color:#ffffff;text-align:left">'
            f'<strong>{row_dup["Asistidor"]}</strong></td>'
            f'<td style="padding:10px 16px;color:#ffffff;text-align:center">'
            f'{int(row_dup["Frecuencia"])}</td></tr>'
            for i, (_, row_dup) in enumerate(inf["dup_rival"].iterrows())
        )
        st.markdown(f"""
        <div style="margin-top:16px;margin-bottom:16px">
            <div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;
            text-transform:uppercase;margin-bottom:8px">Duplas peligrosas de {rival} a neutralizar</div>
            <div style="border-radius:12px;overflow:hidden;border:1px solid #2d3d6b;background:#1a2540">
                <table style="width:100%;border-collapse:collapse">
                    <thead><tr style="background:#1e3a6e">
                        <th style="padding:10px 16px;color:#F5C842;text-align:left;font-weight:600">Rematador</th>
                        <th style="padding:10px 16px;color:#F5C842;text-align:left;font-weight:600">Asistidor</th>
                        <th style="padding:10px 16px;color:#F5C842;text-align:center;font-weight:600">Combinaciones</th>
                    </tr></thead>
                    <tbody>{dup_rows_html}</tbody>
                </table>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ════════════════════════════════════════
    # BLOQUE 4 — JUGADORES CLAVE
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">🌟 Bloque 4 · Jugadores Clave</div>',
                unsafe_allow_html=True)

    col_jk1, col_jk2 = st.columns(2)

    with col_jk1:
        # ── Alineación propia ─────────────────────────────────────
        lineup_df = pd.DataFrame(lineup_local)
        is_map = inf["jug_local_disp"].set_index("Jugador")["IS_indiv"].to_dict()
        lineup_df["IS_individual"] = lineup_df["Jugador"].map(is_map).round(3)
        _lu_rows = "".join(
            f'<tr style="background:{"#1a2540" if i % 2 == 0 else "#1e2b4a"}">'
            f'<td style="padding:8px 14px;color:#ffffff;font-size:0.82rem">{row["Posicion"]}</td>'
            f'<td style="padding:8px 14px;color:#ffffff;font-weight:500">{row["Jugador"]}</td>'
            f'<td style="padding:8px 14px;color:#ffffff;text-align:center">'
            f'{row["IS_individual"] if pd.notna(row["IS_individual"]) else "–"}</td></tr>'
            for i, (_, row) in enumerate(lineup_df.iterrows())
        )
        st.markdown(f"""
        <div style="margin-bottom:16px">
            <div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;margin-bottom:8px">
                Tu alineación · {formacion}</div>
            <div style="border-radius:12px;overflow:hidden;border:1px solid #2d3d6b;background:#1a2540">
                <table style="width:100%;border-collapse:collapse">
                    <thead><tr style="background:#1e3a6e">
                        <th style="padding:9px 14px;color:#F5C842;text-align:left;font-weight:600;font-size:0.82rem">Posición</th>
                        <th style="padding:9px 14px;color:#F5C842;text-align:left;font-weight:600;font-size:0.82rem">Jugador</th>
                        <th style="padding:9px 14px;color:#F5C842;text-align:center;font-weight:600;font-size:0.82rem">IS</th>
                    </tr></thead>
                    <tbody>{_lu_rows}</tbody>
                </table>
            </div>
        </div>
        """, unsafe_allow_html=True)

        top_local_lineup = (
            inf["jug_local_disp"]
            [inf["jug_local_disp"]["Jugador"].isin([r["Jugador"] for r in lineup_local])]
            .sort_values("IS_indiv", ascending=False)
            .head(3)
        )
        if not top_local_lineup.empty:
            st.markdown(
                '<div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;'
                'text-transform:uppercase;margin-bottom:6px">'
                'Jugadores más influyentes de tu alineación</div>',
                unsafe_allow_html=True
            )
            _pos_map_local = {r["Jugador"]: r["Posicion"] for r in lineup_local}
            for _, row_jug in top_local_lineup.iterrows():
                _pct_min_l = row_jug["Minutos"] / (rl["PJ"] * 90) * 100
                _pos_label_l = _pos_map_local.get(row_jug["Jugador"], row_jug.get("PosGrupo", ""))
                _info_box(
                    f"**{row_jug['Jugador']}** ({_pos_label_l}) — "
                    f"IS: **{row_jug['IS_indiv']:.3f}** · {_pct_min_l:.0f}% de minutos disponibles"
                )

    with col_jk2:
        # ── Alineación estimada del rival ─────────────────────────
        if inf["rival_lineup"].empty:
            st.warning("No hay datos de jugadores del rival.")
        else:
            rival_lineup_show = inf["rival_lineup"].copy()
            rival_lineup_show["IS_indiv"] = rival_lineup_show["IS_indiv"].round(3)
            rival_lineup_show["Minutos"]  = rival_lineup_show["Minutos"].astype(int)
            _POR_ORD = {
                "GK":0,"P":0,"PT":0,"POR":0,"GOR":0,"PORTERO":0,
                "DC":1,"DFC":1,"DL":1,"DR":1,"LD":1,"LI":1,"CB":1,"LB":1,"RB":1,
                "DF":1,"DFD":1,"DFI":1,"DEFENSA":1,"LATERAL":1,
                "MC":2,"MF":2,"CM":2,"MED":2,"CENTROCAMPISTA":2,"MEDIOCAMPISTA":2,
                "FC":3,"FW":3,"SS":3,"ST":3,"CF":3,"EI":3,"ED":3,"EC":3,
                "DEL":3,"AT":3,"FWL":3,"FWR":3,"DELANTERO":3,"EXTREMO":3,"ATACANTE":3,
            }
            def _pos_sort_key(p):
                q = str(p).split(",")[0].strip().upper()
                if q in _POR_ORD: return _POR_ORD[q]
                if q.startswith("PORT"): return 0
                if q.startswith("DEL") or q.startswith("EXT") or q.startswith("ATA"): return 3
                if q.startswith("DEF") or q.startswith("LAT"): return 1
                return 2
            rival_lineup_show = rival_lineup_show.sort_values(
                "Posicion", key=lambda s: s.map(_pos_sort_key)
            ).reset_index(drop=True)
            _rl_rows = "".join(
                f'<tr style="background:{"#1a2540" if i % 2 == 0 else "#1e2b4a"}">'
                f'<td style="padding:8px 14px;color:#ffffff;font-size:0.82rem">{row["Posicion"]}</td>'
                f'<td style="padding:8px 14px;color:#ffffff;font-weight:500">{row["Jugador"]}</td>'
                f'<td style="padding:8px 14px;color:#ffffff;text-align:center">{row["IS_indiv"]:.3f}</td></tr>'
                for i, (_, row) in enumerate(rival_lineup_show.iterrows())
            )
            st.markdown(f"""
            <div style="margin-bottom:16px">
                <div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;margin-bottom:8px">
                    Alineación estimada · {rival}</div>
                <div style="border-radius:12px;overflow:hidden;border:1px solid #2d3d6b;background:#1a2540">
                    <table style="width:100%;border-collapse:collapse">
                        <thead><tr style="background:#1e3a6e">
                            <th style="padding:9px 14px;color:#F5C842;text-align:left;font-weight:600;font-size:0.82rem">Posición</th>
                            <th style="padding:9px 14px;color:#F5C842;text-align:left;font-weight:600;font-size:0.82rem">Jugador</th>
                            <th style="padding:9px 14px;color:#F5C842;text-align:center;font-weight:600;font-size:0.82rem">IS</th>
                        </tr></thead>
                        <tbody>{_rl_rows}</tbody>
                    </table>
                </div>
            </div>
            """, unsafe_allow_html=True)

            top_rival = inf["rival_lineup"].sort_values("IS_indiv", ascending=False).head(3)
            st.markdown(
                f'<div style="color:#8892b0;font-size:0.85rem;font-weight:600;letter-spacing:0.06em;'
                f'text-transform:uppercase;margin-bottom:6px">'
                f'Amenazas principales de {rival} a neutralizar</div>',
                unsafe_allow_html=True
            )
            for _, row_jug in top_rival.iterrows():
                _pct_min_r = row_jug["Minutos"] / (rr["PJ"] * 90) * 100
                _info_box(
                    f"**{row_jug['Jugador']}** ({row_jug['Posicion']}) — "
                    f"IS: **{row_jug['IS_indiv']:.3f}** · {_pct_min_r:.0f}% de minutos disponibles"
                )

    # ── Tarjetas de calidad de plantilla (fuera de las columnas para alineación) ──
    def _quality_card(team_name, is_top5, xg11, scorer_name, scorer_goles, scorer_pct,
                      assist_name, assist_ast, assist_pct):
        def _cell(lbl, val_main, val_sub=None):
            sub_html = (f'<div style="color:#a0a8c0;font-size:0.75rem;margin-top:2px">{val_sub}</div>'
                        if val_sub else "")
            return (f'<div style="text-align:center;padding:8px 4px;background:rgba(255,255,255,0.03);'
                    f'border-radius:8px">'
                    f'<div style="color:#8892b0;font-size:0.72rem;margin-bottom:6px;line-height:1.3;'
                    f'font-weight:600;text-transform:uppercase;letter-spacing:0.04em">{lbl}</div>'
                    f'<div style="color:#64b5f6;font-size:1.05rem;font-weight:700">{val_main}</div>'
                    f'{sub_html}</div>')

        cells = []
        if is_top5 is not None:
            cells.append(_cell("IS medio · top 5", f"{is_top5:.3f}"))
        if xg11 is not None:
            cells.append(_cell("xG por partido · once titular", f"{xg11:.3f}"))
        if scorer_name and scorer_goles is not None:
            pct_str = f"({scorer_pct:.0f}% del equipo)" if scorer_pct is not None else ""
            cells.append(_cell("Máximo goleador", scorer_name,
                               f"{scorer_goles} goles {pct_str}".strip()))
        if assist_name and assist_ast is not None:
            pct_str = f"({assist_pct:.0f}% del equipo)" if assist_pct is not None else ""
            cells.append(_cell("Máximo asistente", assist_name,
                               f"{assist_ast} asistencias {pct_str}".strip()))

        if not cells:
            return
        cols_n = min(len(cells), 2)
        grid_html = "".join(cells)
        st.markdown(f"""
        <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:12px;
        padding:16px 18px;border:1px solid #3d4870;margin-top:12px">
            <div style="color:#8892b0;font-size:0.82rem;font-weight:600;letter-spacing:0.05em;
            text-transform:uppercase;margin-bottom:12px">Calidad de plantilla · {team_name}</div>
            <div style="display:grid;grid-template-columns:repeat({cols_n},1fr);gap:8px">{grid_html}</div>
        </div>
        """, unsafe_allow_html=True)

    col_sq1, col_sq2 = st.columns(2)
    with col_sq1:
        _quality_card(
            local,
            inf["jf_avg_is_top5_local"], inf["jf_xg11_local"],
            inf["jf_scorer_name_local"], inf["jf_scorer_goles_local"], inf["jf_scorer_pct_local"],
            inf["jf_assist_name_local"], inf["jf_assist_ast_local"], inf["jf_assist_pct_local"],
        )
    with col_sq2:
        _quality_card(
            rival,
            inf["jf_avg_is_top5_rival"], inf["jf_xg11_rival"],
            inf["jf_scorer_name_rival"], inf["jf_scorer_goles_rival"], inf["jf_scorer_pct_rival"],
            inf["jf_assist_name_rival"], inf["jf_assist_ast_rival"], inf["jf_assist_pct_rival"],
        )

    if inf["jf_avg_is_top5_local"] is not None and inf["jf_avg_is_top5_rival"] is not None:
        delta_is5 = inf["jf_avg_is_top5_local"] - inf["jf_avg_is_top5_rival"]
        if delta_is5 > 0.05:
            _info_box(
                f"Tu IS medio top-5 es superior al del rival "
                f"(**{inf['jf_avg_is_top5_local']:.3f}** vs {inf['jf_avg_is_top5_rival']:.3f}).",
                "Los cambios en el banquillo pueden marcar la diferencia en el tramo final."
            )
        elif delta_is5 < -0.05:
            _info_box(
                f"El rival tiene mejor profundidad de plantilla "
                f"(IS top-5: **{inf['jf_avg_is_top5_rival']:.3f}** vs {inf['jf_avg_is_top5_local']:.3f}).",
                "Gestiona bien la energía del equipo y evita desgastar a los titulares innecesariamente."
            )

    if (inf["jf_scorer_pct_rival"] is not None and inf["jf_scorer_pct_rival"] > 33
            and inf["jf_scorer_name_rival"]):
        _info_box(
            f"El rival concentra el **{inf['jf_scorer_pct_rival']:.0f}%** de sus goles "
            f"en **{inf['jf_scorer_name_rival']}**.",
            "Asigna marcaje específico sobre él — neutralizarle equivale a desactivar el principal recurso goleador del rival."
        )

    if (inf["jf_assist_pct_rival"] is not None and inf["jf_assist_pct_rival"] > 33
            and inf["jf_assist_name_rival"]):
        _info_box(
            f"**{inf['jf_assist_name_rival']}** genera el **{inf['jf_assist_pct_rival']:.0f}%** "
            f"de las asistencias de {rival}.",
            "Córtale el balón en la fase de construcción — limitar su participación desorganiza el juego combinativo rival."
        )

    if inf["ausentes_rival"]:
        _info_box(
            f"Bajas confirmadas de {rival}: **{', '.join(inf['ausentes_rival'])}**.",
            "Adapta la estrategia para explotar los huecos generados por los sustitutos."
        )

    # ════════════════════════════════════════
    # BLOQUE 5 — CONTEXTO
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">🌦️ Bloque 5 · Contexto del Partido</div>',
                unsafe_allow_html=True)

    col_ctx1, col_ctx2 = st.columns(2)

    with col_ctx1:
        # ── Árbitro ───────────────────────────────────────────────
        if arbitro_sel == "(Desconocido)":
            st.markdown(
                '<div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;'
                'padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);'
                'margin-bottom:16px"><div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;'
                'margin-bottom:10px;text-align:center">Análisis del árbitro</div>'
                '<div style="color:#8892b0;font-size:0.9rem">Árbitro no seleccionado.</div></div>',
                unsafe_allow_html=True
            )
        elif inf["arb_faltas"] is not None:
            _delta_f = inf['arb_faltas'] - inf['liga_faltas']
            _delta_a = inf['arb_amarillas'] - inf['liga_amarillas']
            _delta_r = (inf['arb_rojas'] - inf['liga_rojas']) if inf['arb_rojas'] is not None else None
            _delta_p = (inf['arb_penaltis'] - inf['liga_penaltis']) if inf['arb_penaltis'] is not None else None

            _arb_stats = [
                ("Faltas / Partido",    f"{inf['arb_faltas']:.1f}",   f"{_delta_f:+.1f} vs Liga",  "#4caf50" if _delta_f >= 0 else "#ef5350"),
                ("Amarillas / Partido", f"{inf['arb_amarillas']:.2f}", f"{_delta_a:+.2f} vs Liga",  "#ef5350" if _delta_a >= 0 else "#4caf50"),
            ]
            if inf['arb_rojas'] is not None:
                _arb_stats.append(("Rojas / Partido", f"{inf['arb_rojas']:.2f}", f"{_delta_r:+.2f} vs Liga", "#ef5350" if _delta_r >= 0 else "#4caf50"))
            if inf['arb_penaltis'] is not None:
                _arb_stats.append(("Penaltis / Partido", f"{inf['arb_penaltis']:.2f}", f"{_delta_p:+.2f} vs Liga", "#ef5350" if _delta_p >= 0 else "#4caf50"))

            _arb_flex = "".join(
                f'<div style="text-align:center">'
                f'<div style="color:#8892b0;font-size:0.72rem;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">{lbl}</div>'
                f'<div style="color:#64b5f6;font-size:1.4rem;font-weight:700">{val}</div>'
                f'<div style="color:{col};font-size:0.72rem;margin-top:2px">{delta}</div></div>'
                for lbl, val, delta, col in _arb_stats
            )

            _pct_local = inf['arb_pct_local']
            _pct_emp   = inf['arb_pct_emp']
            _pct_vis   = inf['arb_pct_vis']
            _pct_row = ""
            if all(v is not None for v in [_pct_local, _pct_emp, _pct_vis]):
                _pct_row = (
                    f'<div style="display:flex;justify-content:space-around;margin-top:16px;'
                    f'padding-top:12px;border-top:1px solid rgba(255,255,255,0.08)">'
                    f'<div style="text-align:center">'
                    f'<div style="color:#8892b0;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px">Victoria local</div>'
                    f'<div style="color:#64b5f6;font-size:1.3rem;font-weight:700">{_pct_local:.0f}%</div>'
                    f'<div style="color:#8892b0;font-size:0.7rem">Liga: {inf["liga_pct_local"]:.0f}%</div></div>'
                    f'<div style="text-align:center">'
                    f'<div style="color:#8892b0;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px">Empate</div>'
                    f'<div style="color:#64b5f6;font-size:1.3rem;font-weight:700">{_pct_emp:.0f}%</div></div>'
                    f'<div style="text-align:center">'
                    f'<div style="color:#8892b0;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px">Victoria visitante</div>'
                    f'<div style="color:#64b5f6;font-size:1.3rem;font-weight:700">{_pct_vis:.0f}%</div>'
                    f'<div style="color:#8892b0;font-size:0.7rem">Liga: {inf["liga_pct_vis"]:.0f}%</div></div>'
                    f'</div>'
                )

            st.markdown(f"""
            <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
            padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
            margin-bottom:16px">
                <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:16px;text-align:center">
                    Análisis del árbitro</div>
                <div style="display:flex;justify-content:space-around;flex-wrap:wrap;gap:8px">{_arb_flex}</div>
                {_pct_row}
            </div>
            """, unsafe_allow_html=True)

            # ── Faltas y disciplina ────────────────────────────────
            if inf["arb_faltas"] > inf["liga_faltas"] + 1:
                _info_box(
                    f"El árbitro es estricto: pita **{inf['arb_faltas']:.1f} faltas por partido** "
                    f"(media Liga: {inf['liga_faltas']:.1f}).",
                    "Prioriza el pressing organizado sobre las entradas individuales. Cualquier falta en zona peligrosa puede costar caro."
                )
            elif inf["arb_faltas"] < inf["liga_faltas"] - 1:
                _info_box(
                    f"El árbitro es permisivo con el juego físico: solo pita **{inf['arb_faltas']:.1f} faltas por partido** "
                    f"(media Liga: {inf['liga_faltas']:.1f}).",
                    "Puedes apretar la marca y ganar duelos físicos con menos riesgo de falta. Vigila igualmente las amarillas por acumulación."
                )
            else:
                _info_box(
                    f"El árbitro pita **{inf['arb_faltas']:.1f} faltas por partido**, dentro de la media de la Liga ({inf['liga_faltas']:.1f}).",
                    "Sin condicionante disciplinario especial — ajusta el juego según el rival, no según el árbitro."
                )

            if rl["Faltas_com"] > df["Faltas_com"].quantile(0.70):
                _info_box(
                    f"{local} es de los equipos que más faltas comete (**{rl['Faltas_com']:.1f} por partido**) "
                    f"y este árbitro pita **{inf['arb_amarillas']:.2f} amarillas por partido**.",
                    "El riesgo de inferioridad numérica es elevado. Inculca disciplina táctica antes del partido."
                )

            # ── Penaltis ───────────────────────────────────────────
            if inf["arb_penaltis"] is not None:
                if inf["arb_penaltis"] > inf["liga_penaltis"] * 1.3:
                    _info_box(
                        f"Este árbitro señala penaltis con frecuencia: **{inf['arb_penaltis']:.2f} por partido** "
                        f"(media Liga: {inf['liga_penaltis']:.2f}).",
                        "Cuidado con los contactos en el área propia. En ataque, el desborde interior puede ser rentable."
                    )
                elif inf["arb_penaltis"] < inf["liga_penaltis"] * 0.7:
                    _info_box(
                        f"Este árbitro es reacio a señalar penaltis: **{inf['arb_penaltis']:.2f} por partido** "
                        f"(media Liga: {inf['liga_penaltis']:.2f}).",
                        "No cuentes con el penalti como recurso ofensivo. El desborde en el área debe buscar el remate directo."
                    )

            # ── Expulsiones ────────────────────────────────────────
            if inf["arb_rojas"] is not None and inf["arb_rojas"] > inf["liga_rojas"] * 1.4:
                _info_box(
                    f"Este árbitro expulsa más que la media: **{inf['arb_rojas']:.2f} rojas por partido** "
                    f"(media Liga: {inf['liga_rojas']:.2f}).",
                    "Evita gestos y protestas. Una expulsión temprana puede condicionar todo el planteamiento táctico."
                )

            # ── Porcentaje de victorias según localía ──────────────
            if all(v is not None for v in [_pct_local, _pct_vis]):
                if es_local and _pct_local > inf["liga_pct_local"] + 8:
                    _info_box(
                        f"Este árbitro beneficia estadísticamente al equipo local: **{_pct_local:.0f}% de victorias locales** "
                        f"(media Liga: {inf['liga_pct_local']:.0f}%).",
                        "Jugar en casa es una ventaja adicional con este árbitro — aprovecha el factor campo desde el inicio."
                    )
                elif not es_local and _pct_vis > inf["liga_pct_vis"] + 8:
                    _info_box(
                        f"Este árbitro registra un alto porcentaje de victorias visitante: **{_pct_vis:.0f}%** "
                        f"(media Liga: {inf['liga_pct_vis']:.0f}%).",
                        "Jugar fuera no es una desventaja añadida con este árbitro — sal a proponer desde el minuto uno."
                    )
                elif not es_local and _pct_local > inf["liga_pct_local"] + 8:
                    _info_box(
                        f"Este árbitro registra un alto porcentaje de victorias local: **{_pct_local:.0f}%** "
                        f"(media Liga: {inf['liga_pct_local']:.0f}%). Juegas fuera.",
                        "Como visitante, asegura el resultado antes de asumir riesgos. El empate puede ser un buen punto de partida."
                    )

            st.markdown(
                '<div style="background:#161c2e;border-left:2px solid rgba(136,146,176,0.22);'
                'border-radius:0 6px 6px 0;padding:9px 14px;margin-top:4px;margin-bottom:12px">'
                '<span style="color:#8892b0;font-size:0.78rem;font-style:italic">'
                'Nota: el análisis estadístico del EDA (r no significativo, NS) no confirma una correlación '
                'entre el factor arbitral y el rendimiento de los equipos. Los datos anteriores son '
                'indicadores contextuales para la toma de decisiones estratégicas.'
                '</span></div>',
                unsafe_allow_html=True
            )

    with col_ctx2:
        # ── Condiciones meteorológicas ─────────────────────────────
        row_ll_local = df[df["Equipo"] == local].iloc[0]
        row_ll_rival = df[df["Equipo"] == rival].iloc[0]

        if lluvia_partido:
            st.markdown(f"""
            <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
            padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
            margin-bottom:16px">
                <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:14px;text-align:center">
                    Condiciones meteorológicas</div>
                <div style="color:#ccd6f6;font-size:0.9rem;margin-bottom:8px">
                    Se prevé lluvia en el partido.</div>
                <div style="color:#8892b0;font-size:0.82rem;line-height:1.7">
                    {local}: <span style="color:#ccd6f6">{row_ll_local['Total_mm']:.0f} mm acumulados — {row_ll_local['Dias_lluvia']:.0f} días de lluvia esta temporada</span><br>
                    {rival}: <span style="color:#ccd6f6">{row_ll_rival['Total_mm']:.0f} mm acumulados — {row_ll_rival['Dias_lluvia']:.0f} días de lluvia esta temporada</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            med_lluvia = df["Dias_lluvia"].median()
            local_hab  = row_ll_local["Dias_lluvia"] >= med_lluvia
            rival_hab  = row_ll_rival["Dias_lluvia"] >= med_lluvia

            if local_hab and not rival_hab:
                _info_box(
                    f"{local} está más habituado a jugar con lluvia "
                    f"(**{row_ll_local['Dias_lluvia']:.0f} días** vs {row_ll_rival['Dias_lluvia']:.0f} del rival).",
                    "El campo pesado puede beneficiar el juego físico y directo."
                )
            elif rival_hab and not local_hab:
                _info_box(
                    f"El rival tiene mayor experiencia en lluvia "
                    f"(**{row_ll_rival['Dias_lluvia']:.0f} días** vs {row_ll_local['Dias_lluvia']:.0f} propios).",
                    "Prioriza el juego directo y evita el combinativo en campo encharcado."
                )
            else:
                _info_box("Ambos equipos con experiencia similar en lluvia. Sin ventaja climatológica.")

            st.markdown(
                '<div style="background:#161c2e;border-left:2px solid rgba(136,146,176,0.22);'
                'border-radius:0 6px 6px 0;padding:9px 14px;margin-bottom:12px">'
                '<span style="color:#8892b0;font-size:0.78rem;font-style:italic">'
                'Nota: el análisis estadístico de temporada (p=0.349, NS) no confirma '
                'que la habituación a la lluvia prediga resultados. Es un factor contextual.'
                '</span></div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(f"""
            <div style="background:linear-gradient(145deg,#1a2035,#242c4a);border-radius:14px;
            padding:22px;border:1px solid #3d4870;box-shadow:0 4px 20px rgba(0,0,0,0.4);
            margin-bottom:16px">
                <div style="color:#ccd6f6;font-size:1.05rem;font-weight:700;margin-bottom:10px;text-align:center">
                    Condiciones meteorológicas</div>
                <div style="color:#8892b0;font-size:0.9rem">
                    Sin lluvia prevista. No hay condicionante climatológico.</div>
            </div>
            """, unsafe_allow_html=True)

    # ── Resumen ejecutivo ─────────────────────────────────────────
    st.markdown(
        '<div class="section-header" style="margin-top:8px">Resumen ejecutivo</div>',
        unsafe_allow_html=True
    )

    def _action_box(text):
        act = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
        st.markdown(
            f'<div style="background:#142032;border:1px solid rgba(74,158,202,0.55);'
            f'border-radius:10px;padding:14px 18px;font-size:0.97rem;'
            f'line-height:1.6;margin-bottom:10px">'
            f'<span style="color:#F5C842;font-weight:700;font-size:0.78rem;'
            f'letter-spacing:0.06em;text-transform:uppercase">▸ Acción&nbsp;&nbsp;</span>'
            f'<span style="color:#ffffff">{act}</span></div>',
            unsafe_allow_html=True
        )

    # ── Pool de acciones — ordenadas por importancia ─────────────
    _xg_conc_pp   = rr["xG_contra"] / rr["PJ"] if rr["PJ"] > 0 else 0
    _pos_map_re   = {r["Jugador"]: r["Posicion"] for r in lineup_local}
    _lineup_names = list(_pos_map_re.keys())
    _top_own = (
        inf["jug_local_disp"]
        [inf["jug_local_disp"]["Jugador"].isin(_lineup_names)]
        .sort_values("IS_indiv", ascending=False)
    )
    _fl = inf["forma_local"]
    _fr = inf["forma_rival"]

    # Cada entrada: (prioridad, etiqueta_agrupación, texto)
    _pool = []

    # P1 — Orientación táctica (siempre)
    if is_delta > 0.10:
        _pool.append((1, "tactica", "Sal a dominar el partido desde el inicio. Explota las debilidades del rival detectadas en el informe."))
    elif is_delta < -0.10:
        _pool.append((1, "tactica", "Defiende ordenado y sé paciente. Cada ocasión cuenta — no la desperdicies."))
    else:
        _pool.append((1, "tactica", "Partido igualado — los detalles decidirán. Mantén la concentración los 90 minutos."))

    # P2 — Dupla peligrosa rival (si existe)
    if not inf["dup_rival"].empty:
        _td = inf["dup_rival"].iloc[0]
        _pool.append((2, "dupla", f"Asigna marcaje específico a **{_td['Goleador']}** y **{_td['Asistidor']}** — córtales el circuito antes de que conecten."))

    # P3 — Bajas del rival (si las hay)
    if inf["ausentes_rival"]:
        _aus    = ", ".join(f"**{j}**" for j in inf["ausentes_rival"][:2])
        _sufijo = " y otros" if len(inf["ausentes_rival"]) > 2 else ""
        _pool.append((3, "bajas", f"El rival llega con bajas: {_aus}{_sufijo}. Ataca sus huecos desde el inicio."))

    # P4 — Portero rival (solo si es extremo)
    if inf["por_ga90_rival"] is not None:
        if inf["por_ga90_rival"] > 1.5:
            _pool.append((4, "portero", f"El portero de {rival} es vulnerable — dispara con insistencia, no te compliques."))
        elif inf["por_ga90_rival"] < 0.9:
            _pool.append((4, "portero", f"El portero de {rival} es sólido — busca solo ocasiones claras, no dispares por volumen."))

    # P5 — Tu jugador más decisivo
    if not _top_own.empty:
        _best_own = _top_own.iloc[0]
        _pos_own  = _pos_map_re.get(_best_own["Jugador"], "")
        _pool.append((5, "jug_propio", f"Activa a **{_best_own['Jugador']}** ({_pos_own}) desde el inicio — es tu jugador más decisivo."))

    # P6 — Neutralizar al jugador más influyente del rival
    if not inf["rival_lineup"].empty:
        _rivals_s = inf["rival_lineup"].sort_values("IS_indiv", ascending=False)
        if not inf["dup_rival"].empty:
            _ya = {inf["dup_rival"].iloc[0]["Goleador"], inf["dup_rival"].iloc[0]["Asistidor"]}
            _rivals_s = _rivals_s[~_rivals_s["Jugador"].isin(_ya)]
        if not _rivals_s.empty:
            _top_r = _rivals_s.iloc[0]
            _pool.append((6, "jug_rival", f"Neutraliza a **{_top_r['Jugador']}** ({_top_r['Posicion']}) — no le dejes girar ni recibir en espacios."))

    # P7 — Bloque defensivo (rival muy peligroso ofensivamente)
    if rr["xG_pp"] > df["xG_pp"].quantile(0.75):
        _pool.append((7, "bloque", "El rival es muy peligroso ofensivamente — bloque compacto y sin espacios en la espalda de la defensa."))

    # P8 — Eficiencia de tiro / superioridad ofensiva
    if rl["Pct_tiro_area"] < 45:
        _pool.append((8, "tiro", "Trabaja el balón hasta dentro del área antes de rematar — evita los disparos lejanos."))
    elif rl["xG_pp"] > _xg_conc_pp * 1.15:
        _pool.append((8, "tiro", "Tienes superioridad ofensiva clara — mantén el ritmo y remata desde dentro del área."))

    # P9 — Árbitro (solo si está claramente fuera de la media)
    if inf["arb_faltas"] is not None:
        if inf["arb_faltas"] > inf["liga_faltas"] + 1:
            _pool.append((9, "arbitro", "Árbitro estricto — evita las entradas individuales y cuida las amarillas."))
        elif inf["arb_faltas"] < inf["liga_faltas"] - 1:
            _pool.append((9, "arbitro", "Árbitro permisivo — puedes apretar la intensidad defensiva sin miedo."))

    # P10 — Lluvia
    if lluvia_partido:
        _pool.append((10, "lluvia", "Lluvia prevista — juega directo y evita el combinativo en campo pesado."))

    # P11 — Forma reciente (solo si hay diferencia clara)
    if _fl and _fr:
        if _fl.get("pts", 0) > _fr.get("pts", 0) + 0.5:
            _pool.append((11, "forma", "Llegas en mejor forma que el rival — sal con confianza y propón el juego desde el inicio."))
        elif _fr.get("pts", 0) > _fl.get("pts", 0) + 0.5:
            _pool.append((11, "forma", "El rival llega en mejor forma — atención máxima al arranque del partido."))

    # Ordenar por prioridad y construir tarjetas (máx 2 acciones por tarjeta)
    # Pares naturales: (jug_propio + jug_rival), (arbitro + lluvia)
    _pool.sort(key=lambda x: x[0])
    _pair_rules = [{"jug_propio", "jug_rival"}, {"arbitro", "lluvia"}]

    _cards = []
    _used  = set()
    for i, (_, tag_i, txt_i) in enumerate(_pool):
        if i in _used:
            continue
        _partner = None
        for rule in _pair_rules:
            if tag_i in rule:
                other_tags = rule - {tag_i}
                for j, (_, tag_j, txt_j) in enumerate(_pool):
                    if j > i and j not in _used and tag_j in other_tags:
                        _partner = (j, txt_j)
                        break
            if _partner:
                break
        if _partner:
            _cards.append(f"{txt_i} {_partner[1]}")
            _used.add(i)
            _used.add(_partner[0])
        else:
            _cards.append(txt_i)
            _used.add(i)
        if len(_cards) >= 6:
            break

    for _c in _cards:
        _action_box(_c)

    st.markdown(
        f'<div style="color:#8892b0;font-size:0.78rem;font-style:italic;margin-top:8px">'
        f'Informe generado · {local} vs {rival} · '
        f'Árbitro: {arbitro_sel} · Formación: {formacion} · '
        f'Algoritmo Prescriptivo LaLiga 2025-26</div>',
        unsafe_allow_html=True
    )

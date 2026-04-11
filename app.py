"""
app.py — Algoritmo Prescriptivo Táctico · LaLiga 2025-26
=========================================================
Aplicación Streamlit que, dados los parámetros de un partido próximo
(equipos, árbitro, lluvia, alineación propia), genera un informe táctico
prescriptivo en cinco bloques:

  Bloque 1 · Resumen General      → comparativa de IS y métricas clave
  Bloque 2 · Estrategia Ofensiva  → cómo atacar al rival
  Bloque 3 · Estrategia Defensiva → cómo neutralizar al rival
  Bloque 4 · Jugadores Clave      → alineación estimada del rival +
                                    jugadores propios más influyentes
  Bloque 5 · Contexto             → árbitro y condiciones meteorológicas

El Índice de Éxito (IS) calibra la agresividad de las recomendaciones:
  IS = 0.35·norm(xG/PJ) + 0.35·norm(−GC/PJ) + 0.30·norm(Rating)

Fuente de datos: Base de Datos.xlsx  (hoja Equipos, Clasificación, Lluvias,
  Árbitros, Jugadores, Porteros, Lesionados y Sancionados, Duplas Peligrosas)

Uso: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np

# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════
EXCEL = "/Users/NacoLG/Documentos/UFV 4/TFG/Ingeniería del Dato/Base de Datos.xlsx"

FORMACIONES = ["4-3-3", "4-4-2", "4-2-3-1", "4-1-4-1", "3-4-3",
               "3-5-2", "5-3-2", "5-4-1", "4-5-1", "4-3-2-1"]

# Slots por formación: (etiqueta, grupo_filtro)
# grupo_filtro: "portero" | "defensa" | "centro" | "extremo" | "delantero"
# Los slots "centro" mostrarán un mini-selector de rol (Pivote / MC / Mediapunta)
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
        ("Mediapunta",    "centro"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
    ],
    "4-1-4-1": [
        ("Portero",           "portero"),
        ("Lateral Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Lateral Izquierdo", "defensa"),
        ("Pivote",    "centro"),
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
        ("Portero",           "portero"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Carrilero Derecho",    "defensa"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Carrilero Izquierdo",    "defensa"),
        ("Delantero Centro",  "delantero"),
        ("Delantero Centro",  "delantero"),
    ],
    "5-3-2": [
        ("Portero",           "portero"),
        ("Carrilero Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Carrilero Izquierdo", "defensa"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Delantero Centro",  "delantero"),
        ("Delantero Centro",  "delantero"),
    ],
    "5-4-1": [
        ("Portero",           "portero"),
        ("Carrilero Derecho",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Defensa Central",   "defensa"),
        ("Carrilero Izquierdo", "defensa"),
        ("Extremo Derecho",   "extremo"),
        ("Centrocampista",    "centro"),
        ("Centrocampista",    "centro"),
        ("Extremo Izquierdo", "extremo"),
        ("Delantero Centro",  "delantero"),
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
        ("Mediapunta",    "centro"),
        ("Mediapunta",    "centro"),
        ("Delantero Centro",  "delantero"),
    ],
}


# ── Normalización de nombres de equipo ────────────────────────────────────────
# WhoScored usa nombres distintos a los del resto de fuentes.
# Este diccionario traduce los nombres problemáticos al formato estándar.
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
    """Aplica TEAM_ALIASES para homogeneizar el nombre de un equipo entre hojas."""
    if pd.isna(name):
        return ""
    s = str(name).strip()
    return TEAM_ALIASES.get(s, s)


# ══════════════════════════════════════════════════════════════════
# CARGA DE DATOS (cacheada)
# ══════════════════════════════════════════════════════════════════
def pct_num(series):
    """Convierte una serie con valores tipo '58%' a números (58.0)."""
    return (series.astype(str)
            .str.replace("%", "", regex=False)
            .pipe(pd.to_numeric, errors="coerce"))

def minmax_series(s):
    """Normalización Min-Max de una serie. Si max==min, devuelve 0.5 para todos."""
    mn, mx = s.min(), s.max()
    if mx > mn:
        return (s - mn) / (mx - mn)
    return pd.Series(0.5, index=s.index)

def minmax_group(series, groups):
    """
    Normalización Min-Max dentro de cada grupo posicional.

    Normaliza por separado para Porteros, Defensas, Centrocampistas y Delanteros,
    evitando que un portero (con pocas acciones ofensivas) compita directamente
    en la misma escala que un delantero.
    """
    result = pd.Series(0.0, index=series.index)
    for g in groups.unique():
        mask = groups == g
        s = series[mask]
        mn, mx = s.min(), s.max()
        result[mask] = (s - mn) / (mx - mn) if mx > mn else 0.5
    return result


@st.cache_data(show_spinner="Cargando base de datos…")
def load_data():
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
        "PJ":     pd.to_numeric(raw_cla.iloc[3:23, 2], errors="coerce"),
        "GF":     pd.to_numeric(raw_cla.iloc[3:23, 6], errors="coerce"),
        "GC":     pd.to_numeric(raw_cla.iloc[3:23, 7], errors="coerce"),
        "Pts":    pd.to_numeric(raw_cla.iloc[3:23, 9], errors="coerce"),
    }).reset_index(drop=True)

    # ── Lluvias ───────────────────────────────────────────────────
    raw_ll  = pd.read_excel(EXCEL, sheet_name="Lluvias", header=None)
    lluvia = pd.DataFrame({
        "Equipo":      raw_ll.iloc[2:22, 0].apply(normalize_team).values,
        "Total_mm":    pd.to_numeric(raw_ll.iloc[2:22, 3], errors="coerce"),
        "Dias_lluvia": pd.to_numeric(raw_ll.iloc[2:22, 4], errors="coerce"),
    }).reset_index(drop=True)

    # ── Merge maestro ─────────────────────────────────────────────
    df = eq.merge(cla, on="Equipo", how="inner").merge(lluvia, on="Equipo", how="inner")
    df = df.loc[:, ~df.columns.duplicated()].copy()
    df["Pts_pp"] = df["Pts"] / df["PJ"]
    df["xG_pp"]  = df["xG"]  / df["PJ"]
    df["GC_pp"]  = df["GC"]  / df["PJ"]
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
        "Partidos":    pd.to_numeric(raw_arb.iloc[2:, 2], errors="coerce"),
        "FaltasPP":    pd.to_numeric(raw_arb.iloc[2:, 4], errors="coerce"),
        "AmarillasPP": pd.to_numeric(raw_arb.iloc[2:, 6], errors="coerce"),
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
        "Minutos":  pd.to_numeric(raw_por.iloc[3:, 20], errors="coerce"),
        "Rating":   pd.to_numeric(raw_por.iloc[3:,  5], errors="coerce"),
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
        "Minutos":   pd.to_numeric(raw_jug.iloc[3:, 20], errors="coerce"),
        "PctMin":    pd.to_numeric(raw_jug.iloc[3:, 24], errors="coerce"),
        "Rating":    pd.to_numeric(raw_jug.iloc[3:,  5], errors="coerce"),
        "xG_90":     pd.to_numeric(raw_jug.iloc[3:, 139], errors="coerce"),
        "Asist_90":  pd.to_numeric(raw_jug.iloc[3:,  43], errors="coerce"),
        "Entradas":  pd.to_numeric(raw_jug.iloc[3:, 235], errors="coerce"),
        "Interc":    pd.to_numeric(raw_jug.iloc[3:, 241], errors="coerce"),
        "Despejes":  pd.to_numeric(raw_jug.iloc[3:, 244], errors="coerce"),
        "Bloqueos":  pd.to_numeric(raw_jug.iloc[3:, 247], errors="coerce"),
    }).dropna(subset=["Minutos", "Equipo"])
    jug = jug[jug["Minutos"] >= 400].reset_index(drop=True)

    # ── IS individual con normalización dentro del grupo posicional ───────────
    # Cada jugador se compara solo con los de su misma posición para que
    # un portero no quede penalizado frente a un delantero en métricas ofensivas.
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

    # ── Añadir porteros al dataframe de jugadores ─────────────────
    rat_por = por["Rating"].fillna(por["Rating"].median())
    por["norm_rat"]  = minmax_group(rat_por, por["PosGrupo"])
    por["IS_indiv"]  = 0.30 * por["norm_rat"]   # solo rating para porteros
    min_eq_por = jug.groupby("Equipo")["Minutos"].sum().rename("Min_equipo")
    por = por.merge(min_eq_por, on="Equipo", how="left")
    por["IS_contrib"] = por["IS_indiv"] * (por["Minutos"] / por["Min_equipo"].fillna(1))
    # Columnas que jug tiene y por no (rellenar con NaN)
    for col in ["PctMin", "xG_90", "Asist_90", "Entradas", "Interc", "Despejes", "Bloqueos"]:
        por[col] = np.nan
    jug = pd.concat([jug, por], ignore_index=True)

    # ── Lesionados y Sancionados ──────────────────────────────────
    raw_les = pd.read_excel(EXCEL, sheet_name="Lesionados y Sancionados", header=None)
    les = pd.DataFrame({
        "Jugador": raw_les.iloc[2:, 0].values,
        "Equipo":  raw_les.iloc[2:, 1].apply(normalize_team).values,
        "Motivo":  raw_les.iloc[2:, 2].values,
        "Vuelta":  raw_les.iloc[2:, 5].values,
        "ProbJugar": pd.to_numeric(raw_les.iloc[2:, 6], errors="coerce"),
    }).dropna(subset=["Jugador"]).reset_index(drop=True)
    les["Jugador"] = les["Jugador"].astype(str).str.strip()
    les["Equipo"]  = les["Equipo"].astype(str).str.strip()
    # 0 % = confirmado ausente
    ausentes = les[les["ProbJugar"] == 0]["Jugador"].tolist()

    # ── Duplas Peligrosas ─────────────────────────────────────────
    raw_dup = pd.read_excel(EXCEL, sheet_name="Duplas Peligrosas", header=None)
    dup = pd.DataFrame({
        "Goleador":   raw_dup.iloc[1:, 1].values,
        "Asistidor":  raw_dup.iloc[1:, 2].values,
        "Equipo":     raw_dup.iloc[1:, 3].apply(normalize_team).values,
        "Frecuencia": pd.to_numeric(raw_dup.iloc[1:, 4], errors="coerce"),
    }).dropna(subset=["Frecuencia"]).reset_index(drop=True)

    return df, teams, arb, arbitros, jug, les, ausentes, dup


# ══════════════════════════════════════════════════════════════════
# LÓGICA PRESCRIPTIVA
# ══════════════════════════════════════════════════════════════════

def percentil(val, serie):
    """Calcula el percentil de `val` dentro de `serie` (resultado entre 0 y 100)."""
    """Percentil de `val` dentro de `serie` (0–100)."""
    return round(float(np.mean(serie <= val) * 100))

def diferencia_etiqueta(delta, umbral_alto=0.15, umbral_bajo=-0.15):
    """
    Clasifica la diferencia de IS entre local y rival en tres categorías.

    Devuelve una tupla (etiqueta, color_streamlit) donde:
      - FAVORABLE:    ΔIS > umbral_alto   → verde
      - DESFAVORABLE: ΔIS < umbral_bajo   → rojo
      - EQUILIBRADO:  en el rango medio   → naranja
    """
    if delta > umbral_alto:
        return "FAVORABLE", "success"
    if delta < umbral_bajo:
        return "DESFAVORABLE", "error"
    return "EQUILIBRADO", "warning"

def jugadores_para_slot(jug_df, grupo, ausentes):
    """
    Devuelve la lista de jugadores disponibles para un slot de la alineación.

    Filtra por grupo posicional y excluye los jugadores confirmados como bajas.
    Los jugadores con varias posiciones aparecen en el selector de cada
    grupo que les corresponda (ej. un jugador con 'Defensa, Centrocampista'
    aparece en ambos desplegables).

    Args:
        jug_df:   DataFrame de jugadores del equipo local.
        grupo:    Clave de posición ('portero', 'defensa', 'centro', etc.).
        ausentes: Lista de nombres de jugadores descartados (0% probabilidad).
    """
    """
    Devuelve nombres disponibles filtrados por grupo posicional.
    Si un jugador tiene varias posiciones (ej. 'Defensa, Centrocampista'),
    aparece en el desplegable de CADA grupo que le corresponda.
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
        # Normalizar y comparar cada posición listada
        partes = [p.strip().capitalize() for p in str(pos_str).split(",")]
        return bool(grupos_buscados & set(partes))

    candidatos = jug_df[jug_df["Posicion"].apply(cumple)]
    candidatos = candidatos[~candidatos["Jugador"].isin(ausentes)]
    nombres = candidatos.sort_values("Minutos", ascending=False)["Jugador"].tolist()
    return nombres if nombres else ["(sin datos)"]


def ranking_is(df, equipo):
    """Devuelve la posición del equipo en el ranking IS (1 = mejor IS)."""
    orden = df.sort_values("IS", ascending=False)["Equipo"].tolist()
    return orden.index(equipo) + 1 if equipo in orden else "?"

def generar_informe(local, rival, arbitro_sel, lluvia_partido,
                    lineup_local, df, arb, jug, les, ausentes, dup,
                    es_local=True):
    """
    Devuelve un dict con los 5 bloques del informe prescriptivo.
    """
    row_local = df[df["Equipo"] == local].iloc[0]
    row_rival = df[df["Equipo"] == rival].iloc[0]

    # Ausentes CONFIRMADOS de ambos equipos (ProbJugar == 0)
    ausentes_local = les[(les["Equipo"] == local) & (les["ProbJugar"] == 0)]["Jugador"].tolist()
    ausentes_rival = les[(les["Equipo"] == rival) & (les["ProbJugar"] == 0)]["Jugador"].tolist()

    # Jugadores disponibles del rival (para estimar alineación)
    jug_rival = jug[jug["Equipo"] == rival].copy()
    jug_rival_disp = jug_rival[~jug_rival["Jugador"].isin(ausentes_rival)]
    rival_lineup = (jug_rival_disp
                    .sort_values("Minutos", ascending=False)
                    .head(11)[["Jugador", "Posicion", "Minutos", "IS_indiv"]]
                    .reset_index(drop=True))

    # Jugadores disponibles del equipo local
    jug_local = jug[jug["Equipo"] == local].copy()
    jug_local_disp = jug_local[~jug_local["Jugador"].isin(ausentes_local)]

    # IS delta
    is_delta = float(row_local["IS"]) - float(row_rival["IS"])
    estado_label, estado_color = diferencia_etiqueta(is_delta)

    # Árbitro
    arb_row = arb[arb["Arbitro"] == arbitro_sel].head(1)
    arb_faltas = float(arb_row["FaltasPP"].values[0]) if len(arb_row) else None
    arb_amarillas = float(arb_row["AmarillasPP"].values[0]) if len(arb_row) else None
    liga_faltas   = float(arb["FaltasPP"].mean())
    liga_amarillas = float(arb["AmarillasPP"].mean())

    # Duplas del rival
    dup_rival = (dup[dup["Equipo"] == rival]
                 .sort_values("Frecuencia", ascending=False)
                 .head(3))

    return {
        "row_local": row_local,
        "row_rival": row_rival,
        "is_delta": is_delta,
        "estado_label": estado_label,
        "estado_color": estado_color,
        "rival_lineup": rival_lineup,
        "jug_local_disp": jug_local_disp,
        "arb_faltas": arb_faltas,
        "arb_amarillas": arb_amarillas,
        "liga_faltas": liga_faltas,
        "liga_amarillas": liga_amarillas,
        "ausentes_local": ausentes_local,
        "ausentes_rival": ausentes_rival,
        "dup_rival": dup_rival,
        "lluvia": lluvia_partido,
        "es_local": es_local,
    }


# ══════════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Algoritmo Prescriptivo LaLiga",
    page_icon="⚽",
    layout="wide",
)

# CSS suave
st.markdown("""
<style>
.metric-card {
    background: #f8f9fa;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 12px;
    border-left: 4px solid #1a73e8;
}
.section-header {
    font-size: 1.3rem;
    font-weight: 700;
    color: #1a1a2e;
    border-bottom: 2px solid #1a73e8;
    padding-bottom: 6px;
    margin-top: 24px;
    margin-bottom: 12px;
}
.tag-success { color: #27ae60; font-weight: bold; }
.tag-warning { color: #e67e22; font-weight: bold; }
.tag-error   { color: #e74c3c; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("⚽ Algoritmo Prescriptivo Táctico — LaLiga 2025-26")
st.caption("Herramienta de apoyo a la decisión para cuerpos técnicos")

# ── Carga de datos ────────────────────────────────────────────────
df, teams, arb, arbitros, jug, les, ausentes, dup = load_data()

# ── SIDEBAR: parámetros del partido ───────────────────────────────
with st.sidebar:
    st.header("🗓️ Parámetros del partido")

    mi_equipo = st.selectbox("Tu equipo", teams, index=0)
    rival_options = [t for t in teams if t != mi_equipo]
    rival = st.selectbox("Rival", rival_options, index=0)

    es_local = st.radio(
        "¿Dónde juegas?",
        ["🏠 En casa (local)", "✈️ Fuera (visitante)"],
        index=0,
        horizontal=True,
    ) == "🏠 En casa (local)"

    # Asignar local/visitante según selección
    local = mi_equipo

    arbitro_sel = st.selectbox("Árbitro designado", ["(Desconocido)"] + arbitros)

    lluvia_partido = st.toggle("¿Se prevé lluvia?", value=False)

    st.divider()
    st.header("📋 Tu alineación")

    formacion = st.selectbox("Formación", FORMACIONES, index=0)

    # Bajas confirmadas del equipo local
    ausentes_local_confirmed = les[
        (les["Equipo"] == local) & (les["ProbJugar"] == 0)
    ]["Jugador"].tolist()
    jug_equipo = jug[jug["Equipo"] == local].copy()

    if ausentes_local_confirmed:
        st.caption(f"⚠️ Excluidos (0% prob.): {', '.join(ausentes_local_confirmed)}")

    slots = FORMACIONES_SLOTS.get(formacion, [])

    # Asignación greedy: cada slot recibe el jugador con más minutos
    # de su grupo que aún no haya sido asignado a un slot previo.
    asignados = set()
    defaults = []
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
    for i, ((etiqueta, grupo), (candidatos, idx_default)) in enumerate(
        zip(slots, defaults)
    ):
        col_lbl, col_jug = st.columns([2, 3])
        with col_lbl:
            st.markdown(f"<small>{etiqueta}</small>", unsafe_allow_html=True)
        with col_jug:
            jug_sel = st.selectbox(
                etiqueta,
                candidatos,
                index=idx_default,
                key=f"jug_{i}",
                label_visibility="collapsed",
            )
        lineup_local.append({"Posicion": etiqueta, "Jugador": jug_sel})

    generar = st.button("🔍 Generar informe táctico", type="primary", use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# PANTALLA PRINCIPAL
# ══════════════════════════════════════════════════════════════════

if not generar:
    st.info("Configura el partido en el panel izquierdo y pulsa **Generar informe táctico**.")

    # Muestra estado general de todos los equipos
    st.markdown("### Ranking IS — LaLiga 2025-26")
    tabla_is = (df[["Equipo","IS","Pts","xG_pp","GC_pp","Rating"]]
                .sort_values("IS", ascending=False)
                .reset_index(drop=True))
    tabla_is.index = tabla_is.index + 1
    tabla_is.columns = ["Equipo","IS","Pts","xG/PJ","GC/PJ","Rating"]
    tabla_is["IS"]     = tabla_is["IS"].round(3)
    tabla_is["xG/PJ"]  = tabla_is["xG/PJ"].round(2)
    tabla_is["GC/PJ"]  = tabla_is["GC/PJ"].round(2)
    tabla_is["Rating"] = tabla_is["Rating"].round(2)
    st.dataframe(tabla_is, use_container_width=True, height=600)

else:
    # ── Generar informe ───────────────────────────────────────────
    inf = generar_informe(
        local, rival, arbitro_sel, lluvia_partido,
        lineup_local, df, arb, jug, les, ausentes, dup,
        es_local=es_local,
    )

    rl  = inf["row_local"]
    rr  = inf["row_rival"]

    # ════════════════════════════════════════
    # CABECERA DEL PARTIDO
    # ════════════════════════════════════════
    # Iconos según quién es local
    icono_mi   = "🏠" if es_local else "✈️"
    icono_rival = "✈️" if es_local else "🏠"
    rol_mi    = "local"     if es_local else "visitante"
    rol_rival = "visitante" if es_local else "local"

    c1, c2, c3 = st.columns([3, 1, 3])
    with c1:
        st.markdown(f"### {icono_mi} {local}")
        st.markdown(f"IS: **{rl['IS']:.3f}** · Rating: **{rl['Rating']:.2f}**")
        st.markdown(f"#{ranking_is(df, local)} en la clasificación IS")
    with c2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        estado_html = {
            "FAVORABLE":    '<span class="tag-success">FAVORABLE</span>',
            "EQUILIBRADO":  '<span class="tag-warning">EQUILIBRADO</span>',
            "DESFAVORABLE": '<span class="tag-error">DESFAVORABLE</span>',
        }[inf["estado_label"]]
        st.markdown(f"<div style='text-align:center'>{estado_html}</div>", unsafe_allow_html=True)
    with c3:
        st.markdown(f"### {icono_rival} {rival}")
        st.markdown(f"IS: **{rr['IS']:.3f}** · Rating: **{rr['Rating']:.2f}**")
        st.markdown(f"#{ranking_is(df, rival)} en la clasificación IS")

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

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("IS (local)",   f"{rl['IS']:.3f}",
                 delta=f"{is_delta:+.3f} vs rival")
    col_b.metric("Pts/partido",  f"{rl['Pts_pp']:.2f}",
                 delta=f"{rl['Pts_pp']-rr['Pts_pp']:+.2f}")
    col_c.metric("xG/partido",   f"{rl['xG_pp']:.2f}",
                 delta=f"{rl['xG_pp']-rr['xG_pp']:+.2f}")
    col_d.metric("GC/partido",   f"{rl['GC_pp']:.2f}",
                 delta=f"{rl['GC_pp']-rr['GC_pp']:+.2f}",
                 delta_color="inverse")

    # Evaluación narrativa
    if is_delta > 0.20:
        valoracion = (
            f"**{local}** afronta este partido con una ventaja objetiva clara sobre **{rival}** "
            f"(ΔIS = {is_delta:+.3f}). La estrategia debe orientarse a **consolidar el dominio** "
            f"y explotar las debilidades del rival sin asumir riesgos innecesarios."
        )
    elif is_delta > 0.05:
        valoracion = (
            f"El enfrentamiento es **ligeramente favorable** para {local} (ΔIS = {is_delta:+.3f}). "
            f"El margen es pequeño: la eficiencia en las transiciones y la solidez defensiva "
            f"serán determinantes."
        )
    elif is_delta > -0.05:
        valoracion = (
            f"Partido **muy equilibrado** (ΔIS = {is_delta:+.3f}). "
            f"El resultado dependerá de los detalles tácticos y la gestión emocional. "
            f"Reducir errores propios y maximizar cada ocasión generada son prioritarios."
        )
    elif is_delta > -0.20:
        valoracion = (
            f"**{rival}** parte con ventaja objetiva (ΔIS = {is_delta:+.3f}). "
            f"Se recomienda una propuesta defensivamente sólida, con apuesta por el contraataque "
            f"rápido y la explotación de los puntos débiles del rival identificados a continuación."
        )
    else:
        valoracion = (
            f"**{local}** enfrenta a un rival notablemente superior (ΔIS = {is_delta:+.3f}). "
            f"La clave estará en la **organización defensiva compacta**, minimizar los xG encajados "
            f"y aprovechar cualquier situación a balón parado."
        )

    st.markdown(f"> {valoracion}")

    # Comparativa tabla
    compare_data = {
        "Métrica": ["IS", "Pts/PJ", "xG/PJ", "GC/PJ", "Rating", "Posesión %",
                    "Precisión pase %", "Tiros/PJ", "Tiros contra/PJ", "Faltas cometidas"],
        local:  [f"{rl['IS']:.3f}", f"{rl['Pts_pp']:.2f}", f"{rl['xG_pp']:.2f}",
                 f"{rl['GC_pp']:.2f}", f"{rl['Rating']:.2f}", f"{rl['Pos%']:.1f}%",
                 f"{rl['Precision_pase']:.1f}%", f"{rl['Tiros_pp']:.1f}",
                 f"{rl['Tiros_contra_pp']:.1f}", f"{rl['Faltas_com']:.1f}"],
        rival:  [f"{rr['IS']:.3f}", f"{rr['Pts_pp']:.2f}", f"{rr['xG_pp']:.2f}",
                 f"{rr['GC_pp']:.2f}", f"{rr['Rating']:.2f}", f"{rr['Pos%']:.1f}%",
                 f"{rr['Precision_pase']:.1f}%", f"{rr['Tiros_pp']:.1f}",
                 f"{rr['Tiros_contra_pp']:.1f}", f"{rr['Faltas_com']:.1f}"],
    }
    st.dataframe(pd.DataFrame(compare_data).set_index("Métrica"),
                 use_container_width=True)

    # ════════════════════════════════════════
    # BLOQUE 2 — ATAQUE
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">⚔️ Bloque 2 · Estrategia Ofensiva</div>',
                unsafe_allow_html=True)

    col_at1, col_at2 = st.columns(2)

    with col_at1:
        st.markdown(f"**Potencial ofensivo de {local}**")
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("xG/partido", f"{rl['xG_pp']:.2f}")
        col_m2.metric("Tiros a puerta/PJ", f"{rl['Tiros_puerta_pp']:.1f}")
        col_m3.metric("% tiros desde área", f"{rl['Pct_tiro_area']:.0f}%")

    with col_at2:
        st.markdown(f"**Vulnerabilidad defensiva de {rival}**")
        col_m4, col_m5, col_m6 = st.columns(3)
        col_m4.metric("xG concedido/PJ", f"{rr['xG_contra']/rr['PJ']:.2f}" if rr['PJ'] > 0 else "–")
        col_m5.metric("Tiros concedidos/PJ", f"{rr['Tiros_contra_pp']:.1f}")
        col_m6.metric("% paradas portero", f"{rr['SavePct']:.0f}%")

    # Recomendaciones ofensivas
    recs_ataque = []

    # xG propio vs rival
    if rl["xG_pp"] > rr["xG_contra"] / rr["PJ"] * 1.15:
        recs_ataque.append(
            f"🟢 **Superioridad ofensiva clara**: generas {rl['xG_pp']:.2f} xG/PJ frente a los "
            f"{rr['xG_contra']/rr['PJ']:.2f} xG/PJ que concede {rival}. "
            f"Mantén el volumen de llegada y prioriza tiros desde el área ({rl['Pct_tiro_area']:.0f}% actual)."
        )
    elif rl["xG_pp"] < rr["xG_contra"] / rr["PJ"] * 0.85:
        recs_ataque.append(
            f"🔴 **Rival defensivamente sólido**: solo concede {rr['xG_contra']/rr['PJ']:.2f} xG/PJ. "
            f"Prioriza la calidad del disparo (xG/tiro = {rl['xG_tiro']:.3f}) sobre el volumen. "
            f"Busca balones en profundidad para evitar su bloque defensivo."
        )
    else:
        recs_ataque.append(
            f"🟡 **Equilibrio ofensivo-defensivo**: el rival concede {rr['xG_contra']/rr['PJ']:.2f} xG/PJ, "
            f"cercano a lo que generas ({rl['xG_pp']:.2f}). Las transiciones rápidas serán clave."
        )

    # Tiros desde área
    if rl["Pct_tiro_area"] < 45:
        recs_ataque.append(
            f"⚠️ Solo el {rl['Pct_tiro_area']:.0f}% de tus disparos parten desde dentro del área "
            f"(media Liga ≈50%). Trabaja combinaciones interiores para mejorar la calidad de disparo."
        )

    # Pases clave cortos
    if rl["PasesClave_cort"] and rr["Intercepciones"]:
        if rr["Intercepciones"] > df["Intercepciones"].quantile(0.70):
            recs_ataque.append(
                f"⚠️ {rival} es uno de los equipos con mayor número de intercepciones "
                f"({rr['Intercepciones']:.0f}). Varía el juego con pases en profundidad "
                f"y evita el exceso de combinación corta en zonas de presión."
            )

    # Regates
    if rl["Reg_exit"] and rr["Entradas_fall"]:
        if rl["Reg_exit"] > df["Reg_exit"].quantile(0.65):
            recs_ataque.append(
                f"✅ Capacidad de regate por encima de la media ({rl['Reg_exit']:.1f}/PJ). "
                f"Utiliza los perfiles rápidos en banda para desequilibrar 1vs1."
            )

    # Posesión vs zona ofensiva (hallazgo contraintuitivo G1)
    if rl["Pos%"] > 55:
        recs_ataque.append(
            f"💡 **Dato contraintuitivo**: posesión elevada ({rl['Pos%']:.1f}%) NO garantiza "
            f"mayor presencia en zona ofensiva ni más goles (r=−0.035 NS en la Liga). "
            f"Cuida la verticalidad y la transición: la posesión debe ser productiva."
        )

    for rec in recs_ataque:
        st.markdown(rec)

    # Duplas peligrosas del rival a vigilar
    if not inf["dup_rival"].empty:
        st.markdown(f"**Duplas peligrosas de {rival} a neutralizar:**")
        for _, row_dup in inf["dup_rival"].iterrows():
            st.markdown(
                f"- **{row_dup['Goleador']}** (rematador) ← **{row_dup['Asistidor']}** "
                f"(asistidor) · {int(row_dup['Frecuencia'])} combinaciones"
            )

    # ════════════════════════════════════════
    # BLOQUE 3 — DEFENSA
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">🛡️ Bloque 3 · Estrategia Defensiva</div>',
                unsafe_allow_html=True)

    col_def1, col_def2 = st.columns(2)

    with col_def1:
        st.markdown(f"**Amenaza ofensiva de {rival}**")
        col_d1, col_d2, col_d3 = st.columns(3)
        col_d1.metric("xG rival/PJ",  f"{rr['xG_pp']:.2f}")
        col_d2.metric("Tiros rival/PJ", f"{rr['Tiros_pp']:.1f}")
        col_d3.metric("% tiros área rival", f"{rr['Pct_tiro_area']:.0f}%")

    with col_def2:
        st.markdown(f"**Solidez defensiva de {local}**")
        col_d4, col_d5, col_d6 = st.columns(3)
        col_d4.metric("GC/PJ", f"{rl['GC_pp']:.2f}")
        col_d5.metric("Portería a cero", f"{rl['P0']:.0f}")
        col_d6.metric("% paradas portero", f"{rl['SavePct']:.0f}%")

    recs_defensa = []

    # Presión ofensiva del rival
    if rr["xG_pp"] > df["xG_pp"].quantile(0.75):
        recs_defensa.append(
            f"🔴 **Rival muy peligroso ofensivamente** ({rr['xG_pp']:.2f} xG/PJ, top 25% Liga). "
            f"Replantear la línea defensiva: considera un bloque medio bajo y limitar "
            f"los espacios a la espalda de la defensa."
        )
    elif rr["xG_pp"] < df["xG_pp"].quantile(0.30):
        recs_defensa.append(
            f"🟢 **Rival con ataque limitado** ({rr['xG_pp']:.2f} xG/PJ). "
            f"La defensa puede permitirse una línea más adelantada para generar superioridad "
            f"en la salida de balón."
        )
    else:
        recs_defensa.append(
            f"🟡 **Rival con ataque moderado** ({rr['xG_pp']:.2f} xG/PJ). "
            f"Mantén el bloque medio habitual con atención especial en transiciones."
        )

    # Tiros desde área del rival
    if rr["Pct_tiro_area"] > 55:
        recs_defensa.append(
            f"⚠️ {rival} genera el {rr['Pct_tiro_area']:.0f}% de sus tiros desde dentro del área. "
            f"Fundamental el cierre de los centros: los segundos palos y los puntos de penalti "
            f"son zonas críticas a cubrir."
        )

    # Aéreos
    if rr["Pct_aereo"] and rl["Pct_aereo"]:
        if rr["Pct_aereo"] > 55 and rl["Pct_aereo"] < 50:
            recs_defensa.append(
                f"⚠️ Desventaja aérea notable: {rival} gana el {rr['Pct_aereo']:.0f}% de los duelos "
                f"aéreos frente al {rl['Pct_aereo']:.0f}% de {local}. "
                f"Reduce los despejes largos y evita los duelos de segunda jugada."
            )
        elif rl["Pct_aereo"] > rr["Pct_aereo"] + 10:
            recs_defensa.append(
                f"✅ Superioridad aérea ({rl['Pct_aereo']:.0f}% vs {rr['Pct_aereo']:.0f}%). "
                f"Considera el juego de largo a balón parado como herramienta estratégica."
            )

    # Presión alta vs baja del rival
    if rr["Zona_Ata"] and rr["Zona_Def"]:
        if rr["Zona_Ata"] > 35:
            recs_defensa.append(
                f"⚠️ {rival} pasa el {rr['Zona_Ata']:.0f}% del tiempo en zona de ataque. "
                f"Prepara la salida de balón bajo presión: el portero y los centrales "
                f"serán muy exigidos en la fase de construcción."
            )

    # GC propio vs Liga
    if rl["GC_pp"] > df["GC_pp"].quantile(0.70):
        recs_defensa.append(
            f"🔴 Solidez defensiva por mejorar: encajas {rl['GC_pp']:.2f} GC/PJ "
            f"(percentil {percentil(rl['GC_pp'], df['GC_pp'])}% en la Liga). "
            f"Prioridad táctica: reducir los xG concedidos desde dentro del área."
        )

    for rec in recs_defensa:
        st.markdown(rec)

    # ════════════════════════════════════════
    # BLOQUE 4 — JUGADORES CLAVE
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">🌟 Bloque 4 · Jugadores Clave</div>',
                unsafe_allow_html=True)

    col_jk1, col_jk2 = st.columns(2)

    with col_jk1:
        st.markdown(f"**Tu alineación seleccionada ({formacion})**")
        # Mostrar lineup introducido con IS_indiv del jugador
        lineup_df = pd.DataFrame(lineup_local)
        is_map = (inf["jug_local_disp"]
                  .set_index("Jugador")["IS_indiv"]
                  .to_dict())
        lineup_df["IS_individual"] = lineup_df["Jugador"].map(is_map).round(3)
        st.dataframe(lineup_df[["Posicion","Jugador","IS_individual"]],
                     use_container_width=True, hide_index=True)

        # Jugador con mayor IS en la alineación
        top_local_lineup = (
            inf["jug_local_disp"]
            [inf["jug_local_disp"]["Jugador"].isin([r["Jugador"] for r in lineup_local])]
            .sort_values("IS_indiv", ascending=False)
            .head(3)
        )
        if not top_local_lineup.empty:
            st.markdown("**Tus jugadores más influyentes (IS) en la alineación:**")
            for _, row_jug in top_local_lineup.iterrows():
                st.markdown(
                    f"- **{row_jug['Jugador']}** ({row_jug['PosGrupo']}) · "
                    f"IS={row_jug['IS_indiv']:.3f} · Rating={row_jug['Rating']:.2f} · "
                    f"{row_jug['Minutos']:.0f} min"
                )

    with col_jk2:
        st.markdown(f"**Alineación estimada de {rival}**")
        if inf["rival_lineup"].empty:
            st.warning("No hay datos de jugadores del rival.")
        else:
            rival_lineup_show = inf["rival_lineup"].copy()
            rival_lineup_show["IS_indiv"] = rival_lineup_show["IS_indiv"].round(3)
            rival_lineup_show["Minutos"]  = rival_lineup_show["Minutos"].astype(int)
            rival_lineup_show.index = range(1, len(rival_lineup_show)+1)
            st.dataframe(rival_lineup_show[["Jugador","Posicion","Minutos","IS_indiv"]],
                         use_container_width=True)

            # Amenazas del rival
            top_rival = inf["rival_lineup"].sort_values("IS_indiv", ascending=False).head(3)
            st.markdown(f"**Amenazas principales de {rival} a neutralizar:**")
            for _, row_jug in top_rival.iterrows():
                st.markdown(
                    f"- **{row_jug['Jugador']}** ({row_jug['Posicion']}) · "
                    f"IS={row_jug['IS_indiv']:.3f}"
                )

    # Bajas del rival
    if inf["ausentes_rival"]:
        st.info(
            f"✅ Bajas confirmadas de {rival}: **{', '.join(inf['ausentes_rival'])}** — "
            f"aprovecha su ausencia adaptando la estrategia a los sustitutos."
        )

    # ════════════════════════════════════════
    # BLOQUE 5 — CONTEXTO
    # ════════════════════════════════════════
    st.markdown('<div class="section-header">🌦️ Bloque 5 · Contexto del Partido</div>',
                unsafe_allow_html=True)

    col_ctx1, col_ctx2 = st.columns(2)

    with col_ctx1:
        # Árbitro
        st.markdown("**Análisis del árbitro**")
        if arbitro_sel == "(Desconocido)":
            st.caption("Árbitro no seleccionado. No hay análisis disponible.")
        elif inf["arb_faltas"] is not None:
            col_arb1, col_arb2 = st.columns(2)
            col_arb1.metric(
                "Faltas pitadas/PJ",
                f"{inf['arb_faltas']:.1f}",
                delta=f"{inf['arb_faltas'] - inf['liga_faltas']:+.1f} vs media Liga",
            )
            col_arb2.metric(
                "Amarillas/PJ",
                f"{inf['arb_amarillas']:.2f}",
                delta=f"{inf['arb_amarillas'] - inf['liga_amarillas']:+.2f} vs media Liga",
                delta_color="inverse",
            )

            if inf["arb_faltas"] > inf["liga_faltas"] * 1.15:
                st.markdown(
                    f"⚠️ **Árbitro permisivo con el juego físico** "
                    f"({inf['arb_faltas']:.1f} faltas/PJ, +{inf['arb_faltas']-inf['liga_faltas']:.1f} sobre la media). "
                    f"Aprovecha la intensidad en las disputas. Ojo con las acumulaciones de amarillas."
                )
            elif inf["arb_faltas"] < inf["liga_faltas"] * 0.85:
                st.markdown(
                    f"✅ **Árbitro estricto** ({inf['arb_faltas']:.1f} faltas/PJ). "
                    f"Cuida la disciplina táctica: menos faltas disponibles para frenar transiciones rivales."
                )
            else:
                st.markdown(
                    f"🟡 Árbitro dentro de la media ({inf['arb_faltas']:.1f} faltas/PJ). "
                    f"No condiciona el planteamiento táctico."
                )

            # Disciplina del rival con este árbitro como referencia
            if rl["Faltas_com"] > df["Faltas_com"].quantile(0.70):
                st.markdown(
                    f"⚠️ {local} es un equipo con muchas faltas cometidas ({rl['Faltas_com']:.1f}/PJ). "
                    f"Con un árbitro de alta tarjeta ({inf['arb_amarillas']:.2f}/PJ) el riesgo de "
                    f"quedar con inferioridad numérica es elevado."
                )

    with col_ctx2:
        # Lluvia
        st.markdown("**Condiciones meteorológicas**")
        row_ll_local = df[df["Equipo"] == local].iloc[0]
        row_ll_rival = df[df["Equipo"] == rival].iloc[0]

        if lluvia_partido:
            st.markdown(f"🌧️ **Se prevé lluvia en el partido.**")
            st.markdown(
                f"- {local}: {row_ll_local['Total_mm']:.0f} mm acumulados en el estadio "
                f"({row_ll_local['Dias_lluvia']:.0f} días con lluvia esta temporada)"
            )
            st.markdown(
                f"- {rival}: {row_ll_rival['Total_mm']:.0f} mm · "
                f"{row_ll_rival['Dias_lluvia']:.0f} días con lluvia"
            )

            # Habituación a la lluvia (días de lluvia > mediana)
            med_lluvia = df["Dias_lluvia"].median()
            local_hab  = row_ll_local["Dias_lluvia"] >= med_lluvia
            rival_hab  = row_ll_rival["Dias_lluvia"] >= med_lluvia

            if local_hab and not rival_hab:
                st.markdown(
                    "✅ **Ventaja climatológica**: tu equipo está más habituado a jugar "
                    "con lluvia que el rival. El campo pesado puede beneficiar el juego físico "
                    "y directo frente a un equipo de combinación."
                )
            elif rival_hab and not local_hab:
                st.markdown(
                    "⚠️ **Desventaja climatológica**: el rival tiene mayor experiencia "
                    "en condiciones de lluvia. Prioriza el juego directo y reduce "
                    "los pases cortos en el área propia."
                )
            else:
                st.markdown(
                    "🟡 Ambos equipos tienen experiencia similar con lluvia. "
                    "No es un factor diferencial."
                )

            # Nota: hipótesis no confirmada estadísticamente (p=0.349)
            st.caption(
                "_Nota: el análisis estadístico de temporada (p=0.349, NS) no confirma "
                "que la habituación a la lluvia prediga resultados. Es un factor contextual, "
                "no determinante._"
            )
        else:
            st.markdown("☀️ Sin lluvia prevista. No hay condicionante climatológico.")

        # Pérdidas de balón
        st.markdown("**Gestión del balón**")
        if rl["Perdidas"] and rr["Perdidas"]:
            col_p1, col_p2 = st.columns(2)
            col_p1.metric(f"Pérdidas {local}/PJ",  f"{rl['Perdidas']:.1f}")
            col_p2.metric(f"Pérdidas {rival}/PJ", f"{rr['Perdidas']:.1f}")

            if rl["Perdidas"] > rr["Perdidas"] * 1.2:
                st.markdown(
                    f"⚠️ **{local} pierde más balones que {rival}** "
                    f"({rl['Perdidas']:.1f} vs {rr['Perdidas']:.1f}/PJ). "
                    f"Refuerza la seguridad en la salida de balón; las pérdidas en campo propio "
                    f"pueden ser muy costosas contra este rival."
                )
            elif rr["Perdidas"] > rl["Perdidas"] * 1.2:
                st.markdown(
                    f"✅ **{rival} pierde más balones** ({rr['Perdidas']:.1f} vs {rl['Perdidas']:.1f}/PJ). "
                    f"La presión alta tras pérdida rival puede generar situaciones de peligro."
                )

    # ── Resumen ejecutivo final ───────────────────────────────────
    st.divider()
    st.markdown("### 📋 Resumen ejecutivo")

    prioridades = []
    if is_delta > 0.10:
        prioridades.append(f"✅ Partido favorable (ΔIS={is_delta:+.3f}): mantén tu estilo de juego habitual y explota las debilidades detectadas.")
    elif is_delta < -0.10:
        prioridades.append(f"🔴 Partido difícil (ΔIS={is_delta:+.3f}): prioriza la solidez defensiva y el aprovechamiento de las pocas ocasiones que generes.")
    else:
        prioridades.append(f"🟡 Partido equilibrado (ΔIS={is_delta:+.3f}): los detalles marcarán la diferencia.")

    if not inf["dup_rival"].empty:
        top_dup = inf["dup_rival"].iloc[0]
        prioridades.append(
            f"🎯 Neutraliza la dupla clave del rival: **{top_dup['Goleador']}** ← **{top_dup['Asistidor']}** "
            f"({int(top_dup['Frecuencia'])} combinaciones registradas)."
        )

    if inf["ausentes_rival"]:
        prioridades.append(f"✅ Aprovecha la ausencia de: {', '.join(inf['ausentes_rival'])}.")

    if lluvia_partido:
        prioridades.append("🌧️ Prepara el juego para condiciones de campo pesado.")

    for p in prioridades:
        st.markdown(p)

    st.caption(
        f"Informe generado · {local} vs {rival} · "
        f"Árbitro: {arbitro_sel} · "
        f"Formación: {formacion} · "
        f"Algoritmo Prescriptivo La Liga 2025-26"
    )

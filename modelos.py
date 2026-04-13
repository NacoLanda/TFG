"""
modelos.py — Análisis del Dato · Pilar 2 · TFG LaLiga 2025-26
==============================================================
Entrena y evalúa dos modelos predictivos sobre el mismo target:
la diferencia de goles esperada (Goles_local − Goles_visitante).

  Modelo A (Línea Base): Regresión Lineal Múltiple
    → Modelo estadístico simple e interpretable.
    → Asume relación lineal entre las métricas de equipo y el resultado.
    → Métricas: MAE, RMSE, R²

  Modelo B (Avanzado): Random Forest Regressor
    → Conjunto de 300 árboles de decisión con agregación por promedio.
    → Captura relaciones no lineales e interacciones entre variables.
    → Métricas: MAE, RMSE, R²

Evaluación: validación cruzada K-Fold (k=5, aleatorizada, seed=42).
  Se combinan los 670 partidos de ambas temporadas. Cada pliegue actúa
  de test exactamente una vez. Las métricas se calculan sobre las 670
  predicciones out-of-fold (sin data leakage). El modelo final se entrena
  sobre todos los datos y se guarda para app.py.

Features (108): diferencias Δ = local − visitante de:
  · 85 métricas WhoScored (rendimiento de juego, goles por tipo/zona,
       tiros por zona/dirección, pases desglosados, asistencias por tipo,
       paradas por zona, bloqueados detallados)
  · 5  métricas de clasificación (Pts/PJ, GF/PJ, GC/PJ, GD/PJ, Pos%)
  · 10 métricas de portero titular (GA90, %Paradas, %P0, Rating,
       paradas por zona ×3, %PK parados, %pases, pases/pp)
  · 6  métricas de jugadores (IS pilar, IS medio top-5, rating medio top-11,
       % min pilar, Gls/90 goleador, xG/90 medio top-11)
  · 2  variables contextuales (lluvia, ventaja local)

Salida:
  modelos/modelo_a_lineal.pkl    → pipeline Regresión Lineal (entrenado en todos los datos)
  modelos/modelo_b_rf.pkl        → pipeline Random Forest (entrenado en todos los datos)
  modelos/feature_names.pkl      → lista de nombres de features
  modelos/resultados_a.pkl       → predicciones out-of-fold + métricas Modelo A
  modelos/resultados_b.pkl       → predicciones out-of-fold + métricas Modelo B
  modelos/metricas.json          → métricas comparativas de ambos modelos

Uso: python3 modelos.py
     Los .pkl son cargados por app.py y evaluacion.py
"""

import json
import pickle
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import KFold, cross_val_predict
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from xgboost import XGBRegressor

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════
# RUTAS
# Directorios de entrada (datos fuente) y salida (modelos y resultados).
# ══════════════════════════════════════════════════════════════════
BASE_DIR  = Path(__file__).parent
DATOS_DIR = BASE_DIR.parent / "Datos"
PAST_DIR  = DATOS_DIR / "Temporada Pasada"
BD_PATH   = BASE_DIR.parent / "Base de Datos.xlsx"
OUT_DIR   = BASE_DIR / "modelos"
OUT_DIR.mkdir(exist_ok=True)

JUG_PATH_2526 = DATOS_DIR / "Jugadores Unificados.xlsx"
JUG_PATH_2425 = PAST_DIR  / "Jugadores Unificados 24-25.xlsx"

# ══════════════════════════════════════════════════════════════════
# NORMALIZACIÓN DE NOMBRES DE EQUIPOS
# Mapeo canónico para unificar variantes ortográficas entre fuentes.
# ══════════════════════════════════════════════════════════════════
NOMBRE_MAP = {
    "Deportivo Alaves": "Alavés",
    "Atletico Madrid":  "Atlético Madrid",
    "Real Oviedo":      "Oviedo",
    "Alaves":           "Alavés",
    "Atletico":         "Atlético Madrid",
    "Atl Madrid":       "Atlético Madrid",
    "Celta de Vigo":    "Celta Vigo",
    "Betis":            "Real Betis",
    "Sociedad":         "Real Sociedad",
    "Real Valladolid":  "Valladolid",
    "Leganes":          "Leganés",
    "Espanol":          "Espanyol",
}

def normalizar(nombre):
    if pd.isna(nombre):
        return ""
    nombre = str(nombre).strip()
    return NOMBRE_MAP.get(nombre, nombre)


# ══════════════════════════════════════════════════════════════════
# FEATURES DE WHOSCORED — 85 columnas
# Todas se convierten en diferencias Δ = local − visitante.
# ══════════════════════════════════════════════════════════════════
WS_COLS = [
    # ── Rendimiento general ───────────────────────────────────────
    "rating_gen", "tiros_pp_gen", "tirosAP_pp_gen", "regates_pp_gen",
    "faltasFavor_pp_gen", "fueraJuego_pp_gen", "aereos_gen",
    "aciertoPasePct_gen",
    # ── xG ────────────────────────────────────────────────────────
    "xG_fav_gen", "xGDif_fav_gen", "tiros_fav_gen", "xGTiros_fav_gen",
    "xG_con_gen", "xGDif_con_gen",
    # ── Defensa activa ────────────────────────────────────────────
    "tiros_contra_gen", "entradas_pp_gen", "intercep_pp_gen", "faltas_pp_gen",
    "entrada_exito_gen", "entrada_fallo_gen", "intercepciones_gen",
    "despejes_gen", "bloqueados_tiros_gen", "bloqueados_centros_gen",
    "bloqueados_pases_gen",
    # ── Portería (equipo) ─────────────────────────────────────────
    "paradas_total_gen",
    "paradas_pequeña_gen", "paradas_area_gen", "paradas_fuera_gen",
    # ── Duelos aéreos ─────────────────────────────────────────────
    "balonesAereos_ganados_gen", "balonesAereos_perdidos_gen",
    # ── Pases — volumen y precisión ───────────────────────────────
    "pases_total_gen",
    "pases_largosPrecisos_gen", "pases_largosImprecisos_gen",
    "pases_cortosPrecisos_gen", "pases_cortosImprecisos_gen",
    "pasesClave_corto_gen", "pasesClave_largo_gen",
    # ── Pases — tipo (centros, al hueco) ──────────────────────────
    "pase_centros_fav_gen", "pase_alHueco_fav_gen",
    "pase_centros_con_gen", "pase_alHueco_con_gen",
    # ── Asistencias por tipo ──────────────────────────────────────
    "asistencias_centro_gen", "asistencias_corner_gen",
    "asistencias_alHueco_gen", "asistencias_tiroLibre_gen",
    "asistencias_banda_gen",
    # ── Goles por tipo de situación — a favor ─────────────────────
    "gol_juegoAbierto_fav_gen", "gol_contraataque_fav_gen",
    "gol_balonParado_fav_gen", "gol_penalty_fav_gen",
    # ── Goles por tipo de situación — en contra ───────────────────
    "gol_juegoAbierto_con_gen", "gol_contraataque_con_gen",
    "gol_balonParado_con_gen", "gol_penalty_con_gen",
    # ── Goles por zona de tiro ────────────────────────────────────
    "goles_areaPequeña_gen", "goles_area_gen", "goles_fueraArea_gen",
    # ── Tiros por zona de tiro (conteos absolutos) ────────────────
    "tiros_areaPequeña_gen", "tiros_area_gen", "tiros_fueraArea_gen",
    # ── Zonas de tiro — % a favor ─────────────────────────────────
    "zonasTiro_areaPequeña_fav_gen", "zonasTiro_areaPenalty_fav_gen",
    "zonasTiro_areaFuera_fav_gen",
    # ── Zonas de tiro — % en contra ───────────────────────────────
    "zonasTiro_areaPequeña_con_gen", "zonasTiro_areaPenalty_con_gen",
    "zonasTiro_areaFuera_con_gen",
    # ── Dirección del tiro — a favor ──────────────────────────────
    "direccionesTiro_izquierda_fav_gen", "direccionesTiro_centro_fav_gen",
    "direccionesTiro_derecha_fav_gen",
    # ── Dirección del tiro — en contra ────────────────────────────
    "direccionesTiro_izquierda_con_gen", "direccionesTiro_centro_con_gen",
    "direccionesTiro_derecha_con_gen",
    # ── Zonas de acción ───────────────────────────────────────────
    "zonas_accion_ataque_gen", "zonas_accion_mediocampo_gen",
    "zonas_accion_defensa_gen",
    # ── Disciplina ────────────────────────────────────────────────
    "tarjetas_amarilla_gen", "tarjetas_roja_gen",
    # ── Pérdidas ──────────────────────────────────────────────────
    "perdida_desposeido_gen", "perdida_toqueFallido_gen",
    # ── Regates detallados ────────────────────────────────────────
    "regates_exitosos_gen", "regates_fallidos_gen",
    # ── Lado de ataque ────────────────────────────────────────────
    "ladosAtaque_izquierda_gen", "ladosAtaque_centro_gen",
    "ladosAtaque_derecha_gen",
]


# ══════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# Funciones de ingesta para cada fuente: WhoScored, FBref, partidos
# con resultado, porteros titulares y métricas de plantilla.
# ══════════════════════════════════════════════════════════════════

def _to_num(s):
    return pd.to_numeric(s, errors="coerce")


def cargar_whoscored(path, sheet="Equipos"):
    """
    Carga la hoja de equipos de un archivo WhoScored (fila 0 = cabecera).
    Devuelve DataFrame indexado por nombre canónico del equipo.

    Args:
        path:  Ruta al archivo Excel WhoScored.
        sheet: Nombre de la hoja a leer (por defecto "Equipos").

    Returns:
        DataFrame con índice Equipo y columnas WS_COLS disponibles en el archivo.
    """
    df = pd.read_excel(path, sheet_name=sheet, header=0)
    df = df.rename(columns={"equipo": "Equipo"})
    df["Equipo"] = df["Equipo"].apply(normalizar)
    df = df.set_index("Equipo")
    for col in df.columns:
        if col in WS_COLS:
            df[col] = _to_num(df[col].astype(str).str.replace("%", ""))
    cols_disp = [c for c in WS_COLS if c in df.columns]
    return df[cols_disp]


def cargar_clasificacion_2526():
    """
    Extrae Pts/PJ, GF/PJ, GC/PJ, GD/PJ y Pos% de la temporada 25-26.
    Fuente: Base de Datos.xlsx → hojas 'Clasificación' y 'Equipos'.

    Returns:
        DataFrame indexado por Equipo con 5 columnas de clasificación.
    """
    raw_clas = pd.read_excel(BD_PATH, sheet_name="Clasificación", header=None)
    clas = pd.DataFrame({
        "Equipo": raw_clas.iloc[3:, 1].apply(normalizar).values,
        "PJ":     _to_num(raw_clas.iloc[3:, 2]),
        "GF":     _to_num(raw_clas.iloc[3:, 6]),
        "GC":     _to_num(raw_clas.iloc[3:, 7]),
        "Pts":    _to_num(raw_clas.iloc[3:, 9]),
    }).dropna(subset=["PJ"]).copy()
    clas["Pts_pp"] = clas["Pts"] / clas["PJ"]
    clas["GF_pp"]  = clas["GF"]  / clas["PJ"]
    clas["GC_pp"]  = clas["GC"]  / clas["PJ"]
    clas["GD_pp"]  = (clas["GF"] - clas["GC"]) / clas["PJ"]

    raw_eq = pd.read_excel(BD_PATH, sheet_name="Equipos", header=None)
    pos_df = pd.DataFrame({
        "Equipo":  raw_eq.iloc[4:24, 0].apply(normalizar).values,
        "Pos_pct": _to_num(
            raw_eq.iloc[4:24, 3].astype(str).str.replace("%", "")
        ) * 100,
    })

    clas = clas.merge(pos_df, on="Equipo", how="left").set_index("Equipo")
    return clas[["Pts_pp", "GF_pp", "GC_pp", "GD_pp", "Pos_pct"]]


def cargar_clasificacion_2425():
    """
    Extrae Pts/PJ, GF/PJ, GC/PJ, GD/PJ y Pos% de la temporada 24-25.
    Fuente: Datos FBref 24-25.xlsx → hojas 'Tabla General' y 'Squad Standard'.

    Returns:
        DataFrame indexado por Equipo con 5 columnas de clasificación.
    """
    path = PAST_DIR / "Datos FBref 24-25.xlsx"

    raw_gen = pd.read_excel(path, sheet_name="Tabla General", header=None)
    gen = pd.DataFrame({
        "Equipo": raw_gen.iloc[2:, 1].apply(normalizar).values,
        "PJ":     _to_num(raw_gen.iloc[2:, 2]),
        "GF":     _to_num(raw_gen.iloc[2:, 6]),
        "GC":     _to_num(raw_gen.iloc[2:, 7]),
        "Pts":    _to_num(raw_gen.iloc[2:, 9]),
    }).dropna(subset=["PJ"]).copy()
    gen["Pts_pp"] = gen["Pts"] / gen["PJ"]
    gen["GF_pp"]  = gen["GF"]  / gen["PJ"]
    gen["GC_pp"]  = gen["GC"]  / gen["PJ"]
    gen["GD_pp"]  = (gen["GF"] - gen["GC"]) / gen["PJ"]

    raw_std = pd.read_excel(path, sheet_name="Squad Standard", header=None)
    pos_df = pd.DataFrame({
        "Equipo":  raw_std.iloc[2:, 0].apply(normalizar).values,
        "Pos_pct": _to_num(raw_std.iloc[2:, 3]),
    })

    gen = gen.merge(pos_df, on="Equipo", how="left").set_index("Equipo")
    return gen[["Pts_pp", "GF_pp", "GC_pp", "GD_pp", "Pos_pct"]]


def cargar_porteros(path):
    """
    Carga las métricas del portero titular de cada equipo (el de más minutos).

    Columnas del Jugadores Unificados.xlsx → Porteros:
      col7  = Min         col15 = Rating general
      col22 = GA90        col25 = %Paradas     col30 = %P0
      col31 = par_AreaPeq col34 = par_AreaPen  col37 = par_FueraArea
      col44 = %PK parados col45 = %acierto pases  col48 = pases/partido

    Devuelve DataFrame indexado por Equipo con 10 columnas por_ prefijadas.
    """
    raw = pd.read_excel(path, sheet_name="Porteros", header=None)
    df = pd.DataFrame({
        "Equipo":        raw.iloc[3:, 3].apply(normalizar).values,
        "Min":           _to_num(raw.iloc[3:, 7]),
        "GA90":          _to_num(raw.iloc[3:, 22]),
        "pct_par":       _to_num(raw.iloc[3:, 25]),
        "pct_p0":        _to_num(raw.iloc[3:, 30]),
        "rating":        _to_num(raw.iloc[3:, 15]),
        "par_AreaPeq":   _to_num(raw.iloc[3:, 31]),
        "par_AreaPen":   _to_num(raw.iloc[3:, 34]),
        "par_FueraArea": _to_num(raw.iloc[3:, 37]),
        "pct_pk":        _to_num(raw.iloc[3:, 44]),
        "pct_pases":     _to_num(raw.iloc[3:, 45]),
        "pases_pp":      _to_num(raw.iloc[3:, 48]),
    }).dropna(subset=["Min"])

    titular = (
        df.sort_values("Min", ascending=False)
          .groupby("Equipo")
          .first()
          .reset_index()
    )
    titular = titular.set_index("Equipo")[[
        "GA90", "pct_par", "pct_p0",
        "rating", "par_AreaPeq", "par_AreaPen", "par_FueraArea",
        "pct_pk", "pct_pases", "pases_pp",
    ]]
    titular.columns = [
        "por_GA90", "por_pct_par", "por_pct_p0",
        "por_rating", "por_par_AreaPeq", "por_par_AreaPen", "por_par_FueraArea",
        "por_pct_pk", "por_pct_pases", "por_pases_pp",
    ]
    return titular


def cargar_jugadores_features(path):
    """
    Calcula seis métricas de plantilla a partir de jugadores con ≥ 400 min.

    IS_individual = 0.35·norm((xG+Ast)/90) + 0.35·norm(DefScore/90) + 0.30·norm(Rating)
    DefScore/90 = entradas_gan/90 + interc/90 + despejes/90 + bloqueos/90
    Normalización min-max dentro del grupo posicional de cada jugador
    (misma fórmula que en el análisis exploratorio de eda.py).

    Columnas del Jugadores Unificados.xlsx → Jugadores:
      col2  = Posición   col3  = Equipo    col7  = Min
      col22 = Rating     col32 = Gls/90    col33 = Ast/90
      col61 = xG/90      col247= Entradas_gan/90  col250= Interc/90
      col253= Despejes/90  col256= Bloqueos/90

    Métricas producidas:
      jug_pilar_is        → IS máximo del equipo (jugador más valioso)
      jug_avg_is_top5     → IS medio del top-5 (profundidad de plantilla)
      jug_avg_rating_top11→ rating medio de los 11 jugadores con más minutos
      jug_pilar_pct_min   → % de minutos del equipo que acumula el pilar
      jug_top_scorer_g90  → Gls/90 del máximo goleador del equipo
      jug_avg_xg90_top11  → xG/90 medio de los 11 con más minutos
    """
    raw = pd.read_excel(path, sheet_name="Jugadores", header=None)
    df = pd.DataFrame({
        "Posicion": raw.iloc[3:, 2].astype(str).values,
        "Equipo":   raw.iloc[3:, 3].apply(normalizar).values,
        "Min":      _to_num(raw.iloc[3:, 7]),
        "rating":   _to_num(raw.iloc[3:, 22]),
        "g90":      _to_num(raw.iloc[3:, 32]),   # Gls/90
        "ast90":    _to_num(raw.iloc[3:, 33]),   # Ast/90
        "xg90":     _to_num(raw.iloc[3:, 61]),   # xG/90
        "ent90":    _to_num(raw.iloc[3:, 247]),  # Entradas ganadas/90
        "interc90": _to_num(raw.iloc[3:, 250]),  # Intercepciones/90
        "despe90":  _to_num(raw.iloc[3:, 253]),  # Despejes/90
        "bloq90":   _to_num(raw.iloc[3:, 256]),  # Bloqueos/90
    }).dropna(subset=["Min"])

    df = df[df["Min"] >= 400].copy()

    # Grupo posicional (igual que eda.py)
    def _pos_group(pos_str):
        pos = pos_str.split(",")[0].strip().lower()
        if "portero" in pos:
            return "Portero"
        if "defensa" in pos:
            return "Defensa"
        if "delantero" in pos:
            return "Delantero"
        return "Centrocampista"

    df["PosGrupo"] = df["Posicion"].apply(_pos_group)

    # Scores compuestos por 90
    df["OfScore90"]  = df["xg90"].fillna(0) + df["ast90"].fillna(0)
    df["DefScore90"] = (df["ent90"].fillna(0) + df["interc90"].fillna(0)
                        + df["despe90"].fillna(0) + df["bloq90"].fillna(0))

    # Normalización min-max dentro del grupo posicional
    def _mm_group(series, groups):
        result = pd.Series(np.nan, index=series.index)
        for grp in groups.unique():
            mask = groups == grp
            s = series[mask]
            mn, mx = s.min(), s.max()
            result[mask] = (s - mn) / (mx - mn) if mx > mn else 0.5
        return result

    rat_filled = df["rating"].fillna(df.groupby("PosGrupo")["rating"].transform("median"))
    df["norm_of"]  = _mm_group(df["OfScore90"],  df["PosGrupo"])
    df["norm_def"] = _mm_group(df["DefScore90"], df["PosGrupo"])
    df["norm_rat"] = _mm_group(rat_filled,        df["PosGrupo"])

    df["IS"] = (0.35 * df["norm_of"]
                + 0.35 * df["norm_def"]
                + 0.30 * df["norm_rat"])

    filas = []
    for equipo, grupo in df.groupby("Equipo"):
        by_is  = grupo.sort_values("IS", ascending=False)
        by_min = grupo.sort_values("Min", ascending=False)

        pilar_is  = by_is["IS"].iloc[0]
        pilar_min = by_is["Min"].iloc[0]
        total_min = grupo["Min"].sum()

        filas.append({
            "Equipo":               equipo,
            "jug_pilar_is":         pilar_is,
            "jug_avg_is_top5":      by_is["IS"].head(5).mean(),
            "jug_avg_rating_top11": by_min["rating"].head(11).mean(),
            "jug_pilar_pct_min":    pilar_min / total_min if total_min > 0 else np.nan,
            "jug_top_scorer_g90":   grupo["g90"].fillna(0).max(),
            "jug_avg_xg90_top11":   by_min["xg90"].fillna(0).head(11).mean(),
        })

    return pd.DataFrame(filas).set_index("Equipo")


def cargar_partidos(path, tiene_lluvia=True, sheet=None):
    """
    Carga los partidos con resultado y extrae goles locales y visitantes.
    Detecta automáticamente la fila de inicio buscando el guión largo "–".

    Args:
        path:         Ruta al Excel de partidos.
        tiene_lluvia: Si True, lee la columna 8 como indicador binario de lluvia.
        sheet:        Nombre de hoja explícito; si None se usa la hoja activa.

    Returns:
        DataFrame con columnas Home, Away, lluvia, goles_home, goles_away.
    """
    kw = {"sheet_name": sheet} if sheet else {}
    raw = pd.read_excel(path, header=None, **kw)

    inicio = 0
    for i in range(min(8, len(raw))):
        val = str(raw.iloc[i, 5]) if raw.shape[1] > 5 else ""
        if "–" in val:
            inicio = i
            break

    data  = raw.iloc[inicio:].reset_index(drop=True)
    score = data.iloc[:, 5].astype(str)
    mask  = score.str.contains("–", na=False)

    home    = data.iloc[:, 4][mask].apply(normalizar)
    score   = score[mask]
    away    = data.iloc[:, 6][mask].apply(normalizar)

    goles      = score.str.split("–", expand=True)
    goles_home = pd.to_numeric(goles[0].str.strip(), errors="coerce")
    goles_away = pd.to_numeric(goles[1].str.strip(), errors="coerce")

    if tiene_lluvia and data.shape[1] > 8:
        lluvia_vals = data.iloc[:, 8][mask].astype(str).str.lower()
        lluvia_num  = lluvia_vals.str.contains("llovió|lluvia|sí", na=False).astype(int)
    else:
        lluvia_num = pd.Series(0, index=home.index)

    df = pd.DataFrame({
        "Home":       home.values,
        "Away":       away.values,
        "lluvia":     lluvia_num.values,
        "goles_home": goles_home.values,
        "goles_away": goles_away.values,
    })
    return df.dropna(subset=["goles_home"])


# ══════════════════════════════════════════════════════════════════
# FORMA RECIENTE — últimos N partidos por equipo
# Evita data leakage: solo usa partidos ANTERIORES al que se predice.
# ══════════════════════════════════════════════════════════════════

def calcular_forma(partidos_df, ventana=5):
    """
    Para cada partido calcula la media de goles marcados, encajados y
    puntos obtenidos por cada equipo en sus últimos 'ventana' partidos
    ANTERIORES a ese partido (sin data leakage).

    Devuelve DataFrame con índice igual al de partidos_df y columnas:
      forma_gf_local, forma_gc_local, forma_pts_local
      forma_gf_away,  forma_gc_away,  forma_pts_away
    """
    historial = {}   # equipo -> lista de (goles_favor, goles_contra, puntos)

    rows = []
    for idx, row in partidos_df.iterrows():
        home = row["Home"]
        away = row["Away"]

        def get_media(equipo):
            hist  = historial.get(equipo, [])
            recent = hist[-ventana:]
            if not recent:
                return np.nan, np.nan, np.nan
            return (
                np.mean([x[0] for x in recent]),
                np.mean([x[1] for x in recent]),
                np.mean([x[2] for x in recent]),
            )

        gf_h, gc_h, pts_h = get_media(home)
        gf_a, gc_a, pts_a = get_media(away)

        rows.append({
            "forma_gf_local":  gf_h, "forma_gc_local": gc_h, "forma_pts_local": pts_h,
            "forma_gf_away":   gf_a, "forma_gc_away":  gc_a, "forma_pts_away":  pts_a,
        })

        # Actualizar historial con el resultado de ESTE partido
        gh = row["goles_home"]
        ga = row["goles_away"]
        pts_home = 3 if gh > ga else (1 if gh == ga else 0)
        pts_away = 3 if ga > gh else (1 if ga == gh else 0)

        historial.setdefault(home, []).append((gh, ga, pts_home))
        historial.setdefault(away, []).append((ga, gh, pts_away))

    return pd.DataFrame(rows, index=partidos_df.index)


# ══════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL DATASET
# Para cada partido: Δ = home − away de todas las métricas.
# Target único: dif_goles = goles_home − goles_away.
# ══════════════════════════════════════════════════════════════════

def construir_dataset(partidos, ws, clas, por=None, jug=None, forma=None):
    """
    Genera una fila por partido donde cada feature es la diferencia
    entre la métrica del equipo local y la del visitante.

    Args:
        partidos: DataFrame de partidos.
        ws:       DataFrame WhoScored indexado por equipo.
        clas:     DataFrame de clasificación indexado por equipo.
        por:      DataFrame de porteros indexado por equipo (opcional).
        jug:      DataFrame de jugadores indexado por equipo (opcional).

    Returns:
        DataFrame con features + columna target 'dif_goles'.
    """
    filas = []
    cols_ws   = ws.columns.tolist()
    cols_clas = clas.columns.tolist()
    cols_por  = por.columns.tolist() if por  is not None else []
    cols_jug  = jug.columns.tolist() if jug  is not None else []
    todos_equipos = set(ws.index) & set(clas.index)

    for idx, partido in partidos.iterrows():
        home = partido["Home"]
        away = partido["Away"]
        if home not in todos_equipos or away not in todos_equipos:
            continue

        fila = {}

        # ── Diferencias WhoScored (Δ = home − away) ───────────────
        for col in cols_ws:
            fila[f"d_{col}"] = ws.loc[home, col] - ws.loc[away, col]

        # ── Diferencias Clasificación ──────────────────────────────
        for col in cols_clas:
            fila[f"d_{col}"] = clas.loc[home, col] - clas.loc[away, col]

        # ── Diferencias Porteros ───────────────────────────────────
        if por is not None:
            for col in cols_por:
                h_val = por.loc[home, col] if home in por.index else np.nan
                a_val = por.loc[away, col] if away in por.index else np.nan
                fila[f"d_{col}"] = (
                    h_val - a_val
                    if not (np.isnan(h_val) or np.isnan(a_val)) else np.nan
                )

        # ── Diferencias Jugadores ──────────────────────────────────
        if jug is not None:
            for col in cols_jug:
                h_val = jug.loc[home, col] if home in jug.index else np.nan
                a_val = jug.loc[away, col] if away in jug.index else np.nan
                fila[f"d_{col}"] = (
                    h_val - a_val
                    if not (np.isnan(h_val) or np.isnan(a_val)) else np.nan
                )

        # ── Variables contextuales ─────────────────────────────────
        fila["es_local"] = 1
        fila["lluvia"]   = partido["lluvia"]

        # ── Forma reciente (Δ = local − visitante) ─────────────────
        if forma is not None and idx in forma.index:
            f = forma.loc[idx]
            fila["d_forma_gf"]  = (f["forma_gf_local"]  - f["forma_gf_away"]
                                   if not (pd.isna(f["forma_gf_local"]) or pd.isna(f["forma_gf_away"]))
                                   else np.nan)
            fila["d_forma_gc"]  = (f["forma_gc_local"]  - f["forma_gc_away"]
                                   if not (pd.isna(f["forma_gc_local"]) or pd.isna(f["forma_gc_away"]))
                                   else np.nan)
            fila["d_forma_pts"] = (f["forma_pts_local"] - f["forma_pts_away"]
                                   if not (pd.isna(f["forma_pts_local"]) or pd.isna(f["forma_pts_away"]))
                                   else np.nan)

        # ── Target ────────────────────────────────────────────────
        fila["dif_goles"] = partido["goles_home"] - partido["goles_away"]

        filas.append(fila)

    df_out = pd.DataFrame(filas)
    for col in [c for c in df_out.columns if c != "dif_goles"]:
        if df_out[col].isna().any():
            df_out[col] = df_out[col].fillna(df_out[col].median())
    return df_out


# ══════════════════════════════════════════════════════════════════
# MÉTRICAS DE REGRESIÓN
# MAE, RMSE y R² calculados sobre las predicciones out-of-fold.
# ══════════════════════════════════════════════════════════════════

def calcular_metricas(y_real, y_pred):
    """
    Calcula MAE, RMSE y R² redondeados a 4 decimales.

    Args:
        y_real: Array de valores reales.
        y_pred: Array de predicciones.

    Returns:
        Dict con claves 'mae', 'rmse' y 'r2'.
    """
    mae  = mean_absolute_error(y_real, y_pred)
    rmse = np.sqrt(mean_squared_error(y_real, y_pred))
    r2   = r2_score(y_real, y_pred)
    return {"mae": round(mae, 4), "rmse": round(rmse, 4), "r2": round(r2, 4)}


# ══════════════════════════════════════════════════════════════════
# MODELO A — REGRESIÓN LINEAL MÚLTIPLE
# Línea base interpretable; requiere escalado previo (StandardScaler)
# para que los coeficientes sean comparables entre variables.
# ══════════════════════════════════════════════════════════════════

def entrenar_modelo_a(X_all, y_all, feature_names, kf):
    """
    Evalúa la Regresión Lineal Múltiple con K-Fold CV y guarda el modelo final.
    El escalado es necesario para comparar coeficientes entre variables.

    Args:
        X_all:         Matriz de features (n_partidos × n_features).
        y_all:         Vector objetivo (diferencia de goles).
        feature_names: Lista de nombres de las features.
        kf:            Objeto KFold configurado.

    Returns:
        Dict con métricas, coeficientes estandarizados, feature_names,
        y_test e y_pred (predicciones out-of-fold).
    """
    print("\n" + "═"*60)
    print("  MODELO A — Regresión Lineal Múltiple (Línea Base)")
    print("═"*60)

    pipeline_a = Pipeline([
        ("scaler", StandardScaler()),
        ("reg",    LinearRegression()),
    ])

    y_pred_cv = cross_val_predict(pipeline_a, X_all, y_all, cv=kf)
    metricas = calcular_metricas(y_all, y_pred_cv)
    print(f"  MAE  (CV): {metricas['mae']:.4f} goles")
    print(f"  RMSE (CV): {metricas['rmse']:.4f} goles")
    print(f"  R²   (CV): {metricas['r2']:.4f}")

    pipeline_a.fit(X_all, y_all)
    coefs = pipeline_a.named_steps["reg"].coef_
    top_n = 15
    idx_top = np.argsort(np.abs(coefs))[::-1][:top_n]
    print(f"\n  Top {top_n} coeficientes del modelo final (estandarizados):")
    for i in idx_top:
        print(f"    {feature_names[i]:<50} {coefs[i]:+.4f}")

    with open(OUT_DIR / "modelo_a_lineal.pkl", "wb") as f:
        pickle.dump(pipeline_a, f)
    print(f"\n  Guardado: modelos/modelo_a_lineal.pkl")

    return {
        **metricas,
        "coeficientes":  coefs.tolist(),
        "feature_names": feature_names,
        "y_test":        list(y_all),
        "y_pred":        list(y_pred_cv),
    }


# ══════════════════════════════════════════════════════════════════
# MODELO B — RANDOM FOREST REGRESSOR
# Modelo avanzado seleccionado para producción; captura relaciones
# no lineales e interacciones entre variables sin supuestos de linealidad.
# ══════════════════════════════════════════════════════════════════

def entrenar_modelo_b(X_all, y_all, feature_names, kf):
    """
    Evalúa el Random Forest Regressor con K-Fold CV y guarda el modelo final.

    Args:
        X_all:         Matriz de features (n_partidos × n_features).
        y_all:         Vector objetivo (diferencia de goles).
        feature_names: Lista de nombres de las features.
        kf:            Objeto KFold configurado.

    Returns:
        Dict con métricas, importancias Gini, feature_names,
        y_test e y_pred (predicciones out-of-fold).
    """
    print("\n" + "═"*60)
    print("  MODELO B — Random Forest Regressor (Avanzado)")
    print("═"*60)

    pipeline_b = Pipeline([
        ("scaler", StandardScaler()),
        ("rf",     RandomForestRegressor(
            n_estimators=300,
            max_depth=8,
            min_samples_split=10,
            min_samples_leaf=5,
            max_features="sqrt",
            random_state=42,
            n_jobs=-1,
        )),
    ])

    y_pred_cv = cross_val_predict(pipeline_b, X_all, y_all, cv=kf)
    metricas = calcular_metricas(y_all, y_pred_cv)
    print(f"  MAE  (CV): {metricas['mae']:.4f} goles")
    print(f"  RMSE (CV): {metricas['rmse']:.4f} goles")
    print(f"  R²   (CV): {metricas['r2']:.4f}")

    pipeline_b.fit(X_all, y_all)
    importancias = pipeline_b.named_steps["rf"].feature_importances_
    top_n = 20
    idx_top = np.argsort(importancias)[::-1][:top_n]
    print(f"\n  Top {top_n} features del modelo final (importancia Gini):")
    for i in idx_top:
        print(f"    {feature_names[i]:<50} {importancias[i]:.4f}")

    with open(OUT_DIR / "modelo_b_rf.pkl", "wb") as f:
        pickle.dump(pipeline_b, f)
    print(f"\n  Guardado: modelos/modelo_b_rf.pkl")

    return {
        **metricas,
        "importancias":  importancias.tolist(),
        "feature_names": feature_names,
        "y_test":        list(y_all),
        "y_pred":        list(y_pred_cv),
    }


# ══════════════════════════════════════════════════════════════════
# MODELO C — XGBOOST REGRESSOR
# Gradient boosting con árboles superficiales (max_depth=3) y tasa
# de aprendizaje baja (lr=0.02) para reducir el sobreajuste.
# ══════════════════════════════════════════════════════════════════

def entrenar_modelo_c(X_all, y_all, feature_names, kf):
    """
    Evalúa XGBoost Regressor con K-Fold CV y guarda el modelo final.

    Args:
        X_all:         Matriz de features (n_partidos × n_features).
        y_all:         Vector objetivo (diferencia de goles).
        feature_names: Lista de nombres de las features.
        kf:            Objeto KFold configurado.

    Returns:
        Dict con métricas, importancias XGB, feature_names,
        y_test e y_pred (predicciones out-of-fold).
    """
    print("\n" + "═"*60)
    print("  MODELO C — XGBoost Regressor (Gradient Boosting)")
    print("═"*60)

    pipeline_c = Pipeline([
        ("scaler", StandardScaler()),
        ("xgb",    XGBRegressor(
            n_estimators=500,
            max_depth=3,
            learning_rate=0.02,
            min_child_weight=10,
            gamma=1.0,
            subsample=0.8,
            colsample_bytree=0.8,
            reg_alpha=0.5,
            reg_lambda=2.0,
            random_state=42,
            n_jobs=-1,
            verbosity=0,
        )),
    ])

    y_pred_cv = cross_val_predict(pipeline_c, X_all, y_all, cv=kf)
    metricas = calcular_metricas(y_all, y_pred_cv)
    print(f"  MAE  (CV): {metricas['mae']:.4f} goles")
    print(f"  RMSE (CV): {metricas['rmse']:.4f} goles")
    print(f"  R²   (CV): {metricas['r2']:.4f}")

    pipeline_c.fit(X_all, y_all)
    importancias = pipeline_c.named_steps["xgb"].feature_importances_
    top_n = 20
    idx_top = np.argsort(importancias)[::-1][:top_n]
    print(f"\n  Top {top_n} features del modelo final (importancia XGB):")
    for i in idx_top:
        print(f"    {feature_names[i]:<50} {importancias[i]:.4f}")

    with open(OUT_DIR / "modelo_c_xgb.pkl", "wb") as f:
        pickle.dump(pipeline_c, f)
    print(f"\n  Guardado: modelos/modelo_c_xgb.pkl")

    return {
        **metricas,
        "importancias":  importancias.tolist(),
        "feature_names": feature_names,
        "y_test":        list(y_all),
        "y_pred":        list(y_pred_cv),
    }


# ══════════════════════════════════════════════════════════════════
# EJECUCIÓN PRINCIPAL
# Carga datos de ambas temporadas, construye el dataset combinado
# (670 partidos) y entrena los tres modelos con K-Fold CV (k=5).
# ══════════════════════════════════════════════════════════════════

def main():
    print("\n" + "═"*60)
    print("  CARGANDO DATOS")
    print("═"*60)

    # ── Temporada 25-26 ────────────────────────────────────────────
    ws_2526   = cargar_whoscored(DATOS_DIR / "Datos WhoScored.xlsx")
    clas_2526 = cargar_clasificacion_2526()
    par_2526  = cargar_partidos(BD_PATH, tiene_lluvia=True, sheet="Partidos")
    por_2526  = cargar_porteros(JUG_PATH_2526)
    jug_2526  = cargar_jugadores_features(JUG_PATH_2526)
    print(f"  25-26: {len(par_2526)} partidos | "
          f"{len(ws_2526)} equipos WS ({len(ws_2526.columns)} cols) | "
          f"{len(por_2526)} porteros | {len(jug_2526)} equipos jugadores")

    # ── Temporada 24-25 ────────────────────────────────────────────
    ws_2425   = cargar_whoscored(PAST_DIR / "Datos WhoScored 24-25.xlsx")
    clas_2425 = cargar_clasificacion_2425()
    par_2425  = cargar_partidos(PAST_DIR / "Partidos 24-25.xlsx", tiene_lluvia=False)
    por_2425  = cargar_porteros(JUG_PATH_2425)
    jug_2425  = cargar_jugadores_features(JUG_PATH_2425)
    print(f"  24-25: {len(par_2425)} partidos | "
          f"{len(ws_2425)} equipos WS ({len(ws_2425.columns)} cols) | "
          f"{len(por_2425)} porteros | {len(jug_2425)} equipos jugadores")

    # ── Forma reciente (últimos 5 partidos por equipo) ────────────
    forma_2526 = calcular_forma(par_2526)
    forma_2425 = calcular_forma(par_2425)

    # ── Construir un dataset por temporada ────────────────────────
    print("\n  Construyendo features...")
    df_2526 = construir_dataset(par_2526, ws_2526, clas_2526,
                                por=por_2526, jug=jug_2526, forma=forma_2526)
    df_2425 = construir_dataset(par_2425, ws_2425, clas_2425,
                                por=por_2425, jug=jug_2425, forma=forma_2425)
    print(f"  Dataset 25-26: {len(df_2526)} filas")
    print(f"  Dataset 24-25: {len(df_2425)} filas")

    # Combinar temporadas; usar solo las features comunes
    feat_2526 = [c for c in df_2526.columns if c != "dif_goles"]
    feat_2425 = [c for c in df_2425.columns if c != "dif_goles"]
    feature_cols = [c for c in feat_2425 if c in feat_2526]

    df_all = pd.concat(
        [df_2425[feature_cols + ["dif_goles"]],
         df_2526[feature_cols + ["dif_goles"]]],
        ignore_index=True,
    )
    print(f"  Dataset combinado: {len(df_all)} partidos")
    print(f"  Features usadas  : {len(feature_cols)}")

    X_all = df_all[feature_cols].fillna(0).values
    y_all = df_all["dif_goles"].values

    print(f"\n  Target — media: {y_all.mean():.2f}  "
          f"std: {y_all.std():.2f}  rango: [{y_all.min():.0f}, {y_all.max():.0f}]")

    # ── K-Fold ─────────────────────────────────────────────────────
    kf = KFold(n_splits=5, shuffle=True, random_state=42)

    with open(OUT_DIR / "feature_names.pkl", "wb") as f:
        pickle.dump(feature_cols, f)

    # ── Entrenar y evaluar ─────────────────────────────────────────
    res_a = entrenar_modelo_a(X_all, y_all, feature_cols, kf)
    res_b = entrenar_modelo_b(X_all, y_all, feature_cols, kf)
    res_c = entrenar_modelo_c(X_all, y_all, feature_cols, kf)

    # ── Guardar resultados ─────────────────────────────────────────
    with open(OUT_DIR / "resultados_a.pkl", "wb") as f:
        pickle.dump(res_a, f)
    with open(OUT_DIR / "resultados_b.pkl", "wb") as f:
        pickle.dump(res_b, f)
    with open(OUT_DIR / "resultados_c.pkl", "wb") as f:
        pickle.dump(res_c, f)

    metricas_out = {
        "modelo_a": {"mae": res_a["mae"], "rmse": res_a["rmse"], "r2": res_a["r2"]},
        "modelo_b": {"mae": res_b["mae"], "rmse": res_b["rmse"], "r2": res_b["r2"]},
        "modelo_c": {"mae": res_c["mae"], "rmse": res_c["rmse"], "r2": res_c["r2"]},
        "n_partidos": len(df_all),
        "n_folds": 5,
        "n_features": len(feature_cols),
    }
    with open(OUT_DIR / "metricas.json", "w", encoding="utf-8") as f:
        json.dump(metricas_out, f, ensure_ascii=False, indent=2)

    # ── Resumen comparativo ────────────────────────────────────────
    print("\n" + "═"*60)
    print("  RESUMEN COMPARATIVO — validación cruzada 5-Fold")
    print("═"*60)
    print(f"  {'Métrica':<10} {'Modelo A (Lineal)':<20} {'Modelo B (RF)':<18} {'Modelo C (XGB)':<18} Mejor")
    print(f"  {'─'*74}")

    def mejor_3(a, b, c, menor_es_mejor=True):
        vals = {"Modelo A": a, "Modelo B": b, "Modelo C": c}
        return min(vals, key=vals.__getitem__) if menor_es_mejor else max(vals, key=vals.__getitem__)

    print(f"  {'MAE':<10} {res_a['mae']:<20} {res_b['mae']:<18} {res_c['mae']:<18} "
          f"{mejor_3(res_a['mae'], res_b['mae'], res_c['mae'])}")
    print(f"  {'RMSE':<10} {res_a['rmse']:<20} {res_b['rmse']:<18} {res_c['rmse']:<18} "
          f"{mejor_3(res_a['rmse'], res_b['rmse'], res_c['rmse'])}")
    print(f"  {'R²':<10} {res_a['r2']:<20} {res_b['r2']:<18} {res_c['r2']:<18} "
          f"{mejor_3(res_a['r2'], res_b['r2'], res_c['r2'], menor_es_mejor=False)}")

    best_r2 = max(res_a['r2'], res_b['r2'], res_c['r2'])
    if best_r2 == res_c['r2']:
        best_name = "Modelo C (XGBoost)"
    elif best_r2 == res_b['r2']:
        best_name = "Modelo B (Random Forest)"
    else:
        best_name = "Modelo A (Regresión Lineal)"
    print(f"\n  → Modelo seleccionado para producción: {best_name}")

    print("\n  Archivos en modelos/:")
    for f in ["modelo_a_lineal.pkl", "modelo_b_rf.pkl", "modelo_c_xgb.pkl",
              "feature_names.pkl", "resultados_a.pkl",
              "resultados_b.pkl", "metricas.json"]:
        print(f"    {f}")
    print("═"*60 + "\n")


if __name__ == "__main__":
    main()

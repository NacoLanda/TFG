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

Comparativa directa: ambos modelos predicen el mismo target con las mismas
features y se evalúan con las mismas métricas → selección del modelo final
justificada técnica y empíricamente.

Datos:
  Entrenamiento → Temporada 24-25 (380 partidos, temporada completa)
  Test          → Temporada 25-26 (290 partidos, validación temporal)

Features (57): diferencias Δ = local − visitante de métricas WhoScored
  + clasificación (Pts/PJ, GF/PJ, GC/PJ, GD/PJ, Pos%) + contextuales
  (lluvia, árbitro FaltasPP, ArmarillasPP).

Salida:
  modelos/modelo_a_lineal.pkl    → pipeline Regresión Lineal
  modelos/modelo_b_rf.pkl        → pipeline Random Forest
  modelos/feature_names.pkl      → lista de nombres de features
  modelos/resultados_a.pkl       → predicciones + métricas Modelo A
  modelos/resultados_b.pkl       → predicciones + métricas Modelo B
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
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════
# RUTAS
# ══════════════════════════════════════════════════════════════════
BASE_DIR  = Path(__file__).parent
DATOS_DIR = BASE_DIR.parent / "Datos"
PAST_DIR  = DATOS_DIR / "Temporada Pasada"
BD_PATH   = BASE_DIR.parent / "Base de Datos.xlsx"
OUT_DIR   = BASE_DIR / "modelos"
OUT_DIR.mkdir(exist_ok=True)

# ══════════════════════════════════════════════════════════════════
# NORMALIZACIÓN DE NOMBRES DE EQUIPOS
# Las tres fuentes (WhoScored, FBref, Partidos) usan nombres distintos.
# Este mapa unifica todo a un nombre canónico en español.
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
    """Devuelve el nombre canónico del equipo, o el original si no hay mapeo."""
    if pd.isna(nombre):
        return ""
    nombre = str(nombre).strip()
    return NOMBRE_MAP.get(nombre, nombre)


# ══════════════════════════════════════════════════════════════════
# FEATURES DE WHOSCORED
# Columnas que se usarán de cada equipo (sufijo _gen = General).
# Todas se convertirán en diferencias Δ = local − visitante.
# ══════════════════════════════════════════════════════════════════
WS_COLS = [
    # Rendimiento general
    "rating_gen", "tiros_pp_gen", "tirosAP_pp_gen", "regates_pp_gen",
    "faltasFavor_pp_gen", "fueraJuego_pp_gen", "aereos_gen",
    "aciertoPasePct_gen",
    # xG
    "xG_fav_gen", "xGDif_fav_gen", "tiros_fav_gen", "xGTiros_fav_gen",
    "xG_con_gen", "xGDif_con_gen",
    # Defensa activa
    "tiros_contra_gen", "entradas_pp_gen", "intercep_pp_gen", "faltas_pp_gen",
    "entrada_exito_gen", "entrada_fallo_gen", "intercepciones_gen",
    "despejes_gen", "bloqueados_tiros_gen", "bloqueados_centros_gen",
    # Portería
    "paradas_total_gen",
    # Duelos aéreos
    "balonesAereos_ganados_gen", "balonesAereos_perdidos_gen",
    # Pases
    "pases_total_gen", "pasesClave_corto_gen", "pasesClave_largo_gen",
    # Zonas de acción
    "zonas_accion_ataque_gen", "zonas_accion_mediocampo_gen",
    "zonas_accion_defensa_gen",
    # Disciplina
    "tarjetas_amarilla_gen", "tarjetas_roja_gen",
    # Goles por tipo de jugada
    "gol_balonParado_fav_gen", "gol_contraataque_fav_gen",
    "gol_juegoAbierto_fav_gen",
    # Pérdidas
    "perdida_desposeido_gen", "perdida_toqueFallido_gen",
    # Regates detallados
    "regates_exitosos_gen", "regates_fallidos_gen",
    # Lado de ataque
    "ladosAtaque_izquierda_gen", "ladosAtaque_centro_gen",
    "ladosAtaque_derecha_gen",
    # % zonas de tiro
    "zonasTiro_areaPequeña_fav_gen", "zonasTiro_areaPenalty_fav_gen",
    "zonasTiro_areaFuera_fav_gen",
]


# ══════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ══════════════════════════════════════════════════════════════════

def _to_num(s):
    """Convierte una serie a numérico, reemplazando errores con NaN."""
    return pd.to_numeric(s, errors="coerce")


def cargar_whoscored(path, sheet="Equipos"):
    """
    Carga la hoja de equipos de un archivo WhoScored (fila 0 = cabecera).

    Args:
        path:  Ruta al archivo Excel de WhoScored.
        sheet: Nombre de la hoja (por defecto 'Equipos').

    Returns:
        DataFrame indexado por nombre canónico del equipo.
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
    Extrae Pts/PJ, GF/PJ, GC/PJ, GD/PJ y posesión % de la temporada 25-26.

    Fuente: Base de Datos.xlsx → hojas 'Clasificación' y 'Equipos'.

    Returns:
        DataFrame indexado por Equipo con columnas Pts_pp, GF_pp, GC_pp,
        GD_pp, Pos_pct.
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

    # Posesión: col 3 de Equipos viene como decimal (0.503) → ×100
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
    Extrae Pts/PJ, GF/PJ, GC/PJ, GD/PJ y posesión % de la temporada 24-25.

    Fuente: Datos FBref 24-25.xlsx → hojas 'Tabla General' y 'Squad Standard'.

    Returns:
        DataFrame indexado por Equipo con mismas columnas que 2526.
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


def cargar_arbitros_2526():
    """
    Carga FaltasPP y AmarillasPP de cada árbitro de la temporada 25-26.

    Fuente: Base de Datos.xlsx → hoja 'Árbitros'.

    Returns:
        DataFrame con columnas Arbitro, arb_faltas_pp, arb_amarillas_pp.
    """
    raw = pd.read_excel(BD_PATH, sheet_name="Árbitros", header=None)
    arb = pd.DataFrame({
        "Arbitro":          raw.iloc[2:, 0].astype(str).str.strip().values,
        "arb_faltas_pp":    _to_num(raw.iloc[2:, 4]),
        "arb_amarillas_pp": _to_num(raw.iloc[2:, 6]),
    }).dropna(subset=["arb_faltas_pp"])
    return arb[arb["Arbitro"] != "nan"].reset_index(drop=True)


def cargar_arbitros_2425():
    """
    Carga FaltasPP y AmarillasPP de cada árbitro de la temporada 24-25.

    Fuente: Datos Árbitros 24-25.xlsx.

    Returns:
        DataFrame con columnas Arbitro, arb_faltas_pp, arb_amarillas_pp.
    """
    raw = pd.read_excel(PAST_DIR / "Datos Árbitros 24-25.xlsx", header=None)
    arb = pd.DataFrame({
        "Arbitro":          raw.iloc[2:, 0].astype(str).str.strip().values,
        "arb_faltas_pp":    _to_num(raw.iloc[2:, 5]),
        "arb_amarillas_pp": _to_num(raw.iloc[2:, 9]),
    }).dropna(subset=["arb_faltas_pp"])
    return arb[arb["Arbitro"] != "nan"].reset_index(drop=True)


def _match_arbitro(nombre, arb_df):
    """
    Busca el árbitro por coincidencia exacta o por primeros 5 chars del apellido.

    Args:
        nombre: Nombre del árbitro en el dataset de partidos.
        arb_df: DataFrame de árbitros con FaltasPP y AmarillasPP.

    Returns:
        Tupla (faltas_pp, amarillas_pp) o (NaN, NaN) si no se encuentra.
    """
    nombre = str(nombre).strip()
    fila = arb_df[arb_df["Arbitro"] == nombre]
    if len(fila):
        return fila.iloc[0]["arb_faltas_pp"], fila.iloc[0]["arb_amarillas_pp"]
    key = nombre.split()[0].lower()[:5] if nombre else ""
    fila = arb_df[arb_df["Arbitro"].str.split().str[0].str.lower().str[:5] == key]
    if len(fila):
        return fila.iloc[0]["arb_faltas_pp"], fila.iloc[0]["arb_amarillas_pp"]
    return np.nan, np.nan


def cargar_partidos(path, tiene_lluvia=True, sheet=None):
    """
    Carga los partidos con resultado y extrae goles locales y visitantes.

    Detecta automáticamente la fila de inicio de datos buscando el guión
    largo "–" en la columna de resultado.

    Args:
        path:         Ruta al Excel de partidos.
        tiene_lluvia: True si la columna 8 contiene información de lluvia.
        sheet:        Nombre de hoja si el Excel tiene múltiples hojas.

    Returns:
        DataFrame con columnas: Jornada, Home, Away, Arbitro, lluvia,
        goles_home, goles_away.
    """
    kw = {"sheet_name": sheet} if sheet else {}
    raw = pd.read_excel(path, header=None, **kw)

    # Localizar la primera fila con resultado (contiene "–")
    inicio = 0
    for i in range(min(8, len(raw))):
        val = str(raw.iloc[i, 5]) if raw.shape[1] > 5 else ""
        if "–" in val:
            inicio = i
            break

    data = raw.iloc[inicio:].reset_index(drop=True)

    score   = data.iloc[:, 5].astype(str)
    mask    = score.str.contains("–", na=False)

    jornada = data.iloc[:, 0][mask]
    home    = data.iloc[:, 4][mask].apply(normalizar)
    score   = score[mask]
    away    = data.iloc[:, 6][mask].apply(normalizar)
    arbitro = data.iloc[:, 7][mask]

    goles   = score.str.split("–", expand=True)
    goles_home = pd.to_numeric(goles[0].str.strip(), errors="coerce")
    goles_away = pd.to_numeric(goles[1].str.strip(), errors="coerce")

    if tiene_lluvia and data.shape[1] > 8:
        lluvia_vals = data.iloc[:, 8][mask].astype(str).str.lower()
        lluvia_num  = lluvia_vals.str.contains("llovió|lluvia|sí", na=False).astype(int)
    else:
        lluvia_num = pd.Series(0, index=home.index)

    df = pd.DataFrame({
        "Jornada":    pd.to_numeric(jornada.values, errors="coerce"),
        "Home":       home.values,
        "Away":       away.values,
        "Arbitro":    arbitro.values,
        "lluvia":     lluvia_num.values,
        "goles_home": goles_home.values,
        "goles_away": goles_away.values,
    })
    return df.dropna(subset=["goles_home"])


# ══════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL DATASET
# Para cada partido: Δ = home − away de todas las métricas.
# Target único: dif_goles = goles_home − goles_away.
# ══════════════════════════════════════════════════════════════════

def construir_dataset(partidos, ws, clas, arb):
    """
    Combina los datos de partidos con las métricas de equipos.

    Genera una fila por partido donde cada feature es la diferencia
    entre la métrica del equipo local y la del visitante.

    Args:
        partidos: DataFrame de partidos (salida de cargar_partidos).
        ws:       DataFrame WhoScored indexado por equipo.
        clas:     DataFrame de clasificación indexado por equipo.
        arb:      DataFrame de árbitros.

    Returns:
        DataFrame con features + columna target 'dif_goles'.
    """
    filas = []
    cols_ws   = ws.columns.tolist()
    cols_clas = clas.columns.tolist()
    todos_equipos = set(ws.index) & set(clas.index)

    for _, partido in partidos.iterrows():
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

        # ── Variables contextuales ─────────────────────────────────
        fila["es_local"] = 1   # perspectiva siempre desde el home
        fila["lluvia"]   = partido["lluvia"]
        faltas_pp, amarillas_pp = _match_arbitro(partido["Arbitro"], arb)
        fila["arb_faltas_pp"]    = faltas_pp
        fila["arb_amarillas_pp"] = amarillas_pp

        # ── Target ────────────────────────────────────────────────
        fila["dif_goles"] = partido["goles_home"] - partido["goles_away"]

        filas.append(fila)

    df_out = pd.DataFrame(filas)
    # Rellenar NaN de árbitro con la mediana (hay partidos sin árbitro asignado)
    for col in ["arb_faltas_pp", "arb_amarillas_pp"]:
        df_out[col] = df_out[col].fillna(df_out[col].median())
    return df_out


# ══════════════════════════════════════════════════════════════════
# MÉTRICAS DE REGRESIÓN — función compartida por ambos modelos
# ══════════════════════════════════════════════════════════════════

def calcular_metricas(y_real, y_pred):
    """
    Calcula MAE, RMSE y R² entre valores reales y predichos.

    Args:
        y_real: Array de valores reales.
        y_pred: Array de valores predichos.

    Returns:
        Dict con claves mae, rmse, r2.
    """
    mae  = mean_absolute_error(y_real, y_pred)
    rmse = np.sqrt(mean_squared_error(y_real, y_pred))
    r2   = r2_score(y_real, y_pred)
    return {"mae": round(mae, 4), "rmse": round(rmse, 4), "r2": round(r2, 4)}


# ══════════════════════════════════════════════════════════════════
# MODELO A — REGRESIÓN LINEAL MÚLTIPLE
# Línea base interpretable: cada coeficiente indica cuánto sube la
# diferencia de goles por unidad de diferencia en esa métrica.
# ══════════════════════════════════════════════════════════════════

def entrenar_modelo_a(X_train, y_train, X_test, y_test, feature_names):
    """
    Entrena la Regresión Lineal Múltiple con StandardScaler.

    El escalado es necesario para que los coeficientes sean comparables
    entre sí (todas las features quedan en la misma escala).

    Args:
        X_train, y_train: Datos de entrenamiento (temporada 24-25).
        X_test,  y_test:  Datos de evaluación (temporada 25-26).
        feature_names:    Lista de nombres de columnas.

    Returns:
        Dict con métricas, predicciones y coeficientes del modelo.
    """
    print("\n" + "═"*60)
    print("  MODELO A — Regresión Lineal Múltiple (Línea Base)")
    print("═"*60)

    pipeline_a = Pipeline([
        ("scaler", StandardScaler()),
        ("reg",    LinearRegression()),
    ])
    pipeline_a.fit(X_train, y_train)
    y_pred = pipeline_a.predict(X_test)

    metricas = calcular_metricas(y_test, y_pred)
    print(f"  MAE  : {metricas['mae']:.4f} goles")
    print(f"  RMSE : {metricas['rmse']:.4f} goles")
    print(f"  R²   : {metricas['r2']:.4f}")

    # Coeficientes estandarizados (el scaler ya está aplicado internamente)
    coefs = pipeline_a.named_steps["reg"].coef_
    top_n = 15
    idx_top = np.argsort(np.abs(coefs))[::-1][:top_n]
    print(f"\n  Top {top_n} coeficientes (estandarizados):")
    for i in idx_top:
        print(f"    {feature_names[i]:<45} {coefs[i]:+.4f}")

    with open(OUT_DIR / "modelo_a_lineal.pkl", "wb") as f:
        pickle.dump(pipeline_a, f)
    print(f"\n  Guardado: modelos/modelo_a_lineal.pkl")

    return {
        **metricas,
        "coeficientes":  coefs.tolist(),
        "feature_names": feature_names,
        "y_test":        list(y_test),
        "y_pred":        list(y_pred),
    }


# ══════════════════════════════════════════════════════════════════
# MODELO B — RANDOM FOREST REGRESSOR
# Modelo avanzado: 300 árboles que votan por promedio.
# Captura interacciones no lineales invisibles a la regresión lineal.
# ══════════════════════════════════════════════════════════════════

def entrenar_modelo_b(X_train, y_train, X_test, y_test, feature_names):
    """
    Entrena el Random Forest Regressor.

    El RF no requiere escalado pero se mantiene el scaler en el pipeline
    para que app.py pueda usar ambos modelos con la misma interfaz.

    Args:
        X_train, y_train: Datos de entrenamiento (temporada 24-25).
        X_test,  y_test:  Datos de evaluación (temporada 25-26).
        feature_names:    Lista de nombres de columnas.

    Returns:
        Dict con métricas, predicciones e importancias de variables.
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
    pipeline_b.fit(X_train, y_train)
    y_pred = pipeline_b.predict(X_test)

    metricas = calcular_metricas(y_test, y_pred)
    print(f"  MAE  : {metricas['mae']:.4f} goles")
    print(f"  RMSE : {metricas['rmse']:.4f} goles")
    print(f"  R²   : {metricas['r2']:.4f}")

    importancias = pipeline_b.named_steps["rf"].feature_importances_
    top_n = 15
    idx_top = np.argsort(importancias)[::-1][:top_n]
    print(f"\n  Top {top_n} features por importancia (Gini):")
    for i in idx_top:
        print(f"    {feature_names[i]:<45} {importancias[i]:.4f}")

    with open(OUT_DIR / "modelo_b_rf.pkl", "wb") as f:
        pickle.dump(pipeline_b, f)
    print(f"\n  Guardado: modelos/modelo_b_rf.pkl")

    return {
        **metricas,
        "importancias":  importancias.tolist(),
        "feature_names": feature_names,
        "y_test":        list(y_test),
        "y_pred":        list(y_pred),
    }


# ══════════════════════════════════════════════════════════════════
# EJECUCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════

def main():
    print("\n" + "═"*60)
    print("  CARGANDO DATOS")
    print("═"*60)

    # ── Temporada 25-26 (test) ─────────────────────────────────────
    ws_2526   = cargar_whoscored(DATOS_DIR / "Datos WhoScored.xlsx")
    clas_2526 = cargar_clasificacion_2526()
    arb_2526  = cargar_arbitros_2526()
    par_2526  = cargar_partidos(BD_PATH, tiene_lluvia=True, sheet="Partidos")
    print(f"  25-26: {len(par_2526)} partidos | "
          f"{len(ws_2526)} equipos WhoScored | {len(clas_2526)} equipos clas.")

    # ── Temporada 24-25 (entrenamiento) ───────────────────────────
    ws_2425   = cargar_whoscored(PAST_DIR / "Datos WhoScored 24-25.xlsx")
    clas_2425 = cargar_clasificacion_2425()
    arb_2425  = cargar_arbitros_2425()
    par_2425  = cargar_partidos(PAST_DIR / "Partidos 24-25.xlsx", tiene_lluvia=False)
    print(f"  24-25: {len(par_2425)} partidos | "
          f"{len(ws_2425)} equipos WhoScored | {len(clas_2425)} equipos clas.")

    # ── Construir datasets ─────────────────────────────────────────
    print("\n  Construyendo features...")
    df_train = construir_dataset(par_2425, ws_2425, clas_2425, arb_2425)
    df_test  = construir_dataset(par_2526, ws_2526, clas_2526, arb_2526)
    print(f"  Dataset train: {len(df_train)} filas")
    print(f"  Dataset test : {len(df_test)} filas")

    # Usar solo las features presentes en ambas temporadas
    feat_train = [c for c in df_train.columns if c != "dif_goles"]
    feat_test  = [c for c in df_test.columns  if c != "dif_goles"]
    feature_cols = [c for c in feat_train if c in feat_test]
    print(f"  Features usadas: {len(feature_cols)}")

    X_train = df_train[feature_cols].fillna(0).values
    X_test  = df_test[feature_cols].fillna(0).values
    y_train = df_train["dif_goles"].values
    y_test  = df_test["dif_goles"].values

    print(f"\n  Target train — media: {y_train.mean():.2f}  "
          f"std: {y_train.std():.2f}  rango: [{y_train.min():.0f}, {y_train.max():.0f}]")
    print(f"  Target test  — media: {y_test.mean():.2f}  "
          f"std: {y_test.std():.2f}  rango: [{y_test.min():.0f}, {y_test.max():.0f}]")

    # ── Guardar feature names ──────────────────────────────────────
    with open(OUT_DIR / "feature_names.pkl", "wb") as f:
        pickle.dump(feature_cols, f)

    # ── Entrenar ───────────────────────────────────────────────────
    res_a = entrenar_modelo_a(X_train, y_train, X_test, y_test, feature_cols)
    res_b = entrenar_modelo_b(X_train, y_train, X_test, y_test, feature_cols)

    # ── Guardar resultados para evaluacion.py ──────────────────────
    with open(OUT_DIR / "resultados_a.pkl", "wb") as f:
        pickle.dump(res_a, f)
    with open(OUT_DIR / "resultados_b.pkl", "wb") as f:
        pickle.dump(res_b, f)

    # ── JSON de métricas comparativas ─────────────────────────────
    metricas_out = {
        "modelo_a": {"mae": res_a["mae"], "rmse": res_a["rmse"], "r2": res_a["r2"]},
        "modelo_b": {"mae": res_b["mae"], "rmse": res_b["rmse"], "r2": res_b["r2"]},
    }
    with open(OUT_DIR / "metricas.json", "w", encoding="utf-8") as f:
        json.dump(metricas_out, f, ensure_ascii=False, indent=2)

    # ── Resumen comparativo ────────────────────────────────────────
    print("\n" + "═"*60)
    print("  RESUMEN COMPARATIVO — mismo target, mismas métricas")
    print("═"*60)
    print(f"  {'Métrica':<10} {'Modelo A (Lineal)':<22} {'Modelo B (RF)':<20} Ganador")
    print(f"  {'─'*62}")

    def ganador(a, b, menor_es_mejor=True):
        if menor_es_mejor:
            return "Modelo B" if b < a else "Modelo A"
        return "Modelo B" if b > a else "Modelo A"

    print(f"  {'MAE':<10} {res_a['mae']:<22} {res_b['mae']:<20} "
          f"{ganador(res_a['mae'], res_b['mae'])}")
    print(f"  {'RMSE':<10} {res_a['rmse']:<22} {res_b['rmse']:<20} "
          f"{ganador(res_a['rmse'], res_b['rmse'])}")
    print(f"  {'R²':<10} {res_a['r2']:<22} {res_b['r2']:<20} "
          f"{ganador(res_a['r2'], res_b['r2'], menor_es_mejor=False)}")
    print(f"\n  → Modelo seleccionado para app.py: "
          f"{'Modelo B (RF)' if res_b['r2'] > res_a['r2'] else 'Modelo A (Lineal)'}")

    print("\n  Archivos en modelos/:")
    for f in ["modelo_a_lineal.pkl", "modelo_b_rf.pkl",
              "feature_names.pkl", "resultados_a.pkl",
              "resultados_b.pkl", "metricas.json"]:
        print(f"    {f}")
    print("═"*60 + "\n")


if __name__ == "__main__":
    main()

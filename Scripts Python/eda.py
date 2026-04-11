"""
eda.py — Análisis Exploratorio de Datos Táctico · LaLiga 2025-26
=================================================================
Genera los 16 gráficos del TFG (G1–G16) a partir del Excel principal.
Cada gráfico busca relaciones estadísticas con valor prescriptivo:
saber QUÉ métricas realmente predicen el rendimiento (puntos, goles
encajados) para poder hacer recomendaciones tácticas calibradas.

Gráficos generados en Gráficos/:
  G1  heatmap_ofensiva            → Métricas ofensivas correladas con Pts
  G2  heatmap_defensiva           → Métricas defensivas correladas con GC/Pts
  G3  heatmap_posesion_pase       → Posesión y pase correlados con Pts
  G4  correlaciones_sorprendentes → Matriz completa 35 métricas (azul/rojo)
  G5  scatter_xg_goles            → xG vs Goles (eficiencia rematadora)
  G6  cuadrantes_xgdif            → Cuadrantes xGDif ataque vs defensa
  G7  bubble_posesion_zona        → Posesión vs zona ofensiva (con burbuja)
  G8  zonas_accion_apiladas       → Zonas de acción por equipo (barra apilada)
  G9  radar_perfil                → Radares de perfil — Rendimiento y Estilo Táctico
  G10 validacion_is               → Validación del Índice de Éxito vs Puntos reales
  G11 clima_lluvia_is             → IS vs Pts/partido en condiciones de lluvia
  G12 scatter_arbitros            → Faltas/partido árbitros vs Rating equipo
  G13 pilares_concentracion       → Jugador pilar (max On-Off) por equipo
  G14 duplas_peligrosas           → Duplas goleador-asistidor por equipo
  G15 scatter_influencia          → %Min vs IS_individual (4 cuadrantes)
  G16 porteros_rendimiento        → Rendimiento de porteros — Real vs Esperado (xG)

Fuente: Base de Datos.xlsx  (hojas Equipos, Clasificación, Lluvias,
         Árbitros, Jugadores, Duplas Peligrosas)

IS = 0.35·norm(xG/PJ) + 0.35·norm(−GC/PJ) + 0.30·norm(Rating)

Uso: python3 eda.py
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from scipy import stats

# ═══════════════════════════════════════════════════════════════
# CONFIGURACIÓN GLOBAL
# ═══════════════════════════════════════════════════════════════
sns.set_theme(style="whitegrid", font_scale=1.15)
plt.rcParams.update({
    "axes.titlesize": 15,
    "axes.labelsize": 13,
    "xtick.labelsize": 11,
    "ytick.labelsize": 11,
    "legend.fontsize": 10,
})

EXCEL  = "/Users/NacoLG/Documentos/UFV 4/TFG/Ingeniería del Dato/Base de Datos.xlsx"
OUTDIR = "/Users/NacoLG/Documentos/UFV 4/TFG/Ingeniería del Dato/Gráficos/"


# ═══════════════════════════════════════════════════════════════
# CARGA Y LIMPIEZA
# ═══════════════════════════════════════════════════════════════

def pct_num(series):
    """'58%' → 58.0  (valor numérico, no proporción)."""
    return (
        series.astype(str)
              .str.replace("%", "", regex=False)
              .pipe(pd.to_numeric, errors="coerce")
    )

raw_eq  = pd.read_excel(EXCEL, sheet_name="Equipos",       header=None)
raw_cla = pd.read_excel(EXCEL, sheet_name="Clasificación", header=None)
raw_ll  = pd.read_excel(EXCEL, sheet_name="Lluvias",       header=None)

data = raw_eq.iloc[4:24].copy().reset_index(drop=True)   # 20 equipos

# ── Métricas Equipos ────────────────────────────────────────────
eq = pd.DataFrame({
    "Equipo":          data.iloc[:, 0],
    "Pos%":            pd.to_numeric(data.iloc[:, 3],   errors="coerce") * 100,
    "Rating":          pd.to_numeric(data.iloc[:, 4],   errors="coerce"),
    "Goles":           pd.to_numeric(data.iloc[:, 7],   errors="coerce"),
    # Ataque — xG
    "xG":              pd.to_numeric(data.iloc[:, 57],  errors="coerce"),
    "xGDif":           pd.to_numeric(data.iloc[:, 60],  errors="coerce"),  # Goles - xG (>0 eficiente)
    "xG_contra":       pd.to_numeric(data.iloc[:, 63],  errors="coerce"),
    "xGDif_contra":    pd.to_numeric(data.iloc[:, 66],  errors="coerce"),  # GC - xGcontra (>0 rival eficiente)
    "xG_tiro":         pd.to_numeric(data.iloc[:, 69],  errors="coerce"),
    # Ataque — Tiros
    "Tiros_pp":        pd.to_numeric(data.iloc[:, 81],  errors="coerce"),
    "Tiros_contra_pp": pd.to_numeric(data.iloc[:, 84],  errors="coerce"),
    "Tiros_puerta_pp": pd.to_numeric(data.iloc[:, 87],  errors="coerce"),
    "Pct_tiro_area":   pct_num(data.iloc[:, 102]),   # % tiros desde área
    # Defensa activa
    "Entradas_exit":   pd.to_numeric(data.iloc[:, 170], errors="coerce"),
    "Entradas_fall":   pd.to_numeric(data.iloc[:, 173], errors="coerce"),
    "Intercepciones":  pd.to_numeric(data.iloc[:, 176], errors="coerce"),
    "Despejes":        pd.to_numeric(data.iloc[:, 179], errors="coerce"),
    "Bloqueos_tiro":   pd.to_numeric(data.iloc[:, 182], errors="coerce"),
    "Aereos_gan":      pd.to_numeric(data.iloc[:, 191], errors="coerce"),
    "Aereos_per":      pd.to_numeric(data.iloc[:, 194], errors="coerce"),
    # Pases
    "Total_pases":     pd.to_numeric(data.iloc[:, 197], errors="coerce"),
    "Precision_pase":  pd.to_numeric(data.iloc[:, 200], errors="coerce"),
    "Pases_L_prec":    pd.to_numeric(data.iloc[:, 209], errors="coerce"),
    "Pases_L_impr":    pd.to_numeric(data.iloc[:, 212], errors="coerce"),
    "Pases_C_prec":    pd.to_numeric(data.iloc[:, 215], errors="coerce"),
    "PasesClave_cort": pd.to_numeric(data.iloc[:, 224], errors="coerce"),
    # Zonas de acción (vienen como '34%')
    "Zona_Def":        pct_num(data.iloc[:, 266]),
    "Zona_Med":        pct_num(data.iloc[:, 269]),
    "Zona_Ata":        pct_num(data.iloc[:, 272]),
    # Pérdidas
    "Perdidas":        pd.to_numeric(data.iloc[:, 275], errors="coerce")
                       + pd.to_numeric(data.iloc[:, 278], errors="coerce"),
    # Disciplina
    "Faltas_com":      pd.to_numeric(data.iloc[:, 281], errors="coerce"),
    "Amarillas":       pd.to_numeric(data.iloc[:, 284], errors="coerce"),
    # Plantilla — variables para G4
    "Edad":            pd.to_numeric(data.iloc[:,   2], errors="coerce"),  # edad media
    "Reg_exit":        pd.to_numeric(data.iloc[:, 138], errors="coerce"),  # regates exitosos/pp
    "Reg_fall":        pd.to_numeric(data.iloc[:, 141], errors="coerce"),  # regates fallidos/pp
    "Off_pp":          pd.to_numeric(data.iloc[:, 145], errors="coerce"),  # fueras de juego/pp
    "SavePct":         pd.to_numeric(data.iloc[:, 153], errors="coerce"),  # % paradas portero
    "P0":              pd.to_numeric(data.iloc[:, 154], errors="coerce"),  # porterías a cero
}).reset_index(drop=True)

# Ratio pases largos sobre total
eq["Pct_pase_largo"] = (eq["Pases_L_prec"] + eq["Pases_L_impr"]) / eq["Total_pases"] * 100
# Eficiencia de entradas (exitosas / total)
eq["Efic_entrada"] = eq["Entradas_exit"] / (eq["Entradas_exit"] + eq["Entradas_fall"]) * 100
# Ratio ganado de aéreos
eq["Pct_aereo"] = eq["Aereos_gan"] / (eq["Aereos_gan"] + eq["Aereos_per"]) * 100

# ── Clasificación: Pts, GF, GC ──────────────────────────────────
cla = pd.DataFrame({
    "Equipo": raw_cla.iloc[3:23, 1].values,
    "PJ":     pd.to_numeric(raw_cla.iloc[3:23, 2], errors="coerce"),
    "GF":     pd.to_numeric(raw_cla.iloc[3:23, 6], errors="coerce"),
    "GC":     pd.to_numeric(raw_cla.iloc[3:23, 7], errors="coerce"),
    "Pts":    pd.to_numeric(raw_cla.iloc[3:23, 9], errors="coerce"),
}).reset_index(drop=True)

# ── Lluvias ──────────────────────────────────────────────────────
lluvia = pd.DataFrame({
    "Equipo":      raw_ll.iloc[2:22, 0].values,
    "Total_mm":    pd.to_numeric(raw_ll.iloc[2:22, 3], errors="coerce"),
    "Dias_lluvia": pd.to_numeric(raw_ll.iloc[2:22, 4], errors="coerce"),
}).reset_index(drop=True)

# ── Merge maestro ────────────────────────────────────────────────
df = eq.merge(cla, on="Equipo", how="inner").merge(lluvia, on="Equipo", how="inner")
df = df.loc[:, ~df.columns.duplicated()].copy()   # elimina columnas duplicadas si las hubiera
df["Pts_pp"] = df["Pts"] / df["PJ"]   # Puntos por partido
df["xG_pp"]  = df["xG"]  / df["PJ"]   # xG por partido (capacidad ofensiva)
df["GC_pp"]  = df["GC"]  / df["PJ"]   # Goles encajados por partido (solidez defensiva)

# ── Índice de Éxito (IS) v2 ─────────────────────────────────────
# IS = 0.35·norm(xG/PJ) + 0.35·norm(−GC/PJ) + 0.30·norm(Rating)
#
# Los tres componentes son estadísticamente significativos (p<0.001):
#   · xG/PJ      r=+0.871 *** — capacidad ofensiva real por partido
#   · −GC/PJ     r=−0.832 *** — solidez defensiva (invertida)
#   · Rating Gen r=+0.937 *** — calidad colectiva agregada
#
# Se descartó xGDif (r=+0.296, p=0.20, NS): en esta temporada la
# eficiencia rematadora no discrimina entre equipos de forma
# estadísticamente significativa. Incluirla con peso alto sesgaba
# el índice hacia equipos con suerte en conversión (Sevilla
# aparecía en top-5 con 49 GC y solo 31 pts).
#
# Prescripción: IS alto → ajustes tácticos finos.
#               IS bajo → revisión táctica profunda.
def minmax(s):
    return (s - s.min()) / (s.max() - s.min())

df["IS"] = (
    0.35 * minmax(df["xG_pp"]) +
    0.35 * minmax(-df["GC_pp"]) +
    0.30 * minmax(df["Rating"])
)

print(f"Dataset final: {df.shape[0]} equipos × {df.shape[1]} variables")
print(df[["Equipo", "Pts", "xG", "xGDif", "GC"]].to_string(index=False))


# ═══════════════════════════════════════════════════════════════
# UTILIDAD: etiquetar scatter con todos los equipos
# ═══════════════════════════════════════════════════════════════
def label_scatter(ax, x_vals, y_vals, labels, fontsize=8.5):
    """
    Anota cada punto con el nombre del equipo.
    Coloca la etiqueta en el cuadrante libre respecto al centroide.
    """
    cx, cy = np.mean(x_vals), np.mean(y_vals)
    for xi, yi, lbl in zip(x_vals, y_vals, labels):
        dx = 6 if xi >= cx else -6
        dy = 5 if yi >= cy else -7
        ax.annotate(
            lbl, xy=(xi, yi),
            xytext=(dx, dy), textcoords="offset points",
            fontsize=fontsize, color="#1a1a2e",
            arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5),
        )


def save(fig, name):
    path = OUTDIR + name
    fig.savefig(path, dpi=150, bbox_inches="tight")
    print(f"  Guardado: {path}")
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════
# G1 — HEATMAP: Métricas Ofensivas → Resultados
# ---------------------------------------------------------------
# Relación oculta: ¿importa más el volumen de tiros o la calidad
# (xG/tiro)? ¿La posesión realmente genera goles?
# Una correlación alta de xGDif con Pts revela que los equipos
# que materializan sus ocasiones MEJOR de lo esperado escalan
# la clasificación más allá de su dominio.
# ═══════════════════════════════════════════════════════════════
print("\n[G1] Heatmap ofensiva...")
# Orden conceptual: Output → Volumen disparo → Calidad disparo → Posición/posesión
# GF eliminado por ser redundante con Goles (r=0.997)
off_cols = {
    "Pts":               "Pts",
    "Goles":             "Goles",
    "Tiros/pp":          "Tiros_pp",
    "Tiros a puerta/pp": "Tiros_puerta_pp",
    "xG (total)":        "xG",
    "xG/Tiro":           "xG_tiro",
    "% Tiros desde área":"Pct_tiro_area",
    "xGDif (eficiencia)":"xGDif",
    "Posesión %":        "Pos%",
    "Zona Ofensiva %":   "Zona_Ata",
}
df_off  = df[[v for v in off_cols.values()]].rename(columns={v: k for k, v in off_cols.items()})
corr_off = df_off.corr()

# ── Matriz de p-valores para marcar significancia ───────────────
from itertools import combinations
pval_off = pd.DataFrame(np.ones_like(corr_off), index=corr_off.index, columns=corr_off.columns)
for c1, c2 in combinations(corr_off.columns, 2):
    xy = df_off[[c1, c2]].dropna()
    _, p = stats.pearsonr(xy.iloc[:, 0], xy.iloc[:, 1])
    pval_off.loc[c1, c2] = p
    pval_off.loc[c2, c1] = p

# Anotaciones: valor r + estrellas de significancia
annot_off = pd.DataFrame("", index=corr_off.index, columns=corr_off.columns)
for c1 in corr_off.columns:
    for c2 in corr_off.columns:
        if c1 == c2:
            continue
        r = corr_off.loc[c1, c2]
        p = pval_off.loc[c1, c2]
        stars = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "NS"
        annot_off.loc[c1, c2] = f"{r:.2f}\n{stars}"

mask_off = np.triu(np.ones_like(corr_off, dtype=bool))

fig, ax = plt.subplots(figsize=(13, 10))
sns.heatmap(
    corr_off, mask=mask_off,
    annot=annot_off, fmt="",
    annot_kws={"size": 8.5, "va": "center"},
    cmap="RdYlGn", vmin=-1, vmax=1,
    linewidths=0.4, linecolor="white",
    cbar_kws={"shrink": 0.75, "label": "Pearson r"},
    ax=ax,
)
ax.set_title(
    "G1 · Correlación Métricas Ofensivas\n"
    "¿Qué genera realmente los puntos?",
    fontweight="bold", pad=14,
)
ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
ax.set_yticklabels(ax.get_yticklabels(), rotation=0)
fig.text(
    0.5, -0.02,
    "Verde = correlación positiva · Rojo = negativa · "
    "*** p<0.001  ** p<0.01  * p<0.05  NS = no significativo  ·  n = 20 equipos",
    ha="center", fontsize=9, color="gray",
)
save(fig, "g1_heatmap_ofensiva.png")


# ═══════════════════════════════════════════════════════════════
# G2 — HEATMAP: Métricas Defensivas → Resultados
# ---------------------------------------------------------------
# Relación oculta: ¿los equipos que interceptan más son los que
# menos goles encajan? ¿Los despejes correlacionan negativamente
# con Pts (equipos sufridos que también acaban arriba)?
# La diferencia entre xGDif_contra y GC revela cuánto aporta el
# portero vs la organización defensiva colectiva.
# ═══════════════════════════════════════════════════════════════
print("[G2] Heatmap defensiva...")
def_cols = {
    "Pts":                  "Pts",
    "GC/pp":                "GC_pp",
    "xG en contra":         "xG_contra",
    "xGDif contra":         "xGDif_contra",   # >0: rival eficiente vs nosotros (malo)
    "Despejes/pp":          "Despejes",
    "Bloqueos tiro/pp":     "Bloqueos_tiro",
    "Aéreos ganados/pp":    "Aereos_gan",
    "% Aéreos ganados":     "Pct_aereo",
    "Entradas exitosas/pp": "Entradas_exit",
    "Efic. Entradas %":     "Efic_entrada",
    "Intercepciones/pp":    "Intercepciones",
}
df_def = df[[v for v in def_cols.values()]].rename(columns={v: k for k, v in def_cols.items()})
corr_def = df_def.corr()

# ── Matriz de p-valores para marcar significancia ───────────────
pval_def = pd.DataFrame(np.ones_like(corr_def), index=corr_def.index, columns=corr_def.columns)
for c1, c2 in combinations(corr_def.columns, 2):
    xy = df_def[[c1, c2]].dropna()
    _, p = stats.pearsonr(xy.iloc[:, 0], xy.iloc[:, 1])
    pval_def.loc[c1, c2] = p
    pval_def.loc[c2, c1] = p

# Anotaciones: valor r + estrellas de significancia
annot_def = pd.DataFrame("", index=corr_def.index, columns=corr_def.columns)
for c1 in corr_def.columns:
    for c2 in corr_def.columns:
        if c1 == c2:
            continue
        r = corr_def.loc[c1, c2]
        p = pval_def.loc[c1, c2]
        stars = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "NS"
        annot_def.loc[c1, c2] = f"{r:.2f}\n{stars}"

mask = np.triu(np.ones_like(corr_def, dtype=bool))

fig, ax = plt.subplots(figsize=(13, 10))
sns.heatmap(
    corr_def, mask=mask,
    annot=annot_def, fmt="",
    annot_kws={"size": 8.5, "va": "center"},
    cmap="RdYlGn", vmin=-1, vmax=1,
    linewidths=0.4, linecolor="white",
    cbar_kws={"shrink": 0.75, "label": "Pearson r"},
    ax=ax,
)
ax.set_title(
    "G2 · Correlación Métricas Defensivas\n"
    "¿Qué protege realmente la portería?",
    fontweight="bold", pad=14,
)
ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
ax.set_yticklabels(ax.get_yticklabels(), rotation=0)
fig.text(
    0.5, -0.02,
    "Verde = correlación positiva · Rojo = negativa · "
    "*** p<0.001  ** p<0.01  * p<0.05  NS = no significativo  ·  n = 20 equipos  ·  "
    "xGDif contra > 0: el rival marcó más de lo esperado (peligroso). Efic. Entradas = exitosas / total.",
    ha="center", fontsize=9, color="gray",
)
save(fig, "g2_heatmap_defensiva.png")


# ═══════════════════════════════════════════════════════════════
# G3 — HEATMAP: Posesión, Pase y Zonas → Resultados
# ---------------------------------------------------------------
# Relación oculta: ¿la precisión de pase importa más que el
# volumen? ¿Los equipos que juegan en la zona media son
# más "seguros" pero menos efectivos que los que arriesgan
# en zona ofensiva? ¿Las pérdidas de balón penalizan más en
# equipos con alta posesión?
# ═══════════════════════════════════════════════════════════════
print("[G3] Heatmap posesión-pase...")
# Orden conceptual: Resultado → Output ofensivo → Creación → Control → Volumen/Estilo → Zonas
# GF eliminado por redundancia (r=0.92 con xG, r=0.91 con Pts)
pas_cols = {
    "Pts":                  "Pts",
    "xG":                   "xG",
    "Pases clave cortos/pp":"PasesClave_cort",
    "Pérdidas/pp":          "Perdidas",
    "Posesión %":           "Pos%",
    "Precisión pase %":     "Precision_pase",
    "Total pases/pp":       "Total_pases",
    "% Pases largos":       "Pct_pase_largo",
    "Zona Ofensiva %":      "Zona_Ata",
    "Zona Media %":         "Zona_Med",
    "Zona Defensiva %":     "Zona_Def",
}
df_pas = df[[v for v in pas_cols.values()]].rename(columns={v: k for k, v in pas_cols.items()})
corr_pas = df_pas.corr()

# ── Matriz de p-valores para marcar significancia ───────────────
pval_pas = pd.DataFrame(np.ones_like(corr_pas), index=corr_pas.index, columns=corr_pas.columns)
for c1, c2 in combinations(corr_pas.columns, 2):
    xy = df_pas[[c1, c2]].dropna()
    _, p = stats.pearsonr(xy.iloc[:, 0], xy.iloc[:, 1])
    pval_pas.loc[c1, c2] = p
    pval_pas.loc[c2, c1] = p

# Anotaciones: valor r + estrellas de significancia
annot_pas = pd.DataFrame("", index=corr_pas.index, columns=corr_pas.columns)
for c1 in corr_pas.columns:
    for c2 in corr_pas.columns:
        if c1 == c2:
            continue
        r = corr_pas.loc[c1, c2]
        p = pval_pas.loc[c1, c2]
        stars = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "NS"
        annot_pas.loc[c1, c2] = f"{r:.2f}\n{stars}"

mask_pas = np.triu(np.ones_like(corr_pas, dtype=bool))

fig, ax = plt.subplots(figsize=(13, 10))
sns.heatmap(
    corr_pas, mask=mask_pas,
    annot=annot_pas, fmt="",
    annot_kws={"size": 8.5, "va": "center"},
    cmap="RdYlGn", vmin=-1, vmax=1,
    linewidths=0.4, linecolor="white",
    cbar_kws={"shrink": 0.75, "label": "Pearson r"},
    ax=ax,
)
ax.set_title(
    "G3 · Correlación Posesión, Pase y Zonas\n"
    "¿El control del balón se traduce en goles?",
    fontweight="bold", pad=14,
)
ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
ax.set_yticklabels(ax.get_yticklabels(), rotation=0)
fig.text(
    0.5, -0.02,
    "Verde = correlación positiva · Rojo = negativa · "
    "*** p<0.001  ** p<0.01  * p<0.05  NS = no significativo  ·  n = 20 equipos  ·  "
    "% Pases largos = (largos precisos + imprecisos) / total.",
    ha="center", fontsize=9, color="gray",
)
save(fig, "g3_heatmap_posesion_pase.png")


# ═══════════════════════════════════════════════════════════════
# G5 — SCATTER: xG vs Goles Reales (eficiencia rematadora)
# ---------------------------------------------------------------
# Relación oculta: equipos por encima de la diagonal convierten
# mejor de lo esperado (potencial prescriptivo: mantener el
# mismo volumen pero mejorar la selección del tiro).
# El color muestra si ser "eficiente" correlaciona con más Pts.
# ═══════════════════════════════════════════════════════════════
print("[G5] Scatter xG vs Goles...")
fig, ax = plt.subplots(figsize=(11, 8))

norm  = plt.Normalize(df["Pts"].min(), df["Pts"].max())
cmap  = plt.cm.RdYlGn
sc    = ax.scatter(
    df["xG"], df["Goles"],
    c=df["Pts"], cmap=cmap, norm=norm,
    s=130, edgecolors="white", linewidths=0.8, zorder=3,
)
plt.colorbar(sc, ax=ax, label="Puntos en la clasificación", shrink=0.85)

# Línea de referencia xG = Goles
xy_min = min(df["xG"].min(), df["Goles"].min()) - 3
xy_max = max(df["xG"].max(), df["Goles"].max()) + 3
ax.plot([xy_min, xy_max], [xy_min, xy_max],
        color="#555", lw=1.5, ls=":", label="xG = Goles (conversión perfecta)")

# Línea de regresión
slope, intercept, r, p, _ = stats.linregress(df["xG"], df["Goles"])
x_line = np.linspace(df["xG"].min() - 2, df["xG"].max() + 2, 100)
ax.plot(x_line, slope * x_line + intercept,
        color="#2c3e50", lw=2, ls="--",
        label=f"Regresión (r={r:.2f}, p={p:.3f})")

# Sombrear zona de sobrerendimiento
ax.fill_between([xy_min, xy_max], [xy_min, xy_max], [xy_max, xy_max],
                alpha=0.04, color="green")
ax.fill_between([xy_min, xy_max], [xy_min, xy_max], [xy_min, xy_min],
                alpha=0.04, color="red")
ax.text(xy_max - 2, xy_max - 1, "Sobreconvierte", fontsize=9,
        color="darkgreen", ha="right", va="bottom", style="italic")
ax.text(xy_max - 2, xy_min + 1, "Infraconvierte", fontsize=9,
        color="darkred", ha="right", va="top", style="italic")

from adjustText import adjust_text
texts4 = []
for xi, yi, lbl in zip(df["xG"], df["Goles"], df["Equipo"]):
    texts4.append(ax.text(xi, yi, lbl, fontsize=8.5, color="#1a1a2e"))
adjust_text(
    texts4, x=df["xG"].values, y=df["Goles"].values,
    arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5),
    ax=ax,
)
ax.set_xlabel("xG — Goles Esperados (calidad de ocasiones)")
ax.set_ylabel("Goles Reales")
ax.set_title(
    "G5 · Eficiencia Rematadora: xG vs Goles Reales\n"
    "¿Quién saca más partido de sus ocasiones?",
    fontweight="bold",
)
ax.legend(loc="upper left")
fig.text(
    0.5, -0.02,
    "Línea punteada fina: xG = Goles (conversión perfecta — diagonal 45°). "
    "Por encima → sobreconvierte; por debajo → infraconvierte respecto al modelo. "
    "Línea discontinua: regresión real sobre los 20 equipos (r=0.93). "
    "Que la regresión quede por debajo de la diagonal indica que la liga en conjunto infraconvierte.",
    ha="center", fontsize=8.5, color="gray",
)
save(fig, "g5_scatter_xg_goles.png")


# ═══════════════════════════════════════════════════════════════
# G6 — CUADRANTES TÁCTICOS: xGDif Ataque vs xGDif Defensa
# ---------------------------------------------------------------
# Relación oculta: la ventaja real de un equipo no es solo
# cuántas ocasiones crea sino cuántas desperdicia el rival.
# xGDif_ataque > 0  → equipo anota MÁS de lo que debería
# xGDif_def   > 0  → el rival TAMBIÉN anota más de lo que
#                    debería (portero/defensa sufre)
# Cuadrante ideal: ataque +, defensa - (anota más, encaja menos)
# ═══════════════════════════════════════════════════════════════
print("[G6] Cuadrantes tácticos xGDif...")
fig, ax = plt.subplots(figsize=(11, 8))

# Invertir xGDif_contra: negativo = RIVAL falla (beneficia al equipo)
df["xGDif_def_inv"] = -df["xGDif_contra"]   # >0: rival falla = buena defensa

norm = plt.Normalize(df["Pts"].min(), df["Pts"].max())
sizes5 = ((df["Pts"] - df["Pts"].min()) / (df["Pts"].max() - df["Pts"].min()) * 400 + 80)
sc = ax.scatter(
    df["xGDif"], df["xGDif_def_inv"],
    c=df["Pts"], cmap=plt.cm.RdYlGn, norm=norm,
    s=sizes5, edgecolors="white", linewidths=0.8, zorder=3,
)
plt.colorbar(sc, ax=ax, label="Puntos", shrink=0.85)

ax.axhline(0, color="#555", lw=1.2, ls="--", zorder=2)
ax.axvline(0, color="#555", lw=1.2, ls="--", zorder=2)

# Sombreado de cuadrantes (dibujado antes de fijar límites definitivos)
xlim = ax.get_xlim()
ylim = ax.get_ylim()
ax.fill_between([0, xlim[1]], 0, ylim[1],  alpha=0.07, color="green",      zorder=0)
ax.fill_between([xlim[0], 0], 0, ylim[1],  alpha=0.07, color="steelblue",  zorder=0)
ax.fill_between([0, xlim[1]], ylim[0], 0,  alpha=0.07, color="darkorange", zorder=0)
ax.fill_between([xlim[0], 0], ylim[0], 0,  alpha=0.07, color="darkred",    zorder=0)
ax.set_xlim(xlim)
ax.set_ylim(ylim)

# Etiquetas con adjustText
from adjustText import adjust_text
texts5 = []
for xi, yi, lbl in zip(df["xGDif"], df["xGDif_def_inv"], df["Equipo"]):
    texts5.append(ax.text(xi, yi, lbl, fontsize=8.5, color="#1a1a2e"))
adjust_text(
    texts5, x=df["xGDif"].values, y=df["xGDif_def_inv"].values,
    arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5),
    ax=ax,
)

# Leyenda de cuadrantes
from matplotlib.patches import Patch
legend_quad = [
    Patch(facecolor="green",      alpha=0.3, label="Eficiente en ataque y defensa"),
    Patch(facecolor="steelblue",  alpha=0.3, label="Falla en ataque, sólido en defensa"),
    Patch(facecolor="darkorange", alpha=0.3, label="Eficiente en ataque, sufre en defensa"),
    Patch(facecolor="darkred",    alpha=0.3, label="Ineficiente en ataque y defensa"),
]
ax.legend(handles=legend_quad, loc="lower right", fontsize=8.5, framealpha=0.85)

ax.set_xlabel("xGDif Ataque  (Goles − xG creado)  →  positivo = anota más de lo esperado")
ax.set_ylabel("xGDif Defensa  (xG rival − GC)  →  positivo = rival falla sus ocasiones")
ax.set_title(
    "G6 · Cuadrantes Tácticos — Eficiencia en Ambas Fases\n"
    "Tamaño y color = Puntos  ·  ¿Dónde gana realmente cada equipo sus puntos?",
    fontweight="bold",
)
fig.text(
    0.5, -0.02,
    "xGDif Ataque > 0: el equipo anota más goles de los que su xG predice (sobreconvierte). "
    "xGDif Defensa > 0: el rival falla sus ocasiones contra nosotros (infraconvierte).",
    ha="center", fontsize=8.5, color="gray",
)
save(fig, "g6_cuadrantes_xgdif.png")


# ═══════════════════════════════════════════════════════════════
# G7 — BUBBLE: Posesión vs Zona Ofensiva (tamaño = Pts)
# ---------------------------------------------------------------
# Relación oculta: tener el balón ≠ llegar al último tercio.
# Equipos con alta posesión pero baja zona ofensiva juegan en
# bloque medio o usan el balón para agotar al rival, no para
# atacar (posesión estéril). Equipos con baja posesión pero
# alta zona ofensiva son counter-pressing teams.
# ═══════════════════════════════════════════════════════════════
print("[G7] Bubble posesión vs zona ofensiva...")
fig, ax = plt.subplots(figsize=(11, 8))

norm6 = plt.Normalize(df["Pts"].min(), df["Pts"].max())
sizes6 = ((df["Pts"] - df["Pts"].min()) / (df["Pts"].max() - df["Pts"].min()) * 400 + 80)
scatter = ax.scatter(
    df["Pos%"], df["Zona_Ata"],
    s=sizes6,
    c=df["Pts"], cmap="RdYlGn", norm=norm6,
    edgecolors="white", linewidths=0.8,
    alpha=0.88, zorder=3,
)
plt.colorbar(scatter, ax=ax, label="Puntos", shrink=0.85)

# Líneas de mediana como divisor de cuadrantes
med_pos = df["Pos%"].median()
med_ata = df["Zona_Ata"].median()
ax.axvline(med_pos, color="gray", lw=1.2, ls="--", alpha=0.7, zorder=2)
ax.axhline(med_ata, color="gray", lw=1.2, ls="--", alpha=0.7, zorder=2)

# Sombreado de cuadrantes
xlim = ax.get_xlim()
ylim = ax.get_ylim()
ax.fill_between([med_pos, xlim[1]], med_ata, ylim[1], alpha=0.07, color="green",      zorder=0)
ax.fill_between([xlim[0], med_pos], med_ata, ylim[1], alpha=0.07, color="steelblue",  zorder=0)
ax.fill_between([med_pos, xlim[1]], ylim[0], med_ata, alpha=0.07, color="darkorange", zorder=0)
ax.fill_between([xlim[0], med_pos], ylim[0], med_ata, alpha=0.07, color="darkred",    zorder=0)
ax.set_xlim(xlim)
ax.set_ylim(ylim)

# Etiquetas con adjustText
texts6 = []
for xi, yi, lbl in zip(df["Pos%"], df["Zona_Ata"], df["Equipo"]):
    texts6.append(ax.text(xi, yi, lbl, fontsize=8.5, color="#1a1a2e"))
adjust_text(
    texts6, x=df["Pos%"].values, y=df["Zona_Ata"].values,
    arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5),
    ax=ax,
)

# Leyenda combinada: cuadrantes + tamaño de burbuja
from matplotlib.patches import Patch
legend_quad6 = [
    Patch(facecolor="green",      alpha=0.3, label="Alta posesión + alta zona ofensiva (dominadores)"),
    Patch(facecolor="steelblue",  alpha=0.3, label="Baja posesión + alta zona ofensiva (counter-press)"),
    Patch(facecolor="darkorange", alpha=0.3, label="Alta posesión + baja zona ofensiva (posesión estéril)"),
    Patch(facecolor="darkred",    alpha=0.3, label="Baja posesión + baja zona ofensiva (bloque bajo)"),
]
ax.legend(handles=legend_quad6, loc="lower right", fontsize=8, framealpha=0.85)

ax.set_xlabel("Posesión media (%)")
ax.set_ylabel("% de acciones en Zona Ofensiva (1/3 final)")
ax.set_title(
    "G7 · Estilo de Juego: Posesión vs Dominio Territorial\n"
    "Tamaño y color de burbuja = Puntos · Divisor = mediana",
    fontweight="bold",
)
fig.text(
    0.5, -0.02,
    "Las líneas divisorias corresponden a la mediana de posesión y zona ofensiva de la liga. "
    "Ninguna zona predice Pts de forma significativa (ver G3): el estilo no garantiza resultados.",
    ha="center", fontsize=8.5, color="gray",
)
save(fig, "g7_bubble_posesion_zona.png")


# ═══════════════════════════════════════════════════════════════
# G10 — VALIDACIÓN: Índice de Éxito (IS) vs Puntos reales
# ---------------------------------------------------------------
# Demuestra que el IS predice los puntos reales con alta
# significancia estadística, justificando su uso como eje
# del algoritmo prescriptivo.
# ═══════════════════════════════════════════════════════════════
print("\n[G10] Validación IS vs Pts...")

fig, ax = plt.subplots(figsize=(11, 8))

norm7is = plt.Normalize(df["IS"].min(), df["IS"].max())
sc7is = ax.scatter(
    df["IS"], df["Pts"],
    c=df["IS"], cmap="RdYlGn", norm=norm7is,
    s=130, edgecolors="white", linewidths=0.8, zorder=3,
)
plt.colorbar(sc7is, ax=ax, label="Índice de Éxito (IS)", shrink=0.85)

# Línea de regresión + estadísticos
slope7, intercept7, r7, p7, _ = stats.linregress(df["IS"], df["Pts"])
x7 = np.linspace(df["IS"].min() - 0.02, df["IS"].max() + 0.02, 100)
ax.plot(x7, slope7 * x7 + intercept7,
        color="#2c3e50", lw=2, ls="--",
        label=f"Regresión lineal  r = {r7:.3f}  p < 0.001")

# Etiquetas con adjustText
texts7is = []
for xi, yi, lbl in zip(df["IS"], df["Pts"], df["Equipo"]):
    texts7is.append(ax.text(xi, yi, lbl, fontsize=8.5, color="#1a1a2e"))
adjust_text(
    texts7is, x=df["IS"].values, y=df["Pts"].values,
    arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5),
    ax=ax,
)

ax.set_xlabel("Índice de Éxito  (IS = 0.35·xG/pp + 0.35·(−GC/pp) + 0.30·Rating,  escala Min-Max)")
ax.set_ylabel("Puntos reales en la clasificación")
ax.set_title(
    "G10 · Validación del Índice de Éxito\n"
    "¿Predice el IS los puntos reales?",
    fontweight="bold",
)
ax.legend(loc="upper left", framealpha=0.9)
ax.grid(alpha=0.3)
ax.set_axisbelow(True)
fig.text(
    0.5, -0.02,
    "El IS explica el " + f"{r7**2*100:.1f}% de la varianza en puntos (R² = {r7**2:.3f}). "
    "Prescripción: IS alto → ajustes finos · IS bajo → revisión táctica profunda.",
    ha="center", fontsize=9, color="gray",
)
save(fig, "g10_validacion_is.png")


# ═══════════════════════════════════════════════════════════════
# G11 — CLIMA: IS y rendimiento en partidos de lluvia
# ---------------------------------------------------------------
# Hipótesis prescriptiva: los equipos "habituados" (>100 días
# de lluvia en su ciudad) están acostumbrados a campos pesados,
# lluvia lateral y viento. Su IS debería mantenerse o mejorar
# en condiciones lluviosas, mientras que los "no habituados"
# deberían decrecer. Detectar esta diferencia permite al
# algoritmo recomendar ajustes tácticos (más directo, más
# contacto físico) cuando se juega en estadios lluviosos.
# ═══════════════════════════════════════════════════════════════
print("[G11] Análisis climático...")

# ── Clasificar equipos por habituación ─────────────────────────
df["Habituacion"] = df["Dias_lluvia"].apply(
    lambda d: "Habituado (>100 días)" if d > 100 else "No habituado (≤100 días)"
)

# ── Parsear resultados en partidos de lluvia ────────────────────
raw_par = pd.read_excel(EXCEL, sheet_name="Partidos", header=None)
par = pd.DataFrame({
    "Local":      raw_par.iloc[2:, 4].values,
    "Resultado":  raw_par.iloc[2:, 5].values,
    "Visitante":  raw_par.iloc[2:, 6].values,
    "Lluvia":     raw_par.iloc[2:, 8].values,
}).dropna(subset=["Resultado"]).reset_index(drop=True)

# Filtrar partidos con lluvia real (excluir "No Llovió" y "Estadio Cubierto")
par_lluvia = par[
    (par["Lluvia"] != "No Llovió") & (par["Lluvia"] != "Estadio Cubierto")
].copy()

# Normalización de nombres de equipos en Partidos → nombres en df
TEAM_NAMES_PAR = {
    "FC Barcelona": "Barcelona", "Barcelona": "Barcelona",
    "Atletico De Madrid": "Atlético Madrid", "Atletico Madrid": "Atlético Madrid",
    "Atlético Madrid": "Atlético Madrid",
    "CA Osasuna": "Osasuna", "Osasuna": "Osasuna",
    "Deportivo Alaves": "Alavés", "Alaves": "Alavés", "Alavés": "Alavés",
    "Elche CF": "Elche", "Elche": "Elche",
    "Getafe CF": "Getafe", "Getafe": "Getafe",
    "Girona FC": "Girona", "Girona": "Girona",
    "Levante UD": "Levante", "Levante": "Levante",
    "RC Celta": "Celta Vigo", "Celta Vigo": "Celta Vigo",
    "RCD Espanyol De Barcelona": "Espanyol", "Espanyol": "Espanyol",
    "RCD Mallorca": "Mallorca", "Mallorca": "Mallorca",
    "Rayo Vallecano": "Rayo Vallecano",
    "Real Betis": "Real Betis",
    "Real Madrid": "Real Madrid",
    "Real Oviedo SAD": "Oviedo", "Oviedo": "Oviedo",
    "Real Sociedad": "Real Sociedad",
    "Sevilla FC": "Sevilla", "Sevilla": "Sevilla",
    "Valencia CF": "Valencia", "Valencia": "Valencia",
    "Villarreal CF": "Villarreal", "Villarreal": "Villarreal",
    "Athletic Club": "Athletic Club",
}

def parse_pts(resultado, local=True):
    """'2–1' → 3 pts local / 0 pts visitante. Maneja '–' y '-'."""
    res = str(resultado).replace("–", "-").strip()
    try:
        g_local, g_vis = map(int, res.split("-"))
    except ValueError:
        return np.nan
    if local:
        return 3 if g_local > g_vis else (1 if g_local == g_vis else 0)
    else:
        return 3 if g_vis > g_local else (1 if g_vis == g_local else 0)

records = []
for _, row in par_lluvia.iterrows():
    loc = TEAM_NAMES_PAR.get(str(row["Local"]).strip(), None)
    vis = TEAM_NAMES_PAR.get(str(row["Visitante"]).strip(), None)
    pts_l = parse_pts(row["Resultado"], local=True)
    pts_v = parse_pts(row["Resultado"], local=False)
    if loc and not pd.isna(pts_l):
        records.append({"Equipo": loc, "Pts_lluvia": pts_l})
    if vis and not pd.isna(pts_v):
        records.append({"Equipo": vis, "Pts_lluvia": pts_v})

rain_perf = pd.DataFrame(records).groupby("Equipo")["Pts_lluvia"].agg(
    ["mean", "count"]
).reset_index()
rain_perf.columns = ["Equipo", "Pts_pp_lluvia", "Partidos_lluvia"]

# Merge con IS y habituación
df_g8 = df[["Equipo", "IS", "Pts", "Habituacion", "Dias_lluvia", "Total_mm"]].merge(rain_perf, on="Equipo", how="inner")

# ── Figura: dos paneles ───────────────────────────────────────
fig, (ax_sc, ax_hm) = plt.subplots(1, 2, figsize=(18, 8),
                                    gridspec_kw={"width_ratios": [1.3, 1]})
fig.suptitle(
    "G11 · Clima y Rendimiento — ¿Influye la lluvia en el Índice de Éxito?",
    fontsize=13, fontweight="bold", y=1.01,
)

# ── Panel izquierdo: scatter IS vs Pts/pp en lluvia ───────────
palette_g8 = {"Habituado (>100 días)": "#2980b9", "No habituado (≤100 días)": "#e67e22"}
for grupo, gdata in df_g8.groupby("Habituacion"):
    ax_sc.scatter(
        gdata["IS"], gdata["Pts_pp_lluvia"],
        label=f"{grupo}  (n={len(gdata)}, media={gdata['Pts_pp_lluvia'].mean():.2f} pts/pp)",
        color=palette_g8[grupo],
        s=120 + gdata["Dias_lluvia"] * 0.4,
        edgecolors="white", linewidths=0.8, zorder=3,
    )

liga_pts_llv = df_g8["Pts_pp_lluvia"].mean()
ax_sc.axhline(liga_pts_llv, color="#555", lw=1.5, ls="--", alpha=0.7,
              label=f"Media liga en lluvia: {liga_pts_llv:.2f} pts/pp")

# Test t entre grupos
hab   = df_g8[df_g8["Habituacion"] == "Habituado (>100 días)"]["Pts_pp_lluvia"]
nohab = df_g8[df_g8["Habituacion"] == "No habituado (≤100 días)"]["Pts_pp_lluvia"]
_, p_g8 = stats.ttest_ind(hab, nohab)
ax_sc.text(
    0.98, 0.04,
    f"Test t entre grupos: p = {p_g8:.3f} (NS)",
    transform=ax_sc.transAxes, ha="right", fontsize=9, color="#c0392b",
    bbox=dict(boxstyle="round,pad=0.4", facecolor="white", edgecolor="#cccccc"),
)

texts8 = []
for xi, yi, lbl in zip(df_g8["IS"], df_g8["Pts_pp_lluvia"], df_g8["Equipo"]):
    texts8.append(ax_sc.text(xi, yi, lbl, fontsize=8.5, color="#1a1a2e"))
adjust_text(
    texts8, x=df_g8["IS"].values, y=df_g8["Pts_pp_lluvia"].values,
    arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5),
    ax=ax_sc,
)

ax_sc.set_xlabel("Índice de Éxito (IS) — temporada completa")
ax_sc.set_ylabel("Puntos por partido en condiciones de lluvia")
ax_sc.set_title("IS vs Rendimiento en Lluvia\nColor = habituación · Tamaño = días de lluvia",
                fontweight="bold")
ax_sc.legend(loc="upper left", framealpha=0.9, fontsize=8.5)
ax_sc.grid(alpha=0.3)
ax_sc.set_axisbelow(True)

# ── Panel derecho: heatmap correlaciones clima vs IS/Pts ──────
corr_cols_g8 = {
    "Días lluvia":      "Dias_lluvia",
    "Total mm lluvia":  "Total_mm",
    "Pts/pp lluvia":    "Pts_pp_lluvia",
    "IS":               "IS",
    "Pts":              "Pts",
}
df_corr_g8 = df_g8[[v for v in corr_cols_g8.values()]].rename(
    columns={v: k for k, v in corr_cols_g8.items()}
)
corr_g8 = df_corr_g8.corr()

# p-valores
pval_g8 = pd.DataFrame(np.ones_like(corr_g8), index=corr_g8.index, columns=corr_g8.columns)
for c1, c2 in combinations(corr_g8.columns, 2):
    xy = df_corr_g8[[c1, c2]].dropna()
    _, p = stats.pearsonr(xy.iloc[:, 0], xy.iloc[:, 1])
    pval_g8.loc[c1, c2] = p
    pval_g8.loc[c2, c1] = p

annot_g8 = pd.DataFrame("", index=corr_g8.index, columns=corr_g8.columns)
for c1 in corr_g8.columns:
    for c2 in corr_g8.columns:
        if c1 == c2:
            continue
        r = corr_g8.loc[c1, c2]
        p = pval_g8.loc[c1, c2]
        stars = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "NS"
        annot_g8.loc[c1, c2] = f"{r:.2f}\n{stars}"

mask_g8 = np.triu(np.ones_like(corr_g8, dtype=bool))
sns.heatmap(
    corr_g8, mask=mask_g8,
    annot=annot_g8, fmt="",
    annot_kws={"size": 9, "va": "center"},
    cmap="RdYlGn", vmin=-1, vmax=1,
    linewidths=0.4, linecolor="white",
    cbar_kws={"shrink": 0.75, "label": "Pearson r"},
    ax=ax_hm,
)
ax_hm.set_title("Correlaciones: Variables Climáticas vs IS y Pts\n¿Predice la lluvia el rendimiento?",
                fontweight="bold")
ax_hm.set_xticklabels(ax_hm.get_xticklabels(), rotation=35, ha="right")
ax_hm.set_yticklabels(ax_hm.get_yticklabels(), rotation=0)

fig.text(
    0.5, -0.02,
    "Verde = correlación positiva · Rojo = negativa · "
    "*** p<0.001  ** p<0.01  * p<0.05  NS = no significativo  ·  n = 20 equipos",
    ha="center", fontsize=9, color="gray",
)
plt.tight_layout()
save(fig, "g11_clima_lluvia_is.png")


# ═══════════════════════════════════════════════════════════════
# G12 — ÁRBITROS: Ratio F/T medio por equipo vs Rating
# ---------------------------------------------------------------
# F/T (Faltas por Tarjeta) indica la exigencia del árbitro:
# F/T bajo = árbitro estricto (pocas faltas antes de la tarjeta).
# Prescripción: si el algoritmo detecta que a un equipo top le
# asignan sistemáticamente árbitros estrictos, recomendará
# reducir la agresividad táctica para evitar inferioridad
# numérica en fases decisivas del partido.
# ═══════════════════════════════════════════════════════════════
print("[G12] Análisis arbitral...")

raw_arb = pd.read_excel(EXCEL, sheet_name="Árbitros", header=None)

# ── Árbitros GENERAL: FaltasPP y AmarillasPP ───────────────────
arb_gen = pd.DataFrame({
    "Arbitro_full": raw_arb.iloc[2:, 0].values,
    "Partidos_arb": pd.to_numeric(raw_arb.iloc[2:, 2], errors="coerce"),
    "FaltasPP":     pd.to_numeric(raw_arb.iloc[2:, 4], errors="coerce"),
    "AmarillasPP":  pd.to_numeric(raw_arb.iloc[2:, 6], errors="coerce"),
})
arb_gen["Arbitro_full"] = arb_gen["Arbitro_full"].ffill()
arb_gen = arb_gen[arb_gen["Partidos_arb"].notna()].copy()
# F/T = FaltasPP / AmarillasPP  (a nivel de temporada)
arb_gen["FT_temporada"] = arb_gen["FaltasPP"] / arb_gen["AmarillasPP"]
# Clave de match: primer apellido + primera letra nombre
arb_gen["key"] = arb_gen["Arbitro_full"].apply(
    lambda x: str(x).split()[0].lower()[:5] if pd.notna(x) else ""
)

# ── Árbitros JORNADAS ───────────────────────────────────────────
jorn_raw = raw_arb.iloc[2:, 17:27].copy()
jorn_raw.columns = ["Arbitro","Jornada","Local","Visitante",
                    "FaltasL","FaltasV","TarjL","TarjV","FT_L","FT_V"]
jorn_raw["Arbitro"] = jorn_raw["Arbitro"].ffill()
jorn_raw = jorn_raw[jorn_raw["Jornada"].notna()].reset_index(drop=True)
# Convertir FT a numérico (hay '-' cuando no hubo tarjetas)
jorn_raw["FT_L"] = pd.to_numeric(jorn_raw["FT_L"], errors="coerce")
jorn_raw["FT_V"] = pd.to_numeric(jorn_raw["FT_V"], errors="coerce")

# Normalizar nombre árbitro → clave para match con GENERAL
jorn_raw["key"] = jorn_raw["Arbitro"].apply(
    lambda x: str(x).split()[0].lower()[:5] if pd.notna(x) else ""
)

# Join JORNADAS con FT_temporada de GENERAL
jorn = jorn_raw.merge(arb_gen[["key", "FT_temporada", "AmarillasPP"]], on="key", how="left")

# Para cada equipo: calcular F/T medio al que se enfrenta
# (tanto cuando juega de local como de visitante)
records_arb = []
for _, row in jorn.iterrows():
    loc = TEAM_NAMES_PAR.get(str(row["Local"]).strip(), None)
    vis = TEAM_NAMES_PAR.get(str(row["Visitante"]).strip(), None)
    ft_val = row["FT_temporada"]
    am_val = row["AmarillasPP"]
    if loc and not pd.isna(ft_val):
        records_arb.append({"Equipo": loc, "FT_arb": ft_val, "Amarillas_arb": am_val})
    if vis and not pd.isna(ft_val):
        records_arb.append({"Equipo": vis, "FT_arb": ft_val, "Amarillas_arb": am_val})

arb_per_team = (
    pd.DataFrame(records_arb)
    .groupby("Equipo")[["FT_arb", "Amarillas_arb"]]
    .mean()
    .reset_index()
)

df_g9 = df[["Equipo", "Rating", "Pts", "IS"]].merge(arb_per_team, on="Equipo", how="inner")

# ── Scatter: F/T medio vs Rating (color y tamaño = Pts) ─────────
fig, ax = plt.subplots(figsize=(11, 8))

norm9 = plt.Normalize(df_g9["Pts"].min(), df_g9["Pts"].max())
sizes9 = ((df_g9["Pts"] - df_g9["Pts"].min()) / (df_g9["Pts"].max() - df_g9["Pts"].min()) * 400 + 80)
sc9 = ax.scatter(
    df_g9["FT_arb"], df_g9["Rating"],
    c=df_g9["Pts"], cmap="RdYlGn", norm=norm9,
    s=sizes9, edgecolors="white", linewidths=0.8, zorder=3,
)
plt.colorbar(sc9, ax=ax, label="Puntos en clasificación", shrink=0.85)

# Regresión
slope9, intercept9, r9, p9, _ = stats.linregress(df_g9["FT_arb"], df_g9["Rating"])
x9 = np.linspace(df_g9["FT_arb"].min() - 0.1, df_g9["FT_arb"].max() + 0.1, 100)
ax.plot(x9, slope9 * x9 + intercept9, color="#2c3e50", lw=2, ls="--",
        label=f"Regresión  r = {r9:.2f}  p = {p9:.3f} (NS)")

# Medianas como divisor de cuadrantes
med_ft9     = df_g9["FT_arb"].median()
med_rating9 = df_g9["Rating"].median()
ax.axvline(med_ft9,     color="gray", lw=1.2, ls="--", alpha=0.7, zorder=2)
ax.axhline(med_rating9, color="gray", lw=1.2, ls="--", alpha=0.7, zorder=2)

# Sombreado de cuadrantes
xlim9 = ax.get_xlim()
ylim9 = ax.get_ylim()
ax.fill_between([xlim9[0], med_ft9], med_rating9, ylim9[1], alpha=0.07, color="green",      zorder=0)
ax.fill_between([med_ft9, xlim9[1]], med_rating9, ylim9[1], alpha=0.07, color="steelblue",  zorder=0)
ax.fill_between([xlim9[0], med_ft9], ylim9[0], med_rating9, alpha=0.07, color="darkorange", zorder=0)
ax.fill_between([med_ft9, xlim9[1]], ylim9[0], med_rating9, alpha=0.07, color="darkred",    zorder=0)
ax.set_xlim(xlim9)
ax.set_ylim(ylim9)

# Etiquetas con adjustText
texts9 = []
for xi, yi, lbl in zip(df_g9["FT_arb"], df_g9["Rating"], df_g9["Equipo"]):
    texts9.append(ax.text(xi, yi, lbl, fontsize=8.5, color="#1a1a2e"))
adjust_text(
    texts9, x=df_g9["FT_arb"].values, y=df_g9["Rating"].values,
    arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5),
    ax=ax,
)

# Leyenda cuadrantes
legend_q9 = [
    Patch(facecolor="green",      alpha=0.3, label="Árbitro estricto + Rating alto"),
    Patch(facecolor="steelblue",  alpha=0.3, label="Árbitro leniente + Rating alto"),
    Patch(facecolor="darkorange", alpha=0.3, label="Árbitro estricto + Rating bajo"),
    Patch(facecolor="darkred",    alpha=0.3, label="Árbitro leniente + Rating bajo"),
]
ax.legend(handles=[ax.get_lines()[0]] + legend_q9, loc="upper right", fontsize=8.5, framealpha=0.9)

ax.set_xlabel("F/T medio de los árbitros asignados  (Faltas/Tarjeta)  ← estricto | leniente →")
ax.set_ylabel("Rating General del equipo")
ax.set_title(
    "G12 · Factor Arbitral: Exigencia del Árbitro vs Rendimiento\n"
    "¿Los equipos top reciben árbitros más estrictos?",
    fontweight="bold",
)
ax.grid(alpha=0.3)
ax.set_axisbelow(True)
fig.text(
    0.5, -0.02,
    "F/T (Faltas/Tarjeta): valor bajo = árbitro estricto (saca tarjeta con pocas faltas). "
    "Divisor = mediana de liga. Tamaño y color = Puntos.",
    ha="center", fontsize=9, color="gray",
)
save(fig, "g12_scatter_arbitros.png")


# ═══════════════════════════════════════════════════════════════
# G8 — ZONAS DE ACCIÓN: Barra Apilada por Equipo
# ---------------------------------------------------------------
# Los porcentajes de acción en cada tercio del campo revelan
# el "bloque" táctico del equipo: alto (>35% zona ataque),
# medio o bajo (<10% zona defensa suele indicar line alta).
# Prescripción directa: el algoritmo recomendará la altura de
# línea defensiva óptima comparando el perfil de zonas del
# equipo con los patrones de los equipos de éxito similares.
# ═══════════════════════════════════════════════════════════════
print("[G8] Zonas de acción (barras apiladas)...")

# Ordenar por IS descendente (equipos más exitosos arriba)
zonas = df[["Equipo", "Zona_Def", "Zona_Med", "Zona_Ata", "IS", "Pts"]].copy()
zonas = zonas.sort_values("IS", ascending=True).reset_index(drop=True)

# Normalizar a 100%
zona_sum = zonas[["Zona_Def", "Zona_Med", "Zona_Ata"]].sum(axis=1)
zonas["Def_n"] = zonas["Zona_Def"] / zona_sum * 100
zonas["Med_n"] = zonas["Zona_Med"] / zona_sum * 100
zonas["Ata_n"] = zonas["Zona_Ata"] / zona_sum * 100

fig, ax = plt.subplots(figsize=(14, 10))

y_pos = np.arange(len(zonas))
bar_h = 0.68
C_DEF = "#3498db"
C_MED = "#f39c12"
C_ATA = "#e74c3c"

# Fondo alternado para mejorar legibilidad
for i in range(len(zonas)):
    if i % 2 == 0:
        ax.axhspan(i - 0.5, i + 0.5, color="#f8f8f8", zorder=0)

ax.barh(y_pos, zonas["Def_n"], height=bar_h, color=C_DEF, label="1/3 Defensivo", zorder=2)
ax.barh(y_pos, zonas["Med_n"], height=bar_h, left=zonas["Def_n"], color=C_MED, label="1/3 Medio", zorder=2)
ax.barh(y_pos, zonas["Ata_n"], height=bar_h, left=zonas["Def_n"]+zonas["Med_n"], color=C_ATA, label="1/3 Ofensivo", zorder=2)

# Etiquetas de porcentaje siempre dentro de la barra
for i, (_, row) in enumerate(zonas.iterrows()):
    for val, left in [
        (row["Def_n"], 0),
        (row["Med_n"], row["Def_n"]),
        (row["Ata_n"], row["Def_n"]+row["Med_n"]),
    ]:
        ax.text(left + val / 2, i, f"{val:.0f}%",
                ha="center", va="center", fontsize=6.5, color="white",
                fontweight="bold", zorder=3)

# IS al margen derecho
for i, (_, row) in enumerate(zonas.iterrows()):
    ax.text(101.5, i, f"IS={row['IS']:.2f}", va="center", fontsize=8.5, color="#1a1a2e")

# Línea de referencia 33.3%
ax.axvline(33.3, color="gray", lw=1, ls=":", alpha=0.5, zorder=1)

ax.set_yticks(y_pos)
ax.set_yticklabels(zonas["Equipo"], fontsize=10.5)
ax.set_xlim(0, 112)
ax.set_xlabel("Distribución porcentual de acciones por tercio del campo")
ax.set_title(
    "G8 · Zonas de Acción por Equipo — Perfil de Bloque Táctico\n"
    "Ordenado de menor a mayor IS · Margen derecho = Índice de Éxito",
    fontweight="bold",
)
ax.legend(loc="upper left", bbox_to_anchor=(1.01, 1), borderaxespad=0, framealpha=0.9)
ax.grid(axis="x", alpha=0.3)
ax.set_axisbelow(True)
fig.text(
    0.5, -0.02,
    "Prescripción: equipos con >35% en zona ataque → bloque alto. "
    "Equipos con >12% en zona defensa → bloque bajo o medio-bajo. "
    "Línea vertical gris = distribución equilibrada (33.3% cada tercio).",
    ha="center", fontsize=9, color="gray",
)
save(fig, "g8_zonas_accion_apiladas.png")


# ═══════════════════════════════════════════════════════════════
# DATOS JUGADORES (compartido por G13 y G15)
# Filtro: >= 400 minutos para garantizar representatividad
# ═══════════════════════════════════════════════════════════════
print("[G13 + G15] Cargando datos de jugadores...")

raw_jug = pd.read_excel(EXCEL, sheet_name="Jugadores", header=None)

jug = pd.DataFrame({
    "Jugador":   raw_jug.iloc[3:, 0].values,
    "Posicion":  raw_jug.iloc[3:, 2].values,
    "Equipo":    raw_jug.iloc[3:, 3].values,
    "Minutos":   pd.to_numeric(raw_jug.iloc[3:, 20], errors="coerce"),
    "PctMin":    pd.to_numeric(raw_jug.iloc[3:, 24], errors="coerce"),
    "OnOff":     pd.to_numeric(raw_jug.iloc[3:, 33], errors="coerce"),
    "Rating":    pd.to_numeric(raw_jug.iloc[3:,  5], errors="coerce"),
    "xG_90":     pd.to_numeric(raw_jug.iloc[3:, 139], errors="coerce"),
    "Asist_90":  pd.to_numeric(raw_jug.iloc[3:,  43], errors="coerce"),
    "Entradas":  pd.to_numeric(raw_jug.iloc[3:, 235], errors="coerce"),
    "Interc":    pd.to_numeric(raw_jug.iloc[3:, 241], errors="coerce"),
    "Despejes":  pd.to_numeric(raw_jug.iloc[3:, 244], errors="coerce"),
    "Bloqueos":  pd.to_numeric(raw_jug.iloc[3:, 247], errors="coerce"),
}).dropna(subset=["Minutos", "Equipo", "PctMin"])
jug = jug[jug["Minutos"] >= 400].reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════
# G13 — APORTACIÓN INDIVIDUAL AL ÍNDICE DE ÉXITO
# ---------------------------------------------------------------
# IS_individual: proxy del IS con datos de jugador.
#   IS_indiv = 0.35·norm(xG/90) + 0.35·norm(DefScore/90) + 0.30·norm(Rating)
#   DefScore/90 = (Entradas + Interc + Despejes + Bloqueos) / (Min/90)
# Contribución al equipo: IS_indiv × (Minutos / min_totales_equipo)
#
# Panel izquierdo — jugador que más aporta al IS de su equipo.
# Panel derecho — concentración: suma top-3 contribuciones por equipo.
# ═══════════════════════════════════════════════════════════════
print("[G13] Aportación individual al IS...")

# ── Construir IS individual ──────────────────────────────────────
jug_is = jug.copy()

# ── Grupo posicional (para normalizar dentro de cada rol) ────────
_PORTERO   = {"GK", "P", "PT", "POR", "GOR"}
_DEFENSA   = {"DC", "DFC", "DL", "DR", "LD", "LI", "CB", "LB", "RB", "DF"}
_DELANTERO = {"FC", "FW", "SS", "ST", "CF", "EI", "ED", "EC", "DEL", "AT"}

def _pos_group(pos_str):
    if pd.isna(pos_str) or str(pos_str).strip() == "":
        return "Centrocampista"
    p = str(pos_str).split(",")[0].strip().upper()
    if p in _PORTERO:
        return "Portero"
    if p in _DELANTERO or p.startswith("F") or p.startswith("SS"):
        return "Delantero"
    if p.startswith("E"):   # Extremo izq/der → atacante
        return "Delantero"
    if p in _DEFENSA or p.startswith("L"):
        return "Defensa"
    return "Centrocampista"

jug_is["PosGrupo"] = jug_is["Posicion"].apply(_pos_group)

# ── Métricas brutas ──────────────────────────────────────────────
jug_is["OfScore90"] = (jug_is["xG_90"].fillna(0)
                       + jug_is["Asist_90"].fillna(0))   # xG + asist por 90
jug_is["DefScore90"] = (
    jug_is[["Entradas", "Interc", "Despejes", "Bloqueos"]]
    .fillna(0).sum(axis=1) / (jug_is["Minutos"] / 90)
)

# ── Min-Max dentro del grupo posicional ─────────────────────────
def _mm_group(series, groups):
    """Normaliza [0-1] dentro de cada grupo posicional."""
    result = pd.Series(0.0, index=series.index)
    for g in groups.unique():
        mask = groups == g
        s = series[mask]
        mn, mx = s.min(), s.max()
        result[mask] = (s - mn) / (mx - mn) if mx > mn else 0.5
    return result

rat_filled = jug_is["Rating"].fillna(jug_is.groupby("PosGrupo")["Rating"].transform("median"))
jug_is["norm_of"]  = _mm_group(jug_is["OfScore90"],  jug_is["PosGrupo"])
jug_is["norm_def"] = _mm_group(jug_is["DefScore90"], jug_is["PosGrupo"])
jug_is["norm_rat"] = _mm_group(rat_filled,            jug_is["PosGrupo"])

jug_is["IS_indiv"] = (0.35 * jug_is["norm_of"]
                      + 0.35 * jug_is["norm_def"]
                      + 0.30 * jug_is["norm_rat"])

# Minutos totales por equipo (entre jugadores con ≥ 400 min)
min_eq = jug_is.groupby("Equipo")["Minutos"].sum().rename("Min_equipo")
jug_is = jug_is.merge(min_eq, on="Equipo", how="left")
jug_is["IS_contrib"] = jug_is["IS_indiv"] * (jug_is["Minutos"] / jug_is["Min_equipo"])

# ── Jugador top por equipo, ordenado por IS_contrib ─────────────
top_jug = (
    jug_is.sort_values("IS_contrib", ascending=False)
          .groupby("Equipo", sort=False)
          .first()
          .reset_index()[["Equipo", "Jugador", "IS_contrib", "PctMin", "IS_indiv"]]
)
top_jug = top_jug.merge(df[["Equipo", "IS"]], on="Equipo", how="left")
top_jug = top_jug.sort_values("IS_contrib", ascending=True).reset_index(drop=True)

# ── Concentración top-3 contribuciones por equipo ───────────────
conc11 = []
for equipo, grp in jug_is.groupby("Equipo"):
    top3 = grp.nlargest(3, "IS_contrib")
    apellidos = " / ".join([n.split()[-1] for n in top3["Jugador"].tolist()])
    conc11.append({
        "Equipo":    equipo,
        "Conc_IS":   top3["IS_contrib"].sum(),
        "Jugadores": apellidos,
    })
df_conc11 = pd.DataFrame(conc11)
df_conc11 = df_conc11.merge(df[["Equipo", "IS"]], on="Equipo", how="left")
df_conc11 = df_conc11.sort_values("Conc_IS", ascending=True).reset_index(drop=True)

# ── Figura doble ─────────────────────────────────────────────────
fig, (ax_pil, ax_conc) = plt.subplots(1, 2, figsize=(18, 10),
                                       gridspec_kw={"width_ratios": [1.4, 1]})

# ── Panel izquierdo: jugador que más aporta al IS ────────────────
cmap_pil   = plt.cm.RdYlGn
norm_pil   = plt.Normalize(top_jug["IS_contrib"].min(), top_jug["IS_contrib"].max())
colors_pil = cmap_pil(norm_pil(top_jug["IS_contrib"]))

for i in range(len(top_jug)):
    if i % 2 == 0:
        ax_pil.axhspan(i - 0.5, i + 0.5, color="#f8f8f8", zorder=0)

bars_pil = ax_pil.barh(
    range(len(top_jug)), top_jug["IS_contrib"],
    color=colors_pil, height=0.65, edgecolor="white", zorder=2,
)
ax_pil.set_yticks(range(len(top_jug)))
ax_pil.set_yticklabels(top_jug["Equipo"], fontsize=10)

for bar, (_, row), rgba in zip(bars_pil, top_jug.iterrows(), colors_pil):
    lum = 0.299*rgba[0] + 0.587*rgba[1] + 0.114*rgba[2]
    txt_color = "#1a1a2e" if lum > 0.55 else "white"
    lbl = f"{row['Jugador']}  ({row['PctMin']:.0f}% min)"
    ax_pil.text(
        bar.get_width() * 0.03, bar.get_y() + bar.get_height() / 2,
        lbl, va="center", fontsize=8.5, color=txt_color, fontweight="bold", zorder=3,
    )
    ax_pil.text(
        bar.get_width() + 0.001, bar.get_y() + bar.get_height() / 2,
        f"{row['IS_contrib']:.3f}", va="center", fontsize=8.5, color="#1a1a2e",
    )

media_pil = top_jug["IS_contrib"].mean()
ax_pil.axvline(media_pil, color="#2c3e50", lw=1.8, ls="--",
               label=f"Media liga: {media_pil:.3f}")
ax_pil.set_xlabel("Contribución al IS del equipo  (IS_individual × proporción de minutos)")
ax_pil.set_title("Jugador que más aporta al IS\n(mayor contribución con ≥ 400 min)", fontweight="bold")
ax_pil.legend(loc="lower right", fontsize=9)
ax_pil.set_xlim(0, top_jug["IS_contrib"].max() * 1.18)
ax_pil.grid(axis="x", alpha=0.3)
ax_pil.set_axisbelow(True)

# ── Panel derecho: concentración de aportación al IS ────────────
norm_conc   = plt.Normalize(df_conc11["Conc_IS"].min(), df_conc11["Conc_IS"].max())
colors_conc = plt.cm.YlOrRd(norm_conc(df_conc11["Conc_IS"]))

for i in range(len(df_conc11)):
    if i % 2 == 0:
        ax_conc.axhspan(i - 0.5, i + 0.5, color="#f8f8f8", zorder=0)

bars_conc = ax_conc.barh(
    range(len(df_conc11)), df_conc11["Conc_IS"],
    color=colors_conc, height=0.65, edgecolor="white", zorder=2,
)
ax_conc.set_yticks(range(len(df_conc11)))
ax_conc.set_yticklabels(df_conc11["Equipo"], fontsize=10)

for bar, (_, row), rgba in zip(bars_conc, df_conc11.iterrows(), colors_conc):
    lum = 0.299*rgba[0] + 0.587*rgba[1] + 0.114*rgba[2]
    txt_color = "#1a1a2e" if lum > 0.55 else "white"
    ax_conc.text(
        bar.get_width() * 0.03, bar.get_y() + bar.get_height() / 2,
        row["Jugadores"], va="center", fontsize=7.5, color=txt_color,
        fontweight="bold", zorder=3,
    )
    ax_conc.text(
        bar.get_width() + 0.003, bar.get_y() + bar.get_height() / 2,
        f"{row['Conc_IS']:.3f}", va="center", fontsize=8.5, color="#1a1a2e",
    )

media_conc = df_conc11["Conc_IS"].mean()
ax_conc.axvline(media_conc, color="#8e44ad", lw=1.8, ls="--",
                label=f"Media liga: {media_conc:.3f}")
ax_conc.set_xlabel("Suma IS-contribución top-3 (concentración de aportación)")
ax_conc.set_title("Concentración de Aportación al IS\n(suma top-3 contribuciones por equipo)", fontweight="bold")
ax_conc.legend(loc="lower right", fontsize=9)
ax_conc.set_xlim(0, df_conc11["Conc_IS"].max() * 1.18)
ax_conc.grid(axis="x", alpha=0.3)
ax_conc.set_axisbelow(True)

fig.suptitle(
    "G13 · Aportación Individual al Índice de Éxito — LaLiga 2025-26\n"
    "IS_individual = 0.35·(xG+Ast)/90 + 0.35·Def/90 + 0.30·Rating  ·  Norm. por grupo posicional  ·  Ponderado por minutos",
    fontweight="bold", fontsize=13,
)
fig.text(
    0.5, -0.01,
    "Contribución = IS_individual × (minutos / minutos_totales_equipo). "
    "Concentración alta → el IS del equipo depende de pocos jugadores. "
    "Concentración baja → aportación repartida.",
    ha="center", fontsize=9, color="gray",
)
plt.tight_layout(rect=[0, 0.02, 1, 0.96])
save(fig, "g13_pilares_concentracion.png")


# ═══════════════════════════════════════════════════════════════
# G14 — DUPLAS PELIGROSAS: todas las duplas agrupadas por equipo
# ---------------------------------------------------------------
# Cada dupla es una fila independiente. Los equipos aparecen como
# cabecera (fondo azul) con todas sus duplas debajo, ordenadas
# por frecuencia descendente. Los equipos se ordenan por IS.
# El color de cada barra refleja el % que esa dupla representa
# sobre los goles totales del equipo (verde=baja dependencia,
# rojo=alta dependencia).
# ═══════════════════════════════════════════════════════════════
print("[G14] Duplas peligrosas (todas, por equipo)...")

raw_dup = pd.read_excel(EXCEL, sheet_name="Duplas Peligrosas", header=None)
dup = pd.DataFrame({
    "Goleador":   raw_dup.iloc[1:, 1].values,
    "Asistidor":  raw_dup.iloc[1:, 2].values,
    "Equipo":     raw_dup.iloc[1:, 3].values,
    "Frecuencia": pd.to_numeric(raw_dup.iloc[1:, 4], errors="coerce"),
}).dropna(subset=["Frecuencia"]).reset_index(drop=True)

TEAM_MAP_DUP = {
    "Deportivo Alaves": "Alavés", "Alavés": "Alavés",
    "Real Oviedo": "Oviedo", "Real Oviedo SAD": "Oviedo",
    "FC Barcelona": "Barcelona", "Barcelona": "Barcelona",
    "Real Madrid": "Real Madrid",
    "Mallorca": "Mallorca", "RCD Mallorca": "Mallorca",
    "Rayo Vallecano": "Rayo Vallecano",
    "Sevilla": "Sevilla", "Sevilla FC": "Sevilla",
    "Real Betis": "Real Betis",
    "Valencia": "Valencia", "Valencia CF": "Valencia",
    "Celta Vigo": "Celta Vigo", "RC Celta": "Celta Vigo",
}
dup["Equipo"] = dup["Equipo"].map(TEAM_MAP_DUP).fillna(dup["Equipo"])
dup["Dupla"]  = dup["Goleador"] + " ← " + dup["Asistidor"]

goles_totales = df.set_index("Equipo")["Goles"].to_dict()

# ── Construir filas: cabecera de equipo + sus duplas ─────────────
team_order_g12 = df.sort_values("IS", ascending=False)["Equipo"].tolist()

all_rows_g12 = []
for equipo in team_order_g12:
    grp = dup[dup["Equipo"] == equipo].sort_values("Frecuencia", ascending=False)
    if grp.empty:
        continue
    goles_eq = goles_totales.get(equipo, np.nan)
    n_dup = len(grp)
    all_rows_g12.append({
        "type": "header", "equipo": equipo, "n_dup": n_dup,
        "freq": 0.0, "pct": 0.0, "dupla": "",
    })
    for _, r in grp.iterrows():
        pct = float(r["Frecuencia"]) / goles_eq if (not np.isnan(goles_eq) and goles_eq > 0) else 0.0
        all_rows_g12.append({
            "type": "dupla", "equipo": equipo,
            "dupla": r["Dupla"],
            "freq": float(r["Frecuencia"]),
            "pct": pct,
            "goles_eq": goles_eq,
        })

n_rows_g12 = len(all_rows_g12)
max_freq_g12 = max((r["freq"] for r in all_rows_g12 if r["type"] == "dupla"), default=1)

fig_height_g12 = max(14, n_rows_g12 * 0.50 + 2.5)
fig, ax = plt.subplots(figsize=(17, fig_height_g12))

cmap_dup = plt.cm.RdYlGn_r
norm_dup = plt.Normalize(0, 0.30)   # 0 % → 30 % de dependencia

# ── Dibujar fila a fila (y=0 abajo, y=n_rows−1 arriba) ──────────
team_bg_colors = ["#cfe2f3", "#d5eaf7"]  # alternancia de azul por equipo
team_idx = -1

for i, row in enumerate(all_rows_g12):
    y = n_rows_g12 - 1 - i   # primer row (i=0) → arriba

    if row["type"] == "header":
        team_idx += 1
        bg = team_bg_colors[team_idx % 2]
        ax.axhspan(y - 0.48, y + 0.48, color=bg, zorder=0)
        ax.text(
            0.15, y,
            f"{row['equipo']}  —  {row['n_dup']} dupla{'s' if row['n_dup'] > 1 else ''}",
            va="center", fontsize=11, fontweight="bold", color="#1a3a5c", zorder=2,
        )
    else:
        # fondo alternado suave dentro de cada equipo
        if i % 2 == 0:
            ax.axhspan(y - 0.48, y + 0.48, color="#f5f5f5", zorder=0)
        color = cmap_dup(norm_dup(row["pct"]))
        ax.barh(y, row["freq"], color=color, height=0.65,
                edgecolor="white", zorder=2)
        lum = 0.299*color[0] + 0.587*color[1] + 0.114*color[2]
        txt_color = "#1a1a2e" if lum > 0.55 else "white"
        ax.text(
            0.2, y, f"  {row['dupla']}",
            va="center", fontsize=8.5, color=txt_color,
            fontweight="bold", zorder=3,
        )
        ax.text(
            row["freq"] + 0.25, y,
            f"{int(row['freq'])} gol{'es' if row['freq']!=1 else ''}  ({row['pct']:.0%})",
            va="center", fontsize=8, color="#2c3e50",
        )

ax.set_xlim(0, max_freq_g12 + 7)
ax.set_ylim(-0.65, n_rows_g12 - 0.35)
ax.set_yticks([])
ax.set_xlabel("Goles de la dupla en la temporada  ·  (%) = % sobre goles totales del equipo",
              fontsize=10)
ax.grid(axis="x", alpha=0.3)
ax.set_axisbelow(True)

# Colorbar
import matplotlib.ticker as mticker
sm = plt.cm.ScalarMappable(cmap=cmap_dup, norm=norm_dup)
sm.set_array([])
cbar = fig.colorbar(sm, ax=ax, shrink=0.35, pad=0.01, aspect=25)
cbar.set_label("% de los goles del equipo\n(dependencia de la dupla)", fontsize=9)
cbar.ax.yaxis.set_major_formatter(mticker.PercentFormatter(1.0))

equipos_sin_dupla = [e for e in df["Equipo"] if e not in dup["Equipo"].values]
note = ""
if equipos_sin_dupla:
    note = f"Sin duplas registradas: {', '.join(equipos_sin_dupla)}"

fig.suptitle(
    "G14 · Duplas Peligrosas — Todas las Sociedades Goleador-Asistidor\n"
    "Equipos ordenados por IS · Duplas ordenadas por frecuencia · Color = dependencia",
    fontweight="bold", fontsize=13,
)
if note:
    fig.text(0.5, 0.005, note, ha="center", fontsize=8.5, color="gray")

plt.tight_layout(rect=[0, 0.02, 1, 0.96])
save(fig, "g14_duplas_peligrosas.png")


# ═══════════════════════════════════════════════════════════════
# G15 — DÍPTICO: % Minutos vs IS_individual (Mapa de Influencia)
# ---------------------------------------------------------------
# Ejes compartidos:
#   X = % minutos jugados  (uso/confianza del técnico)
#   Y = IS_individual      (calidad: xG/90 + Def/90 + Rating)
# Panel izquierdo  — foco en Q1+Q2 (calidad alta):
#   Q1 → INDISPENSABLE   (alto uso + alta calidad)
#   Q2 → INFRAUTILIZADO  (bajo uso + alta calidad)
# Panel derecho — foco en Q3+Q4 (calidad baja):
#   Q3 → FIJO NEUTRAL    (alto uso + baja calidad) ← vulnerabilidad rival
#   Q4 → PERIFÉRICO      (bajo uso + baja calidad)
# Umbrales: p60 en % minutos, mediana en IS_individual.
# ═══════════════════════════════════════════════════════════════
print("[G15] Díptico % Min vs IS_individual...")

# ── Datos: jug_is calculado en G13 ──────────────────────────────
jug13 = jug_is[["Jugador", "Equipo", "PctMin", "Minutos", "IS_indiv"]].copy()

thr_pct13 = jug13["PctMin"].quantile(0.60)
thr_is13  = jug13["IS_indiv"].median()

def _quad13(row):
    hi_min = row["PctMin"]   >= thr_pct13
    hi_is  = row["IS_indiv"] >= thr_is13
    if hi_min and hi_is:      return "Q1"
    if not hi_min and hi_is:  return "Q2"
    if hi_min and not hi_is:  return "Q3"
    return "Q4"

jug13["Quad"] = jug13.apply(_quad13, axis=1)

QUAD_CFG = {
    "Q1": {"color": "#27ae60", "label": "Indispensable",  "alpha": 0.82, "s": 62},
    "Q2": {"color": "#e67e22", "label": "Infrautilizado", "alpha": 0.82, "s": 62},
    "Q3": {"color": "#2980b9", "label": "Fijo neutral",   "alpha": 0.82, "s": 62},
    "Q4": {"color": "#95a5a6", "label": "Periférico",     "alpha": 0.82, "s": 62},
}
# Alpha rebajada para los cuadrantes no protagonistas en cada panel
ALPHA_DIM = 0.18
SIZE_DIM  = 35

from matplotlib.patches import Patch as Patch13
from matplotlib.lines import Line2D as Line2D13

fig, (ax_L, ax_R) = plt.subplots(1, 2, figsize=(22, 10),
                                  sharey=True, sharex=True)

def _draw_panel(ax, focus_quads, title_detail):
    """Dibuja un panel del díptico. focus_quads = lista de cuadrantes destacados."""

    # Puntos: vivos si están en focus_quads, tenues si no
    for q, cfg in QUAD_CFG.items():
        sub = jug13[jug13["Quad"] == q]
        if q in focus_quads:
            ax.scatter(sub["PctMin"], sub["IS_indiv"],
                       color=cfg["color"], s=cfg["s"], alpha=cfg["alpha"],
                       edgecolors="white", linewidths=0.5, zorder=3)
        else:
            ax.scatter(sub["PctMin"], sub["IS_indiv"],
                       color="#cccccc", s=SIZE_DIM, alpha=ALPHA_DIM,
                       edgecolors="white", linewidths=0.3, zorder=2)

    # Umbrales
    ax.axvline(thr_pct13, color="#2c3e50", lw=1.4, ls="--", alpha=0.55)
    ax.axhline(thr_is13,  color="#8e44ad", lw=1.4, ls="--", alpha=0.55)

    # Sombreado — todos los cuadrantes, extendidos hasta los bordes del gráfico
    ax.autoscale(False)
    xl = ax.get_xlim()
    yl = ax.get_ylim()
    # Usar valores extremos: matplotlib los recorta al borde del eje automáticamente
    BIG = 9999
    shade_map = {
        "Q1": (thr_pct13,  BIG, thr_is13,   BIG, "#27ae60"),
        "Q2": (-BIG, thr_pct13, thr_is13,   BIG, "#e67e22"),
        "Q3": (thr_pct13,  BIG,    -BIG, thr_is13, "#2980b9"),
        "Q4": (-BIG, thr_pct13,   -BIG, thr_is13, "#95a5a6"),
    }
    for q, (x0, x1, y0, y1, col) in shade_map.items():
        ax.fill_between([x0, x1], y0, y1, color=col, alpha=0.08, zorder=0)
    ax.set_xlim(xl)
    ax.set_ylim(yl)

    # Etiquetas de esquina
    corner_map = {
        "Q1": (0.98, 0.95, "INDISPENSABLE",  "right", "top"),
        "Q2": (0.02, 0.95, "INFRAUTILIZADO", "left",  "top"),
        "Q3": (0.98, 0.05, "FIJO NEUTRAL",   "right", "bottom"),
        "Q4": (0.02, 0.05, "PERIFÉRICO",     "left",  "bottom"),
    }
    for q, (cx, cy, lbl, ha, va) in corner_map.items():
        col   = QUAD_CFG[q]["color"]
        alpha = 0.80 if q in focus_quads else 0.18
        fw    = "bold" if q in focus_quads else "normal"
        fs    = 12 if q in focus_quads else 10
        ax.text(cx, cy, lbl, ha=ha, va=va, fontsize=fs, fontweight=fw,
                color=col, alpha=alpha, zorder=1, transform=ax.transAxes)

    # Etiquetas de jugadores con adjustText
    labeled_quads = focus_quads
    sub_labeled = jug13[jug13["Quad"].isin(labeled_quads)]
    texts = []
    for _, row in sub_labeled.iterrows():
        texts.append(ax.text(
            row["PctMin"], row["IS_indiv"], row["Jugador"],
            fontsize=7, color=QUAD_CFG[row["Quad"]]["color"],
            fontweight="bold",
            bbox=dict(boxstyle="round,pad=0.15", facecolor="white",
                      alpha=0.65, edgecolor="none"),
        ))
    if texts:
        adjust_text(
            texts,
            x=sub_labeled["PctMin"].values,
            y=sub_labeled["IS_indiv"].values,
            arrowprops=dict(arrowstyle="-", color="#cccccc", lw=0.6),
            ax=ax,
        )

    ax.set_xlabel("% de Minutos Jugados", fontsize=10)
    ax.set_title(title_detail, fontweight="bold", fontsize=11, pad=8)
    ax.grid(alpha=0.18)
    ax.set_axisbelow(True)

# ── Panel izquierdo: Q1 + Q2 ────────────────────────────────────
_draw_panel(
    ax_L, ["Q1", "Q2"],
    f"Panel A · Jugadores sobre el umbral de calidad\n"
    f"Indispensables (verde) e Infrautilizados (naranja)",
)
ax_L.set_ylabel("IS Individual por posición  ((xG+Ast)/90 + Def/90 + Rating, norm. dentro del rol)", fontsize=10)

# ── Panel derecho: Q3 + Q4 ──────────────────────────────────────
_draw_panel(
    ax_R, ["Q3", "Q4"],
    f"Panel B · Jugadores bajo el umbral de calidad\n"
    f"Fijos neutrales (azul) y Periféricos (gris)",
)

# ── Leyenda compartida en panel derecho ─────────────────────────
legend_handles = [
    Patch13(facecolor="#27ae60", alpha=0.55, label="Indispensable — alto uso + alta calidad"),
    Patch13(facecolor="#e67e22", alpha=0.55, label="Infrautilizado — bajo uso + alta calidad"),
    Patch13(facecolor="#2980b9", alpha=0.55, label="Fijo neutral — alto uso + baja calidad"),
    Patch13(facecolor="#95a5a6", alpha=0.55, label="Periférico — bajo uso + baja calidad"),
    Line2D13([0],[0], color="#2c3e50", lw=1.4, ls="--",
             label=f"Umbral uso (p60 = {thr_pct13:.0f}% min)"),
    Line2D13([0],[0], color="#8e44ad", lw=1.4, ls="--",
             label=f"Umbral calidad (mediana IS = {thr_is13:.2f})"),
]
ax_R.legend(handles=legend_handles, loc="upper left", bbox_to_anchor=(1.01, 1),
            borderaxespad=0, fontsize=9, framealpha=0.92)

fig.suptitle(
    "G15 · Mapa de Influencia Individual — Uso vs Calidad (IS individual)\n"
    "Jugadores con ≥ 400 min  ·  Umbrales: p60 en minutos · mediana en IS",
    fontweight="bold", fontsize=13,
)
fig.text(
    0.5, 0.005,
    "IS individual = 0.35·norm((xG+Ast)/90) + 0.35·norm(Def/90) + 0.30·norm(Rating)  ·  "
    "Normalización dentro del grupo posicional (Portero/Defensa/Centrocampista/Delantero)  ·  "
    "Panel A: Q1 → marcaje prioritario · Q2 → revulsivo  ·  Panel B: Q3 → vulnerabilidad rival",
    ha="center", fontsize=8.5, color="gray",
)
plt.tight_layout(rect=[0, 0.03, 1, 0.96])
save(fig, "g15_scatter_influencia_jugadores.png")


# ═══════════════════════════════════════════════════════════════
# G4 — CORRELACIONES SORPRENDENTES
# ---------------------------------------------------------------
# 8 scatter plots (2×4) con correlaciones contra-intuitivas
# descubiertas entre variables del Diccionario del Excel.
# Color de punto = Pts en la clasificación (rojo→verde).
# ═══════════════════════════════════════════════════════════════
print("[G4] Correlaciones sorprendentes...")

# ── Variables y etiquetas ──────────────────────────────────────
sorpr_cols = {
    "Edad media":        "Edad",
    "Goles marcados":    "Goles",
    "Tarjetas amarillas":"Amarillas",
    "Pases totales":     "Total_pases",
    "Regates exit./pp":  "Reg_exit",
    "Tiros/partido":     "Tiros_pp",
    "Despejes":          "Despejes",
    "% Tiros área":      "Pct_tiro_area",
    "Faltas cometidas":  "Faltas_com",
    "Puntos":            "Pts",
}
# DataFrame limpio (columnas extraídas individualmente para evitar duplicados)
df_sorpr = pd.DataFrame(
    {k: df[v].values for k, v in sorpr_cols.items()}
)
corr_sorpr = df_sorpr.corr()

# ── Matriz de p-valores ────────────────────────────────────────
pval_sorpr = pd.DataFrame(
    np.ones_like(corr_sorpr), index=corr_sorpr.index, columns=corr_sorpr.columns
)
for c1, c2 in combinations(corr_sorpr.columns, 2):
    xy = df_sorpr[[c1, c2]].dropna()
    _, p = stats.pearsonr(xy.iloc[:, 0], xy.iloc[:, 1])
    pval_sorpr.loc[c1, c2] = p
    pval_sorpr.loc[c2, c1] = p

# ── Anotaciones: r + estrellas (mismo formato que G1/G2) ───────
annot_sorpr = pd.DataFrame("", index=corr_sorpr.index, columns=corr_sorpr.columns)
for c1 in corr_sorpr.columns:
    for c2 in corr_sorpr.columns:
        if c1 == c2:
            continue
        r = corr_sorpr.loc[c1, c2]
        p = pval_sorpr.loc[c1, c2]
        stars = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "NS"
        annot_sorpr.loc[c1, c2] = f"{r:.2f}\n{stars}"

mask_sorpr = np.triu(np.ones_like(corr_sorpr, dtype=bool))

fig, ax = plt.subplots(figsize=(13, 10))
sns.heatmap(
    corr_sorpr, mask=mask_sorpr,
    annot=annot_sorpr, fmt="",
    annot_kws={"size": 8.5, "va": "center"},
    cmap="RdYlGn", vmin=-1, vmax=1,
    linewidths=0.4, linecolor="white",
    cbar_kws={"shrink": 0.75, "label": "Pearson r"},
    ax=ax,
)
ax.set_title(
    "G4 · Correlaciones Sorprendentes — Variables del Diccionario\n"
    "Relaciones contra-intuitivas entre métricas tácticas, disciplina y rendimiento",
    fontweight="bold", pad=14,
)
ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
ax.set_yticklabels(ax.get_yticklabels(), rotation=0)

fig.text(
    0.5, -0.02,
    "Verde = correlación positiva · Rojo = negativa · "
    "*** p<0.001  ** p<0.01  * p<0.05  NS = no significativo  ·  n = 20 equipos",
    ha="center", fontsize=9, color="gray",
)
save(fig, "g4_correlaciones_sorprendentes.png")


# ═══════════════════════════════════════════════════════════════
# G9 — RADARES DE PERFIL: Rendimiento y Estilo Táctico
# ---------------------------------------------------------------
# Dos paneles en una figura. Cada grupo se representa como
# la MEDIA del grupo (polígono sólido) + banda min-máx sutil,
# eliminando el solapamiento de líneas individuales.
# Métricas IS (eje oscuro/negrita): xG/pp, Solidez def., Rating
# Métricas adicionales (eje gris): Pases clave/pp, Posesión %
# ═══════════════════════════════════════════════════════════════
print("[G9] Radares de perfil...")
from matplotlib.lines import Line2D

# ── Métricas y normalización ──────────────────────────────────
radar_metrics = {
    "xG/pp":           ("xG_pp",          False),  # IS
    "Solidez\ndefens.":("GC_pp",           True),   # IS — invertido
    "Rating":          ("Rating",          False),  # IS
    "Pases\nclave/pp": ("PasesClave_cort", False),  # adicional
    "Posesión %":      ("Pos%",            False),  # adicional
}
IS_METRICS = {"xG/pp", "Solidez\ndefens.", "Rating"}

radar_df = pd.DataFrame()
radar_df["Equipo"] = df["Equipo"]
radar_df["Pts"]    = df["Pts"]
radar_df["IS"]     = df["IS"]
for label, (col, invert) in radar_metrics.items():
    vals = df[col].copy()
    if invert:
        vals = -vals
    radar_df[label] = minmax(vals)

labels_list   = list(radar_metrics.keys())
N             = len(labels_list)
angles        = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
angles_closed = angles + angles[:1]
median_vals   = [radar_df[m].median() for m in labels_list]
med_closed    = median_vals + median_vals[:1]

# ── Setup de ejes polar ───────────────────────────────────────
def setup_radar(ax):
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_ylim(0, 1)
    ax.set_yticks([0.25, 0.5, 0.75])
    ax.set_yticklabels([])   # eliminamos los labels por defecto
    # Colocamos los valores manualmente pegados al eje superior (ángulo ≈ 0)
    # y ligeramente a la derecha para no solapar con la línea
    for r, lbl in zip([0.25, 0.50, 0.75], ["0.25", "0.50", "0.75"]):
        ax.text(0.08, r, lbl, ha="left", va="center",
                fontsize=6.5, color="#aaaaaa")
    ax.set_xticks(angles)
    ax.set_xticklabels([])
    ax.set_frame_on(False)
    ax.tick_params(length=0)
    ax.grid(False)
    # Anillos
    for r in [0.25, 0.5, 0.75, 1.0]:
        ax.plot(angles_closed, [r] * (N + 1), color="gray", lw=0.4, ls=":", alpha=0.4)
    # Radios con color según IS o adicional
    for angle, label in zip(angles, labels_list):
        col_ax = "#2c3e50" if label in IS_METRICS else "#bbbbbb"
        ax.plot([angle, angle], [0, 1], color=col_ax, lw=0.9, alpha=0.6)
    # Etiquetas de métricas pegadas al borde exterior (radio 1.10)
    for angle, label in zip(angles, labels_list):
        col_ax  = "#2c3e50" if label in IS_METRICS else "#888888"
        weight  = "bold"    if label in IS_METRICS else "normal"
        ax.text(angle, 1.10, label, ha="center", va="center",
                fontsize=8.5, color=col_ax, fontweight=weight)
    # Mediana de liga (solo línea, sin relleno)
    ax.plot(angles_closed, med_closed, color="#555", lw=1.4, ls="--", alpha=0.65, zorder=1)

# ── Función para dibujar grupo como media (solo línea, sin relleno) ──
def plot_group(ax, teams, color):
    grp = radar_df[radar_df["Equipo"].isin(teams)]
    if grp.empty:
        return
    mean_v = [grp[m].mean() for m in labels_list]
    mc = mean_v + mean_v[:1]
    ax.plot(angles_closed, mc, color=color, lw=2.4, zorder=3)

# ── Figura ────────────────────────────────────────────────────
fig = plt.figure(figsize=(20, 9))
plt.subplots_adjust(left=0.02, right=0.98, top=0.82, bottom=0.12, wspace=0.45)
ax_rend  = fig.add_subplot(121, polar=True)
ax_estil = fig.add_subplot(122, polar=True)

# Título general de la figura
fig.suptitle(
    "G9 · Radares de Perfil — Rendimiento y Estilo Táctico\n"
    "¿Qué distingue a los equipos exitosos del resto?",
    fontsize=13, fontweight="bold", y=0.97,
)

# Títulos centrados con su respectivo panel
ax_rend.set_title("Perfil por Rendimiento — Índice de Éxito",
                  fontsize=12, fontweight="bold", pad=55)
ax_estil.set_title("Perfil por Estilo Táctico",
                   fontsize=12, fontweight="bold", pad=55)

# ── Panel izquierdo: Rendimiento ──────────────────────────────
setup_radar(ax_rend)
df_sorted = radar_df.sort_values("IS", ascending=False).reset_index(drop=True)
top5 = df_sorted.iloc[:5]["Equipo"].tolist()
mid5 = df_sorted.iloc[7:12]["Equipo"].tolist()
bot5 = df_sorted.iloc[15:]["Equipo"].tolist()

plot_group(ax_rend, top5, "#27ae60")
plot_group(ax_rend, mid5, "#f39c12")
plot_group(ax_rend, bot5, "#e74c3c")

legend_rend = [
    Line2D([0],[0], color="#27ae60", lw=2.5, label=f"Top 5 IS  ({', '.join(top5)})"),
    Line2D([0],[0], color="#f39c12", lw=2.5, label=f"Zona media ({', '.join(mid5)})"),
    Line2D([0],[0], color="#e74c3c", lw=2.5, label=f"Bottom 5 IS ({', '.join(bot5)})"),
    Line2D([0],[0], color="#555", lw=1.4, ls="--", label="Mediana liga"),
]
ax_rend.legend(handles=legend_rend, loc="upper center",
               bbox_to_anchor=(0.5, -0.08), fontsize=7.5,
               framealpha=0.90, ncol=1)

# ── Panel derecho: Estilo táctico ─────────────────────────────
setup_radar(ax_estil)
groups_estilo = [
    (["Real Madrid", "Rayo Vallecano", "Elche"],        "#2ecc71", "Dominadores"),
    (["Espanyol", "Mallorca", "Levante"],                "#3498db", "Counter-press"),
    (["Barcelona", "Atlético Madrid", "Celta Vigo"],     "#e67e22", "Posesión estéril"),
    (["Villarreal", "Osasuna", "Oviedo"],                "#c0392b", "Bloque bajo"),
]
for teams, color, label in groups_estilo:
    plot_group(ax_estil, teams, color)

legend_estil = [
    Line2D([0],[0], color="#2ecc71", lw=2.5, label="Dominadores (R.Madrid, Rayo, Elche)"),
    Line2D([0],[0], color="#3498db", lw=2.5, label="Counter-press (Espanyol, Mallorca, Levante)"),
    Line2D([0],[0], color="#e67e22", lw=2.5, label="Posesión estéril (Barcelona, Atlético, Celta)"),
    Line2D([0],[0], color="#c0392b", lw=2.5, label="Bloque bajo (Villarreal, Osasuna, Oviedo)"),
    Line2D([0],[0], color="#555", lw=1.4, ls="--", label="Mediana liga"),
]
ax_estil.legend(handles=legend_estil, loc="upper center",
                bbox_to_anchor=(0.5, -0.08), fontsize=7.5,
                framealpha=0.90, ncol=1)

fig.text(0.5, 0.02, "Valores normalizados entre 0 y 1",
         ha="center", va="center", fontsize=9, color="gray")
save(fig, "g9_radar_perfil.png")


# ═══════════════════════════════════════════════════════════════
# G16 — RENDIMIENTO DE PORTEROS: Real vs Esperado (xG)
# ---------------------------------------------------------------
# Scatter xGC/pp vs GC/pp real + ranking por GSA/pp.
# GSA/pp = GC/pp − xGC/pp:  negativo → portero mejor que esperado
#                             positivo → portero peor que esperado
# Color = Pts del equipo (rojo→verde).
# ═══════════════════════════════════════════════════════════════
print("[G16] Rendimiento porteros...")

from adjustText import adjust_text as _adj16

raw_por = pd.read_excel(EXCEL, sheet_name="Porteros", header=None)
df_p16  = raw_por.iloc[3:].copy().reset_index(drop=True)

porteros16 = pd.DataFrame({
    "Jugador":       df_p16.iloc[:, 0],
    "Equipo":        df_p16.iloc[:, 3],
    "Minutos":       pd.to_numeric(df_p16.iloc[:, 20], errors="coerce"),
    "GC_total":      pd.to_numeric(df_p16.iloc[:, 23], errors="coerce"),
    "GC_pp":         pd.to_numeric(df_p16.iloc[:, 24], errors="coerce"),
    "SoTA":          pd.to_numeric(df_p16.iloc[:, 25], errors="coerce"),
    "SavePct":       pd.to_numeric(df_p16.iloc[:, 27], errors="coerce"),
    "P0":            pd.to_numeric(df_p16.iloc[:, 31], errors="coerce"),
    "PctP0":         pd.to_numeric(df_p16.iloc[:, 32], errors="coerce"),
    "Par_APeq_pp":   pd.to_numeric(df_p16.iloc[:, 33], errors="coerce"),
    "Par_Area_pp":   pd.to_numeric(df_p16.iloc[:, 36], errors="coerce"),
    "Par_FA_pp":     pd.to_numeric(df_p16.iloc[:, 39], errors="coerce"),
    "SavePK_pct":    pd.to_numeric(df_p16.iloc[:, 46], errors="coerce"),
}).dropna(subset=["Jugador", "Equipo", "Minutos"])

# Solo portero titular de cada equipo (más minutos)
porteros16 = (
    porteros16[porteros16["Minutos"] >= 900]
    .sort_values("Minutos", ascending=False)
    .drop_duplicates(subset="Equipo", keep="first")
    .reset_index(drop=True)
)

# xG en contra por partido (del equipo) y Pts
eq_ref = pd.DataFrame({
    "Equipo":    raw_eq.iloc[4:24, 0].values,
    "xGC_total": pd.to_numeric(raw_eq.iloc[4:24, 63].values, errors="coerce"),
})
cla_ref = pd.DataFrame({
    "Equipo": raw_cla.iloc[3:23, 1].values,
    "PJ":     pd.to_numeric(raw_cla.iloc[3:23, 2].values, errors="coerce"),
    "Pts":    pd.to_numeric(raw_cla.iloc[3:23, 9].values, errors="coerce"),
}).reset_index(drop=True)
eq_ref = eq_ref.merge(cla_ref, on="Equipo")
eq_ref["xGC_pp"] = eq_ref["xGC_total"] / eq_ref["PJ"]

gk = porteros16.merge(eq_ref[["Equipo", "xGC_pp", "Pts"]], on="Equipo", how="left")
gk["GSA_pp"]   = gk["GC_pp"] - gk["xGC_pp"]   # negativo = mejor que esperado
gk["SoTA_pp"]  = gk["SoTA"] / (gk["Minutos"] / 90)

# Apellido para etiquetas
def apellido(nombre):
    partes = str(nombre).split()
    return partes[-1] if len(partes) > 1 else partes[0]

gk["Label"] = gk["Jugador"].apply(apellido)

# ── Figura: scatter izq. + heatmap der. ───────────────────────
fig, (ax_sc, ax_hm) = plt.subplots(1, 2, figsize=(22, 10),
                                    gridspec_kw={"width_ratios": [1, 1.15]})
fig.suptitle("G16 · Rendimiento de Porteros — Real vs Esperado (xG)\n"
             "LaLiga 2025-26",
             fontweight="bold", fontsize=15)

cmap16 = plt.cm.RdYlGn
norm16 = plt.Normalize(gk["Pts"].min(), gk["Pts"].max())
colors16 = [cmap16(norm16(p)) for p in gk["Pts"]]

# ── Panel izquierdo: Scatter xGC/pp vs GC/pp ──────────────────
ax_sc.scatter(gk["xGC_pp"], gk["GC_pp"],
              c=colors16, s=120, edgecolors="#333333",
              linewidths=0.5, zorder=5)

# Línea diagonal (GC = xGC → rendimiento esperado)
lim_min = min(gk["xGC_pp"].min(), gk["GC_pp"].min()) - 0.05
lim_max = max(gk["xGC_pp"].max(), gk["GC_pp"].max()) + 0.05
ax_sc.plot([lim_min, lim_max], [lim_min, lim_max],
           color="#888888", lw=1.5, ls="--", zorder=3,
           label="Rendimiento esperado (GC = xGC)")

# Sombreado zonas
ax_sc.fill_between([lim_min, lim_max], [lim_min, lim_max], lim_max,
                   color="#3498db", alpha=0.06, zorder=0)
ax_sc.fill_between([lim_min, lim_max], lim_min, [lim_min, lim_max],
                   color="#27ae60", alpha=0.06, zorder=0)
ax_sc.text(0.97, 0.97, "Por encima: encaja más\nde lo esperado",
           transform=ax_sc.transAxes, ha="right", va="top",
           fontsize=8, color="#2c3e50", style="italic")
ax_sc.text(0.03, 0.03, "Por debajo: encaja menos\nde lo esperado ✓",
           transform=ax_sc.transAxes, ha="left", va="bottom",
           fontsize=8, color="#2c3e50", style="italic")

texts_sc = [
    ax_sc.text(row["xGC_pp"], row["GC_pp"], row["Label"],
               fontsize=8, color="#1a1a2e", zorder=7)
    for _, row in gk.iterrows()
]
_adj16(texts_sc, ax=ax_sc,
       arrowprops=dict(arrowstyle="-", color="#aaaaaa", lw=0.5))

ax_sc.set_xlabel("xG en contra / partido (calidad defensiva del equipo)", fontsize=11)
ax_sc.set_ylabel("GC real / partido (portero)", fontsize=11)
ax_sc.set_title("xGC/pp vs GC/pp real", fontweight="bold", fontsize=12)
ax_sc.legend(fontsize=9, loc="upper left")
ax_sc.grid(alpha=0.25)

# ── Panel derecho: heatmap correlaciones métricas portero ─────
hm_cols = {
    "GC/pp":             "GC_pp",
    "% Paradas":         "SavePct",
    "Tiros rec./pp":     "SoTA_pp",
    "% Port. a cero":    "PctP0",
    "Par. Área Peq./pp": "Par_APeq_pp",
    "Par. Área/pp":      "Par_Area_pp",
    "Par. F. Área/pp":   "Par_FA_pp",
    "% Pen. parados":    "SavePK_pct",
    "GSA/pp":            "GSA_pp",
    "Pts equipo":        "Pts",
}
df_hm16 = pd.DataFrame({k: gk[v].values for k, v in hm_cols.items()})
corr_hm16 = df_hm16.corr()

pval_hm16 = pd.DataFrame(np.ones_like(corr_hm16),
                          index=corr_hm16.index, columns=corr_hm16.columns)
for c1, c2 in combinations(corr_hm16.columns, 2):
    xy = df_hm16[[c1, c2]].dropna()
    if len(xy) >= 5:
        _, p = stats.pearsonr(xy.iloc[:, 0], xy.iloc[:, 1])
        pval_hm16.loc[c1, c2] = p
        pval_hm16.loc[c2, c1] = p

annot_hm16 = pd.DataFrame("", index=corr_hm16.index, columns=corr_hm16.columns)
for c1 in corr_hm16.columns:
    for c2 in corr_hm16.columns:
        if c1 == c2:
            continue
        r = corr_hm16.loc[c1, c2]
        p = pval_hm16.loc[c1, c2]
        stars = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "NS"
        annot_hm16.loc[c1, c2] = f"{r:.2f}\n{stars}"

mask_hm16 = np.triu(np.ones_like(corr_hm16, dtype=bool))
sns.heatmap(
    corr_hm16, mask=mask_hm16, ax=ax_hm,
    annot=annot_hm16, fmt="",
    annot_kws={"size": 8, "va": "center"},
    cmap="RdYlGn", vmin=-1, vmax=1,
    linewidths=0.4, linecolor="white",
    cbar_kws={"shrink": 0.65, "label": "Pearson r"},
)
ax_hm.set_title("Correlaciones entre métricas de porteros",
                fontweight="bold", fontsize=12)
ax_hm.set_xticklabels(ax_hm.get_xticklabels(), rotation=35, ha="right", fontsize=9)
ax_hm.set_yticklabels(ax_hm.get_yticklabels(), rotation=0, fontsize=9)

fig.text(
    0.5, -0.02,
    "Solo porteros titulares (≥ 900 min)  ·  n = 20 porteros  ·  "
    "Verde = correlación positiva · Rojo = negativa · "
    "*** p<0.001  ** p<0.01  * p<0.05  NS = no significativo",
    ha="center", fontsize=9, color="gray",
)
plt.tight_layout()
save(fig, "g16_porteros_rendimiento.png")


print("\nTodos los graficos guardados en:", OUTDIR)

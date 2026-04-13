"""
evaluacion.py — Visualización del Modelo · Pilar 2 · TFG LaLiga 2025-26
========================================================================
Genera los gráficos de evaluación de los dos modelos predictivos entrenados
en modelos.py. Debe ejecutarse DESPUÉS de modelos.py.

Ambos modelos predicen el mismo target (Δgoles = goles_local − goles_visitante)
y se evalúan con las mismas métricas (MAE, RMSE, R²), lo que permite
una comparación directa y la selección justificada del modelo final.

Gráficos generados en Gráficos/:
  E1  pred_vs_real_comparativa.png → Predicción vs Real — A, B y C (3 paneles)
  E2  residuos_comparados.png      → Residuos de los tres modelos (lado a lado)
  E3  comparativa_modelos.png      → Barras MAE/RMSE/R² + tabla resumen
  E4  confusion_matrix.png         → Matriz de confusión 3×3 — Modelo B
  E5  roc_curves.png               → Curvas ROC/AUC one-vs-rest — Modelo B
  E6  feature_importance.png       → Importancia de variables (top 20) — Modelo B

Entrada:  modelos/resultados_a.pkl, modelos/resultados_b.pkl, modelos/resultados_c.pkl, modelos/metricas.json
Salida:   Gráficos/E1_*.png … E6_*.png

Uso: python3 evaluacion.py
"""

import json
import pickle
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.transforms import blended_transform_factory
import numpy as np
from sklearn.metrics import confusion_matrix, roc_curve, auc
from sklearn.preprocessing import label_binarize

# ══════════════════════════════════════════════════════════════════
# RUTAS
# Directorios de entrada (resultados de modelos.py) y salida (PNG).
# ══════════════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).parent
MOD_DIR  = BASE_DIR / "modelos"
GRAF_DIR = BASE_DIR.parent / "Gráficos"
GRAF_DIR.mkdir(exist_ok=True)

# ══════════════════════════════════════════════════════════════════
# ESTILO VISUAL — paleta consistente con eda.py
# ══════════════════════════════════════════════════════════════════
AZUL     = "#1a3a5c"
AZUL_MED = "#2e6da4"
AZUL_CLA = "#7fb3d6"
ROJO     = "#c0392b"
VERDE    = "#27ae60"
NARANJA  = "#e67e22"
GRIS     = "#bdc3c7"

plt.rcParams.update({
    "font.family":        "DejaVu Sans",
    "axes.spines.top":    False,
    "axes.spines.right":  False,
    "axes.titlesize":     12,
    "axes.labelsize":     10,
    "xtick.labelsize":    9,
    "ytick.labelsize":    9,
    "figure.dpi":         150,
})


def save(fig, nombre):
    """
    Guarda la figura en Gráficos/ y la cierra.

    Args:
        fig:    Figura de matplotlib a guardar.
        nombre: Nombre del archivo PNG (se antepone la ruta Gráficos/).
    """
    ruta = GRAF_DIR / nombre
    fig.savefig(ruta, bbox_inches="tight", dpi=150)
    plt.close(fig)
    print(f"  Guardado: Gráficos/{nombre}")


def cargar_resultados():
    """
    Carga los resultados de evaluación generados por modelos.py.

    Returns:
        Tupla (res_a, res_b, res_c, metricas) con los dicts de resultados
        de los tres modelos y el JSON de métricas comparativas.
    """
    with open(MOD_DIR / "resultados_a.pkl", "rb") as f:
        res_a = pickle.load(f)
    with open(MOD_DIR / "resultados_b.pkl", "rb") as f:
        res_b = pickle.load(f)
    with open(MOD_DIR / "resultados_c.pkl", "rb") as f:
        res_c = pickle.load(f)
    with open(MOD_DIR / "metricas.json", "r", encoding="utf-8") as f:
        metricas = json.load(f)
    return res_a, res_b, res_c, metricas


# ══════════════════════════════════════════════════════════════════
# FUNCIÓN AUXILIAR — scatter pred vs real con cuadrantes sombreados
# Reutilizada en los tres paneles del gráfico E1 comparativo.
# ══════════════════════════════════════════════════════════════════

def _scatter_pred_real(ax, fig, y_real, y_pred, titulo, mae, rmse, r2):
    """
    Dibuja el scatter de predicción vs real con 4 cuadrantes sombreados
    y bandas de error.

    El color de cada punto representa su error absoluto respecto a la
    diagonal (|real − predicho|): verde = cerca de la diagonal (error ~0),
    rojo = lejos de la diagonal (error grande). Una colorbar lateral
    permite leer el error exacto en goles.

    Cuadrantes (referencia: Δ = goles_local − goles_visitante):
      Verde  (Δreal>0, Δpred>0): ambos aciertan victoria local
      Rojo   (Δreal<0, Δpred>0): modelo predice local, gana visitante
      Azul   (Δreal<0, Δpred<0): ambos aciertan victoria visitante
      Naranja(Δreal>0, Δpred<0): modelo predice visitante, gana local

    Args:
        ax:     Eje de matplotlib donde dibujar.
        fig:    Figura padre (para añadir colorbar).
        y_real: Array de valores reales.
        y_pred: Array de valores predichos.
        titulo: Título del subplot.
        mae, rmse, r2: Métricas para anotar en el gráfico.
    """
    import matplotlib.colors as mcolors

    lim = max(abs(y_real).max(), abs(y_pred).max()) + 0.5
    x_line = np.linspace(-lim, lim, 200)

    # ── Sombreado de cuadrantes ──────────────────────────────────────
    ax.fill_between([0,    lim],  0,    lim,  alpha=0.09, color=VERDE,    zorder=1)
    ax.fill_between([-lim, 0],   0,    lim,  alpha=0.09, color=ROJO,     zorder=1)
    ax.fill_between([-lim, 0],  -lim,  0,    alpha=0.09, color=AZUL_CLA, zorder=1)
    ax.fill_between([0,    lim], -lim,  0,    alpha=0.09, color=NARANJA,  zorder=1)

    # Líneas de separación de cuadrantes
    ax.axhline(0, color=GRIS, lw=0.9, alpha=0.8, zorder=2)
    ax.axvline(0, color=GRIS, lw=0.9, alpha=0.8, zorder=2)

    # ── Bandas de error ±1 y ±2 goles ───────────────────────────────
    ax.fill_between(x_line, x_line - 2, x_line + 2,
                    alpha=0.07, color=AZUL_MED, zorder=3)
    ax.fill_between(x_line, x_line - 1, x_line + 1,
                    alpha=0.15, color=AZUL_MED, zorder=3)
    ax.plot(x_line, x_line, "--", color=GRIS, lw=1.5, zorder=4)

    # ── Puntos coloreados por error absoluto respecto a la diagonal ──
    # Error = |real − predicho| = distancia perpendicular (aprox.) a la diagonal
    error_abs = np.abs(y_real - y_pred)
    cmap = mcolors.LinearSegmentedColormap.from_list(
        "verde_rojo", [VERDE, "#f9e04b", ROJO]
    )
    sc = ax.scatter(y_real, y_pred, c=error_abs, cmap=cmap,
                    vmin=0, vmax=error_abs.max(),
                    alpha=0.75, s=30, zorder=5, edgecolors="none")

    # Colorbar lateral
    cbar = fig.colorbar(sc, ax=ax, fraction=0.046, pad=0.04)
    cbar.set_label("Error absoluto (goles)", fontsize=8)
    cbar.ax.tick_params(labelsize=7)

    ax.set_xlabel("Diferencia de goles real (local − visitante)", fontsize=9)
    ax.set_ylabel("Diferencia de goles predicha", fontsize=9)
    ax.set_xlim(-lim, lim)
    ax.set_ylim(-lim, lim)
    ax.set_title(f"{titulo}\nMAE={mae:.3f}  RMSE={rmse:.3f}  R²={r2:.3f}", fontsize=10)


def _leyenda_cuadrantes():
    """
    Devuelve los handles para la leyenda externa de cuadrantes.

    Returns:
        Lista de Patch con colores y etiquetas de los cuatro cuadrantes.
    """
    return [
        mpatches.Patch(facecolor=VERDE,    alpha=0.5, edgecolor="gray", linewidth=0.5,
                       label="Acierto — local gana: Δreal > 0  y  Δpred > 0"),
        mpatches.Patch(facecolor=ROJO,     alpha=0.5, edgecolor="gray", linewidth=0.5,
                       label="Error — modelo predice victoria local, real fue derrota: Δreal < 0, Δpred > 0"),
        mpatches.Patch(facecolor=AZUL_CLA, alpha=0.5, edgecolor="gray", linewidth=0.5,
                       label="Acierto — visitante gana: Δreal < 0  y  Δpred < 0"),
        mpatches.Patch(facecolor=NARANJA,  alpha=0.5, edgecolor="gray", linewidth=0.5,
                       label="Error — modelo predice derrota local, real fue victoria: Δreal > 0, Δpred < 0"),
    ]


# ══════════════════════════════════════════════════════════════════
# E1 — PREDICCIÓN VS REAL — COMPARATIVA A, B y C (3 paneles)
# Los tres modelos con la misma estructura visual para comparación directa.
# ══════════════════════════════════════════════════════════════════

def grafico_pred_real_comparativa(res_a, res_b, res_c):
    """
    Tres paneles horizontales (A · B · C) con el scatter predicción vs real,
    misma estructura y colores que los gráficos individuales originales.
    Permite comparar visualmente los tres modelos en un solo vistazo.
    """
    print("[E1] Pred vs Real — comparativa A, B y C...")

    modelos = [
        (res_a, "Modelo A — Regresión Lineal"),
        (res_b, "Modelo B — Random Forest"),
        (res_c, "Modelo C — XGBoost"),
    ]

    fig, axes = plt.subplots(1, 3, figsize=(21, 7))

    for ax, (res, titulo) in zip(axes, modelos):
        y_real = np.array(res["y_test"])
        y_pred = np.array(res["y_pred"])
        _scatter_pred_real(ax, fig, y_real, y_pred,
                           titulo, res["mae"], res["rmse"], res["r2"])

    fig.legend(
        handles=_leyenda_cuadrantes(),
        loc="lower center",
        bbox_to_anchor=(0.5, -0.01),
        ncol=2,
        fontsize=8,
        frameon=True,
        title="Interpretación de cuadrantes  (Δ = goles_local − goles_visitante)",
        title_fontsize=8.5,
    )

    fig.suptitle(
        "E1 · Predicción vs Real — Comparativa Modelos A, B y C\n"
        "670 predicciones out-of-fold · Validación cruzada 5-Fold",
        fontsize=12, y=1.01
    )
    fig.subplots_adjust(bottom=0.22, wspace=0.35)
    save(fig, "E1_pred_vs_real_comparativa.png")


# ══════════════════════════════════════════════════════════════════
# E2 — RESIDUOS COMPARADOS (tres modelos, mismo gráfico)
# Permite ver de un vistazo cuál modelo tiene errores más centrados.
# ══════════════════════════════════════════════════════════════════

def grafico_residuos_comparados(res_a, res_b, res_c):
    """
    Histogramas de residuos de los tres modelos en un mismo gráfico.

    Un modelo bien calibrado tiene residuos simétricos centrados en 0.
    Comparar la anchura y el sesgo permite evaluar cuál comete errores
    más pequeños y más sistemáticos.
    """
    print("[E2] Residuos comparados...")
    residuos_a = np.array(res_a["y_test"]) - np.array(res_a["y_pred"])
    residuos_b = np.array(res_b["y_test"]) - np.array(res_b["y_pred"])
    residuos_c = np.array(res_c["y_test"]) - np.array(res_c["y_pred"])

    fig, axes = plt.subplots(1, 3, figsize=(16, 4.5), sharey=False)

    for ax, residuos, titulo, color, m in [
        (axes[0], residuos_a, "Modelo A — Regresión Lineal", AZUL_MED, res_a),
        (axes[1], residuos_b, "Modelo B — Random Forest",    AZUL,     res_b),
        (axes[2], residuos_c, "Modelo C — XGBoost",          NARANJA,  res_c),
    ]:
        ax.hist(residuos, bins=25, color=color, edgecolor="white",
                alpha=0.80, density=True)
        ax.axvline(0, color=ROJO, lw=1.8, linestyle="--", label="Residuo = 0")
        ax.axvline(residuos.mean(), color=VERDE, lw=1.5, linestyle="-",
                   label=f"Media = {residuos.mean():.3f}")
        ax.set_xlabel("Residuo (Real − Predicho)")
        ax.set_ylabel("Densidad")
        ax.set_title(f"{titulo}\nMAE={m['mae']:.3f}  RMSE={m['rmse']:.3f}  R²={m['r2']:.3f}",
                     fontsize=10)
        ax.legend(fontsize=8)

    fig.suptitle(
        "E2 · Distribución de Residuos — Comparativa Modelos A, B y C\n"
        "Un modelo bien calibrado tiene residuos simétricos centrados en 0",
        fontsize=11, y=1.02
    )
    fig.tight_layout()
    save(fig, "E2_residuos_comparados.png")


# ══════════════════════════════════════════════════════════════════
# E3 — COMPARATIVA FINAL: barras + tabla resumen
# Permite seleccionar el modelo ganador con justificación empírica.
# ══════════════════════════════════════════════════════════════════

def grafico_comparativa(metricas, res_a, res_b, res_c):
    """
    Barras comparativas de MAE/RMSE/R² para los tres modelos + tabla resumen.

    Args:
        metricas: Dict JSON con las métricas de los tres modelos.
        res_a:    Dict de resultados del Modelo A.
        res_b:    Dict de resultados del Modelo B.
        res_c:    Dict de resultados del Modelo C.
    """
    print("[E3] Comparativa de modelos...")
    m_a = metricas["modelo_a"]
    m_b = metricas["modelo_b"]
    m_c = metricas["modelo_c"]

    fig = plt.figure(figsize=(14, 5.5))
    gs  = fig.add_gridspec(1, 2, width_ratios=[1.3, 1], wspace=0.35)

    # ── Panel izquierdo: barras agrupadas ──────────────────────────
    ax1 = fig.add_subplot(gs[0])

    metricas_barras = ["MAE\n(goles)", "RMSE\n(goles)", "R²"]
    vals_a = [m_a["mae"], m_a["rmse"], m_a["r2"]]
    vals_b = [m_b["mae"], m_b["rmse"], m_b["r2"]]
    vals_c = [m_c["mae"], m_c["rmse"], m_c["r2"]]

    x     = np.arange(3)
    ancho = 0.24
    bars_a = ax1.bar(x - ancho, vals_a, ancho, color=AZUL_MED,
                     label="Modelo A — Regresión Lineal", edgecolor="white")
    bars_b = ax1.bar(x,         vals_b, ancho, color=NARANJA,
                     label="Modelo B — Random Forest",    edgecolor="white")
    bars_c = ax1.bar(x + ancho, vals_c, ancho, color=VERDE,
                     label="Modelo C — XGBoost",          edgecolor="white")

    for bar in list(bars_a) + list(bars_b) + list(bars_c):
        h = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width() / 2, h + 0.005,
                 f"{h:.3f}", ha="center", va="bottom", fontsize=8, color=AZUL)

    ax1.set_xticks(x)
    ax1.set_xticklabels(metricas_barras, fontsize=10)
    ax1.set_ylabel("Valor de la métrica")
    ax1.set_title("Métricas de Evaluación\n(MAE y RMSE: menor es mejor · R²: mayor es mejor)",
                  fontsize=10)
    ax1.legend(fontsize=8.5, loc="upper center",
               bbox_to_anchor=(0.5, -0.22), ncol=1, frameon=True)

    # Indicador del ganador
    trans = blended_transform_factory(ax1.transData, ax1.transAxes)
    mejor_mae  = "B" if m_b["mae"]  <= m_a["mae"]  and m_b["mae"]  <= m_c["mae"]  else ("A" if m_a["mae"]  <= m_c["mae"]  else "C")
    mejor_rmse = "B" if m_b["rmse"] <= m_a["rmse"] and m_b["rmse"] <= m_c["rmse"] else ("A" if m_a["rmse"] <= m_c["rmse"] else "C")
    mejor_r2   = "B" if m_b["r2"]   >= m_a["r2"]   and m_b["r2"]   >= m_c["r2"]   else ("A" if m_a["r2"]   >= m_c["r2"]   else "C")
    for xi, gan in [(0, mejor_mae), (1, mejor_rmse), (2, mejor_r2)]:
        ax1.text(xi, -0.14, f"★ Mejor: {gan}",
                 ha="center", fontsize=8.5, color=VERDE,
                 fontweight="bold", transform=trans)

    # ── Panel derecho: tabla resumen ───────────────────────────────
    ax2 = fig.add_subplot(gs[1])
    ax2.axis("off")

    n = metricas.get("n_partidos", 670)
    ax2.text(0.5, -0.06,
             f"Validación cruzada 5-Fold · {n} partidos (24-25 + 25-26) · "
             "Predicciones out-of-fold",
             ha="center", va="top", fontsize=8, color="gray",
             transform=ax2.transAxes, style="italic")

    best_r2 = max(m_a["r2"], m_b["r2"], m_c["r2"])
    datos_tabla = [
        ["",              "Modelo A",           "Modelo B",        "Modelo C"],
        ["Tipo",          "Reg. Lineal",         "Random Forest",   "XGBoost"],
        ["MAE",           f"{m_a['mae']:.4f}",   f"{m_b['mae']:.4f}",  f"{m_c['mae']:.4f}"],
        ["RMSE",          f"{m_a['rmse']:.4f}",  f"{m_b['rmse']:.4f}", f"{m_c['rmse']:.4f}"],
        ["R²",            f"{m_a['r2']:.4f}",    f"{m_b['r2']:.4f}",   f"{m_c['r2']:.4f}"],
        ["Modelo final",
         "✓" if m_a["r2"] == best_r2 else "—",
         "✓" if m_b["r2"] == best_r2 else "—",
         "✓" if m_c["r2"] == best_r2 else "—"],
    ]

    tabla = ax2.table(
        cellText=datos_tabla[1:],
        colLabels=datos_tabla[0],
        cellLoc="center", loc="center",
        bbox=[0, 0.08, 1, 0.90],
        colWidths=[0.22, 0.26, 0.26, 0.26],
    )
    tabla.auto_set_font_size(False)
    tabla.set_fontsize(9)

    for j in range(4):
        tabla[(0, j)].set_facecolor(AZUL)
        tabla[(0, j)].set_text_props(color="white", fontweight="bold")

    for i in range(1, len(datos_tabla)):
        color = "#EBF3FB" if i % 2 == 0 else "white"
        for j in range(4):
            tabla[(i, j)].set_facecolor(color)

    for j in range(4):
        tabla[(1, j)].set_text_props(fontweight="bold")

    last_row = len(datos_tabla) - 1
    for j in range(4):
        tabla[(last_row, j)].set_facecolor("#D5E8D4")
        tabla[(last_row, j)].set_text_props(fontweight="bold")

    fig.suptitle(
        "E3 · Comparativa Final — Modelos A (Regresión Lineal), B (Random Forest) y C (XGBoost)\n"
        "Mismo target (Δgoles), mismas 111 features, mismas métricas → selección justificada",
        fontsize=11, fontweight="bold", color=AZUL, y=1.02
    )
    fig.subplots_adjust(bottom=0.18)
    save(fig, "E3_comparativa_modelos.png")


# ══════════════════════════════════════════════════════════════════
# FUNCIÓN AUXILIAR — conversión regresión → clasificación de resultado
# Umbral ±0.5: Δ > 0.5 → local, |Δ| ≤ 0.5 → empate, Δ < −0.5 → visitante.
# ══════════════════════════════════════════════════════════════════

CLASES_STR  = ["Victoria\nlocal", "Empate", "Victoria\nvisitante"]
UMBRAL_CLS  = 0.5

def _clasificar(y, umbral=UMBRAL_CLS):
    """Convierte diferencias de goles continuas a índice de clase.
    0 = Victoria local (Δ > umbral)
    1 = Empate         (|Δ| ≤ umbral)
    2 = Victoria visitante (Δ < -umbral)
    """
    return np.where(y > umbral, 0, np.where(y < -umbral, 2, 1))


# ══════════════════════════════════════════════════════════════════
# E4 — MATRIZ DE CONFUSIÓN (Modelo B)
# Muestra aciertos y confusiones entre las tres clases de resultado.
# ══════════════════════════════════════════════════════════════════

def grafico_confusion_matrix(res_b):
    """
    Matriz de confusión 3×3 para la clasificación de resultado del Modelo B.

    Convierte las predicciones continuas (diferencia de goles) a tres clases
    usando el umbral ±0.5 y visualiza los aciertos (diagonal) y los errores
    de clasificación (fuera de la diagonal).

    Args:
        res_b: Dict de resultados del Modelo B.
    """
    print("[E4] Matriz de confusión...")

    y_real = np.array(res_b["y_test"])
    y_pred = np.array(res_b["y_pred"])
    y_real_cls = _clasificar(y_real)
    y_pred_cls = _clasificar(y_pred)

    cm = confusion_matrix(y_real_cls, y_pred_cls, labels=[0, 1, 2])
    # Normalización por fila = recall por clase
    cm_norm = cm.astype(float) / cm.sum(axis=1, keepdims=True)

    fig, ax = plt.subplots(figsize=(8, 6))
    im = ax.imshow(cm_norm, cmap="Blues", vmin=0, vmax=1, aspect="auto")

    for i in range(3):
        for j in range(3):
            count = cm[i, j]
            pct   = cm_norm[i, j]
            color = "white" if pct > 0.52 else AZUL
            ax.text(j, i, f"{count}\n({pct:.0%})",
                    ha="center", va="center", fontsize=13,
                    color=color, fontweight="bold")

    ax.set_xticks([0, 1, 2])
    ax.set_yticks([0, 1, 2])
    ax.set_xticklabels(CLASES_STR, fontsize=10)
    ax.set_yticklabels(CLASES_STR, fontsize=10)
    ax.set_xlabel("Resultado predicho", fontsize=11, labelpad=10)
    ax.set_ylabel("Resultado real", fontsize=11, labelpad=10)

    # Extender límites para acomodar etiquetas pegadas a la matriz
    ax.set_xlim(-0.5, 3.0)
    ax.set_ylim(2.5, -0.85)   # invertido: fila 0 arriba, margen superior para Prec.

    # Precisión por columna — justo encima de la fila superior (y = −0.65 en coords de datos)
    precision = np.diag(cm) / cm.sum(axis=0)
    for j in range(3):
        ax.text(j, -0.68, f"Prec.\n{precision[j]:.0%}",
                ha="center", va="center", fontsize=8.5,
                color=AZUL_MED)

    # Recall por fila — justo a la derecha de la columna 2 (x = 2.6 en coords de datos)
    recall = np.diag(cm) / cm.sum(axis=1)
    for i in range(3):
        ax.text(2.62, i, f"Rec.\n{recall[i]:.0%}",
                ha="left", va="center", fontsize=8.5,
                color=AZUL_MED)

    fig.colorbar(im, ax=ax, fraction=0.040, pad=0.12,
                 label="Tasa de acierto por clase real (recall normalizado)")

    # Totales reales por clase
    totales = cm.sum(axis=1)
    n = len(y_real)
    distribuciones = "  |  ".join(
        [f"{CLASES_STR[i].replace(chr(10), ' ')}: {totales[i]} ({totales[i]/n:.0%})"
         for i in range(3)]
    )

    fig.suptitle(
        "E4 · Matriz de Confusión — Modelo B (Random Forest)\n"
        "Umbral de clasificación: Δ > +0.5 → local,  |Δ| ≤ 0.5 → empate,  Δ < −0.5 → visitante\n"
        f"Distribución real:  {distribuciones}",
        fontsize=9, y=1.03
    )
    fig.tight_layout()
    save(fig, "E4_confusion_matrix.png")


# ══════════════════════════════════════════════════════════════════
# E5 — CURVAS ROC / AUC (Modelo B — one-vs-rest, 3 clases)
# Usa la predicción continua como score de confianza por clase.
# ══════════════════════════════════════════════════════════════════

def grafico_roc(res_b):
    """
    Curvas ROC one-vs-rest para las tres clases de resultado del Modelo B.

    Usa la predicción continua de diferencia de goles como score de confianza:
      · Victoria local:     score = +Δ̂  (mayor → más probable victoria local)
      · Empate:             score = −|Δ̂| (más próximo a 0 → más probable empate)
      · Victoria visitante: score = −Δ̂  (menor Δ̂ → más probable victoria visitante)

    Args:
        res_b: Dict de resultados del Modelo B.
    """
    print("[E5] Curvas ROC...")

    y_real = np.array(res_b["y_test"])
    y_pred = np.array(res_b["y_pred"])

    y_real_cls = _clasificar(y_real)
    y_bin      = label_binarize(y_real_cls, classes=[0, 1, 2])

    scores = np.column_stack([
        y_pred,           # clase 0: victoria local
        -np.abs(y_pred),  # clase 1: empate
        -y_pred,          # clase 2: victoria visitante
    ])

    colores_cls    = [VERDE, NARANJA, AZUL_MED]
    nombres_cls    = ["Victoria local", "Empate", "Victoria visitante"]
    linestyles_cls = ["-", "--", "-."]

    fig, ax = plt.subplots(figsize=(7, 6.5))

    aucs = []
    for i, (nombre, color, ls) in enumerate(zip(nombres_cls, colores_cls, linestyles_cls)):
        fpr, tpr, _ = roc_curve(y_bin[:, i], scores[:, i])
        roc_auc = auc(fpr, tpr)
        aucs.append(roc_auc)
        ax.plot(fpr, tpr, color=color, lw=2.5, linestyle=ls,
                label=f"{nombre}  (AUC = {roc_auc:.3f})")

    ax.plot([0, 1], [0, 1], linestyle=":", color=GRIS, lw=1.8,
            label="Clasificador aleatorio  (AUC = 0.500)")

    ax.fill_between([0, 1], [0, 1], [1, 1], alpha=0.04, color=VERDE,
                    label="Zona de mejor-que-azar")

    ax.set_xlim([0.0, 1.0])
    ax.set_ylim([0.0, 1.05])
    ax.set_xlabel("Tasa de Falsos Positivos  (1 − Especificidad)", fontsize=10)
    ax.set_ylabel("Tasa de Verdaderos Positivos  (Sensibilidad)", fontsize=10)
    ax.legend(loc="lower right", fontsize=9, frameon=True)
    ax.grid(alpha=0.25)

    macro_auc = np.mean(aucs)
    fig.suptitle(
        "E5 · Curvas ROC — Modelo B (Random Forest) · One-vs-Rest\n"
        f"AUC macro promedio = {macro_auc:.3f}  "
        f"(score = predicción continua de diferencia de goles)",
        fontsize=10, y=1.02
    )
    fig.tight_layout()
    save(fig, "E5_roc_curves.png")


# ══════════════════════════════════════════════════════════════════
# E6 — IMPORTANCIA DE VARIABLES (Modelo B — top 20, por bloque)
# Reducción media de impureza Gini, coloreada por bloque de origen.
# ══════════════════════════════════════════════════════════════════

def grafico_feature_importance(res_b):
    """
    Gráfico de barras horizontales con las 20 variables más importantes
    del Modelo B (Random Forest), coloreadas por bloque de origen.

    La importancia es la reducción media de impureza de Gini acumulada
    a través de los 300 árboles del forest.

    Args:
        res_b: Dict de resultados del Modelo B.
    """
    print("[E6] Importancia de variables...")

    importancias   = np.array(res_b["importancias"])
    feature_names  = res_b["feature_names"]

    top_n  = 20
    idx    = np.argsort(importancias)[::-1][:top_n]
    names  = [feature_names[i] for i in idx]
    vals   = [importancias[i]  for i in idx]

    # ── Asignación de bloque por nombre ──────────────────────────
    CLS_COLS = {"d_Pts_pp", "d_GF_pp", "d_GC_pp", "d_GD_pp", "d_Pos_pct"}
    CTX_COLS = {"lluvia", "arb_faltas_pp", "arb_amarillas_pp"}

    def bloque(name):
        if "d_por_"  in name: return "Portero titular"
        if "d_jug_"  in name: return "Jugadores"
        if name in CLS_COLS:  return "Clasificación"
        if name in CTX_COLS:  return "Contextuales"
        if name == "es_local": return "Ventaja local"
        return "WhoScored"

    BLOCK_COLORS = {
        "WhoScored":       AZUL_MED,
        "Clasificación":   NARANJA,
        "Portero titular": VERDE,
        "Jugadores":       "#8e44ad",
        "Contextuales":    ROJO,
        "Ventaja local":   GRIS,
    }

    # ── Nombres limpios para el eje ───────────────────────────────
    def clean(n):
        n = n.replace("d_", "Δ ")
        n = n.replace("_gen", "").replace("_fav", " (fav)").replace("_con", " (con)")
        n = n.replace("_pp", "/90").replace("_pct", " %")
        return n

    blocks      = [bloque(n)  for n in names]
    colors_list = [BLOCK_COLORS[b] for b in blocks]
    clean_names = [clean(n) for n in names]

    # Invertir para que la barra más importante quede arriba
    y_pos  = np.arange(top_n)
    vals_r = vals[::-1]
    col_r  = colors_list[::-1]
    lbl_r  = clean_names[::-1]

    fig, ax = plt.subplots(figsize=(11, 8))
    bars = ax.barh(y_pos, vals_r, color=col_r, edgecolor="white", alpha=0.88, height=0.72)

    # Anotar valor en cada barra
    for bar, v in zip(bars, vals_r):
        ax.text(bar.get_width() + 0.0005, bar.get_y() + bar.get_height() / 2,
                f"{v:.4f}", va="center", ha="left", fontsize=8, color=AZUL)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(lbl_r, fontsize=9)
    ax.set_xlabel("Importancia (reducción media de impureza Gini)", fontsize=10)
    ax.set_xlim(0, max(vals_r) * 1.18)

    # Línea vertical en la importancia uniforme (1/n_features)
    n_feat = len(feature_names)
    ax.axvline(1 / n_feat, linestyle="--", color=GRIS, lw=1.2, alpha=0.8,
               label=f"Importancia uniforme (1/{n_feat} = {1/n_feat:.4f})")

    # Leyenda de bloques
    bloques_presentes = list(dict.fromkeys(blocks))  # orden de aparición, sin duplicados
    handles = [mpatches.Patch(color=BLOCK_COLORS[b], label=b)
               for b in bloques_presentes]
    ax.legend(handles=handles, title="Bloque de origen", fontsize=8.5,
              title_fontsize=9, loc="lower right", frameon=True)

    # Porcentaje acumulado del top-20
    pct_top20 = sum(vals) / importancias.sum() * 100
    fig.suptitle(
        f"E6 · Importancia de Variables — Modelo B (Random Forest) · Top {top_n}\n"
        f"El top {top_n} acumula el {pct_top20:.1f}% de la importancia total · "
        f"Validación cruzada 5-Fold, {len(importancias)} features",
        fontsize=10, y=1.01
    )
    fig.tight_layout()
    save(fig, "E6_feature_importance.png")


# ══════════════════════════════════════════════════════════════════
# EJECUCIÓN PRINCIPAL
# Carga los resultados de modelos.py y genera los seis gráficos E1–E6.
# ══════════════════════════════════════════════════════════════════

def main():
    print("\n" + "═"*60)
    print("  EVALUACIÓN Y VISUALIZACIÓN DE MODELOS")
    print("═"*60 + "\n")

    res_a, res_b, res_c, metricas = cargar_resultados()

    grafico_pred_real_comparativa(res_a, res_b, res_c)
    grafico_residuos_comparados(res_a, res_b, res_c)
    grafico_comparativa(metricas, res_a, res_b, res_c)
    grafico_confusion_matrix(res_b)
    grafico_roc(res_b)
    grafico_feature_importance(res_b)

    print("\n" + "═"*60)
    print("  Gráficos E1–E6 generados en Gráficos/")
    print("═"*60 + "\n")


if __name__ == "__main__":
    main()

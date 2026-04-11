"""
evaluacion.py — Visualización del Modelo · Pilar 2 · TFG LaLiga 2025-26
========================================================================
Genera los gráficos de evaluación de los dos modelos predictivos entrenados
en modelos.py. Debe ejecutarse DESPUÉS de modelos.py.

Ambos modelos predicen el mismo target (Δgoles = goles_local − goles_visitante)
y se evalúan con las mismas métricas (MAE, RMSE, R²), lo que permite
una comparación directa y la selección justificada del modelo final.

Gráficos generados en Gráficos/:
  E1  pred_vs_real_A.png        → Predicción vs Real — Modelo A (Lineal)
  E2  pred_vs_real_B.png        → Predicción vs Real — Modelo B (RF)
  E3  residuos_comparados.png   → Residuos de ambos modelos (lado a lado)
  E4  comparativa_modelos.png   → Barras MAE/RMSE/R² + tabla resumen

Entrada:  modelos/resultados_a.pkl, modelos/resultados_b.pkl, modelos/metricas.json
Salida:   Gráficos/E1_*.png … E4_*.png

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

# ══════════════════════════════════════════════════════════════════
# RUTAS
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
    """Guarda la figura en Gráficos/ y la cierra."""
    ruta = GRAF_DIR / nombre
    fig.savefig(ruta, bbox_inches="tight", dpi=150)
    plt.close(fig)
    print(f"  Guardado: Gráficos/{nombre}")


def cargar_resultados():
    """
    Carga los resultados de evaluación generados por modelos.py.

    Returns:
        Tupla (res_a, res_b, metricas).
    """
    with open(MOD_DIR / "resultados_a.pkl", "rb") as f:
        res_a = pickle.load(f)
    with open(MOD_DIR / "resultados_b.pkl", "rb") as f:
        res_b = pickle.load(f)
    with open(MOD_DIR / "metricas.json", "r", encoding="utf-8") as f:
        metricas = json.load(f)
    return res_a, res_b, metricas


# ══════════════════════════════════════════════════════════════════
# FUNCIÓN AUXILIAR — scatter pred vs real con cuadrantes sombreados
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
    """Devuelve los handles para la leyenda externa de cuadrantes."""
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
# E1 — PREDICCIÓN VS REAL — MODELO A (Regresión Lineal)
# ══════════════════════════════════════════════════════════════════

def grafico_pred_real_A(res_a):
    """
    Scatter de diferencia de goles predicha vs real para el Modelo A.

    La diagonal representa predicción perfecta. Las bandas sombreadas
    delimitan el error de ±1 y ±2 goles. Los cuadrantes indican si
    el modelo acierta o falla el signo del resultado.

    Args:
        res_a: Dict de resultados del Modelo A.
    """
    print("[E1] Pred vs Real — Modelo A...")
    y_real = np.array(res_a["y_test"])
    y_pred = np.array(res_a["y_pred"])

    fig, ax = plt.subplots(figsize=(7, 7))
    _scatter_pred_real(ax, fig, y_real, y_pred,
                       "Modelo A — Regresión Lineal",
                       res_a["mae"], res_a["rmse"], res_a["r2"])

    fig.legend(
        handles=_leyenda_cuadrantes(),
        loc="lower center",
        bbox_to_anchor=(0.5, -0.01),
        ncol=1,
        fontsize=8,
        frameon=True,
        title="Interpretación de cuadrantes  (Δ = goles_local − goles_visitante)",
        title_fontsize=8.5,
    )

    fig.suptitle(
        "E1 · Predicción vs Real — Modelo A (Regresión Lineal Múltiple)\n"
        "Temporada 25-26 como conjunto de test (validación temporal)",
        fontsize=11, y=1.01
    )
    fig.subplots_adjust(bottom=0.28)
    save(fig, "E1_pred_vs_real_A.png")


# ══════════════════════════════════════════════════════════════════
# E2 — PREDICCIÓN VS REAL — MODELO B (Random Forest)
# ══════════════════════════════════════════════════════════════════

def grafico_pred_real_B(res_b):
    """
    Scatter de diferencia de goles predicha vs real para el Modelo B.

    Mismo formato que E1 para facilitar la comparación visual directa.

    Args:
        res_b: Dict de resultados del Modelo B.
    """
    print("[E2] Pred vs Real — Modelo B...")
    y_real = np.array(res_b["y_test"])
    y_pred = np.array(res_b["y_pred"])

    fig, ax = plt.subplots(figsize=(7, 7))
    _scatter_pred_real(ax, fig, y_real, y_pred,
                       "Modelo B — Random Forest Regressor",
                       res_b["mae"], res_b["rmse"], res_b["r2"])

    fig.legend(
        handles=_leyenda_cuadrantes(),
        loc="lower center",
        bbox_to_anchor=(0.5, -0.01),
        ncol=1,
        fontsize=8,
        frameon=True,
        title="Interpretación de cuadrantes  (Δ = goles_local − goles_visitante)",
        title_fontsize=8.5,
    )

    fig.suptitle(
        "E2 · Predicción vs Real — Modelo B (Random Forest Regressor)\n"
        "Temporada 25-26 como conjunto de test (validación temporal)",
        fontsize=11, y=1.01
    )
    fig.subplots_adjust(bottom=0.28)
    save(fig, "E2_pred_vs_real_B.png")


# ══════════════════════════════════════════════════════════════════
# E3 — RESIDUOS COMPARADOS (ambos modelos, mismo gráfico)
# Permite ver de un vistazo cuál modelo tiene errores más centrados.
# ══════════════════════════════════════════════════════════════════

def grafico_residuos_comparados(res_a, res_b):
    """
    Histogramas de residuos de ambos modelos en un mismo gráfico.

    Un modelo bien calibrado tiene residuos simétricos centrados en 0.
    Comparar la anchura y el sesgo de ambas distribuciones permite
    evaluar cuál comete errores más pequeños y más sistemáticos.

    Args:
        res_a: Dict de resultados del Modelo A.
        res_b: Dict de resultados del Modelo B.
    """
    print("[E3] Residuos comparados...")
    residuos_a = np.array(res_a["y_test"]) - np.array(res_a["y_pred"])
    residuos_b = np.array(res_b["y_test"]) - np.array(res_b["y_pred"])

    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5), sharey=False)

    for ax, residuos, titulo, color, m in [
        (axes[0], residuos_a, "Modelo A — Regresión Lineal", AZUL_MED, res_a),
        (axes[1], residuos_b, "Modelo B — Random Forest",   AZUL,     res_b),
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
        "E3 · Distribución de Residuos — Comparativa Modelo A vs Modelo B\n"
        "Un modelo bien calibrado tiene residuos simétricos centrados en 0",
        fontsize=11, y=1.02
    )
    fig.tight_layout()
    save(fig, "E3_residuos_comparados.png")


# ══════════════════════════════════════════════════════════════════
# E4 — COMPARATIVA FINAL: barras + tabla resumen
# Permite seleccionar el modelo ganador con justificación empírica.
# ══════════════════════════════════════════════════════════════════

def grafico_comparativa(metricas, res_a, res_b):
    """
    Gráfico doble: barras comparativas de MAE/RMSE/R² + tabla resumen.

    Las barras permiten comparar visualmente las tres métricas.
    La tabla incluye la interpretación de cada métrica para facilitar
    la defensa oral del modelo seleccionado.
    La leyenda de las barras y la nota de train/test están fuera de sus
    respectivos elementos para no saturar el gráfico.

    Args:
        metricas: Dict de metricas.json con valores de ambos modelos.
        res_a:    Dict de resultados del Modelo A.
        res_b:    Dict de resultados del Modelo B.
    """
    print("[E4] Comparativa de modelos...")
    m_a = metricas["modelo_a"]
    m_b = metricas["modelo_b"]

    fig = plt.figure(figsize=(13, 5.5))
    gs  = fig.add_gridspec(1, 2, width_ratios=[1.2, 1], wspace=0.35)

    # ── Panel izquierdo: barras agrupadas ──────────────────────────
    ax1 = fig.add_subplot(gs[0])

    metricas_barras = ["MAE\n(goles)", "RMSE\n(goles)", "R²"]
    vals_a = [m_a["mae"], m_a["rmse"], m_a["r2"]]
    vals_b = [m_b["mae"], m_b["rmse"], m_b["r2"]]

    x     = np.arange(3)
    ancho = 0.32
    bars_a = ax1.bar(x - ancho/2, vals_a, ancho, color=AZUL_MED,
                     label="Modelo A — Regresión Lineal", edgecolor="white")
    bars_b = ax1.bar(x + ancho/2, vals_b, ancho, color=NARANJA,
                     label="Modelo B — Random Forest",    edgecolor="white")

    # Anotar valores sobre las barras
    for bar in list(bars_a) + list(bars_b):
        h = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width() / 2, h + 0.005,
                 f"{h:.3f}", ha="center", va="bottom", fontsize=9, color=AZUL)

    ax1.set_xticks(x)
    ax1.set_xticklabels(metricas_barras, fontsize=10)
    ax1.set_ylabel("Valor de la métrica")
    ax1.set_title("Métricas de Evaluación\n(MAE y RMSE: menor es mejor · R²: mayor es mejor)",
                  fontsize=10)

    # Leyenda fuera del área de barras, debajo
    ax1.legend(fontsize=9, loc="upper center",
               bbox_to_anchor=(0.5, -0.20), ncol=1, frameon=True)

    # Estrellas perfectamente alineadas bajo cada grupo — blended transform
    ganadores = [
        (0, "B" if m_b["mae"]  < m_a["mae"]  else "A"),
        (1, "B" if m_b["rmse"] < m_a["rmse"] else "A"),
        (2, "B" if m_b["r2"]   > m_a["r2"]   else "A"),
    ]
    trans = blended_transform_factory(ax1.transData, ax1.transAxes)
    for xi, gan in ganadores:
        ax1.text(xi, -0.13, f"★ Mejor: {gan}",
                 ha="center", fontsize=8.5, color=VERDE,
                 fontweight="bold", transform=trans)

    # ── Panel derecho: tabla resumen ───────────────────────────────
    ax2 = fig.add_subplot(gs[1])
    ax2.axis("off")

    # Nota de train/test fuera de la tabla, como pie de figura
    ax2.text(0.5, -0.06,
             "Entrenamiento: temporada 24-25 (380 partidos)  ·  "
             "Test: temporada 25-26 (290 partidos)",
             ha="center", va="top", fontsize=8, color="gray",
             transform=ax2.transAxes, style="italic")

    datos_tabla = [
        ["",               "Modelo A",              "Modelo B"],
        ["Tipo",           "Regresión Lineal",       "Random Forest"],
        ["MAE",            f"{m_a['mae']:.4f}",      f"{m_b['mae']:.4f}"],
        ["RMSE",           f"{m_a['rmse']:.4f}",     f"{m_b['rmse']:.4f}"],
        ["R²",             f"{m_a['r2']:.4f}",       f"{m_b['r2']:.4f}"],
        ["Interpretable",  "Sí (coeficientes)",      "Parcial (importancias)"],
        ["Modelo final",
         "✓" if m_a["r2"] >= m_b["r2"] else "—",
         "" if m_b["r2"] <= m_a["r2"] else "✓"],
    ]

    tabla = ax2.table(
        cellText=datos_tabla[1:],
        colLabels=datos_tabla[0],
        cellLoc="center", loc="center",
        bbox=[0, 0.08, 1, 0.90],
        colWidths=[0.24, 0.36, 0.40],
    )
    tabla.auto_set_font_size(False)
    tabla.set_fontsize(9.5)

    # Cabecera
    for j in range(3):
        tabla[(0, j)].set_facecolor(AZUL)
        tabla[(0, j)].set_text_props(color="white", fontweight="bold")

    # Filas alternadas
    for i in range(1, len(datos_tabla)):
        color = "#EBF3FB" if i % 2 == 0 else "white"
        for j in range(3):
            tabla[(i, j)].set_facecolor(color)

    # Fila "Tipo" (índice 1) en negrita
    for j in range(3):
        tabla[(1, j)].set_text_props(fontweight="bold")

    # Fila "Modelo final" (última) destacada en verde
    last_row = len(datos_tabla) - 1
    for j in range(3):
        tabla[(last_row, j)].set_facecolor("#D5E8D4")
        tabla[(last_row, j)].set_text_props(fontweight="bold")

    fig.suptitle(
        "E4 · Comparativa Final — Modelo A (Regresión Lineal) vs Modelo B (Random Forest)\n"
        "Mismo target (Δgoles), mismas features, mismas métricas → selección justificada",
        fontsize=11, fontweight="bold", color=AZUL, y=1.02
    )
    fig.subplots_adjust(bottom=0.15)
    save(fig, "E4_comparativa_modelos.png")


# ══════════════════════════════════════════════════════════════════
# EJECUCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════

def main():
    print("\n" + "═"*60)
    print("  EVALUACIÓN Y VISUALIZACIÓN DE MODELOS")
    print("═"*60 + "\n")

    res_a, res_b, metricas = cargar_resultados()

    grafico_pred_real_A(res_a)
    grafico_pred_real_B(res_b)
    grafico_residuos_comparados(res_a, res_b)
    grafico_comparativa(metricas, res_a, res_b)

    print("\n" + "═"*60)
    print("  Gráficos E1–E4 generados en Gráficos/")
    print("═"*60 + "\n")


if __name__ == "__main__":
    main()

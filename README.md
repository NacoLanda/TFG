# Analítica Prescriptiva en Fútbol: Sistema Predictivo Táctico con ML y Big Data
**Autor:** Ignacio Landaluce Gortázar · 4BA1 · Universidad Francisco de Vitoria  
**App desplegada:** [futbolytics.streamlit.app](https://futbolytics.streamlit.app)

---

## Descripción

Herramienta de análisis prescriptivo para LaLiga que genera informes tácticos personalizados integrando más de 100 variables estadísticas en tiempo real. Incluye modelos predictivos de resultado (Regresión Lineal, Random Forest y XGBoost), un índice de éxito propio validado con R²=0.924 sobre puntos reales, y una aplicación Streamlit lista para el cuerpo técnico.

---

## Estructura del repositorio

```
TFG/
│
├── app.py                      # Aplicación Streamlit — interfaz prescriptiva para el cuerpo técnico
├── modelos.py                  # Entrenamiento de modelos predictivos (Δgoles, K-Fold 5)
├── evaluacion.py               # Evaluación comparativa y gráficos E1–E6
├── demo_video.py               # Montaje del vídeo demo (capturas + voz + subtítulos)
├── requirements.txt            # Dependencias para despliegue en Streamlit Cloud
├── Base de datos.xlsx          # Base de datos principal (temporada 25-26)
│
├── modelos/
│   ├── modelo_b_rf.pkl         # Random Forest entrenado — modelo de producción
│   └── feature_names.pkl       # Orden exacto de los 111 features del modelo
│
├── Tablas Excel/               # Salida de scrapers — temporada 25-26
│   ├── Datos WhoScored.xlsx
│   ├── Datos FBref.xlsx
│   ├── Datos Árbitros.xlsx
│   ├── Datos Lluvias.xlsx
│   ├── Jugadores Unificados.xlsx
│   ├── Jugadores FBref.xlsx
│   ├── Duplas Peligrosas.xlsx
│   ├── Lesiones y Sanciones.xlsx
│   └── Partidos.xlsx
│
├── Tablas Excel 24-25/         # Salida de scrapers — temporada 24-25 (datos históricos)
│
├── Scripts Python/             # Scrapers y análisis — temporada 25-26
│   ├── eda.py                  # Análisis exploratorio — genera gráficos G1–G17
│   ├── Actualización.py        # Orquestador: lanza todos los scrapers secuencialmente
│   ├── Otras Fuentes/
│   │   ├── Código Árbitros.py
│   │   ├── Código Duplas Peligrosas.py
│   │   ├── Código FBref.py
│   │   ├── Código Jugadores FBref.py
│   │   ├── Código Lesiones y Sanciones.py
│   │   ├── Código Lluvias.py
│   │   └── Código Partidos.py
│   └── WhoScored/
│       ├── who.py
│       ├── general.py
│       ├── detallado.py
│       ├── posicionales.py
│       ├── situacionales.py
│       ├── jugadores.py
│       └── jugadores2.py
│
└── Scripts Python 24-25/       # Scrapers equivalentes — temporada 24-25
```

---

## Modelos predictivos

**Dataset:** 670 partidos (temporadas 24-25 y 25-26). Target: diferencia de goles local−visitante.  
**111 variables** (diferencias local−visitante): 85 WhoScored + 5 clasificación + 10 portero + 6 jugadores + 3 forma reciente + 2 contextuales.

| Modelo | MAE | RMSE | R² |
|---|---|---|---|
| A — Regresión Lineal | 1.1351 | 1.4644 | 0.2087 |
| **B — Random Forest** | **1.0969** | **1.4188** | **0.2573** |
| C — XGBoost | 1.1570 | 1.4857 | 0.1855 |

Accuracy clasificación (Modelo B): **49.6%** (baseline azar: 36.3%). AUC macro: 0.694.

---

## Índice de Éxito (IS)

```
IS_equipo   = 0.35·minmax(xG/PJ) + 0.35·minmax(−GC/PJ) + 0.30·minmax(Rating)
IS_jugador  = 0.35·norm(OfScore90) + 0.35·norm(DefScore90) + 0.30·norm(Rating)
```

Validación: r=0.87 con Pts (p<0.001), R²=0.924.

---

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Fuentes de datos

| Fuente | Datos |
|---|---|
| WhoScored | Métricas tácticas de equipo y jugador (85 variables) |
| FBref | xG, porteros, partidos, árbitros |
| estadisticaslaliga.es | Clasificación y estadísticas de árbitros |
| FutbolFantasy | Lesionados y sancionados |
| Open-Meteo | Datos climatológicos por ciudad |
| TransferMarkt | Asignaciones de árbitros |

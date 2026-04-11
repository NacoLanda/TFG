# TFG — Algoritmo Prescriptivo LaLiga 2025-26

Trabajo de Fin de Grado · Ingeniería del Dato · Universidad Francisco de Vitoria

Aplicación de análisis y recomendación táctica para cuerpos técnicos de fútbol profesional. A partir de variables de un partido (rival, árbitro, condiciones del encuentro), el sistema genera recomendaciones de alineación, formación y estrategia calibradas por un Índice de Éxito (IS).

Nota: Todos los datos del conjunto de test están actualizados tras la jornada 29.
---

## Estructura del repositorio

```
├── Base de datos.xlsx          # Base de datos principal (20 hojas, temporada 25-26)
├── modelos.py                  # Entrenamiento de modelos predictivos (Δgoles)
├── evaluacion.py               # Visualización y evaluación comparativa de modelos
│
├── Scripts Python/             # Scrapers y análisis — temporada 25-26
│   ├── app.py                  # Aplicación principal (interfaz para el cuerpo técnico)
│   ├── eda.py                  # Análisis exploratorio — genera gráficos G1-G16
│   ├── Actualización.py        # Orquestador: actualiza todos los datos
│   ├── Otras Fuentes/
│   │   ├── Código Árbitros.py
│   │   ├── Código Duplas Peligrosas.py
│   │   ├── Código FBref.py
│   │   ├── Código Jugadores FBref.py
│   │   ├── Código Lesiones y Sanciones.py
│   │   ├── Código Lluvias.py
│   │   └── Código Partidos.py
│   └── WhoScored/
│       ├── who.py              # Orquestador WhoScored
│       ├── general.py
│       ├── detallado.py
│       ├── posicionales.py
│       ├── situacionales.py
│       ├── jugadores.py
│       └── jugadores2.py
│
├── Scripts Python 24-25/       # Scrapers equivalentes — temporada 24-25 (datos históricos)
│
├── Tablas Excel/               # Salida de scrapers — temporada 25-26
└── Tablas Excel 24-25/         # Salida de scrapers — temporada 24-25
```

---

## Modelos predictivos

Ambos modelos predicen **Δgoles = goles_local − goles_visitante** sobre el mismo conjunto de features, lo que permite una comparación directa.

| | Modelo A | Modelo B |
|---|---|---|
| Algoritmo | Regresión Lineal Múltiple | Random Forest Regressor |
| Train | Temporada 24-25 (380 partidos) | Temporada 24-25 (380 partidos) |
| Test | Temporada 25-26 (290 partidos) | Temporada 25-26 (290 partidos) |
| Features | 57 diferenciales local−visitante | 57 diferenciales local−visitante |

```bash
python3 modelos.py      # entrena y guarda modelos en modelos/
python3 evaluacion.py   # genera gráficos E1-E4 en Gráficos/
```

---

## Requisitos

```bash
pip install pandas openpyxl requests beautifulsoup4 selenium scikit-learn matplotlib numpy
```

Selenium requiere [ChromeDriver](https://chromedriver.chromium.org/) compatible con la versión de Chrome instalada.

---

## Flujo de actualización

1. Ejecutar `Actualización.py` para lanzar todos los scrapers secuencialmente.
2. Los datos se guardan en `Tablas Excel/` y se consolidan en `Base de datos.xlsx`.
3. Ejecutar `eda.py` para regenerar los gráficos de análisis.
4. Ejecutar `modelos.py` + `evaluacion.py` para actualizar los modelos y sus visualizaciones.
